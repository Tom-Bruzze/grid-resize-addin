/* ===== Grid Resize Tool - JavaScript ===== */
var CM_TO_POINTS = 28.3465;
var MIN_SIZE_CM = 0.1;
var gridUnitCm = 0.21;

Office.onReady(function (info) {
    if (info.host === Office.HostType.PowerPoint) { initUI(); }
});

function initUI() {
    var gridInput = document.getElementById("gridUnit");
    gridInput.addEventListener("change", function () {
        var val = parseFloat(this.value);
        if (!isNaN(val) && val > 0) { gridUnitCm = val; updatePresetButtons(val); showStatus("Rastereinheit: " + val.toFixed(2) + " cm", "info"); }
    });
    document.querySelectorAll(".preset-btn").forEach(function (btn) {
        btn.addEventListener("click", function () {
            var val = parseFloat(this.getAttribute("data-value"));
            gridUnitCm = val; gridInput.value = val; updatePresetButtons(val);
            showStatus("Rastereinheit: " + val.toFixed(2) + " cm", "info");
        });
    });
    document.querySelectorAll(".tab-btn").forEach(function (btn) {
        btn.addEventListener("click", function () {
            var tabId = this.getAttribute("data-tab");
            document.querySelectorAll(".tab-btn").forEach(function (b) { b.classList.remove("active"); });
            document.querySelectorAll(".tab-content").forEach(function (t) { t.classList.remove("active"); });
            this.classList.add("active");
            document.getElementById(tabId).classList.add("active");
        });
    });
    document.getElementById("shrinkWidth").addEventListener("click", function () { resizeShapes("width", -gridUnitCm); });
    document.getElementById("growWidth").addEventListener("click", function () { resizeShapes("width", gridUnitCm); });
    document.getElementById("shrinkHeight").addEventListener("click", function () { resizeShapes("height", -gridUnitCm); });
    document.getElementById("growHeight").addEventListener("click", function () { resizeShapes("height", gridUnitCm); });
    document.getElementById("shrinkBoth").addEventListener("click", function () { resizeShapes("both", -gridUnitCm); });
    document.getElementById("growBoth").addEventListener("click", function () { resizeShapes("both", gridUnitCm); });
    document.getElementById("propShrink").addEventListener("click", function () { proportionalResize(-gridUnitCm); });
    document.getElementById("propGrow").addEventListener("click", function () { proportionalResize(gridUnitCm); });
    document.getElementById("snapPosition").addEventListener("click", function () { snapToGrid("position"); });
    document.getElementById("snapSize").addEventListener("click", function () { snapToGrid("size"); });
    document.getElementById("snapBoth").addEventListener("click", function () { snapToGrid("both"); });
    document.getElementById("showInfo").addEventListener("click", function () { showShapeInfo(); });
    document.getElementById("matchWidthMax").addEventListener("click", function () { matchDimension("width", "max"); });
    document.getElementById("matchWidthMin").addEventListener("click", function () { matchDimension("width", "min"); });
    document.getElementById("matchHeightMax").addEventListener("click", function () { matchDimension("height", "max"); });
    document.getElementById("matchHeightMin").addEventListener("click", function () { matchDimension("height", "min"); });
    document.getElementById("matchBothMax").addEventListener("click", function () { matchDimension("both", "max"); });
    document.getElementById("matchBothMin").addEventListener("click", function () { matchDimension("both", "min"); });
    document.getElementById("propMatchMax").addEventListener("click", function () { proportionalMatch("max"); });
    document.getElementById("propMatchMin").addEventListener("click", function () { proportionalMatch("min"); });
}

function updatePresetButtons(val) {
    document.querySelectorAll(".preset-btn").forEach(function (btn) {
        btn.classList.toggle("active", Math.abs(parseFloat(btn.getAttribute("data-value")) - val) < 0.001);
    });
}

function showStatus(message, type) {
    var el = document.getElementById("status");
    el.textContent = message;
    el.className = "status visible " + (type || "info");
    setTimeout(function () { el.classList.remove("visible"); }, 3000);
}

function cmToPoints(cm) { return cm * CM_TO_POINTS; }
function pointsToCm(pts) { return pts / CM_TO_POINTS; }
function roundToGrid(valueCm) { return Math.round(valueCm / gridUnitCm) * gridUnitCm; }

function withSelectedShapes(minCount, callback) {
    PowerPoint.run(function (context) {
        var shapes = context.presentation.getSelectedShapes();
        shapes.load("items");
        return context.sync().then(function () {
            if (shapes.items.length < minCount) {
                showStatus(minCount <= 1 ? "Bitte Objekt(e) auswählen!" : "Bitte mindestens " + minCount + " Objekte auswählen!", "error");
                return;
            }
            return callback(context, shapes.items);
        });
    }).catch(function (error) { showStatus("Fehler: " + error.message, "error"); });
}

// ===== Helper: Check if shapes are arranged along an axis =====
// Returns true if shapes don't overlap along the given axis
// axis = "horizontal" → checks if shapes are side-by-side (no horizontal overlap)
// axis = "vertical"   → checks if shapes are stacked (no vertical overlap)
function shapesArrangedAlongAxis(shapeData, axis) {
    var sorted, i;
    if (axis === "horizontal") {
        sorted = shapeData.slice().sort(function (a, b) { return a.left - b.left; });
        for (i = 0; i < sorted.length - 1; i++) {
            var rightEdge = sorted[i].left + sorted[i].width;
            // If next shape starts before this one ends → overlap
            if (sorted[i + 1].left < rightEdge - 0.5) return false;
        }
    } else {
        sorted = shapeData.slice().sort(function (a, b) { return a.top - b.top; });
        for (i = 0; i < sorted.length - 1; i++) {
            var bottomEdge = sorted[i].top + sorted[i].height;
            if (sorted[i + 1].top < bottomEdge - 0.5) return false;
        }
    }
    return true;
}

// ===== Helper: Reposition shapes to preserve gaps along an axis =====
// shapeData = [{shape, left, top, width, height}, ...] with new dimensions already applied on shape
// axis = "horizontal" or "vertical"
// Uses ORIGINAL positions to compute gaps, then repositions with NEW sizes
function repositionToPreserveGaps(shapeData, axis) {
    var sorted, gaps, i;

    if (axis === "horizontal") {
        sorted = shapeData.slice().sort(function (a, b) { return a.left - b.left; });
        // Calculate original gaps
        gaps = [];
        for (i = 0; i < sorted.length - 1; i++) {
            var origRight = sorted[i].left + sorted[i].width; // original right edge
            gaps.push(sorted[i + 1].left - origRight);        // original gap
        }
        // Reposition: first shape stays, others shift
        for (i = 1; i < sorted.length; i++) {
            var prevNewRight = sorted[i - 1].shape.left + sorted[i - 1].shape.width;
            sorted[i].shape.left = prevNewRight + gaps[i - 1];
        }
    } else {
        sorted = shapeData.slice().sort(function (a, b) { return a.top - b.top; });
        // Calculate original gaps
        gaps = [];
        for (i = 0; i < sorted.length - 1; i++) {
            var origBottom = sorted[i].top + sorted[i].height; // original bottom edge
            gaps.push(sorted[i + 1].top - origBottom);         // original gap
        }
        // Reposition: first shape stays, others shift
        for (i = 1; i < sorted.length; i++) {
            var prevNewBottom = sorted[i - 1].shape.top + sorted[i - 1].shape.height;
            sorted[i].shape.top = prevNewBottom + gaps[i - 1];
        }
    }
}

// ===================================================================
// TAB 1: RESIZE SHAPES (spacing preserved for 2+ shapes)
// ===================================================================
function resizeShapes(dimension, deltaCm) {
    withSelectedShapes(1, function (context, items) {
        items.forEach(function (s) { s.load(["left", "top", "width", "height"]); });
        return context.sync().then(function () {
            var dp = cmToPoints(Math.abs(deltaCm));
            var grow = deltaCm > 0;

            // === Single shape: simple resize ===
            if (items.length === 1) {
                var s = items[0];
                if (dimension === "width" || dimension === "both") {
                    var nw = grow ? s.width + dp : s.width - dp;
                    if (nw >= cmToPoints(MIN_SIZE_CM)) s.width = nw;
                }
                if (dimension === "height" || dimension === "both") {
                    var nh = grow ? s.height + dp : s.height - dp;
                    if (nh >= cmToPoints(MIN_SIZE_CM)) s.height = nh;
                }
                return context.sync().then(function () {
                    showStatus((dimension === "both" ? "Breite & Höhe" : dimension === "width" ? "Breite" : "Höhe") + (grow ? " vergrößert" : " verkleinert") + " (" + Math.abs(deltaCm).toFixed(2) + " cm)", "success");
                });
            }

            // === Multiple shapes: resize + preserve spacing ===

            // Snapshot original positions & sizes
            var orig = items.map(function (s) {
                return { shape: s, left: s.left, top: s.top, width: s.width, height: s.height };
            });

            // Detect arrangement
            var isHorizontal = shapesArrangedAlongAxis(orig, "horizontal");
            var isVertical = shapesArrangedAlongAxis(orig, "vertical");

            // --- Apply width resize ---
            if (dimension === "width" || dimension === "both") {
                var wOk = true;
                for (var i = 0; i < orig.length; i++) {
                    if ((grow ? orig[i].width + dp : orig[i].width - dp) < cmToPoints(MIN_SIZE_CM)) { wOk = false; break; }
                }
                if (wOk) {
                    orig.forEach(function (o) {
                        o.shape.width = grow ? o.width + dp : o.width - dp;
                    });
                    // Reposition horizontally only if shapes are arranged side-by-side
                    if (isHorizontal) {
                        repositionToPreserveGaps(orig, "horizontal");
                    }
                }
            }

            // --- Apply height resize ---
            if (dimension === "height" || dimension === "both") {
                var hOk = true;
                for (var i = 0; i < orig.length; i++) {
                    if ((grow ? orig[i].height + dp : orig[i].height - dp) < cmToPoints(MIN_SIZE_CM)) { hOk = false; break; }
                }
                if (hOk) {
                    orig.forEach(function (o) {
                        o.shape.height = grow ? o.height + dp : o.height - dp;
                    });
                    // Reposition vertically only if shapes are stacked
                    if (isVertical) {
                        repositionToPreserveGaps(orig, "vertical");
                    }
                }
            }

            return context.sync().then(function () {
                showStatus((dimension === "both" ? "Breite & Höhe" : dimension === "width" ? "Breite" : "Höhe") + (grow ? " vergrößert" : " verkleinert") + " – Abstände beibehalten", "success");
            });
        });
    });
}

// ===================================================================
// TAB 1: PROPORTIONAL RESIZE (spacing preserved for 2+ shapes)
// ===================================================================
function proportionalResize(deltaCm) {
    withSelectedShapes(1, function (context, items) {
        items.forEach(function (s) { s.load(["left", "top", "width", "height"]); });
        return context.sync().then(function () {
            var dp = cmToPoints(Math.abs(deltaCm));
            var grow = deltaCm > 0;

            // === Single shape ===
            if (items.length === 1) {
                var s = items[0], r = s.height / s.width;
                var nw = grow ? s.width + dp : s.width - dp;
                if (nw >= cmToPoints(MIN_SIZE_CM)) {
                    var nh = nw * r;
                    if (nh >= cmToPoints(MIN_SIZE_CM)) { s.width = nw; s.height = nh; }
                }
                return context.sync().then(function () {
                    showStatus("Proportional " + (grow ? "vergrößert" : "verkleinert") + " (" + Math.abs(deltaCm).toFixed(2) + " cm)", "success");
                });
            }

            // === Multiple shapes ===

            // Snapshot originals
            var orig = items.map(function (s) {
                return { shape: s, left: s.left, top: s.top, width: s.width, height: s.height, ratio: s.height / s.width };
            });

            // Validate all can be resized
            var ok = true;
            orig.forEach(function (o) {
                var nw = grow ? o.width + dp : o.width - dp;
                if (nw < cmToPoints(MIN_SIZE_CM) || nw * o.ratio < cmToPoints(MIN_SIZE_CM)) ok = false;
            });
            if (!ok) { showStatus("Mindestgröße erreicht!", "error"); return context.sync(); }

            // Detect arrangement
            var isHorizontal = shapesArrangedAlongAxis(orig, "horizontal");
            var isVertical = shapesArrangedAlongAxis(orig, "vertical");

            // Apply proportional resize
            orig.forEach(function (o) {
                var nw = grow ? o.width + dp : o.width - dp;
                o.shape.width = nw;
                o.shape.height = nw * o.ratio;
            });

            // Reposition only along axes where shapes are actually arranged
            if (isHorizontal) {
                repositionToPreserveGaps(orig, "horizontal");
            }
            if (isVertical) {
                repositionToPreserveGaps(orig, "vertical");
            }

            return context.sync().then(function () {
                showStatus("Proportional " + (grow ? "vergrößert" : "verkleinert") + " – Abstände beibehalten", "success");
            });
        });
    });
}

// ===== TAB 2: SNAP TO GRID =====
function snapToGrid(mode) {
    withSelectedShapes(1, function (context, items) {
        items.forEach(function (s) { s.load(["left", "top", "width", "height"]); });
        return context.sync().then(function () {
            items.forEach(function (s) {
                if (mode === "position" || mode === "both") {
                    s.left = cmToPoints(roundToGrid(pointsToCm(s.left)));
                    s.top = cmToPoints(roundToGrid(pointsToCm(s.top)));
                }
                if (mode === "size" || mode === "both") {
                    var nw = roundToGrid(pointsToCm(s.width));
                    var nh = roundToGrid(pointsToCm(s.height));
                    if (nw >= MIN_SIZE_CM) s.width = cmToPoints(nw);
                    if (nh >= MIN_SIZE_CM) s.height = cmToPoints(nh);
                }
            });
            return context.sync().then(function () {
                showStatus((mode === "both" ? "Position & Größe" : mode === "position" ? "Position" : "Größe") + " am Raster ausgerichtet ✓", "success");
            });
        });
    });
}

// ===== TAB 2: SHOW SHAPE INFO =====
function showShapeInfo() {
    withSelectedShapes(1, function (context, items) {
        items.forEach(function (s) { s.load(["name", "left", "top", "width", "height"]); });
        return context.sync().then(function () {
            var el = document.getElementById("infoDisplay"), html = "";
            items.forEach(function (s, idx) {
                if (items.length > 1) html += '<div style="font-weight:700;margin-top:' + (idx > 0 ? '8' : '0') + 'px;margin-bottom:4px;">' + (s.name || 'Objekt ' + (idx + 1)) + '</div>';
                html += '<div class="info-item"><span class="info-label">Breite:</span><span class="info-value">' + pointsToCm(s.width).toFixed(2) + ' cm</span></div>';
                html += '<div class="info-item"><span class="info-label">Höhe:</span><span class="info-value">' + pointsToCm(s.height).toFixed(2) + ' cm</span></div>';
                html += '<div class="info-item"><span class="info-label">Links:</span><span class="info-value">' + pointsToCm(s.left).toFixed(2) + ' cm</span></div>';
                html += '<div class="info-item"><span class="info-label">Oben:</span><span class="info-value">' + pointsToCm(s.top).toFixed(2) + ' cm</span></div>';
            });
            el.innerHTML = html; el.classList.add("visible");
            showStatus("Objektinfo geladen ✓", "info");
        });
    });
}

// ===== TAB 3: MATCH DIMENSIONS =====
function matchDimension(dimension, mode) {
    withSelectedShapes(2, function (context, items) {
        items.forEach(function (s) { s.load(["width", "height"]); });
        return context.sync().then(function () {
            var ws = items.map(function (s) { return s.width; }), hs = items.map(function (s) { return s.height; });
            var tw = mode === "max" ? Math.max.apply(null, ws) : Math.min.apply(null, ws);
            var th = mode === "max" ? Math.max.apply(null, hs) : Math.min.apply(null, hs);
            items.forEach(function (s) {
                if (dimension === "width" || dimension === "both") s.width = tw;
                if (dimension === "height" || dimension === "both") s.height = th;
            });
            return context.sync().then(function () {
                showStatus("Alle auf " + (dimension === "both" ? "Größe" : dimension === "width" ? "Breite" : "Höhe") + " des " + (mode === "max" ? "größten" : "kleinsten") + " Objekts ✓", "success");
            });
        });
    });
}

// ===== TAB 3: PROPORTIONAL MATCH =====
function proportionalMatch(mode) {
    withSelectedShapes(2, function (context, items) {
        items.forEach(function (s) { s.load(["width", "height"]); });
        return context.sync().then(function () {
            var ws = items.map(function (s) { return s.width; });
            var tw = mode === "max" ? Math.max.apply(null, ws) : Math.min.apply(null, ws);
            items.forEach(function (s) { var r = s.height / s.width; s.width = tw; s.height = tw * r; });
            return context.sync().then(function () {
                showStatus("Proportional auf Breite des " + (mode === "max" ? "größten" : "kleinsten") + " Objekts ✓", "success");
            });
        });
    });
}
