/* ===== Grid Resize Tool - JavaScript ===== */

// Constants
var CM_TO_POINTS = 28.3465;
var MIN_SIZE_CM = 0.1;

// State
var gridUnitCm = 0.21;

// ===== Office Init =====
Office.onReady(function (info) {
    if (info.host === Office.HostType.PowerPoint) {
        initUI();
    }
});

function initUI() {
    // Grid unit input
    var gridInput = document.getElementById("gridUnit");
    gridInput.addEventListener("change", function () {
        var val = parseFloat(this.value);
        if (!isNaN(val) && val > 0) {
            gridUnitCm = val;
            updatePresetButtons(val);
            showStatus("Rastereinheit: " + val.toFixed(2) + " cm", "info");
        }
    });

    // Preset buttons
    var presets = document.querySelectorAll(".preset-btn");
    presets.forEach(function (btn) {
        btn.addEventListener("click", function () {
            var val = parseFloat(this.getAttribute("data-value"));
            gridUnitCm = val;
            gridInput.value = val;
            updatePresetButtons(val);
            showStatus("Rastereinheit: " + val.toFixed(2) + " cm", "info");
        });
    });

    // Tab navigation
    var tabBtns = document.querySelectorAll(".tab-btn");
    tabBtns.forEach(function (btn) {
        btn.addEventListener("click", function () {
            var tabId = this.getAttribute("data-tab");
            tabBtns.forEach(function (b) { b.classList.remove("active"); });
            document.querySelectorAll(".tab-content").forEach(function (t) { t.classList.remove("active"); });
            this.classList.add("active");
            document.getElementById(tabId).classList.add("active");
        });
    });

    // === TAB 1: Resize buttons ===
    document.getElementById("shrinkWidth").addEventListener("click", function () { resizeShapes("width", -gridUnitCm); });
    document.getElementById("growWidth").addEventListener("click", function () { resizeShapes("width", gridUnitCm); });
    document.getElementById("shrinkHeight").addEventListener("click", function () { resizeShapes("height", -gridUnitCm); });
    document.getElementById("growHeight").addEventListener("click", function () { resizeShapes("height", gridUnitCm); });
    document.getElementById("shrinkBoth").addEventListener("click", function () { resizeShapes("both", -gridUnitCm); });
    document.getElementById("growBoth").addEventListener("click", function () { resizeShapes("both", gridUnitCm); });
    document.getElementById("propShrink").addEventListener("click", function () { proportionalResize(-gridUnitCm); });
    document.getElementById("propGrow").addEventListener("click", function () { proportionalResize(gridUnitCm); });

    // === TAB 2: Snap buttons ===
    document.getElementById("snapPosition").addEventListener("click", function () { snapToGrid("position"); });
    document.getElementById("snapSize").addEventListener("click", function () { snapToGrid("size"); });
    document.getElementById("snapBoth").addEventListener("click", function () { snapToGrid("both"); });
    document.getElementById("showInfo").addEventListener("click", function () { showShapeInfo(); });

    // === TAB 3: Align buttons ===
    document.getElementById("matchWidthMax").addEventListener("click", function () { matchDimension("width", "max"); });
    document.getElementById("matchWidthMin").addEventListener("click", function () { matchDimension("width", "min"); });
    document.getElementById("matchHeightMax").addEventListener("click", function () { matchDimension("height", "max"); });
    document.getElementById("matchHeightMin").addEventListener("click", function () { matchDimension("height", "min"); });
    document.getElementById("matchBothMax").addEventListener("click", function () { matchDimension("both", "max"); });
    document.getElementById("matchBothMin").addEventListener("click", function () { matchDimension("both", "min"); });
    document.getElementById("propMatchMax").addEventListener("click", function () { proportionalMatch("max"); });
    document.getElementById("propMatchMin").addEventListener("click", function () { proportionalMatch("min"); });
}

// ===== Preset Button Highlight =====
function updatePresetButtons(val) {
    document.querySelectorAll(".preset-btn").forEach(function (btn) {
        var bv = parseFloat(btn.getAttribute("data-value"));
        btn.classList.toggle("active", Math.abs(bv - val) < 0.001);
    });
}

// ===== Status Messages =====
function showStatus(message, type) {
    var el = document.getElementById("status");
    el.textContent = message;
    el.className = "status visible " + (type || "info");
    setTimeout(function () {
        el.classList.remove("visible");
    }, 3000);
}

// ===== Helper Functions =====
function cmToPoints(cm) { return cm * CM_TO_POINTS; }
function pointsToCm(pts) { return pts / CM_TO_POINTS; }
function roundToGrid(valueCm) { return Math.round(valueCm / gridUnitCm) * gridUnitCm; }

// ===== Get Selected Shapes Helper =====
function withSelectedShapes(minCount, callback) {
    PowerPoint.run(function (context) {
        var shapes = context.presentation.getSelectedShapes();
        shapes.load("items");
        return context.sync().then(function () {
            if (shapes.items.length < minCount) {
                var msg = minCount <= 1 ? "Bitte Objekt(e) auswählen!" : "Bitte mindestens " + minCount + " Objekte auswählen!";
                showStatus(msg, "error");
                return;
            }
            return callback(context, shapes.items);
        });
    }).catch(function (error) {
        showStatus("Fehler: " + error.message, "error");
    });
}

// ===================================================================
// TAB 1: RESIZE SHAPES (with spacing preservation for multiple shapes)
// ===================================================================
function resizeShapes(dimension, deltaCm) {
    withSelectedShapes(1, function (context, items) {
        items.forEach(function (shape) { shape.load(["left", "top", "width", "height"]); });
        return context.sync().then(function () {
            var deltaPoints = cmToPoints(Math.abs(deltaCm));
            var growing = deltaCm > 0;

            // === Single shape: simple resize (no spacing logic needed) ===
            if (items.length === 1) {
                var shape = items[0];
                if (dimension === "width" || dimension === "both") {
                    var newW = growing ? shape.width + deltaPoints : shape.width - deltaPoints;
                    if (newW >= cmToPoints(MIN_SIZE_CM)) shape.width = newW;
                }
                if (dimension === "height" || dimension === "both") {
                    var newH = growing ? shape.height + deltaPoints : shape.height - deltaPoints;
                    if (newH >= cmToPoints(MIN_SIZE_CM)) shape.height = newH;
                }
                return context.sync().then(function () {
                    var action = growing ? "vergrößert" : "verkleinert";
                    var dimText = dimension === "both" ? "Breite & Höhe" : (dimension === "width" ? "Breite" : "Höhe");
                    showStatus(dimText + " " + action + " (" + Math.abs(deltaCm).toFixed(2) + " cm)", "success");
                });
            }

            // === Multiple shapes: resize + preserve spacing ===

            // Horizontal spacing preservation (for width changes)
            if (dimension === "width" || dimension === "both") {
                // Sort by left position
                var hSorted = items.slice().sort(function (a, b) { return a.left - b.left; });

                // Calculate gaps between shapes (right edge of shape[i] to left edge of shape[i+1])
                var hGaps = [];
                for (var i = 0; i < hSorted.length - 1; i++) {
                    var rightEdge = hSorted[i].left + hSorted[i].width;
                    var gap = hSorted[i + 1].left - rightEdge;
                    hGaps.push(gap);
                }

                // Resize all shapes horizontally
                var hResizeValid = true;
                for (var i = 0; i < hSorted.length; i++) {
                    var newW = growing ? hSorted[i].width + deltaPoints : hSorted[i].width - deltaPoints;
                    if (newW < cmToPoints(MIN_SIZE_CM)) {
                        hResizeValid = false;
                        break;
                    }
                }

                if (hResizeValid) {
                    for (var i = 0; i < hSorted.length; i++) {
                        var newW = growing ? hSorted[i].width + deltaPoints : hSorted[i].width - deltaPoints;
                        hSorted[i].width = newW;
                    }

                    // Reposition shapes to preserve gaps
                    // First shape keeps its position, subsequent shapes shift
                    for (var i = 1; i < hSorted.length; i++) {
                        var prevRightEdge = hSorted[i - 1].left + hSorted[i - 1].width;
                        hSorted[i].left = prevRightEdge + hGaps[i - 1];
                    }
                }
            }

            // Vertical spacing preservation (for height changes)
            if (dimension === "height" || dimension === "both") {
                // Sort by top position
                var vSorted = items.slice().sort(function (a, b) { return a.top - b.top; });

                // Calculate gaps between shapes (bottom edge of shape[i] to top edge of shape[i+1])
                var vGaps = [];
                for (var i = 0; i < vSorted.length - 1; i++) {
                    var bottomEdge = vSorted[i].top + vSorted[i].height;
                    var gap = vSorted[i + 1].top - bottomEdge;
                    vGaps.push(gap);
                }

                // Resize all shapes vertically
                var vResizeValid = true;
                for (var i = 0; i < vSorted.length; i++) {
                    var newH = growing ? vSorted[i].height + deltaPoints : vSorted[i].height - deltaPoints;
                    if (newH < cmToPoints(MIN_SIZE_CM)) {
                        vResizeValid = false;
                        break;
                    }
                }

                if (vResizeValid) {
                    for (var i = 0; i < vSorted.length; i++) {
                        var newH = growing ? vSorted[i].height + deltaPoints : vSorted[i].height - deltaPoints;
                        vSorted[i].height = newH;
                    }

                    // Reposition shapes to preserve gaps
                    for (var i = 1; i < vSorted.length; i++) {
                        var prevBottomEdge = vSorted[i - 1].top + vSorted[i - 1].height;
                        vSorted[i].top = prevBottomEdge + vGaps[i - 1];
                    }
                }
            }

            return context.sync().then(function () {
                var action = growing ? "vergrößert" : "verkleinert";
                var dimText = dimension === "both" ? "Breite & Höhe" : (dimension === "width" ? "Breite" : "Höhe");
                showStatus(dimText + " " + action + " – Abstände beibehalten (" + Math.abs(deltaCm).toFixed(2) + " cm)", "success");
            });
        });
    });
}

// ===================================================================
// TAB 1: PROPORTIONAL RESIZE (with spacing preservation)
// ===================================================================
function proportionalResize(deltaCm) {
    withSelectedShapes(1, function (context, items) {
        items.forEach(function (shape) { shape.load(["left", "top", "width", "height"]); });
        return context.sync().then(function () {
            var deltaPoints = cmToPoints(Math.abs(deltaCm));
            var growing = deltaCm > 0;

            // === Single shape: simple proportional resize ===
            if (items.length === 1) {
                var shape = items[0];
                var ratio = shape.height / shape.width;
                var newW = growing ? shape.width + deltaPoints : shape.width - deltaPoints;
                if (newW >= cmToPoints(MIN_SIZE_CM)) {
                    var newH = newW * ratio;
                    if (newH >= cmToPoints(MIN_SIZE_CM)) {
                        shape.width = newW;
                        shape.height = newH;
                    }
                }
                return context.sync().then(function () {
                    var action = growing ? "vergrößert" : "verkleinert";
                    showStatus("Proportional " + action + " (" + Math.abs(deltaCm).toFixed(2) + " cm)", "success");
                });
            }

            // === Multiple shapes: proportional resize + preserve spacing ===

            // Store original values and compute ratios
            var originals = items.map(function (shape) {
                return {
                    shape: shape,
                    left: shape.left,
                    top: shape.top,
                    width: shape.width,
                    height: shape.height,
                    ratio: shape.height / shape.width
                };
            });

            // Check if all shapes can be resized
            var allValid = true;
            originals.forEach(function (o) {
                var newW = growing ? o.width + deltaPoints : o.width - deltaPoints;
                var newH = newW * o.ratio;
                if (newW < cmToPoints(MIN_SIZE_CM) || newH < cmToPoints(MIN_SIZE_CM)) {
                    allValid = false;
                }
            });

            if (!allValid) {
                showStatus("Mindestgröße erreicht!", "error");
                return context.sync();
            }

            // --- Horizontal spacing ---
            var hSorted = originals.slice().sort(function (a, b) { return a.left - b.left; });
            var hGaps = [];
            for (var i = 0; i < hSorted.length - 1; i++) {
                var rightEdge = hSorted[i].left + hSorted[i].width;
                hGaps.push(hSorted[i + 1].left - rightEdge);
            }

            // --- Vertical spacing ---
            var vSorted = originals.slice().sort(function (a, b) { return a.top - b.top; });
            var vGaps = [];
            for (var i = 0; i < vSorted.length - 1; i++) {
                var bottomEdge = vSorted[i].top + vSorted[i].height;
                vGaps.push(vSorted[i + 1].top - bottomEdge);
            }

            // Resize all shapes
            originals.forEach(function (o) {
                var newW = growing ? o.width + deltaPoints : o.width - deltaPoints;
                var newH = newW * o.ratio;
                o.shape.width = newW;
                o.shape.height = newH;
                o.newWidth = newW;
                o.newHeight = newH;
            });

            // Reposition horizontally (preserve gaps)
            for (var i = 1; i < hSorted.length; i++) {
                var prevRight = hSorted[i - 1].shape.left + hSorted[i - 1].newWidth;
                hSorted[i].shape.left = prevRight + hGaps[i - 1];
            }

            // Reposition vertically (preserve gaps)
            for (var i = 1; i < vSorted.length; i++) {
                var prevBottom = vSorted[i - 1].shape.top + vSorted[i - 1].newHeight;
                vSorted[i].shape.top = prevBottom + vGaps[i - 1];
            }

            return context.sync().then(function () {
                var action = growing ? "vergrößert" : "verkleinert";
                showStatus("Proportional " + action + " – Abstände beibehalten (" + Math.abs(deltaCm).toFixed(2) + " cm)", "success");
            });
        });
    });
}

// ===================================================================
// TAB 2: SNAP TO GRID
// ===================================================================
function snapToGrid(mode) {
    withSelectedShapes(1, function (context, items) {
        items.forEach(function (shape) { shape.load(["left", "top", "width", "height"]); });
        return context.sync().then(function () {
            items.forEach(function (shape) {
                if (mode === "position" || mode === "both") {
                    shape.left = cmToPoints(roundToGrid(pointsToCm(shape.left)));
                    shape.top = cmToPoints(roundToGrid(pointsToCm(shape.top)));
                }
                if (mode === "size" || mode === "both") {
                    var newW = roundToGrid(pointsToCm(shape.width));
                    var newH = roundToGrid(pointsToCm(shape.height));
                    if (newW >= MIN_SIZE_CM) shape.width = cmToPoints(newW);
                    if (newH >= MIN_SIZE_CM) shape.height = cmToPoints(newH);
                }
            });
            return context.sync().then(function () {
                var modeText = mode === "both" ? "Position & Größe" : (mode === "position" ? "Position" : "Größe");
                showStatus(modeText + " am Raster ausgerichtet ✓", "success");
            });
        });
    });
}

// ===================================================================
// TAB 2: SHOW SHAPE INFO
// ===================================================================
function showShapeInfo() {
    withSelectedShapes(1, function (context, items) {
        items.forEach(function (shape) { shape.load(["name", "left", "top", "width", "height"]); });
        return context.sync().then(function () {
            var infoEl = document.getElementById("infoDisplay");
            var html = "";
            items.forEach(function (shape, idx) {
                var w = pointsToCm(shape.width).toFixed(2);
                var h = pointsToCm(shape.height).toFixed(2);
                var l = pointsToCm(shape.left).toFixed(2);
                var t = pointsToCm(shape.top).toFixed(2);
                if (items.length > 1) {
                    html += '<div style="font-weight:700;margin-top:' + (idx > 0 ? '8' : '0') + 'px;margin-bottom:4px;">' + (shape.name || 'Objekt ' + (idx + 1)) + '</div>';
                }
                html += '<div class="info-item"><span class="info-label">Breite:</span><span class="info-value">' + w + ' cm</span></div>';
                html += '<div class="info-item"><span class="info-label">Höhe:</span><span class="info-value">' + h + ' cm</span></div>';
                html += '<div class="info-item"><span class="info-label">Links:</span><span class="info-value">' + l + ' cm</span></div>';
                html += '<div class="info-item"><span class="info-label">Oben:</span><span class="info-value">' + t + ' cm</span></div>';
            });
            infoEl.innerHTML = html;
            infoEl.classList.add("visible");
            showStatus("Objektinfo geladen ✓", "info");
        });
    });
}

// ===================================================================
// TAB 3: MATCH DIMENSIONS
// ===================================================================
function matchDimension(dimension, mode) {
    withSelectedShapes(2, function (context, items) {
        items.forEach(function (shape) { shape.load(["width", "height"]); });
        return context.sync().then(function () {
            var widths = items.map(function (s) { return s.width; });
            var heights = items.map(function (s) { return s.height; });
            var targetW = mode === "max" ? Math.max.apply(null, widths) : Math.min.apply(null, widths);
            var targetH = mode === "max" ? Math.max.apply(null, heights) : Math.min.apply(null, heights);

            items.forEach(function (s) {
                if (dimension === "width" || dimension === "both") { s.width = targetW; }
                if (dimension === "height" || dimension === "both") { s.height = targetH; }
            });

            return context.sync().then(function () {
                var modeText = mode === "max" ? "größten" : "kleinsten";
                var dimText = dimension === "both" ? "Größe" : (dimension === "width" ? "Breite" : "Höhe");
                showStatus("Alle auf " + dimText + " des " + modeText + " Objekts ✔", "success");
            });
        });
    });
}

// ===================================================================
// TAB 3: PROPORTIONAL MATCH
// ===================================================================
function proportionalMatch(mode) {
    withSelectedShapes(2, function (context, items) {
        items.forEach(function (s) { s.load(["width", "height"]); });
        return context.sync().then(function () {
            var widths = items.map(function (s) { return s.width; });
            var targetW = mode === "max" ? Math.max.apply(null, widths) : Math.min.apply(null, widths);
            items.forEach(function (s) {
                var ratio = s.height / s.width;
                s.width = targetW;
                s.height = targetW * ratio;
            });
            return context.sync().then(function () {
                var modeText = mode === "max" ? "größten" : "kleinsten";
                showStatus("Proportional auf Breite des " + modeText + " Objekts ✔", "success");
            });
        });
    });
}

// ===================================================================
// EXTRAS: SET DROEGE SLIDE SIZE
// ===================================================================
function setDroegeSlideSize() {
    // Droege-Format: 27,724 × 19,297 cm
    // Breite: 27.724 × 28.3465 = 785.67 pt → gerundet 786 pt
    // Höhe:  19.297 cm = 547 pt
    var targetWidth = 786;
    var targetHeight = 547;

    PowerPoint.run(function (context) {
        var pageSetup = context.presentation.pageSetup;
        pageSetup.load(["slideWidth", "slideHeight"]);
        return context.sync()
            .then(function () {
                pageSetup.slideWidth = targetWidth;
                return context.sync();
            })
            .then(function () {
                pageSetup.slideHeight = targetHeight;
                return context.sync();
            })
            .then(function () {
                showStatus("Papierformat gesetzt: 27,724 × 19,297 cm (786 × 547 pt) ✔", "success");
            });
    }).catch(function (error) {
        showStatus("Fehler: " + error.message, "error");
    });
}
