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

    // === TAB 4: Copy shadow button ===
    document.getElementById("copyShadowText").addEventListener("click", function () { copyShadowInstructions(); });
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
// TAB 1: RESIZE SHAPES
// ===================================================================
function resizeShapes(dimension, deltaCm) {
    withSelectedShapes(1, function (context, items) {
        items.forEach(function (shape) { shape.load(["width", "height"]); });
        return context.sync().then(function () {
            var deltaPoints = cmToPoints(Math.abs(deltaCm));
            items.forEach(function (shape) {
                if (dimension === "width" || dimension === "both") {
                    var newW = deltaCm > 0 ? shape.width + deltaPoints : shape.width - deltaPoints;
                    if (newW >= cmToPoints(MIN_SIZE_CM)) shape.width = newW;
                }
                if (dimension === "height" || dimension === "both") {
                    var newH = deltaCm > 0 ? shape.height + deltaPoints : shape.height - deltaPoints;
                    if (newH >= cmToPoints(MIN_SIZE_CM)) shape.height = newH;
                }
            });
            return context.sync().then(function () {
                var action = deltaCm > 0 ? "vergrößert" : "verkleinert";
                var dimText = dimension === "both" ? "Breite & Höhe" : (dimension === "width" ? "Breite" : "Höhe");
                showStatus(dimText + " " + action + " (" + Math.abs(deltaCm).toFixed(2) + " cm)", "success");
            });
        });
    });
}

// ===================================================================
// TAB 1: PROPORTIONAL RESIZE
// ===================================================================
function proportionalResize(deltaCm) {
    withSelectedShapes(1, function (context, items) {
        items.forEach(function (shape) { shape.load(["width", "height"]); });
        return context.sync().then(function () {
            var deltaPoints = cmToPoints(Math.abs(deltaCm));
            items.forEach(function (shape) {
                var ratio = shape.height / shape.width;
                var newW = deltaCm > 0 ? shape.width + deltaPoints : shape.width - deltaPoints;
                if (newW >= cmToPoints(MIN_SIZE_CM)) {
                    var newH = newW * ratio;
                    if (newH >= cmToPoints(MIN_SIZE_CM)) {
                        shape.width = newW;
                        shape.height = newH;
                    }
                }
            });
            return context.sync().then(function () {
                var action = deltaCm > 0 ? "vergrößert" : "verkleinert";
                showStatus("Proportional " + action + " (" + Math.abs(deltaCm).toFixed(2) + " cm)", "success");
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
            items.forEach(function (shape) {
                if (dimension === "width" || dimension === "both") shape.width = targetW;
                if (dimension === "height" || dimension === "both") shape.height = targetH;
            });
            return context.sync().then(function () {
                var modeText = mode === "max" ? "größtes" : "kleinstes";
                var dimText = dimension === "both" ? "Größe" : (dimension === "width" ? "Breite" : "Höhe");
                showStatus("Alle auf " + dimText + " des " + modeText + "n Objekts ✓", "success");
            });
        });
    });
}

// ===================================================================
// TAB 3: PROPORTIONAL MATCH
// ===================================================================
function proportionalMatch(mode) {
    withSelectedShapes(2, function (context, items) {
        items.forEach(function (shape) { shape.load(["width", "height"]); });
        return context.sync().then(function () {
            var widths = items.map(function (s) { return s.width; });
            var targetW = mode === "max" ? Math.max.apply(null, widths) : Math.min.apply(null, widths);
            items.forEach(function (shape) {
                var ratio = shape.height / shape.width;
                shape.width = targetW;
                shape.height = targetW * ratio;
            });
            return context.sync().then(function () {
                var modeText = mode === "max" ? "größten" : "kleinsten";
                showStatus("Proportional auf Breite des " + modeText + " Objekts ✓", "success");
            });
        });
    });
}

// ===================================================================
// TAB 4: COPY SHADOW INSTRUCTIONS TO CLIPBOARD
// ===================================================================
function copyShadowInstructions() {
    var text = "Schatten-Einstellungen (Droege Group Standard):\n" +
        "─────────────────────────────────\n" +
        "Typ: Offset unten rechts\n" +
        "Farbe: Schwarz 50% (= #808080)\n" +
        "Transparenz: 30%\n" +
        "Größe: 100%\n" +
        "Weichzeichnen: 0,5 Pt.\n" +
        "Winkel: 45°\n" +
        "Abstand: 1 Pt.\n" +
        "─────────────────────────────────\n" +
        "Pfad: Objekt rechtsklick → Form formatieren → Effekte → Schatten";
    navigator.clipboard.writeText(text).then(function () {
        showStatus("Schatten-Werte in Zwischenablage kopiert ✓", "success");
    }).catch(function () {
        showStatus("Kopieren fehlgeschlagen – bitte manuell kopieren", "error");
    });
}
