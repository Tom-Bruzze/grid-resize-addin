/* ===== Grid Resize Tool Enhanced JavaScript ===== */

// Constants
const CM_TO_POINTS = 28.3465;
const MIN_SIZE_CM = 0.1;

// State
let gridUnitCm = 0.21;

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
    document.getElementById("shrinkWidth").addEventListener("click", function () {
        resizeShapes("width", -gridUnitCm);
    });
    document.getElementById("growWidth").addEventListener("click", function () {
        resizeShapes("width", gridUnitCm);
    });
    document.getElementById("shrinkHeight").addEventListener("click", function () {
        resizeShapes("height", -gridUnitCm);
    });
    document.getElementById("growHeight").addEventListener("click", function () {
        resizeShapes("height", gridUnitCm);
    });
    document.getElementById("shrinkBoth").addEventListener("click", function () {
        resizeShapes("both", -gridUnitCm);
    });
    document.getElementById("growBoth").addEventListener("click", function () {
        resizeShapes("both", gridUnitCm);
    });
    document.getElementById("propShrink").addEventListener("click", function () {
        proportionalResize(-gridUnitCm);
    });
    document.getElementById("propGrow").addEventListener("click", function () {
        proportionalResize(gridUnitCm);
    });

    // === TAB 2: Snap buttons ===
    document.getElementById("snapPosition").addEventListener("click", function () {
        snapToGrid("position");
    });
    document.getElementById("snapSize").addEventListener("click", function () {
        snapToGrid("size");
    });
    document.getElementById("snapBoth").addEventListener("click", function () {
        snapToGrid("both");
    });
    document.getElementById("showInfo").addEventListener("click", function () {
        showShapeInfo();
    });

    // === TAB 3: Align buttons ===
    document.getElementById("matchWidthMax").addEventListener("click", function () {
        matchDimension("width", "max");
    });
    document.getElementById("matchWidthMin").addEventListener("click", function () {
        matchDimension("width", "min");
    });
    document.getElementById("matchHeightMax").addEventListener("click", function () {
        matchDimension("height", "max");
    });
    document.getElementById("matchHeightMin").addEventListener("click", function () {
        matchDimension("height", "min");
    });
    document.getElementById("matchBothMax").addEventListener("click", function () {
        matchDimension("both", "max");
    });
    document.getElementById("matchBothMin").addEventListener("click", function () {
        matchDimension("both", "min");
    });
    document.getElementById("propMatchMax").addEventListener("click", function () {
        proportionalMatch("max");
    });
    document.getElementById("propMatchMin").addEventListener("click", function () {
        proportionalMatch("min");
    });

    // === TAB 4: Extras buttons ===
    document.getElementById("addShadow").addEventListener("click", function () {
        applyShadow();
    });
    document.getElementById("removeShadow").addEventListener("click", function () {
        removeShadow();
    });
}

// ===== Preset Button Highlight =====
function updatePresetButtons(val) {
    document.querySelectorAll(".preset-btn").forEach(function (btn) {
        var bv = parseFloat(btn.getAttribute("data-value"));
        if (Math.abs(bv - val) < 0.001) {
            btn.classList.add("active");
        } else {
            btn.classList.remove("active");
        }
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

// ===== Helper: Points <-> CM =====
function cmToPoints(cm) {
    return cm * CM_TO_POINTS;
}

function pointsToCm(pts) {
    return pts / CM_TO_POINTS;
}

function roundToGrid(valueCm) {
    return Math.round(valueCm / gridUnitCm) * gridUnitCm;
}

// ===================================================================
// TAB 1: RESIZE SHAPES
// ===================================================================
function resizeShapes(dimension, deltaCm) {
    Office.context.document.getSelectedDataAsync(
        Office.CoercionType.SlideRange,
        function (asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                showStatus("Fehler: " + asyncResult.error.message, "error");
                return;
            }

            PowerPoint.run(function (context) {
                var shapes = context.presentation.getSelectedShapes();
                shapes.load("items");

                return context.sync().then(function () {
                    if (shapes.items.length === 0) {
                        showStatus("Bitte Objekt(e) auswählen!", "error");
                        return;
                    }

                    var promises = shapes.items.map(function (shape) {
                        shape.load(["width", "height"]);
                        return context.sync().then(function () {
                            var deltaPoints = cmToPoints(Math.abs(deltaCm));

                            if (dimension === "width" || dimension === "both") {
                                var newW = deltaCm > 0
                                    ? shape.width + deltaPoints
                                    : shape.width - deltaPoints;
                                if (newW >= cmToPoints(MIN_SIZE_CM)) {
                                    shape.width = newW;
                                }
                            }

                            if (dimension === "height" || dimension === "both") {
                                var newH = deltaCm > 0
                                    ? shape.height + deltaPoints
                                    : shape.height - deltaPoints;
                                if (newH >= cmToPoints(MIN_SIZE_CM)) {
                                    shape.height = newH;
                                }
                            }

                            return context.sync();
                        });
                    });

                    return Promise.all(promises).then(function () {
                        var action = deltaCm > 0 ? "vergrößert" : "verkleinert";
                        var dimText = dimension === "both" ? "Breite & Höhe" : (dimension === "width" ? "Breite" : "Höhe");
                        showStatus(dimText + " " + action + " (" + Math.abs(deltaCm).toFixed(2) + " cm)", "success");
                    });
                });
            }).catch(function (error) {
                showStatus("Fehler: " + error.message, "error");
            });
        }
    );
}

// ===================================================================
// TAB 1: PROPORTIONAL RESIZE
// ===================================================================
function proportionalResize(deltaCm) {
    Office.context.document.getSelectedDataAsync(
        Office.CoercionType.SlideRange,
        function (asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                showStatus("Fehler: " + asyncResult.error.message, "error");
                return;
            }

            PowerPoint.run(function (context) {
                var shapes = context.presentation.getSelectedShapes();
                shapes.load("items");

                return context.sync().then(function () {
                    if (shapes.items.length === 0) {
                        showStatus("Bitte Objekt(e) auswählen!", "error");
                        return;
                    }

                    var promises = shapes.items.map(function (shape) {
                        shape.load(["width", "height"]);
                        return context.sync().then(function () {
                            var ratio = shape.height / shape.width;
                            var deltaPoints = cmToPoints(Math.abs(deltaCm));
                            var newW = deltaCm > 0
                                ? shape.width + deltaPoints
                                : shape.width - deltaPoints;

                            if (newW >= cmToPoints(MIN_SIZE_CM)) {
                                var newH = newW * ratio;
                                if (newH >= cmToPoints(MIN_SIZE_CM)) {
                                    shape.width = newW;
                                    shape.height = newH;
                                }
                            }
                            return context.sync();
                        });
                    });

                    return Promise.all(promises).then(function () {
                        var action = deltaCm > 0 ? "vergrößert" : "verkleinert";
                        showStatus("Proportional " + action + " (" + Math.abs(deltaCm).toFixed(2) + " cm)", "success");
                    });
                });
            }).catch(function (error) {
                showStatus("Fehler: " + error.message, "error");
            });
        }
    );
}

// ===================================================================
// TAB 2: SNAP TO GRID
// ===================================================================
function snapToGrid(mode) {
    Office.context.document.getSelectedDataAsync(
        Office.CoercionType.SlideRange,
        function (asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                showStatus("Fehler: " + asyncResult.error.message, "error");
                return;
            }

            PowerPoint.run(function (context) {
                var shapes = context.presentation.getSelectedShapes();
                shapes.load("items");

                return context.sync().then(function () {
                    if (shapes.items.length === 0) {
                        showStatus("Bitte Objekt(e) auswählen!", "error");
                        return;
                    }

                    var promises = shapes.items.map(function (shape) {
                        shape.load(["left", "top", "width", "height"]);
                        return context.sync().then(function () {
                            if (mode === "position" || mode === "both") {
                                var leftCm = pointsToCm(shape.left);
                                var topCm = pointsToCm(shape.top);
                                shape.left = cmToPoints(roundToGrid(leftCm));
                                shape.top = cmToPoints(roundToGrid(topCm));
                            }

                            if (mode === "size" || mode === "both") {
                                var widthCm = pointsToCm(shape.width);
                                var heightCm = pointsToCm(shape.height);
                                var newW = roundToGrid(widthCm);
                                var newH = roundToGrid(heightCm);
                                if (newW >= MIN_SIZE_CM) shape.width = cmToPoints(newW);
                                if (newH >= MIN_SIZE_CM) shape.height = cmToPoints(newH);
                            }

                            return context.sync();
                        });
                    });

                    return Promise.all(promises).then(function () {
                        var modeText = mode === "both" ? "Position & Größe" : (mode === "position" ? "Position" : "Größe");
                        showStatus(modeText + " am Raster ausgerichtet ✓", "success");
                    });
                });
            }).catch(function (error) {
                showStatus("Fehler: " + error.message, "error");
            });
        }
    );
}

// ===================================================================
// TAB 2: SHOW SHAPE INFO
// ===================================================================
function showShapeInfo() {
    Office.context.document.getSelectedDataAsync(
        Office.CoercionType.SlideRange,
        function (asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                showStatus("Fehler: " + asyncResult.error.message, "error");
                return;
            }

            PowerPoint.run(function (context) {
                var shapes = context.presentation.getSelectedShapes();
                shapes.load("items");

                return context.sync().then(function () {
                    if (shapes.items.length === 0) {
                        showStatus("Bitte Objekt(e) auswählen!", "error");
                        return;
                    }

                    var loadPromises = shapes.items.map(function (shape) {
                        shape.load(["name", "left", "top", "width", "height"]);
                        return context.sync();
                    });

                    return Promise.all(loadPromises).then(function () {
                        var infoEl = document.getElementById("infoDisplay");
                        var html = "";

                        shapes.items.forEach(function (shape, idx) {
                            var w = pointsToCm(shape.width).toFixed(2);
                            var h = pointsToCm(shape.height).toFixed(2);
                            var l = pointsToCm(shape.left).toFixed(2);
                            var t = pointsToCm(shape.top).toFixed(2);

                            if (shapes.items.length > 1) {
                                html += "<div style='font-weight:700;margin-top:" + (idx > 0 ? "8" : "0") + "px;margin-bottom:4px;'>" + (shape.name || "Objekt " + (idx + 1)) + "</div>";
                            }

                            html += "<div class='info-item'><span class='info-label'>Breite:</span><span class='info-value'>" + w + " cm</span></div>";
                            html += "<div class='info-item'><span class='info-label'>Höhe:</span><span class='info-value'>" + h + " cm</span></div>";
                            html += "<div class='info-item'><span class='info-label'>Links:</span><span class='info-value'>" + l + " cm</span></div>";
                            html += "<div class='info-item'><span class='info-label'>Oben:</span><span class='info-value'>" + t + " cm</span></div>";
                        });

                        infoEl.innerHTML = html;
                        infoEl.classList.add("visible");
                        showStatus("Objektinfo geladen ✓", "info");
                    });
                });
            }).catch(function (error) {
                showStatus("Fehler: " + error.message, "error");
            });
        }
    );
}

// ===================================================================
// TAB 3: MATCH DIMENSIONS
// ===================================================================
function matchDimension(dimension, mode) {
    Office.context.document.getSelectedDataAsync(
        Office.CoercionType.SlideRange,
        function (asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                showStatus("Fehler: " + asyncResult.error.message, "error");
                return;
            }

            PowerPoint.run(function (context) {
                var shapes = context.presentation.getSelectedShapes();
                shapes.load("items");

                return context.sync().then(function () {
                    if (shapes.items.length < 2) {
                        showStatus("Bitte mindestens 2 Objekte auswählen!", "error");
                        return;
                    }

                    // Load all dimensions
                    shapes.items.forEach(function (shape) {
                        shape.load(["width", "height"]);
                    });

                    return context.sync().then(function () {
                        // Find target values
                        var widths = shapes.items.map(function (s) { return s.width; });
                        var heights = shapes.items.map(function (s) { return s.height; });

                        var targetW, targetH;

                        if (mode === "max") {
                            targetW = Math.max.apply(null, widths);
                            targetH = Math.max.apply(null, heights);
                        } else {
                            targetW = Math.min.apply(null, widths);
                            targetH = Math.min.apply(null, heights);
                        }

                        // Apply
                        shapes.items.forEach(function (shape) {
                            if (dimension === "width" || dimension === "both") {
                                shape.width = targetW;
                            }
                            if (dimension === "height" || dimension === "both") {
                                shape.height = targetH;
                            }
                        });

                        return context.sync().then(function () {
                            var modeText = mode === "max" ? "größtes" : "kleinstes";
                            var dimText = dimension === "both" ? "Größe" : (dimension === "width" ? "Breite" : "Höhe");
                            showStatus("Alle auf " + dimText + " des " + modeText + "n Objekts ✓", "success");
                        });
                    });
                });
            }).catch(function (error) {
                showStatus("Fehler: " + error.message, "error");
            });
        }
    );
}

// ===================================================================
// TAB 3: PROPORTIONAL MATCH
// ===================================================================
function proportionalMatch(mode) {
    Office.context.document.getSelectedDataAsync(
        Office.CoercionType.SlideRange,
        function (asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                showStatus("Fehler: " + asyncResult.error.message, "error");
                return;
            }

            PowerPoint.run(function (context) {
                var shapes = context.presentation.getSelectedShapes();
                shapes.load("items");

                return context.sync().then(function () {
                    if (shapes.items.length < 2) {
                        showStatus("Bitte mindestens 2 Objekte auswählen!", "error");
                        return;
                    }

                    // Load all dimensions
                    shapes.items.forEach(function (shape) {
                        shape.load(["width", "height"]);
                    });

                    return context.sync().then(function () {
                        var widths = shapes.items.map(function (s) { return s.width; });
                        var targetW;

                        if (mode === "max") {
                            targetW = Math.max.apply(null, widths);
                        } else {
                            targetW = Math.min.apply(null, widths);
                        }

                        // Apply proportional scaling
                        shapes.items.forEach(function (shape) {
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
            }).catch(function (error) {
                showStatus("Fehler: " + error.message, "error");
            });
        }
    );
}

// ===================================================================
// TAB 4: SHADOW - ADD SHADOW
// ===================================================================
// Shadow settings:
//   Offset: unten rechts (bottom-right)
//   Transparenz: 30% (=> alpha/opacity = 70% => 0.7)
//   Größe: 100%
//   Weichzeichnen (blur): 0.5 Pt.
//   Winkel: 45°
//   Abstand: 1 Pt.
//   Farbe: Schwarz 50% => #808080 (50% between black and white)
// ===================================================================
function applyShadow() {
    Office.context.document.getSelectedDataAsync(
        Office.CoercionType.SlideRange,
        function (asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                showStatus("Fehler: " + asyncResult.error.message, "error");
                return;
            }

            PowerPoint.run(function (context) {
                var shapes = context.presentation.getSelectedShapes();
                shapes.load("items");

                return context.sync().then(function () {
                    if (shapes.items.length === 0) {
                        showStatus("Bitte Objekt(e) auswählen!", "error");
                        return;
                    }

                    // Shadow parameters
                    // Winkel 45° + Abstand 1 Pt. => offset calculation:
                    // offsetX = distance * cos(angle) = 1 * cos(45°) = 0.7071 Pt.
                    // offsetY = distance * sin(angle) = 1 * sin(45°) = 0.7071 Pt.
                    // "Offset unten rechts" => positive X and Y
                    var angle = 45;
                    var distance = 1; // 1 Pt.
                    var blur = 0.5;   // 0.5 Pt.
                    var transparency = 0.30; // 30%
                    // Schwarz 50% = #808080
                    var shadowColor = "#808080";

                    var promises = shapes.items.map(function (shape) {
                        shape.load("shadow");
                        return context.sync().then(function () {
                            var shadow = shape.shadow;

                            // Set shadow properties
                            shadow.visible = true;
                            shadow.color = shadowColor;
                            shadow.transparency = transparency;
                            shadow.blur = blur;
                            shadow.angle = angle;
                            shadow.distance = distance;

                            return context.sync();
                        });
                    });

                    return Promise.all(promises).then(function () {
                        var count = shapes.items.length;
                        var text = count === 1 ? "1 Objekt" : count + " Objekte";
                        showStatus("Schatten hinzugefügt (" + text + ") ✓", "success");
                    });
                });
            }).catch(function (error) {
                showStatus("Fehler: " + error.message, "error");
            });
        }
    );
}

// ===================================================================
// TAB 4: SHADOW - REMOVE SHADOW
// ===================================================================
function removeShadow() {
    Office.context.document.getSelectedDataAsync(
        Office.CoercionType.SlideRange,
        function (asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                showStatus("Fehler: " + asyncResult.error.message, "error");
                return;
            }

            PowerPoint.run(function (context) {
                var shapes = context.presentation.getSelectedShapes();
                shapes.load("items");

                return context.sync().then(function () {
                    if (shapes.items.length === 0) {
                        showStatus("Bitte Objekt(e) auswählen!", "error");
                        return;
                    }

                    var promises = shapes.items.map(function (shape) {
                        shape.load("shadow");
                        return context.sync().then(function () {
                            shape.shadow.visible = false;
                            return context.sync();
                        });
                    });

                    return Promise.all(promises).then(function () {
                        var count = shapes.items.length;
                        var text = count === 1 ? "1 Objekt" : count + " Objekte";
                        showStatus("Schatten entfernt (" + text + ") ✓", "success");
                    });
                });
            }).catch(function (error) {
                showStatus("Fehler: " + error.message, "error");
            });
        }
    );
}
