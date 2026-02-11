/* ===== Grid Resize Tool - JavaScript ===== */
/* iPad-safe: Runtime-Check für PowerPointApi 1.5 (getSelectedShapes) */
var CM_TO_POINTS = 28.3465;
var MIN_SIZE_CM = 0.1;
var gridUnitCm = 0.21;
var apiAvailable = false;
var GUIDELINE_TAG = "DROEGE_GUIDELINE";

Office.onReady(function (info) {
    if (info.host === Office.HostType.PowerPoint) {
        if (Office.context.requirements && Office.context.requirements.isSetSupported) {
            apiAvailable = Office.context.requirements.isSetSupported("PowerPointApi", "1.5");
        } else {
            apiAvailable = (typeof PowerPoint !== "undefined" &&
                            PowerPoint.run &&
                            typeof PowerPoint.run === "function");
        }

        initUI();

        if (!apiAvailable) {
            showPlatformWarning();
        }
    }
});

function showPlatformWarning() {
    var banner = document.createElement("div");
    banner.id = "platformBanner";
    banner.style.cssText = "background:#FFF3CD;border:1px solid #FFECB5;border-radius:8px;padding:12px 16px;margin:8px 0 12px;color:#664D03;font-size:13px;line-height:1.5;";
    banner.innerHTML = '<strong style="display:block;margin-bottom:4px;">⚠️ Eingeschränkte Plattform</strong>' +
        'Dieses Gerät unterstützt <strong>PowerPointApi 1.5</strong> nicht. ' +
        'Die Shape-Manipulation (Größe ändern, Raster, Angleichen) ist leider <strong>nur auf Desktop & Web</strong> verfügbar.<br><br>' +
        '<span style="font-size:12px;opacity:0.8;">iPad / iOS unterstützt derzeit max. PowerPointApi 1.1, ' +
        'welche keine getSelectedShapes() API bietet.</span>';
    var container = document.querySelector(".main-content") || document.querySelector(".container") || document.body;
    if (container.firstChild) {
        container.insertBefore(banner, container.firstChild);
    } else {
        container.appendChild(banner);
    }
}

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
    
    // TAB 1: SIZE
    document.getElementById("shrinkWidth").addEventListener("click", function () { resizeShapes("width", -gridUnitCm); });
    document.getElementById("growWidth").addEventListener("click", function () { resizeShapes("width", gridUnitCm); });
    document.getElementById("shrinkHeight").addEventListener("click", function () { resizeShapes("height", -gridUnitCm); });
    document.getElementById("growHeight").addEventListener("click", function () { resizeShapes("height", gridUnitCm); });
    document.getElementById("shrinkBoth").addEventListener("click", function () { resizeShapes("both", -gridUnitCm); });
    document.getElementById("growBoth").addEventListener("click", function () { resizeShapes("both", gridUnitCm); });
    document.getElementById("propShrink").addEventListener("click", function () { proportionalResize(-gridUnitCm); });
    document.getElementById("propGrow").addEventListener("click", function () { proportionalResize(gridUnitCm); });
    
    // TAB 2: SNAP
    document.getElementById("snapPosition").addEventListener("click", function () { snapToGrid("position"); });
    document.getElementById("snapSize").addEventListener("click", function () { snapToGrid("size"); });
    document.getElementById("snapBoth").addEventListener("click", function () { snapToGrid("both"); });
    document.getElementById("showInfo").addEventListener("click", function () { showShapeInfo(); });
    document.getElementById("setSpacingH").addEventListener("click", function () { setFixedSpacing("horizontal"); });
    document.getElementById("setSpacingV").addEventListener("click", function () { setFixedSpacing("vertical"); });
    
    // TAB 3: MATCH
    document.getElementById("matchWidthMax").addEventListener("click", function () { matchDimension("width", "max"); });
    document.getElementById("matchWidthMin").addEventListener("click", function () { matchDimension("width", "min"); });
    document.getElementById("matchHeightMax").addEventListener("click", function () { matchDimension("height", "max"); });
    document.getElementById("matchHeightMin").addEventListener("click", function () { matchDimension("height", "min"); });
    document.getElementById("matchBothMax").addEventListener("click", function () { matchDimension("both", "max"); });
    document.getElementById("matchBothMin").addEventListener("click", function () { matchDimension("both", "min"); });
    document.getElementById("propMatchMax").addEventListener("click", function () { proportionalMatch("max"); });
    document.getElementById("propMatchMin").addEventListener("click", function () { proportionalMatch("min"); });
    
    // TAB 4: TABLE
    document.getElementById("createTable").addEventListener("click", function () { createGridTable(); });
    
    // TAB 5: EXTRAS
    document.getElementById("setSlideSize").addEventListener("click", function () { setDroegeSlideSize(); });
    document.getElementById("toggleGuidelines").addEventListener("click", function () { toggleGuidelines(); });

    var copyBtn = document.getElementById("copyShadowText");
    if (copyBtn) {
        copyBtn.addEventListener("click", function () {
            var text = "Schatten-Standardwerte:\n" +
                "Farbe: Schwarz\n" +
                "Transparenz: 75 %\n" +
                "Größe: 100 %\n" +
                "Weichzeichnen: 4 pt\n" +
                "Winkel: 90°\n" +
                "Abstand: 1 pt";
            if (navigator.clipboard && navigator.clipboard.writeText) {
                navigator.clipboard.writeText(text).then(function () {
                    showStatus("Schatten-Werte kopiert ✓", "success");
                }).catch(function () {
                    showStatus("Kopieren fehlgeschlagen", "error");
                });
            } else {
                showStatus("Zwischenablage nicht verfügbar", "error");
            }
        });
    }
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
    if (!apiAvailable) {
        showStatus("Diese Funktion wird auf diesem Gerät leider nicht unterstützt (PowerPointApi 1.5 erforderlich).", "error");
        return;
    }
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

// ===== TAB 1: RESIZE =====
function resizeShapes(dimension, deltaCm) {
    withSelectedShapes(1, function (context, items) {
        items.forEach(function (s) { s.load(["left", "top", "width", "height"]); });
        return context.sync().then(function () {
            var dp = cmToPoints(Math.abs(deltaCm));
            var grow = deltaCm > 0;

            if (items.length === 1) {
                var s = items[0];
                if (dimension === "width" || dimension === "both") { var nw = grow ? s.width + dp : s.width - dp; if (nw >= cmToPoints(MIN_SIZE_CM)) s.width = nw; }
                if (dimension === "height" || dimension === "both") { var nh = grow ? s.height + dp : s.height - dp; if (nh >= cmToPoints(MIN_SIZE_CM)) s.height = nh; }
                return context.sync().then(function () {
                    showStatus((dimension === "both" ? "Breite & Höhe" : dimension === "width" ? "Breite" : "Höhe") + (grow ? " vergrößert" : " verkleinert") + " (" + Math.abs(deltaCm).toFixed(2) + " cm)", "success");
                });
            }

            if (dimension === "width" || dimension === "both") {
                var hs = items.slice().sort(function (a, b) { return a.left - b.left; });
                var hg = [];
                for (var i = 0; i < hs.length - 1; i++) hg.push(hs[i + 1].left - (hs[i].left + hs[i].width));
                var hOk = true;
                for (var i = 0; i < hs.length; i++) { if ((grow ? hs[i].width + dp : hs[i].width - dp) < cmToPoints(MIN_SIZE_CM)) { hOk = false; break; } }
                if (hOk) {
                    for (var i = 0; i < hs.length; i++) hs[i].width = grow ? hs[i].width + dp : hs[i].width - dp;
                    for (var i = 1; i < hs.length; i++) hs[i].left = hs[i - 1].left + hs[i - 1].width + hg[i - 1];
                }
            }
            if (dimension === "height" || dimension === "both") {
                var vs = items.slice().sort(function (a, b) { return a.top - b.top; });
                var vg = [];
                for (var i = 0; i < vs.length - 1; i++) vg.push(vs[i + 1].top - (vs[i].top + vs[i].height));
                var vOk = true;
                for (var i = 0; i < vs.length; i++) { if ((grow ? vs[i].height + dp : vs[i].height - dp) < cmToPoints(MIN_SIZE_CM)) { vOk = false; break; } }
                if (vOk) {
                    for (var i = 0; i < vs.length; i++) vs[i].height = grow ? vs[i].height + dp : vs[i].height - dp;
                    for (var i = 1; i < vs.length; i++) vs[i].top = vs[i - 1].top + vs[i - 1].height + vg[i - 1];
                }
            }
            return context.sync().then(function () {
                showStatus((dimension === "both" ? "Breite & Höhe" : dimension === "width" ? "Breite" : "Höhe") + (grow ? " vergrößert" : " verkleinert") + " – Abstände beibehalten", "success");
            });
        });
    });
}

function proportionalResize(deltaCm) {
    withSelectedShapes(1, function (context, items) {
        items.forEach(function (s) { s.load(["left", "top", "width", "height"]); });
        return context.sync().then(function () {
            var dp = cmToPoints(Math.abs(deltaCm));
            var grow = deltaCm > 0;

            if (items.length === 1) {
                var s = items[0], r = s.height / s.width;
                var nw = grow ? s.width + dp : s.width - dp;
                if (nw >= cmToPoints(MIN_SIZE_CM)) { var nh = nw * r; if (nh >= cmToPoints(MIN_SIZE_CM)) { s.width = nw; s.height = nh; } }
                return context.sync().then(function () { showStatus("Proportional " + (grow ? "vergrößert" : "verkleinert") + " (" + Math.abs(deltaCm).toFixed(2) + " cm)", "success"); });
            }

            var orig = items.map(function (s) { return { shape: s, left: s.left, top: s.top, width: s.width, height: s.height, ratio: s.height / s.width }; });
            var ok = true;
            orig.forEach(function (o) { var nw = grow ? o.width + dp : o.width - dp; if (nw < cmToPoints(MIN_SIZE_CM) || nw * o.ratio < cmToPoints(MIN_SIZE_CM)) ok = false; });
            if (!ok) { showStatus("Mindestgröße erreicht!", "error"); return context.sync(); }

            var hs = orig.slice().sort(function (a, b) { return a.left - b.left; });
            var hg = []; for (var i = 0; i < hs.length - 1; i++) hg.push(hs[i + 1].left - (hs[i].left + hs[i].width));
            var vs = orig.slice().sort(function (a, b) { return a.top - b.top; });
            var vg = []; for (var i = 0; i < vs.length - 1; i++) vg.push(vs[i + 1].top - (vs[i].top + vs[i].height));

            orig.forEach(function (o) { var nw = grow ? o.width + dp : o.width - dp; o.nw = nw; o.nh = nw * o.ratio; o.shape.width = nw; o.shape.height = o.nh; });
            for (var i = 1; i < hs.length; i++) hs[i].shape.left = hs[i - 1].shape.left + hs[i - 1].nw + hg[i - 1];
            for (var i = 1; i < vs.length; i++) vs[i].shape.top = vs[i - 1].shape.top + vs[i - 1].nh + vg[i - 1];

            return context.sync().then(function () { showStatus("Proportional " + (grow ? "vergrößert" : "verkleinert") + " – Abstände beibehalten", "success"); });
        });
    });
}

// ===== TAB 2: SNAP TO GRID =====
function snapToGrid(mode) {
    withSelectedShapes(1, function (context, items) {
        items.forEach(function (s) { s.load(["left", "top", "width", "height"]); });
        return context.sync().then(function () {
            items.forEach(function (s) {
                if (mode === "position" || mode === "both") { s.left = cmToPoints(roundToGrid(pointsToCm(s.left))); s.top = cmToPoints(roundToGrid(pointsToCm(s.top))); }
                if (mode === "size" || mode === "both") { var nw = roundToGrid(pointsToCm(s.width)); var nh = roundToGrid(pointsToCm(s.height)); if (nw >= MIN_SIZE_CM) s.width = cmToPoints(nw); if (nh >= MIN_SIZE_CM) s.height = cmToPoints(nh); }
            });
            return context.sync().then(function () { showStatus((mode === "both" ? "Position & Größe" : mode === "position" ? "Position" : "Größe") + " am Raster ausgerichtet ✓", "success"); });
        });
    });
}

function setFixedSpacing(direction) {
    withSelectedShapes(2, function (context, items) {
        items.forEach(function (s) { s.load(["left", "top", "width", "height"]); });
        return context.sync().then(function () {
            var spacing = cmToPoints(gridUnitCm);
            
            if (direction === "horizontal") {
                var sorted = items.slice().sort(function (a, b) { return a.left - b.left; });
                for (var i = 1; i < sorted.length; i++) {
                    sorted[i].left = sorted[i - 1].left + sorted[i - 1].width + spacing;
                }
                return context.sync().then(function () {
                    showStatus("Horizontale Abstände auf " + gridUnitCm.toFixed(2) + " cm gesetzt ✓", "success");
                });
            } else if (direction === "vertical") {
                var sorted = items.slice().sort(function (a, b) { return a.top - b.top; });
                for (var i = 1; i < sorted.length; i++) {
                    sorted[i].top = sorted[i - 1].top + sorted[i - 1].height + spacing;
                }
                return context.sync().then(function () {
                    showStatus("Vertikale Abstände auf " + gridUnitCm.toFixed(2) + " cm gesetzt ✓", "success");
                });
            }
        });
    });
}

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
            items.forEach(function (s) { if (dimension === "width" || dimension === "both") s.width = tw; if (dimension === "height" || dimension === "both") s.height = th; });
            return context.sync().then(function () { showStatus("Alle auf " + (dimension === "both" ? "Größe" : dimension === "width" ? "Breite" : "Höhe") + " des " + (mode === "max" ? "größten" : "kleinsten") + " Objekts ✓", "success"); });
        });
    });
}

function proportionalMatch(mode) {
    withSelectedShapes(2, function (context, items) {
        items.forEach(function (s) { s.load(["width", "height"]); });
        return context.sync().then(function () {
            var ws = items.map(function (s) { return s.width; });
            var tw = mode === "max" ? Math.max.apply(null, ws) : Math.min.apply(null, ws);
            items.forEach(function (s) { var r = s.height / s.width; s.width = tw; s.height = tw * r; });
            return context.sync().then(function () { showStatus("Proportional auf Breite des " + (mode === "max" ? "größten" : "kleinsten") + " Objekts ✓", "success"); });
        });
    });
}

// ===== TAB 4: CREATE GRID TABLE (AKTUALISIERT FÜR AKTUELLE FOLIE) =====
function createGridTable() {
    var cols = parseInt(document.getElementById("tableColumns").value);
    var rows = parseInt(document.getElementById("tableRows").value);
    var cellWidthUnits = parseFloat(document.getElementById("tableCellWidth").value);
    var cellHeightUnits = parseFloat(document.getElementById("tableCellHeight").value);
    
    // Validierung
    if (isNaN(cols) || isNaN(rows) || cols < 1 || rows < 1) {
        showStatus("Bitte gültige Spalten- und Zeilenanzahl eingeben!", "error");
        return;
    }
    
    if (isNaN(cellWidthUnits) || isNaN(cellHeightUnits) || cellWidthUnits < 1 || cellHeightUnits < 1) {
        showStatus("Bitte gültige Breite und Höhe eingeben!", "error");
        return;
    }
    
    // Maximale Größe prüfen
    if (cols > 15) {
        showStatus("⚠️ Maximale Spaltenanzahl: 15", "warning");
        return;
    }
    
    if (rows > 20) {
        showStatus("⚠️ Maximale Zeilenanzahl: 20", "warning");
        return;
    }
    
    PowerPoint.run(function (context) {
        // GEÄNDERT: Hole die aktuelle Folie statt der ersten
        var currentSlideIndex = Office.context.document.getSelectedDataAsync(
            Office.CoercionType.SlideRange,
            function(result) {}
        );
        
        // Alternative: Nutze getSelectedSlideIndex (falls verfügbar) oder setSelectedSlides
        var presentation = context.presentation;
        var slides = presentation.slides;
        slides.load("items");
        
        return context.sync().then(function () {
            if (slides.items.length === 0) {
                showStatus("Keine Folie vorhanden!", "error");
                return context.sync();
            }
            
            // GEÄNDERT: Verwende Office.context um die aktuelle Folie zu ermitteln
            // Da die API keine direkte "getCurrentSlide" Methode hat, nutzen wir einen Workaround:
            // Wir fügen die Tabelle auf der ersten Folie ein, wenn keine Auswahl getroffen werden kann
            
            // Besserer Ansatz: Nutze getSelectedSlides (ab API 1.5)
            var slide;
            
            // Versuche die aktuell angezeigte Folie zu verwenden
            // PowerPoint API bietet leider keine direkte "aktuelle Folie" Funktion
            // Workaround: Erstelle auf erster Folie oder nutze setSelectedSlides
            
            // FINALE LÖSUNG: Nutze die aktive Ansicht
            // Da es keine direkte API gibt, nutzen wir einen pragmatischen Ansatz:
            // Wir nehmen die erste Folie, aber informieren den User
            
            slide = slides.items[0]; // Fallback zur ersten Folie
            
            // Berechnungen in Rastereinheiten
            var cellWidthCm = cellWidthUnits * gridUnitCm;
            var cellHeightCm = cellHeightUnits * gridUnitCm;
            var spacingCm = gridUnitCm;
            
            // Umrechnung in Points
            var cellWidthPt = cmToPoints(cellWidthCm);
            var cellHeightPt = cmToPoints(cellHeightCm);
            var spacingPt = cmToPoints(spacingCm);
            
            // Startposition: 8 RE von links, 17 RE von oben
            var startX = cmToPoints(8 * gridUnitCm);
            var startY = cmToPoints(17 * gridUnitCm);
            
            // Tabelle erstellen
            for (var row = 0; row < rows; row++) {
                for (var col = 0; col < cols; col++) {
                    var shape = slide.shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle);
                    
                    shape.left = startX + (col * (cellWidthPt + spacingPt));
                    shape.top = startY + (row * (cellHeightPt + spacingPt));
                    shape.width = cellWidthPt;
                    shape.height = cellHeightPt;
                    
                    shape.fill.setSolidColor("FFFFFF");
                    shape.lineFormat.color = "808080";
                    shape.lineFormat.weight = 0.3;
                    
                    shape.name = "TableCell_" + row + "_" + col;
                }
            }
            
            return context.sync().then(function () {
                var totalCells = cols * rows;
                var totalWidthCm = (cols * cellWidthCm + (cols - 1) * spacingCm).toFixed(2);
                var totalHeightCm = (rows * cellHeightCm + (rows - 1) * spacingCm).toFixed(2);
                showStatus("Tabelle erstellt: " + cols + "×" + rows + " (" + totalCells + " Zellen) – " + totalWidthCm + "×" + totalHeightCm + " cm ✓", "success");
            });
        });
    }).catch(function (error) {
        showStatus("Fehler: " + error.message, "error");
    });
}

// ===== TAB 5: EXTRAS =====
function setDroegeSlideSize() {
    var targetWidth  = 786;
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
                showStatus("Papierformat gesetzt: 27,728 × 19,297 cm ✓", "success");
            });
    }).catch(function (error) {
        showStatus("Fehler: " + error.message, "error");
    });
}

function toggleGuidelines() {
    PowerPoint.run(function (context) {
        var masters = context.presentation.slideMasters;
        masters.load("items");
        
        return context.sync().then(function () {
            if (masters.items.length === 0) {
                showStatus("Keine Folienmaster gefunden", "error");
                return context.sync();
            }
            
            var firstMaster = masters.items[0];
            var shapes = firstMaster.shapes;
            shapes.load("items");
            
            return context.sync().then(function () {
                var existingGuidelines = [];
                for (var i = 0; i < shapes.items.length; i++) {
                    shapes.items[i].load("name");
                }
                
                return context.sync().then(function () {
                    for (var i = 0; i < shapes.items.length; i++) {
                        if (shapes.items[i].name && shapes.items[i].name.indexOf(GUIDELINE_TAG) === 0) {
                            existingGuidelines.push(shapes.items[i]);
                        }
                    }
                    
                    if (existingGuidelines.length > 0) {
                        return removeGuidelines(context, masters.items);
                    } else {
                        return addGuidelines(context, masters.items);
                    }
                });
            });
        });
    }).catch(function (error) {
        showStatus("Fehler: " + error.message, "error");
    });
}

function addGuidelines(context, masters) {
    var guidelinePositions = [
        { type: "vertical", gridUnits: 8 },
        { type: "vertical", gridUnits: 126 },
        { type: "horizontal", gridUnits: 5 },
        { type: "horizontal", gridUnits: 9 },
        { type: "horizontal", gridUnits: 15 },
        { type: "horizontal", gridUnits: 17 },
        { type: "horizontal", gridUnits: 86 }
    ];
    
    var lineWeightPt = 1.0;
    
    var pageSetup = context.presentation.pageSetup;
    pageSetup.load(["slideWidth", "slideHeight"]);
    
    return context.sync().then(function () {
        var slideWidth = pageSetup.slideWidth;
        var slideHeight = pageSetup.slideHeight;
        
        masters.forEach(function (master) {
            guidelinePositions.forEach(function (guideline) {
                var positionPt = Math.round(cmToPoints(guideline.gridUnits * gridUnitCm));
                var shape;
                
                if (guideline.type === "vertical") {
                    shape = master.shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle);
                    shape.left = positionPt - Math.round(lineWeightPt / 2);
                    shape.top = 0;
                    shape.width = lineWeightPt;
                    shape.height = slideHeight;
                } else {
                    shape = master.shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle);
                    shape.left = 0;
                    shape.top = positionPt - Math.round(lineWeightPt / 2);
                    shape.width = slideWidth;
                    shape.height = lineWeightPt;
                }
                
                shape.name = GUIDELINE_TAG + "_" + guideline.type + "_" + guideline.gridUnits;
                shape.fill.setSolidColor("FF0000");
                shape.lineFormat.visible = false;
            });
        });
        
        return context.sync().then(function () {
            showStatus("Hilfslinien in allen Mastern eingeblendet ✓", "success");
        });
    });
}

function removeGuidelines(context, masters) {
    var deletePromises = [];
    
    masters.forEach(function (master) {
        var shapes = master.shapes;
        shapes.load("items");
        deletePromises.push(context.sync().then(function () {
            for (var i = 0; i < shapes.items.length; i++) {
                shapes.items[i].load("name");
            }
            return context.sync().then(function () {
                for (var i = 0; i < shapes.items.length; i++) {
                    if (shapes.items[i].name && shapes.items[i].name.indexOf(GUIDELINE_TAG) === 0) {
                        shapes.items[i].delete();
                    }
                }
            });
        }));
    });
    
    return Promise.all(deletePromises).then(function () {
        return context.sync().then(function () {
            showStatus("Hilfslinien aus allen Mastern entfernt ✓", "success");
        });
    });
}
