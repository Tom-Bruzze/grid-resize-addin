/* ═══════════════════════════════════════════════════════════════
   DROEGE Grid Resize Tool – taskpane.js
   ═══════════════════════════════════════════════════════════════
   3 Tabs: Tools (Größe + Angleich + Raster + Abstände)
           Extras (Papierformat + Hilfslinien + Schatten)
           Gantt  (Gantt-Chart Generator)
   ═══════════════════════════════════════════════════════════════ */

/* ── globals ── */
var gridUnitCm = 0.21;
var apiOk = false;
var GTAG = "DG_GUIDE";

function c2p(cm) { return cm / 2.54 * 72; }
function p2c(pt) { return pt * 2.54 / 72; }

function showStatus(msg, cls) {
    var el = document.getElementById("status");
    el.textContent = msg;
    el.className = "sts" + (cls ? " " + cls : "");
}

/* ══════════════════════════════════════════════════════════════
   INIT
   ══════════════════════════════════════════════════════════════ */
Office.onReady(function (info) {
    if (info.host !== Office.HostType.PowerPoint) {
        showStatus("Nur PowerPoint!", "error");
        return;
    }
    apiOk = Office.context.requirements.isSetSupported("PowerPointApi", "1.5");

    function bind(id, fn) {
        var el = document.getElementById(id);
        if (el) el.addEventListener("click", fn);
    }

    /* Presets */
    document.querySelectorAll(".pre").forEach(function (b) {
        b.addEventListener("click", function () {
            document.querySelectorAll(".pre").forEach(function (p) { p.classList.remove("active"); });
            this.classList.add("active");
            var v = parseFloat(this.dataset.value);
            gridUnitCm = v;
            document.getElementById("gridUnit").value = v;
            showStatus("RE = " + v.toFixed(2) + " cm", "info");
        });
    });
    document.getElementById("gridUnit").addEventListener("change", function () {
        var v = parseFloat(this.value);
        if (v > 0) { gridUnitCm = v; showStatus("RE = " + v.toFixed(2) + " cm", "info"); }
    });

    /* Tabs */
    document.querySelectorAll(".tab").forEach(function (b) {
        b.addEventListener("click", function () {
            var id = this.dataset.tab;
            document.querySelectorAll(".tab").forEach(function (t) { t.classList.remove("active"); });
            document.querySelectorAll(".pane").forEach(function (p) { p.classList.remove("active"); });
            this.classList.add("active");
            var pane = document.getElementById(id);
            if (pane) pane.classList.add("active");
        });
    });

    /* ── Größe ── */
    bind("resizeW",    function (e) { resize("width",  e.shiftKey); });
    bind("resizeH",    function (e) { resize("height", e.shiftKey); });
    bind("resizeBoth", function (e) { resize("both",   e.shiftKey); });
    bind("resizeProp", function (e) { resize("prop",   e.shiftKey); });

    /* ── Angleichen ── */
    bind("matchW",    function (e) { matchShapes("width",  e.shiftKey); });
    bind("matchH",    function (e) { matchShapes("height", e.shiftKey); });
    bind("matchBoth", function (e) { matchShapes("both",   e.shiftKey); });
    bind("matchProp", function (e) { matchShapes("prop",   e.shiftKey); });

    /* ── Raster ── */
    bind("snapPos",     function () { snap("position"); });
    bind("snapSize",    function () { snap("size");     });
    bind("snapAll",     function () { snap("both");     });
    bind("showInfo",    function () { shapeInfo();      });

    /* ── Abstände ── */
    bind("spaceH",      function () { spacing("h"); });
    bind("spaceV",      function () { spacing("v"); });

    /* ── Grid-Tabelle ── */
    bind("createTable", function () { createGridTable(); });

    /* ── Extras ── */
    bind("setSlide",      function () { setSlideSize(); });
    bind("toggleGuides",  function () { toggleGuides(); });
    bind("copyShadow",    function () { copyShadowText(); });

    /* ── Gantt ── */
    initGantt();

    showStatus("Bereit", "");
});

/* ══════════════════════════════════════════════════════════════
   HELPER – withShapes
   ══════════════════════════════════════════════════════════════ */
function withShapes(min, cb) {
    if (!apiOk) { showStatus("PowerPointApi 1.5 nötig", "error"); return; }
    PowerPoint.run(function (ctx) {
        var shapes = ctx.presentation.getSelectedShapes();
        shapes.load("items");
        return ctx.sync().then(function () {
            if (shapes.items.length < min) {
                showStatus("Mind. " + min + " Shape(s) wählen!", "warning");
                return ctx.sync();
            }
            return cb(ctx, shapes.items);
        });
    }).catch(function (e) { showStatus("Fehler: " + e.message, "error"); });
}

/* ══════════════════════════════════════════════════════════════
   RESIZE
   ══════════════════════════════════════════════════════════════ */
function resize(mode, shrink) {
    var d = shrink ? -1 : 1;
    var step = c2p(gridUnitCm);

    withShapes(1, function (ctx, items) {
        items.forEach(function (s) { s.load(["width", "height"]); });
        return ctx.sync().then(function () {
            items.forEach(function (s) {
                if (mode === "width" || mode === "both") s.width  = Math.max(step, s.width  + d * step);
                if (mode === "height"|| mode === "both") s.height = Math.max(step, s.height + d * step);
                if (mode === "prop") {
                    var r = s.height / s.width;
                    s.width  = Math.max(step, s.width + d * step);
                    s.height = Math.max(step, s.width * r);
                }
            });
            return ctx.sync().then(function () {
                showStatus((shrink ? "−" : "+") + " " + mode + " ✓", "success");
            });
        });
    });
}

/* ══════════════════════════════════════════════════════════════
   MATCH SHAPES
   ══════════════════════════════════════════════════════════════ */
function matchShapes(mode, useMin) {
    withShapes(2, function (ctx, items) {
        items.forEach(function (s) { s.load(["width", "height"]); });
        return ctx.sync().then(function () {
            var fn = useMin ? Math.min : Math.max;
            var ws = items.map(function (s) { return s.width; });
            var hs = items.map(function (s) { return s.height; });
            var tw = fn.apply(null, ws);
            var th = fn.apply(null, hs);

            items.forEach(function (s) {
                if (mode === "width"  || mode === "both") s.width  = tw;
                if (mode === "height" || mode === "both") s.height = th;
                if (mode === "prop") { s.width = tw; s.height = th; }
            });
            return ctx.sync().then(function () {
                showStatus("Angleich " + (useMin ? "Min" : "Max") + " ✓", "success");
            });
        });
    });
}

/* ══════════════════════════════════════════════════════════════
   SNAP TO GRID
   ══════════════════════════════════════════════════════════════ */
function snap(mode) {
    var step = c2p(gridUnitCm);
    withShapes(1, function (ctx, items) {
        items.forEach(function (s) { s.load(["left", "top", "width", "height"]); });
        return ctx.sync().then(function () {
            items.forEach(function (s) {
                if (mode === "position" || mode === "both") {
                    s.left = Math.round(s.left / step) * step;
                    s.top  = Math.round(s.top  / step) * step;
                }
                if (mode === "size" || mode === "both") {
                    s.width  = Math.max(step, Math.round(s.width  / step) * step);
                    s.height = Math.max(step, Math.round(s.height / step) * step);
                }
            });
            return ctx.sync().then(function () { showStatus("Raster ✓", "success"); });
        });
    });
}

/* ══════════════════════════════════════════════════════════════
   SHAPE INFO
   ══════════════════════════════════════════════════════════════ */
function shapeInfo() {
    withShapes(1, function (ctx, items) {
        items.forEach(function (s) { s.load(["name", "left", "top", "width", "height"]); });
        return ctx.sync().then(function () {
            var s = items[0];
            var gu = gridUnitCm;
            var html = "";
            var rows = [
                ["Name",   s.name || "–"],
                ["X (RE)", (p2c(s.left)   / gu).toFixed(1)],
                ["Y (RE)", (p2c(s.top)    / gu).toFixed(1)],
                ["W (RE)", (p2c(s.width)  / gu).toFixed(1)],
                ["H (RE)", (p2c(s.height) / gu).toFixed(1)],
                ["W (cm)", p2c(s.width).toFixed(2)],
                ["H (cm)", p2c(s.height).toFixed(2)]
            ];
            rows.forEach(function (r) {
                html += '<div class="info-item"><span class="info-label">' + r[0] + '</span><span class="info-value">' + r[1] + '</span></div>';
            });
            var box = document.getElementById("infoBox");
            box.innerHTML = html;
            box.classList.add("visible");
            box.style.display = "block";
            showStatus("Info ✓", "success");
            return ctx.sync();
        });
    });
}

/* ══════════════════════════════════════════════════════════════
   SPACING – Multi-Row / Multi-Column
   ══════════════════════════════════════════════════════════════ */
function spacing(dir) {
    withShapes(2, function (ctx, items) {
        items.forEach(function (s) { s.load(["left", "top", "width", "height"]); });
        return ctx.sync().then(function () {
            var data = items.map(function (s) {
                return { shape: s, left: s.left, top: s.top, width: s.width, height: s.height, x: s.left, y: s.top };
            });
            var sp = c2p(gridUnitCm);
            var tol = c2p(gridUnitCm * 2);

            if (dir === "h") {
                var rows = groupByPos(data, "y", tol);
                rows.forEach(function (row) {
                    row.sort(function (a, b) { return a.left - b.left; });
                    for (var i = 1; i < row.length; i++)
                        row[i].shape.left = row[i - 1].shape.left + row[i - 1].width + sp;
                });
            }
            if (dir === "v") {
                var cols = groupByPos(data, "x", tol);
                cols.forEach(function (col) {
                    col.sort(function (a, b) { return a.top - b.top; });
                    var gaps = [];
                    for (var i = 0; i < col.length - 1; i++)
                        gaps.push(col[i + 1].top - (col[i].top + col[i].height));
                    for (var i = 1; i < col.length; i++)
                        col[i].shape.top = col[i - 1].shape.top + col[i - 1].height + sp;
                });
            }

            return ctx.sync().then(function () { showStatus("Abstand " + dir.toUpperCase() + " ✓", "success"); });
        });
    });
}

function groupByPos(arr, key, tol) {
    var groups = [], used = [];
    for (var i = 0; i < arr.length; i++) {
        if (used[i]) continue;
        var g = [arr[i]]; used[i] = true;
        for (var j = i + 1; j < arr.length; j++) {
            if (!used[j] && Math.abs(arr[i][key] - arr[j][key]) < tol) { g.push(arr[j]); used[j] = true; }
        }
        groups.push(g);
    }
    return groups;
}

/* ══════════════════════════════════════════════════════════════
   GRID TABLE
   ══════════════════════════════════════════════════════════════ */
function createGridTable() {
    var cols = parseInt(document.getElementById("tCols").value);
    var rows = parseInt(document.getElementById("tRows").value);
    var cw   = parseFloat(document.getElementById("tCW").value);
    var ch   = parseFloat(document.getElementById("tCH").value);

    if (isNaN(cols) || isNaN(rows) || cols < 1 || rows < 1) { showStatus("Ungültige Sp/Zl!", "error"); return; }
    if (isNaN(cw) || isNaN(ch) || cw < 1 || ch < 1) { showStatus("Ungültige Zellengröße!", "error"); return; }
    if (cols > 15) { showStatus("Max. 15 Spalten!", "warning"); return; }
    if (rows > 20) { showStatus("Max. 20 Zeilen!", "warning"); return; }

    PowerPoint.run(function (ctx) {
        var sel = ctx.presentation.getSelectedSlides();
        sel.load("items");
        return ctx.sync().then(function () {
            if (sel.items.length > 0) return buildTbl(ctx, sel.items[0], cols, rows, cw, ch);
            var slides = ctx.presentation.slides;
            slides.load("items");
            return ctx.sync().then(function () {
                if (!slides.items.length) { showStatus("Keine Folie!", "error"); return ctx.sync(); }
                return buildTbl(ctx, slides.items[0], cols, rows, cw, ch);
            });
        });
    }).catch(function (e) { showStatus("Fehler: " + e.message, "error"); });
}

function buildTbl(ctx, slide, cols, rows, cwRE, chRE) {
    var wPt = c2p(cwRE * gridUnitCm), hPt = c2p(chRE * gridUnitCm);
    var sp = c2p(gridUnitCm);
    var x0 = c2p(8 * gridUnitCm), y0 = c2p(17 * gridUnitCm);

    for (var r = 0; r < rows; r++) {
        for (var c = 0; c < cols; c++) {
            var s = slide.shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle);
            s.left = x0 + c * (wPt + sp);
            s.top = y0 + r * (hPt + sp);
            s.width = wPt;
            s.height = hPt;
            s.fill.setSolidColor("FFFFFF");
            s.lineFormat.color = "808080";
            s.lineFormat.weight = 0.3;
            s.name = "GridCell_R" + r + "_C" + c;
        }
    }
    return ctx.sync().then(function () {
        showStatus(cols + " × " + rows + " Tabelle ✓", "success");
    });
}

/* ══════════════════════════════════════════════════════════════
   PAPIERFORMAT – 27,728 × 19,297 cm
   ══════════════════════════════════════════════════════════════ */
function setSlideSize() {
    PowerPoint.run(function (ctx) {
        var ps = ctx.presentation.pageSetup;
        ps.load(["slideWidth", "slideHeight"]);
        return ctx.sync().then(function () {
            ps.slideWidth = 786;
            return ctx.sync();
        }).then(function () {
            ps.slideHeight = 547;
            return ctx.sync();
        }).then(function () {
            showStatus("Format: 27,728 × 19,297 cm ✓", "success");
        });
    }).catch(function (e) { showStatus("Fehler: " + e.message, "error"); });
}

/* ══════════════════════════════════════════════════════════════
   HILFSLINIEN (Master Toggle)
   ══════════════════════════════════════════════════════════════ */
function toggleGuides() {
    PowerPoint.run(function (ctx) {
        var masters = ctx.presentation.slideMasters;
        masters.load("items");
        return ctx.sync().then(function () {
            if (!masters.items.length) { showStatus("Kein Master!", "error"); return ctx.sync(); }
            var m = masters.items[0], sh = m.shapes;
            sh.load("items");
            return ctx.sync().then(function () {
                for (var i = 0; i < sh.items.length; i++) sh.items[i].load("name");
                return ctx.sync().then(function () {
                    var ex = [];
                    for (var i = 0; i < sh.items.length; i++)
                        if (sh.items[i].name && sh.items[i].name.indexOf(GTAG) === 0) ex.push(sh.items[i]);
                    return ex.length > 0 ? rmGuides(ctx, masters.items) : addGuides(ctx, masters.items);
                });
            });
        });
    }).catch(function (e) { showStatus("Fehler: " + e.message, "error"); });
}

function addGuides(ctx, masters) {
    var lines = [
        { t: "v", p: 8 }, { t: "v", p: 126 },
        { t: "h", p: 5 }, { t: "h", p: 9 }, { t: "h", p: 15 }, { t: "h", p: 17 }, { t: "h", p: 86 }
    ];
    var ps = ctx.presentation.pageSetup;
    ps.load(["slideWidth", "slideHeight"]);
    return ctx.sync().then(function () {
        var sw = ps.slideWidth, sh = ps.slideHeight;
        masters.forEach(function (master) {
            lines.forEach(function (l) {
                var pt = Math.round(c2p(l.p * gridUnitCm));
                var s = master.shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle);
                if (l.t === "v") { s.left = pt; s.top = 0; s.width = 1; s.height = sh; }
                else             { s.left = 0; s.top = pt; s.width = sw; s.height = 1; }
                s.name = GTAG + "_" + l.t + "_" + l.p;
                s.fill.setSolidColor("FF0000");
                s.lineFormat.visible = false;
            });
        });
        return ctx.sync().then(function () { showStatus("Hilfslinien ein ✓", "success"); });
    });
}

function rmGuides(ctx, masters) {
    var ps = [];
    masters.forEach(function (master) {
        var sh = master.shapes;
        sh.load("items");
        ps.push(ctx.sync().then(function () {
            for (var i = 0; i < sh.items.length; i++) sh.items[i].load("name");
            return ctx.sync().then(function () {
                for (var i = 0; i < sh.items.length; i++)
                    if (sh.items[i].name && sh.items[i].name.indexOf(GTAG) === 0) sh.items[i].delete();
            });
        }));
    });
    return Promise.all(ps).then(function () {
        return ctx.sync().then(function () { showStatus("Hilfslinien aus ✓", "success"); });
    });
}

/* ══════════════════════════════════════════════════════════════
   SCHATTEN-WERTE KOPIEREN
   ══════════════════════════════════════════════════════════════ */
function copyShadowText() {
    var txt = "Schatten:\nFarbe: Schwarz\nTransparenz: 75%\nGröße: 100%\nWeichzeichnen: 4pt\nWinkel: 90°\nAbstand: 1pt";
    if (navigator.clipboard && navigator.clipboard.writeText) {
        navigator.clipboard.writeText(txt).then(function () {
            showStatus("Schatten-Werte kopiert ✓", "success");
        }).catch(function () { showStatus("Kopieren fehlgeschlagen", "error"); });
    } else { showStatus("Zwischenablage nicht verfügbar", "error"); }
}

/* ═══════════════════════════════════════════════════════════════
   ███  GANTT CHART GENERATOR  ███
   ═══════════════════════════════════════════════════════════════
   Erzeugt einen Gantt-Chart als Shapes auf der aktuellen Folie.
   Alle Positionen und Größen basieren auf der Rastereinheit (RE).
   ═══════════════════════════════════════════════════════════════ */

var MONTH_NAMES = ["Jan","Feb","Mär","Apr","Mai","Jun","Jul","Aug","Sep","Okt","Nov","Dez"];
var ganttTaskCount = 3;

function initGantt() {
    /* Monat-Selects befüllen */
    var selStart = document.getElementById("gStartM");
    var selEnd   = document.getElementById("gEndM");
    MONTH_NAMES.forEach(function (m, i) {
        selStart.add(new Option(m, i));
        selEnd.add(new Option(m, i));
    });
    selStart.value = "0";   /* Jan */
    selEnd.value   = "11";  /* Dez */

    /* Initiale Task-Zeilen */
    for (var i = 0; i < ganttTaskCount; i++) addTaskRow(i);

    /* Buttons */
    document.getElementById("ganttAddTask").addEventListener("click", function () {
        addTaskRow(ganttTaskCount);
        ganttTaskCount++;
    });
    document.getElementById("ganttRmTask").addEventListener("click", function () {
        var container = document.getElementById("ganttTasks");
        if (container.children.length > 1) {
            container.removeChild(container.lastElementChild);
            ganttTaskCount--;
        }
    });
    document.getElementById("ganttBuild").addEventListener("click", function () { buildGantt(); });
}

function addTaskRow(idx) {
    var container = document.getElementById("ganttTasks");
    var div = document.createElement("div");
    div.className = "gantt-task";

    /* Name */
    var inp = document.createElement("input");
    inp.type = "text";
    inp.placeholder = "Aufgabe " + (idx + 1);
    inp.className = "gt-name";
    div.appendChild(inp);

    /* Start-Monat */
    var lS = document.createElement("span"); lS.className = "task-lbl"; lS.textContent = "S";
    div.appendChild(lS);
    var sS = document.createElement("select"); sS.className = "gt-start";
    MONTH_NAMES.forEach(function (m, i) { sS.add(new Option(m, i)); });
    sS.value = String(idx);
    div.appendChild(sS);

    /* End-Monat */
    var lE = document.createElement("span"); lE.className = "task-lbl"; lE.textContent = "E";
    div.appendChild(lE);
    var sE = document.createElement("select"); sE.className = "gt-end";
    MONTH_NAMES.forEach(function (m, i) { sE.add(new Option(m, i)); });
    sE.value = String(Math.min(idx + 2, 11));
    div.appendChild(sE);

    container.appendChild(div);
}

/* ── Gantt bauen ── */
function buildGantt() {
    if (!apiOk) { showStatus("PowerPointApi 1.5 nötig", "error"); return; }

    /* Eingaben lesen */
    var startMonth = parseInt(document.getElementById("gStartM").value);
    var startYear  = parseInt(document.getElementById("gStartY").value);
    var endMonth   = parseInt(document.getElementById("gEndM").value);
    var endYear    = parseInt(document.getElementById("gEndY").value);
    var barColor   = document.getElementById("gBarColor").value.replace("#", "");
    var barH_RE    = parseInt(document.getElementById("gBarH").value);
    var headerStyle= document.getElementById("gHeaderStyle").value;
    var gap_RE     = parseInt(document.getElementById("gGap").value);

    /* Monate berechnen */
    var totalMonths = (endYear - startYear) * 12 + (endMonth - startMonth) + 1;
    if (totalMonths < 1 || totalMonths > 36) {
        showStatus("Zeitraum: 1-36 Monate!", "warning");
        return;
    }

    /* Tasks sammeln */
    var taskEls = document.querySelectorAll(".gantt-task");
    var tasks = [];
    taskEls.forEach(function (row, i) {
        var name  = row.querySelector(".gt-name").value || ("Aufgabe " + (i + 1));
        var start = parseInt(row.querySelector(".gt-start").value);
        var end   = parseInt(row.querySelector(".gt-end").value);
        tasks.push({ name: name, start: start, end: end });
    });

    if (tasks.length === 0) { showStatus("Keine Aufgaben!", "warning"); return; }

    /* ── Layout berechnen (alles in RE) ── */
    var RE        = c2p(gridUnitCm);
    var marginL   = 8;    /* RE links (Labelbereich) */
    var marginT   = 17;   /* RE oben */
    var labelW_RE = 25;   /* RE Breite für Task-Labels */
    var headerH_RE = 3;   /* RE Höhe für Header */
    var colW_RE   = Math.max(3, Math.floor((126 - marginL - labelW_RE) / totalMonths));

    PowerPoint.run(function (ctx) {
        var sel = ctx.presentation.getSelectedSlides();
        sel.load("items");
        return ctx.sync().then(function () {
            var slidePromise;
            if (sel.items.length > 0) {
                slidePromise = Promise.resolve(sel.items[0]);
            } else {
                var slides = ctx.presentation.slides;
                slides.load("items");
                slidePromise = ctx.sync().then(function () {
                    if (!slides.items.length) return null;
                    return slides.items[0];
                });
            }
            return slidePromise;
        }).then(function (slide) {
            if (!slide) { showStatus("Keine Folie!", "error"); return ctx.sync(); }

            var x0 = marginL * RE;
            var y0 = marginT * RE;

            /* ── HEADER ── */
            if (headerStyle === "month") {
                for (var m = 0; m < totalMonths; m++) {
                    var mi = (startMonth + m) % 12;
                    var hdr = slide.shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle);
                    hdr.left   = x0 + labelW_RE * RE + m * colW_RE * RE;
                    hdr.top    = y0;
                    hdr.width  = colW_RE * RE;
                    hdr.height = headerH_RE * RE;
                    hdr.fill.setSolidColor("16213E");
                    hdr.lineFormat.color = "2A3A5C";
                    hdr.lineFormat.weight = 0.5;
                    hdr.name = "Gantt_Hdr_" + m;

                    /* Text-Label */
                    var tf = hdr.textFrame;
                    tf.textRange.text = MONTH_NAMES[mi];
                    tf.textRange.font.color = "FFFFFF";
                    tf.textRange.font.size = 8;
                    tf.textRange.font.bold = true;
                    tf.verticalAlignment = "Middle";
                    try { tf.textRange.paragraphFormat.horizontalAlignment = "Center"; } catch(e) {}
                }
            } else {
                /* Quartal-Header */
                var qStart = Math.floor(startMonth / 3);
                var m = 0;
                while (m < totalMonths) {
                    var curMi = (startMonth + m) % 12;
                    var qEnd = (Math.floor(curMi / 3) + 1) * 3;
                    var span = Math.min(qEnd - curMi, totalMonths - m);

                    var hdr = slide.shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle);
                    hdr.left   = x0 + labelW_RE * RE + m * colW_RE * RE;
                    hdr.top    = y0;
                    hdr.width  = span * colW_RE * RE;
                    hdr.height = headerH_RE * RE;
                    hdr.fill.setSolidColor("16213E");
                    hdr.lineFormat.color = "2A3A5C";
                    hdr.lineFormat.weight = 0.5;
                    hdr.name = "Gantt_QHdr_" + m;

                    var qNum = Math.floor(curMi / 3) + 1;
                    var yr = startYear + Math.floor((startMonth + m) / 12);
                    var tf = hdr.textFrame;
                    tf.textRange.text = "Q" + qNum + " " + yr;
                    tf.textRange.font.color = "FFFFFF";
                    tf.textRange.font.size = 8;
                    tf.textRange.font.bold = true;
                    tf.verticalAlignment = "Middle";
                    try { tf.textRange.paragraphFormat.horizontalAlignment = "Center"; } catch(e) {}

                    m += span;
                }
            }

            /* ── TASK ROWS ── */
            for (var t = 0; t < tasks.length; t++) {
                var task = tasks[t];
                var rowY = y0 + (headerH_RE + gap_RE) * RE + t * (barH_RE + gap_RE) * RE;

                /* Label-Box */
                var lbl = slide.shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle);
                lbl.left   = x0;
                lbl.top    = rowY;
                lbl.width  = labelW_RE * RE;
                lbl.height = barH_RE * RE;
                lbl.fill.setSolidColor("0A0E1A");
                lbl.lineFormat.color = "2A3A5C";
                lbl.lineFormat.weight = 0.3;
                lbl.name = "Gantt_Lbl_" + t;

                var ltf = lbl.textFrame;
                ltf.textRange.text = task.name;
                ltf.textRange.font.color = "FFFFFF";
                ltf.textRange.font.size = 7;
                ltf.verticalAlignment = "Middle";
                try { ltf.textRange.paragraphFormat.horizontalAlignment = "Left"; } catch(e) {}

                /* Balken berechnen */
                var barStart = task.start - startMonth;
                var barEnd   = task.end - startMonth;
                if (barStart < 0) barStart = 0;
                if (barEnd >= totalMonths) barEnd = totalMonths - 1;
                var barSpan = barEnd - barStart + 1;

                if (barSpan > 0) {
                    var bar = slide.shapes.addGeometricShape(PowerPoint.GeometricShapeType.roundedRectangle);
                    bar.left   = x0 + labelW_RE * RE + barStart * colW_RE * RE;
                    bar.top    = rowY;
                    bar.width  = barSpan * colW_RE * RE;
                    bar.height = barH_RE * RE;
                    bar.fill.setSolidColor(barColor);
                    bar.lineFormat.visible = false;
                    bar.name = "Gantt_Bar_" + t;
                }

                /* Grid-Linien (vertikale Trenner) */
                for (var m = 0; m <= totalMonths; m++) {
                    var gl = slide.shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle);
                    gl.left   = x0 + labelW_RE * RE + m * colW_RE * RE;
                    gl.top    = rowY;
                    gl.width  = 0.5;
                    gl.height = barH_RE * RE;
                    gl.fill.setSolidColor("2A3A5C");
                    gl.lineFormat.visible = false;
                    gl.name = "Gantt_GL_" + t + "_" + m;
                }
            }

            return ctx.sync().then(function () {
                showStatus("Gantt (" + tasks.length + " Tasks, " + totalMonths + " Mon.) ✓", "success");
            });
        });
    }).catch(function (e) { showStatus("Fehler: " + e.message, "error"); });
}
