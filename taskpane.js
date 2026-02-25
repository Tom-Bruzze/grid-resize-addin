/* ═══════════════════════════════════════════════════════════════
   DROEGE Grid Resize Tool – taskpane.js
   ═══════════════════════════════════════════════════════════════
   2 Tabs: Tools (Größe + Angleich + Raster + Abstände) | Extras
   ═══════════════════════════════════════════════════════════════ */

var CM = 28.3465;
var MIN = 0.1;
var gridUnitCm = 0.21;
var apiOk = false;
var GTAG = "DROEGE_GUIDELINE";

/* ── Office Init ── */
Office.onReady(function (info) {
    if (info.host === Office.HostType.PowerPoint) {
        if (Office.context.requirements && Office.context.requirements.isSetSupported) {
            apiOk = Office.context.requirements.isSetSupported("PowerPointApi", "1.5");
        } else {
            apiOk = (typeof PowerPoint !== "undefined" && PowerPoint.run && typeof PowerPoint.run === "function");
        }
        initUI();
        if (!apiOk) showStatus("PowerPointApi 1.5 nicht verfügbar", "warning");
    }
});

/* ══════════════════════════════════════════════════════════════
   UI INIT
   ══════════════════════════════════════════════════════════════ */
function initUI() {

    /* Rastereinheit: Input */
    var gi = document.getElementById("gridUnit");
    gi.addEventListener("change", function () {
        var v = parseFloat(this.value);
        if (!isNaN(v) && v > 0) {
            gridUnitCm = v;
            hlPre(v);
            showStatus("RE: " + v.toFixed(2) + " cm", "info");
        }
    });

    /* Rastereinheit: Presets */
    document.querySelectorAll(".pre").forEach(function (b) {
        b.addEventListener("click", function () {
            var v = parseFloat(this.dataset.value);
            gridUnitCm = v;
            gi.value = v;
            hlPre(v);
            showStatus("RE: " + v.toFixed(2) + " cm", "info");
        });
    });

    /* Tabs */
    document.querySelectorAll(".tab").forEach(function (b) {
        b.addEventListener("click", function () {
            var id = this.dataset.tab;
            document.querySelectorAll(".tab").forEach(function (t) { t.classList.remove("active"); });
            document.querySelectorAll(".pane").forEach(function (p) { p.classList.remove("active"); });
            this.classList.add("active");
            document.getElementById(id).classList.add("active");
        });
    });

    /* ── Größe (Klick = +, Shift = −) ── */
    shiftBind("resizeW",    function () { resize("width",   gridUnitCm); },
                            function () { resize("width",  -gridUnitCm); });
    shiftBind("resizeH",    function () { resize("height",  gridUnitCm); },
                            function () { resize("height", -gridUnitCm); });
    shiftBind("resizeBoth", function () { resize("both",    gridUnitCm); },
                            function () { resize("both",   -gridUnitCm); });
    shiftBind("resizeProp", function () { propResize( gridUnitCm); },
                            function () { propResize(-gridUnitCm); });

    /* ── Angleichen (Klick = Max, Shift = Min) ── */
    shiftBind("matchW",    function () { matchDim("width",  "max"); },
                           function () { matchDim("width",  "min"); });
    shiftBind("matchH",    function () { matchDim("height", "max"); },
                           function () { matchDim("height", "min"); });
    shiftBind("matchBoth", function () { matchDim("both",   "max"); },
                           function () { matchDim("both",   "min"); });
    shiftBind("matchProp", function () { propMatch("max"); },
                           function () { propMatch("min"); });

    /* ── Raster ── */
    bind("snapPos",     function () { snap("position"); });
    bind("snapSize",    function () { snap("size");     });
    bind("snapAll",     function () { snap("both");     });
    bind("showInfo",    function () { shapeInfo();      });

    /* ── Abstände ── */
    bind("spaceH",      function () { spacing("horizontal"); });
    bind("spaceV",      function () { spacing("vertical");   });

    /* ── Grid-Tabelle ── */
    bind("createTable", function () { createGridTable(); });

    /* ── Extras ── */
    bind("detectFmt",    function () { detectFormat();   });
    bind("toggleGuides", function () { toggleGuides();   });
    bind("copyShadow",   function () { copyShadowText(); });
}

/* ══════════════════════════════════════════════════════════════
   HILFSFUNKTIONEN
   ══════════════════════════════════════════════════════════════ */
function shiftBind(id, fnNormal, fnShift) {
    var el = document.getElementById(id);
    if (!el) return;
    el.addEventListener("click", function (e) { e.shiftKey ? fnShift() : fnNormal(); });
}

function bind(id, fn) {
    var el = document.getElementById(id);
    if (!el) return;
    el.addEventListener("click", fn);
}

function hlPre(val) {
    document.querySelectorAll(".pre").forEach(function (b) {
        b.classList.toggle("active", Math.abs(parseFloat(b.dataset.value) - val) < 0.001);
    });
}

function showStatus(msg, type) {
    var el = document.getElementById("status");
    el.textContent = msg;
    el.className = "sts " + (type || "info");
}

function c2p(cm) { return cm * CM; }
function p2c(pt) { return pt / CM; }
function rnd(v)  { return Math.round(v / gridUnitCm) * gridUnitCm; }

function getTol() {
    var t = c2p(gridUnitCm) * 0.5;
    return t < 5 ? 5 : t;
}

/* Berechnet den Raster-Offset (halber Rand) für eine Foliendimension.
   Robuster Algorithmus: Statt fehleranfälligem Floating-Point-Modulo
   wird die Anzahl passender Rastereinheiten gerundet.
   Funktioniert korrekt für alle Formate (16:9, 4:3, A4, A3, Letter, etc.). */
/* ── Feste Raster-Offsets (0,21 cm) – gemessen in PowerPoint ── */
var GRID_OFFSETS = [
    { name: "16:9",      w: 960.0,   h: 540.0,   ox: 0.10, oy: 0.00 },
    { name: "4:3",       w: 914.4,   h: 685.8,   ox: 0.10, oy: 0.069 },
    { name: "16:10",     w: 720.0,   h: 450.0,   ox: 0.10, oy: 0.17 },
    { name: "A4 quer",   w: 841.89,  h: 595.28,  ox: 0.11, oy: 0.07 },
    { name: "Breitbild", w: 786.0,   h: 547.0,   ox: 0.13, oy: 0.07 }
];
function getGridOffsets(slideW, slideH) {
    var bestIdx = -1, bestDiff = 999;
    for (var i = 0; i < GRID_OFFSETS.length; i++) {
        var f = GRID_OFFSETS[i];
        var d = Math.abs(slideW - f.w) + Math.abs(slideH - f.h);
        if (d < bestDiff) { bestDiff = d; bestIdx = i; }
    }
    if (bestIdx >= 0 && bestDiff < 10) {
        var f = GRID_OFFSETS[bestIdx];
        return { x: c2p(f.ox), y: c2p(f.oy), name: f.name };
    }
    return { x: 0, y: 0, name: "Unbekannt" };
}

function withShapes(min, cb) {
    if (!apiOk) { showStatus("PowerPointApi 1.5 nötig", "error"); return; }
    PowerPoint.run(function (ctx) {
        var shapes = ctx.presentation.getSelectedShapes();
        shapes.load("items");
        return ctx.sync().then(function () {
            if (shapes.items.length < min) {
                showStatus(min <= 1 ? "Bitte Objekt(e) auswählen!" : "Mind. " + min + " Objekte!", "error");
                return;
            }
            return cb(ctx, shapes.items);
        });
    }).catch(function (e) { showStatus("Fehler: " + e.message, "error"); });
}

/* ══════════════════════════════════════════════════════════════
   GRUPPIERUNG: Multi-Row / Multi-Column
   ══════════════════════════════════════════════════════════════ */
function groupByPos(items, axis, tol) {
    var groups = [], used = {};
    var sorted = items.slice().sort(function (a, b) {
        return axis === "y" ? (a.top - b.top) : (a.left - b.left);
    });
    for (var i = 0; i < sorted.length; i++) {
        if (used[i]) continue;
        var g = [sorted[i]]; used[i] = true;
        var ref = axis === "y" ? sorted[i].top : sorted[i].left;
        for (var j = i + 1; j < sorted.length; j++) {
            if (used[j]) continue;
            var p = axis === "y" ? sorted[j].top : sorted[j].left;
            if (Math.abs(p - ref) <= tol) { g.push(sorted[j]); used[j] = true; }
        }
        groups.push(g);
    }
    return groups;
}

/* ══════════════════════════════════════════════════════════════
   RESIZE – Multi-Row / Multi-Column
   ══════════════════════════════════════════════════════════════ */
function resize(dim, deltaCm) {
    withShapes(1, function (ctx, items) {
        items.forEach(function (s) { s.load(["left", "top", "width", "height"]); });
        return ctx.sync().then(function () {
            var dp = c2p(Math.abs(deltaCm)), grow = deltaCm > 0;

            if (items.length === 1) {
                var s = items[0];
                if (dim === "width" || dim === "both") {
                    var nw = grow ? s.width + dp : s.width - dp;
                    if (nw >= c2p(MIN)) s.width = nw;
                }
                if (dim === "height" || dim === "both") {
                    var nh = grow ? s.height + dp : s.height - dp;
                    if (nh >= c2p(MIN)) s.height = nh;
                }
                return ctx.sync().then(function () {
                    var l = dim === "both" ? "W+H" : dim === "width" ? "Breite" : "Höhe";
                    showStatus(l + (grow ? " +" : " −") + Math.abs(deltaCm).toFixed(2) + " cm ✓", "success");
                });
            }

            var tol = getTol(), ri = "", ci = "";

            if (dim === "width" || dim === "both") {
                var rows = groupByPos(items, "y", tol);
                ri = rows.length + " Zl";
                rows.forEach(function (row) {
                    row.sort(function (a, b) { return a.left - b.left; });
                    var gaps = [];
                    for (var i = 0; i < row.length - 1; i++)
                        gaps.push(row[i + 1].left - (row[i].left + row[i].width));
                    var ok = true;
                    for (var i = 0; i < row.length; i++) {
                        if ((grow ? row[i].width + dp : row[i].width - dp) < c2p(MIN)) { ok = false; break; }
                    }
                    if (ok) {
                        for (var i = 0; i < row.length; i++)
                            row[i].width = grow ? row[i].width + dp : row[i].width - dp;
                        for (var i = 1; i < row.length; i++)
                            row[i].left = row[i - 1].left + row[i - 1].width + gaps[i - 1];
                    }
                });
            }

            if (dim === "height" || dim === "both") {
                var cols = groupByPos(items, "x", tol);
                ci = cols.length + " Sp";
                cols.forEach(function (col) {
                    col.sort(function (a, b) { return a.top - b.top; });
                    var gaps = [];
                    for (var i = 0; i < col.length - 1; i++)
                        gaps.push(col[i + 1].top - (col[i].top + col[i].height));
                    var ok = true;
                    for (var i = 0; i < col.length; i++) {
                        if ((grow ? col[i].height + dp : col[i].height - dp) < c2p(MIN)) { ok = false; break; }
                    }
                    if (ok) {
                        for (var i = 0; i < col.length; i++)
                            col[i].height = grow ? col[i].height + dp : col[i].height - dp;
                        for (var i = 1; i < col.length; i++)
                            col[i].top = col[i - 1].top + col[i - 1].height + gaps[i - 1];
                    }
                });
            }

            return ctx.sync().then(function () {
                var l = dim === "both" ? "W+H" : dim === "width" ? "Breite" : "Höhe";
                var d = ri && ci ? " · " + ri + " · " + ci : ri ? " · " + ri : ci ? " · " + ci : "";
                showStatus(l + (grow ? " +" : " −") + Math.abs(deltaCm).toFixed(2) + " cm" + d + " ✓", "success");
            });
        });
    });
}

/* ══════════════════════════════════════════════════════════════
   PROPORTIONAL RESIZE
   ══════════════════════════════════════════════════════════════ */
function propResize(deltaCm) {
    withShapes(1, function (ctx, items) {
        items.forEach(function (s) { s.load(["left", "top", "width", "height"]); });
        return ctx.sync().then(function () {
            var dp = c2p(Math.abs(deltaCm)), grow = deltaCm > 0;

            if (items.length === 1) {
                var s = items[0], r = s.height / s.width;
                var nw = grow ? s.width + dp : s.width - dp;
                if (nw >= c2p(MIN) && nw * r >= c2p(MIN)) { s.width = nw; s.height = nw * r; }
                return ctx.sync().then(function () {
                    showStatus("Prop " + (grow ? "+" : "−") + Math.abs(deltaCm).toFixed(2) + " cm ✓", "success");
                });
            }

            var tol = getTol();
            var data = items.map(function (s) {
                return { shape: s, ratio: s.height / s.width, left: s.left, top: s.top, width: s.width, height: s.height, nw: 0, nh: 0 };
            });

            var ok = true;
            data.forEach(function (o) {
                var nw = grow ? o.width + dp : o.width - dp;
                if (nw < c2p(MIN) || nw * o.ratio < c2p(MIN)) ok = false;
            });
            if (!ok) { showStatus("Mindestgröße erreicht!", "error"); return ctx.sync(); }

            data.forEach(function (o) {
                o.nw = grow ? o.width + dp : o.width - dp;
                o.nh = o.nw * o.ratio;
                o.shape.width = o.nw;
                o.shape.height = o.nh;
            });

            var rows = groupByPos(data, "y", tol);
            rows.forEach(function (row) {
                row.sort(function (a, b) { return a.left - b.left; });
                var gaps = [];
                for (var i = 0; i < row.length - 1; i++)
                    gaps.push(row[i + 1].left - (row[i].left + row[i].width));
                for (var i = 1; i < row.length; i++)
                    row[i].shape.left = row[i - 1].shape.left + row[i - 1].nw + gaps[i - 1];
            });

            var cols = groupByPos(data, "x", tol);
            cols.forEach(function (col) {
                col.sort(function (a, b) { return a.top - b.top; });
                var gaps = [];
                for (var i = 0; i < col.length - 1; i++)
                    gaps.push(col[i + 1].top - (col[i].top + col[i].height));
                for (var i = 1; i < col.length; i++)
                    col[i].shape.top = col[i - 1].shape.top + col[i - 1].nh + gaps[i - 1];
            });

            return ctx.sync().then(function () {
                showStatus("Prop " + (grow ? "+" : "−") + " · " + rows.length + " Zl · " + cols.length + " Sp ✓", "success");
            });
        });
    });
}

/* ══════════════════════════════════════════════════════════════
   SNAP TO GRID
   ══════════════════════════════════════════════════════════════ */
function snap(mode) {
    if (!apiOk) { showStatus("PowerPointApi 1.5 nötig", "error"); return; }

    PowerPoint.run(function (ctx) {
        var sel = ctx.presentation.getSelectedShapes();
        sel.load("items");
        var ps = ctx.presentation.pageSetup;
        ps.load(["slideWidth", "slideHeight"]);

        return ctx.sync().then(function () {
            var items = sel.items;
            if (!items || items.length < 1) {
                showStatus("Mind. 1 Objekt auswählen!", "error");
                return;
            }

            for (var i = 0; i < items.length; i++) {
                items[i].load(["left", "top", "width", "height"]);
            }
            return ctx.sync();
        }).then(function () {
            var items = sel.items;
            if (!items || items.length < 1) return;

            /* ── Raster-Offset (Lookup) ── */
            var gPt = c2p(gridUnitCm);
            var off = getGridOffsets(ps.slideWidth, ps.slideHeight);
            var offsetX = off.x;
            var offsetY = off.y;

            for (var i = 0; i < items.length; i++) {
                var s = items[i];
                if (mode === "position" || mode === "both") {
                    s.left = offsetX + Math.round((s.left - offsetX) / gPt) * gPt;
                    s.top  = offsetY + Math.round((s.top  - offsetY) / gPt) * gPt;
                }
                if (mode === "size" || mode === "both") {
                    var nw = Math.round(s.width  / gPt) * gPt;
                    var nh = Math.round(s.height / gPt) * gPt;
                    if (nw >= c2p(MIN)) s.width  = nw;
                    if (nh >= c2p(MIN)) s.height = nh;
                }
            }

            return ctx.sync().then(function () {
                var l = mode === "both" ? "Pos+Size" : mode === "position" ? "Position" : "Größe";
                showStatus(l + " → " + off.name + " ✓ (X:" + p2c(offsetX).toFixed(2) +
                    " Y:" + p2c(offsetY).toFixed(2) + " cm)", "success");
            });
        });
    }).catch(function (e) {
        showStatus("Snap-Fehler: " + e.message, "error");
    });
}

/* ══════════════════════════════════════════════════════════════
   SPACING – Horizontal & Vertical (REWRITE v3)
   ══════════════════════════════════════════════════════════════
   Komplett neu geschrieben:
   - Verwendet einen eigenen PowerPoint.run() Aufruf
   - Liest zuerst ALLE Properties, speichert sie in einem
     lokalen Array, berechnet neue Positionen, schreibt sie
     zurück und synchronisiert einmal am Ende.
   - Kein verschachteltes ctx.sync() / return-Problem mehr.
   ══════════════════════════════════════════════════════════════ */
function spacing(dir) {
    if (!apiOk) { showStatus("PowerPointApi 1.5 nötig", "error"); return; }

    PowerPoint.run(function (ctx) {
        var sel = ctx.presentation.getSelectedShapes();
        sel.load("items");
        var ps = ctx.presentation.pageSetup;
        ps.load(["slideWidth", "slideHeight"]);

        return ctx.sync().then(function () {
            var items = sel.items;
            if (items.length < 2) {
                showStatus("Mind. 2 Objekte auswählen!", "error");
                return;
            }

            /* Schritt 1: Alle relevanten Properties laden */
            for (var i = 0; i < items.length; i++) {
                items[i].load(["left", "top", "width", "height"]);
            }

            return ctx.sync();
        }).then(function () {
            var items = sel.items;
            if (!items || items.length < 2) return;

            var sp = c2p(gridUnitCm);
            var tol = getTol();

            /* Raster-Offset (Lookup) */
            var gPt = c2p(gridUnitCm);
            var off = getGridOffsets(ps.slideWidth, ps.slideHeight);
            var offsetX = off.x;
            var offsetY = off.y;

            /* Schritt 2: Lokale Kopie der Daten erstellen */
            var data = [];
            for (var i = 0; i < items.length; i++) {
                data.push({
                    idx:    i,
                    shape:  items[i],
                    left:   items[i].left,
                    top:    items[i].top,
                    width:  items[i].width,
                    height: items[i].height
                });
            }

            var groupCount = 0;
            var movedCount = 0;

            if (dir === "horizontal") {
                /* Shapes nach Y-Position gruppieren (= Zeilen) */
                var rows = groupByData(data, "top", tol);
                groupCount = rows.length;

                for (var r = 0; r < rows.length; r++) {
                    var row = rows[r];
                    if (row.length < 2) continue;

                    /* Links nach rechts sortieren */
                    row.sort(function (a, b) { return a.left - b.left; });

                    /* Erstes Shape ins Raster einrasten */
                    var snappedLeft = offsetX + Math.round((row[0].left - offsetX) / gPt) * gPt;
                    row[0].left = snappedLeft;
                    row[0].shape.left = snappedLeft;

                    /* Alle weiteren werden repositioniert */
                    for (var i = 1; i < row.length; i++) {
                        var newLeft = row[i - 1].left + row[i - 1].width + sp;
                        row[i].left = newLeft;
                        row[i].shape.left = newLeft;
                        movedCount++;
                    }
                }
            }

            if (dir === "vertical") {
                /* Shapes nach X-Position gruppieren (= Spalten) */
                var cols = groupByData(data, "left", tol);
                groupCount = cols.length;

                for (var c = 0; c < cols.length; c++) {
                    var col = cols[c];
                    if (col.length < 2) continue;

                    /* Oben nach unten sortieren */
                    col.sort(function (a, b) { return a.top - b.top; });

                    /* Erstes Shape ins Raster einrasten */
                    var snappedTop = offsetY + Math.round((col[0].top - offsetY) / gPt) * gPt;
                    col[0].top = snappedTop;
                    col[0].shape.top = snappedTop;

                    /* Alle weiteren werden repositioniert */
                    for (var i = 1; i < col.length; i++) {
                        var newTop = col[i - 1].top + col[i - 1].height + sp;
                        col[i].top = newTop;
                        col[i].shape.top = newTop;
                        movedCount++;
                    }
                }
            }

            /* Schritt 3: Einmal am Ende synchronisieren */
            return ctx.sync().then(function () {
                var label = dir === "horizontal" ? "H" : "V";
                var gLabel = dir === "horizontal" ? " Zl" : " Sp";
                showStatus(label + "-Abstand " + gridUnitCm.toFixed(2) + " cm · " +
                    groupCount + gLabel + " · " + movedCount + " verschoben ✓", "success");
            });
        });
    }).catch(function (e) {
        showStatus("Spacing-Fehler: " + e.message, "error");
    });
}

/* ──────────────────────────────────────────────────────────────
   Hilfs-Gruppierung für Spacing (arbeitet mit lokalen Daten)
   Gruppiert data-Objekte nach einer Property (top oder left)
   mit einer Toleranz.
   ────────────────────────────────────────────────────────────── */
function groupByData(data, prop, tol) {
    var groups = [];
    var used = [];
    for (var i = 0; i < data.length; i++) used[i] = false;

    /* Sortieren nach der Gruppen-Property */
    var sorted = data.slice().sort(function (a, b) { return a[prop] - b[prop]; });

    /* Mapping: sortierte Indizes zu used-Indizes */
    var idxMap = [];
    for (var i = 0; i < sorted.length; i++) {
        for (var j = 0; j < data.length; j++) {
            if (sorted[i].idx === data[j].idx) { idxMap[i] = j; break; }
        }
    }

    for (var i = 0; i < sorted.length; i++) {
        if (used[idxMap[i]]) continue;
        var g = [sorted[i]];
        used[idxMap[i]] = true;
        var ref = sorted[i][prop];

        for (var j = i + 1; j < sorted.length; j++) {
            if (used[idxMap[j]]) continue;
            if (Math.abs(sorted[j][prop] - ref) <= tol) {
                g.push(sorted[j]);
                used[idxMap[j]] = true;
            }
        }
        groups.push(g);
    }
    return groups;
}

/* ══════════════════════════════════════════════════════════════
   SHAPE INFO
   ══════════════════════════════════════════════════════════════ */
function shapeInfo() {
    withShapes(1, function (ctx, items) {
        items.forEach(function (s) { s.load(["name", "left", "top", "width", "height"]); });
        return ctx.sync().then(function () {
            var el = document.getElementById("infoBox"), html = "";
            items.forEach(function (s, idx) {
                if (items.length > 1)
                    html += '<div style="font-weight:700;margin-top:' + (idx > 0 ? '6' : '0') +
                        'px;color:#e94560;font-size:11px;">' + (s.name || "Objekt " + (idx + 1)) + '</div>';
                html += '<div class="info-item"><span class="info-label">Breite:</span>' +
                    '<span class="info-value">' + p2c(s.width).toFixed(2) + ' cm (' + (p2c(s.width) / gridUnitCm).toFixed(1) + ' RE)</span></div>';
                html += '<div class="info-item"><span class="info-label">Höhe:</span>' +
                    '<span class="info-value">' + p2c(s.height).toFixed(2) + ' cm (' + (p2c(s.height) / gridUnitCm).toFixed(1) + ' RE)</span></div>';
                html += '<div class="info-item"><span class="info-label">Links:</span>' +
                    '<span class="info-value">' + p2c(s.left).toFixed(2) + ' cm</span></div>';
                html += '<div class="info-item"><span class="info-label">Oben:</span>' +
                    '<span class="info-value">' + p2c(s.top).toFixed(2) + ' cm</span></div>';
            });
            el.innerHTML = html;
            el.classList.add("visible");
            showStatus("Objektinfo ✓", "info");
        });
    });
}

/* ══════════════════════════════════════════════════════════════
   MATCH DIMENSIONS
   ══════════════════════════════════════════════════════════════ */
function matchDim(dim, mode) {
    withShapes(2, function (ctx, items) {
        items.forEach(function (s) { s.load(["width", "height"]); });
        return ctx.sync().then(function () {
            var ws = items.map(function (s) { return s.width;  });
            var hs = items.map(function (s) { return s.height; });
            var tw = mode === "max" ? Math.max.apply(null, ws) : Math.min.apply(null, ws);
            var th = mode === "max" ? Math.max.apply(null, hs) : Math.min.apply(null, hs);
            items.forEach(function (s) {
                if (dim === "width"  || dim === "both") s.width  = tw;
                if (dim === "height" || dim === "both") s.height = th;
            });
            return ctx.sync().then(function () {
                var l = dim === "both" ? "W+H" : dim === "width" ? "W" : "H";
                showStatus(l + " → " + (mode === "max" ? "Max" : "Min") + " ✓", "success");
            });
        });
    });
}

/* ══════════════════════════════════════════════════════════════
   PROPORTIONAL MATCH
   ══════════════════════════════════════════════════════════════ */
function propMatch(mode) {
    withShapes(2, function (ctx, items) {
        items.forEach(function (s) { s.load(["width", "height"]); });
        return ctx.sync().then(function () {
            var ws = items.map(function (s) { return s.width; });
            var tw = mode === "max" ? Math.max.apply(null, ws) : Math.min.apply(null, ws);
            items.forEach(function (s) {
                var r = s.height / s.width;
                s.width = tw;
                s.height = tw * r;
            });
            return ctx.sync().then(function () {
                showStatus("Prop → " + (mode === "max" ? "Max" : "Min") + " ✓", "success");
            });
        });
    });
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
function detectFormat() {
    PowerPoint.run(function (ctx) {
        var ps = ctx.presentation.pageSetup;
        ps.load(["slideWidth", "slideHeight"]);
        return ctx.sync().then(function () {
            var off = getGridOffsets(ps.slideWidth, ps.slideHeight);
            showStatus("Erkannt: " + off.name + " | " +
                (ps.slideWidth/28.3465).toFixed(2) + " × " +
                (ps.slideHeight/28.3465).toFixed(2) + " cm (" +
                ps.slideWidth.toFixed(1) + " × " + ps.slideHeight.toFixed(1) + " pt)", "success");
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
