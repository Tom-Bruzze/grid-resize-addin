/* ═══════════════════════════════════════════════════════════
   DROEGE Grid Resize Tool – Complete JavaScript
   Compact Edition mit Shift-Click & Multi-Row/Col Spacing
   ═══════════════════════════════════════════════════════════ */

var CM = 28.3465;
var MIN = 0.1;
var gridUnitCm = 0.21;
var apiOk = false;
var GTAG = "DROEGE_GUIDELINE";

/* ── Office Init ── */
Office.onReady(function(info) {
    if (info.host === Office.HostType.PowerPoint) {
        apiOk = !!(
            Office.context.requirements &&
            Office.context.requirements.isSetSupported &&
            Office.context.requirements.isSetSupported("PowerPointApi", "1.5")
        );
        initUI();
        if (!apiOk) showStatus("PowerPointApi 1.5 nicht verfuegbar", "warning");
    }
});

/* ══════════════════════════════════════════════
   UI INIT
   ══════════════════════════════════════════════ */
function initUI() {

    // ── Grid Unit Input ──
    var gi = document.getElementById("gridUnit");
    gi.addEventListener("change", function() {
        var v = parseFloat(this.value);
        if (!isNaN(v) && v > 0) {
            gridUnitCm = v;
            upPre(v);
            showStatus("RE: " + v.toFixed(2) + " cm", "info");
        }
    });

    // ── Preset Buttons ──
    document.querySelectorAll(".pre").forEach(function(b) {
        b.addEventListener("click", function() {
            var v = parseFloat(this.dataset.value);
            gridUnitCm = v;
            gi.value = v;
            upPre(v);
            showStatus("RE: " + v.toFixed(2) + " cm", "info");
        });
    });

    // ── Tab Navigation ──
    document.querySelectorAll(".tab").forEach(function(b) {
        b.addEventListener("click", function() {
            var id = this.dataset.tab;
            document.querySelectorAll(".tab").forEach(function(t) { t.classList.remove("active"); });
            document.querySelectorAll(".pane").forEach(function(p) { p.classList.remove("active"); });
            this.classList.add("active");
            document.getElementById(id).classList.add("active");
        });
    });

    // ── TAB 1: Resize (Klick = grow, Shift+Klick = shrink) ──
    shiftBind("resizeW",    function() { resize("width",  gridUnitCm); },  function() { resize("width",  -gridUnitCm); });
    shiftBind("resizeH",    function() { resize("height", gridUnitCm); },  function() { resize("height", -gridUnitCm); });
    shiftBind("resizeBoth", function() { resize("both",   gridUnitCm); },  function() { resize("both",   -gridUnitCm); });
    shiftBind("resizeProp", function() { propResize(gridUnitCm); },        function() { propResize(-gridUnitCm); });

    // ── TAB 1: Match (Klick = max, Shift+Klick = min) ──
    shiftBind("matchW",    function() { match("width",  "max"); }, function() { match("width",  "min"); });
    shiftBind("matchH",    function() { match("height", "max"); }, function() { match("height", "min"); });
    shiftBind("matchBoth", function() { match("both",   "max"); }, function() { match("both",   "min"); });
    shiftBind("matchProp", function() { propMatch("max"); },       function() { propMatch("min"); });

    // ── TAB 2: Grid ──
    bind("snapPos",     function() { snap("position"); });
    bind("snapSize",    function() { snap("size"); });
    bind("snapAll",     function() { snap("both"); });
    bind("spaceH",      function() { spacing("horizontal"); });
    bind("spaceV",      function() { spacing("vertical"); });
    bind("showInfo",    function() { shapeInfo(); });
    bind("createTable", function() { createGridTable(); });

    // ── TAB 3: Setup ──
    bind("setSlide",     function() { setSlideSize(); });
    bind("toggleGuides", function() { toggleGuides(); });
    bind("copyShadow",   function() { copyShadow(); });
}

/* ══════════════════════════════════════════════
   HELPERS
   ══════════════════════════════════════════════ */

/* Shift-aware binding: normal click = fnNormal, shift+click = fnShift */
function shiftBind(id, fnNormal, fnShift) {
    document.getElementById(id).addEventListener("click", function(e) {
        e.shiftKey ? fnShift() : fnNormal();
    });
}

function bind(id, fn) {
    document.getElementById(id).addEventListener("click", fn);
}

function upPre(v) {
    document.querySelectorAll(".pre").forEach(function(b) {
        b.classList.toggle("active", Math.abs(parseFloat(b.dataset.value) - v) < 0.001);
    });
}

function showStatus(m, t) {
    var e = document.getElementById("status");
    e.textContent = m;
    e.className = "sts " + (t || "info");
}

function c2p(c) { return c * CM; }
function p2c(p) { return p / CM; }
function rnd(v) { return Math.round(v / gridUnitCm) * gridUnitCm; }

/* Load selected shapes with minimum count check */
function withShapes(min, cb) {
    if (!apiOk) {
        showStatus("Nicht unterstuetzt (API 1.5 noetig)", "error");
        return;
    }
    PowerPoint.run(function(ctx) {
        var sh = ctx.presentation.getSelectedShapes();
        sh.load("items");
        return ctx.sync().then(function() {
            if (sh.items.length < min) {
                showStatus(min <= 1 ? "Bitte Objekt(e) auswaehlen!" : "Min. " + min + " Objekte auswaehlen!", "error");
                return;
            }
            return cb(ctx, sh.items);
        });
    }).catch(function(e) {
        showStatus("Fehler: " + e.message, "error");
    });
}

/* ══════════════════════════════════════════════
   RESIZE (Klick = +1RE, Shift = -1RE)
   ══════════════════════════════════════════════ */
function resize(dim, d) {
    withShapes(1, function(ctx, items) {
        items.forEach(function(s) { s.load(["left", "top", "width", "height"]); });
        return ctx.sync().then(function() {
            var dp = c2p(Math.abs(d));
            var g = d > 0;

            // Single shape: simple resize
            if (items.length === 1) {
                var s = items[0];
                if (dim === "width" || dim === "both") {
                    var nw = g ? s.width + dp : s.width - dp;
                    if (nw >= c2p(MIN)) s.width = nw;
                }
                if (dim === "height" || dim === "both") {
                    var nh = g ? s.height + dp : s.height - dp;
                    if (nh >= c2p(MIN)) s.height = nh;
                }
                return ctx.sync().then(function() {
                    showStatus((dim === "both" ? "W+H" : dim === "width" ? "W" : "H") + (g ? " +" : " -") + Math.abs(d).toFixed(2) + " cm", "success");
                });
            }

            // Multiple shapes: resize + preserve gaps
            if (dim === "width" || dim === "both") {
                var hs = items.slice().sort(function(a, b) { return a.left - b.left; });
                var hg = [];
                for (var i = 0; i < hs.length - 1; i++) {
                    hg.push(hs[i + 1].left - (hs[i].left + hs[i].width));
                }
                var ok = true;
                for (var i = 0; i < hs.length; i++) {
                    if ((g ? hs[i].width + dp : hs[i].width - dp) < c2p(MIN)) { ok = false; break; }
                }
                if (ok) {
                    for (var i = 0; i < hs.length; i++) hs[i].width = g ? hs[i].width + dp : hs[i].width - dp;
                    for (var i = 1; i < hs.length; i++) hs[i].left = hs[i - 1].left + hs[i - 1].width + hg[i - 1];
                }
            }
            if (dim === "height" || dim === "both") {
                var vs = items.slice().sort(function(a, b) { return a.top - b.top; });
                var vg = [];
                for (var i = 0; i < vs.length - 1; i++) {
                    vg.push(vs[i + 1].top - (vs[i].top + vs[i].height));
                }
                var ok = true;
                for (var i = 0; i < vs.length; i++) {
                    if ((g ? vs[i].height + dp : vs[i].height - dp) < c2p(MIN)) { ok = false; break; }
                }
                if (ok) {
                    for (var i = 0; i < vs.length; i++) vs[i].height = g ? vs[i].height + dp : vs[i].height - dp;
                    for (var i = 1; i < vs.length; i++) vs[i].top = vs[i - 1].top + vs[i - 1].height + vg[i - 1];
                }
            }

            return ctx.sync().then(function() {
                showStatus((dim === "both" ? "W+H" : dim === "width" ? "W" : "H") + (g ? " +" : " -") + " Abstaende OK", "success");
            });
        });
    });
}

/* ══════════════════════════════════════════════
   PROPORTIONAL RESIZE
   ══════════════════════════════════════════════ */
function propResize(d) {
    withShapes(1, function(ctx, items) {
        items.forEach(function(s) { s.load(["left", "top", "width", "height"]); });
        return ctx.sync().then(function() {
            var dp = c2p(Math.abs(d));
            var g = d > 0;

            // Single shape
            if (items.length === 1) {
                var s = items[0];
                var r = s.height / s.width;
                var nw = g ? s.width + dp : s.width - dp;
                if (nw >= c2p(MIN)) {
                    var nh = nw * r;
                    if (nh >= c2p(MIN)) { s.width = nw; s.height = nh; }
                }
                return ctx.sync().then(function() {
                    showStatus("Prop " + (g ? "+" : "-") + Math.abs(d).toFixed(2) + " cm", "success");
                });
            }

            // Multiple shapes
            var orig = items.map(function(s) {
                return { shape: s, left: s.left, top: s.top, width: s.width, height: s.height, ratio: s.height / s.width };
            });

            var ok = true;
            orig.forEach(function(o) {
                var nw = g ? o.width + dp : o.width - dp;
                if (nw < c2p(MIN) || nw * o.ratio < c2p(MIN)) ok = false;
            });
            if (!ok) { showStatus("Mindestgroesse!", "error"); return ctx.sync(); }

            var hs = orig.slice().sort(function(a, b) { return a.left - b.left; });
            var hg = [];
            for (var i = 0; i < hs.length - 1; i++) hg.push(hs[i + 1].left - (hs[i].left + hs[i].width));

            var vs = orig.slice().sort(function(a, b) { return a.top - b.top; });
            var vg = [];
            for (var i = 0; i < vs.length - 1; i++) vg.push(vs[i + 1].top - (vs[i].top + vs[i].height));

            orig.forEach(function(o) {
                var nw = g ? o.width + dp : o.width - dp;
                o.nw = nw;
                o.nh = nw * o.ratio;
                o.shape.width = nw;
                o.shape.height = o.nh;
            });

            for (var i = 1; i < hs.length; i++) hs[i].shape.left = hs[i - 1].shape.left + hs[i - 1].nw + hg[i - 1];
            for (var i = 1; i < vs.length; i++) vs[i].shape.top = vs[i - 1].shape.top + vs[i - 1].nh + vg[i - 1];

            return ctx.sync().then(function() {
                showStatus("Prop " + (g ? "+" : "-") + " Abstaende OK", "success");
            });
        });
    });
}

/* ══════════════════════════════════════════════
   SNAP TO GRID
   ══════════════════════════════════════════════ */
function snap(mode) {
    withShapes(1, function(ctx, items) {
        items.forEach(function(s) { s.load(["left", "top", "width", "height"]); });
        return ctx.sync().then(function() {
            items.forEach(function(s) {
                if (mode === "position" || mode === "both") {
                    s.left = c2p(rnd(p2c(s.left)));
                    s.top = c2p(rnd(p2c(s.top)));
                }
                if (mode === "size" || mode === "both") {
                    var nw = rnd(p2c(s.width));
                    var nh = rnd(p2c(s.height));
                    if (nw >= MIN) s.width = c2p(nw);
                    if (nh >= MIN) s.height = c2p(nh);
                }
            });
            return ctx.sync().then(function() {
                showStatus("Eingerastet (" + mode + ")", "success");
            });
        });
    });
}

/* ══════════════════════════════════════════════
   SPACING – Multi-Row / Multi-Column
   Groups shapes by Y (rows) or X (columns),
   then applies fixed spacing within each group.
   ══════════════════════════════════════════════ */

/* Group shapes into rows or columns by position similarity */
function groupByPosition(items, axis, tolerance) {
    var groups = [];
    var used = {};
    var sorted = items.slice().sort(function(a, b) {
        return axis === "y" ? (a.top - b.top) : (a.left - b.left);
    });
    for (var i = 0; i < sorted.length; i++) {
        if (used[i]) continue;
        var grp = [sorted[i]];
        used[i] = true;
        var refPos = axis === "y" ? sorted[i].top : sorted[i].left;
        for (var j = i + 1; j < sorted.length; j++) {
            if (used[j]) continue;
            var pos = axis === "y" ? sorted[j].top : sorted[j].left;
            if (Math.abs(pos - refPos) <= tolerance) {
                grp.push(sorted[j]);
                used[j] = true;
            }
        }
        groups.push(grp);
    }
    return groups;
}

function spacing(dir) {
    withShapes(2, function(ctx, items) {
        items.forEach(function(s) { s.load(["left", "top", "width", "height"]); });
        return ctx.sync().then(function() {
            var sp = c2p(gridUnitCm);

            /* Tolerance: half a grid unit in points, minimum 5pt.
               Shapes within this Y/X range are considered the same row/column */
            var tol = c2p(gridUnitCm) * 0.5;
            if (tol < 5) tol = 5;

            if (dir === "horizontal") {
                /* Group by Y position (= rows), then space each row horizontally */
                var rows = groupByPosition(items, "y", tol);
                rows.forEach(function(row) {
                    if (row.length < 2) return;
                    row.sort(function(a, b) { return a.left - b.left; });
                    for (var i = 1; i < row.length; i++) {
                        row[i].left = row[i - 1].left + row[i - 1].width + sp;
                    }
                });
                var rc = rows.length;
                return ctx.sync().then(function() {
                    showStatus("H-Abstand " + gridUnitCm.toFixed(2) + " cm (" + rc + " Zeile" + (rc > 1 ? "n" : "") + ")", "success");
                });
            } else {
                /* Group by X position (= columns), then space each column vertically */
                var cols = groupByPosition(items, "x", tol);
                cols.forEach(function(col) {
                    if (col.length < 2) return;
                    col.sort(function(a, b) { return a.top - b.top; });
                    for (var i = 1; i < col.length; i++) {
                        col[i].top = col[i - 1].top + col[i - 1].height + sp;
                    }
                });
                var cc = cols.length;
                return ctx.sync().then(function() {
                    showStatus("V-Abstand " + gridUnitCm.toFixed(2) + " cm (" + cc + " Spalte" + (cc > 1 ? "n" : "") + ")", "success");
                });
            }
        });
    });
}

/* ══════════════════════════════════════════════
   SHAPE INFO
   ══════════════════════════════════════════════ */
function shapeInfo() {
    withShapes(1, function(ctx, items) {
        items.forEach(function(s) { s.load(["name", "left", "top", "width", "height"]); });
        return ctx.sync().then(function() {
            var el = document.getElementById("infoDisplay");
            var html = "";
            items.forEach(function(s, idx) {
                if (items.length > 1) {
                    html += '<div style="font-weight:700;margin-top:' + (idx > 0 ? '4' : '0') + 'px;color:#e94560">' + (s.name || 'Obj ' + (idx + 1)) + '</div>';
                }
                html += '<div class="info-item"><span class="info-label">W:</span><span class="info-value">' + p2c(s.width).toFixed(2) + ' cm</span></div>';
                html += '<div class="info-item"><span class="info-label">H:</span><span class="info-value">' + p2c(s.height).toFixed(2) + ' cm</span></div>';
                html += '<div class="info-item"><span class="info-label">X:</span><span class="info-value">' + p2c(s.left).toFixed(2) + ' cm</span></div>';
                html += '<div class="info-item"><span class="info-label">Y:</span><span class="info-value">' + p2c(s.top).toFixed(2) + ' cm</span></div>';
            });
            el.innerHTML = html;
            el.classList.add("visible");
            showStatus("Info geladen", "info");
        });
    });
}

/* ══════════════════════════════════════════════
   MATCH (Klick = Max, Shift = Min)
   ══════════════════════════════════════════════ */
function match(dim, mode) {
    withShapes(2, function(ctx, items) {
        items.forEach(function(s) { s.load(["width", "height"]); });
        return ctx.sync().then(function() {
            var ws = items.map(function(s) { return s.width; });
            var hs = items.map(function(s) { return s.height; });
            var tw = mode === "max" ? Math.max.apply(null, ws) : Math.min.apply(null, ws);
            var th = mode === "max" ? Math.max.apply(null, hs) : Math.min.apply(null, hs);
            items.forEach(function(s) {
                if (dim === "width"  || dim === "both") s.width  = tw;
                if (dim === "height" || dim === "both") s.height = th;
            });
            return ctx.sync().then(function() {
                showStatus("Angeglichen -> " + mode, "success");
            });
        });
    });
}

function propMatch(mode) {
    withShapes(2, function(ctx, items) {
        items.forEach(function(s) { s.load(["width", "height"]); });
        return ctx.sync().then(function() {
            var ws = items.map(function(s) { return s.width; });
            var tw = mode === "max" ? Math.max.apply(null, ws) : Math.min.apply(null, ws);
            items.forEach(function(s) {
                var r = s.height / s.width;
                s.width = tw;
                s.height = tw * r;
            });
            return ctx.sync().then(function() {
                showStatus("Prop angeglichen -> " + mode, "success");
            });
        });
    });
}

/* ══════════════════════════════════════════════
   GRID TABLE
   ══════════════════════════════════════════════ */
function createGridTable() {
    var cols = parseInt(document.getElementById("tCols").value);
    var rows = parseInt(document.getElementById("tRows").value);
    var cw   = parseFloat(document.getElementById("tCW").value);
    var ch   = parseFloat(document.getElementById("tCH").value);

    if (isNaN(cols) || isNaN(rows) || cols < 1 || rows < 1) {
        showStatus("Ungueltige Spalten/Zeilen!", "error"); return;
    }
    if (isNaN(cw) || isNaN(ch) || cw < 1 || ch < 1) {
        showStatus("Ungueltige Zellgroesse!", "error"); return;
    }
    if (cols > 15) { showStatus("Max 15 Spalten!", "warning"); return; }
    if (rows > 20) { showStatus("Max 20 Zeilen!", "warning"); return; }

    PowerPoint.run(function(ctx) {
        var sel = ctx.presentation.getSelectedSlides();
        sel.load("items");
        return ctx.sync().then(function() {
            if (sel.items.length > 0) return buildTable(ctx, sel.items[0], cols, rows, cw, ch);
            var slides = ctx.presentation.slides;
            slides.load("items");
            return ctx.sync().then(function() {
                if (!slides.items.length) { showStatus("Keine Folie!", "error"); return ctx.sync(); }
                return buildTable(ctx, slides.items[0], cols, rows, cw, ch);
            });
        });
    }).catch(function(e) {
        showStatus("Fehler: " + e.message, "error");
    });
}

function buildTable(ctx, slide, cols, rows, cw, ch) {
    var wCm = cw * gridUnitCm;
    var hCm = ch * gridUnitCm;
    var sp   = gridUnitCm;
    var wPt  = c2p(wCm);
    var hPt  = c2p(hCm);
    var spPt = c2p(sp);
    var x0   = c2p(8 * gridUnitCm);
    var y0   = c2p(17 * gridUnitCm);

    for (var r = 0; r < rows; r++) {
        for (var c = 0; c < cols; c++) {
            var s = slide.shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle);
            s.left   = x0 + (c * (wPt + spPt));
            s.top    = y0 + (r * (hPt + spPt));
            s.width  = wPt;
            s.height = hPt;
            s.fill.setSolidColor("FFFFFF");
            s.lineFormat.color  = "808080";
            s.lineFormat.weight = 0.3;
            s.name = "TC_" + r + "_" + c;
        }
    }
    return ctx.sync().then(function() {
        showStatus(cols + "x" + rows + " Tabelle erstellt", "success");
    });
}

/* ══════════════════════════════════════════════
   SETUP: Slide Size
   ══════════════════════════════════════════════ */
function setSlideSize() {
    var tw = 785.5;
    var th = 547;
    PowerPoint.run(function(ctx) {
        var ps = ctx.presentation.pageSetup;
        ps.load(["slideWidth", "slideHeight"]);
        return ctx.sync().then(function() {
            ps.slideWidth = tw;
            return ctx.sync();
        }).then(function() {
            ps.slideHeight = th;
            return ctx.sync();
        }).then(function() {
            showStatus("Format: 27,711 x 19,297 cm", "success");
        });
    }).catch(function(e) {
        showStatus("Fehler: " + e.message, "error");
    });
}

/* ══════════════════════════════════════════════
   SETUP: Guidelines (Toggle)
   ══════════════════════════════════════════════ */
function toggleGuides() {
    PowerPoint.run(function(ctx) {
        var masters = ctx.presentation.slideMasters;
        masters.load("items");
        return ctx.sync().then(function() {
            if (!masters.items.length) {
                showStatus("Kein Master!", "error");
                return ctx.sync();
            }
            var m0 = masters.items[0];
            var sh = m0.shapes;
            sh.load("items");
            return ctx.sync().then(function() {
                for (var i = 0; i < sh.items.length; i++) sh.items[i].load("name");
                return ctx.sync().then(function() {
                    var existing = [];
                    for (var i = 0; i < sh.items.length; i++) {
                        if (sh.items[i].name && sh.items[i].name.indexOf(GTAG) === 0) existing.push(sh.items[i]);
                    }
                    return existing.length > 0
                        ? removeGuides(ctx, masters.items)
                        : addGuides(ctx, masters.items);
                });
            });
        });
    }).catch(function(e) {
        showStatus("Fehler: " + e.message, "error");
    });
}

function addGuides(ctx, masters) {
    var pos = [
        { t: "vertical",   u: 8   },
        { t: "vertical",   u: 126 },
        { t: "horizontal", u: 5   },
        { t: "horizontal", u: 9   },
        { t: "horizontal", u: 15  },
        { t: "horizontal", u: 17  },
        { t: "horizontal", u: 86  }
    ];
    var ps = ctx.presentation.pageSetup;
    ps.load(["slideWidth", "slideHeight"]);
    return ctx.sync().then(function() {
        var sw = ps.slideWidth;
        var sh = ps.slideHeight;
        masters.forEach(function(master) {
            pos.forEach(function(g) {
                var pt = Math.round(c2p(g.u * gridUnitCm));
                var s;
                if (g.t === "vertical") {
                    s = master.shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle);
                    s.left = pt; s.top = 0; s.width = 1; s.height = sh;
                } else {
                    s = master.shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle);
                    s.left = 0; s.top = pt; s.width = sw; s.height = 1;
                }
                s.name = GTAG + "_" + g.t + "_" + g.u;
                s.fill.setSolidColor("FF0000");
                s.lineFormat.visible = false;
            });
        });
        return ctx.sync().then(function() {
            showStatus("Hilfslinien eingeblendet", "success");
        });
    });
}

function removeGuides(ctx, masters) {
    var proms = [];
    masters.forEach(function(master) {
        var sh = master.shapes;
        sh.load("items");
        proms.push(ctx.sync().then(function() {
            for (var i = 0; i < sh.items.length; i++) sh.items[i].load("name");
            return ctx.sync().then(function() {
                for (var i = 0; i < sh.items.length; i++) {
                    if (sh.items[i].name && sh.items[i].name.indexOf(GTAG) === 0) sh.items[i].delete();
                }
            });
        }));
    });
    return Promise.all(proms).then(function() {
        return ctx.sync().then(function() {
            showStatus("Hilfslinien entfernt", "success");
        });
    });
}

/* ══════════════════════════════════════════════
   SETUP: Copy Shadow Values
   ══════════════════════════════════════════════ */
function copyShadow() {
    var t = "Schatten:\nFarbe: Schwarz\nTransparenz: 75%\nGroesse: 100%\nWeichzeichnen: 4pt\nWinkel: 90\nAbstand: 1pt";
    if (navigator.clipboard && navigator.clipboard.writeText) {
        navigator.clipboard.writeText(t).then(function() {
            showStatus("Kopiert", "success");
        }).catch(function() {
            showStatus("Kopieren fehlgeschlagen", "error");
        });
    } else {
        showStatus("Zwischenablage nicht verfuegbar", "error");
    }
}
