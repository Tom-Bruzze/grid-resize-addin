/* ═══════════════════════════════════════════════════════════════
   DROEGE Grid Resize Tool – taskpane.js (Komplett)
   ═══════════════════════════════════════════════════════════════
   Features:
     • Resize (W / H / W+H / Proportional) – Multi-Row/Col
     • Snap to Grid (Position / Size / Both)
     • Spacing (H / V) – Multi-Row/Col
     • Match Dimensions (W / H / W+H / Proportional)
     • Grid-Tabelle erstellen
     • Papierformat 27,728 × 19,297 cm
     • Hilfslinien im Master (Toggle)
     • Schatten-Werte kopieren
   ═══════════════════════════════════════════════════════════════ */

/* ── Konstanten ── */
var CM = 28.3465;                   /* 1 cm = 28.3465 pt         */
var MIN = 0.1;                      /* Mindestgröße in cm        */
var gridUnitCm = 0.21;              /* Default-Rastereinheit     */
var apiOk = false;                  /* PowerPointApi 1.5 da?     */
var GTAG = "DROEGE_GUIDELINE";      /* Name-Prefix Hilfslinien   */

/* ══════════════════════════════════════════════════════════════
   OFFICE INIT
   ══════════════════════════════════════════════════════════════ */
Office.onReady(function (info) {
    if (info.host === Office.HostType.PowerPoint) {

        /* Prüfe ob PowerPointApi 1.5 verfügbar ist */
        if (Office.context.requirements && Office.context.requirements.isSetSupported) {
            apiOk = Office.context.requirements.isSetSupported("PowerPointApi", "1.5");
        } else {
            apiOk = (typeof PowerPoint !== "undefined" &&
                     PowerPoint.run &&
                     typeof PowerPoint.run === "function");
        }

        initUI();

        if (!apiOk) {
            showStatus("PowerPointApi 1.5 nicht verfügbar – nur Desktop/Web", "warning");
        }
    }
});

/* ══════════════════════════════════════════════════════════════
   UI INIT – Alle Event-Listener
   ══════════════════════════════════════════════════════════════ */
function initUI() {

    /* ── Rastereinheit: Custom Input ── */
    var gi = document.getElementById("gridUnit");
    gi.addEventListener("change", function () {
        var v = parseFloat(this.value);
        if (!isNaN(v) && v > 0) {
            gridUnitCm = v;
            highlightPresets(v);
            showStatus("Rastereinheit: " + v.toFixed(2) + " cm", "info");
        }
    });

    /* ── Rastereinheit: Preset-Buttons ── */
    document.querySelectorAll(".pre").forEach(function (btn) {
        btn.addEventListener("click", function () {
            var v = parseFloat(this.dataset.value);
            gridUnitCm = v;
            gi.value = v;
            highlightPresets(v);
            showStatus("Rastereinheit: " + v.toFixed(2) + " cm", "info");
        });
    });

    /* ── Tab-Navigation ── */
    document.querySelectorAll(".tab").forEach(function (btn) {
        btn.addEventListener("click", function () {
            var id = this.dataset.tab;
            document.querySelectorAll(".tab").forEach(function (t) { t.classList.remove("active"); });
            document.querySelectorAll(".pane").forEach(function (p) { p.classList.remove("active"); });
            this.classList.add("active");
            document.getElementById(id).classList.add("active");
        });
    });

    /* ─────────────────────────────────────────────
       TAB 1 – GRÖSSE
       Klick = vergrößern (+1 RE)
       Shift+Klick = verkleinern (−1 RE)
       ───────────────────────────────────────────── */
    shiftBind("resizeW",    function () { resize("width",  gridUnitCm);  },
                            function () { resize("width", -gridUnitCm);  });

    shiftBind("resizeH",    function () { resize("height",  gridUnitCm); },
                            function () { resize("height", -gridUnitCm); });

    shiftBind("resizeBoth", function () { resize("both",  gridUnitCm);   },
                            function () { resize("both", -gridUnitCm);   });

    shiftBind("resizeProp", function () { propResize( gridUnitCm); },
                            function () { propResize(-gridUnitCm); });

    /* ─────────────────────────────────────────────
       TAB 2 – RASTER
       ───────────────────────────────────────────── */
    bind("snapPos",     function () { snap("position"); });
    bind("snapSize",    function () { snap("size");     });
    bind("snapAll",     function () { snap("both");     });
    bind("spaceH",      function () { spacing("horizontal"); });
    bind("spaceV",      function () { spacing("vertical");   });
    bind("showInfo",    function () { shapeInfo();  });
    bind("createTable", function () { createGridTable(); });

    /* ─────────────────────────────────────────────
       TAB 3 – ANGLEICHEN
       Klick = auf größtes Objekt
       Shift+Klick = auf kleinstes Objekt
       ───────────────────────────────────────────── */
    shiftBind("matchW",    function () { matchDim("width",  "max"); },
                           function () { matchDim("width",  "min"); });

    shiftBind("matchH",    function () { matchDim("height", "max"); },
                           function () { matchDim("height", "min"); });

    shiftBind("matchBoth", function () { matchDim("both",   "max"); },
                           function () { matchDim("both",   "min"); });

    shiftBind("matchProp", function () { propMatch("max"); },
                           function () { propMatch("min"); });

    /* ─────────────────────────────────────────────
       TAB 4 – EXTRAS
       ───────────────────────────────────────────── */
    bind("setSlide",     function () { setSlideSize();   });
    bind("toggleGuides", function () { toggleGuides();   });
    bind("copyShadow",   function () { copyShadowText(); });
}

/* ══════════════════════════════════════════════════════════════
   HILFS-FUNKTIONEN
   ══════════════════════════════════════════════════════════════ */

/** Klick = fnNormal, Shift+Klick = fnShift */
function shiftBind(id, fnNormal, fnShift) {
    var el = document.getElementById(id);
    if (!el) return;
    el.addEventListener("click", function (e) {
        e.shiftKey ? fnShift() : fnNormal();
    });
}

/** Einfacher Klick-Binder */
function bind(id, fn) {
    var el = document.getElementById(id);
    if (!el) return;
    el.addEventListener("click", fn);
}

/** Preset-Buttons highlighten */
function highlightPresets(val) {
    document.querySelectorAll(".pre").forEach(function (btn) {
        btn.classList.toggle("active",
            Math.abs(parseFloat(btn.dataset.value) - val) < 0.001);
    });
}

/** Status anzeigen (permanent sichtbar) */
function showStatus(msg, type) {
    var el = document.getElementById("status");
    el.textContent = msg;
    el.className = "sts " + (type || "info");
}

/** Umrechnung cm ↔ pt */
function c2p(cm) { return cm * CM; }
function p2c(pt) { return pt / CM; }

/** Auf nächstes Vielfaches der Rastereinheit runden */
function rnd(valueCm) {
    return Math.round(valueCm / gridUnitCm) * gridUnitCm;
}

/** Toleranz für Zeilen/Spalten-Gruppierung */
function getTolerance() {
    var tol = c2p(gridUnitCm) * 0.5;
    return tol < 5 ? 5 : tol;
}

/* ── withShapes: Kern-Wrapper mit API-Check ── */
function withShapes(minCount, callback) {
    if (!apiOk) {
        showStatus("Nicht unterstützt (PowerPointApi 1.5 nötig)", "error");
        return;
    }
    PowerPoint.run(function (ctx) {
        var shapes = ctx.presentation.getSelectedShapes();
        shapes.load("items");
        return ctx.sync().then(function () {
            if (shapes.items.length < minCount) {
                showStatus(
                    minCount <= 1
                        ? "Bitte Objekt(e) auswählen!"
                        : "Mindestens " + minCount + " Objekte auswählen!",
                    "error"
                );
                return;
            }
            return callback(ctx, shapes.items);
        });
    }).catch(function (err) {
        showStatus("Fehler: " + err.message, "error");
    });
}

/* ══════════════════════════════════════════════════════════════
   GRUPPIERUNG: Multi-Row / Multi-Column
   ══════════════════════════════════════════════════════════════
   Shapes werden nach Y-Position (→ Zeilen) oder
   X-Position (→ Spalten) gruppiert.
   Shapes mit ähnlicher Position (±Toleranz) landen
   in der gleichen Gruppe.
   ══════════════════════════════════════════════════════════════ */
function groupByPosition(items, axis, tolerance) {
    var groups = [];
    var used = {};

    /* Sortiere primär nach der gewählten Achse */
    var sorted = items.slice().sort(function (a, b) {
        return axis === "y" ? (a.top - b.top) : (a.left - b.left);
    });

    for (var i = 0; i < sorted.length; i++) {
        if (used[i]) continue;

        var group = [sorted[i]];
        used[i] = true;

        var refPos = axis === "y" ? sorted[i].top : sorted[i].left;

        for (var j = i + 1; j < sorted.length; j++) {
            if (used[j]) continue;
            var pos = axis === "y" ? sorted[j].top : sorted[j].left;
            if (Math.abs(pos - refPos) <= tolerance) {
                group.push(sorted[j]);
                used[j] = true;
            }
        }
        groups.push(group);
    }
    return groups;
}

/* ══════════════════════════════════════════════════════════════
   RESIZE – Multi-Row / Multi-Column
   ══════════════════════════════════════════════════════════════
   Einzelnes Shape: einfaches Resize
   Mehrere Shapes:
     • Width-Resize  → Gruppiert nach Y (Zeilen),
                        jede Zeile wird horizontal resized,
                        Gaps innerhalb der Zeile bleiben erhalten
     • Height-Resize → Gruppiert nach X (Spalten),
                        jede Spalte wird vertikal resized,
                        Gaps innerhalb der Spalte bleiben erhalten
   ══════════════════════════════════════════════════════════════ */
function resize(dim, deltaCm) {
    withShapes(1, function (ctx, items) {
        items.forEach(function (s) { s.load(["left", "top", "width", "height"]); });
        return ctx.sync().then(function () {
            var dp   = c2p(Math.abs(deltaCm));
            var grow = deltaCm > 0;

            /* ── Einzelnes Shape ── */
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
                    var lbl = dim === "both" ? "W+H" : dim === "width" ? "Breite" : "Höhe";
                    showStatus(lbl + (grow ? " +" : " −") + Math.abs(deltaCm).toFixed(2) + " cm ✓", "success");
                });
            }

            /* ── Mehrere Shapes: Multi-Row/Col ── */
            var tol = getTolerance();
            var rowInfo = "";
            var colInfo = "";

            /* Width-Resize: Zeilen (nach Y gruppiert) */
            if (dim === "width" || dim === "both") {
                var rows = groupByPosition(items, "y", tol);
                rowInfo = rows.length + " Zeile" + (rows.length > 1 ? "n" : "");

                rows.forEach(function (row) {
                    /* Innerhalb der Zeile nach X sortieren */
                    row.sort(function (a, b) { return a.left - b.left; });

                    /* Gaps zwischen Shapes merken */
                    var gaps = [];
                    for (var i = 0; i < row.length - 1; i++) {
                        gaps.push(row[i + 1].left - (row[i].left + row[i].width));
                    }

                    /* Prüfe ob alle Shapes groß genug bleiben */
                    var ok = true;
                    for (var i = 0; i < row.length; i++) {
                        var newW = grow ? row[i].width + dp : row[i].width - dp;
                        if (newW < c2p(MIN)) { ok = false; break; }
                    }

                    if (ok) {
                        /* Breite anpassen */
                        for (var i = 0; i < row.length; i++) {
                            row[i].width = grow ? row[i].width + dp : row[i].width - dp;
                        }
                        /* Positionen nachziehen → Gaps erhalten */
                        for (var i = 1; i < row.length; i++) {
                            row[i].left = row[i - 1].left + row[i - 1].width + gaps[i - 1];
                        }
                    }
                });
            }

            /* Height-Resize: Spalten (nach X gruppiert) */
            if (dim === "height" || dim === "both") {
                var cols = groupByPosition(items, "x", tol);
                colInfo = cols.length + " Spalte" + (cols.length > 1 ? "n" : "");

                cols.forEach(function (col) {
                    /* Innerhalb der Spalte nach Y sortieren */
                    col.sort(function (a, b) { return a.top - b.top; });

                    /* Gaps zwischen Shapes merken */
                    var gaps = [];
                    for (var i = 0; i < col.length - 1; i++) {
                        gaps.push(col[i + 1].top - (col[i].top + col[i].height));
                    }

                    /* Prüfe ob alle Shapes groß genug bleiben */
                    var ok = true;
                    for (var i = 0; i < col.length; i++) {
                        var newH = grow ? col[i].height + dp : col[i].height - dp;
                        if (newH < c2p(MIN)) { ok = false; break; }
                    }

                    if (ok) {
                        /* Höhe anpassen */
                        for (var i = 0; i < col.length; i++) {
                            col[i].height = grow ? col[i].height + dp : col[i].height - dp;
                        }
                        /* Positionen nachziehen → Gaps erhalten */
                        for (var i = 1; i < col.length; i++) {
                            col[i].top = col[i - 1].top + col[i - 1].height + gaps[i - 1];
                        }
                    }
                });
            }

            return ctx.sync().then(function () {
                var lbl = dim === "both" ? "W+H" : dim === "width" ? "Breite" : "Höhe";
                var detail = "";
                if (rowInfo) detail += " · " + rowInfo;
                if (colInfo) detail += " · " + colInfo;
                showStatus(lbl + (grow ? " +" : " −") + Math.abs(deltaCm).toFixed(2) + " cm" + detail + " ✓", "success");
            });
        });
    });
}

/* ══════════════════════════════════════════════════════════════
   PROPORTIONAL RESIZE – Multi-Row / Multi-Column
   ══════════════════════════════════════════════════════════════
   Breite wird um 1 RE verändert, Höhe passt sich
   proportional (Seitenverhältnis) an.
   Bei mehreren Shapes: Gaps bleiben erhalten.
   ══════════════════════════════════════════════════════════════ */
function propResize(deltaCm) {
    withShapes(1, function (ctx, items) {
        items.forEach(function (s) { s.load(["left", "top", "width", "height"]); });
        return ctx.sync().then(function () {
            var dp   = c2p(Math.abs(deltaCm));
            var grow = deltaCm > 0;

            /* ── Einzelnes Shape ── */
            if (items.length === 1) {
                var s = items[0];
                var ratio = s.height / s.width;
                var nw = grow ? s.width + dp : s.width - dp;
                if (nw >= c2p(MIN)) {
                    var nh = nw * ratio;
                    if (nh >= c2p(MIN)) {
                        s.width  = nw;
                        s.height = nh;
                    }
                }
                return ctx.sync().then(function () {
                    showStatus("Proportional " + (grow ? "+" : "−") + Math.abs(deltaCm).toFixed(2) + " cm ✓", "success");
                });
            }

            /* ── Mehrere Shapes ── */
            var tol = getTolerance();

            /* Daten-Objekte mit Ratio und Originalpositionen */
            var data = items.map(function (s) {
                return {
                    shape: s,
                    ratio: s.height / s.width,
                    left:  s.left,
                    top:   s.top,
                    width: s.width,
                    height: s.height,
                    nw: 0,
                    nh: 0
                };
            });

            /* Prüfe Mindestgröße für alle */
            var ok = true;
            data.forEach(function (o) {
                var nw = grow ? o.width + dp : o.width - dp;
                if (nw < c2p(MIN) || nw * o.ratio < c2p(MIN)) ok = false;
            });
            if (!ok) {
                showStatus("Mindestgröße erreicht!", "error");
                return ctx.sync();
            }

            /* Neue Größen berechnen + setzen */
            data.forEach(function (o) {
                o.nw = grow ? o.width + dp : o.width - dp;
                o.nh = o.nw * o.ratio;
                o.shape.width  = o.nw;
                o.shape.height = o.nh;
            });

            /* Horizontal: Zeilen gruppieren, Gaps erhalten */
            var rows = groupByPosition(data, "y", tol);
            rows.forEach(function (row) {
                row.sort(function (a, b) { return a.left - b.left; });
                var gaps = [];
                for (var i = 0; i < row.length - 1; i++) {
                    gaps.push(row[i + 1].left - (row[i].left + row[i].width));
                }
                for (var i = 1; i < row.length; i++) {
                    row[i].shape.left = row[i - 1].shape.left + row[i - 1].nw + gaps[i - 1];
                }
            });

            /* Vertikal: Spalten gruppieren, Gaps erhalten */
            var cols = groupByPosition(data, "x", tol);
            cols.forEach(function (col) {
                col.sort(function (a, b) { return a.top - b.top; });
                var gaps = [];
                for (var i = 0; i < col.length - 1; i++) {
                    gaps.push(col[i + 1].top - (col[i].top + col[i].height));
                }
                for (var i = 1; i < col.length; i++) {
                    col[i].shape.top = col[i - 1].shape.top + col[i - 1].nh + gaps[i - 1];
                }
            });

            return ctx.sync().then(function () {
                showStatus("Proportional " + (grow ? "+" : "−") +
                    " · " + rows.length + " Zeilen · " + cols.length + " Spalten ✓", "success");
            });
        });
    });
}

/* ══════════════════════════════════════════════════════════════
   SNAP TO GRID
   ══════════════════════════════════════════════════════════════ */
function snap(mode) {
    withShapes(1, function (ctx, items) {
        items.forEach(function (s) { s.load(["left", "top", "width", "height"]); });
        return ctx.sync().then(function () {
            items.forEach(function (s) {
                if (mode === "position" || mode === "both") {
                    s.left = c2p(rnd(p2c(s.left)));
                    s.top  = c2p(rnd(p2c(s.top)));
                }
                if (mode === "size" || mode === "both") {
                    var nw = rnd(p2c(s.width));
                    var nh = rnd(p2c(s.height));
                    if (nw >= MIN) s.width  = c2p(nw);
                    if (nh >= MIN) s.height = c2p(nh);
                }
            });
            return ctx.sync().then(function () {
                var lbl = mode === "both" ? "Position + Größe"
                        : mode === "position" ? "Position"
                        : "Größe";
                showStatus(lbl + " am Raster ausgerichtet ✓", "success");
            });
        });
    });
}

/* ══════════════════════════════════════════════════════════════
   SPACING – Multi-Row / Multi-Column
   ══════════════════════════════════════════════════════════════
   Setzt den Abstand zwischen Objekten auf genau 1 RE.
   Bei mehreren Zeilen/Spalten wird jede Gruppe separat behandelt.
   ══════════════════════════════════════════════════════════════ */
function spacing(direction) {
    withShapes(2, function (ctx, items) {
        items.forEach(function (s) { s.load(["left", "top", "width", "height"]); });
        return ctx.sync().then(function () {
            var sp  = c2p(gridUnitCm);
            var tol = getTolerance();

            if (direction === "horizontal") {
                /* Gruppiere nach Y → Zeilen */
                var rows = groupByPosition(items, "y", tol);
                rows.forEach(function (row) {
                    if (row.length < 2) return;
                    row.sort(function (a, b) { return a.left - b.left; });
                    for (var i = 1; i < row.length; i++) {
                        row[i].left = row[i - 1].left + row[i - 1].width + sp;
                    }
                });
                return ctx.sync().then(function () {
                    showStatus("H-Abstand " + gridUnitCm.toFixed(2) + " cm · " +
                        rows.length + " Zeile" + (rows.length > 1 ? "n" : "") + " ✓", "success");
                });

            } else {
                /* Gruppiere nach X → Spalten */
                var cols = groupByPosition(items, "x", tol);
                cols.forEach(function (col) {
                    if (col.length < 2) return;
                    col.sort(function (a, b) { return a.top - b.top; });
                    for (var i = 1; i < col.length; i++) {
                        col[i].top = col[i - 1].top + col[i - 1].height + sp;
                    }
                });
                return ctx.sync().then(function () {
                    showStatus("V-Abstand " + gridUnitCm.toFixed(2) + " cm · " +
                        cols.length + " Spalte" + (cols.length > 1 ? "n" : "") + " ✓", "success");
                });
            }
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
            var el = document.getElementById("infoBox");
            var html = "";

            items.forEach(function (s, idx) {
                if (items.length > 1) {
                    html += '<div style="font-weight:700;margin-top:' +
                        (idx > 0 ? '6' : '0') + 'px;color:#e94560;font-size:11px;">' +
                        (s.name || "Objekt " + (idx + 1)) + '</div>';
                }
                html += '<div class="info-item">' +
                    '<span class="info-label">Breite:</span>' +
                    '<span class="info-value">' + p2c(s.width).toFixed(2) + ' cm  (' +
                    (p2c(s.width) / gridUnitCm).toFixed(1) + ' RE)</span></div>';
                html += '<div class="info-item">' +
                    '<span class="info-label">Höhe:</span>' +
                    '<span class="info-value">' + p2c(s.height).toFixed(2) + ' cm  (' +
                    (p2c(s.height) / gridUnitCm).toFixed(1) + ' RE)</span></div>';
                html += '<div class="info-item">' +
                    '<span class="info-label">Links:</span>' +
                    '<span class="info-value">' + p2c(s.left).toFixed(2) + ' cm</span></div>';
                html += '<div class="info-item">' +
                    '<span class="info-label">Oben:</span>' +
                    '<span class="info-value">' + p2c(s.top).toFixed(2) + ' cm</span></div>';
            });

            el.innerHTML = html;
            el.classList.add("visible");
            showStatus("Objektinfo geladen ✓", "info");
        });
    });
}

/* ══════════════════════════════════════════════════════════════
   MATCH DIMENSIONS  (Klick = Max, Shift = Min)
   ══════════════════════════════════════════════════════════════ */
function matchDim(dimension, mode) {
    withShapes(2, function (ctx, items) {
        items.forEach(function (s) { s.load(["width", "height"]); });
        return ctx.sync().then(function () {
            var ws = items.map(function (s) { return s.width;  });
            var hs = items.map(function (s) { return s.height; });

            var tw = mode === "max" ? Math.max.apply(null, ws) : Math.min.apply(null, ws);
            var th = mode === "max" ? Math.max.apply(null, hs) : Math.min.apply(null, hs);

            items.forEach(function (s) {
                if (dimension === "width"  || dimension === "both") s.width  = tw;
                if (dimension === "height" || dimension === "both") s.height = th;
            });

            return ctx.sync().then(function () {
                var lbl = dimension === "both" ? "W+H" : dimension === "width" ? "Breite" : "Höhe";
                showStatus(lbl + " → " + (mode === "max" ? "größtes" : "kleinstes") + " Objekt ✓", "success");
            });
        });
    });
}

/* ══════════════════════════════════════════════════════════════
   PROPORTIONAL MATCH  (Klick = Max, Shift = Min)
   ══════════════════════════════════════════════════════════════ */
function propMatch(mode) {
    withShapes(2, function (ctx, items) {
        items.forEach(function (s) { s.load(["width", "height"]); });
        return ctx.sync().then(function () {
            var ws = items.map(function (s) { return s.width; });
            var tw = mode === "max" ? Math.max.apply(null, ws) : Math.min.apply(null, ws);

            items.forEach(function (s) {
                var ratio = s.height / s.width;
                s.width  = tw;
                s.height = tw * ratio;
            });

            return ctx.sync().then(function () {
                showStatus("Proportional → " + (mode === "max" ? "größtes" : "kleinstes") + " ✓", "success");
            });
        });
    });
}

/* ══════════════════════════════════════════════════════════════
   GRID-TABELLE ERSTELLEN
   ══════════════════════════════════════════════════════════════ */
function createGridTable() {
    var cols = parseInt(document.getElementById("tCols").value);
    var rows = parseInt(document.getElementById("tRows").value);
    var cw   = parseFloat(document.getElementById("tCW").value);
    var ch   = parseFloat(document.getElementById("tCH").value);

    if (isNaN(cols) || isNaN(rows) || cols < 1 || rows < 1) {
        showStatus("Ungültige Spalten/Zeilen-Angabe!", "error");
        return;
    }
    if (isNaN(cw) || isNaN(ch) || cw < 1 || ch < 1) {
        showStatus("Ungültige Zellengröße!", "error");
        return;
    }
    if (cols > 15) { showStatus("Max. 15 Spalten!", "warning"); return; }
    if (rows > 20) { showStatus("Max. 20 Zeilen!", "warning"); return; }

    PowerPoint.run(function (ctx) {
        var sel = ctx.presentation.getSelectedSlides();
        sel.load("items");
        return ctx.sync().then(function () {
            if (sel.items.length > 0) {
                return buildTable(ctx, sel.items[0], cols, rows, cw, ch);
            }
            /* Fallback: erste Folie */
            var slides = ctx.presentation.slides;
            slides.load("items");
            return ctx.sync().then(function () {
                if (!slides.items.length) {
                    showStatus("Keine Folie vorhanden!", "error");
                    return ctx.sync();
                }
                return buildTable(ctx, slides.items[0], cols, rows, cw, ch);
            });
        });
    }).catch(function (err) {
        showStatus("Fehler: " + err.message, "error");
    });
}

function buildTable(ctx, slide, cols, rows, cellW_RE, cellH_RE) {
    var wPt  = c2p(cellW_RE * gridUnitCm);   /* Zellenbreite in pt  */
    var hPt  = c2p(cellH_RE * gridUnitCm);   /* Zellenhöhe in pt    */
    var spPt = c2p(gridUnitCm);               /* Spacing = 1 RE     */
    var x0   = c2p(8  * gridUnitCm);          /* Startposition X    */
    var y0   = c2p(17 * gridUnitCm);          /* Startposition Y    */

    for (var r = 0; r < rows; r++) {
        for (var c = 0; c < cols; c++) {
            var s = slide.shapes.addGeometricShape(
                PowerPoint.GeometricShapeType.rectangle
            );
            s.left   = x0 + c * (wPt + spPt);
            s.top    = y0 + r * (hPt + spPt);
            s.width  = wPt;
            s.height = hPt;
            s.fill.setSolidColor("FFFFFF");
            s.lineFormat.color  = "808080";
            s.lineFormat.weight = 0.3;
            s.name = "GridCell_R" + r + "_C" + c;
        }
    }

    return ctx.sync().then(function () {
        showStatus(cols + " × " + rows + " Tabelle erstellt ✓", "success");
    });
}

/* ══════════════════════════════════════════════════════════════
   PAPIERFORMAT – 27,728 × 19,297 cm
   ══════════════════════════════════════════════════════════════
   27.728 cm × 28.3465 pt/cm = 785.98 pt ≈ 786 pt
   19.297 cm × 28.3465 pt/cm = 547.00 pt
   ══════════════════════════════════════════════════════════════ */
function setSlideSize() {
    var targetW = 786;   /* 27,728 cm */
    var targetH = 547;   /* 19,297 cm */

    PowerPoint.run(function (ctx) {
        var ps = ctx.presentation.pageSetup;
        ps.load(["slideWidth", "slideHeight"]);

        return ctx.sync()
            .then(function () {
                ps.slideWidth = targetW;
                return ctx.sync();
            })
            .then(function () {
                ps.slideHeight = targetH;
                return ctx.sync();
            })
            .then(function () {
                showStatus("Format: 27,728 × 19,297 cm ✓", "success");
            });
    }).catch(function (err) {
        showStatus("Fehler: " + err.message, "error");
    });
}

/* ══════════════════════════════════════════════════════════════
   HILFSLINIEN IM MASTER (Toggle)
   ══════════════════════════════════════════════════════════════
   Erzeugt dünne rote Linien im Slide-Master als
   visuelle Orientierung. Erneuter Klick entfernt sie.
   ══════════════════════════════════════════════════════════════ */
function toggleGuides() {
    PowerPoint.run(function (ctx) {
        var masters = ctx.presentation.slideMasters;
        masters.load("items");

        return ctx.sync().then(function () {
            if (!masters.items.length) {
                showStatus("Kein Slide-Master vorhanden!", "error");
                return ctx.sync();
            }

            var m0 = masters.items[0];
            var sh = m0.shapes;
            sh.load("items");

            return ctx.sync().then(function () {
                /* Names laden um vorhandene zu erkennen */
                for (var i = 0; i < sh.items.length; i++) {
                    sh.items[i].load("name");
                }
                return ctx.sync().then(function () {
                    var existing = [];
                    for (var i = 0; i < sh.items.length; i++) {
                        if (sh.items[i].name && sh.items[i].name.indexOf(GTAG) === 0) {
                            existing.push(sh.items[i]);
                        }
                    }
                    if (existing.length > 0) {
                        return removeGuides(ctx, masters.items);
                    } else {
                        return addGuides(ctx, masters.items);
                    }
                });
            });
        });
    }).catch(function (err) {
        showStatus("Fehler: " + err.message, "error");
    });
}

function addGuides(ctx, masters) {
    /* Positionen der Hilfslinien in Rastereinheiten */
    var lines = [
        { type: "vertical",   pos:   8 },
        { type: "vertical",   pos: 126 },
        { type: "horizontal", pos:   5 },
        { type: "horizontal", pos:   9 },
        { type: "horizontal", pos:  15 },
        { type: "horizontal", pos:  17 },
        { type: "horizontal", pos:  86 }
    ];

    var ps = ctx.presentation.pageSetup;
    ps.load(["slideWidth", "slideHeight"]);

    return ctx.sync().then(function () {
        var sw = ps.slideWidth;
        var sh = ps.slideHeight;

        masters.forEach(function (master) {
            lines.forEach(function (line) {
                var pt = Math.round(c2p(line.pos * gridUnitCm));
                var s;

                if (line.type === "vertical") {
                    s = master.shapes.addGeometricShape(
                        PowerPoint.GeometricShapeType.rectangle
                    );
                    s.left   = pt;
                    s.top    = 0;
                    s.width  = 1;
                    s.height = sh;
                } else {
                    s = master.shapes.addGeometricShape(
                        PowerPoint.GeometricShapeType.rectangle
                    );
                    s.left   = 0;
                    s.top    = pt;
                    s.width  = sw;
                    s.height = 1;
                }

                s.name = GTAG + "_" + line.type + "_" + line.pos;
                s.fill.setSolidColor("FF0000");
                s.lineFormat.visible = false;
            });
        });

        return ctx.sync().then(function () {
            showStatus("Hilfslinien eingeblendet ✓", "success");
        });
    });
}

function removeGuides(ctx, masters) {
    var promises = [];

    masters.forEach(function (master) {
        var sh = master.shapes;
        sh.load("items");

        promises.push(
            ctx.sync().then(function () {
                for (var i = 0; i < sh.items.length; i++) {
                    sh.items[i].load("name");
                }
                return ctx.sync().then(function () {
                    for (var i = 0; i < sh.items.length; i++) {
                        if (sh.items[i].name && sh.items[i].name.indexOf(GTAG) === 0) {
                            sh.items[i].delete();
                        }
                    }
                });
            })
        );
    });

    return Promise.all(promises).then(function () {
        return ctx.sync().then(function () {
            showStatus("Hilfslinien entfernt ✓", "success");
        });
    });
}

/* ══════════════════════════════════════════════════════════════
   SCHATTEN-WERTE KOPIEREN
   ══════════════════════════════════════════════════════════════ */
function copyShadowText() {
    var text =
        "Schatten-Standardwerte:\n" +
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
}
