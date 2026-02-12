/* 
 ═══════════════════════════════════════════════════════
 DROEGE Grid Resize Tool  –  taskpane.js
 
 Tabs: Tools (Größe + Angleich + Raster + Abstände)
       Extras (Papier + Hilfslinien + GANTT + Schatten)
 ═══════════════════════════════════════════════════════
 */

var CM = 28.3465;
var MIN = 0.1;
var gridUnitCm = 0.21;
var apiOk = false;
var GTAG = "DROEGE_GUIDELINE";

/* ═══ GANTT Constants ═══ */
var GANTT_MAX_W = 118;   /* max Breite in RE              */
var GANTT_MAX_H = 69;    /* max Höhe in RE                */
var GANTT_LEFT  = 8;     /* Abstand links in RE           */
var GANTT_TOP   = 17;    /* Abstand oben  in RE           */
var GANTT_TAG   = "DROEGE_GANTT";

/* ═══════════════════════════════════════════
   Office Init
   ═══════════════════════════════════════════ */
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

/* ═══════════════════════════════════════════
   UI INIT
   ═══════════════════════════════════════════ */
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

  /* Größe (Klick = +, Shift = −) */
  shiftBind("resizeW", function () { resize("width", gridUnitCm); },
    function () { resize("width", -gridUnitCm); });
  shiftBind("resizeH", function () { resize("height", gridUnitCm); },
    function () { resize("height", -gridUnitCm); });
  shiftBind("resizeBoth", function () { resize("both", gridUnitCm); },
    function () { resize("both", -gridUnitCm); });
  shiftBind("resizeProp", function () { propResize( gridUnitCm); },
    function () { propResize(-gridUnitCm); });

  /* Angleichen (Klick = Max, Shift = Min) */
  shiftBind("matchW", function () { matchDim("width", "max"); },
    function () { matchDim("width", "min"); });
  shiftBind("matchH", function () { matchDim("height", "max"); },
    function () { matchDim("height", "min"); });
  shiftBind("matchBoth", function () { matchDim("both", "max"); },
    function () { matchDim("both", "min"); });
  shiftBind("matchProp", function () { propMatch("max"); },
    function () { propMatch("min"); });

  /* Raster */
  bind("snapPos", function () { snap("position"); });
  bind("snapSize", function () { snap("size"); });
  bind("snapAll", function () { snap("both"); });
  bind("showInfo", function () { shapeInfo(); });

  /* Abstände */
  bind("spaceH", function () { spacing("horizontal"); });
  bind("spaceV", function () { spacing("vertical"); });

  /* Grid-Tabelle */
  bind("createTable", function () { createGridTable(); });

  /* Extras */
  bind("setSlide", function () { setSlideSize(); });
  bind("toggleGuides", function () { toggleGuides(); });
  bind("copyShadow", function () { copyShadowText(); });

  /* ═══ GANTT Init ═══ */
  initGantt();
}

/* ═══════════════════════════════════════════
   HILFSFUNKTIONEN
   ═══════════════════════════════════════════ */
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
function rnd(v) { return Math.round(v / gridUnitCm) * gridUnitCm; }

function getTol() {
  var t = c2p(gridUnitCm) * 0.5;
  return t < 5 ? 5 : t;
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

/* ═══════════════════════════════════════════
   GRUPPIERUNG: Multi-Row / Multi-Column
   ═══════════════════════════════════════════ */
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

/* ═══════════════════════════════════════════
   RESIZE Multi-Row / Multi-Column
   ═══════════════════════════════════════════ */
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
        var d = ri && ci ? " " + ri + " " + ci : ri ? " " + ri : ci ? " " + ci : "";
        showStatus(l + (grow ? " +" : " −") + Math.abs(deltaCm).toFixed(2) + " cm" + d + " ✓", "success");
      });
    });
  });
}

/* ═══════════════════════════════════════════
   PROPORTIONAL RESIZE
   ═══════════════════════════════════════════ */
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
        showStatus("Prop " + (grow ? "+" : "−") + " " + rows.length + " Zl " + cols.length + " Sp ✓", "success");
      });
    });
  });
}

/* ═══════════════════════════════════════════
   SNAP TO GRID
   ═══════════════════════════════════════════ */
function snap(mode) {
  withShapes(1, function (ctx, items) {
    items.forEach(function (s) { s.load(["left", "top", "width", "height"]); });
    return ctx.sync().then(function () {
      items.forEach(function (s) {
        if (mode === "position" || mode === "both") {
          s.left = c2p(rnd(p2c(s.left)));
          s.top = c2p(rnd(p2c(s.top)));
        }
        if (mode === "size" || mode === "both") {
          var nw = rnd(p2c(s.width)), nh = rnd(p2c(s.height));
          if (nw >= MIN) s.width = c2p(nw);
          if (nh >= MIN) s.height = c2p(nh);
        }
      });
      return ctx.sync().then(function () {
        var l = mode === "both" ? "Pos+Size" : mode === "position" ? "Position" : "Größe";
        showStatus(l + " → Raster ✓", "success");
      });
    });
  });
}

/* ═══════════════════════════════════════════
   SPACING Multi-Row / Multi-Column
   ═══════════════════════════════════════════ */
function spacing(dir) {
  withShapes(2, function (ctx, items) {
    items.forEach(function (s) { s.load(["left", "top", "width", "height"]); });
    return ctx.sync().then(function () {
      var sp = c2p(gridUnitCm), tol = getTol();

      if (dir === "horizontal") {
        var rows = groupByPos(items, "y", tol);
        rows.forEach(function (row) {
          if (row.length < 2) return;
          row.sort(function (a, b) { return a.left - b.left; });
          for (var i = 1; i < row.length; i++)
            row[i].left = row[i - 1].left + row[i - 1].width + sp;
        });
        return ctx.sync().then(function () {
          showStatus("H-Abstand " + gridUnitCm.toFixed(2) + " cm → " + rows.length + " Zl ✓", "success");
        });
      } else {
        var cols = groupByPos(items, "x", tol);
        cols.forEach(function (col) {
          if (col.length < 2) return;
          col.sort(function (a, b) { return a.top - b.top; });
          for (var i = 1; i < col.length; i++)
            col[i].top = col[i - 1].top + col[i - 1].height + sp;
        });
        return ctx.sync().then(function () {
          showStatus("V-Abstand " + gridUnitCm.toFixed(2) + " cm → " + cols.length + " Sp ✓", "success");
        });
      }
    });
  });
}

/* ═══════════════════════════════════════════
   SHAPE INFO
   ═══════════════════════════════════════════ */
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

/* ═══════════════════════════════════════════
   MATCH DIMENSIONS
   ═══════════════════════════════════════════ */
function matchDim(dim, mode) {
  withShapes(2, function (ctx, items) {
    items.forEach(function (s) { s.load(["width", "height"]); });
    return ctx.sync().then(function () {
      var ws = items.map(function (s) { return s.width; });
      var hs = items.map(function (s) { return s.height; });
      var tw = mode === "max" ? Math.max.apply(null, ws) : Math.min.apply(null, ws);
      var th = mode === "max" ? Math.max.apply(null, hs) : Math.min.apply(null, hs);
      items.forEach(function (s) {
        if (dim === "width" || dim === "both") s.width = tw;
        if (dim === "height" || dim === "both") s.height = th;
      });
      return ctx.sync().then(function () {
        var l = dim === "both" ? "W+H" : dim === "width" ? "W" : "H";
        showStatus(l + " → " + (mode === "max" ? "Max" : "Min") + " ✓", "success");
      });
    });
  });
}

/* ═══════════════════════════════════════════
   PROPORTIONAL MATCH
   ═══════════════════════════════════════════ */
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

/* ═══════════════════════════════════════════
   GRID TABLE
   ═══════════════════════════════════════════ */
function createGridTable() {
  var cols = parseInt(document.getElementById("tCols").value);
  var rows = parseInt(document.getElementById("tRows").value);
  var cw = parseFloat(document.getElementById("tCW").value);
  var ch = parseFloat(document.getElementById("tCH").value);

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
    showStatus(cols + "×" + rows + " Tabelle ✓", "success");
  });
}

/* ═══════════════════════════════════════════
   PAPIERFORMAT 27,728 × 19,297 cm
   ═══════════════════════════════════════════ */
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

/* ═══════════════════════════════════════════
   HILFSLINIEN (Master Toggle)
   ═══════════════════════════════════════════ */
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
          return ex.length > 0
            ? rmGuides(ctx, masters.items)
            : addGuides(ctx, masters.items);
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
        else { s.left = 0; s.top = pt; s.width = sw; s.height = 1; }
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

/* ═══════════════════════════════════════════
   SCHATTEN-WERTE KOPIEREN
   ═══════════════════════════════════════════ */
function copyShadowText() {
  var txt = "Schatten: Schwarz, 75% Transparenz, 100% Größe, 4pt Weichzeichnen, 90° Winkel, 1pt Abstand";
  if (navigator.clipboard && navigator.clipboard.writeText) {
    navigator.clipboard.writeText(txt).then(function () {
      showStatus("Schatten-Werte kopiert ✓", "success");
    }).catch(function () {
      showStatus("Kopieren fehlgeschlagen", "error");
    });
  } else {
    showStatus("Clipboard nicht verfügbar", "warning");
  }
}


/* ═══════════════════════════════════════════════════════
   ██████   ██████   ██   █   ████████  ████████
   █        █    █   ██   █      █         █
   █  ███   ██████   █ █  █      █         █
   █    █   █    █   █  █ █      █         █
   ██████   █    █   █   ██      █         █

   GANTT-DIAGRAMM – Erzeugt auf der aktuellen Folie
   
   Fläche:   max 118 RE breit × max 69 RE hoch
   Position: links 8 RE, oben 17 RE vom Rand
   ═══════════════════════════════════════════════════════ */

var ganttPhaseCount = 0;

function initGantt() {

  /* Default-Datum: heute → +3 Monate */
  var today = new Date();
  var d3m = new Date(today);
  d3m.setMonth(d3m.getMonth() + 3);
  document.getElementById("ganttStart").value = isoDate(today);
  document.getElementById("ganttEnd").value = isoDate(d3m);

  /* 3 Beispiel-Phasen */
  addGanttPhase("Konzeption", today, offsetDays(today, 14), "#2e86c1");
  addGanttPhase("Umsetzung", offsetDays(today, 14), offsetDays(today, 56), "#27ae60");
  addGanttPhase("Abnahme", offsetDays(today, 56), d3m, "#e94560");

  /* Buttons */
  bind("ganttAddPhase", function () {
    var s = new Date(document.getElementById("ganttStart").value);
    var e = new Date(document.getElementById("ganttEnd").value);
    if (isNaN(s.getTime()) || isNaN(e.getTime())) { s = today; e = d3m; }
    addGanttPhase("Phase " + (ganttPhaseCount + 1), s, offsetDays(s, 14), randomColor());
  });

  bind("createGantt", function () { createGanttChart(); });
}

/* ─── Phase-UI hinzufügen ─── */
function addGanttPhase(name, start, end, color) {
  ganttPhaseCount++;
  var id = "gp_" + ganttPhaseCount;
  var div = document.createElement("div");
  div.className = "gantt-phase";
  div.id = id;
  div.innerHTML =
    '<input type="text" value="' + name + '" placeholder="Name" title="Phasenname">' +
    '<input type="date" value="' + isoDate(start) + '" title="Start">' +
    '<input type="date" value="' + isoDate(end) + '" title="Ende">' +
    '<input type="color" value="' + color + '" title="Farbe">' +
    '<button class="gantt-del" title="Entfernen">&times;</button>';
  document.getElementById("ganttPhases").appendChild(div);
  div.querySelector(".gantt-del").addEventListener("click", function () {
    div.remove();
  });
}

/* ─── Hilfsfunktionen ─── */
function isoDate(d) {
  var mm = ("0" + (d.getMonth() + 1)).slice(-2);
  var dd = ("0" + d.getDate()).slice(-2);
  return d.getFullYear() + "-" + mm + "-" + dd;
}

function offsetDays(d, n) {
  var r = new Date(d);
  r.setDate(r.getDate() + n);
  return r;
}

function randomColor() {
  var colors = ["#2e86c1","#27ae60","#e94560","#f39c12","#8e44ad","#1abc9c","#e67e22","#3498db","#d35400","#16a085"];
  return colors[Math.floor(Math.random() * colors.length)];
}

function daysBetween(a, b) {
  return Math.round((b - a) / (1000 * 60 * 60 * 24));
}

function weeksBetween(a, b) {
  return Math.ceil(daysBetween(a, b) / 7);
}

function monthsBetween(a, b) {
  return (b.getFullYear() - a.getFullYear()) * 12 + (b.getMonth() - a.getMonth()) + (b.getDate() > a.getDate() ? 1 : 0);
}

function quartersBetween(a, b) {
  return Math.ceil(monthsBetween(a, b) / 3);
}

function ganttInfo(msg, err) {
  var el = document.getElementById("ganttInfo");
  el.innerHTML = msg;
  el.className = "gantt-info" + (err ? " err" : "");
}

/* ─── Phasen aus UI lesen ─── */
function readPhases() {
  var phases = [];
  var items = document.querySelectorAll(".gantt-phase");
  items.forEach(function (div) {
    var inputs = div.querySelectorAll("input");
    var name  = inputs[0].value || "Phase";
    var start = new Date(inputs[1].value);
    var end   = new Date(inputs[2].value);
    var color = inputs[3].value || "#2e86c1";
    if (!isNaN(start.getTime()) && !isNaN(end.getTime()) && end > start) {
      phases.push({ name: name, start: start, end: end, color: color });
    }
  });
  return phases;
}

/* ═══════════════════════════════════════════
   GANTT ERZEUGEN – Hauptfunktion
   ═══════════════════════════════════════════ */
function createGanttChart() {
  if (!apiOk) { showStatus("PowerPointApi 1.5 nötig", "error"); return; }

  /* Eingaben lesen */
  var projStart = new Date(document.getElementById("ganttStart").value);
  var projEnd   = new Date(document.getElementById("ganttEnd").value);
  var unit      = document.getElementById("ganttUnit").value;
  var labelWRE  = parseInt(document.getElementById("ganttLabelW").value) || 20;
  var showHead  = document.getElementById("ganttHeader").checked;
  var showToday = document.getElementById("ganttToday").checked;
  var barColor  = document.getElementById("ganttBarColor").value;
  var headColor = document.getElementById("ganttHeadColor").value;
  var rowColor  = document.getElementById("ganttRowColor").value;

  if (isNaN(projStart.getTime()) || isNaN(projEnd.getTime())) {
    ganttInfo("❌ Ungültige Datumsangaben!", true); return;
  }
  if (projEnd <= projStart) {
    ganttInfo("❌ Ende muss nach Start liegen!", true); return;
  }

  var phases = readPhases();
  if (phases.length === 0) {
    ganttInfo("❌ Mindestens eine Phase hinzufügen!", true); return;
  }

  /* Zeiteinheiten berechnen */
  var numUnits;
  if (unit === "week")    numUnits = weeksBetween(projStart, projEnd);
  if (unit === "month")   numUnits = monthsBetween(projStart, projEnd);
  if (unit === "quarter") numUnits = quartersBetween(projStart, projEnd);
  if (numUnits < 1) numUnits = 1;

  /* Layout berechnen */
  var numRows   = phases.length + (showHead ? 1 : 0);
  var chartWRE  = GANTT_MAX_W - labelWRE;            /* Breite für Zeitachse */
  var colWRE    = Math.floor(chartWRE / numUnits);    /* Breite pro Zeiteinheit */
  if (colWRE < 1) colWRE = 1;
  var usedWRE   = colWRE * numUnits;                  /* tatsächlich genutzte Breite */
  var rowHRE    = Math.floor(GANTT_MAX_H / numRows);  /* Höhe pro Zeile */
  if (rowHRE < 2) rowHRE = 2;
  if (rowHRE > 6) rowHRE = 6;                         /* max 6 RE pro Zeile */

  var totalDays = daysBetween(projStart, projEnd);
  if (totalDays < 1) totalDays = 1;

  /* Info anzeigen */
  ganttInfo(
    "<b>" + numUnits + "</b> " + (unit === "week" ? "Wochen" : unit === "month" ? "Monate" : "Quartale") +
    " | <b>" + phases.length + "</b> Phasen | Spalte: <b>" + colWRE + " RE</b> | Zeile: <b>" + rowHRE + " RE</b>"
  );

  /* ─── PowerPoint: Shapes erzeugen ─── */
  PowerPoint.run(function (ctx) {
    var sel = ctx.presentation.getSelectedSlides();
    sel.load("items");
    return ctx.sync().then(function () {
      var slide;
      if (sel.items.length > 0) {
        slide = sel.items[0];
      } else {
        var slides = ctx.presentation.slides;
        slides.load("items");
        return ctx.sync().then(function () {
          if (!slides.items.length) { showStatus("Keine Folie!", "error"); return ctx.sync(); }
          return buildGantt(ctx, slides.items[0], projStart, projEnd, unit, numUnits,
            labelWRE, colWRE, rowHRE, usedWRE, phases, showHead, showToday,
            barColor, headColor, rowColor, totalDays);
        });
      }
      return buildGantt(ctx, slide, projStart, projEnd, unit, numUnits,
        labelWRE, colWRE, rowHRE, usedWRE, phases, showHead, showToday,
        barColor, headColor, rowColor, totalDays);
    });
  }).catch(function (e) { showStatus("Fehler: " + e.message, "error"); });
}

/* ═══════════════════════════════════════════
   GANTT SHAPES BAUEN
   ═══════════════════════════════════════════ */
function buildGantt(ctx, slide, projStart, projEnd, unit, numUnits,
                    labelWRE, colWRE, rowHRE, usedWRE, phases, showHead, showToday,
                    barColor, headColor, rowColor, totalDays) {

  var re   = gridUnitCm;
  var x0   = c2p(GANTT_LEFT * re);   /* 8 RE vom linken Rand  */
  var y0   = c2p(GANTT_TOP  * re);   /* 17 RE vom oberen Rand */
  var lbW  = c2p(labelWRE * re);     /* Label-Spaltenbreite   */
  var cW   = c2p(colWRE * re);       /* Spaltenbreite         */
  var rH   = c2p(rowHRE * re);       /* Zeilenhöhe            */
  var gap  = c2p(re);                /* 1 RE Abstand          */

  var curRow = 0;
  var shapeIdx = 0;

  /* ─── KOPFZEILE (Zeitachse) ─── */
  if (showHead) {
    /* Leere Label-Zelle oben links */
    var hdrLabel = slide.shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle);
    hdrLabel.left   = x0;
    hdrLabel.top    = y0;
    hdrLabel.width  = lbW;
    hdrLabel.height = rH;
    hdrLabel.fill.setSolidColor(headColor.replace("#", ""));
    hdrLabel.lineFormat.color = "FFFFFF";
    hdrLabel.lineFormat.weight = 0.3;
    hdrLabel.name = GANTT_TAG + "_hdr_label";

    /* Zeiteinheit-Zellen */
    for (var u = 0; u < numUnits; u++) {
      var hCell = slide.shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle);
      hCell.left   = x0 + lbW + gap + u * (cW + gap);
      hCell.top    = y0;
      hCell.width  = cW;
      hCell.height = rH;
      hCell.fill.setSolidColor(headColor.replace("#", ""));
      hCell.lineFormat.color = "FFFFFF";
      hCell.lineFormat.weight = 0.3;
      hCell.name = GANTT_TAG + "_hdr_" + u;

      /* Label-Text: Woche/Monat/Quartal */
      var label = getUnitLabel(projStart, u, unit);
      var tf = hCell.textFrame;
      tf.autoSizeSetting = PowerPoint.ShapeAutoSize.none;
      var tr = tf.textRange;
      tr.text = label;
      tr.font.size = 7;
      tr.font.color = "FFFFFF";
      tr.font.bold = true;
      tf.verticalAlignment = PowerPoint.TextVerticalAlignment.middle;
      tr.paragraphFormat.alignment = PowerPoint.ParagraphAlignment.center;
    }
    curRow = 1;
  }

  /* ─── PHASEN-ZEILEN ─── */
  for (var p = 0; p < phases.length; p++) {
    var phase = phases[p];
    var rowY = y0 + curRow * (rH + gap);

    /* Label-Zelle (Phasenname) */
    var lb = slide.shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle);
    lb.left   = x0;
    lb.top    = rowY;
    lb.width  = lbW;
    lb.height = rH;
    lb.fill.setSolidColor(rowColor.replace("#", ""));
    lb.lineFormat.color = "CCCCCC";
    lb.lineFormat.weight = 0.3;
    lb.name = GANTT_TAG + "_label_" + p;

    var lbTf = lb.textFrame;
    lbTf.autoSizeSetting = PowerPoint.ShapeAutoSize.none;
    var lbTr = lbTf.textRange;
    lbTr.text = phase.name;
    lbTr.font.size = 7;
    lbTr.font.color = "333333";
    lbTr.font.bold = true;
    lbTf.verticalAlignment = PowerPoint.TextVerticalAlignment.middle;
    lbTr.paragraphFormat.alignment = PowerPoint.ParagraphAlignment.left;

    /* Hintergrund-Zellen (Zeile) */
    for (var u = 0; u < numUnits; u++) {
      var bgCell = slide.shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle);
      bgCell.left   = x0 + lbW + gap + u * (cW + gap);
      bgCell.top    = rowY;
      bgCell.width  = cW;
      bgCell.height = rH;
      bgCell.fill.setSolidColor(u % 2 === 0 ? rowColor.replace("#", "") : "FFFFFF");
      bgCell.lineFormat.color = "E0E0E0";
      bgCell.lineFormat.weight = 0.2;
      bgCell.name = GANTT_TAG + "_bg_" + p + "_" + u;
    }

    /* ─── Balken (Phase) ─── */
    var chartStartX = x0 + lbW + gap;
    var totalChartW = numUnits * (cW + gap) - gap;

    /* Phase-Start relativ zum Projektstart */
    var pStartDay = daysBetween(projStart, phase.start);
    var pEndDay   = daysBetween(projStart, phase.end);
    if (pStartDay < 0) pStartDay = 0;
    if (pEndDay > totalDays) pEndDay = totalDays;
    if (pEndDay <= pStartDay) { curRow++; continue; }

    var barXStart = chartStartX + (pStartDay / totalDays) * totalChartW;
    var barXEnd   = chartStartX + (pEndDay / totalDays) * totalChartW;
    var barW      = barXEnd - barXStart;
    if (barW < c2p(re)) barW = c2p(re);  /* min 1 RE */

    /* Balken etwas kleiner als Zeile (Padding) */
    var barPad = rH * 0.15;
    var bar = slide.shapes.addGeometricShape(PowerPoint.GeometricShapeType.roundedRectangle);
    bar.left   = barXStart;
    bar.top    = rowY + barPad;
    bar.width  = barW;
    bar.height = rH - barPad * 2;
    bar.fill.setSolidColor(phase.color.replace("#", ""));
    bar.lineFormat.visible = false;
    bar.name = GANTT_TAG + "_bar_" + p;

    curRow++;
  }

  /* ─── HEUTE-LINIE ─── */
  if (showToday) {
    var today = new Date();
    var todayDay = daysBetween(projStart, today);
    if (todayDay >= 0 && todayDay <= totalDays) {
      var chartStartX2 = x0 + lbW + gap;
      var totalChartW2 = numUnits * (cW + gap) - gap;
      var todayX = chartStartX2 + (todayDay / totalDays) * totalChartW2;
      var totalH = curRow * (rH + gap);

      var todayLine = slide.shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle);
      todayLine.left   = todayX;
      todayLine.top    = y0;
      todayLine.width  = c2p(0.05);
      todayLine.height = totalH;
      todayLine.fill.setSolidColor("E94560");
      todayLine.lineFormat.visible = false;
      todayLine.name = GANTT_TAG + "_today";
    }
  }

  return ctx.sync().then(function () {
    showStatus("Gantt: " + phases.length + " Phasen × " + numUnits + " Einheiten ✓", "success");
  });
}

/* ─── Zeiteinheit-Label erzeugen ─── */
function getUnitLabel(start, idx, unit) {
  var d = new Date(start);
  if (unit === "week") {
    d.setDate(d.getDate() + idx * 7);
    var dd = ("0" + d.getDate()).slice(-2);
    var mm = ("0" + (d.getMonth() + 1)).slice(-2);
    return "KW" + getWeekNumber(d);
  }
  if (unit === "month") {
    d.setMonth(d.getMonth() + idx);
    var months = ["Jan","Feb","Mrz","Apr","Mai","Jun","Jul","Aug","Sep","Okt","Nov","Dez"];
    return months[d.getMonth()];
  }
  if (unit === "quarter") {
    d.setMonth(d.getMonth() + idx * 3);
    var q = Math.floor(d.getMonth() / 3) + 1;
    return "Q" + q + "/" + (d.getFullYear() % 100);
  }
  return "" + (idx + 1);
}

function getWeekNumber(d) {
  var oneJan = new Date(d.getFullYear(), 0, 1);
  var days = Math.floor((d - oneJan) / (24 * 60 * 60 * 1000));
  return Math.ceil((days + oneJan.getDay() + 1) / 7);
}
