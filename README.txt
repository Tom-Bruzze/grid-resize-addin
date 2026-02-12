═══════════════════════════════════════════════════
 DROEGE Grid Resize Tool – README
═══════════════════════════════════════════════════

 DATEIEN
────────
 taskpane.html   – HTML-Oberfläche (3 Tabs)
 taskpane.css    – Dark Theme Styling
 taskpane.js     – Komplette Logik

 INSTALLATION
─────────────
 1. Dateien in ein PowerPoint Web Add-in einbinden
 2. Office.js wird via CDN geladen
 3. Erfordert PowerPointApi 1.5 (Desktop/Web)

 BEDIENUNG
──────────
 TAB 1 – RESIZE:
   Klick    = +1 Rastereinheit (vergrößern / Max)
   Shift    = -1 Rastereinheit (verkleinern / Min)

 TAB 2 – GRID:
   Snap     = Position/Größe auf Raster einrasten
   H-Dist   = Horizontale Abstände = 1 RE
              (funktioniert über MEHRERE ZEILEN)
   V-Dist   = Vertikale Abstände = 1 RE
              (funktioniert über MEHRERE SPALTEN)
   Tabelle  = Rechteck-Raster erstellen

 TAB 3 – SETUP:
   Format   = Folienformat 27,711 x 19,297 cm
   Guides   = Hilfslinien im Master ein/aus
   Shadow   = Schatten-Werte kopieren

 MULTI-ROW/COL SPACING
──────────────────────
 Objekte werden anhand ihrer Position automatisch
 in Zeilen (horizontal) bzw. Spalten (vertikal)
 gruppiert. So können z.B. 6 Objekte in einem
 2x3 Grid auf einmal korrekt beabstandet werden.

═══════════════════════════════════════════════════
