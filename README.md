# DROEGE Grid Resize Tool v15 26/02/2026 - 21:00
## PowerPoint Office Add-in – Komplette Dokumentation

---

## 1. Übersicht

Das **Grid Resize Tool** ist ein PowerPoint Office Add-in für das DROEGE GROUP Corporate Design. Es ermöglicht pixelgenaues Arbeiten auf einem 0,21 cm (6 pt) Raster und bietet Funktionen für Größenanpassung, Positionierung, Abstandssteuerung und Hilfslinien.

### Dateien im ZIP

| Datei | Beschreibung |
|---|---|
| `taskpane.js` | Kernlogik – alle Funktionen |
| `taskpane.html` | Benutzeroberfläche (Taskpane) |
| `taskpane.css` | Styling / Corporate Design |
| `manifest-grid-resize.xml` | Office Add-in Manifest |
| `README.md` | Diese Dokumentation |

---

## 2. Installation

### Voraussetzungen
- Microsoft PowerPoint (Desktop oder Online)
- PowerPointApi 1.10 oder höher

### Schritte
1. ZIP-Datei entpacken
2. Alle Dateien auf einen Webserver oder localhost ablegen
3. `manifest-grid-resize.xml` in PowerPoint laden:
   - **Windows:** Datei → Optionen → Trust Center → Vertrauenswürdige Add-in-Kataloge
   - **Mac:** Einfügen → Add-ins → Meine Add-ins → Benutzerdefiniertes Add-in hochladen
   - **Online:** Einfügen → Office Add-ins → Mein Add-in hochladen
4. Add-in erscheint im Menüband

---

## 3. Grundkonstanten

| Konstante | Wert | Beschreibung |
|---|---|---|
| `CM` | 28.3465 | Umrechnungsfaktor cm → Points |
| `gridUnitCm` | 0.2117 cm | 1 Rastereinheit (RE) = 6 pt |
| `MIN` | 0.1 cm | Minimale Objektgröße |
| `GTAG` | `DROEGE_GUIDELINE` | Namens-Prefix für Hilfslinien-Shapes |

---

## 4. Unterstützte Papierformate

| Format | Breite (pt) | Höhe (pt) | Offset X (cm) | Offset Y (cm) |
|---|---|---|---|---|
| **16:9** | 720.0 | 405.0 | 0.10 | 0.00 |
| **4:3** | 720.0 | 540.0 | 0.10 | 0.069 |
| **16:10** | 720.0 | 450.0 | 0.10 | 0.17 |
| **A4 quer** | 780.0 | 540.0 | 0.11 | 0.07 |
| **Breitbild** | 960.0 | 540.0 | 0.13 | 0.07 |

Die Formaterkennung (`getGridOffsets`) nutzt einen Nearest-Neighbor-Vergleich mit einer Toleranz von 10 pt.

---

## 5. Benutzeroberfläche – Tabs

Das Tool hat **3 Tabs**:

### Tab 1: Größe
Objekte auf dem Raster vergrößern/verkleinern.

| Button | Funktion | Shift-Klick |
|---|---|---|
| **W +** | Breite + 1 RE (0,21 cm) | Breite − 1 RE |
| **H +** | Höhe + 1 RE | Höhe − 1 RE |
| **↔ +** | Proportional breiter | Proportional schmaler |

**Multi-Row/Multi-Column:** Bei mehreren markierten Objekten werden diese automatisch nach Position gruppiert. Objekte in derselben Zeile/Spalte werden gemeinsam verändert.

### Tab 2: Position
Objekte am Raster ausrichten und Abstände setzen.

| Button | Funktion |
|---|---|
| **W Max** / Shift: **W Min** | Breiten angleichen (max/min) |
| **H Max** / Shift: **H Min** | Höhen angleichen (max/min) |
| **⬒ Max** / Shift: **⬒ Min** | Proportional angleichen |
| **Snap Position** | Position auf Raster snappen |
| **Snap Größe** | Größe auf Raster snappen |
| **Snap Alles** | Position + Größe snappen |
| **ℹ Info** | Shape-Details anzeigen |

#### Abstände (Spacing)
| Button | Funktion |
|---|---|
| **Horizontal** | Horizontalen Abstand auf volles RE runden |
| **Vertikal** | Vertikalen Abstand auf volles RE runden |

**Spacing-Logik (v3 Rewrite):**
- Erkennt automatisch Zeilen/Spalten
- Multi-Row: Jede Zeile wird separat behandelt
- Multi-Column: Jede Spalte wird separat behandelt
- Abstand wird auf das nächste volle RE gerundet

#### Grid-Tabelle
| Eingabe | Beschreibung |
|---|---|
| **Spalten** | Anzahl Spalten (Standard: 4) |
| **Zeilen** | Anzahl Zeilen (Standard: 2) |
| **Zelle B** | Zellenbreite in RE (Standard: 20) |
| **Zelle H** | Zellenhöhe in RE (Standard: 10) |
| **→ Tabelle erstellen** | Erzeugt eine Grid-Tabelle aus Rechtecken |

### Tab 3: Extras
Zusätzliche Werkzeuge.

#### Papierformat
- **🔍 Format erkennen:** Liest `slideWidth`/`slideHeight` aus und zeigt erkanntes Format, Maße in cm und pt in der Statuszeile an

#### Hilfslinien (Master)
- **Ein-/Ausblenden:** Toggle – erstellt oder löscht Hilfslinien im Folienmaster
- **Senkrechte Linien (dynamisch, formatabhängig):**
  - Links: `Offset_X + 7 RE` = Position der linken Linie
  - Rechts: `Folienbreite − Offset_X − 6 RE` = Position der rechten Linie
  - Der Offset wird aus `getGridOffsets()` ermittelt, d.h. je nach erkanntem Format (16:9, 4:3, 16:10, A4 quer, Breitbild) werden die Positionen korrekt berechnet
- **Waagerechte Linien (fest):** RE 5, 9, 15, 17, 86
- Farbe: Rot (#FF0000), Stärke: 1 pt
- Statusmeldung zeigt erkanntes Format und berechnete RE-Positionen an

#### VBA Grid-Raster
- **VBA: Raster 6 pt kopieren:** Kopiert ein Mac-kompatibles VBA-Macro in die Zwischenablage
- **Custom:** Beliebigen pt-Wert eingeben und "Kopieren" klicken
- Das Macro setzt `ActivePresentation.GridDistance` und aktiviert `SnapToGrid`
- Anwendung: In PowerPoint Alt+F11 → VBA-Editor → Macro einfügen und ausführen

#### Schatten-Werte
Zeigt die DROEGE Corporate Schatten-Einstellungen an:

| Parameter | Wert |
|---|---|
| Farbe | Schwarz |
| Transparenz | 75 % |
| Größe | 100 % |
| Weichzeichnen | 4 pt |
| Winkel | 90° |
| Abstand | 1 pt |

- **Werte kopieren:** Kopiert alle Werte als Text in die Zwischenablage

---

## 6. Funktionsreferenz (taskpane.js)

### Hilfsfunktionen

| Funktion | Beschreibung |
|---|---|
| `c2p(cm)` | cm → Points |
| `p2c(pt)` | Points → cm |
| `rnd(v)` | Rundet auf nächste RE |
| `getTol()` | Liefert Toleranzwert (½ RE in pt) |
| `hlPre(val)` | Formatiert cm-Wert mit RE-Angabe für Anzeige |
| `showStatus(msg, type)` | Statusmeldung anzeigen (success/error/warning) |
| `bind(id, fn)` | Button-Klick binden |
| `shiftBind(id, fnNormal, fnShift)` | Normaler Klick + Shift-Klick binden |
| `withShapes(min, cb)` | Shapes laden und Callback ausführen |

### Format-Erkennung

| Funktion | Beschreibung |
|---|---|
| `getGridOffsets(slideW, slideH)` | Erkennt Format anhand von slideWidth/slideHeight, gibt `{x, y, name}` zurück |
| `detectFormat()` | Zeigt erkanntes Format in Statuszeile an |

### Kernfunktionen

| Funktion | Beschreibung |
|---|---|
| `resize(dim, deltaCm)` | Größe ändern (Multi-Row/Multi-Column) |
| `propResize(deltaCm)` | Proportionale Größenänderung |
| `snap(mode)` | Auf Raster snappen ("position", "size", "both") |
| `spacing(dir)` | Abstände ausgleichen ("horizontal", "vertical") |
| `matchDim(dim, mode)` | Dimensionen angleichen ("max", "min") |
| `propMatch(mode)` | Proportional angleichen |
| `shapeInfo()` | Shape-Details anzeigen |

### Gruppierung

| Funktion | Beschreibung |
|---|---|
| `groupByPos(items, axis, tol)` | Gruppiert Shapes nach Position (Zeilen/Spalten) |
| `groupByData(data, prop, tol)` | Gruppiert Datenpunkte nach Eigenschaft |

### Extras

| Funktion | Beschreibung |
|---|---|
| `createGridTable()` | Grid-Tabelle aus Rechtecken erstellen |
| `buildTbl(ctx, slide, cols, rows, cwRE, chRE)` | Tabellen-Builder (intern) |
| `toggleGuides()` | Hilfslinien ein-/ausblenden |
| `addGuides(ctx, masters)` | Hilfslinien erstellen (dynamisch) |
| `rmGuides(ctx, masters)` | Hilfslinien entfernen |
| `copyShadowText()` | Schatten-Werte in Zwischenablage kopieren |
| `detectFormat()` | Papierformat erkennen und anzeigen |
| `copyVBA(pts)` | VBA-Macro für Raster generieren und kopieren |

---

## 7. Snap-Logik (Detail)

Der Snap berechnet den Raster-Offset direkt aus der Foliengeometrie:

```
Rastereinheit (gPt) = gridUnitCm × 28.3465 = 6 pt
Offset X = (slideWidth  % gPt) / 2
Offset Y = (slideHeight % gPt) / 2
```

**Position-Snap:**
```
shape.left = offsetX + round((left − offsetX) / gPt) × gPt
shape.top  = offsetY + round((top  − offsetY) / gPt) × gPt
```

**Size-Snap:**
```
shape.width  = round(width  / gPt) × gPt  (min: 0.1 cm)
shape.height = round(height / gPt) × gPt  (min: 0.1 cm)
```

---

## 8. Hilfslinien-Berechnung (Detail)

### Vertikale Linien (dynamisch)
```
off = getGridOffsets(slideWidth, slideHeight)
offXcm = off.name ≠ "Unbekannt" ? p2c(off.x) : 0

Links:  offXcm + 7 × 0.2117 cm  →  Offset + 7 RE
Rechts: Folienbreite_cm − offXcm − 6 × 0.2117 cm  →  Breite − Offset − 6 RE
```

### Beispiel 16:9 (720 × 405 pt)
```
offXcm = 0.10 cm
Links:  0.10 + 1.4819 = 1.5819 cm ≈ 7.5 RE  →  gerundet 8 RE
Rechts: 25.4 − 0.10 − 1.2702 = 24.03 cm ≈ 113.5 RE  →  gerundet 126 RE
```

### Horizontale Linien (fest)
| RE | cm | Verwendung |
|---|---|---|
| 5 | 1.059 | Obere Begrenzung |
| 9 | 1.905 | Titel-Unterkante |
| 15 | 3.176 | Untertitel-Unterkante |
| 17 | 3.599 | Content-Oberkante |
| 86 | 18.206 | Content-Unterkante |

---

## 9. VBA-Macro (Detail)

Das generierte VBA-Macro setzt das PowerPoint-Raster:

```vba
Sub SetGrid_6pt()
    ' Setzt PowerPoint-Raster auf exakt 6 pt (0.2117 cm)
    ' Mac- und Windows-kompatibel
    On Error Resume Next
    With ActivePresentation
        .GridDistance = 6
        .SnapToGrid = msoTrue
    End With
    Application.DisplayGridLines = msoTrue
    If Err.Number <> 0 Then
        Err.Clear
        MsgBox "GridDistance konnte nicht gesetzt werden.", vbExclamation
        Exit Sub
    End If
    MsgBox "Raster gesetzt auf 6 pt (0.2117 cm)", vbInformation
End Sub
```

**Anwendung:**
1. Button "VBA: Raster 6 pt kopieren" klicken
2. In PowerPoint: Alt+F11 (Windows) bzw. Extras → Makro → Visual Basic-Editor (Mac)
3. Neues Modul einfügen → Macro einfügen
4. F5 zum Ausführen

---

## 10. Versionshistorie

| Version | Änderungen |
|---|---|
| v1–v2 | Grundfunktionen: Resize, Snap |
| v3 | Spacing Rewrite (Multi-Row/Multi-Column) |
| v4 | Grid-Tabelle, Match Dimensions |
| v5 | Proportional Resize/Match |
| v6 | Extras-Tab, Hilfslinien, Schatten |
| v7 | Format-Erkennung, Robuste Snap-Offsets |
| v8 | Dynamische Hilfslinien (formatabhängig) |
| v9 | VBA Grid-Raster, detectFormat, GRID_OFFSETS-Tabelle |
| v10 | Bugfixes, UI-Verbesserungen |
| v11 | Konsolidierung, Code-Cleanup |
| **v12** | **Zusammenführung aller Features: VBA (v9) + dynamische Hilfslinien (v8/v9) + Papierformat-Erkennung (v9) + GRID_OFFSETS-Tabelle** |

---

## 11. Bekannte Einschränkungen

- **PowerPointApi 1.10** muss verfügbar sein (nicht alle Office-Versionen unterstützen dies)
- **VBA-Macro:** Das Macro wird in die Zwischenablage kopiert – der Benutzer muss es manuell im VBA-Editor einfügen
- **Hilfslinien:** Werden als Shapes im Folienmaster erstellt (keine echten PowerPoint-Guides, da die API diese nicht unterstützt)
- **Formaterkennung:** Toleranz ±10 pt – bei stark abweichenden Custom-Formaten wird "Unbekannt" zurückgegeben

---

*DROEGE GROUP – Grid Resize Tool v12 – Erstellt: Februar 2026*
