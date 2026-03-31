# EXCEL ADVANCED DESIGN REFERENCE
## Inspired by Josh Cottrell-Schloemer's "Excel + Design" methodology
## For use by Claude Code when building Excel dashboards via MCP

---

## PHILOSOPHY: TREAT EXCEL LIKE POWERPOINT

The core insight: Excel has almost identical design tools to PowerPoint — shapes, layers, gradients, transparency, images, shadows, reflections, glow effects, 3D formats. 98% of Excel users never touch these. The difference between a "data dump" and a professional dashboard is using these features.

**The goal:** When someone sees the output, their first reaction should be "Is that really Excel?"

**Key mindset shifts:**
- Stop thinking in cells/grids → think in CARDS and LAYERS
- Stop using default colors → use a curated COLOR PALETTE
- Stop showing raw tables → use SHAPES as containers for data
- Charts are not standalone → they sit ON TOP of styled card backgrounds
- Every element must be ALIGNED to a grid (AAE: Always Align Everything)

---

## THE CARD LAYOUT METHOD

Instead of placing data in cells, break the dashboard into rectangular "cards" — each card tells ONE mini-story.

### Anatomy of a Card
```
┌─────────────────────────────────┐
│  CARD TITLE (small, muted)      │  ← Text shape, 9pt, secondary color
│                                 │
│  $1.2M                          │  ← Big number shape, 28pt bold, accent
│  ▲ 12% vs prior                 │  ← Delta indicator, 10pt, green/red
│                                 │
│  ┌─────────────────────────┐    │
│  │   Mini chart (no axes)  │    │  ← Embedded chart, transparent bg
│  └─────────────────────────┘    │
│                                 │
└─────────────────────────────────┘
    ↑ Rounded rectangle shape with:
      - Solid fill OR subtle gradient
      - No border OR 1pt very light border
      - Optional: subtle shadow (offset 2-4pt, 70-80% transparent)
```

### Card Grid Layout
```
┌──────────┐ ┌──────────┐ ┌──────────┐ ┌──────────┐
│  KPI #1  │ │  KPI #2  │ │  KPI #3  │ │  KPI #4  │
│          │ │          │ │          │ │          │
└──────────┘ └──────────┘ └──────────┘ └──────────┘
     8px gap between cards (consistent everywhere)

┌─────────────────────┐ ┌─────────────────────────┐
│                     │ │                         │
│   Main Chart        │ │   Secondary Chart       │
│   (2/3 width)       │ │   (1/3 width)           │
│                     │ │                         │
└─────────────────────┘ └─────────────────────────┘

┌─────────────────────────────────────────────────┐
│                                                 │
│   Full-Width Detail Table or Timeline           │
│                                                 │
└─────────────────────────────────────────────────┘
```

**Critical rules:**
- ALL cards same height within a row
- ALL gaps identical (8-12px)
- Use Excel's Align/Distribute tools (or VBA `.Align` / `.Distribute`)
- Cards snap to an invisible grid

---

## COLOR PALETTES

### Dark Theme (Most Impressive — "Command Center" Feel)
```
Background:     #1B2A4A (deep navy) or #121212 (true dark)
Card fill:      #1E3A5F (slightly lighter navy) or #1E1E1E
Card border:    #2E4057 (subtle, 1pt) or none
Text primary:   #FFFFFF
Text secondary: #8899AA (muted blue-gray) or #B0B0B0
Accent 1:       #00E5FF (cyan) or #4FC3F7 (soft blue)
Accent 2:       #FF6D00 (orange) or #FFB74D (warm amber)
Success:        #66BB6A (green)
Danger:         #EF5350 (red)
Chart series:   #4FC3F7, #FF6D00, #66BB6A, #AB47BC, #EC407A, #FFA726
```

### Light Theme (Enterprise-Preferred)
```
Background:     #F5F6FA (very light gray-blue)
Card fill:      #FFFFFF (white)
Card border:    #E0E4E8 (light gray, 1pt)
Card shadow:    #000000 at 8% opacity, offset 2pt, blur 6pt
Text primary:   #1B2A4A (dark navy)
Text secondary: #6B7B8D (medium gray)
Accent 1:       #2979FF (strong blue)
Accent 2:       #FF6D00 (orange)
Success:        #2E7D32 (green)
Danger:         #C62828 (red)
Chart series:   #2979FF, #FF6D00, #2E7D32, #7B1FA2, #C62828, #00838F
```

### Teal Analytics
```
Background:     #E0F2F1 or #FAFAFA
Card fill:      #FFFFFF
Accent 1:       #00695C (deep teal)
Accent 2:       #FF6F00 (amber)
Headers:        #004D40 fill + white text
```

**Rule: Pick ONE palette and use it everywhere. Never mix. Never use Excel defaults.**

---

## VBA SHAPE TECHNIQUES

### Foundation: Creating a Card Background
```vba
Sub CreateCard(ws As Worksheet, left As Single, top As Single, _
               width As Single, height As Single, _
               fillColor As Long, Optional borderColor As Long = -1)
    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, _
        left, top, width, height)
    
    With shp
        .Name = "Card_" & Format(Now, "hhmmss") & Int(Rnd * 1000)
        
        ' Fill
        .Fill.Solid
        .Fill.ForeColor.RGB = fillColor
        
        ' Border
        If borderColor = -1 Then
            .Line.Visible = msoFalse
        Else
            .Line.Visible = msoTrue
            .Line.ForeColor.RGB = borderColor
            .Line.Weight = 1
        End If
        
        ' Rounded corners (adjust 0-1, lower = more rounded)
        .Adjustments(1) = 0.05
        
        ' No selection handles when not selected
        .Placement = xlFreeFloating
    End With
End Sub
```

### Gradient Fill (Depth/Shadow Effect)
```vba
Sub ApplyGradient(shp As Shape, color1 As Long, color2 As Long, _
                  Optional direction As Long = msoGradientVertical)
    With shp.Fill
        .TwoColorGradient direction, 1
        .ForeColor.RGB = color1  ' Top/left color
        .BackColor.RGB = color2  ' Bottom/right color
    End With
End Sub

' Example: Subtle dark card with depth
' ApplyGradient shp, RGB(30, 58, 95), RGB(27, 42, 74)
```

### Multi-Stop Gradient (Advanced)
```vba
Sub ApplyMultiGradient(shp As Shape)
    With shp.Fill
        .OneColorGradient msoGradientVertical, 1, 0
        .ForeColor.RGB = RGB(0, 128, 128)
        ' Add stops
        .GradientStops.Insert RGB(0, 100, 120), 0.3
        .GradientStops.Insert RGB(0, 80, 100), 0.7
    End With
End Sub
```

### Shadow (Makes Cards "Float")
```vba
Sub ApplyShadow(shp As Shape)
    With shp.Shadow
        .Visible = msoTrue
        .Type = msoShadow21  ' Outer, offset bottom-right
        .ForeColor.RGB = RGB(0, 0, 0)
        .Transparency = 0.8  ' 80% transparent = very subtle
        .OffsetX = 3
        .OffsetY = 3
        .Blur = 8
    End With
End Sub
```

### Text Shape (Free-Floating Label/Metric)
```vba
Sub CreateTextShape(ws As Worksheet, left As Single, top As Single, _
                    width As Single, height As Single, _
                    txt As String, fontSize As Single, _
                    fontColor As Long, Optional isBold As Boolean = False)
    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(msoShapeRectangle, _
        left, top, width, height)
    
    With shp
        ' Make shape invisible — only text shows
        .Fill.Visible = msoFalse
        .Line.Visible = msoFalse
        
        ' Add and format text
        .TextFrame2.TextRange.Text = txt
        With .TextFrame2.TextRange.Font
            .Size = fontSize
            .Fill.ForeColor.RGB = fontColor
            .Bold = isBold
            .Name = "Aptos"  ' or "Segoe UI" or "Calibri"
        End With
        
        ' Text alignment
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignLeft
        
        ' No margins
        .TextFrame2.MarginLeft = 4
        .TextFrame2.MarginTop = 0
        .TextFrame2.MarginRight = 0
        .TextFrame2.MarginBottom = 0
    End With
End Sub
```

### Dynamic Text (Linked to Cell Value)
```vba
Sub CreateLinkedMetric(ws As Worksheet, left As Single, top As Single, _
                       width As Single, height As Single, _
                       cellRef As String, fontSize As Single, _
                       fontColor As Long)
    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(msoShapeRectangle, _
        left, top, width, height)
    
    With shp
        .Fill.Visible = msoFalse
        .Line.Visible = msoFalse
        ' Link to cell — shape text auto-updates when cell changes
        .DrawingObject.Formula = "=" & cellRef
        
        With .TextFrame2.TextRange.Font
            .Size = fontSize
            .Fill.ForeColor.RGB = fontColor
            .Bold = True
            .Name = "Aptos"
        End With
    End With
End Sub
```

### Accent Stripe / Divider
```vba
Sub CreateAccentStripe(ws As Worksheet, left As Single, top As Single, _
                       width As Single, accentColor As Long)
    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(msoShapeRectangle, _
        left, top, width, 3)  ' 3pt tall = thin stripe
    
    With shp
        .Fill.Solid
        .Fill.ForeColor.RGB = accentColor
        .Line.Visible = msoFalse
    End With
End Sub
```

### KPI Card (Complete Component)
```vba
Sub CreateKPICard(ws As Worksheet, left As Single, top As Single, _
                  cardWidth As Single, cardHeight As Single, _
                  title As String, valueCell As String, _
                  deltaText As String, deltaIsPositive As Boolean, _
                  palette As String)
    
    Dim bgColor As Long, textColor As Long, accentColor As Long
    Dim mutedColor As Long, successColor As Long, dangerColor As Long
    
    ' Set colors based on palette
    Select Case palette
        Case "dark"
            bgColor = RGB(30, 58, 95)
            textColor = RGB(255, 255, 255)
            accentColor = RGB(0, 229, 255)
            mutedColor = RGB(136, 153, 170)
            successColor = RGB(102, 187, 106)
            dangerColor = RGB(239, 83, 80)
        Case "light"
            bgColor = RGB(255, 255, 255)
            textColor = RGB(27, 42, 74)
            accentColor = RGB(41, 121, 255)
            mutedColor = RGB(107, 123, 141)
            successColor = RGB(46, 125, 50)
            dangerColor = RGB(198, 40, 40)
    End Select
    
    ' 1. Card background
    Dim card As Shape
    Set card = ws.Shapes.AddShape(msoShapeRoundedRectangle, _
        left, top, cardWidth, cardHeight)
    With card
        .Fill.Solid
        .Fill.ForeColor.RGB = bgColor
        .Line.Visible = msoFalse
        .Adjustments(1) = 0.05
        If palette = "light" Then
            .Shadow.Visible = msoTrue
            .Shadow.ForeColor.RGB = RGB(0, 0, 0)
            .Shadow.Transparency = 0.85
            .Shadow.OffsetX = 2
            .Shadow.OffsetY = 2
            .Shadow.Blur = 6
        End If
    End With
    
    ' 2. Title label
    CreateTextShape ws, left + 12, top + 8, cardWidth - 24, 18, _
        UCase(title), 9, mutedColor, False
    
    ' 3. Big number (linked to cell)
    CreateLinkedMetric ws, left + 12, top + 28, cardWidth - 24, 36, _
        valueCell, 28, textColor
    
    ' 4. Delta indicator
    Dim deltaColor As Long
    deltaColor = IIf(deltaIsPositive, successColor, dangerColor)
    Dim arrow As String
    arrow = IIf(deltaIsPositive, ChrW(&H25B2) & " ", ChrW(&H25BC) & " ")
    CreateTextShape ws, left + 12, top + 68, cardWidth - 24, 16, _
        arrow & deltaText, 10, deltaColor, False
    
    ' 5. Accent stripe at top
    CreateAccentStripe ws, left + 2, top + 2, cardWidth - 4, accentColor
End Sub
```

---

## CHART STYLING (Making Charts Not Look Like Excel)

### The 6-Step Chart Restyle Process
1. **Delete unnecessary elements:** Remove title (use card title instead), remove background fill, remove border
2. **Set chart area to transparent:** `ChartArea.Format.Fill.Visible = msoFalse`
3. **Set plot area to transparent:** `PlotArea.Format.Fill.Visible = msoFalse`
4. **Restyle axes:** Change font to match palette, make gridlines very faint or remove them
5. **Recolor data series:** Apply palette colors to each series
6. **Simplify:** Remove legends if only 1 series, use data labels instead of Y-axis

```vba
Sub RestyleChart(cht As ChartObject, palette As String)
    Dim textColor As Long, gridColor As Long
    
    Select Case palette
        Case "dark"
            textColor = RGB(136, 153, 170)
            gridColor = RGB(45, 64, 87)
        Case "light"
            textColor = RGB(107, 123, 141)
            gridColor = RGB(230, 234, 238)
    End Select
    
    With cht.Chart
        ' Transparent backgrounds
        .ChartArea.Format.Fill.Visible = msoFalse
        .ChartArea.Format.Line.Visible = msoFalse
        .PlotArea.Format.Fill.Visible = msoFalse
        .PlotArea.Format.Line.Visible = msoFalse
        
        ' Remove title (card provides context)
        .HasTitle = False
        
        ' Style axes
        If .HasAxis(xlCategory) Then
            With .Axes(xlCategory)
                .TickLabels.Font.Color = textColor
                .TickLabels.Font.Size = 8
                .TickLabels.Font.Name = "Aptos"
                .Format.Line.Visible = msoFalse
                .MajorTickMark = xlNone
            End With
        End If
        
        If .HasAxis(xlValue) Then
            With .Axes(xlValue)
                .TickLabels.Font.Color = textColor
                .TickLabels.Font.Size = 8
                .TickLabels.Font.Name = "Aptos"
                .Format.Line.Visible = msoFalse
                .MajorTickMark = xlNone
                .HasMajorGridlines = True
                .MajorGridlines.Format.Line.ForeColor.RGB = gridColor
                .MajorGridlines.Format.Line.Weight = 0.5
                .MajorGridlines.Format.Line.DashStyle = msoLineSolid
            End With
        End If
        
        ' Remove legend if single series
        If .SeriesCollection.Count = 1 Then
            .HasLegend = False
        Else
            .HasLegend = True
            .Legend.Font.Color = textColor
            .Legend.Font.Size = 8
            .Legend.Font.Name = "Aptos"
            .Legend.Format.Fill.Visible = msoFalse
            .Legend.Format.Line.Visible = msoFalse
        End If
    End With
End Sub
```

### Recolor Chart Series
```vba
Sub RecolorSeries(cht As ChartObject, palette As String)
    Dim colors() As Long
    
    Select Case palette
        Case "dark"
            colors = Array(RGB(79, 195, 247), RGB(255, 109, 0), _
                          RGB(102, 187, 106), RGB(171, 71, 188), _
                          RGB(236, 64, 122), RGB(255, 167, 38))
        Case "light"
            colors = Array(RGB(41, 121, 255), RGB(255, 109, 0), _
                          RGB(46, 125, 50), RGB(123, 31, 162), _
                          RGB(198, 40, 40), RGB(0, 131, 143))
    End Select
    
    Dim i As Long
    For i = 1 To cht.Chart.SeriesCollection.Count
        If i - 1 <= UBound(colors) Then
            With cht.Chart.SeriesCollection(i)
                .Format.Fill.ForeColor.RGB = colors(i - 1)
                .Format.Line.ForeColor.RGB = colors(i - 1)
                .Format.Fill.Solid
                .Format.Fill.ForeColor.RGB = colors(i - 1)
            End With
        End If
    Next i
End Sub
```

---

## PAGE SETUP (Before Any Design Work)

**⚠️ COM SAFETY: When running via MCP (excel_set_vba_code + excel_run_vba_macro):**
- NEVER use `ActiveWindow` — use `ThisWorkbook.Windows(1)` instead
- NEVER use `ActiveSheet` — use explicit `ThisWorkbook.Sheets("SheetName")`
- NEVER use MsgBox, InputBox, Application.Dialogs, UserForm.Show
- ALWAYS wrap in `Application.ScreenUpdating = False / True`

### COM-Safe Canvas Setup:
```vba
Sub PrepareCanvas()
    Application.ScreenUpdating = False
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Dashboard")
    
    ' Background
    ws.Cells.Interior.Color = RGB(27, 42, 74)
    
    ' Hide gridlines (COM-safe)
    ThisWorkbook.Windows(1).DisplayGridlines = False
    
    Application.ScreenUpdating = True
End Sub
```

---

## COMPLETE DASHBOARD BUILD SEQUENCE (for MCP)

### Phase 1: Canvas Setup
1. `excel_create_sheet` — create "Dashboard" sheet if needed
2. `excel_set_vba_code` — write PrepareCanvas macro
3. `excel_run_vba_macro` — execute it (sets background, hides gridlines)

### Phase 2: Layout Structure
4. Write VBA that creates all card backgrounds in a grid layout
5. Include accent stripes, dividers, title bar shape
6. Run the macro — this creates the visual skeleton

### Phase 3: Content Layer
7. Write VBA that creates text shapes for titles, metrics, deltas
8. Link dynamic text shapes to data cells (`.DrawingObject.Formula`)
9. Run the macro

### Phase 4: Charts
10. `excel_create_chart` — create charts on the sheet
11. Write VBA to restyle each chart (transparent bg, palette colors, clean axes)
12. Write VBA to position/resize charts to sit inside card boundaries
13. Run the macros

### Phase 5: Polish
14. Write VBA for final alignment pass (`.Align`, `.Distribute`)
15. Group related shapes (card bg + its labels + its chart)
16. `excel_capture_screenshot` — verify the result

---

## ADVANCED TECHNIQUES

### In-Cell Bar Charts (No Shape Needed)
```vba
' Uses REPT function with block character
' Cell formula: =REPT("█", ROUND(value/maxValue * 20, 0))
' Format cell font color to accent color
' This creates horizontal bars purely in cells
```

### Sparklines via VBA
```vba
Sub AddSparkline(ws As Worksheet, dataRange As String, targetCell As String)
    ws.Range(targetCell).SparklineGroups.Add _
        Type:=xlSparkLine, _
        SourceData:=dataRange
    
    With ws.Range(targetCell).SparklineGroups(1)
        .SeriesColor.Color = RGB(79, 195, 247)
        .Points.Highpoint.Visible = True
        .Points.Highpoint.Color.Color = RGB(102, 187, 106)
        .LineWeight = 1.5
    End With
End Sub
```

### Navigation Buttons
```vba
Sub CreateNavButton(ws As Worksheet, left As Single, top As Single, _
                    caption As String, macroName As String, _
                    fillColor As Long, textColor As Long)
    Dim btn As Shape
    Set btn = ws.Shapes.AddShape(msoShapeRoundedRectangle, _
        left, top, 120, 32)
    
    With btn
        .Fill.Solid
        .Fill.ForeColor.RGB = fillColor
        .Line.Visible = msoFalse
        .Adjustments(1) = 0.3
        .TextFrame2.TextRange.Text = caption
        .TextFrame2.TextRange.Font.Size = 10
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = textColor
        .TextFrame2.TextRange.Font.Bold = True
        .TextFrame2.TextRange.Font.Name = "Aptos"
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .OnAction = macroName
    End With
End Sub
```

### Title Bar with Gradient
```vba
Sub CreateTitleBar(ws As Worksheet, title As String, _
                   pageWidth As Single, palette As String)
    Dim bar As Shape
    Set bar = ws.Shapes.AddShape(msoShapeRectangle, _
        0, 0, pageWidth, 56)
    
    Dim c1 As Long, c2 As Long, txtColor As Long
    Select Case palette
        Case "dark"
            c1 = RGB(27, 42, 74): c2 = RGB(15, 25, 50)
            txtColor = RGB(255, 255, 255)
        Case "light"
            c1 = RGB(27, 42, 74): c2 = RGB(46, 64, 87)
            txtColor = RGB(255, 255, 255)
    End Select
    
    With bar.Fill
        .TwoColorGradient msoGradientHorizontal, 1
        .ForeColor.RGB = c1
        .BackColor.RGB = c2
    End With
    bar.Line.Visible = msoFalse
    
    Dim ttl As Shape
    Set ttl = ws.Shapes.AddShape(msoShapeRectangle, _
        20, 8, pageWidth - 40, 40)
    With ttl
        .Fill.Visible = msoFalse
        .Line.Visible = msoFalse
        .TextFrame2.TextRange.Text = title
        .TextFrame2.TextRange.Font.Size = 18
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = txtColor
        .TextFrame2.TextRange.Font.Bold = True
        .TextFrame2.TextRange.Font.Name = "Aptos"
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
    End With
End Sub
```

### Row Banding for Data Tables
```vba
Sub ApplyRowBanding(ws As Worksheet, dataRange As String, _
                    evenColor As Long, oddColor As Long)
    Dim rng As Range, row As Range
    Set rng = ws.Range(dataRange)
    
    Dim i As Long: i = 0
    For Each row In rng.Rows
        If i Mod 2 = 0 Then
            row.Interior.Color = evenColor
        Else
            row.Interior.Color = oddColor
        End If
        i = i + 1
    Next row
End Sub
```

---

## ALIGNMENT & GROUPING

```vba
Sub AlignAndGroup(ws As Worksheet, shapeNames As Variant)
    Dim sr As ShapeRange
    Set sr = ws.Shapes.Range(shapeNames)
    
    sr.Align msoAlignTops, msoFalse
    sr.Distribute msoDistributeHorizontally, msoFalse
    sr.Group
End Sub
```

---

## WHAT "ADVANCED" LOOKS LIKE (Text Descriptions)

### Example 1: Sales KPI Dashboard (Dark Theme)
Deep navy background. Top: gradient title bar spanning full width, white text "Q4 Performance Dashboard", subtle logo shape on right. Below title: 4 equal-width KPI cards in a row — Revenue, Orders, Avg Order Value, Customer Retention. Each card has: thin cyan accent stripe at top, muted gray label, large white number linked to cell, green/red delta with arrow. Below KPI row: 2 cards side by side — left card (2/3 width) contains a column chart showing monthly revenue with cyan bars on transparent background, right card (1/3 width) shows a donut chart of revenue by category. Bottom: full-width card with a styled table using row banding (alternating dark navy shades), no gridlines, data in white text.

### Example 2: Executive Summary (Light Theme)
Light gray-blue background (#F5F6FA). Cards are white with subtle drop shadows. Title bar is dark navy with white text. KPI cards have a colored left-border accent (different color per KPI) instead of top stripe. Charts use the blue/orange/green series palette. Table section uses very light gray row banding. Everything feels clean, airy, professional — like a web dashboard rendered in Excel.

### Example 3: Infographic Style
Uses Excel shapes to create a visual flow — circular shapes for key metrics, dotted lines connecting sections to guide the eye, sunburst chart as centerpiece, image fills in shapes for visual interest. Cards arranged in non-grid layout (offset, varied sizes). Uses transparency overlays for depth. This pushes Excel to its limits — looks like it was built in a design tool.

---

## TYPOGRAPHY SCALE

| Element | Size | Weight | Color |
|---------|------|--------|-------|
| Dashboard title | 18-22pt | Bold | White or primary |
| Card title / label | 8-9pt | Regular | Muted/secondary |
| Big metric number | 24-32pt | Bold | White or accent |
| Delta / change indicator | 9-10pt | Regular | Success/danger |
| Chart axis labels | 8pt | Regular | Muted/secondary |
| Table headers | 10pt | Bold | White on dark bg |
| Table data | 10pt | Regular | Primary text |
| Helper/footnote text | 8pt | Regular | Muted/secondary |

**Font: Aptos (modern, Windows 11 default) → Segoe UI → Calibri (fallback)**
Pick ONE. Never mix.

---

## DESIGN QUALITY CHECKLIST

Before calling `excel_capture_screenshot` to verify:

- [ ] No gridlines visible (DisplayGridlines = False)
- [ ] Background color set on entire sheet (not just used range)
- [ ] One font family only (Aptos or Segoe UI or Calibri — pick one)
- [ ] One color palette only — no Excel default blues/oranges mixed in
- [ ] All cards same height within their row
- [ ] Consistent gaps between all cards (8-12px everywhere)
- [ ] Charts have transparent backgrounds — they sit ON cards, not in boxes
- [ ] No chart borders or chart area fills
- [ ] Axis labels match palette text color (not default black)
- [ ] Gridlines either removed or very faint (0.5pt, muted color)
- [ ] Big numbers are BIG (24-32pt) — small labels are SMALL (8-9pt)
- [ ] Accent stripe or color indicator on each card for visual interest
- [ ] Number formats applied — no raw decimals or "General" format
- [ ] Title bar spans full width with gradient fill
- [ ] No default Excel chart colors — all series manually recolored

---

## ANTI-PATTERNS (Never Do These)

- ❌ Default Excel chart colors (the blue/orange/gray defaults)
- ❌ Visible gridlines on a dashboard sheet
- ❌ 3D chart effects (3D bars, 3D pie, perspective)
- ❌ Chart borders or white chart backgrounds sitting on colored cards
- ❌ Multiple font families
- ❌ Thick cell borders (use shapes instead)
- ❌ Data in raw cells without card containers
- ❌ Pie charts with more than 5 segments
- ❌ Gradient fills on text (only on shape backgrounds)
- ❌ MsgBox, InputBox, or any dialog in VBA (blocks COM automation)
- ❌ ActiveWindow / ActiveSheet references in VBA (COM unsafe)
- ❌ Wall-of-numbers without visual hierarchy (big/small/muted/bold)
