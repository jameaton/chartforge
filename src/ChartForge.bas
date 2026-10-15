Attribute VB_Name = "ChartForge"
' ChartForge: native, editable charts for PowerPoint, drawn as grouped
' shapes in a consistent house style. The ribbon buttons call the AF_* subs at the end
' of this module.
'
' Data entry happens in an ordinary PowerPoint table (the "datasheet"). A
' chart stores its datasheet and options in tags on its group, so Edit data
' can bring the table back and any button can redraw the chart.
Option Explicit

' ---- Brand palette ------------------------------------------------------
Private Const BLACK As String = "#0B0B0B"
Private Const WHITE As String = "#FFFFFF"
Private Const YELLOW As String = "#FFED00"
Private Const GRAY1 As String = "#575756"
Private Const GRAY2 As String = "#A8A8A8"
Private Const GRAY3 As String = "#E5E5E5"
Private Const BLUE500 As String = "#0047BA"
Private Const BLUE400 As String = "#4D7ECF"
Private Const BLUE300 As String = "#99B5E3"
Private Const BLUE700 As String = "#00358C"

Private Const DEFAULT_FONT As String = "IBM Plex Sans"

Private Type ChartOpts
    chartType As String
    width As Double
    height As Double
    labels As Boolean
    totals As Boolean
    legend As Boolean
    decimals As Long        ' -1 = automatic
    prefix As String
    suffix As String
    arrowMode As String     ' none | cagr | pct | diff
    arrowFrom As Long       ' 1-based category
    arrowTo As Long         ' 0 = last category
    fontSize As Double
    fontName As String
End Type

' ---- Drawing state (set per chart) -------------------------------------
Private gSld As Slide
Private gOX As Double, gOY As Double, gW As Double, gH As Double
Private gNames() As String
Private gCount As Long
Private gFont As String
Private gPrefix As String
Private gTop As Double, gBottom As Double, gMax As Double, gMin As Double

' ---- Parsed data (set per chart) ---------------------------------------
Private gN As Long, gM As Long
Private gCats() As String, gCatText() As String, gSerNames() As String
Private gV() As Double, gHas() As Boolean, gIsE() As Boolean, gWidths() As Double
Private gTaskName() As String, gStart() As Double, gEnd() As Double, gHasEnd() As Boolean
Private gDates As Boolean
Private gLabelW As Double, gTMin As Double, gTMax As Double
Private gTickT() As Double, gTickL() As String, gNT As Long, gEndT As Double

Private gHintsShown As String

' ========================================================================
' Small helpers
' ========================================================================

Private Function Mx(ByVal a As Double, ByVal b As Double) As Double
    If a > b Then Mx = a Else Mx = b
End Function

Private Function Mn(ByVal a As Double, ByVal b As Double) As Double
    If a < b Then Mn = a Else Mn = b
End Function

Private Function RGBHex(ByVal hex As String) As Long
    RGBHex = RGB(CLng("&H" & Mid$(hex, 2, 2)), CLng("&H" & Mid$(hex, 4, 2)), CLng("&H" & Mid$(hex, 6, 2)))
End Function

' Rough average glyph width for a proportional sans font.
Private Function TextW(ByVal s As String, ByVal size As Double) As Double
    TextW = Len(s) * size * 0.56
End Function

' Shrinks a font so a label of maxLen characters fits a slot, within a floor.
Private Function FitSize(ByVal maxLen As Double, ByVal slot As Double, ByVal fs As Double) As Double
    If maxLen < 1 Then maxLen = 1
    FitSize = Mx(6, Mn(fs, (slot - 2) / (maxLen * 0.56)))
End Function

' Black or white, whichever reads better on the given fill.
Private Function ContrastText(ByVal hex As String) As String
    Dim r As Double, g As Double, b As Double
    r = CLng("&H" & Mid$(hex, 2, 2)) / 255
    g = CLng("&H" & Mid$(hex, 4, 2)) / 255
    b = CLng("&H" & Mid$(hex, 6, 2)) / 255
    If 0.2126 * r + 0.7152 * g + 0.0722 * b > 0.55 Then ContrastText = BLACK Else ContrastText = WHITE
End Function

Private Function SeriesColor(ByVal i As Long) As String
    Dim p As Variant
    p = Array(BLUE500, BLUE300, GRAY1, YELLOW, BLUE700, GRAY2, BLUE400, BLACK)
    SeriesColor = p((i - 1) Mod 8)
End Function

Private Function Fmt(ByVal v As Double, ByVal prefix As String, ByVal suffix As String, ByVal dec As Long) As String
    Dim pat As String, s As String
    If dec > 0 Then pat = "#,##0." & String$(dec, "0") Else pat = "#,##0"
    s = Format$(Abs(v), pat)
    If v < 0 And Val(Replace(s, ",", "")) <> 0 Then
        Fmt = ChrW(8722) & prefix & s & suffix
    Else
        Fmt = prefix & s & suffix
    End If
End Function

Private Function Signed(ByVal v As Double, ByVal dec As Long, ByVal suffix As String) As String
    Dim pat As String, s As String
    If dec > 0 Then pat = "0." & String$(dec, "0") Else pat = "0"
    s = Format$(Abs(v), pat)
    If Val(s) = 0 Then
        Signed = "0" & suffix
    ElseIf v > 0 Then
        Signed = "+" & s & suffix
    Else
        Signed = ChrW(8722) & s & suffix
    End If
End Function

Private Function ParseNum(ByVal s As String, ByRef v As Double) As Boolean
    Dim t As String
    t = Trim$(s)
    t = Replace(t, ",", "")
    t = Replace(t, " ", "")
    t = Replace(t, "%", "")
    t = Replace(t, "$", "")
    t = Replace(t, ChrW(163), "")
    t = Replace(t, ChrW(8364), "")
    t = Replace(t, ChrW(8722), "-")
    If Len(t) > 2 And Left$(t, 1) = "(" And Right$(t, 1) = ")" Then t = "-" & Mid$(t, 2, Len(t) - 2)
    If t = "" Then Exit Function
    If Not IsNumeric(t) Then Exit Function
    v = Val(t)
    ParseNum = True
End Function

' Dates as YYYY-MM-DD, YYYY-MM or DD/MM/YYYY (UK order), else plain numbers.
Private Function ParseTime(ByVal s As String, ByRef t As Double, ByRef isDt As Boolean) As Boolean
    Dim p As Variant, d As Long
    s = Trim$(s)
    If s = "" Then Exit Function
    p = Split(s, "-")
    If UBound(p) >= 1 Then
        If Len(p(0)) = 4 And IsNumeric(p(0)) Then
            d = 1
            If UBound(p) >= 2 Then d = Val(p(2))
            t = CDbl(DateSerial(Val(p(0)), Val(p(1)), d))
            isDt = True
            ParseTime = True
            Exit Function
        End If
    End If
    p = Split(s, "/")
    If UBound(p) = 2 Then
        If Len(p(2)) = 4 Then
            t = CDbl(DateSerial(Val(p(2)), Val(p(1)), Val(p(0))))
            isDt = True
            ParseTime = True
            Exit Function
        End If
    End If
    Dim v As Double
    If ParseNum(s, v) Then
        t = v
        isDt = False
        ParseTime = True
    End If
End Function

Private Function GridKey(ByVal t As String) As String
    Select Case t
        Case "stacked", "clustered", "percent": GridKey = "column"
        Case Else: GridKey = t
    End Select
End Function

Private Function TypeLabel(ByVal t As String) As String
    Select Case t
        Case "stacked": TypeLabel = "Stacked"
        Case "clustered": TypeLabel = "Clustered"
        Case "percent": TypeLabel = "100%"
        Case "waterfall": TypeLabel = "Waterfall"
        Case "mekko": TypeLabel = "Mekko"
        Case "gantt": TypeLabel = "Gantt"
    End Select
End Function

Private Sub Hint(ByVal key As String, ByVal msg As String)
    If InStr(gHintsShown, "[" & key & "]") > 0 Then Exit Sub
    gHintsShown = gHintsShown & "[" & key & "]"
    MsgBox msg, vbInformation, "ChartForge"
End Sub

' ========================================================================
' Grids: 2D String arrays, 1-based, row 1 = header
' ========================================================================

Private Function GCell(g As Variant, ByVal r As Long, ByVal c As Long) As String
    If r <= UBound(g, 1) And c <= UBound(g, 2) Then GCell = Trim$(CStr(g(r, c)))
End Function

Private Function RowBlank(g As Variant, ByVal r As Long) As Boolean
    Dim c As Long
    For c = 1 To UBound(g, 2)
        If GCell(g, r, c) <> "" Then Exit Function
    Next
    RowBlank = True
End Function

' Drops trailing empty rows and columns, so stray cells don't add bars.
Private Function TrimGrid(g As Variant) As Variant
    Dim r As Long, c As Long, lastR As Long, lastC As Long
    lastR = 1
    lastC = 1
    For r = 1 To UBound(g, 1)
        For c = 1 To UBound(g, 2)
            If GCell(g, r, c) <> "" Then
                If r > lastR Then lastR = r
                If c > lastC Then lastC = c
            End If
        Next
    Next
    Dim out() As String
    ReDim out(1 To lastR, 1 To lastC)
    For r = 1 To lastR
        For c = 1 To lastC
            out(r, c) = GCell(g, r, c)
        Next
    Next
    TrimGrid = out
End Function

Private Function Esc(ByVal s As String) As String
    Esc = Replace(Replace(Replace(s, "%", "%25"), "|", "%7C"), "^", "%5E")
End Function

Private Function Unesc(ByVal s As String) As String
    Unesc = Replace(Replace(Replace(s, "%7C", "|"), "%5E", "^"), "%25", "%")
End Function

' Cells joined by "|", rows by "^".
Private Function SerializeGrid(g As Variant) As String
    Dim r As Long, c As Long, s As String
    For r = 1 To UBound(g, 1)
        If r > 1 Then s = s & "^"
        For c = 1 To UBound(g, 2)
            If c > 1 Then s = s & "|"
            s = s & Esc(GCell(g, r, c))
        Next
    Next
    SerializeGrid = s
End Function

Private Function ParseGridText(ByVal s As String) As Variant
    Dim rows As Variant, cells As Variant, r As Long, c As Long, nCols As Long
    rows = Split(s, "^")
    nCols = 1
    For r = 0 To UBound(rows)
        cells = Split(rows(r), "|")
        If UBound(cells) + 1 > nCols Then nCols = UBound(cells) + 1
    Next
    Dim g() As String
    ReDim g(1 To UBound(rows) + 1, 1 To nCols)
    For r = 0 To UBound(rows)
        cells = Split(rows(r), "|")
        For c = 0 To UBound(cells)
            g(r + 1, c + 1) = Unesc(cells(c))
        Next
    Next
    ParseGridText = g
End Function

Private Function SampleGrid(ByVal t As String) As Variant
    Select Case GridKey(t)
        Case "waterfall"
            SampleGrid = ParseGridText("|2024|New launches|Pricing|Loss of exclusivity|FX|2025^" & _
                "Revenue|100|25|8|-18|-4|e")
        Case "mekko"
            SampleGrid = ParseGridText("|US|EU|China|Japan|RoW^" & _
                "Market size|120|80|50|30|40^" & _
                "Pfizer|30|25|10|8|10^" & _
                "Moderna|25|15|5|4|6^" & _
                "GSK|20|18|5|6|8^" & _
                "Other|45|22|30|12|16")
        Case "gantt"
            SampleGrid = ParseGridText("Task|Start|End^" & _
                "Phase I|2026-01-01|2026-06-30^" & _
                "Phase II|2026-05-01|2027-03-31^" & _
                "Phase III|2027-02-01|2028-06-30^" & _
                "Filing|2028-09-01|^" & _
                "Approval|2029-03-01|")
        Case Else
            SampleGrid = ParseGridText("|2022|2023|2024|2025|2026^" & _
                "Vaccines|12|15|19|22|27^" & _
                "Therapeutics|8|9|12|14|15^" & _
                "Diagnostics|4|5|5|6|8")
    End Select
End Function

' Reads the datasheet into the module's parsed-data arrays.
Private Sub ParseFor(ByVal t As String, g As Variant)
    Dim tg As Variant, nR As Long, nC As Long, r As Long, c As Long, s As Long, k As Long
    Dim v As Double, txt As String
    tg = TrimGrid(g)
    nR = UBound(tg, 1)
    nC = UBound(tg, 2)

    If t = "gantt" Then
        Dim st As Double, en As Double, d1 As Boolean, d2 As Boolean
        gM = 0
        gDates = False
        ReDim gTaskName(1 To nR)
        ReDim gStart(1 To nR)
        ReDim gEnd(1 To nR)
        ReDim gHasEnd(1 To nR)
        For r = 2 To nR
            d1 = False
            d2 = False
            If ParseTime(GCell(tg, r, 2), st, d1) Then
                gM = gM + 1
                gTaskName(gM) = GCell(tg, r, 1)
                gStart(gM) = st
                gHasEnd(gM) = ParseTime(GCell(tg, r, 3), en, d2)
                If gHasEnd(gM) Then gEnd(gM) = en Else gEnd(gM) = st
                If d1 Or d2 Then gDates = True
            End If
        Next
        If gM = 0 Then Err.Raise vbObjectError + 10, , "Add at least one task with a start date."
        Exit Sub
    End If

    gN = nC - 1
    If gN < 1 Then Err.Raise vbObjectError + 11, , "Add at least one category column."
    ReDim gCats(1 To gN)
    For c = 2 To nC
        gCats(c - 1) = GCell(tg, 1, c)
    Next

    Dim dataRows() As Long, nd As Long
    ReDim dataRows(1 To nR)
    For r = 2 To nR
        If Not RowBlank(tg, r) Then
            nd = nd + 1
            dataRows(nd) = r
        End If
    Next

    Dim first As Long
    first = 1
    If t = "mekko" Then
        If nd < 2 Then Err.Raise vbObjectError + 12, , "A Mekko needs a width row and at least one series row."
        ReDim gWidths(1 To gN)
        For c = 1 To gN
            If ParseNum(GCell(tg, dataRows(1), c + 1), v) Then gWidths(c) = v Else gWidths(c) = 0
        Next
        first = 2
    End If

    gM = nd - first + 1
    If gM < 1 Then Err.Raise vbObjectError + 13, , "Add at least one row of data."
    ReDim gSerNames(1 To gM)
    ReDim gV(1 To gM, 1 To gN)
    ReDim gHas(1 To gM, 1 To gN)
    ReDim gIsE(1 To gN)
    For s = 1 To gM
        r = dataRows(first + s - 1)
        gSerNames(s) = GCell(tg, r, 1)
        For c = 1 To gN
            txt = GCell(tg, r, c + 1)
            If t = "waterfall" And s = 1 And LCase$(txt) = "e" Then
                gIsE(c) = True
            ElseIf ParseNum(txt, v) Then
                gV(s, c) = v
                gHas(s, c) = True
            End If
        Next
    Next
End Sub

Private Function Decimals(o As ChartOpts) As Long
    Dim s As Long, i As Long
    If o.decimals >= 0 Then
        Decimals = o.decimals
        Exit Function
    End If
    If o.chartType = "mekko" Then
        For i = 1 To gN
            If gWidths(i) <> Int(gWidths(i)) Then
                Decimals = 1
                Exit Function
            End If
        Next
        Exit Function
    End If
    For s = 1 To gM
        For i = 1 To gN
            If gV(s, i) <> Int(gV(s, i)) Then
                Decimals = 1
                Exit Function
            End If
        Next
    Next
End Function

' ========================================================================
' Shape primitives. Coordinates are points relative to the chart's
' top-left corner; everything is clipped to the chart frame so the group's
' bounds equal the chart size and resizing round-trips.
' ========================================================================

Private Sub Track(ByVal shp As Shape)
    gCount = gCount + 1
    shp.Name = gPrefix & gCount
    ReDim Preserve gNames(1 To gCount)
    gNames(gCount) = shp.Name
End Sub

Private Function ClampBox(ByRef x As Double, ByRef y As Double, ByRef w As Double, ByRef h As Double) As Boolean
    Dim x2 As Double, y2 As Double
    x2 = Mn(gW, x + w)
    y2 = Mn(gH, y + h)
    x = Mx(0, x)
    y = Mx(0, y)
    w = x2 - x
    h = y2 - y
    ClampBox = (w > 0.01 And h > 0.01)
End Function

Private Sub StyleShape(ByVal shp As Shape, ByVal fill As String, ByVal stroke As String, ByVal strokeW As Double)
    If fill = "" Then
        shp.Fill.Visible = msoFalse
    Else
        shp.Fill.Visible = msoTrue
        shp.Fill.Solid
        shp.Fill.ForeColor.RGB = RGBHex(fill)
        shp.Fill.Transparency = 0
    End If
    If stroke = "" Then
        shp.Line.Visible = msoFalse
    Else
        shp.Line.Visible = msoTrue
        shp.Line.ForeColor.RGB = RGBHex(stroke)
        shp.Line.Weight = strokeW
    End If
    shp.Shadow.Visible = msoFalse
End Sub

Private Sub AddRect(ByVal x As Double, ByVal y As Double, ByVal w As Double, ByVal h As Double, ByVal fill As String, _
        Optional ByVal stroke As String = "", Optional ByVal strokeW As Double = 0, Optional ByVal rounded As Boolean = False)
    If Not ClampBox(x, y, w, h) Then Exit Sub
    Dim shp As Shape
    If rounded Then
        Set shp = gSld.Shapes.AddShape(msoShapeRoundedRectangle, gOX + x, gOY + y, w, h)
        shp.Adjustments(1) = 0.5
    Else
        Set shp = gSld.Shapes.AddShape(msoShapeRectangle, gOX + x, gOY + y, w, h)
    End If
    StyleShape shp, fill, stroke, strokeW
    Track shp
End Sub

Private Sub AddDiamond(ByVal x As Double, ByVal y As Double, ByVal w As Double, ByVal h As Double, ByVal fill As String)
    If Not ClampBox(x, y, w, h) Then Exit Sub
    Dim shp As Shape
    Set shp = gSld.Shapes.AddShape(msoShapeDiamond, gOX + x, gOY + y, w, h)
    StyleShape shp, fill, "", 0
    Track shp
End Sub

Private Sub AddTriDown(ByVal x As Double, ByVal y As Double, ByVal w As Double, ByVal h As Double, ByVal fill As String)
    If Not ClampBox(x, y, w, h) Then Exit Sub
    Dim shp As Shape
    Set shp = gSld.Shapes.AddShape(msoShapeIsoscelesTriangle, gOX + x, gOY + y, w, h)
    shp.Flip msoFlipVertical
    StyleShape shp, fill, "", 0
    Track shp
End Sub

Private Sub AddLn(ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double, _
        ByVal color As String, ByVal weight As Double, Optional ByVal dash As Boolean = False)
    x1 = Mn(gW, Mx(0, x1))
    x2 = Mn(gW, Mx(0, x2))
    y1 = Mn(gH, Mx(0, y1))
    y2 = Mn(gH, Mx(0, y2))
    If Abs(x2 - x1) < 0.01 And Abs(y2 - y1) < 0.01 Then Exit Sub
    Dim shp As Shape
    Set shp = gSld.Shapes.AddLine(gOX + x1, gOY + y1, gOX + x2, gOY + y2)
    shp.Line.Visible = msoTrue
    shp.Line.ForeColor.RGB = RGBHex(color)
    shp.Line.Weight = weight
    If dash Then shp.Line.DashStyle = msoLineDash
    Track shp
End Sub

Private Sub AddText(ByVal x As Double, ByVal y As Double, ByVal w As Double, ByVal h As Double, ByVal txt As String, _
        ByVal size As Double, Optional ByVal color As String = BLACK, Optional ByVal align As String = "c", _
        Optional ByVal bold As Boolean = False)
    If Not ClampBox(x, y, w, h) Then Exit Sub
    Dim shp As Shape
    Set shp = gSld.Shapes.AddTextbox(msoTextOrientationHorizontal, gOX + x, gOY + y, w, h)
    With shp.TextFrame
        .WordWrap = msoFalse
        .AutoSize = ppAutoSizeNone
        .MarginLeft = 0
        .MarginRight = 0
        .MarginTop = 0
        .MarginBottom = 0
        .VerticalAnchor = msoAnchorMiddle
        .TextRange.Text = txt
        With .TextRange.Font
            .Name = gFont
            .Size = Round(size * 2) / 2
            .Color.RGB = RGBHex(color)
            If bold Then .Bold = msoTrue Else .Bold = msoFalse
        End With
        Select Case align
            Case "l": .TextRange.ParagraphFormat.Alignment = ppAlignLeft
            Case "r": .TextRange.ParagraphFormat.Alignment = ppAlignRight
            Case Else: .TextRange.ParagraphFormat.Alignment = ppAlignCenter
        End Select
    End With
    shp.Left = gOX + x
    shp.Top = gOY + y
    shp.Width = w
    shp.Height = h
    Track shp
End Sub

' Value axis: maps a data value to a y position in the plot area.
Private Function YV(ByVal v As Double) As Double
    YV = gTop + (gMax - v) / (gMax - gMin) * (gBottom - gTop)
End Function

' ========================================================================
' Chart pieces shared by several layouts
' ========================================================================

' Splits a label into two lines at the space that best balances them.
Private Function SplitTwo(ByVal s As String) As Variant
    Dim words As Variant, i As Long, j As Long, a As String, b As String
    Dim bestA As String, bestB As String, best As Double
    words = Split(s, " ")
    If UBound(words) < 1 Then
        SplitTwo = Array(s)
        Exit Function
    End If
    best = 1000000
    For i = 1 To UBound(words)
        a = ""
        b = ""
        For j = 0 To i - 1
            If j > 0 Then a = a & " "
            a = a & words(j)
        Next
        For j = i To UBound(words)
            If j > i Then b = b & " "
            b = b & words(j)
        Next
        If Mx(Len(a), Len(b)) < best Then
            best = Mx(Len(a), Len(b))
            bestA = a
            bestB = b
        End If
    Next
    SplitTwo = Array(bestA, bestB)
End Function

' Category labels under equal-width slots: one line at full size when they
' fit, otherwise wrapped onto two lines before shrinking the font.
Private Sub CategoryLabels(ByVal slot As Double, ByVal fs As Double, ByRef size As Double, ByRef h As Double)
    Dim i As Long, fits As Boolean, maxLen As Double, twoLines As Boolean, parts As Variant
    ReDim gCatText(1 To gN)
    fits = True
    For i = 1 To gN
        If TextW(gCats(i), fs) > slot - 2 Then fits = False
    Next
    If fits Then
        For i = 1 To gN
            gCatText(i) = gCats(i)
        Next
        size = fs
        h = fs + 8
        Exit Sub
    End If
    maxLen = 1
    For i = 1 To gN
        parts = SplitTwo(gCats(i))
        If UBound(parts) = 1 Then
            gCatText(i) = parts(0) & vbCr & parts(1)
            twoLines = True
            maxLen = Mx(maxLen, Mx(Len(parts(0)), Len(parts(1))))
        Else
            gCatText(i) = parts(0)
            maxLen = Mx(maxLen, Len(parts(0)))
        End If
    Next
    size = FitSize(maxLen, slot, fs)
    If twoLines Then h = 2 * size * 1.2 + 6 Else h = size + 8
End Sub

Private Sub LegendRow(ByVal fs As Double)
    Dim x As Double, s As Long, w As Double
    For s = 1 To gM
        w = TextW(gSerNames(s), fs)
        If x + 12 + w > gW And x > 0 Then Exit For
        AddRect x, (fs + 10 - 8) / 2 - 2, 8, 8, SeriesColor(s)
        AddText x + 12, 0, w + 4, fs + 6, gSerNames(s), fs, BLACK, "l"
        x = x + 12 + w + 14
    Next
End Sub

Private Function ArrowRange(o As ChartOpts, ByVal n As Long, ByRef aFrom As Long, ByRef aTo As Long) As Boolean
    Dim t As Long
    If o.arrowMode = "" Or o.arrowMode = "none" Or n < 2 Then Exit Function
    aFrom = o.arrowFrom
    If aFrom < 1 Then aFrom = 1
    If aFrom > n Then aFrom = n
    If o.arrowTo < 1 Or o.arrowTo > n Then aTo = n Else aTo = o.arrowTo
    If aFrom > aTo Then
        t = aFrom
        aFrom = aTo
        aTo = t
    End If
    ArrowRange = (aFrom <> aTo)
End Function

Private Function ArrowLabel(ByVal mode As String, ByVal a As Double, ByVal b As Double, ByVal periods As Long, _
        o As ChartOpts, ByVal dec As Long) As String
    If mode = "diff" Then
        If b - a > 0 Then
            ArrowLabel = "+" & Fmt(b - a, o.prefix, o.suffix, dec)
        Else
            ArrowLabel = Fmt(b - a, o.prefix, o.suffix, dec)
        End If
        Exit Function
    End If
    If a <= 0 Or b < 0 Then
        ArrowLabel = "n/a"
        Exit Function
    End If
    If mode = "pct" Then
        ArrowLabel = Signed((b / a - 1) * 100, 1, "%")
        Exit Function
    End If
    If periods <= 0 Then
        ArrowLabel = "n/a"
        Exit Function
    End If
    ArrowLabel = Signed(((b / a) ^ (1 / periods) - 1) * 100, 1, "%") & " p.a."
End Function

' Elbow arrow between two columns with a label bubble on its crossbar.
Private Sub ArrowPrims(ByVal xa As Double, ByVal xb As Double, ByVal topA As Double, ByVal topB As Double, _
        ByVal yLine As Double, ByVal lbl As String, ByVal fs As Double)
    Dim w As Double, h As Double, cx As Double
    AddLn xa, yLine, xa, topA, BLACK, 0.75
    AddLn xa, yLine, xb, yLine, BLACK, 0.75
    AddLn xb, yLine, xb, topB - 6, BLACK, 0.75
    AddTriDown xb - 3.5, topB - 6, 7, 6, BLACK
    w = TextW(lbl, fs) + 14
    h = fs + 6
    cx = (xa + xb) / 2
    AddRect cx - w / 2, yLine - h / 2, w, h, WHITE, BLACK, 0.75, True
    AddText cx - w / 2, yLine - h / 2, w, h, lbl, fs
End Sub

' ========================================================================
' Layouts
' ========================================================================

Private Sub LayoutColumns(o As ChartOpts)
    Dim n As Long, m As Long, i As Long, s As Long, fs As Double, kind As String
    Dim posSum() As Double, negSum() As Double, tot() As Double
    Dim dec As Long, legendH As Double, arrowH As Double, topH As Double, topLabels As Boolean
    Dim aFrom As Long, aTo As Long, hasArrow As Boolean
    Dim slot As Double, catSize As Double, catH As Double, barW As Double, x0 As Double
    Dim colTop() As Double, strokeC As String, strokeW As Double
    Dim x As Double, yt As Double, yb As Double, v As Double, vv As Double, lbl As String
    Dim bw As Double, hi As Double, sz As Double, up As Double, down As Double
    Dim a As Double, b As Double, hgt As Double, fillC As String

    n = gN
    m = gM
    fs = o.fontSize
    kind = o.chartType

    ReDim posSum(1 To n)
    ReDim negSum(1 To n)
    ReDim tot(1 To n)
    For i = 1 To n
        For s = 1 To m
            If gV(s, i) > 0 Then posSum(i) = posSum(i) + gV(s, i) Else negSum(i) = negSum(i) + gV(s, i)
        Next
        tot(i) = posSum(i) + negSum(i)
    Next
    dec = Decimals(o)

    If o.legend And m > 1 Then legendH = fs + 10
    hasArrow = ArrowRange(o, n, aFrom, aTo)
    If hasArrow Then arrowH = fs + 16
    If kind = "clustered" Then
        topLabels = o.labels
    Else
        topLabels = o.totals And (m > 1 Or kind = "percent")
    End If
    If topLabels Then topH = fs + 4

    slot = gW / n
    CategoryLabels slot, fs, catSize, catH
    gTop = legendH + arrowH + topH
    gBottom = gH - catH
    If gBottom - gTop < 1 Then gBottom = gTop + 1

    gMax = 0
    gMin = 0
    If kind = "percent" Then
        gMax = 1
    ElseIf kind = "clustered" Then
        For i = 1 To n
            For s = 1 To m
                gMax = Mx(gMax, gV(s, i))
                gMin = Mn(gMin, gV(s, i))
            Next
        Next
    Else
        For i = 1 To n
            gMax = Mx(gMax, posSum(i))
            gMin = Mn(gMin, negSum(i))
        Next
    End If
    If gMax = gMin Then gMax = gMin + 1

    barW = slot * 0.65
    ReDim colTop(1 To n)
    If m > 1 Then
        strokeC = WHITE
        strokeW = 0.75
    End If

    For i = 1 To n
        x0 = (i - 1) * slot + (slot - barW) / 2
        If kind = "clustered" Then
            bw = barW / m
            hi = YV(0)
            For s = 1 To m
                v = gV(s, i)
                x = x0 + (s - 1) * bw
                yt = YV(Mx(0, v))
                yb = YV(Mn(0, v))
                If yb - yt > 0 Then AddRect x, yt, bw, yb - yt, SeriesColor(s), strokeC, strokeW
                If o.labels And gHas(s, i) Then
                    lbl = Fmt(v, o.prefix, o.suffix, dec)
                    sz = FitSize(Len(lbl), bw + 8, fs)
                    If v >= 0 Then
                        AddText x - 4, yt - fs - 3, bw + 8, fs + 2, lbl, sz
                    Else
                        AddText x - 4, yb + 1, bw + 8, fs + 2, lbl, sz
                    End If
                End If
                If o.labels And v >= 0 Then hi = Mn(hi, yt - fs - 3) Else hi = Mn(hi, yt)
            Next
            colTop(i) = hi
        Else
            up = 0
            down = 0
            For s = 1 To m
                If kind = "percent" Then
                    If posSum(i) > 0 Then vv = Mx(0, gV(s, i)) / posSum(i) Else vv = 0
                Else
                    vv = gV(s, i)
                End If
                If vv >= 0 Then
                    a = up
                    b = up + vv
                    up = b
                Else
                    a = down + vv
                    b = down
                    down = a
                End If
                yt = YV(b)
                hgt = YV(a) - yt
                If hgt > 0 Then
                    fillC = SeriesColor(s)
                    AddRect x0, yt, barW, hgt, fillC, strokeC, strokeW
                    If o.labels Then
                        If kind = "percent" Then
                            lbl = Fmt(vv * 100, "", "%", 0)
                        Else
                            lbl = Fmt(gV(s, i), o.prefix, o.suffix, dec)
                        End If
                        If hgt >= fs + 2 And TextW(lbl, fs) <= barW - 2 Then
                            AddText x0, yt, barW, hgt, lbl, fs, ContrastText(fillC)
                        End If
                    End If
                End If
            Next
            If kind = "percent" Then yt = YV(1) Else yt = YV(up)
            If topLabels Then
                If kind = "percent" Then
                    lbl = Fmt(posSum(i), o.prefix, o.suffix, dec)
                Else
                    lbl = Fmt(tot(i), o.prefix, o.suffix, dec)
                End If
                AddText x0 - 6, yt - fs - 4, barW + 12, fs + 2, lbl, fs, BLACK, "c", True
            End If
            colTop(i) = yt - topH
        End If
    Next

    AddLn 0, YV(0), gW, YV(0), BLACK, 0.75
    For i = 1 To n
        AddText slot * (i - 1), gBottom + 3, slot, catH - 4, gCatText(i), catSize, GRAY1
    Next
    If legendH > 0 Then LegendRow fs

    If hasArrow Then
        Dim va As Double, vb As Double, xa As Double, xb As Double
        If kind = "clustered" Then
            va = gV(1, aFrom)
            vb = gV(1, aTo)
        ElseIf kind = "percent" Then
            va = posSum(aFrom)
            vb = posSum(aTo)
        Else
            va = tot(aFrom)
            vb = tot(aTo)
        End If
        lbl = ArrowLabel(o.arrowMode, va, vb, aTo - aFrom, o, dec)
        xa = (aFrom - 0.5) * slot
        xb = (aTo - 0.5) * slot
        ArrowPrims xa, xb, colTop(aFrom) - 2, colTop(aTo) - 2, legendH + arrowH / 2, lbl, fs
    End If
End Sub

Private Sub LayoutWaterfall(o As ChartOpts)
    Dim n As Long, i As Long, fs As Double, running As Double, d As Double
    Dim fromV() As Double, toV() As Double, delta() As Double, kindB() As String
    Dim hasBelow As Boolean, slot As Double, catSize As Double, catH As Double
    Dim barW As Double, x0 As Double, yt As Double, hgt As Double, colr As String
    Dim lbl As String, sz As Double, ly As Double, above As Boolean, dec As Long

    n = gN
    fs = o.fontSize
    ReDim fromV(1 To n)
    ReDim toV(1 To n)
    ReDim delta(1 To n)
    ReDim kindB(1 To n)
    For i = 1 To n
        If gIsE(i) Then
            fromV(i) = 0
            toV(i) = running
            kindB(i) = "total"
        ElseIf i = 1 Then
            running = gV(1, 1)
            fromV(i) = 0
            toV(i) = running
            kindB(i) = "total"
        Else
            d = gV(1, i)
            fromV(i) = running
            toV(i) = running + d
            delta(i) = d
            If d >= 0 Then kindB(i) = "up" Else kindB(i) = "down"
            running = running + d
        End If
    Next
    dec = Decimals(o)

    gMax = 0
    gMin = 0
    For i = 1 To n
        gMax = Mx(gMax, Mx(fromV(i), toV(i)))
        gMin = Mn(gMin, Mn(fromV(i), toV(i)))
        If kindB(i) = "down" Or (kindB(i) = "total" And toV(i) < 0) Then hasBelow = True
    Next
    If gMax = gMin Then gMax = gMin + 1

    slot = gW / n
    CategoryLabels slot, fs, catSize, catH
    If o.labels Then gTop = fs + 4 Else gTop = 2
    gBottom = gH - catH
    If o.labels And hasBelow Then gBottom = gBottom - fs - 4
    If gBottom - gTop < 1 Then gBottom = gTop + 1
    barW = slot * 0.65

    For i = 1 To n
        x0 = (i - 1) * slot + (slot - barW) / 2
        yt = YV(Mx(fromV(i), toV(i)))
        hgt = Mx(0.75, YV(Mn(fromV(i), toV(i))) - yt)
        Select Case kindB(i)
            Case "total": colr = BLUE500
            Case "up": colr = BLUE300
            Case Else: colr = GRAY2
        End Select
        AddRect x0, yt, barW, hgt, colr
        If o.labels Then
            If kindB(i) = "total" Then
                lbl = Fmt(toV(i), o.prefix, o.suffix, dec)
            ElseIf delta(i) > 0 Then
                lbl = "+" & Fmt(delta(i), o.prefix, o.suffix, dec)
            Else
                lbl = Fmt(delta(i), o.prefix, o.suffix, dec)
            End If
            above = (kindB(i) = "up") Or (kindB(i) = "total" And toV(i) >= 0)
            sz = FitSize(Len(lbl), slot, fs)
            If above Then ly = yt - fs - 3 Else ly = yt + hgt + 1
            AddText x0 - 6, ly, barW + 12, fs + 2, lbl, sz, BLACK, "c", (kindB(i) = "total")
        End If
        If i < n Then AddLn x0 + barW, YV(toV(i)), x0 + slot, YV(toV(i)), GRAY1, 0.5, True
    Next

    AddLn 0, YV(0), gW, YV(0), BLACK, 0.75
    For i = 1 To n
        AddText slot * (i - 1), gH - catH + 3, slot, catH - 4, gCatText(i), catSize, GRAY1
    Next
End Sub

Private Sub LayoutMekko(o As ChartOpts)
    Dim n As Long, m As Long, i As Long, s As Long, fs As Double, dec As Long
    Dim totalW As Double, legendW As Double, plotW As Double, topH As Double
    Dim plotBottom As Double, plotH As Double, x As Double, w As Double
    Dim colSum As Double, acc As Double, share As Double, yb As Double, h As Double
    Dim fillC As String, lbl As String, maxName As Double
    Dim lastMid() As Double, midv() As Double

    n = gN
    m = gM
    fs = o.fontSize
    For i = 1 To n
        totalW = totalW + Mx(0, gWidths(i))
    Next
    If totalW <= 0 Then Err.Raise vbObjectError + 20, , "The first data row (column widths) needs positive numbers."
    dec = Decimals(o)

    If o.legend Then
        For s = 1 To m
            maxName = Mx(maxName, TextW(gSerNames(s), fs))
        Next
        legendW = Mn(gW * 0.3, maxName + 8)
    End If
    plotW = gW - legendW
    If o.totals Then topH = fs + 4
    plotBottom = gH - (fs + 8)
    plotH = Mx(1, plotBottom - topH)

    ReDim lastMid(1 To m)
    ReDim midv(1 To m)
    For s = 1 To m
        lastMid(s) = -1
    Next
    For i = 1 To n
        w = Mx(0, gWidths(i)) / totalW * plotW
        colSum = 0
        For s = 1 To m
            colSum = colSum + Mx(0, gV(s, i))
        Next
        acc = 0
        For s = 1 To m
            If colSum > 0 Then share = Mx(0, gV(s, i)) / colSum Else share = 0
            yb = plotBottom - acc * plotH
            h = share * plotH
            acc = acc + share
            If h > 0 Then midv(s) = yb - h / 2 Else midv(s) = -1
            If h > 0 And w > 0 Then
                fillC = SeriesColor(s)
                AddRect x, yb - h, w, h, fillC, WHITE, 1
                lbl = Fmt(share * 100, "", "%", 0)
                If o.labels And h >= fs + 2 And TextW(lbl, fs) <= w - 2 Then
                    AddText x, yb - h, w, h, lbl, fs, ContrastText(fillC)
                End If
            End If
        Next
        If w > 0 Then
            If o.totals Then
                lbl = Fmt(gWidths(i), o.prefix, o.suffix, dec)
                AddText x, 0, w, fs + 2, lbl, FitSize(Len(lbl), w, fs), BLACK, "c", True
            End If
            AddText x, plotBottom + 3, w, fs + 4, gCats(i), FitSize(Len(gCats(i)), w, fs), GRAY1
            For s = 1 To m
                lastMid(s) = midv(s)
            Next
        End If
        x = x + w
    Next

    AddLn 0, plotBottom, plotW, plotBottom, BLACK, 0.75
    If o.legend Then
        For s = 1 To m
            If lastMid(s) >= 0 Then
                AddText plotW + 6, lastMid(s) - (fs + 2) / 2, legendW - 6, fs + 2, gSerNames(s), fs, BLACK, "l"
            End If
        Next
    End If
End Sub

Private Function NiceStep(ByVal rng As Double, ByVal target As Double) As Double
    Dim raw As Double, mag As Double, norm As Double, f As Double
    raw = rng / target
    mag = 10 ^ Int(Log(raw) / Log(10))
    norm = raw / mag
    If norm < 1.5 Then
        f = 1
    ElseIf norm < 3 Then
        f = 2
    ElseIf norm < 7 Then
        f = 5
    Else
        f = 10
    End If
    NiceStep = f * mag
End Function

Private Sub AddTick(ByVal t As Double, ByVal lbl As String)
    gNT = gNT + 1
    ReDim Preserve gTickT(1 To gNT)
    ReDim Preserve gTickL(1 To gNT)
    gTickT(gNT) = t
    gTickL(gNT) = lbl
End Sub

' Calendar periods covering [t0, t1]: months, quarters or years by span.
Private Sub MonthTicks(ByVal t0 As Double, ByVal t1 As Double)
    Dim span As Long, stp As Long, yr As Long, mo As Long, t As Double, lbl As String, names As Variant
    names = Array("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")
    span = (Year(t1) - Year(t0)) * 12 + Month(t1) - Month(t0) + 1
    If span <= 14 Then
        stp = 1
    ElseIf span <= 40 Then
        stp = 3
    Else
        stp = 12
    End If
    yr = Year(t0)
    mo = Month(t0)
    If stp = 12 Then mo = 1 Else mo = mo - ((mo - 1) Mod stp)
    gNT = 0
    Do
        t = CDbl(DateSerial(yr, mo, 1))
        If t > t1 And gNT > 0 Then Exit Do
        If stp = 12 Then
            lbl = CStr(yr)
        ElseIf stp = 3 Then
            lbl = "Q" & ((mo - 1) \ 3 + 1) & " " & Right$(CStr(yr), 2)
        Else
            lbl = names(mo - 1) & " " & Right$(CStr(yr), 2)
        End If
        AddTick t, lbl
        mo = mo + stp
        Do While mo > 12
            mo = mo - 12
            yr = yr + 1
        Loop
    Loop
    gEndT = CDbl(DateSerial(yr, mo, 1))
End Sub

Private Sub NumberTicks(ByVal t0 As Double, ByVal t1 As Double)
    Dim stp As Double, t As Double
    stp = NiceStep(Mx(t1 - t0, 1), 6)
    t = Int(t0 / stp) * stp
    gNT = 0
    Do While t <= t1 Or gNT = 0
        AddTick t, CStr(Round(t, 6))
        t = t + stp
    Loop
    gEndT = t
End Sub

Private Function GX(ByVal t As Double) As Double
    If gTMax - gTMin <= 0 Then
        GX = gLabelW
    Else
        GX = gLabelW + (t - gTMin) / (gTMax - gTMin) * (gW - gLabelW)
    End If
End Function

Private Sub LayoutGantt(o As ChartOpts)
    Dim fs As Double, r As Long, k As Long, t0 As Double, t1 As Double
    Dim headerH As Double, rowH As Double, top As Double, sz As Double, d As Double
    Dim xa As Double, xb As Double, nxt As Double, maxName As Double

    fs = o.fontSize
    t0 = gStart(1)
    t1 = gStart(1)
    For r = 1 To gM
        t0 = Mn(t0, Mn(gStart(r), gEnd(r)))
        t1 = Mx(t1, Mx(gStart(r), gEnd(r)))
        maxName = Mx(maxName, TextW(gTaskName(r), fs))
    Next
    If gDates Then MonthTicks t0, t1 Else NumberTicks t0, t1
    gTMin = gTickT(1)
    gTMax = Mx(gEndT, t1)
    gLabelW = Mn(gW * 0.35, maxName + 10)
    headerH = fs + 10
    rowH = (gH - headerH) / gM

    For k = 1 To gNT
        If k < gNT Then nxt = gTickT(k + 1) Else nxt = gTMax
        xa = GX(gTickT(k))
        xb = GX(nxt)
        AddLn xa, headerH, xa, gH, GRAY3, 0.5
        AddText xa, 0, xb - xa, headerH - 2, gTickL(k), FitSize(Len(gTickL(k)), xb - xa, fs - 1), GRAY1
    Next
    AddLn gLabelW, headerH, gW, headerH, BLACK, 0.75

    For r = 1 To gM
        top = headerH + (r - 1) * rowH
        sz = Mn(fs, rowH * 0.7)
        AddText 0, top, gLabelW - 6, rowH, gTaskName(r), sz, BLACK, "l"
        If Not gHasEnd(r) Or gEnd(r) = gStart(r) Then
            d = Mn(rowH * 0.55, 12)
            AddDiamond GX(gStart(r)) - d / 2, top + (rowH - d) / 2, d, d, BLACK
        Else
            xa = GX(Mn(gStart(r), gEnd(r)))
            xb = GX(Mx(gStart(r), gEnd(r)))
            AddRect xa, top + rowH * 0.22, Mx(1, xb - xa), rowH * 0.56, BLUE500
        End If
        If r < gM Then AddLn 0, top + rowH, gW, top + rowH, GRAY3, 0.5
    Next
End Sub

' ========================================================================
' Storing charts and datasheets on slides
' ========================================================================

Private Function DefaultOpts(ByVal t As String) As ChartOpts
    Dim o As ChartOpts
    o.chartType = t
    o.width = 480
    o.height = 270
    o.labels = True
    o.totals = True
    o.legend = True
    o.decimals = -1
    o.arrowMode = "none"
    o.arrowFrom = 1
    o.arrowTo = 0
    o.fontSize = 10
    o.fontName = DEFAULT_FONT
    DefaultOpts = o
End Function

Private Sub SetLongTag(ByVal shp As Shape, ByVal name As String, ByVal value As String)
    Dim n As Long, i As Long
    n = (Len(value) + 199) \ 200
    shp.Tags.Add name & "N", CStr(n)
    For i = 0 To n - 1
        shp.Tags.Add name & i, Mid$(value, i * 200 + 1, 200)
    Next
End Sub

Private Function GetLongTag(ByVal shp As Shape, ByVal name As String) As String
    Dim n As Long, i As Long, s As String
    n = Val(shp.Tags(name & "N"))
    For i = 0 To n - 1
        s = s & shp.Tags(name & i)
    Next
    GetLongTag = s
End Function

Private Function BoolTag(ByVal v As Boolean) As String
    If v Then BoolTag = "1" Else BoolTag = "0"
End Function

Private Sub SaveChart(ByVal grp As Shape, g As Variant, o As ChartOpts)
    grp.Tags.Add "AFCHART", "1"
    grp.Tags.Add "AFTYPE", o.chartType
    grp.Tags.Add "AFLABELS", BoolTag(o.labels)
    grp.Tags.Add "AFTOTALS", BoolTag(o.totals)
    grp.Tags.Add "AFLEGEND", BoolTag(o.legend)
    grp.Tags.Add "AFDEC", CStr(o.decimals)
    grp.Tags.Add "AFPREFIX", "[" & o.prefix & "]"
    grp.Tags.Add "AFSUFFIX", "[" & o.suffix & "]"
    grp.Tags.Add "AFARROW", o.arrowMode
    grp.Tags.Add "AFAFROM", CStr(o.arrowFrom)
    grp.Tags.Add "AFATO", CStr(o.arrowTo)
    grp.Tags.Add "AFFS", CStr(o.fontSize)
    grp.Tags.Add "AFFONT", o.fontName
    SetLongTag grp, "AFGRID", SerializeGrid(g)
End Sub

' Prefix and suffix are stored in brackets so an empty value survives.
Private Function Unbracket(ByVal s As String) As String
    If Len(s) >= 2 And Left$(s, 1) = "[" And Right$(s, 1) = "]" Then Unbracket = Mid$(s, 2, Len(s) - 2)
End Function

Private Function LoadOpts(ByVal shp As Shape) As ChartOpts
    Dim o As ChartOpts
    o = DefaultOpts(shp.Tags("AFTYPE"))
    o.width = shp.Width
    o.height = shp.Height
    o.labels = (shp.Tags("AFLABELS") <> "0")
    o.totals = (shp.Tags("AFTOTALS") <> "0")
    o.legend = (shp.Tags("AFLEGEND") <> "0")
    If shp.Tags("AFDEC") <> "" Then o.decimals = Val(shp.Tags("AFDEC"))
    o.prefix = Unbracket(shp.Tags("AFPREFIX"))
    o.suffix = Unbracket(shp.Tags("AFSUFFIX"))
    If shp.Tags("AFARROW") <> "" Then o.arrowMode = shp.Tags("AFARROW")
    If shp.Tags("AFAFROM") <> "" Then o.arrowFrom = Val(shp.Tags("AFAFROM"))
    If shp.Tags("AFATO") <> "" Then o.arrowTo = Val(shp.Tags("AFATO"))
    If shp.Tags("AFFS") <> "" Then o.fontSize = Val(shp.Tags("AFFS"))
    If shp.Tags("AFFONT") <> "" Then o.fontName = shp.Tags("AFFONT")
    LoadOpts = o
End Function

Private Function LoadGrid(ByVal shp As Shape) As Variant
    LoadGrid = ParseGridText(GetLongTag(shp, "AFGRID"))
End Function

Private Function IsChart(ByVal shp As Shape) As Boolean
    IsChart = (shp.Tags("AFCHART") = "1")
End Function

Private Function FindShape(ByVal sld As Slide, ByVal name As String) As Shape
    If name = "" Then Exit Function
    On Error Resume Next
    Set FindShape = sld.Shapes(name)
    On Error GoTo 0
End Function

Private Function TableGrid(ByVal shp As Shape) As Variant
    Dim r As Long, c As Long, nR As Long, nC As Long, txt As String
    nR = shp.Table.Rows.Count
    nC = shp.Table.Columns.Count
    Dim g() As String
    ReDim g(1 To nR, 1 To nC)
    For r = 1 To nR
        For c = 1 To nC
            txt = shp.Table.Cell(r, c).Shape.TextFrame.TextRange.Text
            g(r, c) = Trim$(Replace(Replace(txt, vbCr, " "), vbLf, " "))
        Next
    Next
    TableGrid = TrimGrid(g)
End Function

Private Sub InsertSheet(ByVal sld As Slide, g As Variant, ByVal t As String, ByVal target As String, _
        ByVal l As Double, ByVal tp As Double)
    Dim nR As Long, nC As Long, r As Long, c As Long, w As Double, h As Double, tb As Shape
    nR = UBound(g, 1)
    nC = UBound(g, 2)
    w = Mn(ActivePresentation.PageSetup.SlideWidth - 40, nC * 80)
    h = nR * 22
    l = Mx(10, Mn(l, ActivePresentation.PageSetup.SlideWidth - w - 10))
    tp = Mx(10, Mn(tp, ActivePresentation.PageSetup.SlideHeight - h - 10))
    Set tb = sld.Shapes.AddTable(nR, nC, l, tp, w, h)
    For r = 1 To nR
        For c = 1 To nC
            With tb.Table.Cell(r, c).Shape.TextFrame.TextRange
                .Text = GCell(g, r, c)
                .Font.Size = 11
            End With
        Next
    Next
    tb.Name = "ChartForge data " & Format$(Now, "hhnnss")
    tb.Tags.Add "AFSHEET", t
    If target <> "" Then tb.Tags.Add "AFTARGET", target
    On Error Resume Next
    tb.Select
End Sub

' Draws a chart and returns its group. If anything fails, the shapes drawn
' so far are removed and the error is passed on.
Private Function DrawChart(ByVal sld As Slide, g As Variant, o As ChartOpts, ByVal l As Double, ByVal tp As Double) As Shape
    Dim i As Long, frameShp As Shape, names() As Variant, grp As Shape, msg As String
    Randomize
    Set gSld = sld
    gOX = l
    gOY = tp
    gW = o.width
    gH = o.height
    gFont = o.fontName
    gCount = 0
    gPrefix = "AF" & Format$(Now, "hhnnss") & Int(Rnd * 100000) & "_"

    On Error GoTo Cleanup
    Set frameShp = sld.Shapes.AddShape(msoShapeRectangle, l, tp, gW, gH)
    StyleShape frameShp, "", "", 0
    Track frameShp

    ParseFor o.chartType, g
    Select Case o.chartType
        Case "waterfall": LayoutWaterfall o
        Case "mekko": LayoutMekko o
        Case "gantt": LayoutGantt o
        Case Else: LayoutColumns o
    End Select

    ReDim names(0 To gCount - 1)
    For i = 1 To gCount
        names(i - 1) = gNames(i)
    Next
    Set grp = sld.Shapes.Range(names).Group
    grp.Name = "ChartForge chart " & gPrefix
    SaveChart grp, g, o
    Set DrawChart = grp
    Exit Function

Cleanup:
    msg = Err.Description
    On Error Resume Next
    For i = 1 To gCount
        sld.Shapes(gNames(i)).Delete
    Next
    On Error GoTo 0
    Err.Raise vbObjectError + 30, , msg
End Function

' Draws the new version over the old chart (same place and size, so a
' resize sticks), then removes the old one.
Private Sub RedrawChart(ByVal old As Shape, g As Variant, o As ChartOpts)
    Dim grp As Shape
    o.width = old.Width
    o.height = old.Height
    Set grp = DrawChart(old.Parent, g, o, old.Left, old.Top)
    old.Delete
    SelectShape grp
End Sub

Private Sub SelectShape(ByVal shp As Shape)
    On Error Resume Next
    shp.Select
End Sub

Private Function SelectedShape() As Shape
    Dim sr As ShapeRange
    On Error GoTo NoSel
    If ActiveWindow.Selection.Type = ppSelectionNone Then Exit Function
    If ActiveWindow.Selection.Type = ppSelectionSlides Then Exit Function
    Set sr = ActiveWindow.Selection.ShapeRange
    If sr.Count <> 1 Then Exit Function
    If sr.Child = msoTrue Then
        Set SelectedShape = sr.ParentGroup
    Else
        Set SelectedShape = sr(1)
    End If
NoSel:
End Function

Private Function CurrentSlide() As Slide
    Dim s As Slide
    On Error Resume Next
    Set s = ActiveWindow.View.Slide
    If s Is Nothing Then Set s = ActiveWindow.Selection.SlideRange(1)
    On Error GoTo 0
    If s Is Nothing Then Err.Raise vbObjectError + 40, , "Open a slide in Normal view first."
    Set CurrentSlide = s
End Function

' Builds a chart from a selected datasheet table. If the table came from
' Edit data, it replaces that chart; otherwise the chart goes where the
' table was.
Private Sub BuildFromSheet(ByVal sheet As Shape, ByVal t As String)
    Dim g As Variant, sld As Slide, old As Shape, o As ChartOpts, grp As Shape
    g = TableGrid(sheet)
    Set sld = sheet.Parent
    Set old = FindShape(sld, sheet.Tags("AFTARGET"))
    If old Is Nothing Then
        o = DefaultOpts(t)
        Set grp = DrawChart(sld, g, o, sheet.Left, sheet.Top)
        sheet.Delete
        SelectShape grp
    Else
        o = LoadOpts(old)
        o.chartType = t
        RedrawChart old, g, o
        sheet.Delete
    End If
End Sub

Private Sub RunType(ByVal t As String)
    Dim shp As Shape, o As ChartOpts
    On Error GoTo Fail
    Set shp = SelectedShape()
    If Not shp Is Nothing Then
        If shp.HasTable Then
            BuildFromSheet shp, t
            Exit Sub
        End If
        If IsChart(shp) Then
            o = LoadOpts(shp)
            If GridKey(o.chartType) = GridKey(t) Then
                o.chartType = t
                RedrawChart shp, LoadGrid(shp), o
                Exit Sub
            End If
        End If
    End If
    InsertSheet CurrentSlide(), SampleGrid(t), t, "", 60, 120
    Hint "sheet", "This table is the chart's datasheet. Replace the sample data (you can paste from Excel), " & _
        "keep the table selected, then click " & TypeLabel(t) & " again to draw the chart."
    Exit Sub
Fail:
    MsgBox Err.Description, vbExclamation, "ChartForge"
End Sub

Private Function SelectedChart() As Shape
    Dim shp As Shape
    Set shp = SelectedShape()
    If Not shp Is Nothing Then
        If IsChart(shp) Then Set SelectedChart = shp
    End If
End Function

Private Sub ChangeOpt(ByVal what As String)
    Dim shp As Shape, o As ChartOpts, s As String, parts As Variant
    On Error GoTo Fail
    Set shp = SelectedChart()
    If shp Is Nothing Then
        MsgBox "Select a ChartForge chart first.", vbInformation, "ChartForge"
        Exit Sub
    End If
    o = LoadOpts(shp)
    Select Case what
        Case "cagr", "pct", "diff", "none", "range"
            If GridKey(o.chartType) <> "column" Then
                MsgBox "Growth arrows work on Stacked, Clustered and 100% charts.", vbInformation, "ChartForge"
                Exit Sub
            End If
            If what = "range" Then
                s = InputBox("Draw the arrow between which columns? Enter two column numbers, e.g. 1,5 (or 1,last).", _
                    "Arrow range", CStr(o.arrowFrom) & "," & IIf(o.arrowTo < 1, "last", CStr(o.arrowTo)))
                If Trim$(s) = "" Then Exit Sub
                parts = Split(s, ",")
                If UBound(parts) < 1 Then Exit Sub
                o.arrowFrom = Val(parts(0))
                If LCase$(Trim$(parts(1))) = "last" Then o.arrowTo = 0 Else o.arrowTo = Val(parts(1))
                If o.arrowMode = "none" Then o.arrowMode = "cagr"
            Else
                o.arrowMode = what
            End If
        Case "labels": o.labels = Not o.labels
        Case "totals": o.totals = Not o.totals
        Case "legend": o.legend = Not o.legend
        Case "decmore"
            If o.decimals < 0 Then o.decimals = 1 Else o.decimals = Mn(4, o.decimals + 1)
        Case "decless"
            If o.decimals <= 0 Then o.decimals = 0 Else o.decimals = o.decimals - 1
        Case "format"
            s = InputBox("Text shown before every number (e.g. $). Leave blank for none.", "Number prefix", o.prefix)
            o.prefix = Trim$(s)
            s = InputBox("Text shown after every number (e.g. bn).", "Number suffix", o.suffix)
            o.suffix = Trim$(s)
        Case "font"
            If o.fontName = DEFAULT_FONT Then o.fontName = "Calibri" Else o.fontName = DEFAULT_FONT
    End Select
    RedrawChart shp, LoadGrid(shp), o
    Exit Sub
Fail:
    MsgBox Err.Description, vbExclamation, "ChartForge"
End Sub

' ========================================================================
' Ribbon callbacks
' ========================================================================

Public Sub AF_Stacked(control As IRibbonControl)
    RunType "stacked"
End Sub

Public Sub AF_Clustered(control As IRibbonControl)
    RunType "clustered"
End Sub

Public Sub AF_Percent(control As IRibbonControl)
    RunType "percent"
End Sub

Public Sub AF_Waterfall(control As IRibbonControl)
    RunType "waterfall"
End Sub

Public Sub AF_Mekko(control As IRibbonControl)
    RunType "mekko"
End Sub

Public Sub AF_Gantt(control As IRibbonControl)
    RunType "gantt"
End Sub

Public Sub AF_EditData(control As IRibbonControl)
    Dim shp As Shape, o As ChartOpts
    On Error GoTo Fail
    Set shp = SelectedChart()
    If shp Is Nothing Then
        MsgBox "Select a ChartForge chart first.", vbInformation, "ChartForge"
        Exit Sub
    End If
    o = LoadOpts(shp)
    InsertSheet shp.Parent, LoadGrid(shp), o.chartType, shp.Name, shp.Left, shp.Top
    Hint "edit", "Edit the table, then click Update (or a chart type) with the table selected to apply your changes."
    Exit Sub
Fail:
    MsgBox Err.Description, vbExclamation, "ChartForge"
End Sub

Public Sub AF_Update(control As IRibbonControl)
    Dim shp As Shape, t As String, o As ChartOpts
    On Error GoTo Fail
    Set shp = SelectedShape()
    If shp Is Nothing Then GoTo NothingSelected
    If shp.HasTable Then
        t = shp.Tags("AFSHEET")
        If t = "" Then t = "stacked"
        BuildFromSheet shp, t
        Exit Sub
    End If
    If IsChart(shp) Then
        o = LoadOpts(shp)
        RedrawChart shp, LoadGrid(shp), o
        Exit Sub
    End If
NothingSelected:
    MsgBox "Select a ChartForge chart or its data table first.", vbInformation, "ChartForge"
    Exit Sub
Fail:
    MsgBox Err.Description, vbExclamation, "ChartForge"
End Sub

Public Sub AF_Cagr(control As IRibbonControl)
    ChangeOpt "cagr"
End Sub

Public Sub AF_PctChange(control As IRibbonControl)
    ChangeOpt "pct"
End Sub

Public Sub AF_Difference(control As IRibbonControl)
    ChangeOpt "diff"
End Sub

Public Sub AF_NoArrow(control As IRibbonControl)
    ChangeOpt "none"
End Sub

Public Sub AF_ArrowRange(control As IRibbonControl)
    ChangeOpt "range"
End Sub

Public Sub AF_Labels(control As IRibbonControl)
    ChangeOpt "labels"
End Sub

Public Sub AF_Totals(control As IRibbonControl)
    ChangeOpt "totals"
End Sub

Public Sub AF_Legend(control As IRibbonControl)
    ChangeOpt "legend"
End Sub

Public Sub AF_DecMore(control As IRibbonControl)
    ChangeOpt "decmore"
End Sub

Public Sub AF_DecLess(control As IRibbonControl)
    ChangeOpt "decless"
End Sub

Public Sub AF_NumberFormat(control As IRibbonControl)
    ChangeOpt "format"
End Sub

Public Sub AF_Font(control As IRibbonControl)
    ChangeOpt "font"
End Sub
