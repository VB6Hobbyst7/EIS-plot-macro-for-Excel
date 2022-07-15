Attribute VB_Name = "Module2"
Sub Setup()
    Dim xSh As Worksheet
    Application.ScreenUpdating = False
    For Each xSh In Worksheets
        xSh.Select
        Call PlotNyquis
    Next
    Application.ScreenUpdating = True
End Sub
Sub PlotNyquis()
Attribute PlotNyquis.VB_Description = "To perform Nyquis Plot for Hioki LCR meter data"
Attribute PlotNyquis.VB_ProcData.VB_Invoke_Func = "P\n14"
'
' PlotNyquis Macro
' To perform Nyquis Plot for Hioki LCR meter data
'
' Keyboard Shortcut: Ctrl+Shift+P
'
    Range("L2").Select
    ActiveCell.FormulaR1C1 = "=RC[-9]*COS(RADIANS(RC[-7]))"
    Range("M2").Select
    ActiveCell.FormulaR1C1 = "=-RC[-10]*SIN(RADIANS(RC[-8]))"
    Range("N2").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=RC[-13]"
    Range("O2").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=RC[-3]"
    Range("P2").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=-RC[-3]"
    Range("L2:P2").Select
    ActiveWindow.ScrollRow = 2
    ActiveWindow.ScrollRow = 3
    ActiveWindow.ScrollRow = 6
    ActiveWindow.ScrollRow = 9
    ActiveWindow.ScrollRow = 12
    ActiveWindow.ScrollRow = 17
    ActiveWindow.ScrollRow = 23
    ActiveWindow.ScrollRow = 30
    ActiveWindow.ScrollRow = 43
    ActiveWindow.ScrollRow = 50
    ActiveWindow.ScrollRow = 57
    ActiveWindow.ScrollRow = 63
    ActiveWindow.ScrollRow = 69
    ActiveWindow.ScrollRow = 75
    ActiveWindow.ScrollRow = 79
    ActiveWindow.ScrollRow = 84
    ActiveWindow.ScrollRow = 89
    ActiveWindow.ScrollRow = 93
    ActiveWindow.ScrollRow = 96
    ActiveWindow.ScrollRow = 97
    ActiveWindow.ScrollRow = 100
    ActiveWindow.ScrollRow = 101
    ActiveWindow.ScrollRow = 103
    ActiveWindow.ScrollRow = 104
    ActiveWindow.ScrollRow = 106
    ActiveWindow.ScrollRow = 108
    ActiveWindow.ScrollRow = 111
    ActiveWindow.ScrollRow = 114
    ActiveWindow.ScrollRow = 117
    ActiveWindow.ScrollRow = 121
    ActiveWindow.ScrollRow = 123
    ActiveWindow.ScrollRow = 126
    ActiveWindow.ScrollRow = 128
    ActiveWindow.ScrollRow = 130
    ActiveWindow.ScrollRow = 132
    ActiveWindow.ScrollRow = 135
    ActiveWindow.ScrollRow = 136
    ActiveWindow.ScrollRow = 137
    ActiveWindow.ScrollRow = 139
    ActiveWindow.ScrollRow = 140
    ActiveWindow.ScrollRow = 142
    ActiveWindow.ScrollRow = 143
    ActiveWindow.ScrollRow = 144
    ActiveWindow.ScrollRow = 145
    ActiveWindow.ScrollRow = 146
    ActiveWindow.ScrollRow = 147
    ActiveWindow.ScrollRow = 148
    ActiveWindow.ScrollRow = 150
    ActiveWindow.ScrollRow = 151
    ActiveWindow.ScrollRow = 152
    ActiveWindow.ScrollRow = 154
    ActiveWindow.ScrollRow = 156
    ActiveWindow.ScrollRow = 158
    Range("L2:P177").Select
    Selection.FillDown
    Range("L177").Select
    ActiveWindow.ScrollRow = 157
    ActiveWindow.ScrollRow = 156
    ActiveWindow.ScrollRow = 154
    ActiveWindow.ScrollRow = 150
    ActiveWindow.ScrollRow = 146
    ActiveWindow.ScrollRow = 139
    ActiveWindow.ScrollRow = 132
    ActiveWindow.ScrollRow = 125
    ActiveWindow.ScrollRow = 117
    ActiveWindow.ScrollRow = 110
    ActiveWindow.ScrollRow = 101
    ActiveWindow.ScrollRow = 84
    ActiveWindow.ScrollRow = 76
    ActiveWindow.ScrollRow = 68
    ActiveWindow.ScrollRow = 61
    ActiveWindow.ScrollRow = 54
    ActiveWindow.ScrollRow = 40
    ActiveWindow.ScrollRow = 33
    ActiveWindow.ScrollRow = 27
    ActiveWindow.ScrollRow = 22
    ActiveWindow.ScrollRow = 18
    ActiveWindow.ScrollRow = 15
    ActiveWindow.ScrollRow = 11
    ActiveWindow.ScrollRow = 10
    ActiveWindow.ScrollRow = 8
    ActiveWindow.ScrollRow = 7
    ActiveWindow.ScrollRow = 5
    ActiveWindow.ScrollRow = 4
    ActiveWindow.ScrollRow = 3
    ActiveWindow.ScrollRow = 2
    ActiveWindow.ScrollRow = 1
    Range("L2:M177").Select
    Range("L177").Activate
    ActiveSheet.Shapes.AddChart2(240, xlXYScatter).Select
    ActiveChart.SetSourceData Source:=Range("$L$2:$M$177")
    ActiveChart.ClearToMatchStyle
    ActiveChart.ChartStyle = 245
    ActiveChart.ChartTitle.Select
    ActiveChart.ChartTitle.Text = "Nyquis plot"
    Selection.Format.TextFrame2.TextRange.Characters.Text = "Nyquis plot"
    With Selection.Format.TextFrame2.TextRange.Characters(1, 11).ParagraphFormat
        .TextDirection = msoTextDirectionLeftToRight
        .Alignment = msoAlignCenter
    End With
    With Selection.Format.TextFrame2.TextRange.Characters(1, 11).Font
        .BaselineOffset = 0
        .Bold = msoTrue
        .Caps = msoNoCaps
        .NameComplexScript = "+mn-cs"
        .NameFarEast = "+mn-ea"
        .Fill.Visible = msoTrue
        .Fill.ForeColor.RGB = RGB(217, 217, 217)
        .Fill.Transparency = 0
        .Fill.Solid
        .Size = 14
        .Italic = msoFalse
        .Kerning = 12
        .Name = "+mn-lt"
        .UnderlineStyle = msoNoUnderline
        .Strike = msoNoStrike
    End With
    ActiveChart.ChartArea.Select
End Sub
