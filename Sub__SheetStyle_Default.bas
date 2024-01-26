Sub SheetStyle_Default()

    'PURPOSE:   Change style of current sheet to the default light theme. Delete background, remove borders all cells, change font to default. Option to set fill color of all cells to no fill.

    Application.ScreenUpdating = False ' pause animations
    Starting_Selection_Address = Replace(Selection.Address, "$", "") ' store currently selected cell range as a string

    ActiveSheet.SetBackgroundPicture Filename:=""
    
    Cells.Select

'    With Selection.Interior
'        .Pattern = xlNone
'        .TintAndShade = 0
'        .PatternTintAndShade = 0
'    End With

    With Selection.Font
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
    End With

    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone

    Range(Starting_Selection_Address).Select ' return selection to the starting selection
    Application.ScreenUpdating = True ' resume animations

End Sub