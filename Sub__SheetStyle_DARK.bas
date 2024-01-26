Sub SheetStyle_DARK()

    'PURPOSE:   Change style of current sheet to a DARK theme. Add gray background stored on local disk (Black, Text 1, Lighter 25%.png), add gray borders around all cells, change font to white. Option to set fill cells to gray color instead of adding gray background.

    Application.ScreenUpdating = False ' pause animations
    Starting_Selection_Address = Replace(Selection.Address, "$", "") ' store currently selected cell range as a string

    ActiveSheet.SetBackgroundPicture Filename:= _
        "C:\Users\cmiller.RTGTX\Proton Drive\hollow_submarine\My files\02 -- Images\Plain-Colors\From Excel\Black, Text 1, Lighter 25%.png"

    Cells.Select

'    With Selection.Interior
'        .Pattern = xlNone
'        .TintAndShade = 0
'        .PatternTintAndShade = 0
'    End With

    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -4.99893185216834E-02
    End With

    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
    Range(Starting_Selection_Address).Select ' return selection to the starting selection
    Application.ScreenUpdating = True ' resume animations

End Sub