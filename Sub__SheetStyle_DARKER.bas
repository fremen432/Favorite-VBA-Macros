Sub SheetStyle_DARKER()

    'PURPOSE:   Change style of current sheet to a dark theme. Add black background stored on local disk (Black, Text 1.png), add dark-gray borders around all cells, change font to light gray. Option to set fill cells to black color instead of adding black background.

    Application.ScreenUpdating = False ' pause animations
    Starting_Selection_Address = Replace(Selection.Address, "$", "") ' store currently selected cell range as a string

    ActiveSheet.SetBackgroundPicture Filename:= _
        "C:\Users\cmiller.RTGTX\Proton Drive\hollow_submarine\My files\02 -- Images\Plain-Colors\From Excel\Black, Text 1.png"

    Cells.Select

'    With Selection.Interior
'        .Pattern = xlSolid
'        .PatternColorIndex = xlAutomatic
'        .ThemeColor = xlThemeColorLight1
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
        .ThemeColor = 2
        .TintAndShade = 0.149998474074526
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ThemeColor = 2
        .TintAndShade = 0.149998474074526
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ThemeColor = 2
        .TintAndShade = 0.149998474074526
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ThemeColor = 2
        .TintAndShade = 0.149998474074526
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ThemeColor = 2
        .TintAndShade = 0.149998474074526
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ThemeColor = 2
        .TintAndShade = 0.149998474074526
        .Weight = xlThin
    End With

    Range(Starting_Selection_Address).Select ' return selection to the starting selection
    Application.ScreenUpdating = True ' resume animations

End Sub