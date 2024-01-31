Sub SheetStyle_BLACK()

    'PURPOSE:   Change style of current sheet to a DARKER theme. Add black background stored on local disk (Black, Text 1.png), add dark-gray borders around all cells, change font to light gray. Option to set fill cells to black color instead of adding black background.

    Application.ScreenUpdating = False ' pause animations
    Starting_Selection_Address = Replace(Selection.Address, "$", "") ' store currently selected cell range as a string

    ActiveSheet.SetBackgroundPicture Filename:= "PATH_TO_BLACK_IMAGE" ' set path to plain black image. Ex: https://commons.wikimedia.org/wiki/File:A_black_image.jpg

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
