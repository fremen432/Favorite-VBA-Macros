Sub AutoFit()

    'PURPOSE:   AutoFit all rows and columns. Set empty cells to default RowHeight and ColumnWidth

    Dim Starting_Selection_Address As String
    Dim Default_RowHeight As Integer
    Dim Default_ColumnWidth As Double

    Starting_Selection_Address = Replace(Selection.Address, "$", "") ' store currently selected cell range as a string
    Default_RowHeight = 15
    Default_ColumnWidth = 8.43
    
    Application.ScreenUpdating = False ' pause animations
    
    Cells.Select ' select all cells in current sheet
    
    ' first, set all cells in worksheet to default RowHeight and ColumnWidth
    Selection.RowHeight = Default_RowHeight
    Selection.ColumnWidth = Default_ColumnWidth
    
    ' next, AutoFit rows and columns
    Selection.Rows.AutoFit
    Selection.Columns.AutoFit
    
    'return selection to the starting selection
    Range(Starting_Selection_Address).Select

    Application.ScreenUpdating = True ' resume animations
    
End Sub
