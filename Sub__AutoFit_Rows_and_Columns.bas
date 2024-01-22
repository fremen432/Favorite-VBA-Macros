Sub AutoFit()
'
' AutoFit all rows and columns and set empty cells to default RowHeight and ColumnWidth
'
    Dim Starting_Selection_Address As String
    Dim Default_RowHeight As Integer
    Dim Default_ColumnWidth As Integer

    Starting_Selection_Address = Replace(Selection.Address, "$", "") ' store currently selected cell range as a string
    Default_RowHeight = 15
    Default_ColumnWidth = 8.43
    
    Application.ScreenUpdating = False ' prevent animations
    Cells.Select
    
    ' first, set all cells in worksheet to default RowHeight and ColumnWidth
    Selection.RowHeight = Default_RowHeight
    Selection.ColumnWidth = Default_ColumnWidth
    
    ' next, AutoFit rows and columns
    Selection.Rows.AutoFit
    Selection.Columns.AutoFit
    
    'return selection to the starting selection
    Range(Starting_Selection_Address).Select
    Application.ScreenUpdating = True
    
End Sub
