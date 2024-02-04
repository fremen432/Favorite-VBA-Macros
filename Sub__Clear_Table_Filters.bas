Sub Clear_Table_Filters()

    'PURPOSE:   Clear all filters in a table. Active cell must be located inside a table. If the Active cell is not located inside a table, give an error.

    On Error Resume Next

    Dim tbl As ListObject
    Set tbl = ActiveCell.ListObject
    
    On Error GoTo ErrorHandler
    
    If tbl Is Nothing Then
        MsgBox "Active cell isn't inside a table", vbExclamation, "Error"
    Else
        If tbl.AutoFilter.FilterMode Then
            ActiveSheet.ShowAllData
        Else
            MsgBox "There are no filters to clear", vbInformation, "Info"
        End If
    End If
    
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical, "Error"
    
End Sub