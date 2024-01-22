Function FileExists(FilePath As String) As Boolean

    'PURPOSE:   Simplification of the native Dir() function. Test whether or not a file exists at the path provided and return 'True' or 'False' value.

    If Dir(FilePath) = "" Then
        FileExists = False
    Else
        FileExists = True
    End If

End Function
