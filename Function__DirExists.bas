Function DirExists(Path) As Boolean

    'PURPOSE:   Simplification of the native Dir() function. Test whether or not a directory exists at the path provided and return 'True' or 'False' value.

    If Dir(Path, vbDirectory) = "" Then
        DirExists = False
    Else
        DirExists = True
    End If

End Function
