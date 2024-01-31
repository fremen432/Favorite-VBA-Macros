Function GetFileExtension(filePath As String) As String
    ' =IFERROR(TRIM(RIGHT(SUBSTITUTE(A1, "\", REPT(" ", LEN(A1))), LEN(A1)-MAX(FIND(".",SUBSTITUTE(A1,"\",""))))), "")
    ' Created using ChatGPT

    Dim lastDotPos As Integer
    Dim extension As String
    
    ' Find the position of the last dot (.) in the filename
    lastDotPos = InStrRev(filePath, ".")
    
    ' Check if a dot was found and it is after the last backslash
    If lastDotPos > InStrRev(filePath, "\") Then
        ' Use MID function to extract the file extension
        extension = Mid(filePath, lastDotPos + 1)
    Else
        ' No dot found or it's before the last backslash, set extension to an empty string
        extension = ""
    End If
    
    ' Return the extracted file extension
    GetFileExtension = extension
End Function