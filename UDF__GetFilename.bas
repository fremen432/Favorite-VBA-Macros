Function GetFilename(filePath As String) As String
    ' =TRIM(RIGHT(SUBSTITUTE(A1, "\", REPT(" ", LEN(A1))), LEN(A1)))
    ' Created using ChatGPT
    
    Dim lastBackslashPos As Integer
    Dim filename As String
    
    ' Find the position of the last backslash
    lastBackslashPos = InStrRev(filePath, "\")
    
    ' Use MID function to extract the filename
    filename = Mid(filePath, lastBackslashPos + 1)
    
    ' Return the extracted filename
    GetFilename = filename
End function