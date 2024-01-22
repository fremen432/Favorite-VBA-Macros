Function SuperTrim(userString As String) As String

    'PURPOSE:   Enhancement of the native TRIM() function. Trim everything before and after text that isn't an alphanumeric character. This includes carriage returns.
    'REFERENCE: https://software-solutions-online.com/vba-regex-guide/#:~:text=To%20start%20using%20RegEx%20in,Expressions%205.5%2C%20then%20click%20OK
    'NOTE:      \W matches any non-alphanumeric characters and the underscore
    
    Dim pattern_trim_LEFT As String
    Dim pattern_trim_RIGHT As String
    
    Dim regexObject_trim_LEFT As RegExp
    Dim regexObject_trim_RIGHT As RegExp
    
    Set regexObject_trim_LEFT = New RegExp
    Set regexObject_trim_RIGHT = New RegExp
    
    pattern_trim_LEFT = "^[\W]+"
    pattern_trim_RIGHT = "[\W]+$"
    
    regexObject_trim_LEFT.pattern = pattern_trim_LEFT
    regexObject_trim_RIGHT.pattern = pattern_trim_RIGHT
    
    'first trim LEFT
    If regexObject_trim_LEFT.Test(userString) = True Then
        userString = regexObject_trim_LEFT.Replace(userString, "")
    End If
    
    'next trim RIGHT
    If regexObject_trim_RIGHT.Test(userString) = True Then
        userString = regexObject_trim_RIGHT.Replace(userString, "")
    End If
    
    SuperTrim = userString
    
End Function
