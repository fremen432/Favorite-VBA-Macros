Function myRegExp_replace(userString, userPattern, replacementString)

    'PURPOSE:   Given a string "userString", find all matches for a RegExp pattern "userPattern" and replace all matches with a replacement string "replacementString"
    'REFERENCE: https://software-solutions-online.com/vba-regex-guide/#:~:text=To%20start%20using%20RegEx%20in,Expressions%205.5%2C%20then%20click%20OK.
    'NOTE:      In order to use Regular Expressions in Microsoft Excel you will need to active RegExp object reference library. Directions in reference above.

    Dim result As String
    Dim regexObject As RegExp
    Set regexObject = New RegExp
    
    With regexObject
        .pattern = userPattern
    End With
        
    If regexObject.Test(userString) <> True Then ' if no match is found, give error message, return userString and exit function
    
        MsgBox "No match for pattern """ & userPattern & """ within the string """ & userString & """"

        myRegExp_replace = userString
        Exit Function
    End If
    
    result = regexObject.Replace(userString, replacementString)
    
    myRegExp_replace = result

End Function
