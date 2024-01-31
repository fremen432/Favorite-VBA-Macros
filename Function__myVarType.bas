Function myVarType(variable)

    'PURPOSE:   An alternate version of VBA's VarType function. Determine variable type, then output a string description of the type instead of giving a number.
    'REFERENCE: https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/vartype-function

    Dim result As Long: result = VarType(variable)
    
If result = 0 Then
        myVarType = "0 | vbEmpty | Empty (uninitialized)"
    ElseIf result = 1 Then
        myVarType = "1 | vbNull | Null (no valid data)"
    ElseIf result = 2 Then
        myVarType = "2 | vbInteger | Integer"
    ElseIf result = 3 Then
        myVarType = "3 | vbLong | Long integer"
    ElseIf result = 4 Then
        myVarType = "4 | vbSingle | Single-precision floating-point number"
    ElseIf result = 5 Then
        myVarType = "5 | vbDouble | Double-precision floating-point number"
    ElseIf result = 6 Then
        myVarType = "6 | vbCurrency | Currency value"
    ElseIf result = 7 Then
        myVarType = "7 | vbDate | Date value"
    ElseIf result = 8 Then
        myVarType = "8 | vbString | String"
    ElseIf result = 9 Then
        myVarType = "9 | vbObject | Object"
    ElseIf result = 10 Then
        myVarType = "10 | vbError | Error value"
    ElseIf result = 11 Then
        myVarType = "11 | vbBoolean | Boolean value"
    ElseIf result = 12 Then
        myVarType = "12 | vbVariant | Variant (used only with arrays of variants)"
    ElseIf result = 13 Then
        myVarType = "13 | vbDataObject | A data access object"
    ElseIf result = 14 Then
        myVarType = "14 | vbDecimal | Decimal value"
    ElseIf result = 17 Then
        myVarType = "17 | vbByte | Byte value"
    ElseIf result = 20 Then
        myVarType = "20 | vbLongLong | LongLong integer (valid on 64-bit platforms only)"
    ElseIf result = 36 Then
        myVarType = "36 | vbUserDefinedType | Variants that contain user-defined types"
    ElseIf result = 8192 Then
        myVarType = "8192 | vbArray | Array (always added to another constant when returned by this function)"
    Else
        myVarType = "Unknown variable type: " + result
    End If

End Function
