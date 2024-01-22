Function myVarType(variable)

    'PURPOSE:   An improved version of VBA's VarType function. Determine variable type, then output a written description of the type instead of giving a number.
    'REFERENCE: https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/vartype-function

    Dim result As Long: result = VarType(variable)
    
    If result = 0 Then
        myVarType = "vbEmpty | Empty (uninitialized)"
    ElseIf result = 1 Then
        myVarType = "vbNull | Null (no valid data)"
    ElseIf result = 2 Then
        myVarType = "vbInteger | Integer"
    ElseIf result = 3 Then
        myVarType = "vbLong | Long integer"
    ElseIf result = 4 Then
        myVarType = "vbSingle | Single-precision floating-point number"
    ElseIf result = 5 Then
        myVarType = "vbDouble | Double-precision floating-point number"
    ElseIf result = 6 Then
        myVarType = "vbCurrency | Currency value"
    ElseIf result = 7 Then
        myVarType = "vbDate | Date value"
    ElseIf result = 8 Then
        myVarType = "vbString | String"
    ElseIf result = 9 Then
        myVarType = "vbObject | Object"
    ElseIf result = 10 Then
        myVarType = "vbError | Error value"
    ElseIf result = 11 Then
        myVarType = "vbBoolean | Boolean value"
    ElseIf result = 12 Then
        myVarType = "vbVariant | Variant (used only with arrays of variants)"
    ElseIf result = 13 Then
        myVarType = "vbDataObject | A data access object"
    ElseIf result = 14 Then
        myVarType = "vbDecimal | Decimal value"
    ElseIf result = 17 Then
        myVarType = "vbByte | Byte value"
    ElseIf result = 20 Then
        myVarType = "vbLongLong | LongLong integer (valid on 64-bit platforms only)"
    ElseIf result = 36 Then
        myVarType = "vbUserDefinedType | Variants that contain user-defined types"
    ElseIf result = 8192 Then
        myVarType = "vbArray | Array (always added to another constant when returned by this function)"
    Else
        myVarType = "Unknown variable type: " + result
    End If

End Function
