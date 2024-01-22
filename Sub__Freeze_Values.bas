Sub Freeze_Values()

    'PURPOSE:   Copy selected cells and paste only their values (not their formula) in the same selected cell range

    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
End Sub
