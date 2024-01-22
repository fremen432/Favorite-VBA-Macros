Sub Call_Python_From_VBA()

    'PURPOSE:   Call and run an existing python script from VBA
    'REFERENCE: https://stackoverflow.com/questions/18135551/how-to-call-python-script-on-excel-vba

    Dim pythonEXE_path As String
    Dim pythonScript_path As String
    
    pythonEXE_path = ""
        'Example: "C:\Users\user123\AppData\Local\Programs\Python\Python310\python.exe"
    pythonScript_path = ""
        'Example: "C:\Users\user123\my_python_script.py"

    pythonEXE_path_Exists = FileExists(pythonEXE_path)
    pythonScript_path_Exists = FileExists(pythonScript_path)
    
    If pythonEXE_path_Exists = False Then
        MsgBox "Python exe file does not exist at this path: " & pythonEXE_path
        Exit Sub
    End If
    
    If pythonScript_path_Exists = False Then
        MsgBox "Python script file does not exist at this path: " & pythonScript_path
        Exit Sub
    End If
    
    If pythonEXE_path_Exists = True And pythonScript_path_Exists = True Then
        RetVal = Shell(pythonEXE_path & " " & pythonScript_path)
    Else
        MsgBox "Error running python script with: " & vbCr & vbTab & pythonEXE_path & vbCr & vbTab & "and" & vbCr & vbTab & pythonScript_path
        
End Sub
