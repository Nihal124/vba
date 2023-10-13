Sub RunPythonScriptWithConfirmation()

    Dim objshell As Object
    Dim PythonScriptPath As String

    Set objshell = VBA.CreateObject("Wscript.Shell")

    ' Specify the path to your Python script
    PythonScriptPath = "C:\VS Code\Final_code_split_file_automation.py"

    ' Use the cmd.exe to run the Python script with py.exe after user confirmation
    Dim cmd As String
    cmd = "cmd /k echo off & echo. & set /p choice=Run Python script? (y/n): & if /i !choice! == y (py.exe """ & PythonScriptPath & """) else (echo Script not run) & pause & exit"

    ' Execute the command
    On Error Resume Next
    Dim result As Integer
    result = objshell.Run(cmd, 1, True)

    If result = 0 Then
        MsgBox "Command Prompt opened."
    Else
        MsgBox "Error opening Command Prompt. Error code: " & result
    End If

    ' Release the object
    Set objshell = Nothing

End Sub
