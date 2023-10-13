Sub RunPythonScript()

    Dim objShell As Object
    Dim PythonScriptPath As String

    Set objShell = VBA.CreateObject("Wscript.Shell")

    ' Specify the path to your Python script
    PythonScriptPath = "C:\VS Code\Final_code_split_file_automation.py"

    ' Use the cmd.exe to run the Python script with py.exe
    Dim cmd As String
    cmd = "cmd /k py.exe """ & PythonScriptPath & """"

    ' Display a message and ask for user input
    Dim userInput As String
    userInput = InputBox("Do you want to run the Python script (y/n)?", "User Input")

    If LCase(userInput) = "y" Then
        ' Execute the command
        On Error Resume Next
        Dim result As Integer
        result = objShell.Run(cmd, 1, True)
        If result = 0 Then
            MsgBox "Python script executed successfully."
        Else
            MsgBox "Error executing Python script. Error code: " & result
        End If
    ElseIf LCase(userInput) = "n" Then
        MsgBox "Python script execution canceled by user."
    Else
        MsgBox "Invalid input. Please enter 'y' to run the script or 'n' to cancel."
    End If

    ' Release the object
    Set objShell = Nothing

End Sub
