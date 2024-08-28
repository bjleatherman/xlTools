Sub SetDistanceInEnvision()

    Dim userVal As Variant: userVal = ActiveCell.Value
    
    ' Python
    Dim scriptPathFile As String: scriptPathFile = "C:\UtilScripts\goToEnvDist.py"
    Dim commandFile As String: commandFile = "python " & scriptPathFile & " " & userVal
    
    ' .Exe
    Dim scriptPath As String: scriptPath = "C:\UtilScripts\goToEnvDist.exe"
    Dim command As String: command = scriptPath & " " & userVal
    
    Dim shell As Object: Set shell = CreateObject("WScript.Shell")
    
    If Not IsNumeric(userVal) Then
        'Debug.Print "bad"
        Exit Sub
    End If
    
    ' Execute the command and wait for it to complete
    exitCode = shell.Run(command, 0, True)

    ' Check the exit code
    If exitCode = 0 Then
        'Debug.Print "Script executed successfully."
    Else
        'Debug.Print "Script execution failed with exit code: " & exitCode
    End If
    
    Set shell = Nothing

End Sub
