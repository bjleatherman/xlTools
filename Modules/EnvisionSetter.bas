Sub SetDistanceInEnvision()

    Dim commandType As String: commandType = "dist"
    Dim userVal As Variant: userVal = ActiveCell.Value
    
    If Not IsNumeric(userVal) Then
        Debug.Print "bad"
        Exit Sub
    End If
    
    Call CallEvisionSetterScript(userVal, commandType)

End Sub

Sub SetTimeInEnvision()

    Dim commandType As String: commandType = "time"
    Dim userVal As Variant: userVal = ActiveCell.Value
    
    If Not IsDate(userVal) Then
        Debug.Print "bad"
        Exit Sub
    End If
    
    userVal = Format(userVal, "mm/dd/yyyy hh:nn:ss")
    
    Call CallEvisionSetterScript(userVal, commandType)
    
End Sub

Sub CallEvisionSetterScript(userVal As Variant, commandType As String)

    Dim scriptPath As String: scriptPath = "C:\UtilScripts\goToEnvDistTime.exe"
    Dim command As String: command = scriptPath & " " & """" & userVal & """" & " " & commandType

    Debug.Print command

    shell command

End Sub

Sub ServerSetDistanceInEnvision()

    Dim userVal As Variant: userVal = ActiveCell.Value
    If Not IsNumeric(userVal) Then
        'Debug.Print "bad"
        Exit Sub
    End If
    
    
    Dim staticUrl As String: staticUrl = "http://127.0.0.1:5000/go_to?dist="
    Dim fullUrl As String: fullUrl = staticUrl & userVal
    'Debug.Print fullUrl
    
    Dim http As Object: Set http = CreateObject("MSXML2.ServerXMLHTTP")
    http.Open "GET", fullUrl, False
    http.Send
    Debug.Print http.responseText
    Set http = Nothing
    
    'Dim shell As Object: Set shell = CreateObject("WScript.shell")
    'Dim exec As Object
    'curlCmd = "curl " & fullUrl
    'Set exec = shell.exec(curlCmd)
    
    ' Loop through the output
    'Do While Not exec.StdOut.AtEndOfStream
    '    output = output & exec.StdOut.ReadLine & vbCrLf
    'Loop
    
    'Debug.Print output

    ' Clean up objects
    'Set exec = Nothing
    'Set shell = Nothing

End Sub


Sub test()

shell "curl 127.0.0.1:5000/go_to?dist=10"

End Sub
