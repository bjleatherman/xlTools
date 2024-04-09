Function RegexMatch(inputString As String, regexPattern As String, Optional matchIndex As Integer = 0) As String
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    
    'Dig num pattern- Dig\s\d+(?:\.[^p]\d?|[A-Z]?)?
    'Dig distance from file name- (\d+.?\d{2})
    
    With regex
        .Global = True
        .MultiLine = True
        .IgnoreCase = True
        .Pattern = regexPattern
    End With
    
    Dim matches As Object
    Set matches = regex.Execute(inputString)
    
    If matches.count > 0 Then
        RegexMatch = matches(matchIndex).Value
    Else
        RegexMatch = ""
    End If
End Function
