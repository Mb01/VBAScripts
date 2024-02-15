Function ReadDictionaryFromFile(filePath As String) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Open
    stream.Charset = "UTF-8"
    stream.LoadFromFile filePath

    Dim txtLine As String
    Dim splitLine() As String

    Do While Not stream.EOS
        txtLine = stream.ReadText(-2) ' Read line
        splitLine = Split(txtLine, vbTab)
        If UBound(splitLine) >= 1 Then
            dict(splitLine(0)) = splitLine(1)
        End If
    Loop
    stream.Close

    Set ReadDictionaryFromFile = dict
End Function

Sub FindReplaceUsingDictionary(dict As Object)
    Dim ws As Worksheet
    Dim key As Variant

    For Each ws In ThisWorkbook.Sheets
        For Each key In dict.keys
            With ws.UsedRange
                .Replace What:=key, Replacement:=dict(key), _
                         LookAt:=xlWhole, SearchOrder:=xlByRows, MatchCase:=True, _
                         SearchFormat:=False, ReplaceFormat:=False
            End With
        Next key
    Next ws
    
    MsgBox "Find and replace complete!", vbInformation
End Sub
