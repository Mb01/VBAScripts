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

Sub ReplaceMatchAny(dict As Object)
    Dim ws As Worksheet
    Dim key As Variant

    For Each ws In ThisWorkbook.Sheets
        For Each key In dict.Keys
            ws.UsedRange.Replace What:=key, Replacement:=dict(key), _
                                 LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=True, _
                                 SearchFormat:=False, ReplaceFormat:=False
        Next key
    Next ws
End Sub

Sub FindReplaceUsingDictionary(dict As Object)
    Dim ws As Worksheet
    Dim key As Variant

    For Each ws In ThisWorkbook.Sheets
        For Each key In dict.Keys
            With ws.UsedRange
                .Replace What:=key, Replacement:=dict(key), _
                         LookAt:=xlWhole, SearchOrder:=xlByRows, MatchCase:=True, _
                         SearchFormat:=False, ReplaceFormat:=False
            End With
        Next key
    Next ws
End Sub

Sub FindReplaceUsingRegex(dict As Object)
    Dim ws As Worksheet
    Dim cell As Range
    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Global = True

    For Each ws In ThisWorkbook.Sheets
        For Each cell In ws.UsedRange
            For Each key In dict.Keys
                regEx.Pattern = key
                If regEx.Test(cell.Value) Then
                    cell.Value = regEx.Replace(cell.Value, dict(key))
                End If
            Next key
        Next cell
    Next ws
End Sub

Sub Main()
    Dim dict As Object
    Dim filePath As String

    ' Replace any part of string that matches
    filePath = "...fullHalf.txt"
    Set dict = ReadDictionaryFromFile(filePath)
    ReplaceMatchAny dict
    
    ' Uses regular expressions to replace
    filePath = "...regexReplace.txt"
    Set dict = ReadDictionaryFromFile(filePath)
    FindReplaceUsingRegex dict

    ' Replaces when entire cell is exact match
    filePath = "...findReplace.txt"
    Set dict = ReadDictionaryFromFile(filePath)
    FindReplaceUsingDictionary dict

End Sub

