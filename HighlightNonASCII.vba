Sub HighlightNonASCIIIncludingTextBoxes()
    Dim ws As Worksheet
    Dim cell As Range
    Dim shape As shape
    Dim i As Integer

    Set ws = ActiveSheet

    ' Check each cell in used range
    For Each cell In ws.UsedRange
        For i = 1 To Len(cell.Value)
            If AscW(Mid(cell.Value, i, 1)) > 127 Then
                cell.Interior.Color = RGB(255, 0, 0) ' Highlight cell in red
                Exit For
            End If
        Next i
    Next cell

    ' Check each shape in the worksheet
    For Each shape In ws.Shapes
        On Error Resume Next ' Start error handling
        CheckAndHighlightShape shape
        If Err.Number <> 0 Then
            Debug.Print "Error with shape: " & shape.Name & " - " & Err.Description
            Err.Clear ' Clear the error
            Exit Sub ' End the function
        End If
        On Error GoTo 0 ' Turn off error handling
    Next shape
End Sub

Sub CheckAndHighlightShape(ByVal shape As shape)
    Dim i As Integer
    Dim char As String
    Dim text As String

    ' Check if the shape is a group
    If shape.Type = msoGroup Then
        ' Iterate through each shape in the group
        For i = 1 To shape.GroupItems.Count
            CheckAndHighlightShape shape.GroupItems(i)
        Next i
    Else
        ' Check if the shape has a text frame and text
        If Not shape.TextFrame2 Is Nothing Then
            If shape.TextFrame2.HasText Then
                ' Using TextFrame2 and TextRange for compatibility with newer Excel versions
                text = shape.TextFrame2.TextRange.text
                For i = 1 To Len(text)
                    char = Mid(text, i, 1)
                    If AscW(char) > 127 Then
                        shape.Fill.ForeColor.RGB = RGB(255, 0, 0) ' Highlight shape in red
                        Exit For
                    End If
                Next i
            End If
        End If
    End If
End Sub

