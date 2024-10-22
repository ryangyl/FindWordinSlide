Attribute VB_Name = "Module3"
Sub findryouki()
Dim oRng As TextRange
Dim oSh As Shape
Dim oSl As slide
Dim x As Long
Set activePresentation1 = activePresentation
For Each oSl In activePresentation1.Slides
For Each oSh In oSl.Shapes
If oSh.HasTextFrame Then
If oSh.TextFrame.HasText Then
For Each oRng In oSh.TextFrame.TextRange.Words
If oRng = "RYOUKI" Then
oSl.Select
GoTo error1234444444
End If
Next
End If
End If
Next
Next
error1234444444:
Exit Sub
End Sub

