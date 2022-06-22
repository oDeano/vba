Attribute VB_Name = "Module_colourConversion"
Sub convertColour()
    
    Dim uf As Object
    Dim inp As String
    Dim hex As String
    Dim r As String
    Dim g As String
    Dim b As String
    Dim cVBA As String
    
    Set uf = CC
    inp = CC.inpHex
    
    If Len(inp) = 7 Then
        hex = Right(inp, 6)
        r = Left(hex, 2)
        g = Mid(hex, 3, 2)
        b = Right(hex, 2)
    ElseIf Len(inp) = 6 Then
        hex = inp
        r = Left(hex, 2)
        g = Mid(hex, 3, 2)
        b = Right(hex, 2)
    End If
    

    
    cVBA = "&H00" & b & g & r & "&"
    
    CC.outVBA.Value = cVBA
    

End Sub
Sub showCC()
    CC.Show
End Sub
Sub copyToClipboard()
    Clipboard CC.outVBA.Value
End Sub
Function Clipboard(varText As Variant) As String
  Dim objCP As Object
  Set objCP = CreateObject("HtmlFile")
  objCP.ParentWindow.ClipboardData.SetData "text", varText
End Function
