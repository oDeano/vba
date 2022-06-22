Attribute VB_Name = "Module_copyToClipboard"
Sub copyToClipboard()
    'Replace CC.outVBA.Value with text to copy
    Clipboard CC.outVBA.Value
End Sub
Function Clipboard(varText As Variant) As String
  Dim objCP As Object
  Set objCP = CreateObject("HtmlFile")
  objCP.ParentWindow.ClipboardData.SetData "text", varText
End Function
