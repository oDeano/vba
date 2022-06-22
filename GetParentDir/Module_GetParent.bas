' Returns parent dir of current workbook

Sub GetParent() As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    GetParent = fso.GetParentFolderName(Application.ActiveWorkbook.path)
End Sub