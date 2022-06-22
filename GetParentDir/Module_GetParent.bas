Attribute VB_Name = "Module_GetParent"
' Returns parent dir of current workbook


Function GetParent() As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    GetParent = fso.GetParentFolderName(Application.ActiveWorkbook.path)
End Function