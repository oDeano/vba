
Sub CopyFile(oldPath As String, newPath As String, targetFolder As String) As Boolean
    Dim v As Variant
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If fso.FolderExists(targetFolder) = True Then
        If fso.FileExists(newPath) = True Then
            v = MsgBox("File already exists!", , "Error!")
            CopyFile = False
            Exit Function
        Else
            fso.CopyFile oldPath, newPath
            CopyFile = True
        End If
    Else
        fso.CreateFolder (targetFolder)
        fso.CopyFile oldPath, newPath
        CopyFile = True
    End If
End Sub