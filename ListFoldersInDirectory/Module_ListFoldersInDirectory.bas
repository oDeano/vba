Attribute VB_Name = "Module1"
Option Explicit

Sub ListFoldersInDirectory()

'This Macro will prompt you to select a file directory, then list all folders in a new worksheet

    Dim objFSO As Object
    Dim objFolders As Object
    Dim objFolder As Object
    Dim strDirectory As String
    Dim arrFolders() As String
    Dim FolderCount As Long
    Dim FolderIndex As Long


    With Application.FileDialog(msoFileDialogFolderPicker)
        .InitialFileName = Application.DefaultFilePath & "\"
        .Title = "Select Folder"
        .Show
        If .SelectedItems.Count = 0 Then
            Exit Sub
        End If
        strDirectory = .SelectedItems(1)
    End With
    
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objFolders = objFSO.GetFolder(strDirectory).SubFolders
    
    FolderCount = objFolders.Count
    
    If FolderCount > 0 Then
        ReDim arrFolders(1 To FolderCount)
        FolderIndex = 0
        For Each objFolder In objFolders
            FolderIndex = FolderIndex + 1
            arrFolders(FolderIndex) = objFolder.Name
        Next objFolder
        Worksheets.Add
        Range("A1").Resize(FolderCount).Value = Application.Transpose(arrFolders)
    Else
        MsgBox "No folders found!", vbExclamation
    End If
    
    Set objFSO = Nothing
    Set objFolders = Nothing
    Set objFolder = Nothing
    
End Sub
