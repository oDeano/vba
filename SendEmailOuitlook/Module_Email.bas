Sub CreateEmail(to As String, subject As String, body As String, Optional attachment As String, Optional cc As String) As Boolean
    If isOutlookOpen = False Then
        CreateEmail = False
        Exit Function
    End If
    Dim outlook As Object
    Set outlook = CreateObject("Outlook.Application")
    Dim newemail As Object
    Set newemail = outlook.CreateItem(0)
    
    With newemail
        .To = to
        .Subject = subject
        .HTMLBody = body
    End With
    If Not eAttach = "" Then
        newemail.Attachments.Add attachment
    End If
    If Not eCC = "" Then
        newemail.CC = cc
    End If
    newemail.Display
    CreateEmail = True
End Sub

Sub isOutlookOpen() As Boolean
    On Error Resume Next
    Set outlook = GetObject(, "Outlook.Application")
    On Error Resume Next
    
    If outlook Is Nothing Then
        isOutlookOpen = False
        MsgBox "Please open Outlook before using this feature."
    Else
        isOutlookOpen = True
    End If
End Sub
