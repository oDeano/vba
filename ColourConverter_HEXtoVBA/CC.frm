VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CC 
   Caption         =   "UserForm1"
   ClientHeight    =   4425
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6885
   OleObjectBlob   =   "CC.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton1_Click()
    Call convertColour
End Sub

Private Sub CommandButton2_Click()
    CC.Hide
End Sub

Private Sub CommandButton3_Click()
    Call copyToClipboard
End Sub
