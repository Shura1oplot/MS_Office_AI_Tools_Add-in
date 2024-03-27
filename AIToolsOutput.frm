VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AIToolsOutput 
   Caption         =   "AI Tools"
   ClientHeight    =   8304.001
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   10236
   OleObjectBlob   =   "AIToolsOutput.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AIToolsOutput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ButtonApply_Click()
    Me.Tag = Me.TextBoxOutput.SelText
    Me.Hide
End Sub

Private Sub ButtonCancel_Click()
    Me.Tag = ""
    Me.Hide
End Sub
