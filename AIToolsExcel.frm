VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AIToolsExcel 
   Caption         =   "AI Tools"
   ClientHeight    =   4680
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   8916.001
   OleObjectBlob   =   "AIToolsExcel.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AIToolsExcel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButtonApply_Click()
    Dim addr As String
    Dim r As Range
    
    addr = Me.Tag
    
    If addr <> "" Then
        Set r = Range(addr)
        r.Value = Me.TextBoxOutput.text
    End If
    
    Me.TextBoxInput.text = ""
    Me.TextBoxOutput.text = ""
    Me.Tag = ""
    Me.Hide
End Sub

Private Sub CommandButtonCancel_Click()
    Me.TextBoxInput.text = ""
    Me.TextBoxOutput.text = ""
    Me.Tag = ""
    Me.Hide
End Sub
