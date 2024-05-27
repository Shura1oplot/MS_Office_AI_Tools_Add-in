VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AIToolsCustomTransformSettings 
   Caption         =   "AI Tools Custom Transform Settings"
   ClientHeight    =   4932
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   7536
   OleObjectBlob   =   "AIToolsCustomTransformSettings.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AIToolsCustomTransformSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub SaveButton_Click()
    Dim temperature As Double

    On Error GoTo InvalidValue
    temperature = CDbl(Me.TemperatureTextBox.Text)
    On Error GoTo 0
    GoTo NoErrors

InvalidValue:
    MsgBox "The temperature should be a number!", vbCritical, "AI Tools: Error"
    Exit Sub

NoErrors:
    SaveSetting "AI tools", "CustomTransform1", "Prompt", Me.PromptTextBox.Text
    SaveSetting "AI tools", "CustomTransform1", "Temperature", Me.TemperatureTextBox.Text
    Me.Hide
End Sub
