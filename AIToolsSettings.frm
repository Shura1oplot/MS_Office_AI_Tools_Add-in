VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AIToolsSettings 
   Caption         =   "AI Tools Settings"
   ClientHeight    =   2856
   ClientLeft      =   48
   ClientTop       =   228
   ClientWidth     =   5868
   OleObjectBlob   =   "AIToolsSettings.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AIToolsSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub SaveButton_Click()
    SaveSetting "AI tools", "API Keys", "openai", Me.OpenAIAPIKeyTextBox.text
    SaveSetting "AI tools", "API Keys", "google", Me.GoogleAIAPIKeyTextBox.text
    SaveSetting "AI tools", "API Keys", "anthropic", Me.AnthropicAPIKeyTextBox.text
    Me.Hide
End Sub

