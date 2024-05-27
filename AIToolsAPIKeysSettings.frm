VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AIToolsAPIKeysSettings 
   Caption         =   "AI Tools API Keys"
   ClientHeight    =   2856
   ClientLeft      =   48
   ClientTop       =   228
   ClientWidth     =   5868
   OleObjectBlob   =   "AIToolsAPIKeysSettings.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AIToolsAPIKeysSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub SaveButton_Click()
    SaveSetting "AI tools", "API Keys", "openai", Me.OpenAIAPIKeyTextBox.Text
    SaveSetting "AI tools", "API Keys", "google", Me.GoogleAIAPIKeyTextBox.Text
    SaveSetting "AI tools", "API Keys", "anthropic", Me.AnthropicAPIKeyTextBox.Text
    Me.Hide
End Sub

