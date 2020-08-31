VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ExpansionForm 
   Caption         =   "Šg‘å"
   ClientHeight    =   6795
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10335
   OleObjectBlob   =   "ExpansionForm.frx":0000
   StartUpPosition =   2  '‰æ–Ê‚Ì’†‰›
End
Attribute VB_Name = "ExpansionForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False







Private Sub CancelButton_Click()
  Unload ExpansionForm
End Sub

Private Sub OK_Button_Click()
  Call Library.showExpansionFormClose(TextBox, ExpansionForm.Caption)
  Unload ExpansionForm
End Sub
