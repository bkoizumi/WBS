VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DispFullScreenForm 
   Caption         =   "全画面解除"
   ClientHeight    =   645
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   1800
   OleObjectBlob   =   "DispFullScreenForm.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "DispFullScreenForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
  Unload Me
  
  Application.ScreenUpdating = False
  
  Application.DisplayFullScreen = False
  ActiveWindow.DisplayHeadings = True
  
  Application.ScreenUpdating = True
End Sub
