VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Frm_Debug 
   Caption         =   "UserForm1"
   ClientHeight    =   4500
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9780
   OleObjectBlob   =   "Frm_debug.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "Frm_debug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetFocus Lib "user32" (ByVal hWnd As Long) As Long


Private Sub UserForm_Activate()
  SetFocus Application.hWnd
End Sub
