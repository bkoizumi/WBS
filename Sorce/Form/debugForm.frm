VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} debugForm 
   Caption         =   "UserForm1"
   ClientHeight    =   4500
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9780
   OleObjectBlob   =   "debugForm.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "debugForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'
'Private Declare Function GetForegroundWindow Lib "user32" () As Long
'Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
'Private Const HWND_TOPMOST As Long = -1
'Private Const SWP_NOSIZE As Long = &H1&
'Private Const SWP_NOMOVE As Long = &H2&
'
'
'Private Sub UserForm_Activate()
'    Call SetWindowPos(GetForegroundWindow, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
'    Me.StartUpPosition = 1
'End Sub


