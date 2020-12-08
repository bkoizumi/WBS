VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Frm_FullScreen 
   Caption         =   "全画面解除"
   ClientHeight    =   645
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   1800
   OleObjectBlob   =   "Frm_FullScreen.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "Frm_FullScreen"
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


'**************************************************************************************************
' * 初期設定
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Private Sub UserForm_Initialize()
  
  'マウスカーソルを標準に設定
  Application.Cursor = xlDefault
End Sub

