VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Frm_FullScreen 
   Caption         =   "�S��ʉ���"
   ClientHeight    =   645
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   1800
   OleObjectBlob   =   "Frm_FullScreen.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
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
' * �����ݒ�
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Private Sub UserForm_Initialize()
  
  '�}�E�X�J�[�\����W���ɐݒ�
  Application.Cursor = xlDefault
End Sub

