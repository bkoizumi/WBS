VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ErrorForm 
   Caption         =   "�G���[���e"
   ClientHeight    =   3765
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9450
   OleObjectBlob   =   "ErrorForm.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "ErrorForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************************************************
' * �����ݒ�
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Private Sub UserForm_Initialize()
  
  '�}�E�X�J�[�\����W���ɐݒ�
  Application.Cursor = xlDefault
End Sub



'**************************************************************************************************
' * �L�����Z��
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Private Sub Cancel_Click()

  Unload Me
End Sub



