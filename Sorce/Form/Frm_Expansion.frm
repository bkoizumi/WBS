VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Frm_Expansion 
   Caption         =   "�g��"
   ClientHeight    =   6795
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10335
   OleObjectBlob   =   "Frm_Expansion.frx":0000
   StartUpPosition =   2  '��ʂ̒���
End
Attribute VB_Name = "Frm_Expansion"
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








Private Sub CancelButton_Click()
  Unload ExpansionForm
End Sub

Private Sub OK_Button_Click()
  Call Library.showExpansionFormClose(TextBox, ExpansionForm.Caption)
  Unload ExpansionForm
End Sub
