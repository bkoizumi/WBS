Attribute VB_Name = "ProgressBar"

'**************************************************************************************************
' * �v���O���X�o�[�\������
' *
' * http://www.ne.jp/asahi/fuji/lake/excel/progress_01.html
'**************************************************************************************************
'Option Explicit

Public mypbProgCnt As Long       'Progress �J�E���^�[�ϐ�
Public mypbSCount As Long        '������

Dim myJobCnt As Long             '���ݐi�s���̉�
Dim myBarSize As Long            '�v���O���X�o�[�T�C�Y


'**************************************************************************************************
' * �v���O���X�o�[�\���J�n
' *
'**************************************************************************************************
Public Sub showStart()
  Dim myMsg1 As String
  
  myMsg1 = " �ҋ@��"

  '�_�C�A���O�֕\��
  With FProgress
    .StartUpPosition = 0
    If ThisWorkbook.Worksheets("�ݒ�").Range("B3") = "develop" Then
      .top = Application.top + 50
    Else
      .top = Application.top + (ActiveWindow.Width / 4)
    End If
    .Left = Application.Left + (ActiveWindow.Height / 2)
    .Caption = myMsg1
    
    '�v���O���X�o�[�̘g�̕���
    With .Label1
      .BorderStyle = fmBorderStyleSingle       '�g����
      .SpecialEffect = fmSpecialEffectSunken
      .Height = 15
      .Left = 12
      .Width = 250
      .top = 30
    End With

    '�v���O���X�o�[�̃o�[�̕���
    With .Label2
      .BackColor = RGB(90, 248, 82)
'        .BorderStyle = fmBorderStyleSingle       '�g����
      .SpecialEffect = fmSpecialEffectRaised
      .Height = 13
      .Left = 13
      .Width = 0
      .top = 31
    End With

    '�i���󋵕\���̕��� ( % )
    With .Label3
      .TextAlign = fmTextAlignCenter
      .Caption = "0%"
      .BackStyle = 0
      .Height = 14
      .Left = 12
      .Width = 250
      .top = 32
      .Font.Size = 10
      .Font.Bold = False
    End With
    
    
    '���b�Z�[�W�\���̕���
    With .Label4
      '.TextAlign = fmTextAlignCenter
      '.SpecialEffect = fmSpecialEffectEtched   '�g������
      '.SpecialEffect = fmSpecialEffectRaised   '�����オ��
      '.SpecialEffect = fmSpecialEffectBump
      .Caption = "�ҋ@��"
      .Height = 14
      .Left = 12
      .Width = 250
      .top = 9
      .Font.Size = 9
      .Font.Bold = False
    End With

    myBarSize = .Label3.Width
  End With

  FProgress.Show vbModeless
End Sub


'**************************************************************************************************
' * �v���O���X�o�[�\���X�V
' *
'**************************************************************************************************
Public Sub showCount(ProgressBarTitle As String, mypbProgCnt As Long, mypbSCount As Long, myMsg1 As String)
  Dim myMsg2 As String
  
  If mypbProgCnt > 0 Then
    myJobCnt = Int(mypbProgCnt / mypbSCount * 100)
    myMsg2 = mypbProgCnt & "/" & mypbSCount & " (" & Int(myJobCnt) & "%)"
    
    With FProgress
      .Caption = ProgressBarTitle
      .Label2.Width = Int(myBarSize * myJobCnt / 100)
      .Label3.Caption = myMsg2
      .Label4.Caption = " ������ �c�@" & myMsg1
    End With
  ElseIf mypbProgCnt = 0 Then
    myJobCnt = Int(mypbProgCnt / mypbSCount * 100)
    myMsg2 = ""
    
    With FProgress
      .Caption = ProgressBarTitle
      .Label2.Width = Int(myBarSize * myJobCnt / 100)
      .Label3.Caption = myMsg2
      .Label4.Caption = " ������ �c�@" & myMsg1
    End With
  
  End If
  If setVal("debugMode") = "develop" And myMsg1 <> "" Then
    Call Library.showDebugForm(ProgressBarTitle, myMsg1)
  End If
  

  DoEvents
  
End Sub


'**************************************************************************************************
' * �v���O���X�o�[�\���I��
' *
'**************************************************************************************************
Public Sub showEnd()
  
  '�_�C�A���O�����
  Unload FProgress
  
End Sub

