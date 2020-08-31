VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FilterForm 
   Caption         =   "�^�X�N���o - WBS"
   ClientHeight    =   8055
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10965
   OleObjectBlob   =   "FilterForm.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "FilterForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_TOPMOST As Long = -1
Private Const SWP_NOSIZE As Long = &H1&
Private Const SWP_NOMOVE As Long = &H2&











Private Sub UserForm_Activate()
    Call SetWindowPos(GetForegroundWindow, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
    Me.StartUpPosition = 1
End Sub



'Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
'    If CloseMode = vbFormControlMenu Then
'        Cancel = True
'    End If
'End Sub



Private Sub UserForm_Initialize()
    
  Dim item As Variant, list As Variant
  Dim col As Collection
  
  '�S���҃��X�g�̎擾
  Call init.setting
  Call Task.�S���Ғ��o(col)
  With memberList
    For Each item In col
      .AddItem item
    Next
  End With
  With taskLeveList
    For list = 1 To Application.WorksheetFunction.Max(Range("B6:B" & mainSheet.Cells(Rows.count, 2).End(xlUp).row))
      .AddItem list
    Next
  End With
  Call Task.�^�X�N�����o(col)
  With taskNameList
    For Each item In col
      .AddItem item
    Next
  End With
  
  If setVal("workMode") = "CD���p" Then
    filterTaskName.Value = True
    Call filterTaskName_Click
  Else
    memberList.Enabled = False
    taskLeveList.Enabled = False
    taskNameList.Enabled = False
  End If
End Sub


'�S���҂Ńt�B���^�[��I��
Private Sub filterAssign_Click()
  frameAssign.BackColor = &HC0C0C0
  frameTaskLevel.BackColor = &H8000000F
  frameTaskName.BackColor = &H8000000F
  
  memberList.Enabled = True
  taskLeveList.Enabled = False
  taskNameList.Enabled = False
End Sub

'�^�X�N���x���Ńt�B���^�[��I��
Private Sub filterTaskLeve_Click()
  frameAssign.BackColor = &H8000000F
  frameTaskLevel.BackColor = &HC0C0C0
  frameTaskName.BackColor = &H8000000F
  
  memberList.Enabled = False
  taskLeveList.Enabled = True
  taskNameList.Enabled = False
End Sub

'�^�X�N���Ńt�B���^�[��I��
Private Sub filterTaskName_Click()
  frameAssign.BackColor = &H8000000F
  frameTaskLevel.BackColor = &H8000000F
  frameTaskName.BackColor = &HC0C0C0
  
  memberList.Enabled = False
  taskLeveList.Enabled = False
  taskNameList.Enabled = True
End Sub




'�������s
Private Sub run_Click()
  Select Case True
    Case memberList.Enabled = True
      Call Task.�S���҃t�B���^�[(FilterForm.memberList.Value)
      
    Case taskLeveList.Enabled = True
    
    Case taskNameList.Enabled = True
      Call Task.�^�X�N���t�B���^�[(FilterForm.taskNameList.Value)
    
    
    Case Else
  End Select
  

End Sub

Private Sub Cancel_Click()
  Unload Me
End Sub


