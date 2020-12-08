VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Frm_Filter 
   Caption         =   "�^�X�N���o - WBS"
   ClientHeight    =   8055
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10965
   OleObjectBlob   =   "Frm_Filter.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "Frm_Filter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
#If Win64 Then
  Private Declare PtrSafe Function GetForegroundWindow Lib "user32" () As LongPtr
  Private Declare PtrSafe Function SetWindowPos Lib "user32" (ByVal hWnd As LongPtr, ByVal hWndInsertAfter As LongPtr, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
#Else
  Private Declare Function GetForegroundWindow Lib "user32" () As Long
  Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
#End If
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

'**************************************************************************************************
' * �����ݒ�
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Private Sub UserForm_Initialize()
    
  Dim item As Variant, list As Variant
  Dim col As Collection
  
  
  '�}�E�X�J�[�\����W���ɐݒ�
  Application.Cursor = xlDefault
  
  '�S���҃��X�g�̎擾
  Call init.setting
  Call Task.�S���Ғ��o(col)
  With memberList
    For Each item In col
      .AddItem item
    Next
  End With
  
  With taskLeveList
    For list = 1 To Application.WorksheetFunction.Max(Range("B6:B" & sheetMain.Cells(Rows.count, 2).End(xlUp).row))
      .AddItem list
    Next
  End With
  
  Call Task.�^�X�N�����o(col)
  With taskNameList
    .RowSource = "�ݒ�!" & Range(setVal("cell_DataExtract") & "3:" & setVal("cell_DataExtract") & sheetSetting.Cells(Rows.count, Library.getColumnNo(setVal("cell_DataExtract"))).End(xlUp).row).Address
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
  Dim filterName As String
  Dim endLine As Long
  
  Select Case True
    Case memberList.Enabled = True
      Call Task.�S���҃t�B���^�[(FilterForm.memberList.Value)
      
    Case taskLeveList.Enabled = True
    
    Case taskNameList.Enabled = True
      With FilterForm.taskNameList
        For i = 0 To .ListCount - 1
            If .Selected(i) = True Then
              If filterName = "" Then
                filterName = .list(i)
              Else
                filterName = filterName & "<>" & .list(i)
              End If
            End If
        Next i
      End With
      Call Task.�^�X�N���t�B���^�[(filterName)
    
    
    Case Else
  End Select
  

End Sub

Private Sub Cancel_Click()
  Unload Me
End Sub


