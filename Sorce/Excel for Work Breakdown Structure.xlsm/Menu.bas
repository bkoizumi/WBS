Attribute VB_Name = "Menu"
'**************************************************************************************************
' * �ݒ�
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Sub M_Help()
  Call init.setting
  helpSheet.Visible = True
  helpSheet.Select
End Sub



'**************************************************************************************************
' * �V���[�g�J�b�g�ݒ�
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Sub M_�V���[�g�J�b�g�ݒ�()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  
  Call init.setting(True)
  endLine = Cells(Rows.count, 7).End(xlUp).row
  
  '�ݒ������
  Call M_�V���[�g�J�b�g�ݒ����
  
  For line = 3 To endLine
    If setSheet.Range(setVal("cell_ShortcutKey") & line) <> "" Then
      Application.MacroOptions Macro:="Menu." & setSheet.Range(setVal("cell_ShortcutFuncName") & line), ShortcutKey:=setSheet.Range(setVal("cell_ShortcutKey") & line)
    End If
  Next
  '�C���f���g�̃V���[�g�J�b�g
  Application.OnKey "%{LEFT}", "Menu.M_�C���f���g��"
  Application.OnKey "%{RIGHT}", "Menu.M_�C���f���g��"
  Application.OnKey "%{F1}", "Menu.M_�^�X�N�\��_�W��"
  Application.OnKey "%{F2}", "Menu.M_�^�X�N�\��_�`�[���v�����i�["
End Sub


Sub M_�V���[�g�J�b�g�ݒ����()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  
  Call init.setting
  endLine = Cells(Rows.count, 7).End(xlUp).row
  
  '�ݒ������
  For line = 3 To endLine
    If setSheet.Range("J" & line) <> "" Then
      Application.MacroOptions Macro:="Menu." & setSheet.Range("H" & line), ShortcutKey:=""
    End If
  Next
  
  '�C���f���g�̃V���[�g�J�b�g
  Application.OnKey "%{LEFT}", ""
  Application.OnKey "%{RIGHT}", ""
  Application.OnKey "%{F1}", ""
  Application.OnKey "%{F2}", ""
End Sub

Sub optionKey()
Attribute optionKey.VB_ProcData.VB_Invoke_Func = "O\n14"
  Call M_�I�v�V������ʕ\��
End Sub
Sub centerKey()
End Sub
Sub filterKey()
End Sub
Sub clearFilterKey()
End Sub
Sub taskCheckKey()
Attribute taskCheckKey.VB_ProcData.VB_Invoke_Func = "C\n14"
End Sub
Sub makeGanttKey()
Attribute makeGanttKey.VB_ProcData.VB_Invoke_Func = "t\n14"
End Sub
Sub clearGanttKey()
Attribute clearGanttKey.VB_ProcData.VB_Invoke_Func = "D\n14"
End Sub
Sub dispAllKey()
End Sub
Sub taskControlKey()
End Sub
Sub ScaleKey()
End Sub








Sub M_�I�v�V������ʕ\��()
Attribute M_�I�v�V������ʕ\��.VB_ProcData.VB_Invoke_Func = " \n14"
  
  Call Library.startScript
  Call init.setting(True)
  
  Call WBS_Option.�I�v�V������ʕ\��
  
  Call M_�J�����_�[����
  Call M_�K���g�`���[�g����
  Call WBS_Option.�\����ݒ�
  
  
  Call Library.endScript(True)
End Sub


Sub M_����ւ�()
  Call init.setting
  
  Call Library.startScript
  Call Check.���ڗ�`�F�b�N
  Call init.setting(True)
  
  Call Library.endScript(True)
End Sub

Sub M_�J�����_�[����()

  Call init.setting(True)
  Call Library.startScript
  
  '�S�Ă̍s���\��
  Cells.EntireColumn.Hidden = False
  Cells.EntireRow.Hidden = False
  
  Call Calendar.makeCalendar
  
  Call WBS_Option.�����̒S���ҍs���\��
  Call WBS_Option.�\����ݒ�
  
  Call Library.endScript
End Sub




'**************************************************************************************************
' * ����
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Sub M_�s�n�C���C�g()
  Call Library.startScript
  Call WBS_Option.setLineColor
  Call Library.endScript(True)
End Sub


'--------------------------------------
Sub M_�S�f�[�^�폜()
  If MsgBox("�f�[�^���폜���܂�", vbYesNo + vbExclamation) = vbNo Then
    End
  End If
  
  Call Library.startScript
  Call WBS_Option.clearAll
  Call Library.endScript
End Sub


Sub M_�S���()
Attribute M_�S���.VB_ProcData.VB_Invoke_Func = " \n14"
  Application.ScreenUpdating = False
  ActiveWindow.DisplayHeadings = False
  Application.DisplayFullScreen = True
  
  With DispFullScreenForm
    .StartUpPosition = 0
    .top = Application.top + 300
    .Left = Application.Left + 30
  End With
  Application.ScreenUpdating = True
  DispFullScreenForm.Show vbModeless
End Sub

Sub M_�^�X�N����()
Attribute M_�^�X�N����.VB_ProcData.VB_Invoke_Func = " \n14"
End Sub

Sub M_�X�P�[��()
Attribute M_�X�P�[��.VB_ProcData.VB_Invoke_Func = " \n14"
End Sub


'**************************************************************************************************
' * WBS
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Sub M_�^�X�N�`�F�b�N()
Attribute M_�^�X�N�`�F�b�N.VB_ProcData.VB_Invoke_Func = "C\n14"


  Call init.setting
  mainSheet.Select
  
  Call Library.startScript
  Call ProgressBar.showStart
  
  Call Check.�^�X�N���X�g�m�F
  
  Call ProgressBar.showEnd
  Call Library.endScript(True)

End Sub


Sub M_�t�B���^�[()
Attribute M_�t�B���^�[.VB_ProcData.VB_Invoke_Func = " \n14"
  Call init.setting
  
  With FilterForm
    .StartUpPosition = 0
    .top = Application.top + (ActiveWindow.Width / 8)
    .Left = Application.Left + (ActiveWindow.Height / 8)
  End With
  
  FilterForm.Show
End Sub


Sub M_���ׂĕ\��()
Attribute M_���ׂĕ\��.VB_ProcData.VB_Invoke_Func = " \n14"
  Call Library.startScript
  Rows("6:" & Rows.count).EntireRow.Hidden = False
  
  Call WBS_Option.�����̒S���ҍs���\��
  Call Library.endScript
End Sub


Sub M_�i���R�s�[()
  Call Task.�i���R�s�[
End Sub

Sub M_�C���f���g��()
  Dim selectedCells As Range
  Dim targetCell As Range
  
  On Error Resume Next
  
  Call Library.startScript
  Call init.setting
  mainSheet.Select
   
  Set selectedCells = Selection
  
  For Each targetCell In selectedCells
    Cells(targetCell.row, getColumnNo(setVal("cell_TaskArea"))).InsertIndent 1
  Next
  Call Library.endScript
End Sub


Sub M_�C���f���g��()
  Dim selectedCells As Range
  Dim targetCell As Range
  
  On Error Resume Next
  
  Call Library.startScript
  Call init.setting
  mainSheet.Select
   
  Set selectedCells = Selection
  
  For Each targetCell In selectedCells
    Cells(targetCell.row, getColumnNo(setVal("cell_TaskArea"))).InsertIndent -1
  Next
  Call Library.endScript
End Sub


'�i�����ݒ�----------------------------------------------------------------------------------------
Sub M_�i�����ݒ�(progress As Long)
  Call Task.�i�����ݒ�(progress)
End Sub

'�^�X�N�̃����N�ݒ�/����---------------------------------------------------------------------------
Sub M_�^�X�N�̃����N�ݒ�()
  Call Library.startScript
  Call init.setting
  
  Call Task.taskLink
  
  Call Library.endScript
End Sub

Sub M_�^�X�N�̃����N����()
  Call Library.startScript
  Call init.setting
  
  Call Task.taskUnlink
  
  Call Library.endScript
End Sub

Sub M_�^�X�N�̑}��()
  Call Library.startScript
  Call init.setting
  
  Call Task.�^�X�N�̑}��
  
  Call Library.endScript(True)
End Sub

Sub M_�^�X�N�̍폜()
  Call Library.startScript
  Call init.setting
  
  Call Task.�^�X�N�̍폜
  
  Call Library.endScript(True)
End Sub

'�\�����[�h----------------------------------------------------------------------------------------
Sub M_�^�X�N�\��_�W��()
  Call Library.startScript
  
  Call init.setting
  If setVal("debugMode") <> "develop" Then
    mainSheet.Visible = True
    TeamsPlannerSheet.Visible = xlSheetVeryHidden
  End If
  
  Call init.setting(True)
  Call WBS_Option.�^�X�N�\��_�W��
  Call WBS_Option.setLineColor
  
  Application.Goto Reference:=Range("A6"), Scroll:=True
  Call Library.endScript

End Sub

Sub M_�^�X�N�\��_�^�X�N()
  Call Library.startScript
  Call init.setting(True)
  
  Call WBS_Option.viewTask
  Call WBS_Option.setLineColor
  
  Call Library.endScript
End Sub

Sub M_�^�X�N�\��_�`�[���v�����i�[()
  Call Library.startScript
  Call init.setting(True)
  
  Call WBS_Option.�^�X�N�\��_�`�[���v�����i�[
  Call WBS_Option.setLineColor
  
  Application.Goto Reference:=Range("A6"), Scroll:=True
  
  Call Library.endScript
End Sub


Sub M_�^�X�N�ɃX�N���[��()
  Call Library.startScript
  Call init.setting
  
  Call WBS_Option.�^�X�N�ɃX�N���[��
  Call Library.endScript
End Sub

Sub M_�^�C�����C���ɒǉ�()
  Call Library.startScript
  Call init.setting
  
  Call Chart.�^�C�����C���ɒǉ�(ActiveCell.row)
  Call Library.endScript(True)
End Sub
'**************************************************************************************************
' * �K���g�`���[�g
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'�N���A--------------------------------------------------------------------------------------------
Sub M_�K���g�`���[�g�N���A()
Attribute M_�K���g�`���[�g�N���A.VB_ProcData.VB_Invoke_Func = "D\n14"
  Call Library.startScript
  Call Chart.�K���g�`���[�g�폜
  Call Library.endScript
End Sub

'�����̂�------------------------------------------------------------------------------------------
Sub M_�K���g�`���[�g�����̂�()
Attribute M_�K���g�`���[�g�����̂�.VB_ProcData.VB_Invoke_Func = "A\n14"
  Call init.setting
  Call Library.startScript
  Call ProgressBar.showStart
  Call Library.showDebugForm("�K���g�`���[�g����", "�����J�n")
  
  Call Chart.�K���g�`���[�g����
  
  Call Library.showDebugForm("�K���g�`���[�g����", "��������")
  Call ProgressBar.showEnd
  Call Library.endScript(True)
End Sub


'����----------------------------------------------------------------------------------------------
Sub M_�K���g�`���[�g����()
Attribute M_�K���g�`���[�g����.VB_ProcData.VB_Invoke_Func = "t\n14"
  Call init.setting
  
  Call Library.startScript
  Call ProgressBar.showStart
  
  Call Check.�^�X�N���X�g�m�F
  Call Chart.�K���g�`���[�g����
  
  Call ProgressBar.showEnd
  Call Library.endScript(True)
  Application.EnableEvents = True
End Sub



'�Z���^�[----------------------------------------------------------------------------------------------
Sub M_�Z���^�[()
Attribute M_�Z���^�[.VB_ProcData.VB_Invoke_Func = " \n14"

  Call init.setting
  Call Library.startScript
  Call ProgressBar.showStart
  Call Library.showDebugForm("�Z���^�[�ֈړ�", "�����J�n")
  
  Call Chart.�Z���^�[
  
  Call Library.showDebugForm("�Z���^�[�ֈړ�", "��������")
  Call ProgressBar.showEnd
  Call Library.endScript
End Sub


'**************************************************************************************************
' * import
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'Excel�t�@�C��-------------------------------------------------------------------------------------
Sub M_Excel�C���|�[�g()
  
  Call Library.startScript
  Call Library.showDebugForm("�t�@�C���C���|�[�g", "�����J�n")
  If MsgBox("�f�[�^���폜���܂�", vbYesNo + vbExclamation) = vbYes Then
    Call WBS_Option.clearAll
  Else
    Call WBS_Option.clearCalendar
  End If
  Call ProgressBar.showStart
  
  
  Call import.�t�@�C���C���|�[�g
  Call Calendar.�����ݒ�
  Call import.makeCalendar
  
  If setVal("lineColorFlg") = "True" Then
    setVal("lineColorFlg") = False
    Call WBS_Option.setLineColor
  Else
  End If
  
  
  Call ProgressBar.showEnd
  Call Library.endScript(True)
  
  Call WBS_Option.saveAndRefresh
  
  Err.Clear
  Call Library.showNotice(200, "�C���|�[�g")
End Sub
















