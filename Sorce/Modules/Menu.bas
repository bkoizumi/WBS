Attribute VB_Name = "Menu"
Sub M_xxxxxxxxxxxxxxxxx()
  Call Library.startScript
  Call ProgressBar.showStart


  Call ProgressBar.showEnd
  Call Library.endScript
End Sub



'**************************************************************************************************
' * �ݒ�
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Sub M_Help()
  Call init.setting

End Sub



Sub M_�V���[�g�J�b�g�ݒ�()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  
  Call init.setting
  endLine = Cells(Rows.count, 7).End(xlUp).row
  
  '�ݒ������
  For line = 3 To endLine
    Application.MacroOptions Macro:="Menu." & setSheet.Range("G" & line), ShortcutKey:=""
  Next
  
  For line = 3 To endLine
    If setSheet.Range("I" & line) <> "" Then
      Application.MacroOptions Macro:="Menu." & setSheet.Range("G" & line), ShortcutKey:=setSheet.Range("I" & line)
    End If
  Next
  '�C���f���g�̃V���[�g�J�b�g
  Application.OnKey "%{LEFT}", "Menu.M_�C���f���g��"
  Application.OnKey "%{RIGHT}", "Menu.M_�C���f���g��"
  Application.OnKey "%{F1}", "WBS_Option.�\��_�W��"
  Application.OnKey "%{F2}", "WBS_Option.�\��_�K���g�`���[�g"
  
  
End Sub


Sub M_�I�v�V������ʕ\��()
Attribute M_�I�v�V������ʕ\��.VB_ProcData.VB_Invoke_Func = " \n14"
  Call init.setting
  
  Call WBS_Option.�I�v�V������ʕ\��

End Sub


Sub M_�J�����_�[����()
  Call Library.startScript
  Call ProgressBar.showStart
  
  Call Library.showDebugForm("�J�����_�[����", "�����J�n")
  Call Calendar.makeCalendar
  
  Call Library.showDebugForm("�J�����_�[����", "��������")
  Call ProgressBar.showEnd
  Call Library.endScript

End Sub

Sub M_�C���f���g��()
  If ActiveCell.Column = 3 And ActiveSheet.Name = "WBS" Then
    Selection.InsertIndent 1
    Cells(ActiveCell.row, 2).FormulaR1C1 = "=getIndentLevel(ROW())"
  End If
  
End Sub


Sub M_�C���f���g��()
  If ActiveCell.Column = 3 And ActiveSheet.Name = "WBS" Then
    Selection.InsertIndent -1
    Cells(ActiveCell.row, 2).FormulaR1C1 = "=getIndentLevel(ROW())"
  End If
  
End Sub


'**************************************************************************************************
' * ����
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Sub M_�I���s�F�t�ؑ�()
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
' * �^�X�N
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Sub M_�^�X�N�`�F�b�N()
Attribute M_�^�X�N�`�F�b�N.VB_ProcData.VB_Invoke_Func = "C\n14"
  Call Library.startScript
  Call ProgressBar.showStart
  
  Call Library.showDebugForm("�^�X�N�`�F�b�N", "�����J�n")
  Call Check.�^�X�N���X�g�m�F
  
  Call Library.showDebugForm("�^�X�N�`�F�b�N", "��������")
  Call ProgressBar.showEnd
  Call Library.endScript(True)
End Sub


Sub M_�t�B���^�[()
Attribute M_�t�B���^�[.VB_ProcData.VB_Invoke_Func = " \n14"
  Call init.setting
  mainSheet.Select
  
  With FilterForm
    .StartUpPosition = 0
    .top = Application.top + (ActiveWindow.Width / 8)
    .Left = Application.Left + (ActiveWindow.Height / 8)
  End With
  
  FilterForm.Show
End Sub

Sub M_���ׂĕ\��()
Attribute M_���ׂĕ\��.VB_ProcData.VB_Invoke_Func = " \n14"
  Cells.EntireRow.Hidden = False
End Sub


Sub M_�i���R�s�[()
  Call Task.�i���R�s�[
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
  Call Library.startScript
  Call ProgressBar.showStart
  Call Library.showDebugForm("�K���g�`���[�g����", "�����J�n")
  
  Call Check.�^�X�N���X�g�m�F
  Call Chart.�K���g�`���[�g����
  
  Call Library.showDebugForm("�K���g�`���[�g����", "��������")
  Call ProgressBar.showEnd
  Call Library.endScript(True)
  Application.EnableEvents = True
End Sub



'�Z���^�[----------------------------------------------------------------------------------------------
Sub M_�Z���^�[()
Attribute M_�Z���^�[.VB_ProcData.VB_Invoke_Func = " \n14"
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
  
  
  Call ProgressBar.showEnd
  Call Library.endScript
  
  Call WBS_Option.saveAndRefresh
  
  Err.Clear
  Call Library.showNotice(200, "�C���|�[�g")
End Sub


















