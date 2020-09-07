Attribute VB_Name = "ctl_ribbon"

Private ribbonUI As IRibbonUI ' ���{��
Private rbButton_Visible As Boolean ' �{�^���̕\���^��\��
Private rbButton_Enabled As Boolean ' �{�^���̗L���^����


'**************************************************************************************************
' * ���{�����j���[�ݒ�
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'�ǂݍ��ݎ�����------------------------------------------------------------------------------------
Function onLoad(ribbon As IRibbonUI)
  Set ribbonUI = ribbon
  
  ribbonUI.ActivateTab ("WBSTab")
  
  '���{���̕\�����X�V����
  ribbonUI.Invalidate


End Function






Public Sub getLabel(control As IRibbonControl, ByRef setRibbonVal)
  setRibbonVal = getRibbonMenu(control.ID, 2)
End Sub

Sub getonAction(control As IRibbonControl)
  Dim setRibbonVal As String

  setRibbonVal = getRibbonMenu(control.ID, 3)
  Application.run setRibbonVal

End Sub


'Supertip�̓��I�\��
Public Sub getSupertip(control As IRibbonControl, ByRef setRibbonVal)
  setRibbonVal = getRibbonMenu(control.ID, 5)
End Sub

Public Sub getDescription(control As IRibbonControl, ByRef setRibbonVal)
  setRibbonVal = getRibbonMenu(control.ID, 6)
End Sub

Public Sub getsize(control As IRibbonControl, ByRef setRibbonVal)
  Dim getVal As String
  getVal = getRibbonMenu(control.ID, 4)

  Select Case getVal
    Case "large"
      setRibbonVal = 1
    Case "normal"
      setRibbonVal = 0
    Case Else
  End Select


End Sub

'Ribbon�V�[�g������e���擾
Function getRibbonMenu(menuId As String, offsetVal As Long)

  Dim getString As String
  Dim FoundCell As Range
  Dim ribSheet As Worksheet
  Dim endLine As Long

  On Error GoTo catchError

  Call Library.startScript
  Set ribSheet = ThisWorkbook.Worksheets("Ribbon")

  endLine = ribSheet.Cells(Rows.count, 1).End(xlUp).row

  getRibbonMenu = Application.VLookup(menuId, ribSheet.Range("A2:F" & endLine), offsetVal, False)
  Call Library.endScript


  Exit Function
'�G���[������=====================================================================================
catchError:
  getRibbonMenu = "�G���["

End Function
'**************************************************************************************************
' * ����
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'�I���s�F�t�ؑ�------------------------------------------------------------------------------------
Function setLineColor(control As IRibbonControl)
  Call menu.M_�I���s�F�t�ؑ�
End Function

'**************************************************************************************************
' * �ݒ�
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'Help----------------------------------------------------------------------------------------------
Function Help(control As IRibbonControl)
  Call menu.M_Help
End Function

'�I�v�V����----------------------------------------------------------------------------------------
Function dispOption(control As IRibbonControl)
  Call menu.M_�I�v�V������ʕ\��
End Function

'�S�f�[�^�폜--------------------------------------------------------------------------------------
Function clearAll(control As IRibbonControl)
  Call menu.M_�S�f�[�^�폜
End Function

'����----------------------------------------------------------------------------------------------
Function makeCalendar(control As IRibbonControl)
  Call menu.M_�J�����_�[����
End Function

'�S��ʕ\��----------------------------------------------------------------------------------------
Function DispFullScreen(control As IRibbonControl)
  Call menu.M_�S���
End Function


'**************************************************************************************************
' * WBS
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'�^�X�N���X�g�m�F----------------------------------------------------------------------------------
Function chkTaskList(control As IRibbonControl)
  Call menu.M_�^�X�N�`�F�b�N
End Function

'�t�B���^�[----------------------------------------------------------------------------------------
Function setFilter(control As IRibbonControl)
  Call menu.M_�t�B���^�[
End Function

'���ׂĕ\��----------------------------------------------------------------------------------------
Function dispAllList(control As IRibbonControl)
  Call menu.M_���ׂĕ\��
End Function

'�i���R�s�[----------------------------------------------------------------------------------------
Function copyProgress(control As IRibbonControl)
  Call menu.M_�i���R�s�[
End Function

'�C���f���g----------------------------------------------------------------------------------------
Function taskOutdent(control As IRibbonControl)
  Call menu.M_�C���f���g��
End Function
Function taskIndent(control As IRibbonControl)
  Call menu.M_�C���f���g��
End Function

'�i�����ݒ�----------------------------------------------------------------------------------------
Function progress_0(control As IRibbonControl)
  Call menu.M_�i�����ݒ�(0)
End Function
Function progress_25(control As IRibbonControl)
  Call menu.M_�i�����ݒ�(25)
End Function
Function progress_50(control As IRibbonControl)
  Call menu.M_�i�����ݒ�(50)
End Function
Function progress_75(control As IRibbonControl)
  Call menu.M_�i�����ݒ�(75)
End Function
Function progress_100(control As IRibbonControl)
  Call menu.M_�i�����ݒ�(100)
End Function

'�^�X�N�̃����N------------------------------------------------------------------------------------
Function taskLink(control As IRibbonControl)
  Call menu.M_�^�X�N�̃����N�ݒ�
End Function
Function taskUnlink(control As IRibbonControl)
  Call menu.M_�^�X�N�̃����N����
End Function


'�\�����[�h----------------------------------------------------------------------------------------
Function viewNormal(control As IRibbonControl)
  Call menu.M_�^�X�N�\��_�W��
End Function

Function viewTask(control As IRibbonControl)
  Call menu.M_�^�X�N�\��_�^�X�N
End Function

Function viewResources(control As IRibbonControl)
  Call menu.M_�^�X�N�\��_���\�[�X
End Function

Function viewSetting(control As IRibbonControl)
  Call menu.M_�^�X�N�\��_�ݒ�
End Function








'**************************************************************************************************
' * �K���g�`���[�g
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'�N���A--------------------------------------------------------------------------------------------
Function clearChart(control As IRibbonControl)
  Call menu.M_�K���g�`���[�g�N���A
End Function

'����----------------------------------------------------------------------------------------------
Function makeChart(control As IRibbonControl)
  Call menu.M_�K���g�`���[�g����
End Function

'�Z���^�[----------------------------------------------------------------------------------------------
Function setCenter(control As IRibbonControl)
  Call menu.M_�Z���^�[
End Function
'**************************************************************************************************
' * import
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'Excel�t�@�C��-------------------------------------------------------------------------------------
Function importExcel(control As IRibbonControl)
  Call menu.M_Excel�C���|�[�g
End Function