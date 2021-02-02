Attribute VB_Name = "init"
'���[�N�u�b�N�p�ϐ�==============================
Public ThisBook As Workbook
Public targetBook As Workbook

'���[�N�V�[�g�p�ϐ�==============================
Public sheetNotice As Worksheet
Public sheetHelp As Worksheet
Public sheetSetting As Worksheet
Public tmpSheet As Worksheet
Public sheetMain As Worksheet
Public sheetTeamsPlanner As Worksheet

'�O���[�o���ϐ�==================================
Public Const thisAppName = "WorkBreakdownStructure4Excel"
Public Const thisAppVersion = "0.0.3.0"


Public setVal As Collection
Public getVal As Collection
Public memberColor As Object

Public sheetMainName As String
Public sheetTeamsPlannerName As String

'���W�X�g���o�^�p�T�u�L�[
Public Const RegistryKey As String = "WBS"
Public Const RegistrySubKey As String = "Main"
Public Const RibbonTabName As String = "WBSTab"
Public RegistryRibbonName As String


'���O�t�@�C��
Public logFile As String

'�K���g�`���[�g�I��
Public selectShapesName(0) As Variant
Public changeShapesName As String


Public deleteFlg As Boolean

'**************************************************************************************************
' * �ݒ�N���A
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function clearSetting()
  Set sheetHelp = Nothing
  Set sheetNotice = Nothing
  Set sheetSetting = Nothing
  Set sheetMain = Nothing
  Set tmpSheet = Nothing
  Set sheetTeamsPlanner = Nothing
  
  Set setVal = Nothing
  Set memberColor = Nothing

  
End Function
'**************************************************************************************************
' * �ݒ�
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function setting(Optional reCheckFlg As Boolean)
  Dim line As Long
  
'  On Error GoTo catchError
  
  If logFile <> "" And reCheckFlg <> True Then
    Exit Function
  End If
  
  If ThisWorkbook.Worksheets("�ݒ�").Range("B3") = "develop" Then
    ThisWorkbook.Save
  End If
  
Label_reset:
  
  '�u�b�N�̐ݒ�------------------------------------------------------------------------------------
  Set ThisBook = ThisWorkbook
  ThisBook.Activate
  
  '���[�N�V�[�g���̐ݒ�----------------------------------------------------------------------------
  sheetMainName = "���C��"
  sheetTeamsPlannerName = "�`�[���v�����i�["
  Set sheetHelp = ThisBook.Worksheets("Help")
  Set sheetNotice = ThisBook.Worksheets("Notice")
  Set sheetSetting = ThisBook.Worksheets("�ݒ�")
  Set tmpSheet = ThisBook.Worksheets("Tmp")
  
  If sheetSetting.Range("B9") = "Normal" Then
    Set sheetMain = ThisBook.Worksheets(sheetMainName)
    Set sheetTeamsPlanner = ThisBook.Worksheets(sheetTeamsPlannerName)
  ElseIf sheetSetting.Range("B9") = "TeamsPlanner" Then
    Set sheetMain = ThisBook.Worksheets(sheetTeamsPlannerName)
    Set sheetTeamsPlanner = ThisBook.Worksheets(sheetMainName)
  End If
  
  Set setVal = New Collection
  Set memberColor = CreateObject("Scripting.Dictionary")
  
  
  '�����l�ݒ�--------------------------------------------------------------------------------------
  '���ԁA���
  Select Case True
    Case sheetSetting.Range("B6") = ""
      sheetSetting.Range("B6") = Format(DateAdd("d", 0, Date), "yyyy/mm/dd")
    
    Case sheetSetting.Range("B7") = ""
      sheetSetting.Range("A7") = Format(DateAdd("d", 60, Date), "yyyy/mm/dd")
    
    Case sheetSetting.Range("B8") = ""
      sheetSetting.Range("B8") = Format(DateAdd("d", 0, Date), "yyyy/mm/dd")
  End Select
  
  If sheetSetting.Range("B4") = "CD��" Then
    sheetSetting.Range("B8") = Format(Date, "yyyy/mm/dd")
  End If
  
  'VBA�p�̐ݒ�l�擾-------------------------------------------------------------------------------
  With setVal
    For line = 3 To sheetSetting.Cells(Rows.count, 1).End(xlUp).row
      If sheetSetting.Range("A" & line) <> "" Then
       .Add item:=sheetSetting.Range("B" & line), Key:=sheetSetting.Range("A" & line)
      End If
    Next
    For line = 3 To sheetSetting.Cells(Rows.count, 4).End(xlUp).row
      If sheetSetting.Range("D" & line) <> "" Then
       .Add item:=sheetSetting.Range("E" & line), Key:=sheetSetting.Range("D" & line)
      End If
    Next
  End With
  
  '�V���[�g�J�b�g�L�[�ݒ�--------------------------------------------------------------------------
  With setVal
    For line = 3 To sheetSetting.Cells(Rows.count, 7).End(xlUp).row
      .Add item:=sheetSetting.Range(setVal("cell_ShortcutKey") & line), Key:=sheetSetting.Range(setVal("cell_ShortcutFuncName") & line)
    Next
  End With
  

  '�S���ҐF�ǂݍ���--------------------------------------------------------------------------------
  For line = 3 To sheetSetting.Cells(Rows.count, 11).End(xlUp).row
    If sheetSetting.Range("K" & line).Value <> "" Then
      memberColor.Add sheetSetting.Range("K" & line).Value, sheetSetting.Range("K" & line).Interior.Color
    End If
  Next line


  '�t�@�C���֘A�ݒ�--------------------------------------------------------------------------------
  logFile = ThisBook.Path & "\ExcelMacro.log"
  
  
  '���W�X�g���֘A�ݒ�------------------------------------------------------------------------------
  RegistryRibbonName = "RP_" & ActiveWorkbook.Name
  
  
  
  
'  If reCheckFlg = True Then
'    Call Check.���ڗ�`�F�b�N
'    reCheckFlg = False
'    Call clearSetting
'
'    GoTo Label_reset
'  End If
  
  Call ���O��`
  Exit Function
  
'�G���[������--------------------------------------------------------------------------------------
catchError:
  logFile = ""
'  Call Library.showNotice(Err.Number, Err.Description, True)
  
  GoTo Label_reset
  
End Function

'**************************************************************************************************
' * �x���ݒ�
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function chkHollyday(chkDate As Date, HollydayName As String)
  Dim line As Long, endLine As Long
  Dim strFilMessage() As Date
  
  '�x������
  Call GetHollyday(CDate(chkDate), HollydayName)
  
  '�y���𔻒�
  If HollydayName = "" Then
    If Weekday(chkDate) = vbSunday Then
      HollydayName = "Sunday"
    ElseIf Weekday(chkDate) = vbSaturday Then
      HollydayName = "Saturday"
    End If
  End If
End Function


'**************************************************************************************************
' * ���O��`
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function ���O��`()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  Dim Name As Object
  
  On Error GoTo catchError

  For Each Name In Names
    If Name.Visible = False Then
      Name.Visible = True
    End If
    If Not Name.Name Like "*!Print_Area" And Not Name.Name Like "*!Print_Titles" Then
      Name.Delete
    End If
  Next
  
  'VBA�p�̐ݒ�
  For line = 3 To sheetSetting.Range("B5")
    If sheetSetting.Range("A" & line) <> "" Then
      sheetSetting.Range(setVal("cell_LevelInfo") & line).Name = sheetSetting.Range("A" & line)
    End If
  Next
  
  '�V���[�g�J�b�g�L�[�̐ݒ�
  endLine = sheetSetting.Cells(Rows.count, Library.getColumnNo(setVal("cell_ShortcutFuncName"))).End(xlUp).row
  For line = 3 To endLine
    If sheetSetting.Range(setVal("cell_ShortcutFuncName") & line) <> "" Then
      sheetSetting.Range(setVal("cell_ShortcutKey") & line).Name = sheetSetting.Range(setVal("cell_ShortcutFuncName") & line)
    End If
  Next
  
  
  endLine = sheetSetting.Cells(Rows.count, 11).End(xlUp).row
  If setVal("workMode") = "CD��" Then
    sheetSetting.Range(setVal("cell_AssignorList") & "3:" & setVal("cell_AssignorList") & endLine).Name = "�S����"
  Else
    sheetSetting.Range(setVal("cell_AssignorList") & "4:" & setVal("cell_AssignorList") & endLine).Name = "�S����"
  End If
  endLine = sheetSetting.Cells(Rows.count, 17).End(xlUp).row
  sheetSetting.Range(setVal("cell_CompanyHoliday") & "3:" & setVal("cell_CompanyHoliday") & endLine).Name = "�x�����X�g"

  Exit Function
'�G���[������--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(Err.Number, Err.Description, True)
  
End Function


'**************************************************************************************************
' * �V�[�g�̕\��/��\��
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function noDispSheet()

  Call init.setting
  tmpSheet.Visible = xlSheetVeryHidden
  sheetNotice.Visible = xlSheetVeryHidden
  Worksheets("�T���v��").Visible = xlSheetVeryHidden
  sheetTeamsPlanner.Visible = xlSheetVeryHidden
  
  Worksheets(sheetMainName).Select
End Function



Function dispSheet()

  Call init.setting
  Worksheets("Help").Visible = True
  Worksheets("Tmp").Visible = True
  Worksheets("Notice").Visible = True
  Worksheets("�ݒ�").Visible = True
  Worksheets("�T���v��").Visible = True
  
  Worksheets(sheetTeamsPlannerName).Visible = True
  Worksheets(sheetMainName).Visible = True
  
  Worksheets(sheetMainName).Select
  
End Function

