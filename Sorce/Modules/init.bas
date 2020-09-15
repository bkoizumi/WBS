Attribute VB_Name = "init"
'���[�N�u�b�N�p�ϐ�------------------------------
Public ThisBook As Workbook

'���[�N�V�[�g�p�ϐ�------------------------------
Public noticeCodeSheet As Worksheet
Public helpSheet As Worksheet
Public setSheet As Worksheet
Public tmpSheet As Worksheet
Public mainSheet As Worksheet
Public TeamsPlannerSheet As Worksheet



'�O���[�o���ϐ�----------------------------------
Public Const thisAppName = "Excel for Work Breakdown Structure"

Public setVal As Collection
Public getVal As Collection
Public memberColor As Object

Public mainSheetName As String
Public TeamsPlannerSheetName As String

'���W�X�g���o�^�p�T�u�L�[
Public Const RegistrySubKey As String = "WBS"

'���O�t�@�C��
Public logFile As String

'�K���g�`���[�g�I��
Public selectShapesName(0) As Variant
Public changeShapesName As String


'***********************************************************************************************************************************************
' * �ݒ�N���A
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'***********************************************************************************************************************************************
Function clearSetting()
  Set helpSheet = Nothing
  Set noticeCodeSheet = Nothing
  Set setSheet = Nothing
  Set mainSheet = Nothing
  Set tmpSheet = Nothing
  Set TeamsPlannerSheet = Nothing
  
  Set setVal = Nothing
  Set memberColor = Nothing

  
End Function
'***********************************************************************************************************************************************
' * �ݒ�
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'***********************************************************************************************************************************************
Function setting(Optional reCheckFlg As Boolean)
  Dim line As Long
  
  On Error GoTo catchError
  
  If logFile <> "" And setVal("debugMode") = setSheet.Range("B3") And reCheckFlg <> True Then
    Exit Function
  End If

Label_reset:
  Call Library.showDebugForm("setting", CStr(reCheckFlg))

  '�u�b�N�̐ݒ�
  Set ThisBook = ThisWorkbook
  ThisBook.Activate
  
  '���[�N�V�[�g���̐ݒ�
  mainSheetName = "���C��"
  TeamsPlannerSheetName = "�`�[���v�����i�["
  Set helpSheet = ThisBook.Worksheets("Help")
  Set noticeCodeSheet = ThisBook.Worksheets("Notice")
  Set setSheet = ThisBook.Worksheets("�ݒ�")
  Set mainSheet = ThisBook.Worksheets(mainSheetName)
  Set tmpSheet = ThisBook.Worksheets("Tmp")
  Set TeamsPlannerSheet = ThisBook.Worksheets(TeamsPlannerSheetName)
  
  Set setVal = New Collection
  Set memberColor = CreateObject("Scripting.Dictionary")
  
  
  '���ԁA����������͎��̏����l
  Select Case True
    Case setSheet.Range("B7") = ""
      setSheet.Range("B7") = Format(DateAdd("d", 0, Date), "yyyy/mm/dd")
    
    Case setSheet.Range("B8") = ""
      setSheet.Range("A8") = Format(DateAdd("d", 60, Date), "yyyy/mm/dd")
    
    Case setSheet.Range("B9") = ""
      setSheet.Range("B9") = Format(DateAdd("d", 0, Date), "yyyy/mm/dd")
  End Select
  
  '�ݒ�l�̓ǂݍ���
  With setVal
    For line = 3 To setSheet.Cells(Rows.count, 1).End(xlUp).row
      If setSheet.Range("A" & line) <> "" Then
       .Add item:=setSheet.Range("B" & line), Key:=setSheet.Range("A" & line)
      End If
    Next
    For line = 3 To setSheet.Cells(Rows.count, 4).End(xlUp).row
      If setSheet.Range("D" & line) <> "" Then
       .Add item:=setSheet.Range("E" & line), Key:=setSheet.Range("D" & line)
      End If
    Next
  End With
  
  '�V���[�g�J�b�g�L�[�̐ݒ�ǉ�
  With setVal
    For line = 3 To setSheet.Cells(Rows.count, 7).End(xlUp).row
      .Add item:=setSheet.Range(setVal("cell_ShortcutKey") & line), Key:=setSheet.Range(setVal("cell_ShortcutFuncName") & line)
    Next
  End With
  

  '�S���ҐF�ǂݍ���
  For line = 3 To setSheet.Cells(Rows.count, 11).End(xlUp).row
    If setSheet.Range("K" & line).Value <> "" Then
      memberColor.Add setSheet.Range("K" & line).Value, setSheet.Range("K" & line).Interior.Color
    End If
  Next line

'  lineColor = setSheet.Range("E3").Interior.Color
'  SaturdayColor = setSheet.Range("E4").Interior.Color
'  SundayColor = setSheet.Range("E5").Interior.Color
'  CompanyHolidayColor = setSheet.Range("E6").Interior.Color
'
'  taskLevel1Color = setSheet.Range("E7").Interior.Color
'  taskLevel2Color = setSheet.Range("E8").Interior.Color
'  taskLevel3Color = setSheet.Range("E9").Interior.Color
  
  logFile = ThisBook.Path & "\ExcelMacro.log"
  
  If reCheckFlg = True Then
    Call Check.���ڗ�`�F�b�N
    reCheckFlg = False
    Call clearSetting
    
    GoTo Label_reset
  End If
  
  Call ���O��`
  Exit Function
  
'�G���[������=====================================================================================
catchError:
  logFile = ""
'  Set setVal = Nothing
'  Set setVal = New Collection
'
'  With setVal
'    .Add item:="ABC", Key:="debugMode"
'  End With

'  Call Library.showNotice(Err.Number, Err.Description, True)
  
  GoTo Label_reset
  
End Function

'***********************************************************************************************************************************************
' * �x���ݒ�
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'***********************************************************************************************************************************************
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
  
'  On Error GoTo catchError

  For Each Name In Names
    If Name.Visible = False Then
      Name.Visible = True
    End If
    If Not Name.Name Like "*!Print_Area" And Not Name.Name Like "*!Print_Titles" Then
      Name.Delete
    End If
  Next
  
  'VBA�p�̐ݒ�
  For line = 3 To setSheet.Range("B5")
    If setSheet.Range("A" & line) <> "" Then
      setSheet.Range(setVal("cell_LevelInfo") & line).Name = setSheet.Range("A" & line)
    End If
  Next
  
  '�V���[�g�J�b�g�L�[�̐ݒ�
  endLine = setSheet.Cells(Rows.count, Library.getColumnNo(setVal("cell_ShortcutFuncName"))).End(xlUp).row
  For line = 3 To endLine
    If setSheet.Range(setVal("cell_ShortcutFuncName") & line) <> "" Then
      setSheet.Range(setVal("cell_ShortcutKey") & line).Name = setSheet.Range(setVal("cell_ShortcutFuncName") & line)
    End If
  Next
  
  
  endLine = setSheet.Cells(Rows.count, 11).End(xlUp).row
  setSheet.Range(setVal("cell_AssignorList") & "3:" & setVal("cell_AssignorList") & endLine).Name = "�S����"

  endLine = setSheet.Cells(Rows.count, 17).End(xlUp).row
  setSheet.Range(setVal("cell_CompanyHoliday") & "3:" & setVal("cell_CompanyHoliday") & endLine).Name = "�x�����X�g"

  Exit Function
'�G���[������=====================================================================================
catchError:
  Call Library.showNotice(Err.Number, Err.Description, True)
  
End Function


'***********************************************************************************************************************************************
' * �V�[�g�̕\��/��\��
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'***********************************************************************************************************************************************
Function noDispSheet()

  Call init.setting
  'Worksheets("Help").Visible = xlSheetVeryHidden
  Worksheets("Tmp").Visible = xlSheetVeryHidden
  Worksheets("Notice").Visible = xlSheetVeryHidden
'  Worksheets("�ݒ�").Visible = xlSheetVeryHidden
  Worksheets("�T���v��").Visible = xlSheetVeryHidden
  Worksheets(TeamsPlannerSheetName).Visible = xlSheetVeryHidden
  
  Worksheets(mainSheetName).Select
End Function



Function dispSheet()

  Call init.setting
  Worksheets("Help").Visible = True
  Worksheets("Tmp").Visible = True
  Worksheets("Notice").Visible = True
  Worksheets("�ݒ�").Visible = True
  Worksheets("�T���v��").Visible = True
  
  Worksheets(TeamsPlannerSheetName).Visible = True
  Worksheets(mainSheetName).Visible = True
  
  Worksheets(mainSheetName).Select
  
End Function




































