Attribute VB_Name = "Calendar"


'**************************************************************************************************
' * �J�����_�[�����ݒ�
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function �����ݒ�()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  
  Columns("A:A").ColumnWidth = 4
  Columns(setVal("cell_LevelInfo") & ":" & setVal("cell_LineInfo")).ColumnWidth = 3
    
  '��ƍ���
  Columns(setVal("cell_TaskArea") & ":" & setVal("cell_TaskArea")).ColumnWidth = 40
  
  '�\���
  Columns(setVal("cell_PlanStart") & ":" & setVal("cell_PlanEnd")).ColumnWidth = 6
  Columns(setVal("cell_PlanStart") & ":" & setVal("cell_PlanEnd")).NumberFormatLocal = "m/d;@"
  
  '�S����
  Columns(setVal("cell_Assign") & ":" & setVal("cell_Assign")).ColumnWidth = 10
  
  '�^�X�N�z��
  Columns(setVal("cell_TaskAllocation") & ":" & setVal("cell_TaskAllocation")).ColumnWidth = 10
  
  '��s�^�X�N
  Columns(setVal("cell_Task") & ":" & setVal("cell_Task")).ColumnWidth = 10
  
  '���ѓ�
  Columns(setVal("cell_AchievementStart") & ":" & setVal("cell_AchievementEnd")).ColumnWidth = 6
  Columns(setVal("cell_AchievementStart") & ":" & setVal("cell_AchievementEnd")).NumberFormatLocal = "m/d;@"
  
  '�i����
  Columns(setVal("cell_ProgressLast") & ":" & setVal("cell_Progress")).ColumnWidth = 6
  Columns(setVal("cell_ProgressLast") & ":" & setVal("cell_Progress")).NumberFormatLocal = "0_ ;[��]-0 "
  
  '�^�X�N���
  Columns(setVal("cell_TaskInfoP") & ":" & setVal("cell_TaskInfoC")).ColumnWidth = 8
  Columns(setVal("cell_TaskInfoC") & ":" & setVal("cell_WorkLoadA")).NumberFormatLocal = "@"
  
  
  '��ƍH��
  Columns(setVal("cell_WorkLoadP") & ":" & setVal("cell_WorkLoadA")).ColumnWidth = 7
  Columns(setVal("cell_WorkLoadP") & ":" & setVal("cell_WorkLoadA")).NumberFormatLocal = "0.0_ ;[��]-0.0 "
  
  
  '�x���H��
  Columns(setVal("cell_LateOrEarly") & ":" & setVal("cell_LateOrEarly")).ColumnWidth = 8
  Columns(setVal("cell_LateOrEarly") & ":" & setVal("cell_LateOrEarly")).NumberFormatLocal = "0.00_ ;[��]-0.00 "
  
  '���l
  Columns(setVal("cell_Note") & ":" & setVal("cell_Note")).ColumnWidth = 30

  
  '�J�����_�[����
  With Columns(setVal("calendarStartCol") & ":XFD")
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    .ColumnWidth = 2.5
  End With
  
  Cells.RowHeight = 20
  Rows("5:5").RowHeight = 35
  Range(setVal("calendarStartCol") & "3:XFD3").NumberFormatLocal = "m""��"""
  Range(setVal("calendarStartCol") & "4:XFD4").NumberFormatLocal = "d"

End Function


'**************************************************************************************************
' * �J�����_�[�폜
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function clearCalendar()

  Call init.setting
  Columns(setVal("calendarStartCol") & ":XFD").Delete Shift:=xlToLeft
  Range("I5:" & setVal("cell_Note") & 5).ClearContents
  sheetSetting.Range(setVal("cell_HolidayListDay") & "3:" & setVal("cell_HolidayListName") & sheetSetting.Cells(Rows.count, Library.getColumnNo(setVal("cell_HolidayListDay"))).End(xlUp).row + 1).ClearContents
  
  
  Application.Goto Reference:=Range("A6"), Scroll:=True
  
End Function


'**************************************************************************************************
' * �J�����_�[����
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function makeCalendar()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long, endRowLine As Long
  Dim today As Date
  Dim HollydayName As String
  
  Call init.setting
  Call WBS_Option.�I���V�[�g�m�F
  Call clearCalendar
  
  today = setVal("startDay")
  line = Range(setVal("calendarStartCol") & 1).Column
  Call Library.showDebugForm("�J�����_�[����", "")
  
  
  Do While today <= setVal("endDay")
    Cells(4, line) = today
    Call ctl_ProgressBar.showCount("�J�����_�[����", 0, 100, "")
    
    If Format(today, "d") = 1 Or line = Library.getColumnNo(setVal("calendarStartCol")) Then
      Cells(3, line) = today
      Cells(3, line).NumberFormatLocal = "m""��"""
      Range(Cells(3, line), Cells(4, line)).Select
      Call �r��.����

    ElseIf DateSerial(Format(today, "yyyy"), Format(today, "m") + 1, 1) - 1 = today Or today = setVal("endDay") Then
      Cells(4, line).Select
      Call �r��.����
      Cells(3, line).Select
      Range(Selection, Selection.End(xlToLeft)).Merge

    Else
      Cells(4, line).Select
      Call �r��.����
    End If
    
    '�x���̐ݒ�==================================
    Call init.chkHollyday(today, HollydayName)
    Select Case HollydayName
      Case "Saturday"
        Cells(4, line).Interior.Color = setVal("SaturdayColor")
        
      Case "Sunday"
        Cells(4, line).Interior.Color = setVal("SundayColor")
      Case ""
      Case Else
        If HollydayName <> "��Ўw��x��" Then
          Cells(4, line).Interior.Color = setVal("SundayColor")
        Else
          Cells(4, line).Interior.Color = setVal("CompanyHolidayColor")
        End If
        '�x�������R�����g��
        If TypeName(Cells(4, line).Comment) = "Nothing" Then
          Cells(4, line).AddComment HollydayName
        Else
          Cells(4, line).ClearComments
          Cells(4, line).AddComment HollydayName
        End If
        
        '���Ԓ��̋x�����X�g�ݒ�
        endRowLine = sheetSetting.Cells(Rows.count, Library.getColumnNo(setVal("cell_HolidayListDay"))).End(xlUp).row + 1
        sheetSetting.Range(setVal("cell_HolidayListDay") & endRowLine) = today
        sheetSetting.Range(setVal("cell_HolidayListName") & endRowLine) = HollydayName
    End Select
    
    '�����ݒ�
    Cells(3, line).NumberFormatLocal = "m""��"""
    Cells(4, line).NumberFormatLocal = "d"
    
    line = line + 1
    today = today + 1
  Loop
  
  '�J�����_�[�̍ŏI��ݒ�
  Range("calendarEndCol") = Library.getColumnName(line - 1)
  
  
  Range(Cells(4, Library.getColumnNo(setVal("calendarStartCol"))), Cells(4, line - 1)).Select
  Call Library.resetComment
    
  Range(Selection, Selection.End(xlToLeft)).Merge
  Range(Cells(3, line - 1), Cells(6, line - 1)).Select
  Call �r��.�ŏI��
  
  Range(setVal("calendarStartCol") & "1:" & setVal("calendarStartCol") & 6).Select
  Call �r��.��d��
  Range(Cells(4, Library.getColumnNo(setVal("calendarStartCol"))), Cells(4, line - 1)).Copy
  Range(Cells(5, Library.getColumnNo(setVal("calendarStartCol"))), Cells(6, line - 1)).Select
  Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
  Range(Cells(3, Library.getColumnNo(setVal("calendarStartCol"))), Cells(6, line - 1)).Select
  Call �r��.����

  endLine = Cells(Rows.count, Library.getColumnNo(setVal("cell_LineInfo"))).End(xlUp).row
  If endLine < 6 And Range(setVal("cell_TaskArea") & 6) = "" Then
    endLine = 25
  End If
  Rows("6:" & endLine).Select
  Selection.RowHeight = 20
    
  Range("A6:B6").Select
  Selection.Style = "���l"

  Call �����ݒ�
  Call �s�����R�s�[(6, endLine)
  
  If ActiveSheet.Name = sheetMainName Then
    Call init.���O��`
  End If

  
  Application.Goto Reference:=Range("A1"), Scroll:=True
End Function


'**************************************************************************************************
' * �s�����R�s�[
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function �s�����R�s�[(startLine As Long, endLine As Long)
  Dim line As Long
  Dim taskLevel As Long
  Dim taskLevelRange As Range
  Dim cell_LineInfo As Long
  
'  On Error GoTo catchError
  
  cell_LineInfo = 1
  '�^�X�N���L�ڂ���Ă���ꍇ�A�^�X�N���x����l�Ƃ��ăR�s�[
  sheetMain.Calculate
  If Range(setVal("cell_TaskArea") & startLine) <> "" Then
    Range("B" & startLine & ":B" & endLine).Copy
    Range("B" & startLine & ":B" & endLine).PasteSpecial Paste:=xlPasteValues
  End If
  
  '�����̃R�s�[���y�[�X�g
  Rows("4:4").Copy
  Rows(startLine & ":" & endLine).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
  Application.CutCopyMode = False
  
  '�^�X�N���x���̐ݒ�
  If ActiveSheet.Name = sheetMainName Then
    For line = 6 To endLine
      If Range(setVal("cell_TaskArea") & line) <> "" Then
        taskLevel = Range(setVal("cell_LevelInfo") & line) - 1
        If taskLevel > 0 Then
          Range(setVal("cell_TaskArea") & line).InsertIndent taskLevel
        End If
      End If
      
      If Range(setVal("cell_Info") & line) <> setVal("TaskInfoStr_Multi") Then
        Range("A" & line) = cell_LineInfo
        cell_LineInfo = cell_LineInfo + 1
      Else
        Range("A" & line) = Range("A" & line - 1)
      End If
      
      Range(setVal("cell_LineInfo") & line).FormulaR1C1 = "=ROW()-5"
      Set taskLevelRange = Range(setVal("cell_TaskArea") & line)
      Range(setVal("cell_LevelInfo") & line).Formula = "=getIndentLevel(" & taskLevelRange.Address(ReferenceStyle:=xlA1, RowAbsolute:=False, ColumnAbsolute:=False) & ")"
      Set taskLevelRange = Nothing
    Next
  End If
  
  With Range(setVal("cell_Assign") & startLine & ":" & setVal("cell_Assign") & endLine).Validation
    .Delete
    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
    xlBetween, Formula1:="=�S����"
    .IgnoreBlank = True
    .InCellDropdown = True
    .InputTitle = ""
    .ErrorTitle = ""
    .InputMessage = ""
    .ErrorMessage = ""
    .IMEMode = xlIMEModeNoControl
    .ShowInput = True
    .ShowError = False
  End With

  With Range(setVal("cell_TaskArea") & startLine & ":" & setVal("cell_TaskArea") & endLine).Validation
    .Delete
    .Add Type:=xlValidateInputOnly, AlertStyle:=xlValidAlertStop, Operator _
    :=xlBetween
    .IgnoreBlank = True
    .InCellDropdown = False
    .InputTitle = ""
    .ErrorTitle = ""
    .InputMessage = ""
    .ErrorMessage = ""
    .IMEMode = xlIMEModeOn
    .ShowInput = True
    .ShowError = True
  End With
  

  Exit Function
'�G���[������--------------------------------------------------------------------------------------
catchError:
  Call Library.endScript
End Function
