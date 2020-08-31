Attribute VB_Name = "Calendar"
' *************************************************************************************************
' * �J�����_�[�֘A�֐�
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
' *************************************************************************************************


'**************************************************************************************************
' * �J�����_�[�����ݒ�
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function �����ݒ�()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long

  Call init.setting
  mainSheet.Select
  
  Columns("A:A").ColumnWidth = 4
  Columns("B:B").ColumnWidth = 3
  
  '��ƍ���
  Columns(setVal("cell_TaskArea") & ":" & setVal("cell_TaskArea")).ColumnWidth = 40
  
  '�\���
  Columns(setVal("cell_PlanStart") & ":" & setVal("cell_PlanEnd")).ColumnWidth = 6
  Columns(setVal("cell_PlanStart") & ":" & setVal("cell_PlanEnd")).NumberFormatLocal = "m/d;@"
  
  '�S����
  Columns(setVal("cell_AssignP") & ":" & setVal("cell_AssignP")).ColumnWidth = 7
  
  '�^�X�N
  Columns(setVal("cell_TaskA") & ":" & setVal("cell_TaskB")).ColumnWidth = 5
  
  '���ѓ�
  Columns(setVal("cell_AchievementStart") & ":" & setVal("cell_AchievementEnd")).ColumnWidth = 6
  Columns(setVal("cell_AchievementStart") & ":" & setVal("cell_AchievementEnd")).NumberFormatLocal = "m/d;@"
  
  '�i����
  Columns(setVal("cell_ProgressLast") & ":" & setVal("cell_Progress")).ColumnWidth = 6
  Columns(setVal("cell_ProgressLast") & ":" & setVal("cell_Progress")).NumberFormatLocal = "0_ ;[��]-0 "
  
  
  '��ƍH��
  Columns(setVal("cell_WorkLoadP") & ":" & setVal("cell_WorkLoadA")).ColumnWidth = 7
  Columns(setVal("cell_WorkLoadP") & ":" & setVal("cell_WorkLoadA")).NumberFormatLocal = "0.0_ ;[��]-0.0 "
  
  
  '�x���H��
  Columns(setVal("cell_LateOrEarly") & ":" & setVal("cell_LateOrEarly")).ColumnWidth = 10
  Columns(setVal("cell_LateOrEarly") & ":" & setVal("cell_LateOrEarly")).NumberFormatLocal = "0.00_ ;[��]-0.00 "
  
  '���l
  Columns(setVal("cell_Note") & ":" & setVal("cell_Note")).ColumnWidth = 40
'  Columns(setVal("cell_Note") & ":" & setVal("cell_WorkLoadA")).NumberFormatLocal = "0.0_ ;[��]-0.0 "
  
  
  '�J�����_�[����
  With Columns(setVal("calendarStartCol") & ":XFD")
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    .ColumnWidth = 2.5
  End With
  
  Cells.RowHeight = 20
  Rows("4:4").RowHeight = 20
  Rows("5:5").RowHeight = 35
  Range("W3:XFD3").NumberFormatLocal = "m""��"""
  Range("W4:XFD4").NumberFormatLocal = "d"

End Function


'**************************************************************************************************
' * �J�����_�[�폜
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function clearCalendar()

  Call init.setting
  mainSheet.Select
  Columns(setVal("calendarStartCol") & ":XFD").Delete Shift:=xlToLeft
  Range("I5:" & setVal("cell_Note") & 5).ClearContents
  setSheet.Range("O3:P" & setSheet.Cells(Rows.count, 15).End(xlUp).row + 1).ClearContents
  
  
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
  Call clearCalendar
  mainSheet.Select
  
  today = setVal("startDay")
  line = Range(setVal("calendarStartCol") & "1").Column
  
  Do While today <= setVal("endDay")
    Cells(4, line) = today
    
    If Format(today, "d") = 1 Or line = Library.getColumnNo(setVal("calendarStartCol")) Then
      Cells(3, line) = today
      Cells(3, line).NumberFormatLocal = "m""��"""
      Range(Cells(3, line), Cells(4, line)).Select
      Call �r��.����

    ElseIf DateSerial(Format(today, "yyyy"), Format(today, "m") + 1, 1) - 1 = today Then
      Cells(4, line).Select
      Call �r��.����
      Cells(3, line - 1).Select
      Range(Selection, Selection.End(xlToLeft)).Merge

    Else
      Cells(4, line).Select
      Call �r��.����
    End If
    
    '�x���̐ݒ�----------------------------------
    Call init.chkHollyday(today, HollydayName, True)
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
        endRowLine = setSheet.Cells(Rows.count, 15).End(xlUp).row + 1
        setSheet.Range("O" & endRowLine) = today
        setSheet.Range("P" & endRowLine) = HollydayName
    End Select
    
    '�����ݒ�
    Cells(3, line).NumberFormatLocal = "m""��"""
    Cells(4, line).NumberFormatLocal = "d"
    
    line = line + 1
    today = today + 1
  Loop
  Range(Cells(4, 23), Cells(4, line - 1)).Select
  Call Library.resetComment
    
  Cells(3, line - 1).Select
  Range(Selection, Selection.End(xlToLeft)).Merge
  Range(Cells(3, line - 1), Cells(6, line - 1)).Select
  Call �r��.�ŏI��
  
  Range(setVal("cell_Note") & "1:" & setVal("cell_Note") & 6).Select
  Call �r��.��d��
  Range(Cells(4, Library.getColumnNo(setVal("calendarStartCol"))), Cells(4, line - 1)).Copy
  Range(Cells(5, Library.getColumnNo(setVal("calendarStartCol"))), Cells(6, line - 1)).Select
  Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
  Range(Cells(3, Library.getColumnNo(setVal("calendarStartCol"))), Cells(6, line - 1)).Select
  Call �r��.����

  endLine = Cells(Rows.count, 3).End(xlUp).row
  If endLine = 5 Then
    endLine = 25
  End If
  Rows("6:" & endLine).Select
  Selection.RowHeight = 20
    
  Range("A6:B6").Select
  Selection.Style = "���l"

  Call �����ݒ�
  Call �s�����R�s�[(6, endLine)
  Call init.���O��`
  
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
  
  On Error GoTo catchError
  
  Call init.setting
  Call Library.startScript
  
  '�^�X�N���L�ڂ���Ă���ꍇ�A�^�X�N���x����l�Ƃ��ăR�s�[
  Application.CalculateFull
  If Range("C" & startLine) <> "" Then
    Range("B" & startLine & ":B" & endLine).Copy
    Range("B" & startLine & ":B" & endLine).PasteSpecial Paste:=xlPasteValues
  End If
  
  '�����̃R�s�[���y�[�X�g
  Rows("4:4").Copy
  Rows(startLine & ":" & endLine).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
  
  '�^�X�N���x���̐ݒ�
  For line = 6 To endLine
    taskLevel = Range("B" & line) - 1
    If taskLevel > 0 Then
      Range(setVal("cell_TaskArea") & line).InsertIndent taskLevel
    End If
  Next
  
  
  Range("A" & startLine & ":A" & endLine).FormulaR1C1 = "=ROW()-5"
'  Range("B" & startLine & ":B" & endLine).FormulaR1C1 = _
'      "=IF(RC[1]<>"""",1,IF(RC[2]<>"""",2,IF(RC[3]<>"""",3,IF(RC[4]<>"""",4,IF(RC[5]<>"""",5,IF(RC[6]<>"""",6,""""))))))"
  
  Range("B" & startLine & ":B" & endLine).FormulaR1C1 = "=getIndentLevel(ROW())"
  
  
  
  Application.CutCopyMode = False
  
  With Range(setVal("cell_AssignP") & startLine & ":" & setVal("cell_AssignA") & endLine).Validation
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
      .ShowError = True
  End With
  
  
  With Range("C" & startLine & ":" & setVal("cell_TaskAreaEnd") & endLine).Validation
    .Delete
    .Add Type:=xlValidateInputOnly, AlertStyle:=xlValidAlertStop, Operator _
    :=xlBetween
    .IgnoreBlank = True
    .InCellDropdown = True
    .InputTitle = ""
    .ErrorTitle = ""
    .InputMessage = ""
    .ErrorMessage = ""
    .IMEMode = xlIMEModeOn
    .ShowInput = True
    .ShowError = True
  End With




  
  Call Library.endScript
  

  Exit Function
'�G���[������=====================================================================================
catchError:
  Call Library.endScript
End Function
