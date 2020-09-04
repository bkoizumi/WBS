Attribute VB_Name = "Calendar"


'**************************************************************************************************
' * カレンダー書式設定
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function 書式設定()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long

'  Call init.setting
'  mainSheet.Select
  
  Columns("A:A").ColumnWidth = 4
  Columns("B:B").ColumnWidth = 3
  Columns("C:C").ColumnWidth = 3
  

    '作業項目
    Columns(setVal("cell_TaskArea") & ":" & setVal("cell_TaskArea")).ColumnWidth = 40
    
    '予定日
    Columns(setVal("cell_PlanStart") & ":" & setVal("cell_PlanEnd")).ColumnWidth = 6
    Columns(setVal("cell_PlanStart") & ":" & setVal("cell_PlanEnd")).NumberFormatLocal = "m/d;@"
    
    '担当者
    Columns(setVal("cell_Assign") & ":" & setVal("cell_Assign")).ColumnWidth = 10
    
    '先行タスク
    Columns(setVal("cell_Task") & ":" & setVal("cell_Task")).ColumnWidth = 10
    
    '実績日
    Columns(setVal("cell_AchievementStart") & ":" & setVal("cell_AchievementEnd")).ColumnWidth = 6
    Columns(setVal("cell_AchievementStart") & ":" & setVal("cell_AchievementEnd")).NumberFormatLocal = "m/d;@"
    
    '進捗率
    Columns(setVal("cell_ProgressLast") & ":" & setVal("cell_Progress")).ColumnWidth = 6
    Columns(setVal("cell_ProgressLast") & ":" & setVal("cell_Progress")).NumberFormatLocal = "0_ ;[赤]-0 "
    
    'タスク情報
    Columns(setVal("cell_TaskInfoP") & ":" & setVal("cell_TaskInfoC")).ColumnWidth = 8
    Columns(setVal("cell_TaskInfoC") & ":" & setVal("cell_WorkLoadA")).NumberFormatLocal = "@"
    
    
    '作業工数
    Columns(setVal("cell_WorkLoadP") & ":" & setVal("cell_WorkLoadA")).ColumnWidth = 7
    Columns(setVal("cell_WorkLoadP") & ":" & setVal("cell_WorkLoadA")).NumberFormatLocal = "0.0_ ;[赤]-0.0 "
    
    
    '遅早工数
    Columns(setVal("cell_LateOrEarly") & ":" & setVal("cell_LateOrEarly")).ColumnWidth = 10
    Columns(setVal("cell_LateOrEarly") & ":" & setVal("cell_LateOrEarly")).NumberFormatLocal = "0.00_ ;[赤]-0.00 "
    
    '備考
    Columns(setVal("cell_Note") & ":" & setVal("cell_Note")).ColumnWidth = 40

  
  'カレンダー部分
  With Columns(setVal("calendarStartCol") & ":XFD")
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    .ColumnWidth = 2.5
  End With
  
  Cells.RowHeight = 20
  Rows("4:4").RowHeight = 20
  Rows("5:5").RowHeight = 35
  Range("W3:XFD3").NumberFormatLocal = "m""月"""
  Range("W4:XFD4").NumberFormatLocal = "d"

End Function


'**************************************************************************************************
' * カレンダー削除
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function clearCalendar()

  Call init.setting
  Columns(setVal("calendarStartCol") & ":XFD").Delete Shift:=xlToLeft
  Range("I5:" & setVal("cell_Note") & 5).ClearContents
  setSheet.Range("O3:P" & setSheet.Cells(Rows.count, 15).End(xlUp).row + 1).ClearContents
  
  
  Application.Goto Reference:=Range("A6"), Scroll:=True
  
End Function


'**************************************************************************************************
' * カレンダー生成
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function makeCalendar()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long, endRowLine As Long
  Dim today As Date
  Dim HollydayName As String
  
  
  Call clearCalendar
  Call WBS_Option.選択シート確認
  
  today = setVal("startDay")
  line = Range(setVal("calendarStartCol") & "1").Column
  
  Do While today <= setVal("endDay")
    Cells(4, line) = today
    
    If Format(today, "d") = 1 Or line = Library.getColumnNo(setVal("calendarStartCol")) Then
      Cells(3, line) = today
      Cells(3, line).NumberFormatLocal = "m""月"""
      Range(Cells(3, line), Cells(4, line)).Select
      Call 罫線.月初

    ElseIf DateSerial(Format(today, "yyyy"), Format(today, "m") + 1, 1) - 1 = today Then
      Cells(4, line).Select
      Call 罫線.月末
      Cells(3, line - 1).Select
      Range(Selection, Selection.End(xlToLeft)).Merge

    Else
      Cells(4, line).Select
      Call 罫線.月中
    End If
    
    '休日の設定----------------------------------
    Call init.chkHollyday(today, HollydayName)
    Select Case HollydayName
      Case "Saturday"
        Cells(4, line).Interior.Color = setVal("SaturdayColor")
        
      Case "Sunday"
        Cells(4, line).Interior.Color = setVal("SundayColor")
      Case ""
      Case Else
        If HollydayName <> "会社指定休日" Then
          Cells(4, line).Interior.Color = setVal("SundayColor")
        Else
          Cells(4, line).Interior.Color = setVal("CompanyHolidayColor")
        End If
        '休日名をコメントに
        If TypeName(Cells(4, line).Comment) = "Nothing" Then
          Cells(4, line).AddComment HollydayName
        Else
          Cells(4, line).ClearComments
          Cells(4, line).AddComment HollydayName
        End If
        
        '期間中の休日リスト設定
        endRowLine = setSheet.Cells(Rows.count, 15).End(xlUp).row + 1
        setSheet.Range(setVal("cell_HolidayListDay") & endRowLine) = today
        setSheet.Range(setVal("cell_HolidayListName") & endRowLine) = HollydayName
    End Select
    
    '書式設定
    Cells(3, line).NumberFormatLocal = "m""月"""
    Cells(4, line).NumberFormatLocal = "d"
    
    line = line + 1
    today = today + 1
  Loop
  
  'カレンダーの最終列設定
  Range("calendarEndCol") = Library.getColumnName(line - 1)
  
  
  Range(Cells(4, Library.getColumnNo(setVal("calendarStartCol"))), Cells(4, line - 1)).Select
  Call Library.resetComment
    
  Cells(3, line - 1).Select
  Range(Selection, Selection.End(xlToLeft)).Merge
  Range(Cells(3, line - 1), Cells(6, line - 1)).Select
  Call 罫線.最終日
  
  Range(setVal("cell_Note") & "1:" & setVal("cell_Note") & 6).Select
  Call 罫線.二重線
  Range(Cells(4, Library.getColumnNo(setVal("calendarStartCol"))), Cells(4, line - 1)).Copy
  Range(Cells(5, Library.getColumnNo(setVal("calendarStartCol"))), Cells(6, line - 1)).Select
  Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
  Range(Cells(3, Library.getColumnNo(setVal("calendarStartCol"))), Cells(6, line - 1)).Select
  Call 罫線.横線

  endLine = Cells(Rows.count, 1).End(xlUp).row
  If endLine = 5 Then
    endLine = 25
  End If
  Rows("6:" & endLine).Select
  Selection.RowHeight = 20
    
  Range("A6:B6").Select
  Selection.Style = "数値"

  Call 書式設定
  Call 行書式コピー(6, endLine)
  
  If ActiveSheet.Name = "メイン" Then
    Call init.名前定義
  End If

  
  Application.Goto Reference:=Range("A1"), Scroll:=True
End Function


'**************************************************************************************************
' * 行書式コピー
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function 行書式コピー(startLine As Long, endLine As Long)
  Dim line As Long
  Dim taskLevel As Long
  Dim taskLevelRange As Range
  
'  On Error GoTo catchError
  
 
  'タスクが記載されている場合、タスクレベルを値としてコピー
  Application.CalculateFull
  If Range(setVal("cell_TaskArea") & startLine) <> "" Then
    Range("B" & startLine & ":B" & endLine).Copy
    Range("B" & startLine & ":B" & endLine).PasteSpecial Paste:=xlPasteValues
  End If
  
  '書式のコピー＆ペースト
  Rows("4:4").Copy
  Rows(startLine & ":" & endLine).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
  Application.CutCopyMode = False
  
  'タスクレベルの設定
  For line = 6 To endLine
    If Range(setVal("cell_TaskArea") & line) <> "" Then
      taskLevel = Range("B" & line) - 1
      If taskLevel > 0 Then
        Range(setVal("cell_TaskArea") & line).InsertIndent taskLevel
      End If
    End If
    
    Range("A" & line).FormulaR1C1 = "=ROW()-5"
    Set taskLevelRange = Range(setVal("cell_TaskArea") & line)
    Range("B" & line).FormulaR1C1 = "=getIndentLevel(" & taskLevelRange.Address(ReferenceStyle:=xlR1C1) & ")"
    Set taskLevelRange = Nothing
  Next
  
  With Range(setVal("cell_Assign") & startLine & ":" & setVal("cell_Assign") & endLine).Validation
      .Delete
      .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
      xlBetween, Formula1:="=担当者"
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
  
  With Range(setVal("cell_TaskArea") & startLine & ":" & setVal("cell_TaskArea") & endLine).Validation
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
  

  Exit Function
'エラー発生時=====================================================================================
catchError:
  Call Library.endScript
End Function
