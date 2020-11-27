Attribute VB_Name = "Task"
'**************************************************************************************************
' * タスク名抽出
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function タスク名抽出(taskList As Collection)
  Dim line As Long, endLine As Long, count As Long

'  On Error GoTo catchError

  Call init.setting
  Set taskList = New Collection
  count = 1
  
  endLine = sheetSetting.Cells(Rows.count, Library.getColumnNo(setVal("cell_DataExtract"))).End(xlUp).row
  count = count + 1
  For line = 3 To endLine
    If sheetSetting.Range(setVal("cell_DataExtract") & line) <> "" Then
      With taskList
        .Add item:=sheetSetting.Range(setVal("cell_DataExtract") & line).Value, Key:=str(count)
      End With
      count = count + 1
    End If
  Next
  Exit Function
  
'エラー発生時--------------------------------------------------------------------------------------
catchError:

End Function


'**************************************************************************************************
' * 担当者抽出
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function 担当者抽出(memberList As Collection)
  Dim line As Long, endLine As Long, count As Long
  Dim assignor As String
  
  
'  On Error GoTo catchError

  Call init.setting
  sheetMain.Select
  Set memberList = New Collection
  count = 1

  
  endLine = sheetMain.Cells(Rows.count, 1).End(xlUp).row
  With memberList
    .Add item:="工程", Key:=str(count)
  End With
  count = count + 1
  NoAssignorFlg = False
  
  For line = 6 To endLine
    assignor = sheetMain.Range(setVal("cell_Assign") & line).Value
    If assignor <> "" Then
        If isCollection(memberList, assignor) = False Then
          With memberList
            .Add item:=assignor, Key:=str(count)
          End With
          count = count + 1
        End If
    
    ElseIf assignor = "" And NoAssignorFlg = False Then
      With memberList
        .Add item:="未割り当て", Key:=str(count)
      End With
      count = count + 1
      NoAssignorFlg = True
    End If
  Next





'  For line = 6 To endLine
'    If sheetMain.Range(setVal("cell_Assign") & line).Value <> "" Then
'      For Each assignName In Split(sheetMain.Range(setVal("cell_Assign") & line).Value, ",")
'        assignor = assignName
'        If assignor <> "" And isCollection(memberList, assignor) = False Then
'          With memberList
'            .Add item:=assignor, Key:=str(count)
'          End With
'          count = count + 1
'        End If
'      Next
'    End If
'  Next
  Exit Function
'エラー発生時--------------------------------------------------------------------------------------
catchError:

End Function



Function isCollection(col As Collection, query) As Boolean
  Dim item
  
  For Each item In col
    If item = query Then
      isCollection = True
      Exit Function
    End If
  Next
  isCollection = False
End Function


'**************************************************************************************************
' * 担当者フィルター
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function 担当者フィルター(filterName As String)
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long

'  On Error GoTo catchError

  Unload FilterForm
  Call Library.startScript
  Call ProgressBar.showStart
  Call init.setting
  
  sheetMain.Select
  Cells.EntireRow.Hidden = False
  endLine = Cells(Rows.count, 1).End(xlUp).row
  
  For line = 6 To endLine
    Call ProgressBar.showCount("担当者フィルター", line, endLine, "")
    
    If filterName = "未割り当て" And Range(setVal("cell_Assign") & line).Text = "" Then
    ElseIf Range(setVal("cell_Assign") & line).Text = filterName Or Range(setVal("cell_Assign") & line).Text = filterName Then
    Else
      Rows(line & ":" & line).EntireRow.Hidden = True
    End If
  Next
  Call ProgressBar.showEnd
  Call Library.endScript
  Exit Function
'エラー発生時--------------------------------------------------------------------------------------
catchError:

End Function
  
'**************************************************************************************************
' * タスク名フィルター
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function タスク名フィルター(filterNames As String)
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long

  On Error GoTo catchError
  Call Library.showDebugForm("タスク名フィルター", "開始")

  Unload FilterForm
  Call Library.startScript
  Call init.setting
  
  sheetMain.Select
  
  '非表示行を全て表示
  Cells.EntireRow.Hidden = False
  
  endLine = Cells(Rows.count, 1).End(xlUp).row
  
  For line = 6 To endLine
    DoEvents
    For Each filterName In Split(filterNames, "<>")
      DoEvents
      
      If Range(setVal("cell_Info") & line) = setVal("TaskInfoStr_Multi") Then
        Rows(line & ":" & line).EntireRow.Hidden = True
      ElseIf Range(setVal("cell_TaskArea") & line) Like "*" & filterName & "*" Then
        Rows(line & ":" & line).EntireRow.Hidden = False
        Exit For
      Else
        Rows(line & ":" & line).EntireRow.Hidden = True
      End If
    Next
  Next
  
  
  Call Library.endScript
  Exit Function
'エラー発生時--------------------------------------------------------------------------------------
catchError:

End Function


'**************************************************************************************************
' * 進捗コピー
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function 進捗コピー()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  On Error GoTo catchError
  
  Call Library.startScript
  Call init.setting
  sheetMain.Select
 
  endLine = Cells(Rows.count, 1).End(xlUp).row
  
  Range(setVal("cell_Progress") & 6 & ":" & setVal("cell_Progress") & endLine).Copy
  Range(setVal("cell_ProgressLast") & 6).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
  Application.CutCopyMode = False
  
  Call Library.endScript(True)
  Exit Function
'エラー発生時--------------------------------------------------------------------------------------
catchError:

  Call Library.showNotice(Err.Number, Err.Description, True)
End Function


'**************************************************************************************************
' * 進捗率設定
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function 進捗率設定(progress As Long)
  Dim line As Long
  On Error GoTo catchError
  
  Call Library.startScript
  Call init.setting
  sheetMain.Select
  
  line = ActiveCell.row
  If Range(setVal("cell_TaskArea") & line) <> "" Then
    Range(setVal("cell_Progress") & line) = progress
  End If
  
  Call Library.endScript(True)
  
  Exit Function
'エラー発生時--------------------------------------------------------------------------------------
catchError:

  Call Library.showNotice(Err.Number, Err.Description, True)
End Function


'**************************************************************************************************
' * タスクのリンク設定
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function taskLink()
  Dim line As Long, oldLine As Long
  Dim selectedCells As Range
  Dim targetCell As Range
  Dim newStartDay As Date, newEndDay As Date
  Dim HollydayName As String
    
'  On Error GoTo catchError

  sheetMain.Select
   
  oldLine = 0
  Set selectedCells = Selection
  
  For Each targetCell In selectedCells
    If oldLine <> 0 Then
      If Range(setVal("cell_Task") & targetCell.row) = "" Then
        Range(setVal("cell_Task") & targetCell.row) = oldLine
      Else
        Range(setVal("cell_Task") & targetCell.row) = Range(setVal("cell_Task") & targetCell.row) & "," & oldLine
      End If
      
      '先行タスクの終了日+1を開始日に設定
      newStartDay = Range(setVal("cell_PlanEnd") & oldLine) + 1
      Call init.chkHollyday(newStartDay, HollydayName)
      Do While HollydayName <> ""
        newStartDay = newStartDay + 1
        Call init.chkHollyday(newStartDay, HollydayName)
      Loop
      Range(setVal("cell_PlanStart") & targetCell.row) = newStartDay
      
      '終了日を再設定
      newEndDay = Range(setVal("cell_PlanEnd") & targetCell.row) + Range(setVal("cell_WorkLoadP") & targetCell.row)
      Call init.chkHollyday(newEndDay, HollydayName)
      Do While HollydayName <> ""
        newEndDay = newEndDay + 1
        Call init.chkHollyday(newEndDay, HollydayName)
      Loop
      Range(setVal("cell_PlanEnd") & targetCell.row) = newEndDay
      
'      Range(setVal("cell_Info") & targetCell.row) = setVal("TaskInfoStr_Change")
    End If
    oldLine = targetCell.row
  Next

  Exit Function
'エラー発生時--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(Err.Number, Err.Description, True)
End Function


'**************************************************************************************************
' * タスクのリンク解除
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function taskUnlink()
  Dim line As Long, oldLine As Long
  Dim selectedCells As Range
  Dim targetCell As Range
    
'  On Error GoTo catchError
  Call Library.startScript
  Call init.setting
  sheetMain.Select
   
  oldLine = 0
  Set selectedCells = Selection
  
  For Each targetCell In selectedCells
    Range(setVal("cell_Task") & targetCell.row) = ""
  Next


  Exit Function
'エラー発生時--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(Err.Number, Err.Description, True)
End Function


'**************************************************************************************************
' * タスクの挿入
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function タスクの挿入()
  Dim taskLevelRange As Range
'  On Error GoTo catchError
  

  Rows("4:4").Copy
  Rows(Selection.row & ":" & Selection.row).Insert Shift:=xlDown
  Range("A" & Selection.row).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
  Application.CutCopyMode = False
  Range(setVal("cell_Info") & Selection.row & ":XFD" & Selection.row).ClearContents
  Range(setVal("cell_Info") & Selection.row & ":XFD" & Selection.row).ClearComments
  
  Range("A" & Selection.row) = Range("A" & Selection.row - 1) + 1
  
  
  Set taskLevelRange = Range(setVal("cell_TaskArea") & Selection.row)
  Range(setVal("cell_LevelInfo") & Selection.row).Formula = "=getIndentLevel(" & taskLevelRange.Address(ReferenceStyle:=xlA1, RowAbsolute:=False, ColumnAbsolute:=False) & ")"
  Set taskLevelRange = Nothing

  
  Range(setVal("cell_LineInfo") & Selection.row).FormulaR1C1 = "=ROW()-5"
 
  Call WBS_Option.行番号再設定

  Exit Function
'エラー発生時--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(Err.Number, Err.Description, True)
End Function


'**************************************************************************************************
' * タスクの削除
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function タスクの削除()
  Dim selectedCells As Range

'  On Error GoTo catchError
  Call Library.startScript
  Call init.setting
  sheetMain.Select


  Rows(Selection(1).row & ":" & Selection(Selection.count).row).Delete Shift:=xlUp
  Call WBS_Option.行番号再設定

  Call Library.endScript(True)

  Exit Function
'エラー発生時--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(Err.Number, Err.Description, True)
End Function




