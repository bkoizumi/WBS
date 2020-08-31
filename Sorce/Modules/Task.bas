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
  
  endLine = setSheet.Cells(Rows.count, 18).End(xlUp).row
  count = count + 1
  For line = 3 To endLine
    If setSheet.Range("R" & line) <> "" Then
      With taskList
        .Add item:=setSheet.Range("R" & line).Value, Key:=str(count)
      End With
      count = count + 1
    End If
  Next
  Exit Function
  
'エラー発生時=====================================================================================
catchError:

End Function


'**************************************************************************************************
' * 担当者抽出
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function 担当者抽出(memberList As Collection)
  Dim line As Long, endLine As Long, count As Long

'  On Error GoTo catchError

  Call init.setting
  Set memberList = New Collection
  count = 1
  
  endLine = Cells(Rows.count, 1).End(xlUp).row
  With memberList
    .Add item:="工程", Key:=str(count)
  End With
  count = count + 1
  
    For line = 6 To endLine
      If mainSheet.Range(setVal("cell_AssignP") & line).Value <> "" And isCollection(memberList, mainSheet.Range(setVal("cell_AssignP") & line).Value) = False Then
        With memberList
          .Add item:=mainSheet.Range(setVal("cell_AssignP") & line).Value, Key:=str(count)
        End With
        count = count + 1
      End If
    Next
  Exit Function
'エラー発生時=====================================================================================
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
  
  mainSheet.Select
  Cells.EntireRow.Hidden = False
  endLine = Cells(Rows.count, 1).End(xlUp).row
  
  For line = 6 To endLine
    Call ProgressBar.showCount("担当者フィルター", line, endLine, "")
    
    If Range(setVal("cell_AssignP") & line).Text = filterName Or Range(setVal("cell_AssignA") & line).Text = filterName Then
    Else
      Rows(line & ":" & line).EntireRow.Hidden = True
    End If
  Next
  Call ProgressBar.showEnd
  Call Library.endScript
  Exit Function
'エラー発生時=====================================================================================
catchError:

End Function
  
'**************************************************************************************************
' * タスク名フィルター
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function タスク名フィルター(filterName As String)
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long

'  On Error GoTo catchError

  Unload FilterForm
  Call Library.startScript
  Call ProgressBar.showStart
  Call init.setting
  
  mainSheet.Select
  Cells.EntireRow.Hidden = False
  endLine = Cells(Rows.count, 1).End(xlUp).row
  
  For line = 6 To endLine
    Call ProgressBar.showCount("タスク名フィルター", line, endLine, "")
    
    If Library.TEXTJOIN(" ", True, Range("C" & line & ":" & setVal("cell_TaskAreaEnd") & line)) Like "*" & filterName & "*" Then
    Else
      Rows(line & ":" & line).EntireRow.Hidden = True
    End If
  Next
  Call ProgressBar.showEnd
  Call Library.endScript
  Exit Function
'エラー発生時=====================================================================================
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
  mainSheet.Select
 
  endLine = Cells(Rows.count, 1).End(xlUp).row
  
  Range(setVal("cell_Progress") & 6 & ":" & setVal("cell_Progress") & endLine).Copy
  Range(setVal("cell_ProgressLast") & 6).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
  Application.CutCopyMode = False
  
  Call Library.endScript(True)
  Exit Function
'エラー発生時=====================================================================================
catchError:

  Call Library.showNotice(Err.Number, Err.Description, True)
End Function















