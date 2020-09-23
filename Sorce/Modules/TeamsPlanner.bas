Attribute VB_Name = "TeamsPlanner"
'**************************************************************************************************
' * データ移行
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function データ移行()
  Dim line As Long, endLine As Long, startLine As Long, endColLine As Long
  Dim assignName As Collection
  
  
'  On Error GoTo catchError

  
  TeamsPlannerSheet.Select
  Call WBS_Option.clearAll
  
  mainSheet.Calculate
  
  '担当者リストの取得
  Call Task.担当者抽出(assignName)

  endLine = mainSheet.Cells(Rows.count, 1).End(xlUp).row
  rowLine = 6
  
    For Each assignor In assignName
      If assignor <> "工程" Then
        For line = 6 To endLine
          If mainSheet.Range(setVal("cell_Assign") & line) Like "*,*" Then
          
          ElseIf mainSheet.Range(setVal("cell_Assign") & line) = assignor Or mainSheet.Range(setVal("cell_Assign") & line) = "" Then
            TeamsPlannerSheet.Range("A" & rowLine) = mainSheet.Range("A" & line)
            TeamsPlannerSheet.Range("B" & rowLine) = mainSheet.Range(setVal("cell_LevelInfo") & line)
            
            TeamsPlannerSheet.Range("C" & rowLine) = mainSheet.Range(setVal("cell_Info") & line)
            TeamsPlannerSheet.Range("D" & rowLine) = mainSheet.Range(setVal("cell_LineInfo") & line)
            
            TeamsPlannerSheet.Range("E" & rowLine) = assignor
            TeamsPlannerSheet.Range("F" & rowLine) = mainSheet.Range(setVal("cell_TaskArea") & line)
            TeamsPlannerSheet.Range("G" & rowLine) = mainSheet.Range(setVal("cell_PlanStart") & line)
            TeamsPlannerSheet.Range("H" & rowLine) = mainSheet.Range(setVal("cell_PlanEnd") & line)
            TeamsPlannerSheet.Range("I" & rowLine) = mainSheet.Range(setVal("cell_AchievementStart") & line)
            TeamsPlannerSheet.Range("J" & rowLine) = mainSheet.Range(setVal("cell_AchievementEnd") & line)
            TeamsPlannerSheet.Range("K" & rowLine) = mainSheet.Range(setVal("cell_ProgressLast") & line)
            TeamsPlannerSheet.Range("L" & rowLine) = mainSheet.Range(setVal("cell_Progress") & line)
            
            TeamsPlannerSheet.Range("M" & rowLine) = mainSheet.Range(setVal("cell_TaskAllocation") & line)
            
            TeamsPlannerSheet.Range("N" & rowLine) = mainSheet.Range(setVal("cell_Task") & line)
            TeamsPlannerSheet.Range("O" & rowLine) = mainSheet.Range(setVal("cell_TaskInfoP") & line)
            TeamsPlannerSheet.Range("P" & rowLine) = mainSheet.Range(setVal("cell_TaskInfoC") & line)
            TeamsPlannerSheet.Range("Q" & rowLine) = mainSheet.Range(setVal("cell_WorkLoadP") & line)
            TeamsPlannerSheet.Range("R" & rowLine) = mainSheet.Range(setVal("cell_WorkLoadA") & line)
            TeamsPlannerSheet.Range("S" & rowLine) = mainSheet.Range(setVal("cell_LateOrEarly") & line)
            TeamsPlannerSheet.Range("T" & rowLine) = mainSheet.Range(setVal("cell_Note") & line)
            
            rowLine = rowLine + 1
          End If
        Next
      End If
    Next

  '書式のコピー＆ペースト
  Rows("4:4").Copy
  Rows("6:" & rowLine - 1).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
  Application.CutCopyMode = False

  'リソースシート用に設定値を変更
  setVal("cell_TaskArea") = "F"
  setVal("cell_PlanStart") = "G"
  setVal("cell_PlanEnd") = "H"
  setVal("cell_Assign") = "E"
  setVal("cell_AchievementStart") = "I"
  setVal("cell_AchievementEnd") = "J"
  setVal("cell_ProgressLast") = "K"
  setVal("cell_Progress") = "L"
  
  setVal("cell_TaskAllocation") = "M"
  
  setVal("cell_Task") = "N"
  setVal("cell_TaskInfoP") = "O"
  setVal("cell_TaskInfoC") = "P"
  setVal("cell_WorkLoadP") = "Q"
  setVal("cell_WorkLoadA") = "R"
  setVal("cell_LateOrEarly") = "S"
  setVal("cell_Note") = "T"
  
  setVal("setLightning") = False
  
  
  Call Calendar.makeCalendar
  Call Chart.ガントチャート生成

  endLine = Cells(Rows.count, 1).End(xlUp).row
  
  '担当者のプルダウン削除
  Range(setVal("cell_Assign") & "6:" & setVal("cell_Assign") & endLine).Validation.Delete
  
  startLine = 6
  Range(setVal("cell_Assign") & endLine + 1) = "sss"
  For line = 7 To endLine + 1
    If Range(setVal("cell_Assign") & line) = Range(setVal("cell_Assign") & line - 1) Then
    Else
      Range(setVal("cell_Assign") & startLine & ":" & setVal("cell_Assign") & line - 1).Select
      Range(setVal("cell_Assign") & startLine & ":" & setVal("cell_Assign") & line - 1).Merge
      Range("A" & startLine & ":" & setVal("calendarEndCol") & line - 1).Borders(xlEdgeBottom).LineStyle = xlDouble
      
      startLine = line
    End If
  Next
  Range(setVal("cell_Assign") & endLine + 1).ClearContents

  'レベルの再設定
  For line = 6 To endLine
    Range(setVal("cell_LevelInfo") & line) = Cells(line, getColumnNo(setVal("cell_TaskArea"))).IndentLevel + 1
  Next

  Call Library.endScript(True)

  Exit Function

'エラー発生時=====================================================================================
catchError:
  Call Library.showNotice(Err.Number, Err.Description, True)
End Function
