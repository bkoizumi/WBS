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
          If mainSheet.Range(setVal("cell_Assign") & line) Like "*" & assignor & "*" Then
            TeamsPlannerSheet.Range("A" & rowLine) = mainSheet.Range("A" & line)
            TeamsPlannerSheet.Range("B" & rowLine) = mainSheet.Range("B" & line).Value
            TeamsPlannerSheet.Range("D" & rowLine) = assignor
            TeamsPlannerSheet.Range("E" & rowLine) = mainSheet.Range(setVal("cell_TaskArea") & line)
            TeamsPlannerSheet.Range("F" & rowLine) = mainSheet.Range(setVal("cell_PlanStart") & line)
            TeamsPlannerSheet.Range("G" & rowLine) = mainSheet.Range(setVal("cell_PlanEnd") & line)
            TeamsPlannerSheet.Range("H" & rowLine) = mainSheet.Range(setVal("cell_AchievementStart") & line)
            TeamsPlannerSheet.Range("I" & rowLine) = mainSheet.Range(setVal("cell_AchievementEnd") & line)
            TeamsPlannerSheet.Range("J" & rowLine) = mainSheet.Range(setVal("cell_ProgressLast") & line)
            TeamsPlannerSheet.Range("K" & rowLine) = mainSheet.Range(setVal("cell_Progress") & line)
            
            For Each taskAllocation In Split(mainSheet.Range(setVal("cell_TaskAllocation") & line), ",")
              If taskAllocation Like "*" & assignor & "*" Then
                allocationRate = Split(taskAllocation, "<>")
                TeamsPlannerSheet.Range("L" & rowLine) = allocationRate(1)
              End If
            Next
            
            TeamsPlannerSheet.Range("M" & rowLine) = mainSheet.Range(setVal("cell_Task") & line)
            TeamsPlannerSheet.Range("N" & rowLine) = mainSheet.Range(setVal("cell_TaskInfoP") & line)
            TeamsPlannerSheet.Range("O" & rowLine) = mainSheet.Range(setVal("cell_TaskInfoC") & line)
            TeamsPlannerSheet.Range("P" & rowLine) = mainSheet.Range(setVal("cell_WorkLoadP") & line)
            TeamsPlannerSheet.Range("Q" & rowLine) = mainSheet.Range(setVal("cell_WorkLoadA") & line)
            TeamsPlannerSheet.Range("R" & rowLine) = mainSheet.Range(setVal("cell_LateOrEarly") & line)
            TeamsPlannerSheet.Range("S" & rowLine) = mainSheet.Range(setVal("cell_Note") & line)
            
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
  setVal("cell_TaskArea") = "E"
  setVal("cell_PlanStart") = "F"
  setVal("cell_PlanEnd") = "G"
  setVal("cell_Assign") = "D"
  setVal("cell_AchievementStart") = "H"
  setVal("cell_AchievementEnd") = "I"
  setVal("cell_ProgressLast") = "J"
  setVal("cell_Progress") = "K"
  
  setVal("cell_TaskAllocation") = "L"
  
  setVal("cell_Task") = "M"
  setVal("cell_TaskInfoP") = "N"
  setVal("cell_TaskInfoC") = "O"
  setVal("cell_WorkLoadP") = "P"
  setVal("cell_WorkLoadA") = "Q"
  setVal("cell_LateOrEarly") = "R"
  setVal("cell_Note") = "S"
  
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
    Range("B" & line) = Cells(line, getColumnNo(setVal("cell_TaskArea"))).IndentLevel + 1
  Next

  Call Library.endScript(True)

  Exit Function

'エラー発生時=====================================================================================
catchError:
  Call Library.showNotice(Err.Number, Err.Description, True)
End Function
