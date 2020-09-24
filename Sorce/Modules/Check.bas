Attribute VB_Name = "Check"
'**************************************************************************************************
' * 項目チェック
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function 項目列チェック()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  Dim itemName As String
  Dim defaultLine As Long
  
'  Call init.setting
  defaultLine = setSheet.Range("B4")
  startLine = 4

  setSheet.Range("A" & defaultLine & ":B100").ClearContents
  
'  mainSheet.Select
 
  For colLine = 1 To 20
    If mainSheet.Cells(2, colLine) <> "" Then
      itemName = mainSheet.Cells(2, colLine)
    Else
      GoTo Label_nextFor
    End If
    
    line = setSheet.Cells(Rows.count, 1).End(xlUp).row + 1
    If line < defaultLine Then
      line = defaultLine
    End If
    
    Select Case itemName
      Case "#"
      Case "Lv"
      Case "Info"
      Case "タスク名"
        setSheet.Range("cell_TaskArea") = Library.getColumnName(colLine)
    
    
      Case "予定日"
        setSheet.Range("A" & line) = "cell_PlanStart"
        setSheet.Range(setVal("cell_LevelInfo") & line) = Library.getColumnName(colLine)
        
        setSheet.Range("A" & line + 1) = "cell_PlanEnd"
        setSheet.Range("B" & line + 1) = Library.getColumnName(colLine + 1)
      
      Case "担当者"
        setSheet.Range("A" & line) = "cell_Assign"
        setSheet.Range(setVal("cell_LevelInfo") & line) = Library.getColumnName(colLine)

      Case "実績日"
        setSheet.Range("A" & line) = "cell_AchievementStart"
        setSheet.Range(setVal("cell_LevelInfo") & line) = Library.getColumnName(colLine)
        
        setSheet.Range("A" & line + 1) = "cell_AchievementEnd"
        setSheet.Range("B" & line + 1) = Library.getColumnName(colLine + 1)
        
      Case "進捗率"
        setSheet.Range("A" & line) = "cell_ProgressLast"
        setSheet.Range(setVal("cell_LevelInfo") & line) = Library.getColumnName(colLine)
        
        setSheet.Range("A" & line + 1) = "cell_Progress"
        setSheet.Range("B" & line + 1) = Library.getColumnName(colLine + 1)
                
      Case "先行タスク"
        setSheet.Range("A" & line) = "cell_Task"
        setSheet.Range(setVal("cell_LevelInfo") & line) = Library.getColumnName(colLine)
        
      Case "タスク配分"
        setSheet.Range("A" & line) = "cell_TaskAllocation"
        setSheet.Range(setVal("cell_LevelInfo") & line) = Library.getColumnName(colLine)
        
      Case "タスク情報"
        setSheet.Range("A" & line) = "cell_TaskInfoP"
        setSheet.Range(setVal("cell_LevelInfo") & line) = Library.getColumnName(colLine)
        
        setSheet.Range("A" & line + 1) = "cell_TaskInfoC"
        setSheet.Range("B" & line + 1) = Library.getColumnName(colLine + 1)
                        
                        
      Case "作業工数"
        setSheet.Range("A" & line) = "cell_WorkLoadP"
        setSheet.Range(setVal("cell_LevelInfo") & line) = Library.getColumnName(colLine)
        
        setSheet.Range("A" & line + 1) = "cell_WorkLoadA"
        setSheet.Range("B" & line + 1) = Library.getColumnName(colLine + 1)
        
      Case "遅早工数"
        setSheet.Range("A" & line) = "cell_LateOrEarly"
        setSheet.Range(setVal("cell_LevelInfo") & line) = Library.getColumnName(colLine)
        
      Case "備考"
        setSheet.Range("A" & line) = "cell_Note"
        setSheet.Range(setVal("cell_LevelInfo") & line) = Library.getColumnName(colLine)
        
        'カレンダー開始セル
        setSheet.Range("A" & line + 1) = "calendarStartCol"
        setSheet.Range("B" & line + 1) = Library.getColumnName(colLine + 1)
        
      Case Else
    End Select

Label_nextFor:
  Next

'  init.logFile = ""
'  Call init.setting
End Function


'**************************************************************************************************
' * タスクリスト確認
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function タスクリスト確認()

  Dim line As Long, endLine As Long
  Dim tmpLine As Long, startTaskLine As Long, endTaskLine As Long
  Dim targetLevel As Long, progress As Long, progressCnt As Long, lateOrEarly As Double, lateOrEarlyCnt As Long
  Dim workStartDay As Date, workEndDay As Date, chkDay As Date
  Dim errorFlg As Boolean, chlkFlg As Boolean
  Dim ErrorMeg As String
  Dim workLoadP As Long
  
'  On Error GoTo catchError
   Call Library.showDebugForm("タスクリスト確認", "開始")

    
  ' 強制的に再計算させる
  Application.CalculateFull
  
  Call init.setting
  mainSheet.Select
  ErrorMeg = ""

  '入力チェック------------------------------------------------------------------------------------
  errorFlg = False
  Call ProgressBar.showCount("タスクリスト確認", 0, 10, "入力チェック")
      
  'タスク名の設定
  endLine = Cells(Rows.count, 1).End(xlUp).row
  
  'カレンダーの開始日とタスクの開始日を確認
  tmpSheet.Cells.Delete Shift:=xlUp
  endLine = Cells(Rows.count, Library.getColumnNo(setVal("cell_PlanStart"))).End(xlUp).row
  If endLine > 6 Then
    Range(setVal("cell_PlanStart") & 6 & ":" & setVal("cell_PlanStart") & endLine).Copy
    tmpSheet.Range("A1").PasteSpecial
    
    tmpSheet.Sort.SortFields.Clear
    tmpSheet.Sort.SortFields.Add Key:=Range("A1"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    Application.CutCopyMode = False
    
    With tmpSheet.Sort
        .SetRange Columns("A:A")
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    If setVal("startDay") > Application.WorksheetFunction.Min(tmpSheet.Columns("A:A")) Then
      Call Library.showDebugForm("タスクリスト確認", "カレンダーの期間よりもタスクの開始日が過去に設定されています")
      ErrorMeg = ErrorMeg & "カレンダーの期間よりもタスクの開始日が過去に設定されています" & vbCrLf
      errorFlg = True
    End If
  End If
  

  'カレンダーの終了日とタスクの終了日を確認
  tmpSheet.Cells.Delete Shift:=xlUp
  endLine = Cells(Rows.count, Library.getColumnNo(setVal("cell_PlanEnd"))).End(xlUp).row
  If endLine > 6 Then
    Range(setVal("cell_PlanEnd") & 6 & ":" & setVal("cell_PlanEnd") & endLine).Copy
    tmpSheet.Range("A1").PasteSpecial
    
    tmpSheet.Sort.SortFields.Clear
    tmpSheet.Sort.SortFields.Add Key:=Range("A1"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    Application.CutCopyMode = False
    
    With tmpSheet.Sort
        .SetRange Columns("A:A")
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    If setVal("endDay") < Application.WorksheetFunction.Max(tmpSheet.Range("A1:A" & Rows.count)) Then
      Call Library.showDebugForm("タスクリスト確認", "カレンダーの期間よりもタスクの終了日が未来に設定されています")
      ErrorMeg = ErrorMeg & "カレンダーの期間よりもタスクの終了日が未来に設定されています" & vbCrLf
      ErrorMeg = ErrorMeg & "　カレンダーの終了日:" & setVal("endDay")
      ErrorMeg = ErrorMeg & "　タスクの終了日:" & Format(Application.WorksheetFunction.Max(tmpSheet.Range("A1:A" & Rows.count)), "yyyy/mm/dd")
      errorFlg = True
    End If
  End If
  
  'カレンダーの開始日とタスクの実績開始日を確認
  tmpSheet.Cells.Delete Shift:=xlUp
  endLine = Cells(Rows.count, Library.getColumnNo(setVal("cell_AchievementStart"))).End(xlUp).row
  If endLine > 6 Then
    Range(setVal("cell_AchievementStart") & 6 & ":" & setVal("cell_AchievementStart") & endLine).Copy
    tmpSheet.Range("A1").PasteSpecial
    
    tmpSheet.Sort.SortFields.Clear
    tmpSheet.Sort.SortFields.Add Key:=Range("A1"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    Application.CutCopyMode = False
    
    With tmpSheet.Sort
        .SetRange Columns("A:A")
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    If setVal("startDay") > Application.WorksheetFunction.Min(tmpSheet.Columns("A:A")) Then
      Call Library.showDebugForm("タスクリスト確認", "カレンダーの期間よりもタスクの実績開始日が過去に設定されています")
      ErrorMeg = ErrorMeg & "カレンダーの期間よりもタスクの実績開始日が過去に設定されています" & vbCrLf
      errorFlg = True
    End If
  End If

  'カレンダーの終了日とタスクの実績終了日を確認
  tmpSheet.Cells.Delete Shift:=xlUp
  endLine = Cells(Rows.count, Library.getColumnNo(setVal("cell_AchievementEnd"))).End(xlUp).row
  If endLine > 6 Then
    Range(setVal("cell_AchievementEnd") & 6 & ":" & setVal("cell_AchievementEnd") & endLine).Copy
    tmpSheet.Range("A1").PasteSpecial
    
    tmpSheet.Sort.SortFields.Clear
    tmpSheet.Sort.SortFields.Add Key:=Range("A1"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    Application.CutCopyMode = False
    
    With tmpSheet.Sort
        .SetRange Columns("A:A")
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    If setVal("endDay") < Application.WorksheetFunction.Max(tmpSheet.Range("A1:A" & Rows.count)) Then
      Call Library.showDebugForm("タスクリスト確認", "カレンダーの期間よりもタスクの実績終了日が未来に設定されています")
      ErrorMeg = ErrorMeg & "カレンダーの期間よりもタスクの実績終了日が未来に設定されています" & vbCrLf
      errorFlg = True
    End If
  End If
  
  If Not (setVal("startDay") <= setVal("baseDay") And setVal("baseDay") <= setVal("endDay")) Then
    Call Library.showDebugForm("基準日がカレンダーの期間外に設定されています")
    ErrorMeg = ErrorMeg & "基準日がカレンダーの期間外に設定されています" & vbCrLf
'    errorFlg = True
    Call WBS_Option.エラー情報表示(ErrorMeg)
    
  End If
  
  '予定日のチェック--------------------------------------------------------------------------------
  endLine = Cells(Rows.count, 1).End(xlUp).row
  For line = 6 To endLine
    If Range(setVal("cell_PlanStart") & line) > Range(setVal("cell_PlanEnd") & line) Then
      Range(setVal("cell_PlanStart") & line).Style = "Error"
      ErrorMeg = ErrorMeg & line & "行:予定日(開始日)が予定日(終了日)より未来です。" & vbCrLf
      errorFlg = True
    End If
  
    '作業工数(予定)の算出------------------------------------------------------------------------
    workLoadP = WorksheetFunction.NetworkDays_Intl(Range(setVal("cell_PlanStart") & line), Range(setVal("cell_PlanEnd") & line), "0000011", Range("休日リスト"))
    If Range(setVal("cell_PlanStart") & line) <> "" And Range(setVal("cell_PlanEnd") & line) <> "" Then
      If Range(setVal("cell_WorkLoadP") & line) > workLoadP And Not (Range(setVal("cell_WorkLoadP") & line).Formula Like "=*") Then
        Range(setVal("cell_WorkLoadP") & line).Style = "Error"
        ErrorMeg = ErrorMeg & line & "行:作業工数(予定)が実際の期間より多いです。　　入力値:" & Range(setVal("cell_WorkLoadP") & line) & " 計算値:" & workLoadP & vbCrLf
        errorFlg = True
      End If
    End If
  
    '進捗確認------------------------------------------------------------------------
    If Range(setVal("cell_Progress") & line) > 0 Then
      If Range(setVal("cell_AchievementStart") & line) = "" Then
        Range(setVal("cell_AchievementStart") & line).Style = "Error"
        ErrorMeg = ErrorMeg & line & "行:実績日(開始)が入力されていません" & vbCrLf
        errorFlg = True
      End If
      If Range(setVal("cell_AchievementEnd") & line) = "" And Range(setVal("cell_Progress") & line) = 100 Then
        Range(setVal("cell_AchievementEnd") & line).Style = "Error"
        ErrorMeg = ErrorMeg & line & "行:実績日(終了)が入力されていません" & vbCrLf
        errorFlg = True
      End If
    End If
  
  
  
  Next
  
  'エラー情報の表示--------------------------------------------------------------------------------
  tmpSheet.Cells.Delete Shift:=xlUp
  If errorFlg = True Then
    Call WBS_Option.エラー情報表示(ErrorMeg)
    GoTo catchError
  End If
  
  
  endLine = Cells(Rows.count, 1).End(xlUp).row
  
  '遅早工数、タスク情報のクリア
  Range(setVal("cell_LateOrEarly") & "4:" & setVal("cell_LateOrEarly") & endLine).ClearContents
  Range(setVal("cell_TaskInfoP") & 6 & ":" & setVal("cell_TaskInfoC") & endLine).ClearContents
  
  'マクロで設定した実績、工程をクリア
  For line = 6 To endLine
    If Range(setVal("cell_WorkLoadP") & line).Formula Like "=*" Then
      Range(setVal("cell_WorkLoadP") & line).ClearContents
    End If
    
    If Range(setVal("cell_WorkLoadA") & line).Formula Like "=*" Then
      Range(setVal("cell_WorkLoadA") & line).ClearContents
    End If
    
    If Range(setVal("cell_Assign") & line) = "工程" Then
      Range(setVal("cell_Assign") & line) = ""
    End If
  Next
  
  
  '親タスクなら、担当者(予定)に「工程」を割り当て
  Call Library.showDebugForm("タスクリスト確認", "親/子タスク判定")
  For line = 6 To endLine
    Call ProgressBar.showCount("タスクリスト確認", line, endLine, "親/子タスク判定")
    
    'タスクレベルが1ならリセット
    If Range(setVal("cell_LevelInfo") & line) = 1 Then
      parentTaskLine = ""
    ElseIf Range(setVal("cell_LevelInfo") & line) = "" Then
      GoTo Label_nextFor
    End If
    If Range(setVal("cell_LevelInfo") & line) < Range("B" & line + 1) And Range("B" & line + 1) <> "" Then
      endTaskLine = line + 1
      Do While Range(setVal("cell_LevelInfo") & line).Value <= Range("B" & endTaskLine).Value And Range("B" & endTaskLine) <> ""
        Call ProgressBar.showCount("タスクリスト確認", endTaskLine, endLine, "親/子タスク判定")
      
        If Range(setVal("cell_LevelInfo") & line).Value >= Range("B" & endTaskLine).Value Then
          endTaskLine = endTaskLine - 1
          Exit Do
        End If
        endTaskLine = endTaskLine + 1
      Loop
      If Range(setVal("cell_LevelInfo") & line).Value >= Range("B" & endTaskLine).Value Then
        endTaskLine = endTaskLine - 1
      End If
      Range(setVal("cell_Assign") & line) = "工程"
      
      'タスクレベルによる色分け
      Select Case Range(setVal("cell_LevelInfo") & line)
        Case "1"
          If setVal("lineColor_TaskLevel1") = 16777215 Then
          Else
            Range("A" & line & ":" & setVal("cell_Note") & line).Interior.Color = setVal("lineColor_TaskLevel1")
          End If
        Case "2"
          If setVal("lineColor_TaskLevel2") = 16777215 Then
          Else
            Range("A" & line & ":" & setVal("cell_Note") & line).Interior.Color = setVal("lineColor_TaskLevel2")
          End If
        Case "3"
          If setVal("lineColor_TaskLevel2") = 16777215 Then
          Else
            Range("A" & line & ":" & setVal("cell_Note") & line).Interior.Color = setVal("lineColor_TaskLevel3")
          End If
      End Select
      
      
      '作業工数(実績)の算出--------------------------------------------------------------------------
      If Range(setVal("cell_PlanStart") & line) <> "" And Range(setVal("cell_PlanEnd") & line) <> "" Then
        If Range(setVal("cell_WorkLoadA") & line).Formula Like "=*" Or Range(setVal("cell_WorkLoadA") & line) = "" Then
          If Range(setVal("cell_PlanStart") & line) <= setVal("baseDay") Then
            Range(setVal("cell_WorkLoadA") & line) = "=" & WorksheetFunction.NetworkDays_Intl(Range(setVal("cell_PlanStart") & line), setVal("baseDay"), "0000011", Range("休日リスト"))
          ElseIf Range(setVal("cell_AchievementStart") & line) <= setVal("baseDay") Then
            Range(setVal("cell_WorkLoadA") & line) = "=" & WorksheetFunction.NetworkDays_Intl(Date, Range(setVal("cell_PlanStart") & line), "0000011", Range("休日リスト"))
          End If
        End If
      End If
      
      '子タスクの範囲を保存
      Range(setVal("cell_TaskInfoC") & line) = line + 1 & ":" & endTaskLine
      
      '親タスク情報
      Range(setVal("cell_TaskInfoP") & line) = parentTaskLine
    
      parentTaskLine = line
    Else
      '親タスク情報
      Range(setVal("cell_TaskInfoP") & line) = parentTaskLine
    End If
Label_nextFor:
  Next
  
  '親タスク情報
  Call Library.showDebugForm("タスクリスト確認", "親タスク情報の確認")
  For line = endLine To 6 Step -1
    If Range(setVal("cell_Assign") & line) = "工程" Then
      Range(setVal("cell_TaskInfoP") & line).Select
      For lineP = line - 1 To 6 Step -1
        If Range(setVal("cell_Assign") & lineP) = "工程" And (Range(setVal("cell_LevelInfo") & line) > Range(setVal("cell_LevelInfo") & lineP)) Then
          Range(setVal("cell_TaskInfoP") & lineP).Select
          Range(setVal("cell_TaskInfoP") & line) = Range(setVal("cell_LevelInfo") & lineP).row
          Exit For
        End If
      Next
    End If
  Next
  
  
  '子タスクのデータ確認
  Call Library.showDebugForm("タスクリスト確認", "子タスクのデータ確認")
  For line = 6 To endLine
    Call ProgressBar.showCount("タスクリスト確認", line, endLine, "子タスクのデータ確認")

'    Call Library.showDebugForm("タスクリスト確認", "　" & Range(setVal("cell_Info") & line))
    
    'Levelがなければループを抜ける
    If Range(setVal("cell_LevelInfo") & line) = "" Then Exit For
    
    If Range(setVal("cell_Assign") & line) <> "工程" And Range(setVal("cell_Info") & line) <> setVal("TaskInfoStr_Multi") Then
      '実績日(開始と終了)が入力されていれば、進捗を100にする---------------------------------------
      If Range(setVal("cell_AchievementStart") & line) <> "" And Range(setVal("cell_AchievementEnd") & line) <> "" Then
        Range(setVal("cell_Progress") & line) = 100
      End If
    
      '作業工数(予定)の算出--------------------------------------------------------------------------
      If Range(setVal("cell_PlanStart") & line) <> "" And Range(setVal("cell_PlanEnd") & line) <> "" Then
        If Range(setVal("cell_WorkLoadP") & line).Formula Like "=*" Or Range(setVal("cell_WorkLoadP") & line) = "" Then
          Range(setVal("cell_WorkLoadP") & line) = "=" & WorksheetFunction.NetworkDays_Intl(Range(setVal("cell_PlanStart") & line), Range(setVal("cell_PlanEnd") & line), "0000011", Range("休日リスト"))
        End If
      End If
      
      '作業工数(実績)の算出--------------------------------------------------------------------------
      If Range(setVal("cell_PlanStart") & line) <> "" And Range(setVal("cell_PlanEnd") & line) <> "" Then
        If Range(setVal("cell_WorkLoadA") & line).Formula Like "=*" Or Range(setVal("cell_WorkLoadA") & line) = "" Then
          If Range(setVal("cell_PlanStart") & line) <= setVal("baseDay") Then
            Range(setVal("cell_WorkLoadA") & line) = "=" & WorksheetFunction.NetworkDays_Intl(Range(setVal("cell_PlanStart") & line), setVal("baseDay"), "0000011", Range("休日リスト"))
'          Else
'            Range(setVal("cell_WorkLoadA") & line) = "=" & WorksheetFunction.NetworkDays_Intl(Date, Range(setVal("cell_PlanStart") & line), "0000011", Range("休日リスト"))
          End If
        End If
      End If
      
      '進捗率の設定
      '作業予定日を経過しているが、未入力の場合
      If Range(setVal("cell_Progress") & line) = "" And Range(setVal("cell_PlanStart") & line) < setVal("baseDay") Then
'        Range(setVal("cell_Progress") & line) = "=" & 0
      End If
      
      '遅早工数の計算--------------------------------------------------------------------------------
'      Range(setVal("cell_Progress") & line).Select
      
      '遅早工数=(作業工数_実績-(作業工数_予定/進捗率))*-1
      If Range(setVal("cell_Progress") & line) = 100 And Range(setVal("cell_PlanEnd") & line) < setVal("baseDay") Then
        Range(setVal("cell_LateOrEarly") & line) = 0
        
      ElseIf Range(setVal("cell_Progress") & line) <> "" Then
        Range(setVal("cell_LateOrEarly") & line) = (Range(setVal("cell_WorkLoadA") & line) - (Range(setVal("cell_WorkLoadP") & line) * Range(setVal("cell_Progress") & line) / 100)) * -1
      End If
    End If
  Next
  
  '親タスクのデータ確認============================================================================
  Call Library.showDebugForm("タスクリスト確認", "親タスクのデータ確認")
  For line = 6 To endLine
    Call ProgressBar.showCount("タスクリスト確認", line, endLine, "親タスクのデータ確認")
    If Range(setVal("cell_TaskInfoC") & line) <> "" Then
      taskAreas = Split(Range(setVal("cell_TaskInfoC") & line), ":")
      
      '予定日(開始)設定----------------------------------------------------------------------------------
      workStartDay = Application.WorksheetFunction.Max(Range(setVal("cell_PlanStart") & taskAreas(0) & ":" & setVal("cell_PlanStart") & taskAreas(1)))
      For tmpLine = taskAreas(0) To taskAreas(1)
        If workStartDay > Range(setVal("cell_PlanStart") & tmpLine) And Range(setVal("cell_PlanStart") & tmpLine) <> "" Then
          workStartDay = Range(setVal("cell_PlanStart") & tmpLine)
        End If
      Next
      If workStartDay <> 0 Then
        Range(setVal("cell_PlanStart") & line) = workStartDay
      End If
      
      '予定日(終了)設定----------------------------------------------------------------------------------
      workEndDay = Application.WorksheetFunction.Min(Range(setVal("cell_PlanEnd") & taskAreas(0) & ":" & setVal("cell_PlanEnd") & taskAreas(1)))
      For tmpLine = taskAreas(0) To taskAreas(1)
        If workEndDay < Range(setVal("cell_PlanEnd") & tmpLine) And Range(setVal("cell_PlanEnd") & tmpLine) <> "" Then
          workEndDay = Range(setVal("cell_PlanEnd") & tmpLine)
        End If
      Next
      If workEndDay <> 0 Then
        Range(setVal("cell_PlanEnd") & line) = workEndDay
      End If
      
      
      '作業工数(予定)の算出------------------------------------------------------------------------
      If Range(setVal("cell_PlanStart") & line) <> "" And Range(setVal("cell_PlanEnd") & line) <> "" Then
        Range(setVal("cell_WorkLoadP") & line) = "=" & WorksheetFunction.NetworkDays_Intl(Range(setVal("cell_PlanStart") & line), Range(setVal("cell_PlanEnd") & line), "0000011", Range("休日リスト"))
      End If
      
      
      '実績日の設定--------------------------------------------------------------------------------
      workStartDay = Application.WorksheetFunction.Max(Range(setVal("cell_AchievementStart") & taskAreas(0) & ":" & setVal("cell_AchievementStart") & taskAreas(1)))
      
      For tmpLine = taskAreas(0) To taskAreas(1)
        If workStartDay > Range(setVal("cell_AchievementStart") & tmpLine) And Range(setVal("cell_AchievementStart") & tmpLine) <> "" Then
          workStartDay = Range(setVal("cell_AchievementStart") & tmpLine)
        End If
      Next
      If workStartDay <> 0 Then
        Range(setVal("cell_AchievementStart") & line) = workStartDay
      End If
      
      If WorksheetFunction.CountBlank(Range(setVal("cell_AchievementEnd") & taskAreas(0) & ":" & setVal("cell_AchievementEnd") & taskAreas(1))) = 0 Then
        Range(setVal("cell_AchievementEnd") & line) = Application.WorksheetFunction.Max(Range(setVal("cell_AchievementEnd") & taskAreas(0) & ":" & setVal("cell_AchievementEnd") & taskAreas(1)))
      End If
      
      '進捗の計算----------------------------------------------------------------------------------
      progress = 0
      progressCnt = 0
      For tmpLine = taskAreas(0) To taskAreas(1)
        If Range(setVal("cell_Assign") & tmpLine) <> "工程" Then
          progress = progress + Range(setVal("cell_Progress") & tmpLine)
          progressCnt = progressCnt + 1
        End If
      Next
      If progressCnt = 0 Or progress = 0 Then
        Range(setVal("cell_Progress") & line) = ""
      Else
        Range(setVal("cell_Progress") & line) = progress / progressCnt
      End If
  
      '遅早工数の計算--------------------------------------------------------------------------------
      lateOrEarly = 0
      lateOrEarlyCnt = 0
      For tmpLine = taskAreas(0) To taskAreas(1)
        If Range(setVal("cell_Assign") & tmpLine) <> "工程" Then
          lateOrEarly = lateOrEarly + Range(setVal("cell_LateOrEarly") & tmpLine)
          lateOrEarlyCnt = lateOrEarlyCnt + 1
        End If
      Next
       'Range(setVal("cell_LateOrEarly") & line).Select
      If lateOrEarlyCnt = 0 Then
        Range(setVal("cell_LateOrEarly") & line) = ""
      Else
        Columns(setVal("cell_LateOrEarly") & ":" & setVal("cell_LateOrEarly")).NumberFormatLocal = "0.00_ ;[赤]-0.00 "
        Range(setVal("cell_LateOrEarly") & line) = lateOrEarly
      End If
    End If
  Next
  Call Library.showDebugForm("タスクリスト確認", "全体の進捗計算")
  
  '全体の進捗の計算--------------------------------------------------------------------------------
  progressCnt = 0
  progress = 0
  lateOrEarly = 0
  For line = 6 To endLine
    Call ProgressBar.showCount("タスクリスト確認", line, endLine, "全タスクのデータ集計")
    
    If Range(setVal("cell_Assign") & line).Text <> "工程" Then
      Range(setVal("cell_Assign") & line).Select
      progress = progress + Range(setVal("cell_Progress") & line)
      progressCnt = progressCnt + 1
      lateOrEarly = lateOrEarly + Range(setVal("cell_LateOrEarly") & line)
    End If
    
    '進捗が100%なら非表示------------------------------------
    If setVal("setDispProgress100") = True And Range(setVal("cell_Progress") & line) = 100 Then
      Rows(line & ":" & line).EntireRow.Hidden = True
      
    End If
    
  Next
  If progressCnt > 1 Then
    Range(setVal("cell_Progress") & 5) = progress / progressCnt
    Range(setVal("cell_LateOrEarly") & 5) = lateOrEarly
  ElseIf progressCnt = 1 Then
    Range(setVal("cell_Progress") & 5) = progress
    Range(setVal("cell_LateOrEarly") & 5) = lateOrEarly
  End If
  
  Call Library.showDebugForm("タスクリスト確認", "終了")
  
  Exit Function
'エラー発生時=====================================================================================
catchError:
  Call Library.endScript

End Function

