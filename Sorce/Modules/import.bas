Attribute VB_Name = "import"
'ワークブック用変数==============================


'ワークシート用変数==============================
'Dim masterSheet As Worksheet

'グローバル変数==================================



'**************************************************************************************************
' * import用機能
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function ファイルインポート()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  Dim dataDirPath As String, filePath As String
  
'  On Error GoTo catchError
  
  Call init.setting
  sheetMain.Select
  
  dataDirPath = Library.getRegistry("dataDirPath")
  If dataDirPath = "" Then
    dataDirPath = init.ThisBook.Path
  End If
  
  Call Library.showNotice(1, "プロジェクト")
  
  filePaths = Library.getFilesPath(dataDirPath, "", "プロジェクトごとのファイルを選択してください", 1)
  If filePaths(0) = "" Then
    Call Library.showDebugForm("ファイルインポート", "ファイル選択キャンセル")
    
    Call Library.showNotice(100, , True)
    End
  End If

  If deleteFlg = True Then
    Call WBS_Option.clearAll
  Else
    Call WBS_Option.clearCalendar
  End If




  For i = 0 To UBound(filePaths)
    filePath = filePaths(i)
    Call ctl_ProgressBar.showCount("ファイルインポート", i + 1, UBound(filePaths) + 1, "対象：" & Dir(filePath))
    
    '指定ファイルオープンし、シートの存在確認
    Set targetBook = Workbooks.Open(fileName:=filePath, ReadOnly:=True)
    Windows(targetBook.Name).WindowState = xlMinimized
    Call Library.startScript
    targetBook.Activate
    
    If Library.chkSheetExists("メイン") = True Then
      Call データコピー(filePath)
    ElseIf Library.chkSheetExists("calendar") = True Then
      Call CD部用.データコピー(filePath)
    Else
      Call Library.showNotice(405, "<" & Dir(filePath) & "ファイルにメイン または calendar>")
    End If
  Next

  dataDirPath = Replace(filePath, "\" & Dir(filePath), "")
  Call Library.setRegistry("dataDirPath", dataDirPath)

  Exit Function
'エラー発生時--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(Err.Number, Err.Description, True)

End Function


'**************************************************************************************************
' * データコピー
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function データコピー(filePath As String)
  Dim line As Long, endLine As Long, targetEndLine As Long, endColLine As Long
  Dim tmpEndLine As Long
  Dim targetSetVal As Collection
  Dim targetLevel As Long
  Dim prgbarMeg As String
  Dim prgbarCnt As Long
  Dim taskLevelRange As Range
  
  
  On Error GoTo catchError
  
  Set targetSetVal = New Collection
  prgbarCnt = 0

  Call ctl_ProgressBar.showCount("ファイルインポート", prgbarCnt, 100, "対象：" & Dir(filePath))
  
  If Library.chkSheetExists("メイン") = True Then
    'インポートファイルの設定読み込み
    With targetSetVal
      For line = 3 To targetBook.Sheets("設定").Cells(Rows.count, 1).End(xlUp).row
        If targetBook.Sheets("設定").Range("A" & line) <> "" Then
         .Add item:=targetBook.Sheets("設定").Range(setVal("cell_LevelInfo") & line), Key:=targetBook.Sheets("設定").Range("A" & line)
        End If
      Next
    End With
  
    endLine = sheetMain.Cells(Rows.count, 1).End(xlUp).row + 1
    
    'ファイル名をタスクとして登録
    prgbarMeg = "ファイル名をタスクとして登録"
    prgbarCnt = prgbarCnt + 1
    Call ctl_ProgressBar.showCount(Dir(filePath), prgbarCnt, 11, prgbarMeg)
    
    sheetMain.Range("B" & endLine) = 1
    sheetMain.Range(setVal("cell_TaskArea") & endLine) = Dir(filePath)
    sheetMain.Range(setVal("cell_Note") & endLine) = filePath

    endLine = endLine + 1
    
    Call Library.showDebugForm("ファイルインポート", "インポート開始")
    targetEndLine = targetBook.Worksheets(sheetMainName).Cells(Rows.count, 1).End(xlUp).row
    
    '#〜タスク名をコピー
    prgbarMeg = "タスク名列までをコピー"
    prgbarCnt = prgbarCnt + 1
    Call ctl_ProgressBar.showCount(Dir(filePath), prgbarCnt, 11, prgbarMeg)
    
    targetBook.Worksheets(sheetMainName).Range("A6:" & targetSetVal("cell_TaskArea") & targetEndLine).Copy
    sheetMain.Range("A" & endLine).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
    '予定日をコピー
    prgbarMeg = "予定日をコピー"
    prgbarCnt = prgbarCnt + 1
    Call ctl_ProgressBar.showCount(Dir(filePath), prgbarCnt, 11, prgbarMeg)

    targetBook.Worksheets(sheetMainName).Range(targetSetVal("cell_PlanStart") & "6:" & targetSetVal("cell_PlanEnd") & targetEndLine).Copy
    sheetMain.Range(setVal("cell_PlanStart") & endLine).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
    '担当者をコピー
    prgbarMeg = "担当者をコピー"
    prgbarCnt = prgbarCnt + 1
    Call ctl_ProgressBar.showCount(Dir(filePath), prgbarCnt, 11, prgbarMeg)
    
    targetBook.Worksheets(sheetMainName).Range(targetSetVal("cell_Assign") & "6:" & targetSetVal("cell_Assign") & targetEndLine).Copy
    sheetMain.Range(setVal("cell_Assign") & endLine).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        
    '実績日をコピー
    prgbarMeg = "実績日をコピー"
    prgbarCnt = prgbarCnt + 1
    Call ctl_ProgressBar.showCount(Dir(filePath), prgbarCnt, 11, prgbarMeg)
    
    targetBook.Worksheets(sheetMainName).Range(targetSetVal("cell_AchievementStart") & "6:" & targetSetVal("cell_AchievementEnd") & targetEndLine).Copy
    sheetMain.Range(setVal("cell_AchievementStart") & endLine).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        
    '進捗率をコピー
    prgbarMeg = "A〜C列コピー"
    prgbarCnt = prgbarCnt + 1
    Call ctl_ProgressBar.showCount(Dir(filePath), prgbarCnt, 11, prgbarMeg)
    
    targetBook.Worksheets(sheetMainName).Range(targetSetVal("cell_ProgressLast") & "6:" & targetSetVal("cell_Progress") & targetEndLine).Copy
    sheetMain.Range(setVal("cell_ProgressLast") & endLine).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        
    '先行タスクをコピー
    prgbarMeg = "先行タスクをコピー"
    prgbarCnt = prgbarCnt + 1
    Call ctl_ProgressBar.showCount(Dir(filePath), prgbarCnt, 11, prgbarMeg)
    
    targetBook.Worksheets(sheetMainName).Range(targetSetVal("cell_Task") & "6:" & targetSetVal("cell_Task") & targetEndLine).Copy
    sheetMain.Range(setVal("cell_Task") & endLine).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        
    'タスク配分をコピー
    prgbarMeg = "タスク配分をコピー"
    prgbarCnt = prgbarCnt + 1
    Call ctl_ProgressBar.showCount(Dir(filePath), prgbarCnt, 11, prgbarMeg)
    
    targetBook.Worksheets(sheetMainName).Range(targetSetVal("cell_TaskAllocation") & "6:" & targetSetVal("cell_TaskAllocation") & targetEndLine).Copy
    sheetMain.Range(setVal("cell_TaskAllocation") & endLine).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        
    '作業工数をコピー
    prgbarMeg = "作業工数をコピー"
    prgbarCnt = prgbarCnt + 1
    Call ctl_ProgressBar.showCount(Dir(filePath), prgbarCnt, 11, prgbarMeg)
    
    targetBook.Worksheets(sheetMainName).Range(targetSetVal("cell_WorkLoadP") & "6:" & targetSetVal("cell_WorkLoadA") & targetEndLine).Copy
    sheetMain.Range(setVal("cell_WorkLoadP") & endLine).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        
    '遅早工数をコピー
    prgbarMeg = "遅早工数をコピー"
    prgbarCnt = prgbarCnt + 1
    Call ctl_ProgressBar.showCount(Dir(filePath), prgbarCnt, 11, prgbarMeg)
    
    targetBook.Worksheets(sheetMainName).Range(targetSetVal("cell_LateOrEarly") & "6:" & targetSetVal("cell_LateOrEarly") & targetEndLine).Copy
    sheetMain.Range(setVal("cell_LateOrEarly") & endLine).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
         
    '備考をコピー
    prgbarMeg = "備考をコピー"
    prgbarCnt = prgbarCnt + 1
    Call ctl_ProgressBar.showCount(Dir(filePath), prgbarCnt, 11, prgbarMeg)
    
    targetBook.Worksheets(sheetMainName).Range(targetSetVal("cell_Note") & "6:" & targetSetVal("cell_Note") & targetEndLine).Copy
    sheetMain.Range(setVal("cell_Note") & endLine).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
  End If
    
  Set targetSetVal = Nothing
  targetBook.Close
  
  ThisWorkbook.Activate
  sheetMain.Select
  Call Library.startScript
  
  tmpEndLine = Cells(Rows.count, 1).End(xlUp).row
  
  'レベルの再設定
  Call ctl_ProgressBar.showCount(Dir(filePath), 0, tmpEndLine, "レベルの再設定")
  For line = endLine To tmpEndLine
    Call ctl_ProgressBar.showCount("レベルの再設定", line, tmpEndLine, "")

    targetLevel = sheetMain.Range(setVal("cell_LevelInfo") & line) + 1
    sheetMain.Range(setVal("cell_LevelInfo") & line) = targetLevel
    If targetLevel <> 0 Then
      sheetMain.Range(setVal("cell_Info") & line).InsertIndent targetLevel
    End If
  Next
  sheetMain.Calculate
  
  Exit Function
'エラー発生時--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(Err.Number, Err.Description, True)

End Function


'**************************************************************************************************
' * カレンダー用日程取得
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function カレンダー用日程取得()
  Dim line As Long, endLine As Long, tmpLine As Long
  Dim dataDirPath As String, filePath As String
  Dim workStartDay As Date, workEndDay As Date
  
'  On Error GoTo catchError
  Call init.setting
  sheetMain.Select
  
  endLine = Cells(Rows.count, 1).End(xlUp).row
  workStartDay = Application.WorksheetFunction.Max(Range(setVal("cell_PlanStart") & "6:" & setVal("cell_PlanStart") & endLine))
  For tmpLine = 6 To endLine
    If workStartDay > Range(setVal("cell_PlanStart") & tmpLine) And Range(setVal("cell_PlanStart") & tmpLine) <> "" Then
      workStartDay = Range(setVal("cell_PlanStart") & tmpLine) - 10
    End If
  Next
  If workStartDay <> 0 Then
    Range("startDay") = workStartDay
  End If

  workEndDay = Application.WorksheetFunction.Min(Range(setVal("cell_PlanStart") & "6:" & setVal("cell_PlanStart") & endLine))
  For tmpLine = 6 To endLine
    If workEndDay < Range(setVal("cell_PlanEnd") & tmpLine) And Range(setVal("cell_PlanEnd") & tmpLine) <> "" Then
      workEndDay = Range(setVal("cell_PlanEnd") & tmpLine) + 30
    End If
  Next
  If workStartDay <> 0 Then
    Range("endDay") = workEndDay
  End If
  
  If workStartDay <= Date And Date <= workEndDay Then
    Range("baseDay") = Date
  End If

  Call Calendar.makeCalendar
  sheetMain.Calculate
  Call Check.タスクリスト確認
  Call Chart.ガントチャート生成
  
  Exit Function
'エラー発生時--------------------------------------------------------------------------------------
catchError:

End Function
