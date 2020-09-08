Attribute VB_Name = "import"
'ワークブック用変数------------------------------
Dim targetBook As Workbook

'ワークシート用変数------------------------------
'Dim masterSheet As Worksheet

'グローバル変数----------------------------------



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
  mainSheet.Select
  
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

  For i = 0 To UBound(filePaths)
    filePath = filePaths(i)
    Call Library.showDebugForm("ファイルインポート", "対象：" & Dir(filePath))
    Call ProgressBar.showCount("ファイルインポート", i + 1, UBound(filePaths) + 1, "対象：" & Dir(filePath))
    
    If setVal("workMode") = "default" Then
      Call データコピー(filePath)
    ElseIf setVal("workMode") = "CD部用" Then
      Call CD部用.データコピー(filePath)
    End If
  Next

  dataDirPath = Replace(filePath, "\" & Dir(filePath), "")
  Call Library.setRegistry("dataDirPath", dataDirPath)

  Exit Function
'エラー発生時=====================================================================================
catchError:


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

  Set targetBook = Workbooks.Open(filePath, , True)
  Call Library.startScript
  targetBook.Activate
  
  
  Call ProgressBar.showCount("ファイルインポート", prgbarCnt, 100, "対象：" & Dir(filePath))
  
  If Library.chkSheetName("WBS") = True Then
    Call Library.showDebugForm("ファイルインポート", "WBS シート発見")
    
    'インポートファイルの設定読み込み
    With targetSetVal
      For line = 3 To targetBook.Sheets("設定").Cells(Rows.count, 1).End(xlUp).row
        If targetBook.Sheets("設定").Range("A" & line) <> "" Then
         .Add item:=targetBook.Sheets("設定").Range("B" & line), Key:=targetBook.Sheets("設定").Range("A" & line)
        End If
      Next
    End With
  
    endLine = mainSheet.Cells(Rows.count, 1).End(xlUp).row + 1
    
    'ファイル名をタスクとして登録
    prgbarMeg = "ファイル名をタスクとして登録"
    prgbarCnt = prgbarCnt + 1
    Call Library.showDebugForm(Dir(filePath), prgbarMeg)
    Call ProgressBar.showCount(Dir(filePath), prgbarCnt, 10, prgbarMeg)
    
    mainSheet.Range("B" & endLine) = 0
    mainSheet.Range(setVal("cell_Info") & endLine) = Dir(filePath)
    mainSheet.Range(setVal("cell_Note") & endLine) = filePath

    endLine = endLine + 1
    
    Call Library.showDebugForm("ファイルインポート", "インポート開始")
    targetEndLine = targetBook.Worksheets(mainSheetName).Cells(Rows.count, 1).End(xlUp).row
    
    '#～タスク名をコピー
    prgbarMeg = "A～C列コピー"
    prgbarCnt = prgbarCnt + 1
    Call Library.showDebugForm(Dir(filePath), prgbarMeg)
    Call ProgressBar.showCount(Dir(filePath), prgbarCnt, 10, prgbarMeg)
    
    targetBook.Worksheets(mainSheetName).Range("A6:" & targetSetVal("cell_TaskArea") & targetEndLine).Copy
    mainSheet.Range("A" & endLine).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
    '予定日をコピー
    prgbarMeg = "予定日をコピー"
    prgbarCnt = prgbarCnt + 1
    Call Library.showDebugForm(Dir(filePath), prgbarMeg)
    Call ProgressBar.showCount(Dir(filePath), prgbarCnt, 10, prgbarMeg)

    targetBook.Worksheets(mainSheetName).Range(targetSetVal("cell_PlanStart") & "6:" & targetSetVal("cell_PlanEnd") & targetEndLine).Copy
    mainSheet.Range(setVal("cell_PlanStart") & endLine).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
    '担当者をコピー
    prgbarMeg = "担当者をコピー"
    prgbarCnt = prgbarCnt + 1
    Call Library.showDebugForm(Dir(filePath), prgbarMeg)
    Call ProgressBar.showCount(Dir(filePath), prgbarCnt, 10, prgbarMeg)
    
    targetBook.Worksheets(mainSheetName).Range(targetSetVal("cell_Assign") & "6:" & targetSetVal("cell_Assign") & targetEndLine).Copy
    mainSheet.Range(setVal("cell_Assign") & endLine).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        
    '実績日をコピー
    prgbarMeg = "実績日をコピー"
    prgbarCnt = prgbarCnt + 1
    Call Library.showDebugForm(Dir(filePath), prgbarMeg)
    Call ProgressBar.showCount(Dir(filePath), prgbarCnt, 10, prgbarMeg)
    
    targetBook.Worksheets(mainSheetName).Range(targetSetVal("cell_AchievementStart") & "6:" & targetSetVal("cell_AchievementEnd") & targetEndLine).Copy
    mainSheet.Range(setVal("cell_AchievementStart") & endLine).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        
    '進捗率をコピー
    prgbarMeg = "A～C列コピー"
    prgbarCnt = prgbarCnt + 1
    Call Library.showDebugForm(Dir(filePath), prgbarMeg)
    Call ProgressBar.showCount(Dir(filePath), prgbarCnt, 10, prgbarMeg)
    
    targetBook.Worksheets(mainSheetName).Range(targetSetVal("cell_ProgressLast") & "6:" & targetSetVal("cell_Progress") & targetEndLine).Copy
    mainSheet.Range(setVal("cell_ProgressLast") & endLine).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        
    'タスクをコピー
    prgbarMeg = "タスクをコピー"
    prgbarCnt = prgbarCnt + 1
    Call Library.showDebugForm(Dir(filePath), prgbarMeg)
    Call ProgressBar.showCount(Dir(filePath), prgbarCnt, 10, prgbarMeg)
    
    targetBook.Worksheets(mainSheetName).Range(targetSetVal("cell_TaskA") & "6:" & targetSetVal("cell_TaskB") & targetEndLine).Copy
    mainSheet.Range(setVal("cell_TaskA") & endLine).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        
    '作業工数をコピー
    prgbarMeg = "作業工数をコピー"
    prgbarCnt = prgbarCnt + 1
    Call Library.showDebugForm(Dir(filePath), prgbarMeg)
    Call ProgressBar.showCount(Dir(filePath), prgbarCnt, 10, prgbarMeg)
    
    targetBook.Worksheets(mainSheetName).Range(targetSetVal("cell_WorkLoadP") & "6:" & targetSetVal("cell_WorkLoadA") & targetEndLine).Copy
    mainSheet.Range(setVal("cell_WorkLoadP") & endLine).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        
    '遅早工数をコピー
    prgbarMeg = "遅早工数をコピー"
    prgbarCnt = prgbarCnt + 1
    Call Library.showDebugForm(Dir(filePath), prgbarMeg)
    Call ProgressBar.showCount(Dir(filePath), prgbarCnt, 10, prgbarMeg)
    
    targetBook.Worksheets(mainSheetName).Range(targetSetVal("cell_LateOrEarly") & "6:" & targetSetVal("cell_LateOrEarly") & targetEndLine).Copy
    mainSheet.Range(setVal("cell_LateOrEarly") & endLine).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
         
    '備考をコピー
    prgbarMeg = "備考をコピー"
    prgbarCnt = prgbarCnt + 1
    Call Library.showDebugForm(Dir(filePath), prgbarMeg)
    Call ProgressBar.showCount(Dir(filePath), prgbarCnt, 10, prgbarMeg)
    
    targetBook.Worksheets(mainSheetName).Range(targetSetVal("cell_Note") & "6:" & targetSetVal("cell_Note") & targetEndLine).Copy
    mainSheet.Range(setVal("cell_Note") & endLine).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
  End If
    
  Set targetSetVal = Nothing
  targetBook.Close
  
  ThisWorkbook.Activate
  mainSheet.Select
  
  Call Library.showDebugForm("ファイルインポート", "WBSシート A列設定")
  tmpEndLine = Cells(Rows.count, 3).End(xlUp).row
  
  'レベルの再設定
  Call Library.showDebugForm(Dir(filePath), prgbarMeg)
  For line = endLine To tmpEndLine
    Call ProgressBar.showCount(Dir(filePath), line, tmpEndLine, "レベルの再設定")
    targetLevel = mainSheet.Range("B" & line)
    If targetLevel <> 0 Then
      mainSheet.Range(setVal("cell_Info") & line).InsertIndent targetLevel
    End If
    
    mainSheet.Range("A" & line).FormulaR1C1 = "=ROW()-5"
    
    Set taskLevelRange = Range(setVal("cell_TaskArea") & line)
    Range("B" & line).FormulaR1C1 = "=getIndentLevel(" & taskLevelRange.Address(ReferenceStyle:=xlR1C1) & ")"
    Set taskLevelRange = Nothing
  Next
  

  
  Application.CalculateFull
  
  Exit Function
'エラー発生時=====================================================================================
catchError:

End Function


'**************************************************************************************************
' * xxxxxxxxxx
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function makeCalendar()
  Dim line As Long, endLine As Long, tmpLine As Long
  Dim dataDirPath As String, filePath As String
  Dim workStartDay As Date, workEndDay As Date
  
'  On Error GoTo catchError
  Call init.setting
  mainSheet.Select
  
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


  Call Calendar.makeCalendar
  Application.CalculateFull
  Call Check.タスクリスト確認
  Call Chart.ガントチャート生成
  Application.Goto Reference:=Range("A1"), Scroll:=True
  
  
  
  Exit Function
'エラー発生時=====================================================================================
catchError:

End Function
