Attribute VB_Name = "CD部用"

Function データコピー(filePath As String)
  Dim line As Long, endLine As Long, targetEndLine As Long, endColLine As Long
  Dim tmpEndLine As Long, lineCount As Long
  Dim targetCalSheet As Worksheet
  Dim taskLevelRange As Range
  
  On Error GoTo catchError

  Call Library.showDebugForm("ファイルインポート", "CD部用ファイル")
  
  Set targetCalSheet = targetBook.Worksheets("calendar")
  targetBook.Activate
  targetCalSheet.Select
  targetCalSheet.Range("B2").Copy
  
  endLine = sheetMain.Cells(Rows.count, 1).End(xlUp).row + 1
  sheetMain.Range(setVal("cell_TaskArea") & endLine).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
  sheetMain.Range(setVal("cell_Assign") & endLine) = "工程"
  sheetMain.Range(setVal("cell_Assign") & endLine) = "工程"
  sheetMain.Range("A" & endLine) = endLine - 5
  sheetMain.Range(setVal("cell_LineInfo") & endLine).FormulaR1C1 = "=ROW()-5"
  
  'ファイル名を備考に格納
  sheetMain.Range(setVal("cell_Note") & endLine) = filePath
  
  endLine = endLine + 1
  targetEndLine = Cells(Rows.count, 2).End(xlUp).row - 1
  For line = 6 To targetEndLine
    If targetCalSheet.Range("B" & line) <> "" Then
      Call ctl_ProgressBar.showCount(Dir(filePath), line, targetEndLine, targetCalSheet.Range("B" & line))
      
      If targetCalSheet.Range("B" & line) Like "<*" Then
        sheetMain.Range(setVal("cell_TaskArea") & endLine) = targetCalSheet.Range("B" & line).Value
        sheetMain.Range(setVal("cell_PlanStart") & endLine) = targetCalSheet.Range("C" & line).Value
        sheetMain.Range(setVal("cell_PlanEnd") & endLine) = targetCalSheet.Range("D" & line).Value
        sheetMain.Range(setVal("cell_TaskArea") & endLine).InsertIndent 2
        
        
      ElseIf targetCalSheet.Range("B" & line) <> "" Then
        sheetMain.Range(setVal("cell_TaskArea") & endLine) = targetCalSheet.Range("B" & line).Value
        sheetMain.Range(setVal("cell_PlanStart") & endLine) = targetCalSheet.Range("C" & line).Value
        sheetMain.Range(setVal("cell_PlanEnd") & endLine) = targetCalSheet.Range("D" & line).Value
        
        sheetMain.Range(setVal("cell_TaskArea") & endLine).InsertIndent 1
        sheetMain.Range(setVal("cell_Assign") & endLine) = "工程"
        sheetMain.Range(setVal("cell_Assign") & endLine) = "工程"
      End If
        
      '要員設定
'      If targetCalSheet.Range("B" & line) Like "<TCI>*" Then
'        Select Case True
'          Case targetCalSheet.Range("B" & line) Like "*構成案*"
'            sheetMain.Range(setVal("cell_Assign") & endLine) = "[TBD]PL"
'          Case targetCalSheet.Range("B" & line) Like "*デザイン*"
'            sheetMain.Range(setVal("cell_Assign") & endLine) = "[TBD]De"
'          Case targetCalSheet.Range("B" & line) Like "*コーディング*"
'            sheetMain.Range(setVal("cell_Assign") & endLine) = "[TBD]HT"
'          Case targetCalSheet.Range("B" & line) Like "*公開*"
'            sheetMain.Range(setVal("cell_Assign") & endLine) = "公開"
'          Case Else
'            sheetMain.Range(setVal("cell_Assign") & endLine) = "TBD"
'        End Select
'      ElseIf targetCalSheet.Range("B" & line) Like "<御社>*" Then
'        sheetMain.Range(setVal("cell_Assign") & endLine) = "○社A様"
'      End If
      
      
      sheetMain.Range("A" & endLine) = endLine - 5
      sheetMain.Range(setVal("cell_LineInfo") & endLine).FormulaR1C1 = "=ROW()-5"
      
      endLine = endLine + 1
    End If
  Next
  
  Set taskLevelRange = Range(setVal("cell_TaskArea") & endLine)
  Range("B" & endLine).Formula = "=getIndentLevel(" & taskLevelRange.Address(ReferenceStyle:=xlA1, RowAbsolute:=False, ColumnAbsolute:=False) & ")"
  Set taskLevelRange = Nothing
  
  targetBook.Close
  ThisWorkbook.Activate
  Call Library.startScript
  sheetMain.Select
  
  Call WBS_Option.行番号再設定
  Call WBS_Option.タスクレベルの設定
  
  Call Library.showDebugForm("ファイルインポート", "コピー完了")

  
  Exit Function
'エラー発生時--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(Err.Number, Err.Description, True)
End Function
