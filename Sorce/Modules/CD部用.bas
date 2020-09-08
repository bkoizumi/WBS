Attribute VB_Name = "CD部用"

Function データコピー(filePath As String)
  Dim line As Long, endLine As Long, targetEndLine As Long, endColLine As Long
  Dim tmpEndLine As Long, lineCount As Long
  Dim targetCalSheet As Worksheet
  Dim taskLevelRange As Range
  
'  On Error GoTo catchError

  Set targetBook = Workbooks.Open(filePath, , True)
  targetBook.Activate
  Call Library.startScript
  
  If Library.chkSheetName("calendar") = True Then
    Call Library.showDebugForm("ファイルインポート", "　calendar シート発見")
    
    Set targetCalSheet = targetBook.Worksheets("calendar")
    targetCalSheet.Select
    targetCalSheet.Range("B2").Copy
    
    endLine = mainSheet.Cells(Rows.count, 1).End(xlUp).row + 1
    mainSheet.Range(setVal("cell_Info") & endLine).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    mainSheet.Range(setVal("cell_Assign") & endLine) = "工程"
    mainSheet.Range(setVal("cell_Assign") & endLine) = "工程"
    
    'ファイル名を備考に格納
    Call Library.showDebugForm("ファイルインポート", "　ファイル名を備考に格納")
    mainSheet.Range(setVal("cell_Note") & endLine) = Dir(filePath)
    
    endLine = endLine + 1
    targetEndLine = Cells(Rows.count, 2).End(xlUp).row - 1
    For line = 6 To targetEndLine
    
      Call Library.showDebugForm("ファイルインポート", "　" & targetCalSheet.Range("B" & line))
      Call ProgressBar.showCount(Dir(filePath), line, targetEndLine, targetCalSheet.Range("B" & line))
      
      If targetCalSheet.Range("B" & line) Like "<*" Then
        Call Library.showDebugForm("ファイルインポート", "　　タスク名設定")
        targetCalSheet.Range("B" & line).Copy
        mainSheet.Range(setVal("cell_Info") & endLine).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        mainSheet.Range(setVal("cell_Info") & endLine).InsertIndent 2
      
        targetCalSheet.Range(setVal("cell_Info") & line & ":D" & line).Copy
        mainSheet.Range(setVal("cell_PlanStart") & endLine).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
      
      ElseIf targetCalSheet.Range("B" & line) <> "" Then
        Call Library.showDebugForm("ファイルインポート", "　　工程設定")
        targetCalSheet.Range("B" & line).Copy
        mainSheet.Range(setVal("cell_Info") & endLine).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        mainSheet.Range(setVal("cell_Info") & endLine).InsertIndent 1
        
      
        targetCalSheet.Range(setVal("cell_Info") & line & ":D" & line).Copy
        mainSheet.Range(setVal("cell_PlanStart") & endLine).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
      
        mainSheet.Range(setVal("cell_Assign") & endLine) = "工程"
        mainSheet.Range(setVal("cell_Assign") & endLine) = "工程"
      End If
      
      '要員設定
      If targetCalSheet.Range("B" & line) Like "<TCI>*" Then
        Call Library.showDebugForm("ファイルインポート", "　　要員設定")
        Select Case True
          Case targetCalSheet.Range("B" & line) Like "*構成案*"
            mainSheet.Range(setVal("cell_Assign") & endLine) = "[TBD]PL"
          Case targetCalSheet.Range("B" & line) Like "*デザイン*"
            mainSheet.Range(setVal("cell_Assign") & endLine) = "[TBD]De"
          Case targetCalSheet.Range("B" & line) Like "*コーディング*"
            mainSheet.Range(setVal("cell_Assign") & endLine) = "[TBD]HT"
          Case targetCalSheet.Range("B" & line) Like "*公開*"
            mainSheet.Range(setVal("cell_Assign") & endLine) = "公開"
          
          
          Case Else
            mainSheet.Range(setVal("cell_Assign") & endLine) = "TBD"
        End Select
      ElseIf targetCalSheet.Range("B" & line) Like "<御社>*" Then
        Call Library.showDebugForm("ファイルインポート", "　　要員設定")
        mainSheet.Range(setVal("cell_Assign") & endLine) = "○社A様"
      End If
      endLine = endLine + 1
    Next
  
    Set taskLevelRange = Range(setVal("cell_TaskArea") & endLine)
    Range("B" & endLine).FormulaR1C1 = "=getIndentLevel(" & taskLevelRange.Address(ReferenceStyle:=xlR1C1) & ")"
    Set taskLevelRange = Nothing
  
  End If
  
  targetBook.Close
  ThisWorkbook.Activate
  Call Library.startScript
  mainSheet.Select
  
  Call Library.showDebugForm("ファイルインポート", "　WBSシート A、B列設定")
  tmpEndLine = Cells(Rows.count, 3).End(xlUp).row
  
  mainSheet.Range("A6" & ":A" & tmpEndLine).FormulaR1C1 = "=ROW()-5"
  

'  endLine = mainSheet.Cells(Rows.count, 1).End(xlUp).row
'  mainSheet.Range("B6:B" & endLine).Select
  Call Library.showDebugForm("ファイルインポート", "コピー完了")

  
  Exit Function
'エラー発生時=====================================================================================
catchError:

End Function
