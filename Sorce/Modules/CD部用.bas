Attribute VB_Name = "CD���p"

Function �f�[�^�R�s�[(filePath As String)
  Dim line As Long, endLine As Long, targetEndLine As Long, endColLine As Long
  Dim tmpEndLine As Long, lineCount As Long
  Dim targetCalSheet As Worksheet
  Dim taskLevelRange As Range
  
'  On Error GoTo catchError

  Set targetBook = Workbooks.Open(filePath, , True)
  targetBook.Activate
  Call Library.startScript
  
  If Library.chkSheetName("calendar") = True Then
    Call Library.showDebugForm("�t�@�C���C���|�[�g", "�@calendar �V�[�g����")
    
    Set targetCalSheet = targetBook.Worksheets("calendar")
    targetCalSheet.Select
    targetCalSheet.Range("B2").Copy
    
    endLine = mainSheet.Cells(Rows.count, 1).End(xlUp).row + 1
    mainSheet.Range(setVal("cell_Info") & endLine).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    mainSheet.Range(setVal("cell_Assign") & endLine) = "�H��"
    mainSheet.Range(setVal("cell_Assign") & endLine) = "�H��"
    
    '�t�@�C��������l�Ɋi�[
    Call Library.showDebugForm("�t�@�C���C���|�[�g", "�@�t�@�C��������l�Ɋi�[")
    mainSheet.Range(setVal("cell_Note") & endLine) = Dir(filePath)
    
    endLine = endLine + 1
    targetEndLine = Cells(Rows.count, 2).End(xlUp).row - 1
    For line = 6 To targetEndLine
    
      Call Library.showDebugForm("�t�@�C���C���|�[�g", "�@" & targetCalSheet.Range(setVal("cell_LevelInfo") & line))
      Call ProgressBar.showCount(Dir(filePath), line, targetEndLine, targetCalSheet.Range(setVal("cell_LevelInfo") & line))
      
      If targetCalSheet.Range(setVal("cell_LevelInfo") & line) Like "<*" Then
        Call Library.showDebugForm("�t�@�C���C���|�[�g", "�@�@�^�X�N���ݒ�")
        targetCalSheet.Range(setVal("cell_LevelInfo") & line).Copy
        mainSheet.Range(setVal("cell_Info") & endLine).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        mainSheet.Range(setVal("cell_Info") & endLine).InsertIndent 2
      
        targetCalSheet.Range(setVal("cell_Info") & line & ":D" & line).Copy
        mainSheet.Range(setVal("cell_PlanStart") & endLine).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
      
      ElseIf targetCalSheet.Range(setVal("cell_LevelInfo") & line) <> "" Then
        Call Library.showDebugForm("�t�@�C���C���|�[�g", "�@�@�H���ݒ�")
        targetCalSheet.Range(setVal("cell_LevelInfo") & line).Copy
        mainSheet.Range(setVal("cell_Info") & endLine).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        mainSheet.Range(setVal("cell_Info") & endLine).InsertIndent 1
        
      
        targetCalSheet.Range(setVal("cell_Info") & line & ":D" & line).Copy
        mainSheet.Range(setVal("cell_PlanStart") & endLine).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
      
        mainSheet.Range(setVal("cell_Assign") & endLine) = "�H��"
        mainSheet.Range(setVal("cell_Assign") & endLine) = "�H��"
      End If
      
      '�v���ݒ�
      If targetCalSheet.Range(setVal("cell_LevelInfo") & line) Like "<TCI>*" Then
        Call Library.showDebugForm("�t�@�C���C���|�[�g", "�@�@�v���ݒ�")
        Select Case True
          Case targetCalSheet.Range(setVal("cell_LevelInfo") & line) Like "*�\����*"
            mainSheet.Range(setVal("cell_Assign") & endLine) = "[TBD]PL"
          Case targetCalSheet.Range(setVal("cell_LevelInfo") & line) Like "*�f�U�C��*"
            mainSheet.Range(setVal("cell_Assign") & endLine) = "[TBD]De"
          Case targetCalSheet.Range(setVal("cell_LevelInfo") & line) Like "*�R�[�f�B���O*"
            mainSheet.Range(setVal("cell_Assign") & endLine) = "[TBD]HT"
          Case targetCalSheet.Range(setVal("cell_LevelInfo") & line) Like "*���J*"
            mainSheet.Range(setVal("cell_Assign") & endLine) = "���J"
          
          
          Case Else
            mainSheet.Range(setVal("cell_Assign") & endLine) = "TBD"
        End Select
      ElseIf targetCalSheet.Range(setVal("cell_LevelInfo") & line) Like "<���>*" Then
        Call Library.showDebugForm("�t�@�C���C���|�[�g", "�@�@�v���ݒ�")
        mainSheet.Range(setVal("cell_Assign") & endLine) = "����A�l"
      End If
      endLine = endLine + 1
    Next
  
    Set taskLevelRange = Range(setVal("cell_TaskArea") & endLine)
    Range("B" & endLine).Formula = "=getIndentLevel(" & taskLevelRange.Address(ReferenceStyle:=xlA1, RowAbsolute:=False, ColumnAbsolute:=False) & ")"
    Set taskLevelRange = Nothing
  
  End If
  
  targetBook.Close
  ThisWorkbook.Activate
  Call Library.startScript
  mainSheet.Select
  
  Call Library.showDebugForm("�t�@�C���C���|�[�g", "�@WBS�V�[�g A�AB��ݒ�")
  tmpEndLine = Cells(Rows.count, 3).End(xlUp).row
  
  mainSheet.Range("A6" & ":A" & tmpEndLine).FormulaR1C1 = "=ROW()-5"
  

'  endLine = mainSheet.Cells(Rows.count, 1).End(xlUp).row
'  mainSheet.Range("B6:B" & endLine).Select
  Call Library.showDebugForm("�t�@�C���C���|�[�g", "�R�s�[����")

  
  Exit Function
'�G���[������=====================================================================================
catchError:

End Function
