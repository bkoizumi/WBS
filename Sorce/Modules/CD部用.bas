Attribute VB_Name = "CD���p"

Function �f�[�^�R�s�[(filePath As String)
  Dim line As Long, endLine As Long, targetEndLine As Long, endColLine As Long
  Dim tmpEndLine As Long, lineCount As Long
  Dim targetCalSheet As Worksheet
  Dim taskLevelRange As Range
  
'  On Error GoTo catchError

  Call Library.showDebugForm("�t�@�C���C���|�[�g", "CD���p�t�@�C��")
  Set targetBook = Workbooks.Open(filePath, , True)
  targetBook.Activate
  Call Library.startScript
  
  If Library.chkSheetName("calendar") = True Then
    
    
    Set targetCalSheet = targetBook.Worksheets("calendar")
    targetCalSheet.Select
    targetCalSheet.Range("B2").Copy
    
    endLine = mainSheet.Cells(Rows.count, 1).End(xlUp).row + 1
    mainSheet.Range(setVal("cell_TaskArea") & endLine).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    mainSheet.Range(setVal("cell_Assign") & endLine) = "�H��"
    mainSheet.Range(setVal("cell_Assign") & endLine) = "�H��"
    mainSheet.Range("A" & endLine) = endLine - 5
    mainSheet.Range(setVal("cell_LineInfo") & endLine).FormulaR1C1 = "=ROW()-5"
    
    '�t�@�C��������l�Ɋi�[
    mainSheet.Range(setVal("cell_Note") & endLine) = Dir(filePath)
    
    endLine = endLine + 1
    targetEndLine = Cells(Rows.count, 2).End(xlUp).row - 1
    For line = 6 To targetEndLine
      If targetCalSheet.Range("B" & line) <> "" Then
        Call ProgressBar.showCount(Dir(filePath), line, targetEndLine, targetCalSheet.Range("B" & line))
        
        If targetCalSheet.Range("B" & line) Like "<*" Then
          targetCalSheet.Range("B" & line).Copy
          mainSheet.Range(setVal("cell_TaskArea") & endLine).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
          mainSheet.Range(setVal("cell_TaskArea") & endLine).InsertIndent 2
        
          targetCalSheet.Range("C" & line & ":D" & line).Copy
          mainSheet.Range(setVal("cell_PlanStart") & endLine).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        
        ElseIf targetCalSheet.Range("B" & line) <> "" Then
          targetCalSheet.Range("B" & line).Copy
          mainSheet.Range(setVal("cell_TaskArea") & endLine).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
          mainSheet.Range(setVal("cell_TaskArea") & endLine).InsertIndent 1

          
        
          targetCalSheet.Range("C" & line & ":D" & line).Copy
          mainSheet.Range(setVal("cell_PlanStart") & endLine).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        
          mainSheet.Range(setVal("cell_Assign") & endLine) = "�H��"
          mainSheet.Range(setVal("cell_Assign") & endLine) = "�H��"
        End If
        
        '�v���ݒ�
'        If targetCalSheet.Range("B" & line) Like "<TCI>*" Then
'          Select Case True
'            Case targetCalSheet.Range("B" & line) Like "*�\����*"
'              mainSheet.Range(setVal("cell_Assign") & endLine) = "[TBD]PL"
'            Case targetCalSheet.Range("B" & line) Like "*�f�U�C��*"
'              mainSheet.Range(setVal("cell_Assign") & endLine) = "[TBD]De"
'            Case targetCalSheet.Range("B" & line) Like "*�R�[�f�B���O*"
'              mainSheet.Range(setVal("cell_Assign") & endLine) = "[TBD]HT"
'            Case targetCalSheet.Range("B" & line) Like "*���J*"
'              mainSheet.Range(setVal("cell_Assign") & endLine) = "���J"
'            Case Else
'              mainSheet.Range(setVal("cell_Assign") & endLine) = "TBD"
'          End Select
'        ElseIf targetCalSheet.Range("B" & line) Like "<���>*" Then
'          mainSheet.Range(setVal("cell_Assign") & endLine) = "����A�l"
'        End If
        
        
        mainSheet.Range("A" & endLine) = endLine - 5
        mainSheet.Range(setVal("cell_LineInfo") & endLine).FormulaR1C1 = "=ROW()-5"
        
        endLine = endLine + 1
      End If
    Next
  
    Set taskLevelRange = Range(setVal("cell_TaskArea") & endLine)
    Range("B" & endLine).Formula = "=getIndentLevel(" & taskLevelRange.Address(ReferenceStyle:=xlA1, RowAbsolute:=False, ColumnAbsolute:=False) & ")"
    Set taskLevelRange = Nothing
  
  Else
    Call Library.showNotice(100, , True)
    End
  End If
  
  targetBook.Close
  ThisWorkbook.Activate
  Call Library.startScript
  mainSheet.Select
  
  Call WBS_Option.�s�ԍ��Đݒ�
  Call WBS_Option.�^�X�N���x���̐ݒ�
  
  Call Library.showDebugForm("�t�@�C���C���|�[�g", "�R�s�[����")

  
  Exit Function
'�G���[������=====================================================================================
catchError:

End Function
