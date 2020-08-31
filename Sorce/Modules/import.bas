Attribute VB_Name = "import"
'���[�N�u�b�N�p�ϐ�------------------------------
Dim targetBook As Workbook

'���[�N�V�[�g�p�ϐ�------------------------------
'Dim masterSheet As Worksheet

'�O���[�o���ϐ�----------------------------------



'**************************************************************************************************
' * import�p�@�\
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function �t�@�C���C���|�[�g()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  Dim dataDirPath As String, filePath As String
  
  
'  On Error GoTo catchError
  Call init.setting
  mainSheet.Select
  
  dataDirPath = Library.getRegistry("dataDirPath")
  If dataDirPath = "" Then
    dataDirPath = init.ThisBook.Path
  End If
  
  Call Library.showNotice(1, "�v���W�F�N�g")
  
  filePaths = Library.getFilesPath(dataDirPath, "", "�v���W�F�N�g���Ƃ̃t�@�C����I�����Ă�������", 1)
  If filePaths(0) = "" Then
    Call Library.showDebugForm("�t�@�C���C���|�[�g", "�t�@�C���I���L�����Z��")
    Call Library.showNotice(100, , True)
    End
  End If

  For i = 0 To UBound(filePaths)
    filePath = filePaths(i)
    Call Library.showDebugForm("�t�@�C���C���|�[�g", "�ΏہF" & Dir(filePath))
    Call ProgressBar.showCount("�t�@�C���C���|�[�g", i + 1, UBound(filePaths) + 1, "�ΏہF" & Dir(filePath))
    
    If setVal("workMode") = "default" Then
      Call �f�[�^�R�s�[(filePath)
    ElseIf setVal("workMode") = "CD���p" Then
      Call CD���p.�f�[�^�R�s�[(filePath)
    End If
  Next

  dataDirPath = Replace(filePath, "\" & Dir(filePath), "")
  Call Library.setRegistry("dataDirPath", dataDirPath)

  Exit Function
'�G���[������=====================================================================================
catchError:


End Function


'**************************************************************************************************
' * �f�[�^�R�s�[
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function �f�[�^�R�s�[(filePath As String)
  Dim line As Long, endLine As Long, targetEndLine As Long, endColLine As Long
  Dim tmpEndLine As Long
  
'  On Error GoTo catchError

  Set targetBook = Workbooks.Open(filePath, , True)
  Call Library.startScript
  targetBook.Activate
  
  If Library.chkSheetName("calendar") = True Then
    Call Library.showDebugForm("�t�@�C���C���|�[�g", "calendar �V�[�g����")
    
    targetBook.Sheets("calendar").Select
    targetBook.Worksheets("calendar").Range("B2").Copy
    
    endLine = mainSheet.Cells(Rows.count, 1).End(xlUp).row + 1
    mainSheet.Range("C" & endLine).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
    '�t�@�C��������l�Ɋi�[
    Call Library.showDebugForm("�t�@�C���C���|�[�g", "�t�@�C��������l�Ɋi�[")
    mainSheet.Range(setVal("cell_Note") & endLine) = Dir(filePath)
    
    
    Call Library.showDebugForm("�t�@�C���C���|�[�g", "B6�Z���R�s�[")
    targetEndLine = Cells(Rows.count, 2).End(xlUp).row - 1
    targetBook.Worksheets("calendar").Range("B6").Copy
    endLine = endLine + 1
    
    mainSheet.Range("D" & endLine).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
    
    Call Library.showDebugForm("�t�@�C���C���|�[�g", "B��R�s�[")
    targetBook.Worksheets("calendar").Range("B7:B" & targetEndLine).Copy
    mainSheet.Range("E" & endLine + 1).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
    Call Library.showDebugForm("�t�@�C���C���|�[�g", "C��R�s�[")
    targetBook.Worksheets("calendar").Range("C6:D" & targetEndLine).Copy
    mainSheet.Range(setVal("cell_PlanStart") & endLine).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
  
  End If
  
  Call Library.showDebugForm("�t�@�C���C���|�[�g", "WBS�V�[�g A��ݒ�")
  tmpEndLine = Cells(Rows.count, 5).End(xlUp).row
  mainSheet.Range("A6" & ":A" & tmpEndLine).FormulaR1C1 = "=ROW()-5"
  mainSheet.Range("B6" & ":B" & tmpEndLine).FormulaR1C1 = _
      "=IF(RC[1]<>"""",1,IF(RC[2]<>"""",2,IF(RC[3]<>"""",3,IF(RC[4]<>"""",4,IF(RC[5]<>"""",5,IF(RC[6]<>"""",6,""""))))))"
    
  targetBook.Close

  ThisWorkbook.Activate
  mainSheet.Select
  endLine = mainSheet.Cells(Rows.count, 1).End(xlUp).row
  mainSheet.Range("B6:B" & endLine).Select

  
  Exit Function
'�G���[������=====================================================================================
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
  Call Check.�^�X�N���X�g�m�F
  Call Chart.�K���g�`���[�g����
  Application.Goto Reference:=Range("A1"), Scroll:=True
  
  
  
  Exit Function
'�G���[������=====================================================================================
catchError:

End Function
