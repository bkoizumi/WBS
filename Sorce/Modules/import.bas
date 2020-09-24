Attribute VB_Name = "import"
'���[�N�u�b�N�p�ϐ�------------------------------


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
    
    '�w��t�@�C���I�[�v�����A�V�[�g�̑��݊m�F
    Set targetBook = Workbooks.Open(FileName:=filePath, ReadOnly:=True)
    Windows(targetBook.Name).WindowState = xlMinimized
    Call Library.startScript
    targetBook.Activate
    
    If Library.chkSheetName("���C��") = True Then
      Call �f�[�^�R�s�[(filePath)
    ElseIf Library.chkSheetName("calendar") = True Then
      Call CD���p.�f�[�^�R�s�[(filePath)
    Else
      Call Library.showNotice(405, "�Y����", True)
      End
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
  Dim targetSetVal As Collection
  Dim targetLevel As Long
  Dim prgbarMeg As String
  Dim prgbarCnt As Long
  Dim taskLevelRange As Range
  
  
'  On Error GoTo catchError
  
  Set targetSetVal = New Collection
  prgbarCnt = 0

  Call ProgressBar.showCount("�t�@�C���C���|�[�g", prgbarCnt, 100, "�ΏہF" & Dir(filePath))
  
  If Library.chkSheetName("���C��") = True Then
    Call Library.showDebugForm("�t�@�C���C���|�[�g", "WBS �V�[�g����")
    
    '�C���|�[�g�t�@�C���̐ݒ�ǂݍ���
    With targetSetVal
      For line = 3 To targetBook.Sheets("�ݒ�").Cells(Rows.count, 1).End(xlUp).row
        If targetBook.Sheets("�ݒ�").Range("A" & line) <> "" Then
         .Add item:=targetBook.Sheets("�ݒ�").Range(setVal("cell_LevelInfo") & line), Key:=targetBook.Sheets("�ݒ�").Range("A" & line)
        End If
      Next
    End With
  
    endLine = mainSheet.Cells(Rows.count, 1).End(xlUp).row + 1
    
    '�t�@�C�������^�X�N�Ƃ��ēo�^
    prgbarMeg = "�t�@�C�������^�X�N�Ƃ��ēo�^"
    prgbarCnt = prgbarCnt + 1
    Call Library.showDebugForm(Dir(filePath), prgbarMeg)
    Call ProgressBar.showCount(Dir(filePath), prgbarCnt, 11, prgbarMeg)
    
    mainSheet.Range("B" & endLine) = 1
    mainSheet.Range(setVal("cell_TaskArea") & endLine) = Dir(filePath)
    mainSheet.Range(setVal("cell_Note") & endLine) = filePath

    endLine = endLine + 1
    
    Call Library.showDebugForm("�t�@�C���C���|�[�g", "�C���|�[�g�J�n")
    targetEndLine = targetBook.Worksheets(mainSheetName).Cells(Rows.count, 1).End(xlUp).row
    
    '#�`�^�X�N�����R�s�[
    prgbarMeg = "�^�X�N����܂ł��R�s�["
    prgbarCnt = prgbarCnt + 1
    Call Library.showDebugForm(Dir(filePath), prgbarMeg)
    Call ProgressBar.showCount(Dir(filePath), prgbarCnt, 11, prgbarMeg)
    
    targetBook.Worksheets(mainSheetName).Range("A6:" & targetSetVal("cell_TaskArea") & targetEndLine).Copy
    mainSheet.Range("A" & endLine).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
    '�\������R�s�[
    prgbarMeg = "�\������R�s�["
    prgbarCnt = prgbarCnt + 1
    Call Library.showDebugForm(Dir(filePath), prgbarMeg)
    Call ProgressBar.showCount(Dir(filePath), prgbarCnt, 11, prgbarMeg)

    targetBook.Worksheets(mainSheetName).Range(targetSetVal("cell_PlanStart") & "6:" & targetSetVal("cell_PlanEnd") & targetEndLine).Copy
    mainSheet.Range(setVal("cell_PlanStart") & endLine).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
    '�S���҂��R�s�[
    prgbarMeg = "�S���҂��R�s�["
    prgbarCnt = prgbarCnt + 1
    Call Library.showDebugForm(Dir(filePath), prgbarMeg)
    Call ProgressBar.showCount(Dir(filePath), prgbarCnt, 11, prgbarMeg)
    
    targetBook.Worksheets(mainSheetName).Range(targetSetVal("cell_Assign") & "6:" & targetSetVal("cell_Assign") & targetEndLine).Copy
    mainSheet.Range(setVal("cell_Assign") & endLine).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        
    '���ѓ����R�s�[
    prgbarMeg = "���ѓ����R�s�["
    prgbarCnt = prgbarCnt + 1
    Call Library.showDebugForm(Dir(filePath), prgbarMeg)
    Call ProgressBar.showCount(Dir(filePath), prgbarCnt, 11, prgbarMeg)
    
    targetBook.Worksheets(mainSheetName).Range(targetSetVal("cell_AchievementStart") & "6:" & targetSetVal("cell_AchievementEnd") & targetEndLine).Copy
    mainSheet.Range(setVal("cell_AchievementStart") & endLine).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        
    '�i�������R�s�[
    prgbarMeg = "A�`C��R�s�["
    prgbarCnt = prgbarCnt + 1
    Call Library.showDebugForm(Dir(filePath), prgbarMeg)
    Call ProgressBar.showCount(Dir(filePath), prgbarCnt, 11, prgbarMeg)
    
    targetBook.Worksheets(mainSheetName).Range(targetSetVal("cell_ProgressLast") & "6:" & targetSetVal("cell_Progress") & targetEndLine).Copy
    mainSheet.Range(setVal("cell_ProgressLast") & endLine).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        
    '��s�^�X�N���R�s�[
    prgbarMeg = "��s�^�X�N���R�s�["
    prgbarCnt = prgbarCnt + 1
    Call Library.showDebugForm(Dir(filePath), prgbarMeg)
    Call ProgressBar.showCount(Dir(filePath), prgbarCnt, 11, prgbarMeg)
    
    targetBook.Worksheets(mainSheetName).Range(targetSetVal("cell_Task") & "6:" & targetSetVal("cell_Task") & targetEndLine).Copy
    mainSheet.Range(setVal("cell_Task") & endLine).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        
    '�^�X�N�z�����R�s�[
    prgbarMeg = "�^�X�N�z�����R�s�["
    prgbarCnt = prgbarCnt + 1
    Call Library.showDebugForm(Dir(filePath), prgbarMeg)
    Call ProgressBar.showCount(Dir(filePath), prgbarCnt, 11, prgbarMeg)
    
    targetBook.Worksheets(mainSheetName).Range(targetSetVal("cell_TaskAllocation") & "6:" & targetSetVal("cell_TaskAllocation") & targetEndLine).Copy
    mainSheet.Range(setVal("cell_TaskAllocation") & endLine).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        
    '��ƍH�����R�s�[
    prgbarMeg = "��ƍH�����R�s�["
    prgbarCnt = prgbarCnt + 1
    Call Library.showDebugForm(Dir(filePath), prgbarMeg)
    Call ProgressBar.showCount(Dir(filePath), prgbarCnt, 11, prgbarMeg)
    
    targetBook.Worksheets(mainSheetName).Range(targetSetVal("cell_WorkLoadP") & "6:" & targetSetVal("cell_WorkLoadA") & targetEndLine).Copy
    mainSheet.Range(setVal("cell_WorkLoadP") & endLine).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        
    '�x���H�����R�s�[
    prgbarMeg = "�x���H�����R�s�["
    prgbarCnt = prgbarCnt + 1
    Call Library.showDebugForm(Dir(filePath), prgbarMeg)
    Call ProgressBar.showCount(Dir(filePath), prgbarCnt, 11, prgbarMeg)
    
    targetBook.Worksheets(mainSheetName).Range(targetSetVal("cell_LateOrEarly") & "6:" & targetSetVal("cell_LateOrEarly") & targetEndLine).Copy
    mainSheet.Range(setVal("cell_LateOrEarly") & endLine).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
         
    '���l���R�s�[
    prgbarMeg = "���l���R�s�["
    prgbarCnt = prgbarCnt + 1
    Call Library.showDebugForm(Dir(filePath), prgbarMeg)
    Call ProgressBar.showCount(Dir(filePath), prgbarCnt, 11, prgbarMeg)
    
    targetBook.Worksheets(mainSheetName).Range(targetSetVal("cell_Note") & "6:" & targetSetVal("cell_Note") & targetEndLine).Copy
    mainSheet.Range(setVal("cell_Note") & endLine).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
  End If
    
  Set targetSetVal = Nothing
  targetBook.Close
  
  ThisWorkbook.Activate
  mainSheet.Select
  Call Library.startScript
  
  Call Library.showDebugForm("�t�@�C���C���|�[�g", "WBS�V�[�g A��ݒ�")
  tmpEndLine = Cells(Rows.count, 1).End(xlUp).row
  
  '���x���̍Đݒ�
  Call Library.showDebugForm(Dir(filePath), "���x���̍Đݒ�")
  For line = endLine To tmpEndLine
    Call ProgressBar.showCount(Dir(filePath), line, tmpEndLine, "���x���̍Đݒ�")
    targetLevel = mainSheet.Range(setVal("cell_LevelInfo") & line) + 1
    mainSheet.Range(setVal("cell_LevelInfo") & line) = targetLevel
    If targetLevel <> 0 Then
      mainSheet.Range(setVal("cell_Info") & line).InsertIndent targetLevel
    End If
  Next
  Application.CalculateFull
  
  Exit Function
'�G���[������=====================================================================================
catchError:

End Function


'**************************************************************************************************
' * �J�����_�[�p�����擾
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function �J�����_�[�p�����擾()
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
  
  If workStartDay <= Date And Date <= workEndDay Then
    Range("baseDay") = Date
  End If

  Call Calendar.makeCalendar
  Application.CalculateFull
  Call Check.�^�X�N���X�g�m�F
  Call Chart.�K���g�`���[�g����
  
  Exit Function
'�G���[������=====================================================================================
catchError:

End Function
