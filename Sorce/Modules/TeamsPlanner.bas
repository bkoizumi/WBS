Attribute VB_Name = "TeamsPlanner"
'**************************************************************************************************
' * �f�[�^�ڍs
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function �f�[�^�ڍs()
  Dim line As Long, endLine As Long, startLine As Long, endColLine As Long
  Dim assignName As Collection

'  On Error GoTo catchError
  
  Call WBS_Option.clearAll
  
  sheetMain.Calculate
  
  '�S���҃��X�g�̎擾
  Call Task.�S���Ғ��o(assignName)

  endLine = sheetMain.Cells(Rows.count, 1).End(xlUp).row
  rowLine = 6
  sheetTeamsPlanner.Select
  
    For Each assignor In assignName
      If assignor <> "�H��" Then
        For line = 6 To endLine
          If sheetMain.Range(setVal("cell_Assign") & line) Like "*,*" Then
          
          ElseIf sheetMain.Range(setVal("cell_Assign") & line) = assignor Or (sheetMain.Range(setVal("cell_Assign") & line) = "" And assignor = "�����蓖��") Then
            sheetTeamsPlanner.Range("A" & rowLine) = sheetMain.Range("A" & line)
            sheetTeamsPlanner.Range("B" & rowLine) = sheetMain.Range(setVal("cell_LevelInfo") & line)
            
            sheetTeamsPlanner.Range("C" & rowLine) = sheetMain.Range(setVal("cell_Info") & line)
            sheetTeamsPlanner.Range("D" & rowLine) = sheetMain.Range(setVal("cell_LineInfo") & line)
            
            sheetTeamsPlanner.Range("E" & rowLine) = assignor
            sheetTeamsPlanner.Range("F" & rowLine) = sheetMain.Range(setVal("cell_TaskArea") & line)
            sheetTeamsPlanner.Range("G" & rowLine) = sheetMain.Range(setVal("cell_PlanStart") & line)
            sheetTeamsPlanner.Range("H" & rowLine) = sheetMain.Range(setVal("cell_PlanEnd") & line)
            sheetTeamsPlanner.Range("I" & rowLine) = sheetMain.Range(setVal("cell_AchievementStart") & line)
            sheetTeamsPlanner.Range("J" & rowLine) = sheetMain.Range(setVal("cell_AchievementEnd") & line)
            sheetTeamsPlanner.Range("K" & rowLine) = sheetMain.Range(setVal("cell_ProgressLast") & line)
            sheetTeamsPlanner.Range("L" & rowLine) = sheetMain.Range(setVal("cell_Progress") & line)
            
            sheetTeamsPlanner.Range("M" & rowLine) = sheetMain.Range(setVal("cell_TaskAllocation") & line)
            
            sheetTeamsPlanner.Range("N" & rowLine) = sheetMain.Range(setVal("cell_Task") & line)
            sheetTeamsPlanner.Range("O" & rowLine) = sheetMain.Range(setVal("cell_TaskInfoP") & line)
            sheetTeamsPlanner.Range("P" & rowLine) = sheetMain.Range(setVal("cell_TaskInfoC") & line)
            sheetTeamsPlanner.Range("Q" & rowLine) = sheetMain.Range(setVal("cell_WorkLoadP") & line)
            sheetTeamsPlanner.Range("R" & rowLine) = sheetMain.Range(setVal("cell_WorkLoadA") & line)
            sheetTeamsPlanner.Range("S" & rowLine) = sheetMain.Range(setVal("cell_LateOrEarly") & line)
            sheetTeamsPlanner.Range("T" & rowLine) = sheetMain.Range(setVal("cell_Note") & line)
            
            rowLine = rowLine + 1
          End If
        Next
      End If
    Next

  '�����̃R�s�[���y�[�X�g
  Rows("4:4").Copy
  Rows("6:" & rowLine - 1).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
  Application.CutCopyMode = False

  '���\�[�X�V�[�g�p�ɐݒ�l��ύX
  Call Check.���ڗ�`�F�b�N
  setVal("setLightning") = False
  
  
  
  Call Calendar.makeCalendar
  Call Chart.�K���g�`���[�g����

  endLine = Cells(Rows.count, 1).End(xlUp).row
  
  '�S���҂̃v���_�E���폜
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

  '���x���̍Đݒ�
  For line = 6 To endLine
    Range(setVal("cell_LevelInfo") & line) = Cells(line, getColumnNo(setVal("cell_TaskArea"))).IndentLevel + 1
  Next

  Call Library.endScript(True)

  Exit Function

'�G���[������--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(Err.Number, Err.Description, True)
End Function
