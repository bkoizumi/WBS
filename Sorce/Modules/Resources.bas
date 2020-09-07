Attribute VB_Name = "Resources"
'**************************************************************************************************
' * �f�[�^�ڍs
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function �f�[�^�ڍs()
  Dim line As Long, endLine As Long, startLine As Long, endColLine As Long
  Dim assignName As Collection
  
  
'  On Error GoTo catchError

  
  ResourcesSheet.Select
  Call WBS_Option.clearAll
  
  mainSheet.Calculate
  
  '�S���҃��X�g�̎擾
  Call Task.�S���Ғ��o(assignName)

  endLine = mainSheet.Cells(Rows.count, 1).End(xlUp).row
  rowLine = 6
  
    For Each assignor In assignName
      If assignor <> "�H��" Then
        For line = 6 To endLine
          If mainSheet.Range(setVal("cell_Assign") & line) Like "*" & assignor & "*" Then
            ResourcesSheet.Range("A" & rowLine) = mainSheet.Range("A" & line)
            ResourcesSheet.Range("B" & rowLine) = mainSheet.Range("B" & line).Value
            ResourcesSheet.Range("D" & rowLine) = assignor
            ResourcesSheet.Range("E" & rowLine) = mainSheet.Range(setVal("cell_TaskArea") & line)
            ResourcesSheet.Range("F" & rowLine) = mainSheet.Range(setVal("cell_PlanStart") & line)
            ResourcesSheet.Range("G" & rowLine) = mainSheet.Range(setVal("cell_PlanEnd") & line)
            ResourcesSheet.Range("H" & rowLine) = mainSheet.Range(setVal("cell_AchievementStart") & line)
            ResourcesSheet.Range("I" & rowLine) = mainSheet.Range(setVal("cell_AchievementEnd") & line)
            ResourcesSheet.Range("J" & rowLine) = mainSheet.Range(setVal("cell_ProgressLast") & line)
            ResourcesSheet.Range("K" & rowLine) = mainSheet.Range(setVal("cell_Progress") & line)
            
            For Each taskAllocation In Split(mainSheet.Range(setVal("cell_TaskAllocation") & line), ",")
              If taskAllocation Like "*" & assignor & "*" Then
                allocationRate = Split(taskAllocation, "<>")
                ResourcesSheet.Range("L" & rowLine) = allocationRate(1)
              End If
            Next
            
            ResourcesSheet.Range("M" & rowLine) = mainSheet.Range(setVal("cell_Task") & line)
            ResourcesSheet.Range("N" & rowLine) = mainSheet.Range(setVal("cell_TaskInfoP") & line)
            ResourcesSheet.Range("O" & rowLine) = mainSheet.Range(setVal("cell_TaskInfoC") & line)
            ResourcesSheet.Range("P" & rowLine) = mainSheet.Range(setVal("cell_WorkLoadP") & line)
            ResourcesSheet.Range("Q" & rowLine) = mainSheet.Range(setVal("cell_WorkLoadA") & line)
            ResourcesSheet.Range("R" & rowLine) = mainSheet.Range(setVal("cell_LateOrEarly") & line)
            ResourcesSheet.Range("S" & rowLine) = mainSheet.Range(setVal("cell_Note") & line)
            
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
  setVal("cell_TaskArea") = "E"
  setVal("cell_PlanStart") = "F"
  setVal("cell_PlanEnd") = "G"
  setVal("cell_Assign") = "D"
  setVal("cell_AchievementStart") = "H"
  setVal("cell_AchievementEnd") = "I"
  setVal("cell_ProgressLast") = "J"
  setVal("cell_Progress") = "K"
  
  setVal("cell_TaskAllocation") = "L"
  
  setVal("cell_Task") = "M"
  setVal("cell_TaskInfoP") = "N"
  setVal("cell_TaskInfoC") = "O"
  setVal("cell_WorkLoadP") = "P"
  setVal("cell_WorkLoadA") = "Q"
  setVal("cell_LateOrEarly") = "R"
  setVal("cell_Note") = "S"
  
  setVal("setLightning") = False
  
  
  Call Calendar.makeCalendar
  Call Chart.�K���g�`���[�g����

  endLine = Cells(Rows.count, 1).End(xlUp).row
  
  '�S���҂̃v���_�E���폜
  Range("D6:D" & endLine).Validation.Delete
  
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
    Range("B" & line) = Cells(line, getColumnNo(setVal("cell_TaskArea"))).IndentLevel + 1
  Next

  Call Library.endScript(True)

  Exit Function

'�G���[������=====================================================================================
catchError:
  Call Library.showNotice(Err.Number, Err.Description, True)
End Function
