Attribute VB_Name = "Check"
'**************************************************************************************************
' * ���ڃ`�F�b�N
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function ���ڗ�`�F�b�N()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  Dim itemName As String
  Dim defaultLine As Long
  
  Call init.setting
  defaultLine = sheetSetting.Range("B5")
  startLine = 4

  sheetSetting.Range("A" & defaultLine & ":B100").ClearContents
  
'  sheetMain.Select
 
  For colLine = 1 To 20
    If Cells(2, colLine) <> "" Then
      itemName = Cells(2, colLine)
    Else
      GoTo Label_nextFor
    End If
    
    line = sheetSetting.Cells(Rows.count, 1).End(xlUp).row + 1
    If line < defaultLine Then
      line = defaultLine
    End If
    
    Select Case itemName
      Case "#"
      Case "Lv"
      Case "Info"
      Case "�^�X�N��"
        sheetSetting.Range("A" & line) = "cell_TaskArea"
        sheetSetting.Range("B" & line) = Library.getColumnName(colLine)
    
    
      Case "�\���"
        sheetSetting.Range("A" & line) = "cell_PlanStart"
        sheetSetting.Range("B" & line) = Library.getColumnName(colLine)
        
        sheetSetting.Range("A" & line + 1) = "cell_PlanEnd"
        sheetSetting.Range("B" & line + 1) = Library.getColumnName(colLine + 1)
      
      Case "�S����"
        sheetSetting.Range("A" & line) = "cell_Assign"
        sheetSetting.Range("B" & line) = Library.getColumnName(colLine)

      Case "���ѓ�"
        sheetSetting.Range("A" & line) = "cell_AchievementStart"
        sheetSetting.Range("B" & line) = Library.getColumnName(colLine)
        
        sheetSetting.Range("A" & line + 1) = "cell_AchievementEnd"
        sheetSetting.Range("B" & line + 1) = Library.getColumnName(colLine + 1)
        
      Case "�i����"
        sheetSetting.Range("A" & line) = "cell_ProgressLast"
        sheetSetting.Range("B" & line) = Library.getColumnName(colLine)
        
        sheetSetting.Range("A" & line + 1) = "cell_Progress"
        sheetSetting.Range("B" & line + 1) = Library.getColumnName(colLine + 1)
                
      Case "��s�^�X�N"
        sheetSetting.Range("A" & line) = "cell_Task"
        sheetSetting.Range("B" & line) = Library.getColumnName(colLine)
        
      Case "�^�X�N�z��"
        sheetSetting.Range("A" & line) = "cell_TaskAllocation"
        sheetSetting.Range("B" & line) = Library.getColumnName(colLine)
        
      Case "�^�X�N���"
        sheetSetting.Range("A" & line) = "cell_TaskInfoP"
        sheetSetting.Range("B" & line) = Library.getColumnName(colLine)
        
        sheetSetting.Range("A" & line + 1) = "cell_TaskInfoC"
        sheetSetting.Range("B" & line + 1) = Library.getColumnName(colLine + 1)
                        
                        
      Case "��ƍH��"
        sheetSetting.Range("A" & line) = "cell_WorkLoadP"
        sheetSetting.Range("B" & line) = Library.getColumnName(colLine)
        
        sheetSetting.Range("A" & line + 1) = "cell_WorkLoadA"
        sheetSetting.Range("B" & line + 1) = Library.getColumnName(colLine + 1)
        
      Case "�x���H��"
        sheetSetting.Range("A" & line) = "cell_LateOrEarly"
        sheetSetting.Range("B" & line) = Library.getColumnName(colLine)
        
      Case "���l"
        sheetSetting.Range("A" & line) = "cell_Note"
        sheetSetting.Range("B" & line) = Library.getColumnName(colLine)
        
        '�J�����_�[�J�n�Z��
        sheetSetting.Range("A" & line + 1) = "calendarStartCol"
        sheetSetting.Range("B" & line + 1) = Library.getColumnName(colLine + 1)
        
      Case Else
    End Select

Label_nextFor:
  Next

  Call init.setting(True)
End Function


'**************************************************************************************************
' * �^�X�N���X�g�m�F
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function �^�X�N���X�g�m�F()

  Dim line As Long, endLine As Long
  Dim tmpLine As Long, startTaskLine As Long, endTaskLine As Long
  Dim targetLevel As Long, progress As Long, progressCnt As Long, lateOrEarly As Double, lateOrEarlyCnt As Long
  Dim workStartDay As Date, workEndDay As Date, chkDay As Date
  Dim errorFlg As Boolean, chlkFlg As Boolean
  Dim ErrorMeg As String
  Dim workLoadP As Long
  
'  On Error GoTo catchError
  Call init.setting
  Call Library.showDebugForm("�^�X�N���X�g�m�F", "�J�n")
  sheetMain.Select
  
  '�����I�ɍČv�Z������
  sheetMain.Calculate
  
  ErrorMeg = ""

  '���̓`�F�b�N====================================================================================
  errorFlg = False
  Call ctl_ProgressBar.showCount("�^�X�N���X�g�m�F", 0, 10, "���̓`�F�b�N")
      
  '�^�X�N���̐ݒ�
  endLine = Cells(Rows.count, 1).End(xlUp).row
  
  Call �s�����R�s�[(6, endLine)
    
  '�J�����_�[�̊J�n���ƃ^�X�N�̊J�n�����m�F
  tmpSheet.Cells.Delete Shift:=xlUp
  endLine = Cells(Rows.count, Library.getColumnNo(setVal("cell_PlanStart"))).End(xlUp).row
  If endLine > 6 Then
    Range(setVal("cell_PlanStart") & 6 & ":" & setVal("cell_PlanStart") & endLine).Copy
    tmpSheet.Range("A1").PasteSpecial
    
    tmpSheet.Sort.SortFields.Clear
    tmpSheet.Sort.SortFields.Add Key:=Range("A1"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    Application.CutCopyMode = False
    
    With tmpSheet.Sort
        .SetRange Columns("A:A")
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    If setVal("startDay") > Application.WorksheetFunction.Min(tmpSheet.Columns("A:A")) Then
      Call Library.showDebugForm("�^�X�N���X�g�m�F", "�J�����_�[�̊��Ԃ����^�X�N�̊J�n�����ߋ��ɐݒ肳��Ă��܂�")
      ErrorMeg = ErrorMeg & "�J�����_�[�̊��Ԃ����^�X�N�̊J�n�����ߋ��ɐݒ肳��Ă��܂�" & vbCrLf
      errorFlg = True
    End If
  End If
  

  '�J�����_�[�̏I�����ƃ^�X�N�̏I�������m�F
  tmpSheet.Cells.Delete Shift:=xlUp
  endLine = Cells(Rows.count, Library.getColumnNo(setVal("cell_PlanEnd"))).End(xlUp).row
  If endLine > 6 Then
    Range(setVal("cell_PlanEnd") & 6 & ":" & setVal("cell_PlanEnd") & endLine).Copy
    tmpSheet.Range("A1").PasteSpecial
    
    tmpSheet.Sort.SortFields.Clear
    tmpSheet.Sort.SortFields.Add Key:=Range("A1"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    Application.CutCopyMode = False
    
    With tmpSheet.Sort
        .SetRange Columns("A:A")
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    If setVal("endDay") < Application.WorksheetFunction.Max(tmpSheet.Range("A1:A" & Rows.count)) Then
      Call Library.showDebugForm("�^�X�N���X�g�m�F", "�J�����_�[�̊��Ԃ����^�X�N�̏I�����������ɐݒ肳��Ă��܂�")
      ErrorMeg = ErrorMeg & "�J�����_�[�̊��Ԃ����^�X�N�̏I�����������ɐݒ肳��Ă��܂�" & vbCrLf
      ErrorMeg = ErrorMeg & "�@�J�����_�[�̏I����:" & setVal("endDay")
      ErrorMeg = ErrorMeg & "�@�^�X�N�̏I����:" & Format(Application.WorksheetFunction.Max(tmpSheet.Range("A1:A" & Rows.count)), "yyyy/mm/dd")
      errorFlg = True
    End If
  End If
  
  '�J�����_�[�̊J�n���ƃ^�X�N�̎��ъJ�n�����m�F
  tmpSheet.Cells.Delete Shift:=xlUp
  endLine = Cells(Rows.count, Library.getColumnNo(setVal("cell_AchievementStart"))).End(xlUp).row
  If endLine > 6 Then
    Range(setVal("cell_AchievementStart") & 6 & ":" & setVal("cell_AchievementStart") & endLine).Copy
    tmpSheet.Range("A1").PasteSpecial
    
    tmpSheet.Sort.SortFields.Clear
    tmpSheet.Sort.SortFields.Add Key:=Range("A1"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    Application.CutCopyMode = False
    
    With tmpSheet.Sort
        .SetRange Columns("A:A")
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    If setVal("startDay") > Application.WorksheetFunction.Min(tmpSheet.Columns("A:A")) Then
      Call Library.showDebugForm("�^�X�N���X�g�m�F", "�J�����_�[�̊��Ԃ����^�X�N�̎��ъJ�n�����ߋ��ɐݒ肳��Ă��܂�")
      ErrorMeg = ErrorMeg & "�J�����_�[�̊��Ԃ����^�X�N�̎��ъJ�n�����ߋ��ɐݒ肳��Ă��܂�" & vbCrLf
      errorFlg = True
    End If
  End If

  '�J�����_�[�̏I�����ƃ^�X�N�̎��яI�������m�F
  tmpSheet.Cells.Delete Shift:=xlUp
  endLine = Cells(Rows.count, Library.getColumnNo(setVal("cell_AchievementEnd"))).End(xlUp).row
  If endLine > 6 Then
    Range(setVal("cell_AchievementEnd") & 6 & ":" & setVal("cell_AchievementEnd") & endLine).Copy
    tmpSheet.Range("A1").PasteSpecial
    
    tmpSheet.Sort.SortFields.Clear
    tmpSheet.Sort.SortFields.Add Key:=Range("A1"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    Application.CutCopyMode = False
    
    With tmpSheet.Sort
        .SetRange Columns("A:A")
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    If setVal("endDay") < Application.WorksheetFunction.Max(tmpSheet.Range("A1:A" & Rows.count)) Then
      Call Library.showDebugForm("�^�X�N���X�g�m�F", "�J�����_�[�̊��Ԃ����^�X�N�̎��яI�����������ɐݒ肳��Ă��܂�")
      ErrorMeg = ErrorMeg & "�J�����_�[�̊��Ԃ����^�X�N�̎��яI�����������ɐݒ肳��Ă��܂�" & vbCrLf
      errorFlg = True
    End If
  End If
  
  On Error Resume Next

  If Not (setVal("startDay") <= setVal("baseDay") And setVal("baseDay") <= setVal("endDay")) And setVal("workMode") <> "CD��" And setVal("setLightning") <> False Then
    Call Library.showDebugForm("������J�����_�[�̊��ԊO�ɐݒ肳��Ă��܂�")
    ErrorMeg = ErrorMeg & "������J�����_�[�̊��ԊO�ɐݒ肳��Ă��܂�" & vbCrLf
'    errorFlg = True
    Call WBS_Option.�G���[���\��(ErrorMeg)
    
  End If
  
  '�\����̃`�F�b�N===============================================================================-
  endLine = Cells(Rows.count, 1).End(xlUp).row
  For line = 6 To endLine
    If Range(setVal("cell_PlanStart") & line) > Range(setVal("cell_PlanEnd") & line) Then
      Range(setVal("cell_PlanStart") & line).Style = "Error"
      ErrorMeg = ErrorMeg & line & "�s:�\���(�J�n��)���\���(�I����)��薢���ł��B" & vbCrLf
      errorFlg = True
    End If
  
    '��ƍH��(�\��)�̎Z�o========================================================================
    workLoadP = WorksheetFunction.NetworkDays_Intl(Range(setVal("cell_PlanStart") & line), Range(setVal("cell_PlanEnd") & line), "0000011", Range("�x�����X�g"))
    If Range(setVal("cell_PlanStart") & line) <> "" And Range(setVal("cell_PlanEnd") & line) <> "" Then
      If Range(setVal("cell_WorkLoadP") & line) > workLoadP And Not (Range(setVal("cell_WorkLoadP") & line).Formula Like "=*") Then
        Range(setVal("cell_WorkLoadP") & line).Style = "Error"
        ErrorMeg = ErrorMeg & line & "�s:��ƍH��(�\��)�����ۂ̊��Ԃ�葽���ł��B�@�@���͒l:" & Range(setVal("cell_WorkLoadP") & line) & " �v�Z�l:" & workLoadP & vbCrLf
        errorFlg = True
      End If
    End If
    '�i���m�F========================================================================
    If Range(setVal("cell_Progress") & line) > 0 Then
      If Range(setVal("cell_AchievementStart") & line) = "" Then
        Range(setVal("cell_AchievementStart") & line).Style = "Error"
        ErrorMeg = ErrorMeg & line & "�s:���ѓ�(�J�n)�����͂���Ă��܂���" & vbCrLf
        errorFlg = True
      End If
      If Range(setVal("cell_AchievementEnd") & line) = "" And Range(setVal("cell_Progress") & line) = 100 Then
        Range(setVal("cell_AchievementEnd") & line).Style = "Error"
        ErrorMeg = ErrorMeg & line & "�s:���ѓ�(�I��)�����͂���Ă��܂���" & vbCrLf
        errorFlg = True
      End If
    End If
  
  
  
  Next
  
  '�G���[���̕\��===============================================================================-
  tmpSheet.Cells.Delete Shift:=xlUp
  If errorFlg = True Then
    Call WBS_Option.�G���[���\��(ErrorMeg)
    GoTo catchError
  End If
  
  
  endLine = Cells(Rows.count, 1).End(xlUp).row
  
  '�x���H���A�^�X�N���̃N���A
  Range(setVal("cell_LateOrEarly") & "4:" & setVal("cell_LateOrEarly") & endLine).ClearContents
  Range(setVal("cell_TaskInfoP") & 6 & ":" & setVal("cell_TaskInfoC") & endLine).ClearContents
  
  '�}�N���Őݒ肵�����сA�H�����N���A
  For line = 6 To endLine
    If Range(setVal("cell_WorkLoadP") & line).Formula Like "=*" Then
      Range(setVal("cell_WorkLoadP") & line).ClearContents
    End If
    
    If Range(setVal("cell_WorkLoadA") & line).Formula Like "=*" Then
      Range(setVal("cell_WorkLoadA") & line).ClearContents
    End If
    
    If Range(setVal("cell_Assign") & line) = "�H��" Then
      Range(setVal("cell_Assign") & line) = ""
    End If
  Next
  
  
  '�e�^�X�N�Ȃ�A�S����(�\��)�Ɂu�H���v�����蓖��
  Call Library.showDebugForm("�^�X�N���X�g�m�F", "�e/�q�^�X�N����")
  For line = 6 To endLine
    Call ctl_ProgressBar.showCount("�e/�q�^�X�N����", line, endLine, "")
    
    '�^�X�N���x����1�Ȃ烊�Z�b�g
    If Range(setVal("cell_LevelInfo") & line) = 1 Then
      parentTaskLine = ""
    ElseIf Range(setVal("cell_LevelInfo") & line) = "" Then
      GoTo Label_nextFor
    End If
    If Range(setVal("cell_LevelInfo") & line) < Range("B" & line + 1) And Range("B" & line + 1) <> "" Then
      endTaskLine = line + 1
      Do While Range(setVal("cell_LevelInfo") & line).Value <= Range("B" & endTaskLine).Value And Range("B" & endTaskLine) <> ""
        Call ctl_ProgressBar.showCount("�e/�q�^�X�N����", endTaskLine, endLine, "")
      
        If Range(setVal("cell_LevelInfo") & line).Value >= Range("B" & endTaskLine).Value Then
          endTaskLine = endTaskLine - 1
          Exit Do
        End If
        endTaskLine = endTaskLine + 1
      Loop
      If Range(setVal("cell_LevelInfo") & line).Value >= Range("B" & endTaskLine).Value Then
        endTaskLine = endTaskLine - 1
      End If
      Range(setVal("cell_Assign") & line) = "�H��"
      
      '�^�X�N���x���ɂ��F����
      Select Case Range(setVal("cell_LevelInfo") & line)
        Case "1"
          If setVal("lineColor_TaskLevel1") = 16777215 Then
          Else
            Range("A" & line & ":" & setVal("cell_Note") & line).Interior.Color = setVal("lineColor_TaskLevel1")
          End If
        Case "2"
          If setVal("lineColor_TaskLevel2") = 16777215 Then
          Else
            Range("A" & line & ":" & setVal("cell_Note") & line).Interior.Color = setVal("lineColor_TaskLevel2")
          End If
        Case "3"
          If setVal("lineColor_TaskLevel2") = 16777215 Then
          Else
            Range("A" & line & ":" & setVal("cell_Note") & line).Interior.Color = setVal("lineColor_TaskLevel3")
          End If
      End Select
      
      
      '��ƍH��(����)�̎Z�o=========================================================================-
      If Range(setVal("cell_PlanStart") & line) <> "" And Range(setVal("cell_PlanEnd") & line) <> "" Then
        If Range(setVal("cell_WorkLoadA") & line).Formula Like "=*" Or Range(setVal("cell_WorkLoadA") & line) = "" Then
          If Range(setVal("cell_PlanStart") & line) <= setVal("baseDay") Then
            Range(setVal("cell_WorkLoadA") & line) = "=" & WorksheetFunction.NetworkDays_Intl(Range(setVal("cell_PlanStart") & line), setVal("baseDay"), "0000011", Range("�x�����X�g"))
          ElseIf Range(setVal("cell_AchievementStart") & line) <= setVal("baseDay") Then
            Range(setVal("cell_WorkLoadA") & line) = "=" & WorksheetFunction.NetworkDays_Intl(Date, Range(setVal("cell_PlanStart") & line), "0000011", Range("�x�����X�g"))
          End If
        End If
      End If
      
      '�q�^�X�N�͈̔͂�ۑ�
      Range(setVal("cell_TaskInfoC") & line) = line + 1 & ":" & endTaskLine
      
      '�e�^�X�N���
      Range(setVal("cell_TaskInfoP") & line) = parentTaskLine
    
      parentTaskLine = line
    Else
      '�e�^�X�N���
      Range(setVal("cell_TaskInfoP") & line) = parentTaskLine
    End If
Label_nextFor:
  Next
  
  '�e�^�X�N���
  Call Library.showDebugForm("�^�X�N���X�g�m�F", "�e�^�X�N���̊m�F")
  For line = endLine To 6 Step -1
    If Range(setVal("cell_Assign") & line) = "�H��" Then
      Range(setVal("cell_TaskInfoP") & line).Select
      For lineP = line - 1 To 6 Step -1
        If Range(setVal("cell_Assign") & lineP) = "�H��" And (Range(setVal("cell_LevelInfo") & line) > Range(setVal("cell_LevelInfo") & lineP)) Then
          Range(setVal("cell_TaskInfoP") & lineP).Select
          Range(setVal("cell_TaskInfoP") & line) = Range(setVal("cell_LevelInfo") & lineP).row
          Exit For
        End If
      Next
    End If
  Next
  
  
  '�q�^�X�N�̃f�[�^�m�F
  Call Library.showDebugForm("�^�X�N���X�g�m�F", "�q�^�X�N�̃f�[�^�m�F")
  For line = 6 To endLine
    Call ctl_ProgressBar.showCount("�q�^�X�N�̃f�[�^�m�F", line, endLine, "")

'    Call Library.showDebugForm("�^�X�N���X�g�m�F", "�@" & Range(setVal("cell_Info") & line))
    
    'Level���Ȃ���΃��[�v�𔲂���
    If Range(setVal("cell_LevelInfo") & line) = "" Then Exit For
    
    If Range(setVal("cell_Assign") & line) <> "�H��" And Range(setVal("cell_Info") & line) <> setVal("TaskInfoStr_Multi") Then
      '���ѓ�(�J�n�ƏI��)�����͂���Ă���΁A�i����100�ɂ���
      If Range(setVal("cell_AchievementStart") & line) <> "" And Range(setVal("cell_AchievementEnd") & line) <> "" Then
        Range(setVal("cell_Progress") & line) = 100
      End If
      
      '--------------------------------------------------------------------------------------------
      '��ƍH��(�\��)�̎Z�o
      If Range(setVal("cell_PlanStart") & line) <> "" And Range(setVal("cell_PlanEnd") & line) <> "" Then
        If Range(setVal("cell_WorkLoadP") & line).Formula Like "=*" Or Range(setVal("cell_WorkLoadP") & line) = "" Then
          Range(setVal("cell_WorkLoadP") & line) = "=" & WorksheetFunction.NetworkDays_Intl(Range(setVal("cell_PlanStart") & line), Range(setVal("cell_PlanEnd") & line), "0000011", Range("�x�����X�g"))
        End If
      End If
      
      '--------------------------------------------------------------------------------------------
      '��ƍH��(����)�̎Z�o
      If Range(setVal("cell_PlanStart") & line) <> "" And Range(setVal("cell_PlanEnd") & line) <> "" Then
        If Range(setVal("cell_WorkLoadA") & line).Formula Like "=*" Or Range(setVal("cell_WorkLoadA") & line) = "" Then
          If Range(setVal("cell_PlanStart") & line) <= setVal("baseDay") Then
            Range(setVal("cell_WorkLoadA") & line) = "=" & WorksheetFunction.NetworkDays_Intl(Range(setVal("cell_PlanStart") & line), setVal("baseDay"), "0000011", Range("�x�����X�g"))
          End If
        End If
      End If
      
      '�i�����̐ݒ�
      '��Ɨ\������o�߂��Ă��邪�A�����͂̏ꍇ
      If Range(setVal("cell_Progress") & line) = "" And Range(setVal("cell_PlanStart") & line) < setVal("baseDay") Then
'        Range(setVal("cell_Progress") & line) = "=" & 0
      End If
      
      '--------------------------------------------------------------------------------------------
      '�x���H���̌v�Z
'      Range(setVal("cell_Progress") & line).Select
      
      '�x���H��=(��ƍH��_����-(��ƍH��_�\��/�i����))*-1
      If Range(setVal("cell_Progress") & line) = 100 And Range(setVal("cell_PlanEnd") & line) < setVal("baseDay") Then
        Range(setVal("cell_LateOrEarly") & line) = 0
        
      ElseIf Range(setVal("cell_Progress") & line) <> "" Then
        Range(setVal("cell_LateOrEarly") & line) = (Range(setVal("cell_WorkLoadA") & line) - (Range(setVal("cell_WorkLoadP") & line) * Range(setVal("cell_Progress") & line) / 100)) * -1
      End If
    End If
  Next
  
  '�e�^�X�N�̃f�[�^�m�F============================================================================
  Call Library.showDebugForm("�^�X�N���X�g�m�F", "�e�^�X�N�̃f�[�^�m�F")
  For line = 6 To endLine
    Call ctl_ProgressBar.showCount("�e�^�X�N�̃f�[�^�m�F", line, endLine, "")
    If Range(setVal("cell_TaskInfoC") & line) <> "" Then
      taskAreas = Split(Range(setVal("cell_TaskInfoC") & line), ":")
      
      '�\���(�J�n)�ݒ�==================================================================================
      workStartDay = Application.WorksheetFunction.Max(Range(setVal("cell_PlanStart") & taskAreas(0) & ":" & setVal("cell_PlanStart") & taskAreas(1)))
      For tmpLine = taskAreas(0) To taskAreas(1)
        If workStartDay > Range(setVal("cell_PlanStart") & tmpLine) And Range(setVal("cell_PlanStart") & tmpLine) <> "" Then
          workStartDay = Range(setVal("cell_PlanStart") & tmpLine)
        End If
      Next
      If workStartDay <> 0 Then
        'Range(setVal("cell_PlanStart") & line) = workStartDay
        Range(setVal("cell_PlanStart") & line) = "=Min(" & setVal("cell_PlanStart") & taskAreas(0) & ":" & setVal("cell_PlanStart") & taskAreas(1) & ")"
      End If
      
      '�\���(�I��)�ݒ�==================================================================================
      workEndDay = Application.WorksheetFunction.Min(Range(setVal("cell_PlanEnd") & taskAreas(0) & ":" & setVal("cell_PlanEnd") & taskAreas(1)))
      For tmpLine = taskAreas(0) To taskAreas(1)
        If workEndDay < Range(setVal("cell_PlanEnd") & tmpLine) And Range(setVal("cell_PlanEnd") & tmpLine) <> "" Then
          workEndDay = Range(setVal("cell_PlanEnd") & tmpLine)
        End If
      Next
      If workEndDay <> 0 Then
'        Range(setVal("cell_PlanEnd") & line) = workEndDay
        Range(setVal("cell_PlanEnd") & line) = "=Max(" & setVal("cell_PlanEnd") & taskAreas(0) & ":" & setVal("cell_PlanEnd") & taskAreas(1) & ")"
      End If
      
      
      '��ƍH��(�\��)�̎Z�o========================================================================
      If Range(setVal("cell_PlanStart") & line) <> "" And Range(setVal("cell_PlanEnd") & line) <> "" Then
        Range(setVal("cell_WorkLoadP") & line) = "=" & WorksheetFunction.NetworkDays_Intl(Range(setVal("cell_PlanStart") & line), Range(setVal("cell_PlanEnd") & line), "0000011", Range("�x�����X�g"))
      End If
      
      
      '���ѓ��̐ݒ�===============================================================================-
      workStartDay = Application.WorksheetFunction.Max(Range(setVal("cell_AchievementStart") & taskAreas(0) & ":" & setVal("cell_AchievementStart") & taskAreas(1)))
      
      For tmpLine = taskAreas(0) To taskAreas(1)
        If workStartDay > Range(setVal("cell_AchievementStart") & tmpLine) And Range(setVal("cell_AchievementStart") & tmpLine) <> "" Then
          workStartDay = Range(setVal("cell_AchievementStart") & tmpLine)
        End If
      Next
      If workStartDay <> 0 Then
        Range(setVal("cell_AchievementStart") & line) = workStartDay
      End If
      
      If WorksheetFunction.CountBlank(Range(setVal("cell_AchievementEnd") & taskAreas(0) & ":" & setVal("cell_AchievementEnd") & taskAreas(1))) = 0 Then
        Range(setVal("cell_AchievementEnd") & line) = Application.WorksheetFunction.Max(Range(setVal("cell_AchievementEnd") & taskAreas(0) & ":" & setVal("cell_AchievementEnd") & taskAreas(1)))
      End If
      
      '�i���̌v�Z==================================================================================
      progress = 0
      progressCnt = 0
      For tmpLine = taskAreas(0) To taskAreas(1)
        If Range(setVal("cell_Assign") & tmpLine) <> "�H��" Then
          progress = progress + Range(setVal("cell_Progress") & tmpLine)
          progressCnt = progressCnt + 1
        End If
      Next
      If progressCnt = 0 Or progress = 0 Then
        Range(setVal("cell_Progress") & line) = ""
      Else
        Range(setVal("cell_Progress") & line) = progress / progressCnt
      End If
  
      '�x���H���̌v�Z===============================================================================-
      lateOrEarly = 0
      lateOrEarlyCnt = 0
      For tmpLine = taskAreas(0) To taskAreas(1)
        If Range(setVal("cell_Assign") & tmpLine) <> "�H��" Then
          lateOrEarly = lateOrEarly + Range(setVal("cell_LateOrEarly") & tmpLine)
          lateOrEarlyCnt = lateOrEarlyCnt + 1
        End If
      Next
       'Range(setVal("cell_LateOrEarly") & line).Select
      If lateOrEarlyCnt = 0 Then
        Range(setVal("cell_LateOrEarly") & line) = ""
      Else
        Columns(setVal("cell_LateOrEarly") & ":" & setVal("cell_LateOrEarly")).NumberFormatLocal = "0.00_ ;[��]-0.00 "
        Range(setVal("cell_LateOrEarly") & line) = lateOrEarly
      End If
    End If
  Next
  Call Library.showDebugForm("�^�X�N���X�g�m�F", "�S�̂̐i���v�Z")
  
  '�S�̂̐i���̌v�Z===============================================================================-
  progressCnt = 0
  progress = 0
  lateOrEarly = 0
  For line = 6 To endLine
    Call ctl_ProgressBar.showCount("�S�^�X�N�̃f�[�^�W�v", line, endLine, "")
    
    If Range(setVal("cell_Assign") & line).Text <> "�H��" Then
      Range(setVal("cell_Assign") & line).Select
      progress = progress + Range(setVal("cell_Progress") & line)
      progressCnt = progressCnt + 1
      lateOrEarly = lateOrEarly + Range(setVal("cell_LateOrEarly") & line)
    End If
    
    '�i����100%�Ȃ��\��====================================
    If setVal("setDispProgress100") = True And Range(setVal("cell_Progress") & line) = 100 Then
      Rows(line & ":" & line).EntireRow.Hidden = True
      
    End If
    
  Next
  If progressCnt > 1 Then
    Range(setVal("cell_Progress") & 5) = progress / progressCnt
    Range(setVal("cell_LateOrEarly") & 5) = lateOrEarly
  ElseIf progressCnt = 1 Then
    Range(setVal("cell_Progress") & 5) = progress
    Range(setVal("cell_LateOrEarly") & 5) = lateOrEarly
  End If
  
  Call Library.showDebugForm("�^�X�N���X�g�m�F", "�I��")
  
  '�����I�ɍČv�Z������
  sheetMain.Calculate
  
  
  Exit Function
'�G���[������--------------------------------------------------------------------------------------
catchError:
  Call Library.endScript

End Function

