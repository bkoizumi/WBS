Attribute VB_Name = "Check"
'**************************************************************************************************
' * ���ڃ`�F�b�N
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function ���ڗ�`�F�b�N()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  Dim itemName As String
  
  line = 30
  endLine = Cells(Rows.count, 1).End(xlUp).row
  Range("A" & line & ":B" & endLine).ClearContents
  startLine = 4
  
  Call init.setting
 
  For colLine = startLine To 15 Step 2
    If mainSheet.Cells(2, colLine) <> "" Then
      itemName = mainSheet.Cells(2, colLine)
    End If
    
    Select Case itemName
      Case "�\���"
        setSheet.Range("A" & line) = "cell_PlanStart"
        setSheet.Range("B" & line) = Library.getColumnName(colLine)
        
        setSheet.Range("A" & line + 1) = "cell_PlanEnd"
        setSheet.Range("B" & line + 1) = Library.getColumnName(colLine + 1)
      
      Case "�S����"
        setSheet.Range("A" & line) = "cell_AssignP"
        setSheet.Range("B" & line) = Library.getColumnName(colLine)
        
        setSheet.Range("A" & line + 1) = "cell_AssignA"
        setSheet.Range("B" & line + 1) = Library.getColumnName(colLine + 1)
        
      Case "���ѓ�"
        setSheet.Range("A" & line) = "cell_AchievementStart"
        setSheet.Range("B" & line) = Library.getColumnName(colLine)
        
        setSheet.Range("A" & line + 1) = "cell_AchievementEnd"
        setSheet.Range("B" & line + 1) = Library.getColumnName(colLine + 1)
        
      Case "�i����"
        setSheet.Range("A" & line) = "cell_ProgressLast"
        setSheet.Range("B" & line) = Library.getColumnName(colLine)
        
        setSheet.Range("A" & line + 1) = "cell_Progress"
        setSheet.Range("B" & line + 1) = Library.getColumnName(colLine + 1)
                
      Case "�^�X�N"
        setSheet.Range("A" & line) = "cell_TaskA"
        setSheet.Range("B" & line) = Library.getColumnName(colLine)
        
        setSheet.Range("A" & line + 1) = "cell_TaskB"
        setSheet.Range("B" & line + 1) = Library.getColumnName(colLine + 1)
                
      Case "��ƍH��"
        setSheet.Range("A" & line) = "cell_WorkLoadP"
        setSheet.Range("B" & line) = Library.getColumnName(colLine)
        
        setSheet.Range("A" & line + 1) = "cell_WorkLoadA"
        setSheet.Range("B" & line + 1) = Library.getColumnName(colLine + 1)
        
      Case Else
    End Select

    line = line + 2
  Next

'  colLine = colLine - 1
  '�x���H��----------------------------
  setSheet.Range("A" & line) = "cell_LateOrEarly"
  setSheet.Range("B" & line) = Library.getColumnName(colLine)

  '���l----------------------------
  colLine = colLine + 1
  line = line + 1
  setSheet.Range("A" & line) = "cell_Note"
  setSheet.Range("B" & line) = Library.getColumnName(colLine)

  '�J�����_�[�J�n�Z��
  colLine = colLine + 1
  line = line + 1
  setSheet.Range("A" & line) = "calendarStartCol"
  setSheet.Range("B" & line) = Library.getColumnName(colLine)
  
  Call init.setting
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
  Dim errorMeg As String
  Dim workLoadP As Long
  
'  On Error GoTo catchError
    
  ' �����I�ɍČv�Z������
  Application.CalculateFull
  
  Call init.setting
  Call Library.startScript
  mainSheet.Select
  errorMeg = ""

  '���̓`�F�b�N------------------------------------------------------------------------------------
  errorFlg = False
  Call ProgressBar.showCount("�^�X�N�m�F", 0, 10, "���̓`�F�b�N")
      
  '�^�X�N���̐ݒ�
  endLine = Cells(Rows.count, 1).End(xlUp).row
  
  '�J�����_�[�̊J�n���ƃ^�X�N�̊J�n�����m�F
  tmpSheet.Cells.Delete Shift:=xlUp
  endLine = Cells(Rows.count, Library.getColumnNo(setVal("cell_PlanStart"))).End(xlUp).row
  If endLine > 6 Then
    mainSheet.Range(setVal("cell_PlanStart") & 6 & ":" & setVal("cell_PlanStart") & endLine).Copy
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
      Call Library.showDebugForm("�^�X�N�m�F", "�J�����_�[�̊��Ԃ����^�X�N�̊J�n�����ߋ��ɐݒ肳��Ă��܂�")
      errorMeg = errorMeg & "�J�����_�[�̊��Ԃ����^�X�N�̊J�n�����ߋ��ɐݒ肳��Ă��܂�" & vbCrLf
  '    Call Library.showNotice(400)
      errorFlg = True
    End If
  End If
  

  '�J�����_�[�̏I�����ƃ^�X�N�̏I�������m�F
  tmpSheet.Cells.Delete Shift:=xlUp
  endLine = Cells(Rows.count, Library.getColumnNo(setVal("cell_PlanEnd"))).End(xlUp).row
  If endLine > 6 Then
    mainSheet.Range(setVal("cell_PlanEnd") & 6 & ":" & setVal("cell_PlanEnd") & endLine).Copy
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
      Call Library.showDebugForm("�^�X�N�m�F", "�J�����_�[�̊��Ԃ����^�X�N�̏I�����������ɐݒ肳��Ă��܂�")
      errorMeg = errorMeg & "�J�����_�[�̊��Ԃ����^�X�N�̏I�����������ɐݒ肳��Ă��܂�" & vbCrLf
      errorMeg = errorMeg & "�@�J�����_�[�̏I����:" & setVal("endDay")
      errorMeg = errorMeg & "�@�^�X�N�̏I����:" & Format(Application.WorksheetFunction.Max(tmpSheet.Range("A1:A" & Rows.count)), "yyyy/mm/dd")
      
      
  '    Call Library.showNotice(401)
      errorFlg = True
    End If
  End If
  
  '�J�����_�[�̊J�n���ƃ^�X�N�̎��ъJ�n�����m�F
  tmpSheet.Cells.Delete Shift:=xlUp
  endLine = Cells(Rows.count, Library.getColumnNo(setVal("cell_AchievementStart"))).End(xlUp).row
  If endLine > 6 Then
    mainSheet.Range(setVal("cell_AchievementStart") & 6 & ":" & setVal("cell_AchievementStart") & endLine).Copy
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
      Call Library.showDebugForm("�^�X�N�m�F", "�J�����_�[�̊��Ԃ����^�X�N�̎��ъJ�n�����ߋ��ɐݒ肳��Ă��܂�")
      errorMeg = errorMeg & "�J�����_�[�̊��Ԃ����^�X�N�̎��ъJ�n�����ߋ��ɐݒ肳��Ă��܂�" & vbCrLf
  '    Call Library.showNotice(402)
      errorFlg = True
    End If
  End If

  '�J�����_�[�̏I�����ƃ^�X�N�̎��яI�������m�F
  tmpSheet.Cells.Delete Shift:=xlUp
  endLine = Cells(Rows.count, Library.getColumnNo(setVal("cell_AchievementEnd"))).End(xlUp).row
  If endLine > 6 Then
    mainSheet.Range(setVal("cell_AchievementEnd") & 6 & ":" & setVal("cell_AchievementEnd") & endLine).Copy
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
      Call Library.showDebugForm("�^�X�N�m�F", "�J�����_�[�̊��Ԃ����^�X�N�̎��яI�����������ɐݒ肳��Ă��܂�")
      errorMeg = errorMeg & "�J�����_�[�̊��Ԃ����^�X�N�̎��яI�����������ɐݒ肳��Ă��܂�" & vbCrLf
  '    Call Library.showNotice(403)
      errorFlg = True
    End If
  End If
  
  If Not (setVal("startDay") <= setVal("baseDay") And setVal("baseDay") <= setVal("endDay")) Then
    Call Library.showDebugForm("������J�����_�[�̊��ԊO�ɐݒ肳��Ă��܂�")
    errorMeg = errorMeg & "������J�����_�[�̊��ԊO�ɐݒ肳��Ă��܂�" & vbCrLf
    errorFlg = True
  End If
  
  '��ƍH��(�\��)�̎Z�o------------------------------------------------------------------------
  endLine = Cells(Rows.count, 1).End(xlUp).row
  For line = 6 To endLine
    workLoadP = WorksheetFunction.NetworkDays_Intl(Range(setVal("cell_PlanStart") & line), Range(setVal("cell_PlanEnd") & line), "0000011", Range("�x�����X�g"))
    If Range(setVal("cell_PlanStart") & line) <> "" And Range(setVal("cell_PlanEnd") & line) <> "" Then
      If Range(setVal("cell_WorkLoadP") & line) > workLoadP And Not (Range(setVal("cell_WorkLoadP") & line).Formula Like "=*") Then
        Range(setVal("cell_WorkLoadP") & line).Style = "Error"
        errorMeg = errorMeg & line & "�s:��ƍH��(�\��)�����ۂ̊��Ԃ�葽���ł��B�@�@���͒l:" & Range(setVal("cell_WorkLoadP") & line) & " �v�Z�l:" & workLoadP & vbCrLf
        errorFlg = True
      End If
    End If
  Next
  
  '�G���[���̕\��--------------------------------------------------------------------------------
  tmpSheet.Cells.Delete Shift:=xlUp
  If errorFlg = True Then
    Call WBS_Option.�G���[���\��(errorMeg)
    GoTo catchError
  End If
  
  
  endLine = Cells(Rows.count, 1).End(xlUp).row
  
  '�x���H���̃N���A
  Range(setVal("cell_LateOrEarly") & "4:" & setVal("cell_LateOrEarly") & endLine).ClearContents
  
  '�e�^�X�N�Ȃ�A�S����(�\��)�Ɂu�H���v�����蓖��
  For line = 6 To endLine
    'Level���Ȃ���΃��[�v�𔲂���
    If mainSheet.Range("B" & line) = "" Then Exit For
    
    Call ProgressBar.showCount("�^�X�N�m�F", line, endLine, "�e�^�X�N����")
    If mainSheet.Range("B" & line) < mainSheet.Range("B" & line + 1) And mainSheet.Range("B" & line + 1) <> "" Then
      endTaskLine = line + 1
      Do While mainSheet.Range("B" & line).Value <= mainSheet.Range("B" & endTaskLine).Value And mainSheet.Range("B" & endTaskLine) <> ""
        Call ProgressBar.showCount("�^�X�N�m�F", endTaskLine, endLine, "�e�^�X�N����")
      
        mainSheet.Rows(line & ":" & endTaskLine).Select
        If Range("B" & line).Value >= Range("B" & endTaskLine).Value Then
          endTaskLine = endTaskLine - 1
          Exit Do
        End If
        endTaskLine = endTaskLine + 1
      Loop
      If mainSheet.Range("B" & line).Value >= mainSheet.Range("B" & endTaskLine).Value Then
        endTaskLine = endTaskLine - 1
      End If
      mainSheet.Rows(line & ":" & endTaskLine).Select
      Range(setVal("cell_AssignP") & line) = "�H��"
      
      '�^�X�N���x���ɂ��F����
      Select Case Range("B" & line)
        Case "1"
          If setVal("lineColor_TaskLevel1") = 16777215 Then
          Else
            mainSheet.Range("A" & line & ":" & setVal("cell_Note") & line).Interior.Color = setVal("lineColor_TaskLevel1")
          End If
        Case "2"
          If setVal("lineColor_TaskLevel2") = 16777215 Then
          Else
            mainSheet.Range("A" & line & ":" & setVal("cell_Note") & line).Interior.Color = setVal("lineColor_TaskLevel2")
          End If
        Case "3"
          If setVal("lineColor_TaskLevel2") = 16777215 Then
          Else
            mainSheet.Range("A" & line & ":" & setVal("cell_Note") & line).Interior.Color = setVal("lineColor_TaskLevel3")
          End If
      End Select
      
      
      '��ƍH��(����)�̎Z�o--------------------------------------------------------------------------
      If Range(setVal("cell_PlanStart") & line) <> "" And Range(setVal("cell_PlanEnd") & line) <> "" Then
        If Range(setVal("cell_WorkLoadA") & line).Formula Like "=*" Or Range(setVal("cell_WorkLoadA") & line) = "" Then
          If Range(setVal("cell_PlanStart") & line) <= Date Then
            Range(setVal("cell_WorkLoadA") & line) = "=" & WorksheetFunction.NetworkDays_Intl(Range(setVal("cell_PlanStart") & line), setVal("baseDay"), "0000011", Range("�x�����X�g"))
          ElseIf Range(setVal("cell_AchievementStart") & line) <= Date Then
            Range(setVal("cell_WorkLoadA") & line) = "=" & WorksheetFunction.NetworkDays_Intl(Date, Range(setVal("cell_PlanStart") & line), "0000011", Range("�x�����X�g"))
          End If
        End If
      End If
      
      '�e�^�X�N�͈̔͂��ꎞ�ۑ�------------------
      mainSheet.Range(setVal("cell_LateOrEarly") & line).NumberFormatLocal = "@"
      mainSheet.Range(setVal("cell_LateOrEarly") & line) = line + 1 & ":" & endTaskLine
    End If
  Next
  
  '�q�^�X�N�̃f�[�^�m�F
  Call Library.showDebugForm("�^�X�N�m�F", "�q�^�X�N�̃f�[�^�m�F")
  For line = 6 To endLine
    Call ProgressBar.showCount("�^�X�N�m�F", line, endLine, "�q�^�X�N�̃f�[�^�m�F")

    Call Library.showDebugForm("�^�X�N�m�F", "�@" & mainSheet.Range("C" & line))
    
    'Level���Ȃ���΃��[�v�𔲂���
    If mainSheet.Range("B" & line) = "" Then Exit For
    
    If Range(setVal("cell_AssignP") & line) <> "�H��" Then
      '���ѓ�(�J�n�ƏI��)�����͂���Ă���΁A�i����100�ɂ���---------------------------------------
      If mainSheet.Range(setVal("cell_AchievementStart") & line) <> "" And mainSheet.Range(setVal("cell_AchievementEnd") & line) <> "" Then
        mainSheet.Range(setVal("cell_Progress") & line) = 100
      End If
    
      '��ƍH��(�\��)�̎Z�o--------------------------------------------------------------------------
      If Range(setVal("cell_PlanStart") & line) <> "" And Range(setVal("cell_PlanEnd") & line) <> "" Then
        If Range(setVal("cell_WorkLoadP") & line).Formula Like "=*" Or Range(setVal("cell_WorkLoadP") & line) = "" Then
          Range(setVal("cell_WorkLoadP") & line) = "=" & WorksheetFunction.NetworkDays_Intl(Range(setVal("cell_PlanStart") & line), Range(setVal("cell_PlanEnd") & line), "0000011", Range("�x�����X�g"))
        End If
      End If
      
      '��ƍH��(����)�̎Z�o--------------------------------------------------------------------------
      If Range(setVal("cell_PlanStart") & line) <> "" And Range(setVal("cell_PlanEnd") & line) <> "" Then
        If Range(setVal("cell_WorkLoadA") & line).Formula Like "=*" Or Range(setVal("cell_WorkLoadA") & line) = "" Then
          If Range(setVal("cell_PlanStart") & line) <= Date Then
            Range(setVal("cell_WorkLoadA") & line) = "=" & WorksheetFunction.NetworkDays_Intl(Range(setVal("cell_PlanStart") & line), setVal("baseDay"), "0000011", Range("�x�����X�g"))
'          Else
'            Range(setVal("cell_WorkLoadA") & line) = "=" & WorksheetFunction.NetworkDays_Intl(Date, Range(setVal("cell_PlanStart") & line), "0000011", Range("�x�����X�g"))
          End If
        End If
      End If
      
      '�i�����̐ݒ�
      '��Ɨ\������o�߂��Ă��邪�A�����͂̏ꍇ
      If Range(setVal("cell_Progress") & line) = "" And Range(setVal("cell_PlanStart") & line) <= setVal("baseDay") Then
        Range(setVal("cell_Progress") & line) = "=" & 0
      End If
      
      '�x���H���̌v�Z--------------------------------------------------------------------------------
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
  For line = 6 To endLine
    'Level���Ȃ���΃��[�v�𔲂���
    If mainSheet.Range("B" & line) = "" Then Exit For
    
    Call ProgressBar.showCount("�^�X�N�m�F", line, endLine, "�e�^�X�N�̃f�[�^�m�F")
    If Range(setVal("cell_AssignP") & line) = "�H��" Then
      taskAreas = Split(Range(setVal("cell_LateOrEarly") & line), ":")
      
      '�\���(�J�n)�ݒ�----------------------------------------------------------------------------------
      workStartDay = Application.WorksheetFunction.Max(Range(setVal("cell_PlanStart") & taskAreas(0) & ":" & setVal("cell_PlanStart") & taskAreas(1)))
      For tmpLine = taskAreas(0) To taskAreas(1)
        If workStartDay > Range(setVal("cell_PlanStart") & tmpLine) And Range(setVal("cell_PlanStart") & tmpLine) <> "" Then
          workStartDay = Range(setVal("cell_PlanStart") & tmpLine)
        End If
      Next
      If workStartDay <> 0 Then
        Range(setVal("cell_PlanStart") & line) = workStartDay
      End If
      
      '�\���(�I��)�ݒ�----------------------------------------------------------------------------------
      workEndDay = Application.WorksheetFunction.Min(Range(setVal("cell_PlanEnd") & taskAreas(0) & ":" & setVal("cell_PlanEnd") & taskAreas(1)))
      For tmpLine = taskAreas(0) To taskAreas(1)
        If workEndDay < Range(setVal("cell_PlanEnd") & tmpLine) And Range(setVal("cell_PlanEnd") & tmpLine) <> "" Then
          workEndDay = Range(setVal("cell_PlanEnd") & tmpLine)
        End If
      Next
      If workEndDay <> 0 Then
        Range(setVal("cell_PlanEnd") & line) = workEndDay
      End If
      
      
      '��ƍH��(�\��)�̎Z�o------------------------------------------------------------------------
      If Range(setVal("cell_PlanStart") & line) <> "" And Range(setVal("cell_PlanEnd") & line) <> "" Then
        Range(setVal("cell_WorkLoadP") & line) = "=" & WorksheetFunction.NetworkDays_Intl(Range(setVal("cell_PlanStart") & line), Range(setVal("cell_PlanEnd") & line), "0000011", Range("�x�����X�g"))
      End If
      
      
      '���ѓ��̐ݒ�--------------------------------------------------------------------------------
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
      
      '�i���̌v�Z----------------------------------------------------------------------------------
      progress = 0
      progressCnt = 0
      For tmpLine = taskAreas(0) To taskAreas(1)
        If Range(setVal("cell_AssignP") & tmpLine) <> "�H��" Then
          progress = progress + Range(setVal("cell_Progress") & tmpLine)
          progressCnt = progressCnt + 1
        End If
      Next
      If progressCnt = 0 Or progress = 0 Then
        Range(setVal("cell_Progress") & line) = ""
      Else
        Range(setVal("cell_Progress") & line) = progress / progressCnt
      End If
  
      '�x���H���̌v�Z--------------------------------------------------------------------------------
      lateOrEarly = 0
      lateOrEarlyCnt = 0
      For tmpLine = taskAreas(0) To taskAreas(1)
        If Range(setVal("cell_AssignP") & tmpLine) <> "�H��" Then
          lateOrEarly = lateOrEarly + Range(setVal("cell_LateOrEarly") & tmpLine)
          lateOrEarlyCnt = lateOrEarlyCnt + 1
        End If
      Next
       Range(setVal("cell_LateOrEarly") & line).Select
      If lateOrEarlyCnt = 0 Then
        Range(setVal("cell_LateOrEarly") & line) = ""
      Else
        Columns(setVal("cell_LateOrEarly") & ":" & setVal("cell_LateOrEarly")).NumberFormatLocal = "0.00_ ;[��]-0.00 "
        Range(setVal("cell_LateOrEarly") & line) = lateOrEarly
      End If
    End If
  Next
  
  '�S�̂̐i���̌v�Z--------------------------------------------------------------------------------
  progressCnt = 0
  progress = 0
  lateOrEarly = 0
  For line = 6 To endLine
    Call ProgressBar.showCount("�^�X�N�m�F", line, endLine, "�S�^�X�N�̃f�[�^�W�v")
    
    If Range(setVal("cell_AssignP") & line).Text <> "�H��" Then
      mainSheet.Range(setVal("cell_AssignP") & line).Select
      progress = progress + mainSheet.Range(setVal("cell_Progress") & line)
      progressCnt = progressCnt + 1
      lateOrEarly = lateOrEarly + mainSheet.Range(setVal("cell_LateOrEarly") & line)
    End If
  Next
  If progressCnt > 1 Then
    Range(setVal("cell_Progress") & 5) = progress / progressCnt
    Range(setVal("cell_LateOrEarly") & 5) = lateOrEarly
  ElseIf progressCnt = 1 Then
    Range(setVal("cell_Progress") & 5) = progress
    Range(setVal("cell_LateOrEarly") & 5) = lateOrEarly
  End If
  

  
  
  
  
  Exit Function
'�G���[������=====================================================================================
catchError:
  Call Library.endScript

End Function

