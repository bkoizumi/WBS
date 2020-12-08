Attribute VB_Name = "WBS_Option"

'**************************************************************************************************
' * �E�N���b�N���j���[
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function �E�N���b�N���j���[(Target As Range, Cancel As Boolean)
  Dim menu01 As CommandBarControl
  
  Call init.setting
  
  '�W����ԂɃ��Z�b�g
  Application.CommandBars("Cell").Reset

  With CommandBars("Cell").Controls.Add(Before:=1, Type:=msoControlPopup)
    .Caption = "WBS"
    With .Controls.Add(Temporary:=True)
      .Caption = "�^�X�N�ɃX�N���[��"
      .OnAction = "menu.M_�^�X�N�ɃX�N���[��"
    End With
    With .Controls.Add(Temporary:=True)
      .Caption = "�^�C�����C���ɒǉ�"
      .OnAction = "menu.M_�^�C�����C���ɒǉ�"
    End With
    With .Controls.Add(Temporary:=True)
      .BeginGroup = True
      .Caption = "�^�X�N�̃��x���グ"
      .FaceId = 3161
      .OnAction = "menu.M_�C���f���g��"
    End With
    With .Controls.Add(Temporary:=True)
      .Caption = "�^�X�N�̃��x������"
      .FaceId = 3162
      .OnAction = "menu.M_�C���f���g��"
    End With
    With .Controls.Add(Temporary:=True)
      .BeginGroup = True
      .Caption = "�^�X�N�̑}��"
      .FaceId = 296
      .OnAction = "menu.M_�^�X�N�̑}��"
    End With
    With .Controls.Add(Temporary:=True)
      .Caption = "�^�X�N�̍폜"
      .FaceId = 293
      .OnAction = "menu.M_�^�X�N�̍폜"
    End With
  End With
  
  
'  If setVal("debugMode") <> "develop" Then
'    '�W����ԂɃ��Z�b�g
'    Application.CommandBars("Cell").Reset
'
'    If setVal("debugMode") <> "develop" Then
'      '�E�N���b�N���j���[���N���A
'      For Each menu01 In Application.CommandBars("Cell").Controls
'        'Debug.Print menu01.Caption
'        Select Case True
'          Case menu01.Caption Like "�؂���*"
'          Case menu01.Caption Like "�R�s�[*"
'          Case menu01.Caption Like "�����ƒl�̃N���A*"
'          Case menu01.Caption Like "�\��t��*"
''          Case menu01.Caption Like "�Z���̏����ݒ�*"
''          Case menu01.Caption Like "�}��*"
''          Case menu01.Caption Like "�폜*"
''          Case menu01.Caption Like "�R�����g��*"
'          Case Else
'            menu01.Visible = False
'        End Select
'      Next
'    End If
'  End If
'
'
'
'
'  With Application.CommandBars("Cell").Controls.Add(Temporary:=True)
'      .Caption = "�^�X�N�ɃX�N���[��"
'      .OnAction = "menu.M_�^�X�N�ɃX�N���[��"
'  End With
'
'  With Application.CommandBars("Cell").Controls.Add(Temporary:=True)
'      .Caption = "�^�C�����C���ɒǉ�"
'      .OnAction = "menu.M_�^�C�����C���ɒǉ�"
'  End With
'
'  With Application.CommandBars("Cell").Controls.Add(Temporary:=True)
'      .BeginGroup = True
'      .Caption = "�^�X�N�̃��x���グ"
'      .FaceId = 3161
'      .OnAction = "menu.M_�C���f���g��"
'  End With
'
'  With Application.CommandBars("Cell").Controls.Add(Temporary:=True)
'      .Caption = "�^�X�N�̃��x������"
'      .FaceId = 3162
'      .OnAction = "menu.M_�C���f���g��"
'  End With
'
'  With Application.CommandBars("Cell").Controls.Add(Temporary:=True)
'      .BeginGroup = True
'      .Caption = "�^�X�N�̑}��"
'      .FaceId = 296
'      .OnAction = "menu.M_�^�X�N�̑}��"
'  End With
'
'  With Application.CommandBars("Cell").Controls.Add(Temporary:=True)
'      .Caption = "�^�X�N�̍폜"
'      .FaceId = 293
'      .OnAction = "menu.M_�^�X�N�̍폜"
'  End With


  Application.CommandBars("Cell").ShowPopup
  Application.CommandBars("Cell").Reset
  Cancel = True
End Function



' *************************************************************************************************
' * �J�����_�[�֘A�֐�
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
' *************************************************************************************************
Function �I���V�[�g�m�F()

  If ActiveSheet.Name = sheetMainName Or ActiveSheet.Name = sheetTeamsPlannerName Then
  Else
    Call Library.showNotice(454, , True)
  End If


End Function

'**************************************************************************************************
' * saveAndRefresh
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function saveAndRefresh()
  
  Application.EnableEvents = True
  ActiveWorkbook.Save
  ActiveWorkbook.RefreshAll

  Call Library.endScript
End Function


'**************************************************************************************************
' * ���t�Z������
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function ���t�Z������(chkDay As Date, Optional chlkFlg As Boolean)
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  Dim chkCell As Range
  
  
'  On Error GoTo catchError
  
  endColLine = Cells(4, Columns.count).End(xlToLeft).Column
  ���t�Z������ = Library.getColumnName(Range(Cells(4, Library.getColumnNo(setVal("calendarStartCol"))), Cells(4, endColLine)).Find(chkDay).Column)



'  Set chkCell = Range( _
'                       Cells(4, Library.getColumnNo(setVal("calendarStartCol"))), Cells(4, endColLine) _
'                                            ).Find(chkDay, SearchOrder:=xlByColumns)
'
'���t�Z������ = chkCell.Column
  
  Exit Function
'�G���[������--------------------------------------------------------------------------------------
catchError:
'  Call Library.showNotice(Err.Number, Err.Description)
  ���t�Z������ = setVal("calendarStartCol")

End Function


'**************************************************************************************************
' * �C�i�Y�}���p���t�v�Z
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function �C�i�Y�}���p���t�v�Z(baseDay As Date, calDay As Double) As Date
  Dim cntDay As Integer
  Dim resultDay As Date
  Dim loopFlg As Boolean
  Dim chk As Variant
  
  loopFlg = True
  resultDay = baseDay
  cntDay = 0
  If Application.WorksheetFunction.RoundUp(calDay, 0) <> 0 Then
    Do While loopFlg = True
      Select Case Sgn(calDay)
        Case 1
            resultDay = resultDay + 1
        Case -1
            resultDay = resultDay - 1
      End Select
      
      On Error Resume Next
      chk = ""
      chk = WorksheetFunction.VLookup(CLng(resultDay), Range("�x�����X�g"), 2, False)
      On Error GoTo 0
      
      If Weekday(resultDay) = 1 Or Weekday(resultDay) = 7 Then
        chk = "�y��"
      ElseIf IsEmpty(chk) Or chk = "" Then
        Select Case Sgn(calDay)
          Case 1
              cntDay = cntDay + 1
          Case -1
              cntDay = cntDay - 1
        End Select
      End If
      If cntDay = Application.WorksheetFunction.RoundUp(calDay, 0) Then
        loopFlg = False
      End If
    Loop
  Else
  
  End If
 �C�i�Y�}���p���t�v�Z = resultDay
End Function


'**************************************************************************************************
' * �I���s�̐F�t�ؑ�
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function setLineColor()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  Dim targetArea As String
  Dim setFlg As String
  
  Call init.setting
    
  endLine = Cells(Rows.count, 1).End(xlUp).row
  endColLine = Cells(4, Columns.count).End(xlToLeft).Column
  
  setFlg = setVal("lineColorFlg")
  
  If setFlg = "True" Then
    targetArea = "A4:" & Library.getColumnName(endColLine) & endLine
    Call Library.unsetLineColor(targetArea)
    
    setVal("lineColorFlg") = False
  Else
    '�^�X�N�G���A
    If ActiveSheet.Name = sheetMainName Then
      targetArea = "A6:" & setVal("calendarStartCol") & endLine
    ElseIf ActiveSheet.Name = sheetTeamsPlannerName Then
      targetArea = "F6:" & setVal("calendarStartCol") & endLine
    End If
    
    Call Library.setLineColor(targetArea, False, setVal("lineColor"))
    
    '�J�����_�[�G���A
    targetArea = setVal("calendarStartCol") & "4:" & Library.getColumnName(endColLine) & endLine
    Call Library.setLineColor(targetArea, True, setVal("lineColor"))
  
    setVal("lineColorFlg") = True
  End If
  
  sheetSetting.Range("lineColorFlg") = setVal("lineColorFlg")
End Function

'**************************************************************************************************
' * �V�[�g���̑S�f�[�^�폜
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function clearAll()
  Call init.setting
  
  Call �I���V�[�g�m�F
  Call Library.delSheetData(6)
  
  Columns(setVal("calendarStartCol") & ":XFD").Delete Shift:=xlToLeft
  
  '�S�̂̐i���Ȃǂ��폜
  Range("I5:" & setVal("cell_Note") & 5).ClearContents
  
  
  Range(setVal("calendarStartCol") & "1:" & setVal("calendarStartCol") & 5).Borders(xlEdgeLeft).LineStyle = xlDouble
'  sheetSetting.Range("O3:P" & sheetSetting.Cells(Rows.count, 15).End(xlUp).row + 1).ClearContents
  
  Application.Goto Reference:=Range("A6"), Scroll:=True
    
End Function

'**************************************************************************************************
' * �V�[�g���̑S�f�[�^�폜
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function clearCalendar()
  Call init.setting
  sheetMain.Select
  Columns(setVal("calendarStartCol") & ":XFD").Delete Shift:=xlToLeft
  
  '�S�̂̐i���Ȃǂ��폜
  Range("I5:" & setVal("cell_Note") & 5).ClearContents
  Application.Goto Reference:=Range("A6"), Scroll:=True
  
End Function




'**************************************************************************************************
' * �G���[���\��
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function �G���[���\��(ErrorMeg As String)

  With ErrorForm
    .StartUpPosition = 0
    .top = Application.top + (ActiveWindow.Width / 8)
    .Left = Application.Left + (ActiveWindow.Height / 8)
    .errMeg.Text = ErrorMeg
  End With
  
  ErrorForm.Show vbModeless

End Function


'**************************************************************************************************
' * �\����ݒ�
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function �\����ݒ�()
  Dim line As Long, endLine As Long
  Dim viewLineName As Variant
  
  On Error GoTo catchError
  
  Call init.setting(True)
  sheetMain.Select
  
  Columns(setVal("cell_PlanStart") & ":" & setVal("cell_PlanEnd")).EntireColumn.Hidden = setVal("view_Plan")
  Columns(setVal("cell_Assign") & ":" & setVal("cell_Assign")).EntireColumn.Hidden = setVal("view_Assign")
  Columns(setVal("cell_ProgressLast") & ":" & setVal("cell_Progress")).EntireColumn.Hidden = setVal("view_Progress")
  
  Columns(setVal("cell_AchievementStart") & ":" & setVal("cell_AchievementEnd")).EntireColumn.Hidden = setVal("view_Achievement")
  Columns(setVal("cell_Task") & ":" & setVal("cell_Task")).EntireColumn.Hidden = setVal("view_Task")
  Columns(setVal("cell_TaskInfoP") & ":" & setVal("cell_TaskInfoC")).EntireColumn.Hidden = setVal("view_TaskInfo")
  
  Columns(setVal("cell_WorkLoadP") & ":" & setVal("cell_WorkLoadA")).EntireColumn.Hidden = setVal("view_WorkLoad")
  
  Columns(setVal("cell_LateOrEarly") & ":" & setVal("cell_LateOrEarly")).EntireColumn.Hidden = setVal("view_LateOrEarly")
  Columns(setVal("cell_Note") & ":" & setVal("cell_Note")).EntireColumn.Hidden = setVal("view_Note")


  Columns(setVal("cell_LineInfo") & ":" & setVal("cell_LineInfo")).EntireColumn.Hidden = setVal("view_LineInfo")
  Columns(setVal("cell_TaskAllocation") & ":" & setVal("cell_TaskAllocation")).EntireColumn.Hidden = setVal("view_TaskAllocation")

Exit Function
'�G���[������--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(Err.Number, Err.Description, True)
End Function


'**************************************************************************************************
' * xxxxxxxxxx
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function �^�X�N�\��_�W��()
  Dim line As Long, endLine As Long, rowLine As Long, endColLine As Long
  
  
  On Error GoTo catchError

  
  Call init.setting
  
  endLine = sheetTeamsPlanner.Cells(Rows.count, 1).End(xlUp).row
  
  Rows("6:" & Rows.count).EntireRow.Hidden = False
  
  '�`�[���v�����i�[�ŕύX�����\������i�[
  For line = 6 To endLine
    If sheetTeamsPlanner.Range(("C") & line) Like "*" & setVal("TaskInfoStr_Change") & "*" Then
      rowLine = sheetTeamsPlanner.Range(("D") & line) + 5
      
      sheetMain.Range(setVal("cell_PlanStart") & rowLine) = sheetTeamsPlanner.Range(("G") & line)
      sheetMain.Range(setVal("cell_PlanEnd") & rowLine) = sheetTeamsPlanner.Range(("H") & line)
    End If
  Next
  
  sheetMain.Visible = True
  sheetTeamsPlanner.Visible = xlSheetVeryHidden
    
  sheetMain.Select
  sheetMain.ScrollArea = ""
  Cells.EntireColumn.Hidden = False

  '�E�C���h�E�g�̌Œ�
  Range(setVal("calendarStartCol") & 6).Select
  ActiveWindow.FreezePanes = False
  ActiveWindow.FreezePanes = True
  
  Call Chart.�K���g�`���[�g����
  Call WBS_Option.�����̒S���ҍs���\��
  Call �\����ݒ�
  
  

  Exit Function
'�G���[������--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(Err.Number, Err.Description, True)
End Function


'**************************************************************************************************
' * xxxxxxxxxx
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function viewTask()
  On Error GoTo catchError

  Call Library.startScript
  Call init.setting
  
  ActiveWindow.FreezePanes = False
  
  sheetMain.Columns(setVal("calendarStartCol") & ":" & Library.getColumnName(Cells(4, Columns.count).End(xlToLeft).Column)).EntireColumn.Hidden = True
  sheetMain.ScrollArea = "A1:P" & Rows.count
  
  '�E�C���h�E�g�̌Œ�
  Range("A6").Select
  ActiveWindow.FreezePanes = True
    
    
  Call Library.endScript(True)

  Exit Function
'�G���[������--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(Err.Number, Err.Description, True)
End Function


'**************************************************************************************************
' * �^�X�N�\��_�`�[���v�����i�[
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function �^�X�N�\��_�`�[���v�����i�[()
'  On Error GoTo catchError
  
  sheetTeamsPlanner.Select
  
  Call Calendar.makeCalendar
  Call TeamsPlanner.�f�[�^�ڍs
  
  sheetTeamsPlanner.Columns("I:S").EntireColumn.Hidden = True
  
  sheetMain.Visible = xlSheetVeryHidden
  Call Library.endScript

  Exit Function
'�G���[������--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(Err.Number, Err.Description, True)
End Function


'**************************************************************************************************
' * xxxxxxxxxx
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function viewSetting()
  On Error GoTo catchError

  Call Library.startScript
  Call Library.endScript(True)

  Exit Function
'�G���[������--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(Err.Number, Err.Description, True)
End Function


'**************************************************************************************************
' * xxxxxxxxxx
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function �i�����̃o�[�ݒ�()
  On Error GoTo catchError
  
  Range("K4").Select
  Selection.FormatConditions.AddDatabar
  Selection.FormatConditions(Selection.FormatConditions.count).ShowValue = True
  Selection.FormatConditions(Selection.FormatConditions.count).SetFirstPriority
  With Selection.FormatConditions(1)
    .MinPoint.Modify newtype:=xlConditionValueNumber, newvalue:=0
    .MaxPoint.Modify newtype:=xlConditionValueNumber, newvalue:=100
  End With
  With Selection.FormatConditions(1).BarColor
  .Color = RGB(102, 153, 255)
    .TintAndShade = 0
'  Select Case Selection.Value
'    Case 0 To 49
'      .Color = RGB(255, 0, 0)
'    Case 50 To 74
'      .Color = RGB(102, 153, 255)
'    Case 75 To 100
'      .Color = RGB(102, 153, 255)
'    Case Else
'  End Select
  End With
  Selection.FormatConditions(1).BarFillType = xlDataBarFillSolid
  Selection.FormatConditions(1).Direction = xlLTR
  Selection.FormatConditions(1).NegativeBarFormat.ColorType = xlDataBarColor
  Selection.FormatConditions(1).BarBorder.Type = xlDataBarBorderNone
  Selection.FormatConditions(1).AxisPosition = xlDataBarAxisAutomatic


  Exit Function
'�G���[������--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(Err.Number, Err.Description, True)
End Function


'**************************************************************************************************
' * xxxxxxxxxx
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function setTaskLevel()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long

  Dim taskLevelRange As Range
  
  Call init.setting
  line = 6
  
  Set taskLevelRange = Range(setVal("cell_TaskArea") & line)
  Range(setVal("cell_LevelInfo") & line).Formula = "=getIndentLevel(" & taskLevelRange.Address(ReferenceStyle:=xlA1, RowAbsolute:=False, ColumnAbsolute:=False) & ")"
  Set taskLevelRange = Nothing


  Exit Function
'�G���[������--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(Err.Number, Err.Description, True)
End Function


'**************************************************************************************************
' * �S���҂𕡐��I��
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function �S���҂̕����I��()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  
  With Frm_Assignor
    .StartUpPosition = 0
    .top = Application.top + (ActiveWindow.Width / 8)
    .Left = Application.Left + (ActiveWindow.Height / 8)
   .Show vbModeless
  End With

  Exit Function
'�G���[������--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(Err.Number, Err.Description, True)
End Function




'**************************************************************************************************
' * �����̒S���ҍs���\��
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function �����̒S���ҍs���\��()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  
  endLine = Cells(Rows.count, 1).End(xlUp).row
   
  For line = 6 To endLine
    If Range(setVal("cell_Info") & line) = "�|" Then
      Range(setVal("cell_Info") & line) = "�{"
    ElseIf Range(setVal("cell_Info") & line) = setVal("TaskInfoStr_Multi") Then
      Rows(line & ":" & line).EntireRow.Hidden = True
    End If
  Next

  Exit Function
'�G���[������--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(Err.Number, Err.Description, True)
End Function

'**************************************************************************************************
' * �^�X�N���x���̐ݒ�
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function �^�X�N���x���̐ݒ�()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long


  If ActiveSheet.Name = sheetMainName Then
'    Rows("6:" & Rows.count).EntireRow.Hidden = False
    
    endLine = Cells(Rows.count, 1).End(xlUp).row
    For line = 6 To endLine
      If Range(setVal("cell_TaskArea") & line) <> "" Then
        If Range(setVal("cell_LevelInfo") & line) = "" Then
          Range(setVal("cell_LevelInfo") & line) = Library.getIndentLevel(Range(setVal("cell_TaskArea") & line))
        End If
        
        
        taskLevel = Range(setVal("cell_LevelInfo") & line) - 1
        If taskLevel > 0 Then
          If Range(setVal("cell_TaskArea") & line).IndentLevel <> 0 Then
            Range(setVal("cell_TaskArea") & line).InsertIndent -Range(setVal("cell_TaskArea") & line).IndentLevel
          End If
          Range(setVal("cell_TaskArea") & line).InsertIndent taskLevel
        End If
      End If
    Next
  End If


End Function



'**************************************************************************************************
' * �^�X�N�ɃX�N���[��
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function �^�X�N�ɃX�N���[��()
  Dim line As Long, activeCellRowLine As Long, activeCellColLine As Long
  On Error GoTo catchError

  activeCellRowLine = ActiveCell.row
  activeCellColLine = ActiveCell.Column
  
  targetColumn = Library.getColumnNo(WBS_Option.���t�Z������(Range(setVal("cell_PlanStart") & activeCellRowLine) - 1))
  ActiveWindow.ScrollColumn = targetColumn
  
  Exit Function
'�G���[������--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(Err.Number, Err.Description, True)
End Function


'**************************************************************************************************
' * �s�ԍ��Đݒ�
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function �s�ԍ��Đݒ�()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  
  On Error GoTo catchError

  Call init.setting
  endLine = Cells(Rows.count, 1).End(xlUp).row
  
  For line = 6 To endLine
    If line = 6 Then
      Range("A" & line) = 1
    ElseIf Range(setVal("cell_Info") & line) = setVal("TaskInfoStr_Multi") Then
      Range("A" & line) = Range("A" & line - 1)
    
    ElseIf Range(setVal("cell_TaskArea") & line) = "" Then
      Range("A" & line) = ""
    Else
      Range("A" & line) = Range("A" & line - 1) + 1
    End If
  Next
  
  Exit Function
'�G���[������--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(Err.Number, Err.Description, True)
End Function
