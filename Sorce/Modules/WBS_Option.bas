Attribute VB_Name = "WBS_Option"

'**************************************************************************************************
' * �E�N���b�N���j���[
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function �E�N���b�N���j���[(Target As Range, Cancel As Boolean)
  Dim menu01 As CommandBarControl
  
  Call init.setting
  
  If setVal("debugMode") <> "develop" Then
    '�W����ԂɃ��Z�b�g
    Application.CommandBars("Cell").Reset

    If setVal("debugMode") <> "develop" Then
      '�E�N���b�N���j���[���N���A
      For Each menu01 In Application.CommandBars("Cell").Controls
        'Debug.Print menu01.Caption
        Select Case True
          Case menu01.Caption Like "�؂���*"
          Case menu01.Caption Like "�R�s�[*"
          Case menu01.Caption Like "�����ƒl�̃N���A*"
          Case menu01.Caption Like "�\��t��*"
'          Case menu01.Caption Like "�Z���̏����ݒ�*"
'          Case menu01.Caption Like "�}��*"
'          Case menu01.Caption Like "�폜*"
'          Case menu01.Caption Like "�R�����g��*"
          Case Else
            menu01.Visible = False
        End Select
      Next
    End If
  End If
  


  
  With Application.CommandBars("Cell").Controls.Add(Temporary:=True)
      .Caption = "�^�X�N�ɃX�N���[��"
      .OnAction = "menu.M_�^�X�N�ɃX�N���[��"
  End With
  
  With Application.CommandBars("Cell").Controls.Add(Temporary:=True)
      .Caption = "�^�C�����C���ɒǉ�"
      .OnAction = "menu.M_�^�C�����C���ɒǉ�"
  End With
  
  With Application.CommandBars("Cell").Controls.Add(Temporary:=True)
      .BeginGroup = True
      .Caption = "�^�X�N�̃��x���グ"
      .FaceId = 3161
      .OnAction = "menu.M_�C���f���g��"
  End With

  With Application.CommandBars("Cell").Controls.Add(Temporary:=True)
      .Caption = "�^�X�N�̃��x������"
      .FaceId = 3162
      .OnAction = "menu.M_�C���f���g��"
  End With

  With Application.CommandBars("Cell").Controls.Add(Temporary:=True)
      .BeginGroup = True
      .Caption = "�^�X�N�̑}��"
      .FaceId = 296
      .OnAction = "menu.M_�^�X�N�̑}��"
  End With

  With Application.CommandBars("Cell").Controls.Add(Temporary:=True)
      .Caption = "�^�X�N�̍폜"
      .FaceId = 293
      .OnAction = "menu.M_�^�X�N�̍폜"
  End With


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

  If ActiveSheet.Name = mainSheetName Or ActiveSheet.Name = TeamsPlannerSheetName Then
  Else
    Call Library.showNotice(404, , True)
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
  
  
  On Error GoTo catchError
  
  endColLine = Cells(4, Columns.count).End(xlToLeft).Column
  ���t�Z������ = Library.getColumnName(Range(Cells(4, Library.getColumnNo(setVal("calendarStartCol"))), Cells(4, endColLine)).Find(chkDay).Column)



'  Set chkCell = Range( _
'                       Cells(4, Library.getColumnNo(setVal("calendarStartCol"))), Cells(4, endColLine) _
'                                            ).Find(chkDay, SearchOrder:=xlByColumns)
'
'���t�Z������ = chkCell.Column
  
  Exit Function
'�G���[������=====================================================================================
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
    If ActiveSheet.Name = mainSheetName Then
      targetArea = "A6:" & setVal("calendarStartCol") & endLine
    ElseIf ActiveSheet.Name = TeamsPlannerSheetName Then
      targetArea = "F6:" & setVal("calendarStartCol") & endLine
    End If
    
    Call Library.setLineColor(targetArea, False, setVal("lineColor"))
    
    '�J�����_�[�G���A
    targetArea = setVal("calendarStartCol") & "4:" & Library.getColumnName(endColLine) & endLine
    Call Library.setLineColor(targetArea, True, setVal("lineColor"))
  
    setVal("lineColorFlg") = True
  End If
  
  setSheet.Range("lineColorFlg") = setVal("lineColorFlg")
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
'  setSheet.Range("O3:P" & setSheet.Cells(Rows.count, 15).End(xlUp).row + 1).ClearContents
  
  Application.Goto Reference:=Range("A6"), Scroll:=True
    
End Function

'**************************************************************************************************
' * �V�[�g���̑S�f�[�^�폜
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function clearCalendar()
  Call init.setting
  mainSheet.Select
  Columns(setVal("calendarStartCol") & ":XFD").Delete Shift:=xlToLeft
  
  '�S�̂̐i���Ȃǂ��폜
  Range("I5:" & setVal("cell_Note") & 5).ClearContents
  Application.Goto Reference:=Range("A6"), Scroll:=True
  
End Function

'**************************************************************************************************
' * ���[�U�[�t�H�[���p�̉摜�쐬
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function �I�v�V������ʕ\��()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  Dim images As Variant, tmpObjChart As Variant
  Dim CompanyHolidayList As String, dataExtractList As String
  
'  On Error GoTo catchError
  
  mainSheet.Select
  
  With optionForm
    .StartUpPosition = 0
    .top = Application.top + (ActiveWindow.Width / 8)
    .Left = Application.Left + (ActiveWindow.Height / 8)
    
    '�}���`�y�[�W�̕\��
    .MultiPage1.Value = 0
    
    '���ԁA����̏����l
    .startDay.Text = setVal("startDay")
    .endDay.Text = setVal("endDay")
    .baseDay.Text = setVal("baseDay")
    
    .setLightning.Value = setVal("setLightning")
    .setDispProgress100.Value = setVal("setDispProgress100")
        
    '�X�^�C���֘A
    .lineColor.BackColor = setVal("lineColor")
    .SaturdayColor.BackColor = setVal("SaturdayColor")
    .SundayColor.BackColor = setVal("SundayColor")
    .CompanyHolidayColor.BackColor = setVal("CompanyHolidayColor")
    .lineColor_Plan.BackColor = setVal("lineColor_Plan")
    .lineColor_Achievement.BackColor = setVal("lineColor_Achievement")
    .lineColor_Lightning.BackColor = setVal("lineColor_Lightning")
    .lineColor_TaskLevel1.BackColor = setVal("lineColor_TaskLevel1")
    .lineColor_TaskLevel2.BackColor = setVal("lineColor_TaskLevel2")
    .lineColor_TaskLevel3.BackColor = setVal("lineColor_TaskLevel3")
    
    
    '�V���[�g�J�b�g�L�[�֘A
    .optionKey.Value = setVal("optionKey")
    .centerKey.Value = setVal("centerKey")
    .filterKey.Value = setVal("filterKey")
    .clearFilterKey.Value = setVal("clearFilterKey")
    .taskCheckKey.Value = setVal("taskCheckKey")
    .makeGanttKey.Value = setVal("makeGanttKey")
    .clearGanttKey.Value = setVal("clearGanttKey")
    .dispAllKey.Value = setVal("dispAllKey")
    .taskControlKey.Value = setVal("taskControlKey")
    .ScaleKey.Value = setVal("ScaleKey")
    
    '�S����
    .Assign01.Text = setSheet.Range(setVal("cell_AssignorList") & 4)
    .Assign02.Text = setSheet.Range(setVal("cell_AssignorList") & 5)
    .Assign03.Text = setSheet.Range(setVal("cell_AssignorList") & 6)
    .Assign04.Text = setSheet.Range(setVal("cell_AssignorList") & 7)
    .Assign05.Text = setSheet.Range(setVal("cell_AssignorList") & 8)
    .Assign06.Text = setSheet.Range(setVal("cell_AssignorList") & 9)
    .Assign07.Text = setSheet.Range(setVal("cell_AssignorList") & 10)
    .Assign08.Text = setSheet.Range(setVal("cell_AssignorList") & 11)
    .Assign09.Text = setSheet.Range(setVal("cell_AssignorList") & 12)
    .Assign10.Text = setSheet.Range(setVal("cell_AssignorList") & 13)
    .Assign11.Text = setSheet.Range(setVal("cell_AssignorList") & 14)
    .Assign12.Text = setSheet.Range(setVal("cell_AssignorList") & 15)
    .Assign13.Text = setSheet.Range(setVal("cell_AssignorList") & 16)
    .Assign14.Text = setSheet.Range(setVal("cell_AssignorList") & 17)
    .Assign15.Text = setSheet.Range(setVal("cell_AssignorList") & 18)
    .Assign16.Text = setSheet.Range(setVal("cell_AssignorList") & 19)
    .Assign17.Text = setSheet.Range(setVal("cell_AssignorList") & 20)
    .Assign18.Text = setSheet.Range(setVal("cell_AssignorList") & 21)
    .Assign19.Text = setSheet.Range(setVal("cell_AssignorList") & 22)
    .Assign20.Text = setSheet.Range(setVal("cell_AssignorList") & 23)
    .Assign21.Text = setSheet.Range(setVal("cell_AssignorList") & 24)
    .Assign22.Text = setSheet.Range(setVal("cell_AssignorList") & 25)
    .Assign23.Text = setSheet.Range(setVal("cell_AssignorList") & 26)
    .Assign24.Text = setSheet.Range(setVal("cell_AssignorList") & 27)
    .Assign25.Text = setSheet.Range(setVal("cell_AssignorList") & 28)
    .Assign26.Text = setSheet.Range(setVal("cell_AssignorList") & 29)
    .Assign27.Text = setSheet.Range(setVal("cell_AssignorList") & 30)
    .Assign28.Text = setSheet.Range(setVal("cell_AssignorList") & 31)
    .Assign29.Text = setSheet.Range(setVal("cell_AssignorList") & 32)
    .Assign30.Text = setSheet.Range(setVal("cell_AssignorList") & 33)
    .Assign31.Text = setSheet.Range(setVal("cell_AssignorList") & 34)
    .Assign32.Text = setSheet.Range(setVal("cell_AssignorList") & 35)
    .Assign33.Text = setSheet.Range(setVal("cell_AssignorList") & 36)
    .Assign34.Text = setSheet.Range(setVal("cell_AssignorList") & 37)
    .Assign35.Text = setSheet.Range(setVal("cell_AssignorList") & 38)
    
    '�S���ҐF
    .AssignColor01.BackColor = setSheet.Range(setVal("cell_AssignorList") & 4).Interior.Color
    .AssignColor02.BackColor = setSheet.Range(setVal("cell_AssignorList") & 5).Interior.Color
    .AssignColor03.BackColor = setSheet.Range(setVal("cell_AssignorList") & 6).Interior.Color
    .AssignColor04.BackColor = setSheet.Range(setVal("cell_AssignorList") & 7).Interior.Color
    .AssignColor05.BackColor = setSheet.Range(setVal("cell_AssignorList") & 8).Interior.Color
    .AssignColor06.BackColor = setSheet.Range(setVal("cell_AssignorList") & 9).Interior.Color
    .AssignColor07.BackColor = setSheet.Range(setVal("cell_AssignorList") & 10).Interior.Color
    .AssignColor08.BackColor = setSheet.Range(setVal("cell_AssignorList") & 11).Interior.Color
    .AssignColor09.BackColor = setSheet.Range(setVal("cell_AssignorList") & 12).Interior.Color
    .AssignColor10.BackColor = setSheet.Range(setVal("cell_AssignorList") & 13).Interior.Color
    .AssignColor11.BackColor = setSheet.Range(setVal("cell_AssignorList") & 14).Interior.Color
    .AssignColor12.BackColor = setSheet.Range(setVal("cell_AssignorList") & 15).Interior.Color
    .AssignColor13.BackColor = setSheet.Range(setVal("cell_AssignorList") & 16).Interior.Color
    .AssignColor14.BackColor = setSheet.Range(setVal("cell_AssignorList") & 17).Interior.Color
    .AssignColor15.BackColor = setSheet.Range(setVal("cell_AssignorList") & 18).Interior.Color
    .AssignColor16.BackColor = setSheet.Range(setVal("cell_AssignorList") & 19).Interior.Color
    .AssignColor17.BackColor = setSheet.Range(setVal("cell_AssignorList") & 20).Interior.Color
    .AssignColor18.BackColor = setSheet.Range(setVal("cell_AssignorList") & 21).Interior.Color
    .AssignColor19.BackColor = setSheet.Range(setVal("cell_AssignorList") & 22).Interior.Color
    .AssignColor20.BackColor = setSheet.Range(setVal("cell_AssignorList") & 23).Interior.Color
    .AssignColor21.BackColor = setSheet.Range(setVal("cell_AssignorList") & 24).Interior.Color
    .AssignColor22.BackColor = setSheet.Range(setVal("cell_AssignorList") & 25).Interior.Color
    .AssignColor23.BackColor = setSheet.Range(setVal("cell_AssignorList") & 26).Interior.Color
    .AssignColor24.BackColor = setSheet.Range(setVal("cell_AssignorList") & 27).Interior.Color
    .AssignColor25.BackColor = setSheet.Range(setVal("cell_AssignorList") & 28).Interior.Color
    .AssignColor26.BackColor = setSheet.Range(setVal("cell_AssignorList") & 29).Interior.Color
    .AssignColor27.BackColor = setSheet.Range(setVal("cell_AssignorList") & 30).Interior.Color
    .AssignColor28.BackColor = setSheet.Range(setVal("cell_AssignorList") & 31).Interior.Color
    .AssignColor29.BackColor = setSheet.Range(setVal("cell_AssignorList") & 32).Interior.Color
    .AssignColor30.BackColor = setSheet.Range(setVal("cell_AssignorList") & 33).Interior.Color
    .AssignColor31.BackColor = setSheet.Range(setVal("cell_AssignorList") & 34).Interior.Color
    .AssignColor32.BackColor = setSheet.Range(setVal("cell_AssignorList") & 35).Interior.Color
    .AssignColor33.BackColor = setSheet.Range(setVal("cell_AssignorList") & 36).Interior.Color
    .AssignColor34.BackColor = setSheet.Range(setVal("cell_AssignorList") & 37).Interior.Color
    .AssignColor35.BackColor = setSheet.Range(setVal("cell_AssignorList") & 38).Interior.Color
    
    '�S���ҒP��
    .unitCost01.Text = setSheet.Range(setVal("cell_unitCostorList") & 4)
    .unitCost02.Text = setSheet.Range(setVal("cell_unitCostorList") & 5)
    .unitCost03.Text = setSheet.Range(setVal("cell_unitCostorList") & 6)
    .unitCost04.Text = setSheet.Range(setVal("cell_unitCostorList") & 7)
    .unitCost05.Text = setSheet.Range(setVal("cell_unitCostorList") & 8)
    .unitCost06.Text = setSheet.Range(setVal("cell_unitCostorList") & 9)
    .unitCost07.Text = setSheet.Range(setVal("cell_unitCostorList") & 10)
    .unitCost08.Text = setSheet.Range(setVal("cell_unitCostorList") & 11)
    .unitCost09.Text = setSheet.Range(setVal("cell_unitCostorList") & 12)
    .unitCost10.Text = setSheet.Range(setVal("cell_unitCostorList") & 13)
    .unitCost11.Text = setSheet.Range(setVal("cell_unitCostorList") & 14)
    .unitCost12.Text = setSheet.Range(setVal("cell_unitCostorList") & 15)
    .unitCost13.Text = setSheet.Range(setVal("cell_unitCostorList") & 16)
    .unitCost14.Text = setSheet.Range(setVal("cell_unitCostorList") & 17)
    .unitCost15.Text = setSheet.Range(setVal("cell_unitCostorList") & 18)
    .unitCost16.Text = setSheet.Range(setVal("cell_unitCostorList") & 19)
    .unitCost17.Text = setSheet.Range(setVal("cell_unitCostorList") & 20)
    .unitCost18.Text = setSheet.Range(setVal("cell_unitCostorList") & 21)
    .unitCost19.Text = setSheet.Range(setVal("cell_unitCostorList") & 22)
    .unitCost20.Text = setSheet.Range(setVal("cell_unitCostorList") & 23)
    .unitCost21.Text = setSheet.Range(setVal("cell_unitCostorList") & 24)
    .unitCost22.Text = setSheet.Range(setVal("cell_unitCostorList") & 25)
    .unitCost23.Text = setSheet.Range(setVal("cell_unitCostorList") & 26)
    .unitCost24.Text = setSheet.Range(setVal("cell_unitCostorList") & 27)
    .unitCost25.Text = setSheet.Range(setVal("cell_unitCostorList") & 28)
    .unitCost26.Text = setSheet.Range(setVal("cell_unitCostorList") & 29)
    .unitCost27.Text = setSheet.Range(setVal("cell_unitCostorList") & 30)
    .unitCost28.Text = setSheet.Range(setVal("cell_unitCostorList") & 31)
    .unitCost29.Text = setSheet.Range(setVal("cell_unitCostorList") & 32)
    .unitCost30.Text = setSheet.Range(setVal("cell_unitCostorList") & 33)
    .unitCost31.Text = setSheet.Range(setVal("cell_unitCostorList") & 34)
    .unitCost32.Text = setSheet.Range(setVal("cell_unitCostorList") & 35)
    .unitCost33.Text = setSheet.Range(setVal("cell_unitCostorList") & 36)
    .unitCost34.Text = setSheet.Range(setVal("cell_unitCostorList") & 37)
    .unitCost35.Text = setSheet.Range(setVal("cell_unitCostorList") & 38)

    
    '��Ўw��x��
    For line = 3 To setSheet.Cells(Rows.count, Library.getColumnNo(setVal("cell_CompanyHoliday"))).End(xlUp).row
      If setSheet.Range(setVal("cell_CompanyHoliday") & line) <> "" Then
        If CompanyHolidayList = "" Then
          CompanyHolidayList = setSheet.Range(setVal("cell_CompanyHoliday") & line)
        Else
          CompanyHolidayList = CompanyHolidayList & vbNewLine & setSheet.Range(setVal("cell_CompanyHoliday") & line)
        End If
      End If
    Next
    .CompanyHoliday.Text = CompanyHolidayList
    
    '���o�^�X�N
    For line = 3 To setSheet.Cells(Rows.count, Library.getColumnNo(setVal("cell_DataExtract"))).End(xlUp).row
      If setSheet.Range(setVal("cell_DataExtract") & line) <> "" Then
        If dataExtractList = "" Then
          dataExtractList = setSheet.Range(setVal("cell_DataExtract") & line)
        Else
          dataExtractList = dataExtractList & vbNewLine & setSheet.Range(setVal("cell_DataExtract") & line)
        End If
      End If
    Next
    .dataExtract.Text = dataExtractList
    
    
    '�\���ݒ�
    .view_Plan.Value = setVal("view_Plan")
    .view_Assign.Value = setVal("view_Assign")
    .view_Progress.Value = setVal("view_Progress")
    .view_Achievement.Value = setVal("view_Achievement")
    .view_Task.Value = setVal("view_Task")
    .view_TaskInfo.Value = setVal("view_TaskInfo")
    .view_TaskAllocation.Value = setVal("view_TaskAllocation")
    .view_LineInfo.Value = setVal("view_LineInfo")
    
    .view_WorkLoad.Value = setVal("view_WorkLoad")
    .view_LateOrEarly.Value = setVal("view_LateOrEarly")
    .view_Note.Value = setVal("view_Note")
    
    .viewGant_TaskName.Value = setVal("viewGant_TaskName")
    .viewGant_Assignor.Value = setVal("viewGant_Assignor")
  
  End With
  
  optionForm.Show

  Exit Function
'�G���[������=====================================================================================
catchError:

End Function


'**************************************************************************************************
' * �ݒ�l�i�[
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function �I�v�V�����ݒ�l�i�[()

  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  Dim CompanyHoliday As Variant, dataExtract As Variant

  On Error Resume Next
  
  Call ProgressBar.showStart
  setSheet.Select
  For line = 3 To setSheet.Range("B5")
    Call ProgressBar.showCount("�I�v�V�����ݒ�l�i�[", line, setSheet.Range("B5"), setSheet.Range("A" & line))
    
    Select Case setSheet.Range("A" & line)
      Case "baseDay"
        If getVal(setSheet.Range("A" & line)) = Format(Now, "yyyy/mm/dd") Then
          setSheet.Range(setVal("cell_LevelInfo") & line).FormulaR1C1 = "=TODAY()"
        Else
          setSheet.Range(setVal("cell_LevelInfo") & line) = getVal(setSheet.Range("A" & line))
        End If
      
      Case ""
      Case Else
        setSheet.Range(setVal("cell_LevelInfo") & line) = getVal(setSheet.Range("A" & line))
    End Select
  Next
  
  '�V���[�g�J�b�g�L�[�̐ݒ�
  endLine = Cells(Rows.count, Library.getColumnNo(setVal("cell_ShortcutFuncName"))).End(xlUp).row
  For line = 3 To endLine
    Call ProgressBar.showCount("�I�v�V�����ݒ�l�i�[", line, setSheet.Range("B5"), "�V���[�g�J�b�g�L�[�ݒ�")
    
    Range(Range(setVal("cell_ShortcutFuncName") & line)).Select
    Range(Range(setVal("cell_ShortcutFuncName") & line)) = getVal(Range(setVal("cell_ShortcutFuncName") & line))
  Next
  
  '��Ўw��x���̐ݒ�
  line = 3
  setSheet.Range(setVal("cell_CompanyHoliday") & "3:" & setVal("cell_CompanyHoliday") & Cells(Rows.count, Library.getColumnNo(setVal("cell_CompanyHoliday"))).End(xlUp).row).ClearContents
  For Each CompanyHoliday In Split(getVal("CompanyHoliday"), vbNewLine)
    DoEvents
    setSheet.Range(setVal("cell_CompanyHoliday") & line) = CompanyHoliday
    line = line + 1
  Next

  '���o�^�X�N�̐ݒ�
  line = 3
  setSheet.Range(setVal("cell_DataExtract") & "3:" & setVal("cell_DataExtract") & Cells(Rows.count, Library.getColumnNo(setVal("cell_DataExtract"))).End(xlUp).row).ClearContents
  For Each dataExtract In Split(getVal("dataExtract"), vbNewLine)
    Call ProgressBar.showCount("�I�v�V�����ݒ�l�i�[", line, 100, "���o�^�X�N�̐ݒ�")
    
    setSheet.Range(setVal("cell_DataExtract") & line) = dataExtract
    line = line + 1
  Next


  '�S����
  setSheet.Range(setVal("cell_AssignorList") & "4:" & setVal("cell_AssignorList") & Cells(Rows.count, Library.getColumnNo(setVal("cell_AssignorList"))).End(xlUp).row).Clear
  For line = 4 To 38
    Call ProgressBar.showCount("�I�v�V�����ݒ�l�i�[", line, 38, "�S���҂̐ݒ�")
    
    setSheet.Range(setVal("cell_AssignorList") & line) = getVal("Assign" & Format(line - 3, "00"))
    setSheet.Range(setVal("cell_AssignorList") & line).Interior.Color = getVal("AssignColor" & Format(line - 3, "00"))
    
    setSheet.Range(setVal("cell_unitCostorList") & line) = getVal("unitCost" & Format(line - 3, "00"))
    
  Next
  setSheet.Range(setVal("cell_AssignorList") & "3:" & setVal("cell_AssignorList") & 38).Select
  Call �r��.�͂݌r��
  Call menu.M_�V���[�g�J�b�g�ݒ�
  
  Application.Goto Reference:=Range("A1"), Scroll:=True
  mainSheet.Select
  

  
  Set getVal = Nothing
  
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
  mainSheet.Select
  
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
'�G���[������=====================================================================================
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
  endLine = TeamsPlannerSheet.Cells(Rows.count, 1).End(xlUp).row
  
  Rows("6:" & Rows.count).EntireRow.Hidden = False
  
  '�`�[���v�����i�[�ŕύX�����\������i�[
  For line = 6 To endLine
    If TeamsPlannerSheet.Range(("C") & line) Like "*" & setVal("TaskInfoStr_Change") & "*" Then
      rowLine = TeamsPlannerSheet.Range(("D") & line) + 5
      
      mainSheet.Range(setVal("cell_PlanStart") & rowLine) = TeamsPlannerSheet.Range(("G") & line)
      mainSheet.Range(setVal("cell_PlanEnd") & rowLine) = TeamsPlannerSheet.Range(("H") & line)
    End If
  Next
  
  mainSheet.Visible = True
  TeamsPlannerSheet.Visible = xlSheetVeryHidden
    
  mainSheet.Select
  mainSheet.ScrollArea = ""
  Cells.EntireColumn.Hidden = False

  '�E�C���h�E�g�̌Œ�
  Range(setVal("calendarStartCol") & 6).Select
  ActiveWindow.FreezePanes = False
  ActiveWindow.FreezePanes = True
  
  Call Chart.�K���g�`���[�g����
  Call WBS_Option.�����̒S���ҍs���\��
  Call �\����ݒ�
  
  

  Exit Function
'�G���[������=====================================================================================
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
  
  mainSheet.Columns(setVal("calendarStartCol") & ":" & Library.getColumnName(Cells(4, Columns.count).End(xlToLeft).Column)).EntireColumn.Hidden = True
  mainSheet.ScrollArea = "A1:P" & Rows.count
  
  '�E�C���h�E�g�̌Œ�
  Range("A6").Select
  ActiveWindow.FreezePanes = True
    
    
  Call Library.endScript(True)

  Exit Function
'�G���[������=====================================================================================
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
  
  
  TeamsPlannerSheet.Visible = True
  
  Cells.EntireColumn.Hidden = False
  Call TeamsPlanner.�f�[�^�ڍs
  
  If setVal("debugMode") <> "develop" Then
    TeamsPlannerSheet.Columns("I:S").EntireColumn.Hidden = True
  End If
  
  mainSheet.Visible = xlSheetVeryHidden
  Call Library.endScript

  Exit Function
'�G���[������=====================================================================================
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
'�G���[������=====================================================================================
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
'�G���[������=====================================================================================
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
'�G���[������=====================================================================================
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
  
  With AssignorForm
    .StartUpPosition = 0
    .top = Application.top + (ActiveWindow.Width / 8)
    .Left = Application.Left + (ActiveWindow.Height / 8)
   .Show vbModeless
  End With

  Exit Function
'�G���[������=====================================================================================
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
'�G���[������=====================================================================================
catchError:
  Call Library.showNotice(Err.Number, Err.Description, True)
End Function

'***********************************************************************************************************************************************
' * �^�X�N���x���̐ݒ�
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'***********************************************************************************************************************************************
Function �^�X�N���x���̐ݒ�()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long


  If ActiveSheet.Name = mainSheetName Then
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
'�G���[������=====================================================================================
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
'�G���[������=====================================================================================
catchError:
  Call Library.showNotice(Err.Number, Err.Description, True)
End Function
