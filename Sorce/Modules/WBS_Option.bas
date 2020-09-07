Attribute VB_Name = "WBS_Option"
' *************************************************************************************************
' * �J�����_�[�֘A�֐�
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
' *************************************************************************************************
Function �I���V�[�g�m�F()

  If ActiveSheet.Name = "���C��" Or ActiveSheet.Name = "���\�[�X" Then
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
  
  On Error GoTo catchError
  
  endColLine = Cells(4, Columns.count).End(xlToLeft).Column
  ���t�Z������ = Library.getColumnName(Range(Cells(4, Library.getColumnNo(setVal("calendarStartCol"))), Cells(4, endColLine)).Find(chkDay).Column)

  
  Exit Function
'�G���[������=====================================================================================
catchError:
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
    targetArea = "A6:" & setVal("calendarStartCol") & endLine
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
  Dim CompanyHolidayList As String
  
'  On Error GoTo catchError
  
  Call Library.startScript
  '�V�[�g���̉摜���t�@�C���Ƃ��ĕۑ�����
'  For Each images In setSheet.Shapes
'    If images.Name = "msoLineSingle" Or images.Name = "msoLineThinThin" Then
'      Set tmpObjChart = setSheet.ChartObjects.Add(0, 0, images.Width, images.Height).Chart
'
'      images.CopyPicture
'      For Each zu In setSheet.ChartObjects
'       cname = zu.Name
'      Next
'      setSheet.ChartObjects(cname).Activate
'      ActiveChart.Paste
'      tmpObjChart.Export FileName:=ThisWorkbook.Path & "\" & images.Name & ".jpg", filterName:="JPG"
'      tmpObjChart.Parent.Delete
'      Set tmpObjChart = Nothing
'    End If
'  Next
  mainSheet.Select
  Call Library.endScript
  
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

'    '�K���g�`���[�g�̐��`�摜�T���v���̓ǂݍ���
'    .ganttChartLineTypeImg1.Picture = LoadPicture(ThisWorkbook.Path & "\" & "msoLineSingle.jpg")
'    .ganttChartLineTypeImg2.Picture = LoadPicture(ThisWorkbook.Path & "\" & "msoLineThinThin.jpg")
    
'    If setVal("ganttChartLineType") = "Type1" Then
'      .ganttChartLineType1.Value = True
'    ElseIf setVal("ganttChartLineType") = "Type2" Then
'      .ganttChartLineType2.Value = True
'    End If
    
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
    
    
  End With
  
'  Kill ThisWorkbook.Path & "\" & "msoLineSingle.jpg"
'  Kill ThisWorkbook.Path & "\" & "msoLineThinThin.jpg"
  
  
  'optionForm.Show vbModeless
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
  Dim CompanyHoliday As Variant

  On Error Resume Next
  Call Library.startScript
  
  setSheet.Select
  For line = 3 To 28
    Select Case setSheet.Range("A" & line)
      Case "baseDay"
        If getVal(setSheet.Range("A" & line)) = Format(Now, "yyyy/mm/dd") Then
          setSheet.Range("B" & line).FormulaR1C1 = "=TODAY()"
        Else
          setSheet.Range("B" & line) = getVal(setSheet.Range("A" & line))
        End If
      
      Case ""
      Case Else
        setSheet.Range("B" & line) = getVal(setSheet.Range("A" & line))
    End Select
  Next
  
  '��Ўw��x���̐ݒ�
  line = 3
  setSheet.Range("M3:M" & Cells(Rows.count, 13).End(xlUp).row).ClearContents
  For Each CompanyHoliday In Split(getVal("CompanyHoliday"), vbNewLine)
    setSheet.Range("M" & line) = CompanyHoliday
    line = line + 1
  Next CompanyHoliday
  setSheet.Range("M3:M37").Select
  Call �r��.�͂݌r��



  '�S����
  setSheet.Range("K3:K" & Cells(Rows.count, 11).End(xlUp).row).Clear
  For line = 3 To 37
    setSheet.Range("K" & line) = getVal("Assign" & Format(line - 2, "00"))
    setSheet.Range("K" & line).Interior.Color = getVal("AssignColor" & Format(line - 2, "00"))
  Next
  setSheet.Range("K3:K37").Select
  Call �r��.�͂݌r��
  
  Application.Goto Reference:=Range("A1"), Scroll:=True
  mainSheet.Select
  Call Library.endScript
End Function


'**************************************************************************************************
' * �G���[���\��
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function �G���[���\��(errorMeg As String)

  With ErrorForm
    .StartUpPosition = 0
    .top = Application.top + (ActiveWindow.Width / 8)
    .Left = Application.Left + (ActiveWindow.Height / 8)
    .errMeg.Text = errorMeg
  End With
  
  ErrorForm.Show vbModeless

End Function


'**************************************************************************************************
' * xxxxxxxxxx
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function ��\����ݒ�()
  Dim line As Long, endLine As Long
  
'  On Error GoTo catchError
  If setVal("debugMode") <> "develop" Then
    Exit Function
  End If

  endLine = setSheet.Cells(Rows.count, 4).End(xlUp).row
  For line = 3 To endLine
    If setSheet.Range("E" & line) = False Then
      Select Case setSheet.Range("D" & line)
      Case "view_Plan"
        Columns(setVal("cell_PlanStart") & ":" & setVal("cell_PlanEnd")).EntireColumn.Hidden = True
      
      Case "view_Assign"
        Columns(setVal("cell_Assign") & ":" & setVal("cell_Assign")).EntireColumn.Hidden = True
      
      Case "view_Progress"
        Columns(setVal("cell_Assign") & ":" & setVal("cell_Assign")).EntireColumn.Hidden = True
      
      Case "view_Achievement"
        Columns(setVal("cell_AchievementStart") & ":" & setVal("cell_AchievementEnd")).EntireColumn.Hidden = True
      
      Case "view_Task"
        Columns(setVal("cell_Task") & ":" & setVal("cell_Task")).EntireColumn.Hidden = True
      
      Case "view_TaskInfo"
        Columns(setVal("cell_TaskInfoP") & ":" & setVal("cell_TaskInfoC")).EntireColumn.Hidden = True
      
      Case "view_WorkLoad"
        Columns(setVal("cell_WorkLoadP") & ":" & setVal("cell_WorkLoadA")).EntireColumn.Hidden = True
      
      Case "view_LateOrEarly"
        Columns(setVal("cell_LateOrEarly") & ":" & setVal("cell_LateOrEarly")).EntireColumn.Hidden = True
      
      Case "view_Note"
        Columns(setVal("cell_Note") & ":" & setVal("cell_Note")).EntireColumn.Hidden = True
      
      Case Else
      End Select
    End If
  Next
  



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
Function viewNormal()
  On Error GoTo catchError

  mainSheet.Select
  mainSheet.ScrollArea = ""
  Cells.EntireColumn.Hidden = False

  '�E�C���h�E�g�̌Œ�
  Range(setVal("calendarStartCol") & 6).Select
  ActiveWindow.FreezePanes = False
  ActiveWindow.FreezePanes = True
  
  Call ��\����ݒ�
  
  

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
' * xxxxxxxxxx
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function viewResources()
  On Error GoTo catchError

  If setVal("debugMode") <> "develop" Then
    Worksheets("���C��").Visible = xlSheetVeryHidden
    Worksheets("���\�[�X").Visible = True
  End If
  ResourcesSheet.Select
  
  Cells.EntireColumn.Hidden = False
  Call Resources.�f�[�^�ڍs
  
  If setVal("debugMode") <> "develop" Then
    Columns("L:Q").EntireColumn.Hidden = True
  End If

  
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
  Range("B" & line).FormulaR1C1 = "=getIndentLevel(" & taskLevelRange.Address(ReferenceStyle:=xlR1C1) & ")"
  Set taskLevelRange = Nothing

  Range("D" & line).Select


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
  On Error GoTo catchError

  endLine = Cells(Rows.count, 1).End(xlUp).row
  
  
  AssignorForm.Show vbModeless

  Exit Function
'�G���[������=====================================================================================
catchError:
  Call Library.showNotice(Err.Number, Err.Description, True)
End Function










