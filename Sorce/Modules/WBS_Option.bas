Attribute VB_Name = "WBS_Option"
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
' * 日付セル検索
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function 日付セル検索(chkDay As Date, Optional chlkFlg As Boolean)
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  
  On Error GoTo catchError
  
  endColLine = mainSheet.Cells(4, Columns.count).End(xlToLeft).Column
  日付セル検索 = Library.getColumnName(mainSheet.Range(Cells(4, Library.getColumnNo(setVal("calendarStartCol"))), Cells(4, endColLine)).Find(chkDay).Column)

  
  Exit Function
'エラー発生時=====================================================================================
catchError:
  日付セル検索 = setVal("calendarStartCol")

End Function


'**************************************************************************************************
' * イナズマ線用日付計算
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function イナズマ線用日付計算(baseDay As Date, calDay As Double) As Date
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
      chk = WorksheetFunction.VLookup(CLng(resultDay), Range("休日リスト"), 2, False)
      On Error GoTo 0
      
      If Weekday(resultDay) = 1 Or Weekday(resultDay) = 7 Then
        chk = "土日"
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
 イナズマ線用日付計算 = resultDay
End Function


'**************************************************************************************************
' * 選択行の色付切替
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function setLineColor()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  Dim targetArea As String
  Dim setFlg As String
  
  Call init.setting
  mainSheet.Select
    
  endLine = mainSheet.Cells(Rows.count, 1).End(xlUp).row
  endColLine = mainSheet.Cells(4, Columns.count).End(xlToLeft).Column
  
  setFlg = setVal("lineColorFlg")
  
  If setFlg = "True" Then
    targetArea = "A4:" & Library.getColumnName(endColLine) & endLine
    Call Library.unsetLineColor(targetArea)
    
    setVal("lineColorFlg") = False
  Else
    'タスクエリア
    targetArea = "A6:" & setVal("calendarStartCol") & endLine
    Call Library.setLineColor(targetArea, False, setVal("lineColor"))
    
    'カレンダーエリア
    targetArea = setVal("calendarStartCol") & "4:" & Library.getColumnName(endColLine) & endLine
    Call Library.setLineColor(targetArea, True, setVal("lineColor"))
  
    setVal("lineColorFlg") = True
  End If
  
  setSheet.Range("lineColorFlg") = setVal("lineColorFlg")
End Function

'**************************************************************************************************
' * シート内の全データ削除
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function clearAll()
  Call init.setting
  mainSheet.Select
  Call Library.delSheetData(6)
  
  Columns(setVal("calendarStartCol") & ":XFD").Delete Shift:=xlToLeft
  
  '全体の進捗などを削除
  Range("I5:" & setVal("cell_Note") & 5).ClearContents
  
  
  Range(setVal("calendarStartCol") & "1:" & setVal("calendarStartCol") & 5).Borders(xlEdgeLeft).LineStyle = xlDouble
'  setSheet.Range("O3:P" & setSheet.Cells(Rows.count, 15).End(xlUp).row + 1).ClearContents
  
  Application.Goto Reference:=Range("A6"), Scroll:=True
    
End Function

'**************************************************************************************************
' * シート内の全データ削除
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function clearCalendar()
  Call init.setting
  mainSheet.Select
  Columns(setVal("calendarStartCol") & ":XFD").Delete Shift:=xlToLeft
  
  '全体の進捗などを削除
  Range("I5:" & setVal("cell_Note") & 5).ClearContents
  Application.Goto Reference:=Range("A6"), Scroll:=True
  
End Function

'**************************************************************************************************
' * ユーザーフォーム用の画像作成
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function オプション画面表示()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  Dim images As Variant, tmpObjChart As Variant
  Dim CompanyHolidayList As String
  
'  On Error GoTo catchError
  
  Call Library.startScript
  'シート内の画像をファイルとして保存する
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
    
    'マルチページの表示
    .MultiPage1.Value = 0
    
    '期間、基準日の初期値
    .startDay.Text = setVal("startDay")
    .endDay.Text = setVal("endDay")
    .baseDay.Text = setVal("baseDay")
    
    .setLightning.Value = setVal("setLightning")
    .setDispProgress100.Value = setVal("setDispProgress100")

'    'ガントチャートの線形画像サンプルの読み込み
'    .ganttChartLineTypeImg1.Picture = LoadPicture(ThisWorkbook.Path & "\" & "msoLineSingle.jpg")
'    .ganttChartLineTypeImg2.Picture = LoadPicture(ThisWorkbook.Path & "\" & "msoLineThinThin.jpg")
    
'    If setVal("ganttChartLineType") = "Type1" Then
'      .ganttChartLineType1.Value = True
'    ElseIf setVal("ganttChartLineType") = "Type2" Then
'      .ganttChartLineType2.Value = True
'    End If
    
    'スタイル関連
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
    
    
    'ショートカットキー関連
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
    
    
    
    
    
    
    '担当者
    .Assign01.Text = setSheet.Range("K3")
    .Assign02.Text = setSheet.Range("K4")
    .Assign03.Text = setSheet.Range("K5")
    .Assign04.Text = setSheet.Range("K6")
    .Assign05.Text = setSheet.Range("K7")
    .Assign06.Text = setSheet.Range("K8")
    .Assign07.Text = setSheet.Range("K9")
    .Assign08.Text = setSheet.Range("K10")
    .Assign09.Text = setSheet.Range("K11")
    .Assign10.Text = setSheet.Range("K12")
    .Assign11.Text = setSheet.Range("K13")
    .Assign12.Text = setSheet.Range("K14")
    .Assign13.Text = setSheet.Range("K15")
    .Assign14.Text = setSheet.Range("K16")
    .Assign15.Text = setSheet.Range("K17")
    .Assign16.Text = setSheet.Range("K18")
    .Assign17.Text = setSheet.Range("K19")
    .Assign18.Text = setSheet.Range("K20")
    .Assign19.Text = setSheet.Range("K21")
    .Assign20.Text = setSheet.Range("K22")
    .Assign21.Text = setSheet.Range("K23")
    .Assign22.Text = setSheet.Range("K24")
    .Assign23.Text = setSheet.Range("K25")
    .Assign24.Text = setSheet.Range("K26")
    .Assign25.Text = setSheet.Range("K27")
    .Assign26.Text = setSheet.Range("K28")
    .Assign27.Text = setSheet.Range("K29")
    .Assign28.Text = setSheet.Range("K30")
    .Assign29.Text = setSheet.Range("K31")
    .Assign30.Text = setSheet.Range("K32")
    .Assign31.Text = setSheet.Range("K33")
    .Assign32.Text = setSheet.Range("K34")
    .Assign33.Text = setSheet.Range("K35")
    .Assign34.Text = setSheet.Range("K36")
    .Assign35.Text = setSheet.Range("K37")
    
    
    .AssignColor01.BackColor = setSheet.Range("K3").Interior.Color
    .AssignColor02.BackColor = setSheet.Range("K4").Interior.Color
    .AssignColor03.BackColor = setSheet.Range("K5").Interior.Color
    .AssignColor04.BackColor = setSheet.Range("K6").Interior.Color
    .AssignColor05.BackColor = setSheet.Range("K7").Interior.Color
    .AssignColor06.BackColor = setSheet.Range("K8").Interior.Color
    .AssignColor07.BackColor = setSheet.Range("K9").Interior.Color
    .AssignColor08.BackColor = setSheet.Range("K10").Interior.Color
    .AssignColor09.BackColor = setSheet.Range("K11").Interior.Color
    .AssignColor10.BackColor = setSheet.Range("K12").Interior.Color
    .AssignColor11.BackColor = setSheet.Range("K13").Interior.Color
    .AssignColor12.BackColor = setSheet.Range("K14").Interior.Color
    .AssignColor13.BackColor = setSheet.Range("K15").Interior.Color
    .AssignColor14.BackColor = setSheet.Range("K16").Interior.Color
    .AssignColor15.BackColor = setSheet.Range("K17").Interior.Color
    .AssignColor16.BackColor = setSheet.Range("K18").Interior.Color
    .AssignColor17.BackColor = setSheet.Range("K19").Interior.Color
    .AssignColor18.BackColor = setSheet.Range("K20").Interior.Color
    .AssignColor19.BackColor = setSheet.Range("K21").Interior.Color
    .AssignColor20.BackColor = setSheet.Range("K22").Interior.Color
    .AssignColor21.BackColor = setSheet.Range("K23").Interior.Color
    .AssignColor22.BackColor = setSheet.Range("K24").Interior.Color
    .AssignColor23.BackColor = setSheet.Range("K25").Interior.Color
    .AssignColor24.BackColor = setSheet.Range("K26").Interior.Color
    .AssignColor25.BackColor = setSheet.Range("K27").Interior.Color
    .AssignColor26.BackColor = setSheet.Range("K28").Interior.Color
    .AssignColor27.BackColor = setSheet.Range("K29").Interior.Color
    .AssignColor28.BackColor = setSheet.Range("K30").Interior.Color
    .AssignColor29.BackColor = setSheet.Range("K31").Interior.Color
    .AssignColor30.BackColor = setSheet.Range("K32").Interior.Color
    .AssignColor31.BackColor = setSheet.Range("K33").Interior.Color
    .AssignColor32.BackColor = setSheet.Range("K34").Interior.Color
    .AssignColor33.BackColor = setSheet.Range("K35").Interior.Color
    .AssignColor34.BackColor = setSheet.Range("K36").Interior.Color
    .AssignColor35.BackColor = setSheet.Range("K37").Interior.Color
  
    '会社指定休日
    For line = 3 To setSheet.Cells(Rows.count, 13).End(xlUp).row
      If setSheet.Range("M" & line) <> "" Then
        If CompanyHolidayList = "" Then
          CompanyHolidayList = setSheet.Range("M" & line)
        Else
          CompanyHolidayList = CompanyHolidayList & vbNewLine & setSheet.Range("M" & line)
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
'エラー発生時=====================================================================================
catchError:

End Function


'**************************************************************************************************
' * 設定値格納
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function オプション設定値格納()

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
  
  '会社指定休日の設定
  line = 3
  setSheet.Range("M3:M" & Cells(Rows.count, 13).End(xlUp).row).ClearContents
  For Each CompanyHoliday In Split(getVal("CompanyHoliday"), vbNewLine)
    setSheet.Range("M" & line) = CompanyHoliday
    line = line + 1
  Next CompanyHoliday
  setSheet.Range("M3:M37").Select
  Call 罫線.囲み罫線



  '担当者
  setSheet.Range("K3:K" & Cells(Rows.count, 11).End(xlUp).row).Clear
  For line = 3 To 37
    setSheet.Range("K" & line) = getVal("Assign" & Format(line - 2, "00"))
    setSheet.Range("K" & line).Interior.Color = getVal("AssignColor" & Format(line - 2, "00"))
  Next
  setSheet.Range("K3:K37").Select
  Call 罫線.囲み罫線
  
  Application.Goto Reference:=Range("A1"), Scroll:=True
  mainSheet.Select
  Call Library.endScript
End Function


'**************************************************************************************************
' * エラー情報表示
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function エラー情報表示(errorMeg As String)

  With ErrorForm
    .StartUpPosition = 0
    .top = Application.top + (ActiveWindow.Width / 8)
    .Left = Application.Left + (ActiveWindow.Height / 8)
    .errMeg.Text = errorMeg
  End With
  
  ErrorForm.Show vbModeless

End Function





'**************************************************************************************************
' * 表示_標準
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function 表示_標準()
    Cells.Select
    Selection.EntireColumn.Hidden = False

End Function


'**************************************************************************************************
' * 表示_ガントチャート
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function 表示_ガントチャート()
    Columns("F:Q").EntireColumn.Hidden = True
  
End Function
