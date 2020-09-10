Attribute VB_Name = "Chart"



'**************************************************************************************************
' * ガントチャート生成
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function ガントチャート生成()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  Dim startColumn As String, endColumn As String
  
  Call WBS_Option.選択シート確認
  
  
  Rows("6:" & Rows.count).EntireRow.Hidden = False
  
  Call ガントチャート削除
  endLine = Cells(Rows.count, 2).End(xlUp).row
  
  For line = 6 To endLine
    '計画線生成------------------------------------
    If Not (mainSheet.Range(setVal("cell_PlanStart") & line) = "" Or mainSheet.Range(setVal("cell_PlanEnd") & line) = "") Then
      Call 計画線設定(line)
    End If

    '実績線生成------------------------------------
    If Not (mainSheet.Range(setVal("cell_AchievementStart") & line) = "") Then
      Call 実績線設定(line)
    End If
    
    'タイムラインへの追加------------------------------------
    If (mainSheet.Range(setVal("cell_Info") & line) = setVal("TaskInfoStr_TimeLine")) Then
      Call タイムラインに追加(line)
    End If
    
    'イナズマ線生成------------------------------
    Call イナズマ線設定(line)
  Next
  For line = 6 To endLine
    Call タスクのリンク設定(line)
  Next

  If ActiveSheet.Name = mainSheetName Then
    Call WBS_Option.複数の担当者行を非表示
  End If

End Function


'**************************************************************************************************
' * ガントチャート削除
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function ガントチャート削除()
  Dim shp As Shape
  Dim rng As Range
  
  On Error Resume Next
'  mainSheet.Select
  endLine = Cells(Rows.count, 1).End(xlUp).row
  endColLine = Cells(4, Columns.count).End(xlToLeft).Column
  
  Set rng = Range(Cells(5, Library.getColumnNo(setVal("calendarStartCol"))), Cells(endLine, endColLine))
  
  For Each shp In ActiveSheet.Shapes
    If Not Intersect(Range(shp.TopLeftCell, shp.BottomRightCell), rng) Is Nothing Then
      If (shp.Name Like "Drop Down*") Or (shp.Name Like "Comment*") Then
      Else
        'Debug.Print shp.Name
        shp.Select
        shp.Delete
      End If
    End If
  Next

End Function


'**************************************************************************************************
' * 計画線設定
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function 計画線設定(line As Long)

  Dim rngStart As Range, rngEnd As Range
  Dim BX As Single, BY As Single, EX As Single, EY As Single
  Dim lColorValue As Long, Red As Long, Green As Long, Blue As Long
  Dim ProcessShape As Shape
  
  startColumn = WBS_Option.日付セル検索(Range(setVal("cell_PlanStart") & line))
  endColumn = WBS_Option.日付セル検索(Range(setVal("cell_PlanEnd") & line))
  
'  'Shapeを配置するための基準となるセル
'  Set rngStart = mainSheet.Range(startColumn & line)
'  Set rngEnd = mainSheet.Range(endColumn & line)
'
'  'セルのLeft、Top、Widthプロパティを利用して位置決め
'  BX = rngStart.Left
'  BY = rngStart.top + (rngStart.Height / 2)
'  EX = rngEnd.Left + rngEnd.Width
'  EY = rngEnd.top + (rngEnd.Height / 2)
  
  '担当者別の色設定------------------------------
  lColorValue = 0
  If Range(setVal("cell_Assign") & line) <> "" Then
    lColorValue = memberColor.item(Range(setVal("cell_Assign") & line).Value)
  ElseIf Range(setVal("cell_Assign") & line) <> "" Then
    lColorValue = memberColor.item(Range(setVal("cell_Assign") & line).Value)
  End If
  
  
  
  If lColorValue <> 0 And ActiveSheet.Name = mainSheetName Then
    Call Library.getRGB(lColorValue, Red, Green, Blue)
  Else
    Call Library.getRGB(setVal("lineColor_Plan"), Red, Green, Blue)
  End If

  If Range(setVal("cell_Assign") & line) = "工程" Or Range(setVal("cell_Assign") & line) = "工程" Then
    With Range(startColumn & line & ":" & endColumn & line)
      Set ProcessShape = ActiveSheet.Shapes.AddShape(Type:=msoShapePentagon, Left:=.Left, top:=.top, Width:=.Width, Height:=.Height)
      
      With ProcessShape
        .Name = "タスク_" & line
        .Fill.ForeColor.RGB = RGB(Red, Green, Blue)
        .Fill.Transparency = 0.6
'        .TextFrame.Characters.Text = Range(setVal("cell_TaskArea") & line)
'        .TextFrame.Characters.Font.Size = 12
        .TextFrame2.WordWrap = msoFalse
        .TextFrame.HorizontalOverflow = xlOartHorizontalOverflowOverflow
        .TextFrame.VerticalOverflow = xlOartHorizontalOverflowOverflow
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .TextFrame2.TextRange.Font.NameComplexScript = "メイリオ"
        .TextFrame2.TextRange.Font.NameFarEast = "メイリオ"
        .TextFrame2.TextRange.Font.Name = "メイリオ"
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
        .TextFrame2.TextRange.Font.Size = 12
        .TextFrame2.TextRange.Font.Bold = msoTrue
      End With
    End With
    Set ProcessShape = Nothing
    ActiveSheet.Shapes.Range(Array("タスク_" & line)).Select
    Selection.Formula = "=" & Range(setVal("cell_TaskArea") & line).Address(False, False)
    Selection.ShapeRange.TextFrame2.TextRange.Font.Size = 12
    Selection.ShapeRange.TextFrame2.TextRange.Font.Bold = msoTrue
  
  Else
    With Range(startColumn & line & ":" & endColumn & line)
      Set ProcessShape = ActiveSheet.Shapes.AddShape(Type:=msoShapeRectangle, Left:=.Left, top:=.top + 5, Width:=.Width, Height:=10)
      
      With ProcessShape
        .Name = "タスク_" & line
        .Fill.ForeColor.RGB = RGB(Red, Green, Blue)
        .Fill.Transparency = 0.6
        
        '.TextFrame.Characters.Text = Range(setVal("cell_TaskArea") & line)
        .TextFrame.Characters.Font.Size = 9
        .TextFrame2.TextRange.Font.Bold = msoTrue
        .TextFrame2.WordWrap = msoFalse
        .TextFrame.HorizontalOverflow = xlOartHorizontalOverflowOverflow
        .TextFrame.VerticalOverflow = xlOartHorizontalOverflowOverflow
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .TextFrame2.TextRange.Font.NameComplexScript = "メイリオ"
        .TextFrame2.TextRange.Font.NameFarEast = "メイリオ"
        .TextFrame2.TextRange.Font.Name = "メイリオ"
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
        .TextFrame2.TextRange.Font.Size = 9
    
        If setVal("viewGant_TaskName") = True Then
          ActiveSheet.Shapes.Range(Array("タスク_" & line)).Select
          Selection.Formula = "=" & Range(setVal("cell_TaskArea") & line).Address(False, False)
        End If
        
        .OnAction = "beforeChangeShapes"
      End With
    End With
    Set ProcessShape = Nothing

    '担当者名を表示
    If setVal("viewGant_Assignor") = True Then
      startColumn = WBS_Option.日付セル検索(Range(setVal("cell_PlanEnd") & line) + 1)
      endColumn = WBS_Option.日付セル検索(Range(setVal("cell_PlanEnd") & line) + 3)

      With Range(startColumn & line & ":" & endColumn & line)
        Set ProcessShape = ActiveSheet.Shapes.AddShape(Type:=msoShapeRoundedRectangle, Left:=.Left + 10, top:=.top, Width:=.Width + 10, Height:=10)
        
        With ProcessShape
          .Name = "担当者_" & line
          .Fill.ForeColor.RGB = RGB(255, 255, 255)
          .Fill.Transparency = 0
          '.TextFrame.Characters.Text = Range(setVal("cell_Assign") & line)
          .TextFrame.Characters.Font.Size = 9
          .TextFrame2.TextRange.Font.Bold = msoTrue
          .TextFrame2.WordWrap = msoFalse
          .TextFrame.HorizontalOverflow = xlOartHorizontalOverflowOverflow
          .TextFrame.VerticalOverflow = xlOartHorizontalOverflowOverflow
          .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignLeft
          .TextFrame2.VerticalAnchor = msoAnchorMiddle
          .TextFrame2.TextRange.Font.NameComplexScript = "メイリオ"
          .TextFrame2.TextRange.Font.NameFarEast = "メイリオ"
          .TextFrame2.TextRange.Font.Name = "メイリオ"
          .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
          .TextFrame2.TextRange.Font.Size = 9
        End With
      End With
      Set ProcessShape = Nothing
      ActiveSheet.Shapes.Range(Array("担当者_" & line)).Select
      Selection.Formula = "=" & Range(setVal("cell_Assign") & line).Address(False, False)
    End If
  End If
End Function


'**************************************************************************************************
' * 実績線設定
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function 実績線設定(line As Long)

  Dim rngStart As Range, rngEnd As Range
  Dim BX As Single, BY As Single, EX As Single, EY As Single
  Dim lColorValue As Long, Red As Long, Green As Long, Blue As Long
'    lColorValue = setSheet.Range(setVal("cell_ProgressEnd") & line).Interior.Color
  
  startColumn = WBS_Option.日付セル検索(Range(setVal("cell_AchievementStart") & line))
  
  Call Library.showDebugForm("実績線設定", Library.TEXTJOIN("", True, Range(setVal("cell_Info") & line & ":" & setVal("cell_TaskAreaEnd") & line)))
  Call Library.showDebugForm("実績線設定", "　開始日:" & Range(setVal("cell_AchievementStart") & line))
  Call Library.showDebugForm("実績線設定", "　終了日:" & Range(setVal("cell_AchievementEnd") & line))
  Call Library.showDebugForm("実績線設定", "　進捗　:" & Range(setVal("cell_Progress") & line))
  
  
  If Range(setVal("cell_AchievementEnd") & line) = "" Then
    endColumn = WBS_Option.日付セル検索(Date)
    
  ElseIf Range(setVal("cell_Progress") & line) = 100 Then
    If Range(setVal("cell_PlanEnd") & line) < Range(setVal("cell_AchievementEnd") & line) Then
      endColumn = WBS_Option.日付セル検索(Range(setVal("cell_AchievementEnd") & line))
    Else
      endColumn = WBS_Option.日付セル検索(Range(setVal("cell_PlanEnd") & line))
    End If
  Else
    endColumn = WBS_Option.日付セル検索(Range(setVal("cell_AchievementEnd") & line))
  End If
  
  
  Call Library.getRGB(setVal("lineColor_Achievement"), Red, Green, Blue)
  

  'Shapeを配置するための基準となるセル
  Set rngStart = Range(startColumn & line)
  Set rngEnd = Range(endColumn & line)
  
  'セルのLeft、Top、Widthプロパティを利用して位置決め
  BX = rngStart.Left
  BY = rngStart.top + (rngStart.Height / 2)
  EX = rngEnd.Left + rngEnd.Width
  EY = rngEnd.top + (rngEnd.Height / 2)
  
  
  With ActiveSheet.Shapes.AddLine(BX, BY, EX, EY).line
    If Range(setVal("cell_Assign") & line) = "工程" Or Range(setVal("cell_Assign") & line) = "工程" Then
      ActiveSheet.Shapes.Range(Array("工程_" & line)).Select
      Selection.ShapeRange.ZOrder msoBringToFront
      .Weight = 4
    Else
      .Weight = 4
    End If
  .Style = msoLineSolid
  .ForeColor.RGB = RGB(Red, Green, Blue)

 End With

End Function


'**************************************************************************************************
' * イナズマ線設定
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function イナズマ線設定(line As Long)

  Dim rngStart As Range, rngEnd As Range, rngBase As Range, rngChkDay As Range
  Dim BX As Single, BY As Single, EX As Single, EY As Single
  Dim startColumn As String, endColumn As String, baseColumn As String, chkDayColumn As String
  Dim progress As Long, lateOrEarly As Double
  Dim extensionDay As Integer
  Dim chkDay As Date
  Dim Red As Long, Green As Long, Blue As Long
  
  If Not (setVal("startDay") <= setVal("baseDay") And setVal("baseDay") <= setVal("endDay")) Then
    If setVal("setLightning") = True Then
      Call Library.showNotice(50)
      setVal("setLightning") = False
      Range("setLightning") = False
    End If
    Exit Function
    
  End If
  
  baseColumn = WBS_Option.日付セル検索(setVal("baseDay"))
  
  'イナズマ線の色取得
  Call Library.getRGB(setVal("lineColor_Lightning"), Red, Green, Blue)

  If Range(setVal("cell_PlanStart") & line) <> "" Then
    startColumn = WBS_Option.日付セル検索(Range(setVal("cell_PlanStart") & line))
  Else
    startColumn = baseColumn
  End If
  
  If Range(setVal("cell_AchievementEnd") & line) <> "" Then
    endColumn = WBS_Option.日付セル検索(Range(setVal("cell_AchievementEnd") & line))
  
  ElseIf Range(setVal("cell_PlanEnd") & line) <> "" Then
    endColumn = WBS_Option.日付セル検索(Range(setVal("cell_PlanEnd") & line))
  Else
    endColumn = baseColumn
  End If
    
  'Shapeを配置するための基準となるセル
  Set rngStart = Range(startColumn & line)
  Set rngEnd = Range(endColumn & line)
  Set rngBase = Range(baseColumn & line)
  
  'イナズマ線を引かない場合は、基準日のみ引く
  If setVal("setLightning") = False Then
    BX = rngBase.Left + rngBase.Width
    BY = rngBase.top
    EX = rngBase.Left + rngBase.Width
    EY = rngBase.top + rngBase.Height
    
    With ActiveSheet.Shapes.AddLine(BX, BY, EX, EY).line
      .Weight = 2
      .Style = msoLineSolid
      .ForeColor.RGB = RGB(Red, Green, Blue)
    End With
    Exit Function
  End If
  
  '遅早工数の値
  If Range(setVal("cell_LateOrEarly") & line) = 0 Or Range(setVal("cell_LateOrEarly") & line) = "" Then
    BX = rngBase.Left + rngBase.Width
    BY = rngBase.top
    EX = rngBase.Left + rngBase.Width
    EY = rngBase.top + rngBase.Height
    
    With ActiveSheet.Shapes.AddLine(BX, BY, EX, EY).line
      .Weight = 2
      .Style = msoLineSolid
      .ForeColor.RGB = RGB(Red, Green, Blue)
    End With
  Else
    chkDay = WBS_Option.イナズマ線用日付計算(setVal("baseDay"), Range(setVal("cell_LateOrEarly") & line))
    chkDayColumn = WBS_Option.日付セル検索(chkDay)
    
    Set rngChkDay = Range(chkDayColumn & line)
    
    BX = rngBase.Left + rngBase.Width
    BY = rngBase.top
    EX = rngChkDay.Left + rngChkDay.Width
    EY = rngBase.top + (rngBase.Height / 2)
    
    With ActiveSheet.Shapes.AddLine(BX, BY, EX, EY).line
      .Weight = 2
      .Style = msoLineSolid
      .ForeColor.RGB = RGB(Red, Green, Blue)
    End With

    BX = rngChkDay.Left + rngChkDay.Width
    BY = rngBase.top + (rngBase.Height / 2)
    EX = rngBase.Left + rngBase.Width
    EY = rngBase.top + rngBase.Height
    
    With ActiveSheet.Shapes.AddLine(BX, BY, EX, EY).line
      .Weight = 2
      .Style = msoLineSolid
      .ForeColor.RGB = RGB(Red, Green, Blue)
    End With
  End If
  
End Function


'**************************************************************************************************
' * タイムラインに追加
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function タイムラインに追加(line As Long)
  Dim ShapeTopStart As Long, count As Long
  Dim shp As Shape
  Dim rng As Range

'  On Error GoTo catchError
  
  Call init.setting
  
  startColumn = WBS_Option.日付セル検索(Range(setVal("cell_PlanStart") & line))
  endColumn = WBS_Option.日付セル検索(Range(setVal("cell_PlanEnd") & line))


  If Library.chkShapeName("タイムライン_" & line) Then
    ActiveSheet.Shapes("タイムライン_" & line).Delete
  End If
  
  On Error Resume Next
  count = 0
  Range(startColumn & "5:" & endColumn & 5).Select
  For Each shp In ActiveSheet.Shapes
    Set rng = Range(shp.TopLeftCell, shp.BottomRightCell)
    If Not (Intersect(rng, Selection) Is Nothing) Then
      count = count + 1
    End If
  Next
  If count <> 0 Then
    ShapeTopStart = 10 * count
  Else
    ShapeTopStart = 0
  End If
  On Error GoTo 0



  'タイムライン行の幅を広げる
  If count >= 3 Then
    Rows("5:5").RowHeight = Rows("5:5").RowHeight + 10
  End If
  
  
  With Range(startColumn & "5:" & endColumn & 5)
    Set ProcessShape = ActiveSheet.Shapes.AddShape(Type:=msoShapeRoundedRectangle, Left:=.Left, top:=.top + ShapeTopStart, Width:=.Width, Height:=10)
    
    With ProcessShape
      .Name = "タイムライン_" & line
      .Fill.ForeColor.RGB = RGB(102, 102, 255)
      .Fill.Transparency = 0.6
      .TextFrame2.WordWrap = msoFalse
      .TextFrame.HorizontalOverflow = xlOartHorizontalOverflowOverflow
      .TextFrame.VerticalOverflow = xlOartHorizontalOverflowOverflow
      .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
      .TextFrame2.VerticalAnchor = msoAnchorMiddle
      .TextFrame2.TextRange.Font.NameComplexScript = "メイリオ"
      .TextFrame2.TextRange.Font.NameFarEast = "メイリオ"
      .TextFrame2.TextRange.Font.Name = "メイリオ"
      .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
      .TextFrame2.TextRange.Font.Size = 9
    End With
  End With
  ActiveSheet.Shapes.Range(Array("タイムライン_" & line)).Select
  Selection.Formula = "=" & Range(setVal("cell_TaskArea") & line).Address(False, False)
  
  If Range(setVal("cell_Info") & line) = "" Then
    Range(setVal("cell_Info") & line) = setVal("TaskInfoStr_TimeLine")
  ElseIf Range(setVal("cell_Info") & line) Like "*" & setVal("TaskInfoStr_TimeLine") & "*" Then
  Else
    Range(setVal("cell_Info") & line) = Range(setVal("cell_Info") & line) & "," & setVal("TaskInfoStr_TimeLine")
  End If



  Exit Function
'エラー発生時=====================================================================================
catchError:
  Call Library.showNotice(Err.Number, Err.Description, True)
End Function


'**************************************************************************************************
' * センター
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function センター()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  Dim baseColumn As String

  
  On Error GoTo catchError

  If setVal("startDay") >= setVal("baseDay") - 10 Then
  
  Else
    baseColumn = WBS_Option.日付セル検索(setVal("baseDay") - 10)
    Application.Goto Reference:=Range(baseColumn & 6), Scroll:=True
  End If
  


  Exit Function
'エラー発生時=====================================================================================
catchError:

End Function





'**************************************************************************************************
' * ガントチャート選択
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function beforeChangeShapes()

  Call Library.startScript
  ActiveSheet.Shapes.Range(Array(Application.Caller)).Select
  changeShapesName = Application.Caller
  
'  Call Library.setArrayPush(selectShapesName, Application.Caller)
  
  With ActiveSheet.Shapes(changeShapesName)
    .ScaleWidth 0.9792388451, msoFalse, msoScaleFromBottomRight
    .ScaleWidth 0.9792388451, msoFalse, msoScaleFromTopLeft
  End With
  
  Call Library.endScript
End Function


'**************************************************************************************************
' * Chartから予定日を操作
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function changeShapes()
  Dim rng As Range
  Dim newStartDay As Date, newEndDay As Date
  Dim HollydayName As String
  
  Call Library.showDebugForm("Chartから予定日を操作", "処理開始")
  
  
  With ActiveSheet.Shapes(changeShapesName)
    .ScaleWidth 0.9792388451, msoFalse, msoScaleFromBottomRight
    .ScaleWidth 0.9792388451, msoFalse, msoScaleFromTopLeft
  End With

  With ActiveSheet.Shapes(changeShapesName)
    Set rng = Range(.TopLeftCell, .BottomRightCell)
  End With
  If rng.Address(False, False) Like "*:*" Then
    tmp = Split(rng.Address(False, False), ":")
    changeShapesName = Replace(changeShapesName, "タスク_", "")
    mainSheet.Range(setVal("cell_PlanStart") & changeShapesName) = Range(getColumnName(Range(tmp(0)).Column) & 4)
    mainSheet.Range(setVal("cell_PlanEnd") & changeShapesName) = Range(getColumnName(Range(tmp(1)).Column) & 4)
  Else
    tmp = rng.Address(False, False)
    changeShapesName = Replace(changeShapesName, "タスク_", "")
    mainSheet.Range(setVal("cell_PlanStart") & changeShapesName) = Range(getColumnName(Range(tmp).Column) & 4)
    mainSheet.Range(setVal("cell_PlanEnd") & changeShapesName) = Range(getColumnName(Range(tmp).Column) & 4)
  End If
  
  '先行タスクの終了日+1を開始日に設定
  newStartDay = mainSheet.Range(setVal("cell_PlanStart") & changeShapesName)
  Call init.chkHollyday(newStartDay, HollydayName)
  Do While HollydayName <> ""
    newStartDay = newStartDay - 1
    Call init.chkHollyday(newStartDay, HollydayName)
  Loop
  Range(setVal("cell_PlanStart") & changeShapesName) = newStartDay
  
  '終了日を再設定
  newEndDay = mainSheet.Range(setVal("cell_PlanEnd") & changeShapesName)
  Call init.chkHollyday(newEndDay, HollydayName)
  Do While HollydayName <> ""
    newEndDay = newEndDay + 1
    Call init.chkHollyday(newEndDay, HollydayName)
  Loop
  Range(setVal("cell_PlanEnd") & changeShapesName) = newEndDay
  
  If ActiveSheet.Name = TeamsPlannerSheetName Then
    If Range(setVal("cell_Info") & changeShapesName) = "" Then
      Range(setVal("cell_Info") & changeShapesName) = setVal("TaskInfoStr_Change")
    ElseIf Range(setVal("cell_Info") & changeShapesName) Like "*" & setVal("TaskInfoStr_Change") & "*" Then
    Else
      Range(setVal("cell_Info") & changeShapesName) = Range(setVal("cell_Info") & changeShapesName) & "," & setVal("TaskInfoStr_Change")
    End If
  End If
  
  ActiveSheet.Shapes("タスク_" & changeShapesName).Delete
  If setVal("viewGant_Assignor") = True Then
    ActiveSheet.Shapes("担当者_" & changeShapesName).Delete
  End If
  
  Call 計画線設定(CLng(changeShapesName))
  
  changeShapesName = ""

  Call Library.showDebugForm("Chartから予定日を操作", "処理終了")
End Function


'**************************************************************************************************
' * タスクのリンク設定
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function タスクのリンク設定(line As Long)
  Dim startTask As String, thisTask As String
  Dim interval As Long
  Dim tmpLine As Variant
  
'  On Error GoTo catchError


  If Range(setVal("cell_Task") & line) = "" Then
    Exit Function
  Else
    For Each tmpLine In Split(Range(setVal("cell_Task") & line), ",")
      startTask = "タスク_" & tmpLine
      thisTask = "タスク_" & line
    
      'カギ線コネクタ生成
      ActiveSheet.Shapes.AddConnector(msoConnectorElbow, 1153.2352755906, 9.7059055118, 1206.1764566929, 30).Select
      Selection.ShapeRange.line.EndArrowheadStyle = msoArrowheadTriangle
      Selection.ShapeRange.ConnectorFormat.BeginConnect ActiveSheet.Shapes(startTask), 4
      Selection.ShapeRange.ConnectorFormat.EndConnect ActiveSheet.Shapes(thisTask), 2
      Selection.Name = "先行タスク設定_" & line
      
      interval = DateDiff("d", Range(setVal("cell_PlanEnd") & tmpLine), Range(setVal("cell_PlanStart") & line))
      If interval < 1 Then
        interval = interval * -1
      End If
      
     If interval = 0 Then
     ElseIf interval < 2 Then
        Selection.ShapeRange.Flip msoFlipHorizontal
     End If
      
      With Selection.ShapeRange.line
        .Visible = msoTrue
        .Weight = 1.5
        .ForeColor.RGB = RGB(166, 166, 166)
      End With
    Next
  End If


  Exit Function
'エラー発生時=====================================================================================
catchError:
  Call Library.showNotice(Err.Number, Err.Description, True)
End Function








