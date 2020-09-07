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
    
    'イナズマ線生成------------------------------
      Call イナズマ線設定(line)
    If setVal("setLightning") = True Then
    End If
    
  Next
  For line = 6 To endLine
    Call タスクのリンク設定(line)
  Next
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
  
  Set rng = Range(Cells(6, Library.getColumnNo(setVal("calendarStartCol"))), Cells(endLine, endColLine))
  
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
  If lColorValue <> 0 Then
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
        .TextFrame.Characters.Text = Range(setVal("cell_TaskArea") & line)
        .TextFrame.Characters.Font.Size = 12
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
      End With
    End With
    Set ProcessShape = Nothing
  
  Else
    With Range(startColumn & line & ":" & endColumn & line)
      Set ProcessShape = ActiveSheet.Shapes.AddShape(Type:=msoShapeRectangle, Left:=.Left, top:=.top + 5, Width:=.Width, Height:=10)
      
      With ProcessShape
        .Name = "タスク_" & line
        .Fill.ForeColor.RGB = RGB(Red, Green, Blue)
        .Fill.Transparency = 0.6
        
        If setVal("viewGant_TaskName") = True Then
          .TextFrame.Characters.Text = Range(setVal("cell_TaskArea") & line)
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
          .TextFrame.Characters.Text = Range(setVal("cell_Assign") & line)
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
        End With
      End With
      Set ProcessShape = Nothing
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
  
  Call Library.showDebugForm("実績線設定", Library.TEXTJOIN("", True, Range("C" & line & ":" & setVal("cell_TaskAreaEnd") & line)))
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
    Call Library.showNotice(50)
    setVal("setLightning") = False
    Range("setLightning") = False
'    Exit Function
  End If
  
  Call Library.showDebugForm("イナズマ線設定", Range("C" & line))
  
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
'  If IsEmpty(Range(setVal("cell_Progress") & line)) Then
'    progress = -1
'  Else
'    progress = Range(setVal("cell_Progress") & line).Value
'  End If
  
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
  ActiveSheet.Shapes.Range(Array(Application.Caller)).Select
  changeShapesName = Application.Caller
  
'  Call Library.setArrayPush(selectShapesName, Application.Caller)
  
  With ActiveSheet.Shapes(changeShapesName)
    .ScaleWidth 0.9792388451, msoFalse, msoScaleFromBottomRight
    .ScaleWidth 0.9792388451, msoFalse, msoScaleFromTopLeft
  End With
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
  
  Call init.setting

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
  
  Range("C" & changeShapesName) = "変"

  
  ActiveSheet.Shapes("タスク_" & changeShapesName).Delete
  If setVal("viewGant_Assignor") = True Then
    ActiveSheet.Shapes("担当者_" & changeShapesName).Delete
  End If
  
  Call 計画線設定(CLng(changeShapesName))
  
  changeShapesName = ""

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








