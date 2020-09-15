Attribute VB_Name = "Chart"



'**************************************************************************************************
' * �K���g�`���[�g����
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function �K���g�`���[�g����()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  Dim startColumn As String, endColumn As String
  
  Call WBS_Option.�I���V�[�g�m�F
  
  
  Rows("6:" & Rows.count).EntireRow.Hidden = False
  
  Call �K���g�`���[�g�폜
  endLine = Cells(Rows.count, 2).End(xlUp).row
  
  For line = 6 To endLine
    '�v�������------------------------------------
    If Not (mainSheet.Range(setVal("cell_PlanStart") & line) = "" Or mainSheet.Range(setVal("cell_PlanEnd") & line) = "") Then
      Call �v����ݒ�(line)
    End If

    '���ѐ�����------------------------------------
    If mainSheet.Range(setVal("cell_Progress") & line) >= 0 Then
      Call ���ѐ��ݒ�(line)
    End If
    
    '�^�C�����C���ւ̒ǉ�------------------------------------
    If (mainSheet.Range(setVal("cell_Info") & line) = setVal("TaskInfoStr_TimeLine")) Then
      Call �^�C�����C���ɒǉ�(line)
    End If
    
    '�C�i�Y�}������------------------------------
    Call �C�i�Y�}���ݒ�(line)
    
    '�i����100%�Ȃ��\��------------------------------------
    If setVal("setDispProgress100") = True And mainSheet.Range(setVal("cell_Progress") & line) = 100 Then
      Rows(line & ":" & line).EntireRow.Hidden = True
      
    End If

  Next
  For line = 6 To endLine
    Call �^�X�N�̃����N�ݒ�(line)
  Next

  If ActiveSheet.Name = mainSheetName Then
    Call WBS_Option.�����̒S���ҍs���\��
  End If

End Function


'**************************************************************************************************
' * �K���g�`���[�g�폜
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function �K���g�`���[�g�폜()
  Dim shp As Shape
  Dim rng As Range
  
  On Error Resume Next
  
  Set rng = Range(Cells(5, Library.getColumnNo(setVal("calendarStartCol"))), Cells(Rows.count, Columns.count))
  
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
' * �v����ݒ�
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function �v����ݒ�(line As Long)

  Dim rngStart As Range, rngEnd As Range
  Dim BX As Single, BY As Single, EX As Single, EY As Single
  Dim lColorValue As Long, Red As Long, Green As Long, Blue As Long
  Dim ProcessShape As Shape
  
  startColumn = WBS_Option.���t�Z������(Range(setVal("cell_PlanStart") & line))
  endColumn = WBS_Option.���t�Z������(Range(setVal("cell_PlanEnd") & line))
  
'  'Shape��z�u���邽�߂̊�ƂȂ�Z��
'  Set rngStart = mainSheet.Range(startColumn & line)
'  Set rngEnd = mainSheet.Range(endColumn & line)
'
'  '�Z����Left�ATop�AWidth�v���p�e�B�𗘗p���Ĉʒu����
'  BX = rngStart.Left
'  BY = rngStart.top + (rngStart.Height / 2)
'  EX = rngEnd.Left + rngEnd.Width
'  EY = rngEnd.top + (rngEnd.Height / 2)
  
  '�S���ҕʂ̐F�ݒ�------------------------------
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

  If Range(setVal("cell_Assign") & line) = "�H��" Or Range(setVal("cell_Assign") & line) = "�H��" Then
    With Range(startColumn & line & ":" & endColumn & line)
      Set ProcessShape = ActiveSheet.Shapes.AddShape(Type:=msoShapePentagon, Left:=.Left, top:=.top, Width:=.Width, Height:=.Height)
      
      With ProcessShape
        .Name = "�^�X�N_" & line
        .Fill.ForeColor.RGB = RGB(Red, Green, Blue)
        .Fill.Transparency = 0.6
'        .TextFrame.Characters.Text = Range(setVal("cell_TaskArea") & line)
'        .TextFrame.Characters.Font.Size = 12
        .TextFrame2.WordWrap = msoFalse
        .TextFrame.HorizontalOverflow = xlOartHorizontalOverflowOverflow
        .TextFrame.VerticalOverflow = xlOartHorizontalOverflowOverflow
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .TextFrame2.TextRange.Font.NameComplexScript = "���C���I"
        .TextFrame2.TextRange.Font.NameFarEast = "���C���I"
        .TextFrame2.TextRange.Font.Name = "���C���I"
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
        .TextFrame2.TextRange.Font.Size = 12
        .TextFrame2.TextRange.Font.Bold = msoTrue
      End With
    End With
    Set ProcessShape = Nothing
    ActiveSheet.Shapes.Range(Array("�^�X�N_" & line)).Select
    Selection.Formula = "=" & Range(setVal("cell_TaskArea") & line).Address(False, False)
    Selection.ShapeRange.TextFrame2.TextRange.Font.Size = 12
    Selection.ShapeRange.TextFrame2.TextRange.Font.Bold = msoTrue
  
  Else
    With Range(startColumn & line & ":" & endColumn & line)
      'Set ProcessShape = ActiveSheet.Shapes.AddShape(Type:=msoShapeRectangle, Left:=.Left, top:=.top + 5, Width:=.Width, Height:=10)
      Set ProcessShape = ActiveSheet.Shapes.AddShape(Type:=msoShapeRectangle, Left:=.Left, top:=.top, Width:=.Width, Height:=.Height)
      
      With ProcessShape
        .Name = "�^�X�N_" & line
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
        .TextFrame2.TextRange.Font.NameComplexScript = "���C���I"
        .TextFrame2.TextRange.Font.NameFarEast = "���C���I"
        .TextFrame2.TextRange.Font.Name = "���C���I"
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
        .TextFrame2.TextRange.Font.Size = 9
    
        If setVal("viewGant_TaskName") = True Then
          ActiveSheet.Shapes.Range(Array("�^�X�N_" & line)).Select
          Selection.Formula = "=" & Range(setVal("cell_TaskArea") & line).Address(False, False)
        End If
        
        .OnAction = "beforeChangeShapes"
      End With
    End With
    Set ProcessShape = Nothing

    '�S���Җ���\��
    If setVal("viewGant_Assignor") = True Then
      startColumn = WBS_Option.���t�Z������(Range(setVal("cell_PlanEnd") & line) + 1)
      endColumn = WBS_Option.���t�Z������(Range(setVal("cell_PlanEnd") & line) + 3)

      With Range(startColumn & line & ":" & endColumn & line)
        Set ProcessShape = ActiveSheet.Shapes.AddShape(Type:=msoShapeRoundedRectangle, Left:=.Left + 10, top:=.top, Width:=.Width + 10, Height:=10)
        
        With ProcessShape
          .Name = "�S����_" & line
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
          .TextFrame2.TextRange.Font.NameComplexScript = "���C���I"
          .TextFrame2.TextRange.Font.NameFarEast = "���C���I"
          .TextFrame2.TextRange.Font.Name = "���C���I"
          .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
          .TextFrame2.TextRange.Font.Size = 9
        End With
      End With
      Set ProcessShape = Nothing
      ActiveSheet.Shapes.Range(Array("�S����_" & line)).Select
      Selection.Formula = "=" & Range(setVal("cell_Assign") & line).Address(False, False)
    End If
  End If
End Function


'**************************************************************************************************
' * ���ѐ��ݒ�
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function ���ѐ��ݒ�(line As Long)

  Dim rngStart As Range, rngEnd As Range
  Dim BX As Single, BY As Single, EX As Single, EY As Single
  Dim lColorValue As Long, Red As Long, Green As Long, Blue As Long
  Dim ProcessShape As Shape
  Dim shapesWith As Long
  
'    lColorValue = setSheet.Range(setVal("cell_ProgressEnd") & line).Interior.Color
  
'  Call Library.showDebugForm("���ѐ��ݒ�", Range(setVal("cell_TaskArea") & line))
'  Call Library.showDebugForm("���ѐ��ݒ�", "�@�J�n��:" & Range(setVal("cell_AchievementStart") & line))
'  Call Library.showDebugForm("���ѐ��ݒ�", "�@�I����:" & Range(setVal("cell_AchievementEnd") & line))
'  Call Library.showDebugForm("���ѐ��ݒ�", "�@�i���@:" & Range(setVal("cell_Progress") & line))
  
  If Range(setVal("cell_AchievementStart") & line) = "" Then
    startColumn = WBS_Option.���t�Z������(Range(setVal("cell_PlanStart") & line))
  Else
    startColumn = WBS_Option.���t�Z������(Range(setVal("cell_AchievementStart") & line))
  End If
  
  If Range(setVal("cell_AchievementEnd") & line) = "" Then
    endColumn = WBS_Option.���t�Z������(Range(setVal("cell_PlanEnd") & line))
  
  '�i����100%�̂Ƃ�
  ElseIf Range(setVal("cell_Progress") & line) = 100 Then
    If Range(setVal("cell_AchievementEnd") & line) < Range(setVal("cell_PlanEnd") & line) Then
      endColumn = WBS_Option.���t�Z������(Range(setVal("cell_PlanEnd") & line))
    Else
      endColumn = WBS_Option.���t�Z������(Range(setVal("cell_AchievementEnd") & line))
    End If
  
  Else
    endColumn = WBS_Option.���t�Z������(Range(setVal("cell_AchievementEnd") & line))
  End If

  
  
  Call Library.getRGB(setVal("lineColor_Achievement"), Red, Green, Blue)

  
  With Range(startColumn & line & ":" & endColumn & line)
    .Select
    
    If Range(setVal("cell_Progress") & line) = "" Or Range(setVal("cell_Progress") & line) = 0 Then
      shapesWith = 0
    Else
      shapesWith = .Width * (Range(setVal("cell_Progress") & line) / 100)
    End If
    
    If Range(setVal("cell_Assign") & line) = "�H��" Or Range(setVal("cell_Assign") & line) = "�H��" Then
      Set ProcessShape = ActiveSheet.Shapes.AddShape(Type:=msoShapePentagon, Left:=.Left, top:=.top + 5, Width:=shapesWith, Height:=10)
    Else
      Set ProcessShape = ActiveSheet.Shapes.AddShape(Type:=msoShapeRectangle, Left:=.Left, top:=.top + 5, Width:=shapesWith, Height:=10)
    End If
    
    With ProcessShape
      .Name = "����_" & line
      .Fill.ForeColor.RGB = RGB(Red, Green, Blue)
      .Fill.Transparency = 0.6
    End With
  End With
  Set ProcessShape = Nothing
    
    
    


End Function


'**************************************************************************************************
' * �C�i�Y�}���ݒ�
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function �C�i�Y�}���ݒ�(line As Long)

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
  
  '�C�i�Y�}���̐F�擾
  Call Library.getRGB(setVal("lineColor_Lightning"), Red, Green, Blue)
  
  baseColumn = WBS_Option.���t�Z������(setVal("baseDay"))
  
  '�^�C�����C����Ɉ���
  If line = 6 Then
    Set rngBase = Range(baseColumn & 5)
    
    '�����R�l�N�^����
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, rngBase.Left + 10, rngBase.top, rngBase.Left + 10, rngBase.top + rngBase.Height).Select
    With Selection
      .Name = "�C�i�Y�}��B_5"
      .ShapeRange.line.Weight = 3
      .ShapeRange.line.ForeColor.RGB = RGB(Red, Green, Blue)
      .ShapeRange.line.Transparency = 0.6
    End With

    Set ProcessShape = Nothing
  End If
  
  Set rngBase = Range(baseColumn & line)
  
  
  
  
  '�C�i�Y�}���������Ȃ��ꍇ�́A����݈̂���
  If setVal("setLightning") = False Or Range(setVal("cell_Progress") & line) = "" Or Range(setVal("cell_LateOrEarly") & line) = 0 Then
    
    '�����R�l�N�^����
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, rngBase.Left + 10, rngBase.top, rngBase.Left + 10, rngBase.top + rngBase.Height).Select
    With Selection
      .Name = "�C�i�Y�}��B_" & line
      .ShapeRange.line.Weight = 3
      .ShapeRange.line.ForeColor.RGB = RGB(Red, Green, Blue)
      .ShapeRange.line.Transparency = 0.6
    End With

    Set ProcessShape = Nothing
    Exit Function
  
  '�i����0%�ȏ�̏ꍇ�́A�C�i�Y�}��������
  ElseIf Range(setVal("cell_Progress") & line) >= 0 Then
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, rngBase.Left + 10, rngBase.top, rngBase.Left + 10, rngBase.top + rngBase.Height).Select
    With Selection
      .Name = "�C�i�Y�}��S_" & line
      .ShapeRange.line.Weight = 3
      .ShapeRange.line.ForeColor.RGB = RGB(Red, Green, Blue)
      .ShapeRange.line.Transparency = 0.6
    End With
    Selection.ShapeRange.ConnectorFormat.EndConnect ActiveSheet.Shapes("����_" & line), 4
  
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, rngBase.Left + 10, rngBase.top, rngBase.Left + 10, rngBase.top + rngBase.Height).Select
    With Selection
      .Name = "�C�i�Y�}��S_" & line
      .ShapeRange.line.Weight = 3
      .ShapeRange.line.ForeColor.RGB = RGB(Red, Green, Blue)
      .ShapeRange.line.Transparency = 0.6
    End With
    Selection.ShapeRange.ConnectorFormat.BeginConnect ActiveSheet.Shapes("����_" & line), 4
    
    
'
'      startTask = "�^�X�N_" & tmpLine
'      thisTask = "�^�X�N_" & line
'
'    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 1153.2352755906, 9.7059055118, 1206.1764566929, 30).Select
'    Selection.ShapeRange.line.EndArrowheadStyle = msoArrowheadTriangle
'
'    Selection.ShapeRange.ConnectorFormat.EndConnect ActiveSheet.Shapes(thisTask), 2
'    Selection.Name = "�C�i�Y�}��_" & line
  
      
  End If
Exit Function





  If Range(setVal("cell_PlanStart") & line) <> "" Then
    startColumn = WBS_Option.���t�Z������(Range(setVal("cell_PlanStart") & line))
  Else
    startColumn = baseColumn
  End If
  
  If Range(setVal("cell_AchievementEnd") & line) <> "" Then
    endColumn = WBS_Option.���t�Z������(Range(setVal("cell_AchievementEnd") & line))
  
  ElseIf Range(setVal("cell_PlanEnd") & line) <> "" Then
    endColumn = WBS_Option.���t�Z������(Range(setVal("cell_PlanEnd") & line))
  Else
    endColumn = baseColumn
  End If
    
  'Shape��z�u���邽�߂̊�ƂȂ�Z��
  Set rngStart = Range(startColumn & line)
  Set rngEnd = Range(endColumn & line)

  
  '�x���H���̒l
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
    chkDay = WBS_Option.�C�i�Y�}���p���t�v�Z(setVal("baseDay"), Range(setVal("cell_LateOrEarly") & line))
    chkDayColumn = WBS_Option.���t�Z������(chkDay)
    
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
' * �^�C�����C���ɒǉ�
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function �^�C�����C���ɒǉ�(line As Long)
  Dim ShapeTopStart As Long, count As Long
  Dim shp As Shape
  Dim rng As Range

'  On Error GoTo catchError
  
  Call init.setting
  
  startColumn = WBS_Option.���t�Z������(Range(setVal("cell_PlanStart") & line))
  endColumn = WBS_Option.���t�Z������(Range(setVal("cell_PlanEnd") & line))


  If Library.chkShapeName("�^�C�����C��_" & line) Then
    ActiveSheet.Shapes("�^�C�����C��_" & line).Delete
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



  '�^�C�����C���s�̕����L����
  If count >= 3 Then
    Rows("5:5").RowHeight = Rows("5:5").RowHeight + 10
  End If
  
  
  With Range(startColumn & "5:" & endColumn & 5)
    Set ProcessShape = ActiveSheet.Shapes.AddShape(Type:=msoShapeRoundedRectangle, Left:=.Left, top:=.top + ShapeTopStart, Width:=.Width, Height:=10)
    
    With ProcessShape
      .Name = "�^�C�����C��_" & line
      .Fill.ForeColor.RGB = RGB(102, 102, 255)
      .Fill.Transparency = 0.6
      .TextFrame2.WordWrap = msoFalse
      .TextFrame.HorizontalOverflow = xlOartHorizontalOverflowOverflow
      .TextFrame.VerticalOverflow = xlOartHorizontalOverflowOverflow
      .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
      .TextFrame2.VerticalAnchor = msoAnchorMiddle
      .TextFrame2.TextRange.Font.NameComplexScript = "���C���I"
      .TextFrame2.TextRange.Font.NameFarEast = "���C���I"
      .TextFrame2.TextRange.Font.Name = "���C���I"
      .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
      .TextFrame2.TextRange.Font.Size = 9
    End With
  End With
  ActiveSheet.Shapes.Range(Array("�^�C�����C��_" & line)).Select
  Selection.Formula = "=" & Range(setVal("cell_TaskArea") & line).Address(False, False)
  
  If Range(setVal("cell_Info") & line) = "" Then
    Range(setVal("cell_Info") & line) = setVal("TaskInfoStr_TimeLine")
  ElseIf Range(setVal("cell_Info") & line) Like "*" & setVal("TaskInfoStr_TimeLine") & "*" Then
  Else
    Range(setVal("cell_Info") & line) = Range(setVal("cell_Info") & line) & "," & setVal("TaskInfoStr_TimeLine")
  End If



  Exit Function
'�G���[������=====================================================================================
catchError:
  Call Library.showNotice(Err.Number, Err.Description, True)
End Function


'**************************************************************************************************
' * �Z���^�[
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function �Z���^�[()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  Dim baseDate As Date
  Dim baseColumn As String
  
  
'  On Error GoTo catchError

  If setVal("startDay") >= setVal("baseDay") - 10 Then
    baseDate = setVal("startDay")
  Else
    baseDate = setVal("baseDay") - 10
    
  End If
  
  baseColumn = WBS_Option.���t�Z������(baseDate)
  Application.Goto Reference:=Range(baseColumn & 6), Scroll:=True


  Exit Function
'�G���[������=====================================================================================
catchError:
  Call Library.showNotice(Err.Number, Err.Description, True)
End Function





'**************************************************************************************************
' * �K���g�`���[�g�I��
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
' * Chart����\����𑀍�
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function changeShapes()
  Dim rng As Range
  Dim newStartDay As Date, newEndDay As Date
  Dim HollydayName As String
  
  Call Library.showDebugForm("Chart����\����𑀍�", "�����J�n")
  
  
  With ActiveSheet.Shapes(changeShapesName)
    .ScaleWidth 0.9792388451, msoFalse, msoScaleFromBottomRight
    .ScaleWidth 0.9792388451, msoFalse, msoScaleFromTopLeft
  End With

  With ActiveSheet.Shapes(changeShapesName)
    Set rng = Range(.TopLeftCell, .BottomRightCell)
  End With
  If rng.Address(False, False) Like "*:*" Then
    tmp = Split(rng.Address(False, False), ":")
    changeShapesName = Replace(changeShapesName, "�^�X�N_", "")
    mainSheet.Range(setVal("cell_PlanStart") & changeShapesName) = Range(getColumnName(Range(tmp(0)).Column) & 4)
    mainSheet.Range(setVal("cell_PlanEnd") & changeShapesName) = Range(getColumnName(Range(tmp(1)).Column) & 4)
  Else
    tmp = rng.Address(False, False)
    changeShapesName = Replace(changeShapesName, "�^�X�N_", "")
    mainSheet.Range(setVal("cell_PlanStart") & changeShapesName) = Range(getColumnName(Range(tmp).Column) & 4)
    mainSheet.Range(setVal("cell_PlanEnd") & changeShapesName) = Range(getColumnName(Range(tmp).Column) & 4)
  End If
  
  '��s�^�X�N�̏I����+1���J�n���ɐݒ�
  newStartDay = mainSheet.Range(setVal("cell_PlanStart") & changeShapesName)
  Call init.chkHollyday(newStartDay, HollydayName)
  Do While HollydayName <> ""
    newStartDay = newStartDay - 1
    Call init.chkHollyday(newStartDay, HollydayName)
  Loop
  Range(setVal("cell_PlanStart") & changeShapesName) = newStartDay
  
  '�I�������Đݒ�
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
  
  ActiveSheet.Shapes("�^�X�N_" & changeShapesName).Delete
  If Range(setVal("cell_Task") & CLng(changeShapesName)) <> "" Then
    ActiveSheet.Shapes("��s�^�X�N�ݒ�_" & CLng(changeShapesName)).Delete
  End If
  If setVal("viewGant_Assignor") = True Then
    ActiveSheet.Shapes("�S����_" & changeShapesName).Delete
  End If
  
  Call �v����ݒ�(CLng(changeShapesName))
  Call �^�X�N�̃����N�ݒ�(CLng(changeShapesName))
  
  changeShapesName = ""

  Call Library.showDebugForm("Chart����\����𑀍�", "�����I��")
End Function


'**************************************************************************************************
' * �^�X�N�̃����N�ݒ�
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function �^�X�N�̃����N�ݒ�(line As Long)
  Dim startTask As String, thisTask As String
  Dim interval As Long
  Dim tmpLine As Variant
  
'  On Error GoTo catchError


  If Range(setVal("cell_Task") & line) = "" Then
    Exit Function
  Else
    For Each tmpLine In Split(Range(setVal("cell_Task") & line), ",")
      startTask = "�^�X�N_" & tmpLine
      thisTask = "�^�X�N_" & line
    
      '�J�M���R�l�N�^����
      ActiveSheet.Shapes.AddConnector(msoConnectorElbow, 1153.2352755906, 9.7059055118, 1206.1764566929, 30).Select
      Selection.ShapeRange.line.EndArrowheadStyle = msoArrowheadTriangle
      Selection.ShapeRange.ConnectorFormat.BeginConnect ActiveSheet.Shapes(startTask), 4
      Selection.ShapeRange.ConnectorFormat.EndConnect ActiveSheet.Shapes(thisTask), 2
      Selection.Name = "��s�^�X�N�ݒ�_" & line
      
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
        .ForeColor.RGB = RGB(0, 0, 0)
      End With
    Next
  End If


  Exit Function
'�G���[������=====================================================================================
catchError:
  Call Library.showNotice(Err.Number, Err.Description, True)
End Function








