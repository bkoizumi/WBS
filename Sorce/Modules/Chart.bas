Attribute VB_Name = "Chart"



'**************************************************************************************************
' * �K���g�`���[�g����
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function �K���g�`���[�g����()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  Dim startColumn As String, endColumn As String
  
  Call init.setting
  Call �K���g�`���[�g�폜
  endLine = mainSheet.Cells(Rows.count, 2).End(xlUp).row
  
  For line = 6 To endLine
    '�v�������------------------------------------
    If Not (mainSheet.Range(setVal("cell_PlanStart") & line) = "" Or mainSheet.Range(setVal("cell_PlanEnd") & line) = "") Then
      Call �v����ݒ�(line)
    End If

    '���ѐ�����------------------------------------
    If Not (mainSheet.Range(setVal("cell_AchievementStart") & line) = "") Then
      Call ���ѐ��ݒ�(line)
    End If
    
    '�C�i�Y�}������------------------------------
      Call �C�i�Y�}���ݒ�(line)
    If setVal("setLightning") = True Then
    End If
  Next
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
  mainSheet.Select
  endLine = mainSheet.Cells(Rows.count, 1).End(xlUp).row
  endColLine = mainSheet.Cells(4, Columns.count).End(xlToLeft).Column
  
  Set rng = mainSheet.Range(Cells(6, Library.getColumnNo(setVal("calendarStartCol"))), Cells(endLine, endColLine))
  
  For Each shp In mainSheet.Shapes
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
  
  startColumn = WBS_Option.���t�Z������(mainSheet.Range(setVal("cell_PlanStart") & line))
  endColumn = WBS_Option.���t�Z������(mainSheet.Range(setVal("cell_PlanEnd") & line))
  
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
  If Range(setVal("cell_AssignA") & line) <> "" Then
    lColorValue = memberColor.item(Range(setVal("cell_AssignA") & line).Value)
  ElseIf Range(setVal("cell_AssignP") & line) <> "" Then
    lColorValue = memberColor.item(Range(setVal("cell_AssignP") & line).Value)
  End If
  If lColorValue <> 0 Then
    Call Library.getRGB(lColorValue, Red, Green, Blue)
  Else
    Call Library.getRGB(setVal("lineColor_Plan"), Red, Green, Blue)
  End If

  If mainSheet.Range(setVal("cell_AssignP") & line) = "�H��" Or mainSheet.Range(setVal("cell_AssignA") & line) = "�H��" Then
    With mainSheet.Range(startColumn & line & ":" & endColumn & line)
      Set ProcessShape = ActiveSheet.Shapes.AddShape(Type:=msoShapePentagon, Left:=.Left, top:=.top, Width:=.Width, Height:=.Height)
      
      With ProcessShape
        .Name = "�H��_" & line
        .Fill.ForeColor.RGB = RGB(Red, Green, Blue)
        .Fill.Transparency = 0.6
        .TextFrame.Characters.Text = mainSheet.Range("C" & line)
        .TextFrame.Characters.Font.Size = 12
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
      End With
    End With
    Set ProcessShape = Nothing
  
  Else
    With mainSheet.Range(startColumn & line & ":" & endColumn & line)
      Set ProcessShape = ActiveSheet.Shapes.AddShape(Type:=msoShapeRectangle, Left:=.Left, top:=.top + 5, Width:=.Width, Height:=10)
      
      With ProcessShape
        .Name = "�^�X�N_" & line
        .Fill.ForeColor.RGB = RGB(Red, Green, Blue)
        .Fill.Transparency = 0.6
        .OnAction = "beforeChangeShapes"
      End With
    End With
    Set ProcessShape = Nothing

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
'    lColorValue = setSheet.Range(setVal("cell_ProgressEnd") & line).Interior.Color
  
  startColumn = WBS_Option.���t�Z������(mainSheet.Range(setVal("cell_AchievementStart") & line))
  
  Call Library.showDebugForm("���ѐ��ݒ�", Library.TEXTJOIN("", True, mainSheet.Range("C" & line & ":" & setVal("cell_TaskAreaEnd") & line)))
  Call Library.showDebugForm("���ѐ��ݒ�", "�@�J�n��:" & mainSheet.Range(setVal("cell_AchievementStart") & line))
  Call Library.showDebugForm("���ѐ��ݒ�", "�@�I����:" & mainSheet.Range(setVal("cell_AchievementEnd") & line))
  Call Library.showDebugForm("���ѐ��ݒ�", "�@�i���@:" & mainSheet.Range(setVal("cell_Progress") & line))
  
  
  If mainSheet.Range(setVal("cell_AchievementEnd") & line) = "" Then
    endColumn = WBS_Option.���t�Z������(Date)
    
  ElseIf mainSheet.Range(setVal("cell_Progress") & line) = 100 Then
    If mainSheet.Range(setVal("cell_PlanEnd") & line) < mainSheet.Range(setVal("cell_AchievementEnd") & line) Then
      endColumn = WBS_Option.���t�Z������(mainSheet.Range(setVal("cell_AchievementEnd") & line))
    Else
      endColumn = WBS_Option.���t�Z������(mainSheet.Range(setVal("cell_PlanEnd") & line))
    End If
  Else
    endColumn = WBS_Option.���t�Z������(mainSheet.Range(setVal("cell_AchievementEnd") & line))
  End If
  
  
  Call Library.getRGB(setVal("lineColor_Achievement"), Red, Green, Blue)
  

  'Shape��z�u���邽�߂̊�ƂȂ�Z��
  Set rngStart = mainSheet.Range(startColumn & line)
  Set rngEnd = mainSheet.Range(endColumn & line)
  
  '�Z����Left�ATop�AWidth�v���p�e�B�𗘗p���Ĉʒu����
  BX = rngStart.Left
  BY = rngStart.top + (rngStart.Height / 2)
  EX = rngEnd.Left + rngEnd.Width
  EY = rngEnd.top + (rngEnd.Height / 2)
  
  
  With ActiveSheet.Shapes.AddLine(BX, BY, EX, EY).line
    If mainSheet.Range(setVal("cell_AssignP") & line) = "�H��" Or mainSheet.Range(setVal("cell_AssignA") & line) = "�H��" Then
      ActiveSheet.Shapes.Range(Array("�H��_" & line)).Select
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
  
  If setVal("setLightning") = False Then
    Exit Function
  ElseIf Not (setVal("startDay") <= setVal("baseDay") And setVal("baseDay") <= setVal("endDay")) Then
    Call Library.showNotice(50)
    setVal("setLightning") = False
    Exit Function
  End If
  
  Call Library.showDebugForm("�C�i�Y�}���ݒ�", mainSheet.Range("C" & line))
  
  baseColumn = WBS_Option.���t�Z������(setVal("baseDay"))
  
  '�C�i�Y�}���̐F�擾
  Call Library.getRGB(setVal("lineColor_Lightning"), Red, Green, Blue)

  If mainSheet.Range(setVal("cell_PlanStart") & line) <> "" Then
    startColumn = WBS_Option.���t�Z������(mainSheet.Range(setVal("cell_PlanStart") & line))
  Else
    startColumn = baseColumn
  End If
  
  If mainSheet.Range(setVal("cell_AchievementEnd") & line) <> "" Then
    endColumn = WBS_Option.���t�Z������(mainSheet.Range(setVal("cell_AchievementEnd") & line))
  
  ElseIf mainSheet.Range(setVal("cell_PlanEnd") & line) <> "" Then
    endColumn = WBS_Option.���t�Z������(mainSheet.Range(setVal("cell_PlanEnd") & line))
  Else
    endColumn = baseColumn
  End If
'  If IsEmpty(mainSheet.Range(setVal("cell_Progress") & line)) Then
'    progress = -1
'  Else
'    progress = mainSheet.Range(setVal("cell_Progress") & line).Value
'  End If
  
  'Shape��z�u���邽�߂̊�ƂȂ�Z��
  Set rngStart = mainSheet.Range(startColumn & line)
  Set rngEnd = mainSheet.Range(endColumn & line)
  Set rngBase = mainSheet.Range(baseColumn & line)
  
  '�C�i�Y�}���������Ȃ��ꍇ�́A����݈̂���
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
    
    Set rngChkDay = mainSheet.Range(chkDayColumn & line)
    
    
    
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
' * �Z���^�[
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function �Z���^�[()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  Dim baseColumn As String

  
  On Error GoTo catchError

  If setVal("startDay") >= setVal("baseDay") - 10 Then
  
  Else
    baseColumn = WBS_Option.���t�Z������(setVal("baseDay") - 10)
    Application.Goto Reference:=Range(baseColumn & 6), Scroll:=True
  End If
  


  Exit Function
'�G���[������=====================================================================================
catchError:

End Function





'**************************************************************************************************
' * �K���g�`���[�g�I��
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
' * Chart����\����𑀍�
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function changeShapes()
  Dim rng As Range

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
    changeShapesName = Replace(changeShapesName, "�^�X�N_", "")
    mainSheet.Range(setVal("cell_PlanStart") & changeShapesName) = Range(getColumnName(Range(tmp(0)).Column) & 4)
    mainSheet.Range(setVal("cell_PlanEnd") & changeShapesName) = Range(getColumnName(Range(tmp(1)).Column) & 4)
  Else
    tmp = rng.Address(False, False)
    changeShapesName = Replace(changeShapesName, "�^�X�N_", "")
    mainSheet.Range(setVal("cell_PlanStart") & changeShapesName) = Range(getColumnName(Range(tmp).Column) & 4)
    mainSheet.Range(setVal("cell_PlanEnd") & changeShapesName) = Range(getColumnName(Range(tmp).Column) & 4)
  End If
  
  ActiveSheet.Shapes("�^�X�N_" & changeShapesName).Delete
  Call �v����ݒ�(CLng(changeShapesName))
  
  changeShapesName = ""

End Function










