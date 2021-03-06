Attribute VB_Name = "罫線"
Sub 月初()
  With Selection.Borders(xlEdgeLeft)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
  End With
  With Selection.Borders(xlEdgeRight)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlHairline
  End With

End Sub

Sub 月中()
  With Selection.Borders(xlEdgeLeft)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlHairline
  End With
  With Selection.Borders(xlEdgeRight)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlHairline
  End With
End Sub

Sub 月末()
  With Selection.Borders(xlEdgeLeft)
    .LineStyle = xlContinuous
    .ColorIndex = xlAutomatic
    .TintAndShade = 0
    .Weight = xlHairline
  End With
  With Selection.Borders(xlEdgeRight)
    .LineStyle = xlContinuous
    .ColorIndex = xlAutomatic
    .TintAndShade = 0
    .Weight = xlThin
  End With
End Sub

Sub 最終日()
  With Selection.Borders(xlEdgeRight)
    .LineStyle = xlDouble
    .Weight = xlThin
  End With
End Sub

Sub 二重線()
  With Selection.Borders(xlEdgeLeft)
    .LineStyle = xlDouble
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThick
  End With
End Sub



Sub 横線()
  With Selection.Borders(xlEdgeTop)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
  End With
  With Selection.Borders(xlEdgeBottom)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
  End With
  With Selection.Borders(xlInsideHorizontal)
    .LineStyle = xlContinuous
    .ColorIndex = 0
    .TintAndShade = 0
    .Weight = xlThin
  End With
End Sub



Sub 囲み罫線()
    With Selection
      .Borders(xlDiagonalDown).LineStyle = xlNone
      .Borders(xlDiagonalUp).LineStyle = xlNone
      .Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Borders(xlEdgeTop).LineStyle = xlContinuous
      .Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Borders(xlEdgeRight).LineStyle = xlContinuous
      .Borders(xlInsideVertical).LineStyle = xlContinuous
      .Borders(xlInsideHorizontal).LineStyle = xlContinuous
    End With
End Sub

