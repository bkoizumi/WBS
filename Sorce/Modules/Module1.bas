Attribute VB_Name = "Module1"
Sub Macro2()
Attribute Macro2.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro2 Macro
'

'
    ActiveSheet.Shapes.AddConnector(msoConnectorElbow, 890.7692125984, _
        593.0768503937, 940.3846456693, 595.961496063).Select
      Selection.ShapeRange.line.EndArrowheadStyle = msoArrowheadTriangle
    Selection.ShapeRange.ConnectorFormat.BeginConnect ActiveSheet.Shapes("タスク_28") _
        , 3
    Selection.ShapeRange.ScaleWidth 0.9418599115, msoFalse, msoScaleFromBottomRight
    Selection.ShapeRange.ScaleHeight 12.1998635185, msoFalse, msoScaleFromTopLeft
    Selection.ShapeRange.ConnectorFormat.EndConnect ActiveSheet.Shapes("タスク_29"), _
        2
    Selection.ShapeRange.ScaleWidth 0.5079376771, msoFalse, msoScaleFromTopLeft
    Selection.ShapeRange.ScaleHeight 2, msoFalse, msoScaleFromBottomRight
    Selection.ShapeRange.ScaleHeight 0.5, msoFalse, msoScaleFromTopLeft
    Selection.ShapeRange.ScaleHeight 0.1895416674, msoFalse, _
        msoScaleFromBottomRight
    Selection.ShapeRange.IncrementRotation -90
    Selection.ShapeRange.Flip msoFlipVertical
    Selection.ShapeRange.IncrementRotation 90
    
    With Selection.ShapeRange.line
      .Visible = msoTrue
      .Weight = 1.5
    End With
End Sub
