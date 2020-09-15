Attribute VB_Name = "Module1"
Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
'

'
    ActiveSheet.Shapes.Range(Array("Straight Connector 193")).Select
    Selection.ShapeRange.Name = "é¿ê—13"
    Selection.Name = "é¿ê—13"
End Sub
Sub Macro2()
Attribute Macro2.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro2 Macro
'

'
    With Selection.ShapeRange.line
      .Visible = msoTrue
      .ForeColor.RGB = RGB(255, 0, 0)
      .Transparency = 0
    End With
End Sub
Sub Macro3()
Attribute Macro3.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro3 Macro
'

'
    With Selection.ShapeRange.line
      .Visible = msoTrue
      .Weight = 2.25
    End With
    With Selection.ShapeRange.line
      .Visible = msoTrue
      .ForeColor.RGB = RGB(255, 0, 0)
      .Transparency = 0
    End With
End Sub
