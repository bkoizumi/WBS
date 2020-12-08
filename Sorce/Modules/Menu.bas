Attribute VB_Name = "Menu"
'**************************************************************************************************
' * 設定
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Sub M_Help()
  Call init.setting
  sheetHelp.Visible = True
  sheetHelp.Select
End Sub



'**************************************************************************************************
' * ショートカット設定
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Sub M_ショートカット設定()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  
  Call init.setting(True)
  endLine = Cells(Rows.count, 7).End(xlUp).row
  
  '設定を解除
  Call M_ショートカット設定解除
  
  For line = 3 To endLine
    If sheetSetting.Range(setVal("cell_ShortcutKey") & line) <> "" Then
      Application.MacroOptions Macro:="Menu." & sheetSetting.Range(setVal("cell_ShortcutFuncName") & line), ShortcutKey:=sheetSetting.Range(setVal("cell_ShortcutKey") & line)
    End If
  Next
  'インデントのショートカット
  Application.OnKey "%{LEFT}", "Menu.M_インデント減"
  Application.OnKey "%{RIGHT}", "Menu.M_インデント増"
  Application.OnKey "%{F1}", "Menu.M_タスク表示_標準"
  Application.OnKey "%{F2}", "Menu.M_タスク表示_チームプランナー"
End Sub


Sub M_ショートカット設定解除()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  
  Call init.setting
  endLine = Cells(Rows.count, 7).End(xlUp).row
  
  '設定を解除
  For line = 3 To endLine
    If sheetSetting.Range("J" & line) <> "" Then
      Application.MacroOptions Macro:="Menu." & sheetSetting.Range("H" & line), ShortcutKey:=""
    End If
  Next
  
  'インデントのショートカット
  Application.OnKey "%{LEFT}", ""
  Application.OnKey "%{RIGHT}", ""
  Application.OnKey "%{F1}", ""
  Application.OnKey "%{F2}", ""
End Sub

Sub optionKey()
Attribute optionKey.VB_ProcData.VB_Invoke_Func = "O\n14"
  Call M_オプション画面表示
End Sub
Sub centerKey()
  Call M_センター
End Sub
Sub filterKey()
  Call M_フィルター
End Sub
Sub clearFilterKey()
  Call M_すべて表示
End Sub
Sub taskCheckKey()
Attribute taskCheckKey.VB_ProcData.VB_Invoke_Func = "C\n14"
  Call M_タスクチェック
End Sub
Sub makeGanttKey()
Attribute makeGanttKey.VB_ProcData.VB_Invoke_Func = "t\n14"
  Call M_ガントチャート生成
End Sub
Sub clearGanttKey()
Attribute clearGanttKey.VB_ProcData.VB_Invoke_Func = "D\n14"
  Call M_ガントチャートクリア
End Sub
Sub dispAllKey()
  Call M_すべて表示
End Sub
Sub taskControlKey()
'  Call M_
End Sub
Sub ScaleKey()
'  Call M_
End Sub








Sub M_オプション画面表示()
Attribute M_オプション画面表示.VB_ProcData.VB_Invoke_Func = " \n14"
  
  Call Library.startScript
  Call init.setting(True)
  
  Call Ctl_Option.オプション画面表示
  
  Call M_カレンダー生成(True)
  Call M_ガントチャート生成
  Call WBS_Option.表示列設定
  
  Call ctl_ProgressBar.showEnd
  Call Library.endScript(True)
End Sub


Sub M_列入替え()
  Call init.setting
  
  Call Library.startScript
  Call Check.項目列チェック
  Call init.setting(True)
  
  Call Library.endScript(True)
End Sub

Sub M_カレンダー生成(Optional flg As Boolean = False)

  Call init.setting(True)
  
  If flg = False Then
    Call Library.startScript
    Call ctl_ProgressBar.showStart
  End If
  
  '全ての行列を表示
  Cells.EntireColumn.Hidden = False
  Cells.EntireRow.Hidden = False
  
  Call Calendar.makeCalendar
  
  Call WBS_Option.複数の担当者行を非表示
  Call WBS_Option.表示列設定
  
  If flg = False Then
    Call ctl_ProgressBar.showEnd
    Call Library.endScript
  End If
End Sub




'**************************************************************************************************
' * 共通
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Sub M_行ハイライト()
  Call Library.startScript
  Call WBS_Option.setLineColor
  Call Library.endScript(True)
End Sub


'==================================================================================================
Sub M_全データ削除()
  If MsgBox("データを削除します", vbYesNo + vbExclamation) = vbNo Then
    End
  End If
  
  Call Library.startScript
  Call WBS_Option.clearAll
  Call Library.endScript
End Sub


Sub M_全画面()
Attribute M_全画面.VB_ProcData.VB_Invoke_Func = " \n14"
  Application.ScreenUpdating = False
  ActiveWindow.DisplayHeadings = False
  Application.DisplayFullScreen = True
  Application.ScreenUpdating = True
  
  With Frm_FullScreen
    .StartUpPosition = 0
    .top = Application.top + 300
    .Left = Application.Left + 30
    .Show vbModeless
  End With
End Sub

Sub M_タスク操作()
Attribute M_タスク操作.VB_ProcData.VB_Invoke_Func = " \n14"
End Sub

Sub M_スケール()
Attribute M_スケール.VB_ProcData.VB_Invoke_Func = " \n14"
End Sub


'**************************************************************************************************
' * WBS
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Sub M_タスクチェック()
Attribute M_タスクチェック.VB_ProcData.VB_Invoke_Func = "C\n14"


  Call init.setting
  sheetMain.Select
  
  Call Library.startScript
  Call ctl_ProgressBar.showStart
  
  Call Check.タスクリスト確認
  
  Call ctl_ProgressBar.showEnd
  Call Library.endScript(True)

End Sub


Sub M_フィルター()
Attribute M_フィルター.VB_ProcData.VB_Invoke_Func = " \n14"
  Call init.setting
  
  With FilterForm
    .StartUpPosition = 0
    .top = Application.top + (ActiveWindow.Width / 8)
    .Left = Application.Left + (ActiveWindow.Height / 8)
  End With
  
  FilterForm.Show
End Sub


Sub M_すべて表示()
Attribute M_すべて表示.VB_ProcData.VB_Invoke_Func = " \n14"
  Call Library.startScript
  Rows("6:" & Rows.count).EntireRow.Hidden = False
  
  Call WBS_Option.複数の担当者行を非表示
  Call Library.endScript
End Sub


Sub M_進捗コピー()
  Call Task.進捗コピー
End Sub

Sub M_インデント増()
  Dim selectedCells As Range
  Dim targetCell As Range
  
  On Error Resume Next
  
  Call Library.startScript
  Call init.setting
  sheetMain.Select
   
  Set selectedCells = Selection
  
  For Each targetCell In selectedCells
    Cells(targetCell.row, getColumnNo(setVal("cell_TaskArea"))).InsertIndent 1
  Next
  Call Library.endScript
End Sub


Sub M_インデント減()
  Dim selectedCells As Range
  Dim targetCell As Range
  
  On Error Resume Next
  
  Call Library.startScript
  Call init.setting
  sheetMain.Select
   
  Set selectedCells = Selection
  
  For Each targetCell In selectedCells
    Cells(targetCell.row, getColumnNo(setVal("cell_TaskArea"))).InsertIndent -1
  Next
  Call Library.endScript
End Sub


'進捗率設定========================================================================================
Sub M_進捗率設定(progress As Long)
  Call Task.進捗率設定(progress)
End Sub

'タスクのリンク設定/解除===========================================================================
Sub M_タスクのリンク設定()
  Call Library.startScript
  Call init.setting
  
  Call Task.taskLink
  
  Call Library.endScript
End Sub

Sub M_タスクのリンク解除()
  Call Library.startScript
  Call init.setting
  
  Call Task.taskUnlink
  
  Call Library.endScript
End Sub

Sub M_タスクの挿入()
  Call Library.startScript
  Call init.setting
  
  Call Task.タスクの挿入
  
  Call Library.endScript(True)
End Sub

Sub M_タスクの削除()
  Call Library.startScript
  Call init.setting
  
  Call Task.タスクの削除
  
  Call Library.endScript(True)
End Sub

'表示モード========================================================================================
Sub M_タスク表示_標準()
  Call Library.startScript
  
  Range("viewMode") = "Normal"
  Call init.setting(True)
  
  sheetMain.Visible = True
  sheetTeamsPlanner.Visible = True
  
  sheetMain.Select
  Call Check.項目列チェック
  
  Call WBS_Option.タスク表示_標準
  Call WBS_Option.setLineColor
  
  Application.Goto Reference:=Range("A6"), Scroll:=True
  sheetTeamsPlanner.Visible = xlSheetVeryHidden
  
  Call Library.endScript

End Sub

'==================================================================================================
Sub M_タスク表示_チームプランナー()
  Call Library.startScript
  Call init.setting(True)
  
  sheetMain.Visible = True
  sheetTeamsPlanner.Visible = True
  
  sheetMain.Select
  Call Check.項目列チェック
  
  sheetTeamsPlanner.Select
  Cells.EntireRow.Hidden = False
  Cells.EntireColumn.Hidden = False
  
  
  Call WBS_Option.タスク表示_チームプランナー
  Call WBS_Option.setLineColor
  
  Application.Goto Reference:=Range("A6"), Scroll:=True
  sheetSetting.Range("viewMode") = "TeamsPlanner"
  
  
  Call Library.endScript
End Sub

'==================================================================================================
Sub M_タスク表示_タスク()
  Call Library.startScript
  Call init.setting(True)
  
  Call WBS_Option.viewTask
  Call WBS_Option.setLineColor
  
  Call Library.endScript
End Sub


'==================================================================================================
Sub M_タスクにスクロール()
  Call Library.startScript
  Call init.setting
  
  Call WBS_Option.タスクにスクロール
  Call Library.endScript
End Sub

'==================================================================================================
Sub M_タイムラインに追加()
  Call Library.startScript
  Call init.setting
  
  Call Chart.タイムラインに追加(ActiveCell.row)
  Call Library.endScript(True)
End Sub


'**************************************************************************************************
' * ガントチャート
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'クリア===========================================================================================-
Sub M_ガントチャートクリア()
Attribute M_ガントチャートクリア.VB_ProcData.VB_Invoke_Func = "D\n14"
  Call Library.startScript
  Call Chart.ガントチャート削除
  Call Library.endScript
End Sub

'生成のみ==========================================================================================
Sub M_ガントチャート生成のみ()
Attribute M_ガントチャート生成のみ.VB_ProcData.VB_Invoke_Func = "A\n14"
  Call init.setting
  Call Library.startScript
  Call ctl_ProgressBar.showStart
  Call Library.showDebugForm("ガントチャート生成", "処理開始")
  
  Call Chart.ガントチャート生成
  
  Call Library.showDebugForm("ガントチャート生成", "処理完了")
  Call ctl_ProgressBar.showEnd
  Call Library.endScript(True)
End Sub


'生成==============================================================================================
Sub M_ガントチャート生成()
Attribute M_ガントチャート生成.VB_ProcData.VB_Invoke_Func = "t\n14"
  Call init.setting
  
  Call Library.startScript
  Call ctl_ProgressBar.showStart
  
  If Range("viewMode") = "Normal" Then
    Call Check.タスクリスト確認
  End If
  Call Chart.ガントチャート生成
  
  Call ctl_ProgressBar.showEnd
  Call Library.endScript(True)
  Application.EnableEvents = True
End Sub



'センター==============================================================================================
Sub M_センター()
Attribute M_センター.VB_ProcData.VB_Invoke_Func = " \n14"

  Call init.setting
  Call Library.startScript
  Call ctl_ProgressBar.showStart
  Call Library.showDebugForm("センターへ移動", "処理開始")
  
  Call Chart.センター
  
  Call Library.showDebugForm("センターへ移動", "処理完了")
  Call ctl_ProgressBar.showEnd
  Call Library.endScript(True)
End Sub


'**************************************************************************************************
' * M_インポート
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'Excelファイル=====================================================================================
Sub M_インポートExcel()
  
  Call Library.startScript
  Call Library.showDebugForm("ファイルインポート", "処理開始")
  
  Call init.setting
  endLine = sheetMain.Cells(Rows.count, 1).End(xlUp).row
  
  If setVal("workMode") = "CD部" Then
    deleteFlg = False
  Else
    If MsgBox("データを削除します", vbYesNo + vbExclamation) = vbYes Then
      deleteFlg = True
    Else
      deleteFlg = False
    End If
  End If
  Call ctl_ProgressBar.showStart
  
  
  Call import.ファイルインポート
  Call Calendar.書式設定
  Call import.カレンダー用日程取得
  
  If setVal("lineColorFlg") = "True" Then
    setVal("lineColorFlg") = False
    Call WBS_Option.setLineColor
  Else
  End If
  
  Call WBS_Option.表示列設定
  Call M_画面再描写
  
  Call ctl_ProgressBar.showEnd
  Call Library.endScript
  
  Call WBS_Option.saveAndRefresh
  
  Application.Goto Reference:=Range("A6"), Scroll:=True


  Err.Clear
  Call Library.showNotice(200, "インポート")
End Sub




Sub M_画面再描写()
  Dim zoomLevel As Integer
  
  zoomLevel = ActiveWindow.Zoom
  ActiveWindow.Zoom = 100
  ActiveWindow.Zoom = zoomLevel

End Sub











