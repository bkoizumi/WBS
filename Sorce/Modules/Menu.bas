Attribute VB_Name = "Menu"
Sub M_xxxxxxxxxxxxxxxxx()
  Call Library.startScript
  Call ProgressBar.showStart

  Call init.setting(True)
  

  Call ProgressBar.showEnd
  Call Library.endScript
End Sub



'**************************************************************************************************
' * 設定
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Sub M_Help()
  Call init.setting

End Sub



Sub M_ショートカット設定()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  
  Call init.setting
  endLine = Cells(Rows.count, 7).End(xlUp).row
  
  '設定を解除
  For line = 3 To endLine
    If setSheet.Range("J" & line) <> "" Then
      Application.MacroOptions Macro:="Menu." & setSheet.Range("H" & line), ShortcutKey:=""
    End If
  Next
  
  For line = 3 To endLine
    If setSheet.Range("I" & line) <> "" Then
      Application.MacroOptions Macro:="Menu." & setSheet.Range("G" & line), ShortcutKey:=setSheet.Range("I" & line)
    End If
  Next
  'インデントのショートカット
  Application.OnKey "%{LEFT}", "Menu.M_インデント減"
  Application.OnKey "%{RIGHT}", "Menu.M_インデント増"
  Application.OnKey "%{F1}", "WBS_Option.表示_標準"
  Application.OnKey "%{F2}", "WBS_Option.表示_ガントチャート"
  
  If setVal("debugMode") <> "develop" Then
    Application.OnKey "^v", "Menu.M_貼り付け"
  End If

End Sub


Sub M_オプション画面表示()
Attribute M_オプション画面表示.VB_ProcData.VB_Invoke_Func = " \n14"
  Call init.setting(True)
  
  Call Library.startScript
  endLine = setSheet.Cells(Rows.count, 7).End(xlUp).row
  setSheet.Range("I3:I" & endLine).Copy
  setSheet.Range("J3:J" & endLine).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
  Application.CutCopyMode = False
  Call WBS_Option.オプション画面表示
  
  setSheet.Range("J3:J" & endLine).ClearContents
  Call Library.endScript(True)
End Sub


Sub M_カレンダー生成()

  Call init.setting(True)
  
  
  Call Library.startScript
  Call ProgressBar.showStart
  
  Call Library.showDebugForm("カレンダー生成", "処理開始")
  Call Calendar.makeCalendar
  
  Call Library.showDebugForm("カレンダー生成", "処理完了")
  
  Call WBS_Option.表示列設定
  Call ProgressBar.showEnd
  Call Library.endScript

End Sub




'**************************************************************************************************
' * 共通
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Sub M_選択行色付切替()
  Call Library.startScript
  Call WBS_Option.setLineColor
  Call Library.endScript(True)
End Sub


'--------------------------------------
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
  
  With DispFullScreenForm
    .StartUpPosition = 0
    .top = Application.top + 300
    .Left = Application.Left + 30
  End With
  Application.ScreenUpdating = True
  DispFullScreenForm.Show vbModeless
End Sub

Sub M_タスク操作()
Attribute M_タスク操作.VB_ProcData.VB_Invoke_Func = " \n14"
End Sub

Sub M_スケール()
Attribute M_スケール.VB_ProcData.VB_Invoke_Func = " \n14"
End Sub


Function M_貼り付け()
  Selection.PasteSpecial Paste:=xlPasteAllExceptBorders, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
End Function


'**************************************************************************************************
' * WBS
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Sub M_タスクチェック()
Attribute M_タスクチェック.VB_ProcData.VB_Invoke_Func = "C\n14"

  Call init.setting
  mainSheet.Select
  
  Application.CalculateFull
  Call Library.startScript
  Call ProgressBar.showStart
  
  Call Library.showDebugForm("タスクチェック", "処理開始")
  Call Check.タスクリスト確認
  
  Call Library.showDebugForm("タスクチェック", "処理完了")
  Call ProgressBar.showEnd
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
  mainSheet.Select
   
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
  mainSheet.Select
   
  Set selectedCells = Selection
  
  For Each targetCell In selectedCells
    Cells(targetCell.row, getColumnNo(setVal("cell_TaskArea"))).InsertIndent -1
  Next
  Call Library.endScript
End Sub


'進捗率設定----------------------------------------------------------------------------------------
Sub M_進捗率設定(progress As Long)
  Call Task.進捗率設定(progress)
End Sub

'タスクのリンク設定/解除---------------------------------------------------------------------------
Sub M_タスクのリンク設定()
  Call Task.taskLink
End Sub
Sub M_タスクのリンク解除()
  Call Task.taskUnlink
End Sub
Sub M_タスクの挿入()
  Call Task.rTaskInsert
End Sub
Sub M_タスクの削除()
  Call Task.rTaskDell
End Sub

'表示モード----------------------------------------------------------------------------------------
Sub M_タスク表示_標準()
  Call Library.startScript
  
  Call init.setting
  If setVal("debugMode") <> "develop" Then
    mainSheet.Visible = True
    TeamsPlannerSheet.Visible = xlSheetVeryHidden
  End If
  
  Call init.setting(True)
  Call WBS_Option.viewNormal
  Call Library.endScript

End Sub

Sub M_タスク表示_タスク()
  Call Library.startScript
  Call init.setting(True)
  
  Call WBS_Option.viewTask
  
  Call Library.endScript
End Sub

Sub M_タスク表示_チームプランナー()
  Call Library.startScript
  Call init.setting(True)
  
  Call WBS_Option.viewTeamsPlanner
  Call Library.endScript
End Sub

Sub M_タスク表示_設定()
  Call WBS_Option.viewSetting
  Call Library.endScript
End Sub



'**************************************************************************************************
' * ガントチャート
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'クリア--------------------------------------------------------------------------------------------
Sub M_ガントチャートクリア()
Attribute M_ガントチャートクリア.VB_ProcData.VB_Invoke_Func = "D\n14"
  Call Library.startScript
  Call Chart.ガントチャート削除
  Call Library.endScript
End Sub

'生成のみ------------------------------------------------------------------------------------------
Sub M_ガントチャート生成のみ()
Attribute M_ガントチャート生成のみ.VB_ProcData.VB_Invoke_Func = "A\n14"
  Call init.setting
  Call Library.startScript
  Call ProgressBar.showStart
  Call Library.showDebugForm("ガントチャート生成", "処理開始")
  
  Call Chart.ガントチャート生成
  
  Call Library.showDebugForm("ガントチャート生成", "処理完了")
  Call ProgressBar.showEnd
  Call Library.endScript(True)
End Sub


'生成----------------------------------------------------------------------------------------------
Sub M_ガントチャート生成()
Attribute M_ガントチャート生成.VB_ProcData.VB_Invoke_Func = "t\n14"
  Call init.setting
  
  Call Library.startScript
  Call ProgressBar.showStart
  Call Library.showDebugForm("ガントチャート生成", "処理開始")
  
  Call Check.タスクリスト確認
  Call Chart.ガントチャート生成
  
  Call Library.showDebugForm("ガントチャート生成", "処理完了")
  Call ProgressBar.showEnd
  Call Library.endScript(True)
  Application.EnableEvents = True
End Sub



'センター----------------------------------------------------------------------------------------------
Sub M_センター()
Attribute M_センター.VB_ProcData.VB_Invoke_Func = " \n14"
  Call Library.startScript
  Call ProgressBar.showStart
  Call Library.showDebugForm("センターへ移動", "処理開始")
  
  Call Chart.センター
  
  Call Library.showDebugForm("センターへ移動", "処理完了")
  Call ProgressBar.showEnd
  Call Library.endScript
End Sub


'**************************************************************************************************
' * import
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'Excelファイル-------------------------------------------------------------------------------------
Sub M_Excelインポート()
  
  Call Library.startScript
  Call Library.showDebugForm("ファイルインポート", "処理開始")
  If MsgBox("データを削除します", vbYesNo + vbExclamation) = vbYes Then
    Call WBS_Option.clearAll
  Else
    Call WBS_Option.clearCalendar
  End If
  Call ProgressBar.showStart
  
  
  Call import.ファイルインポート
  Call Calendar.書式設定
  Call import.makeCalendar
  
  
  Call ProgressBar.showEnd
  Call Library.endScript
  
  Call WBS_Option.saveAndRefresh
  
  Err.Clear
  Call Library.showNotice(200, "インポート")
End Sub
















