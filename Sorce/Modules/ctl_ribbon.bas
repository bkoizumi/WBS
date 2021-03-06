Attribute VB_Name = "ctl_ribbon"
#If VBA7 And Win64 Then
  Private Declare PtrSafe Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal cbLen As LongPtr)
#Else
  Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal cbLen As Long)
#End If


Public ribbonUI As IRibbonUI        ' リボン
Private rbButton_Visible As Boolean ' ボタンの表示／非表示
Private rbButton_Enabled As Boolean ' ボタンの有効／無効

'トグルボタン
Public PressT_B015 As Boolean


'**************************************************************************************************
' * リボンメニュー設定
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'読み込み時処理====================================================================================
Function onLoad(ribbon As IRibbonUI)
  On Error GoTo catchError
  
  Set ribbonUI = ribbon
  
  Call init.setting(True)
  Call Library.delRegistry(RegistryRibbonName)

  Call Library.setRegistry(RegistryRibbonName, CStr(ObjPtr(ribbonUI)))
  
  ribbonUI.ActivateTab (RibbonTabName)
  
  'リボンの表示を更新する
  ribbonUI.Invalidate

  Exit Function
'エラー発生時--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(400, "Ctl_Ribbon.onLoad", True)
End Function



'リボンタブアクティブ==============================================================================
Function setRibbonTabActive()
  Dim regVal As String
  
  On Error GoTo catchError
  
  regVal = Library.getRegistry(RegistryRibbonName)
  If regVal = "" Then
    Exit Function
  End If

   If ribbonUI Is Nothing Then
    #If VBA7 And Win64 Then
      Set ribbonUI = GetRibbon(CLngPtr(regVal))
    #Else
      Set ribbonUI = GetRibbon(CLng(regVal))
    #End If
  End If
  ribbonUI.ActivateTab (RibbonTabName)
  ribbonUI.Invalidate
  
  Exit Function
'エラー発生時--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(400, "Ctl_Ribbon.setRibbonTabActive", True)
End Function


'--------------------------------------------------------------------------------------------------
#If VBA7 And Win64 Then
Function GetRibbon(ByVal lRibbonPointer As LongPtr) As Object
  Dim p As LongPtr
#Else
Function GetRibbon(ByVal lRibbonPointer As Long) As Object
  Dim p As Long
#End If
  Dim ribbonObj As Object
  
  MoveMemory ribbonObj, lRibbonPointer, LenB(lRibbonPointer)
  Set GetRibbon = ribbonObj
  p = 0: MoveMemory ribbonObj, p, LenB(p) '後始末
End Function


'トグルボタン制御==================================================================================
Sub getPressed(control As IRibbonControl, ByRef returnedVal)
  Select Case control.ID
  
    Case "T_B015"         'タイムラインに追加
      If Range(setVal("cell_Info") & ActiveCell.row) Like "" Then
        returnedVal = True
      Else
        returnedVal = False
      End If
    Case Else
  End Select
End Sub


'Label 設定========================================================================================
Public Sub getLabel(control As IRibbonControl, ByRef setRibbonVal)
  setRibbonVal = getRibbonMenu(control.ID, 2)
End Sub


'onAction==========================================================================================
Sub getonAction(control As IRibbonControl)
  Dim setRibbonVal As String

  setRibbonVal = getRibbonMenu(control.ID, 3)
  Application.run setRibbonVal

End Sub


'Supertip 設定=====================================================================================
Public Sub getSupertip(control As IRibbonControl, ByRef setRibbonVal)
  setRibbonVal = getRibbonMenu(control.ID, 5)
End Sub


'Description 設定==================================================================================
Public Sub getDescription(control As IRibbonControl, ByRef setRibbonVal)
  setRibbonVal = getRibbonMenu(control.ID, 6)
End Sub


'size 設定=========================================================================================
Public Sub getsize(control As IRibbonControl, ByRef setRibbonVal)
  Dim getVal As String
  getVal = getRibbonMenu(control.ID, 4)

  Select Case getVal
    Case "large"
      setRibbonVal = 1
    Case "normal"
      setRibbonVal = 0
    Case Else
  End Select
End Sub


'有効/無効切り替え ========================================================================
Function getEnabled(control As IRibbonControl, ByRef returnedVal)
  Call init.setting
  If setVal("workMode") = "CD部" Then
    returnedVal = False
  Else
    returnedVal = True
  End If
  
End Function



'Ribbonシートから内容を取得========================================================================
Function getRibbonMenu(menuId As String, offsetVal As Long)
  Dim getString As String
  Dim FoundCell As Range
  Dim ribSheet As Worksheet
  Dim endLine As Long

  On Error GoTo catchError

  Call Library.startScript
  Set ribSheet = ThisWorkbook.Worksheets("Ribbon")

  endLine = ribSheet.Cells(Rows.count, 1).End(xlUp).row

  getRibbonMenu = Application.VLookup(menuId, ribSheet.Range("A2:F" & endLine), offsetVal, False)
  Call Library.endScript


  Exit Function

'エラー発生時------------------------------------
catchError:
  getRibbonMenu = "エラー"

End Function



'**************************************************************************************************
' * 共通
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'選択行色付切替====================================================================================
Function setLineColor(control As IRibbonControl)
  Call menu.M_行ハイライト
End Function

'**************************************************************************************************
' * 設定
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'Help==============================================================================================
Function Help(control As IRibbonControl)
  Call menu.M_Help
End Function

'オプション========================================================================================
Function dispOption(control As IRibbonControl)
  Call Library.showDebugForm("オプション画面表示", "処理開始")
  Call menu.M_オプション画面表示
  Call Library.showDebugForm("オプション画面表示", "処理終了")
End Function

'列入替え========================================================================================
Function changeColumn(control As IRibbonControl)
  Call Library.showDebugForm("列入替え", "処理開始")
  Call menu.M_列入替え
  Call Library.showDebugForm("列入替え", "処理終了")
End Function

'全データ削除=====================================================================================-
Function clearAll(control As IRibbonControl)
  Call Library.showDebugForm("全データ削除", "処理開始")
  Call menu.M_全データ削除
  Call Library.showDebugForm("全データ削除", "処理終了")
End Function

'生成==============================================================================================
Function makeCalendar(control As IRibbonControl)
  Call Library.showDebugForm("カレンダー生成", "処理開始")
  Call menu.M_カレンダー生成
  Call Library.showDebugForm("カレンダー生成", "処理完了")
End Function

'全画面表示========================================================================================
Function DispFullScreen(control As IRibbonControl)
  Call menu.M_全画面
End Function


'**************************************************************************************************
' * WBS
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'タスクリスト確認==================================================================================
Function chkTaskList(control As IRibbonControl)
  Call Library.showDebugForm("タスクリスト確認", "処理開始")
  Call menu.M_タスクチェック
  Call Library.showDebugForm("タスクリスト確認", "処理終了")
End Function

'フィルター========================================================================================
Function setFilter(control As IRibbonControl)
  Call Library.showDebugForm("フィルター", "処理開始")
  Call menu.M_フィルター
  Call Library.showDebugForm("フィルター", "処理終了")
End Function

'すべて表示========================================================================================
Function dispAllList(control As IRibbonControl)
  Call Library.showDebugForm("すべて表示", "処理開始")
  Call menu.M_すべて表示
  Call Library.showDebugForm("すべて表示", "処理終了")
End Function

'進捗コピー========================================================================================
Function copyProgress(control As IRibbonControl)
  Call Library.showDebugForm("進捗コピー", "処理開始")
  Call menu.M_進捗コピー
  Call Library.showDebugForm("進捗コピー", "処理終了")
End Function

'インデント========================================================================================
Function taskOutdent(control As IRibbonControl)
  Call menu.M_インデント増
End Function
Function taskIndent(control As IRibbonControl)
  Call menu.M_インデント減
End Function

'進捗率設定========================================================================================
Function progress_0(control As IRibbonControl)
  Call menu.M_進捗率設定(0)
End Function
Function progress_25(control As IRibbonControl)
  Call menu.M_進捗率設定(25)
End Function
Function progress_50(control As IRibbonControl)
  Call menu.M_進捗率設定(50)
End Function
Function progress_75(control As IRibbonControl)
  Call menu.M_進捗率設定(75)
End Function
Function progress_100(control As IRibbonControl)
  Call menu.M_進捗率設定(100)
End Function

'タスクのリンク====================================================================================
Function taskLink(control As IRibbonControl)
  Call menu.M_タスクのリンク設定
End Function
Function taskUnlink(control As IRibbonControl)
  Call menu.M_タスクのリンク解除
End Function


'表示モード========================================================================================
Function viewNormal(control As IRibbonControl)
  Call menu.M_タスク表示_標準
End Function

Function viewTask(control As IRibbonControl)
  Call menu.M_タスク表示_タスク
End Function

Function viewTeamsPlanner(control As IRibbonControl)
  Call menu.M_タスク表示_チームプランナー
End Function

'タスクにスクロール========================================================================================
Function scrollTask(control As IRibbonControl)
  
  Call Library.showDebugForm("タスクにスクロール", "処理開始")
  Call menu.M_タスクにスクロール
  Call Library.showDebugForm("タスクにスクロール", "処理終了")
  
End Function

'タイムラインに追加========================================================================================
Function addTimeLine(control As IRibbonControl)
  Call menu.M_タイムラインに追加
End Function


'**************************************************************************************************
' * ガントチャート
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'クリア===========================================================================================-
Function clearChart(control As IRibbonControl)
  Call menu.M_ガントチャートクリア
End Function

'生成==============================================================================================
Function makeChart(control As IRibbonControl)
  Call Library.showDebugForm("ガントチャート生成", "処理開始")
  Call menu.M_ガントチャート生成
  Call Library.showDebugForm("ガントチャート生成", "処理終了")
End Function

'センター==============================================================================================
Function setCenter(control As IRibbonControl)
  Call menu.M_センター
End Function


'**************************************************************************************************
' * import
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'Excelファイル=====================================================================================
Function importExcel(control As IRibbonControl)
  Call menu.M_インポートExcel
End Function

