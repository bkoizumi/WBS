Attribute VB_Name = "init"
'ワークブック用変数==============================
Public ThisBook As Workbook
Public targetBook As Workbook

'ワークシート用変数==============================
Public sheetNotice As Worksheet
Public sheetHelp As Worksheet
Public sheetSetting As Worksheet
Public tmpSheet As Worksheet
Public sheetMain As Worksheet
Public sheetTeamsPlanner As Worksheet

'グローバル変数==================================
Public Const thisAppName = "Work Breakdown Structure for Excel"
Public Const thisAppVersion = "0.0.3.0"


Public setVal As Collection
Public getVal As Collection
Public memberColor As Object

Public sheetMainName As String
Public sheetTeamsPlannerName As String

'レジストリ登録用サブキー
Public Const RegistryKey As String = "B.Koizumi"
Public Const RegistrySubKey As String = "WBS"
Public Const RibbonTabName As String = "WBSTab"
Public RegistryRibbonName As String


'ログファイル
Public logFile As String

'ガントチャート選択
Public selectShapesName(0) As Variant
Public changeShapesName As String


Public deleteFlg As Boolean

'**************************************************************************************************
' * 設定クリア
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function clearSetting()
  Set sheetHelp = Nothing
  Set sheetNotice = Nothing
  Set sheetSetting = Nothing
  Set sheetMain = Nothing
  Set tmpSheet = Nothing
  Set sheetTeamsPlanner = Nothing
  
  Set setVal = Nothing
  Set memberColor = Nothing

  
End Function
'**************************************************************************************************
' * 設定
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function setting(Optional reCheckFlg As Boolean)
  Dim line As Long
  
'  On Error GoTo catchError
  
  If ThisWorkbook.Worksheets("設定").Range("B3") = "develop" Then
    ThisWorkbook.Save
  End If
  
  If logFile <> "" And reCheckFlg <> True Then
    Exit Function
  End If

Label_reset:
  
  'ブックの設定------------------------------------------------------------------------------------
  Set ThisBook = ThisWorkbook
  ThisBook.Activate
  
  'ワークシート名の設定----------------------------------------------------------------------------
  sheetMainName = "メイン"
  sheetTeamsPlannerName = "チームプランナー"
  Set sheetHelp = ThisBook.Worksheets("Help")
  Set sheetNotice = ThisBook.Worksheets("Notice")
  Set sheetSetting = ThisBook.Worksheets("設定")
  Set tmpSheet = ThisBook.Worksheets("Tmp")
  
  If sheetSetting.Range("B9") = "Normal" Then
    Set sheetMain = ThisBook.Worksheets(sheetMainName)
    Set sheetTeamsPlanner = ThisBook.Worksheets(sheetTeamsPlannerName)
  ElseIf sheetSetting.Range("B9") = "TeamsPlanner" Then
    Set sheetMain = ThisBook.Worksheets(sheetTeamsPlannerName)
    Set sheetTeamsPlanner = ThisBook.Worksheets(sheetMainName)
  End If
  
  Set setVal = New Collection
  Set memberColor = CreateObject("Scripting.Dictionary")
  
  
  '初期値設定--------------------------------------------------------------------------------------
  '期間、基準日
  Select Case True
    Case sheetSetting.Range("B6") = ""
      sheetSetting.Range("B6") = Format(DateAdd("d", 0, Date), "yyyy/mm/dd")
    
    Case sheetSetting.Range("B7") = ""
      sheetSetting.Range("A7") = Format(DateAdd("d", 60, Date), "yyyy/mm/dd")
    
    Case sheetSetting.Range("B8") = ""
      sheetSetting.Range("B8") = Format(DateAdd("d", 0, Date), "yyyy/mm/dd")
  End Select
  
  If sheetSetting.Range("B4") = "CD部" Then
    sheetSetting.Range("B8") = Format(Date, "yyyy/mm/dd")
  End If
  
  'VBA用の設定値取得-------------------------------------------------------------------------------
  With setVal
    For line = 3 To sheetSetting.Cells(Rows.count, 1).End(xlUp).row
      If sheetSetting.Range("A" & line) <> "" Then
       .Add item:=sheetSetting.Range("B" & line), Key:=sheetSetting.Range("A" & line)
      End If
    Next
    For line = 3 To sheetSetting.Cells(Rows.count, 4).End(xlUp).row
      If sheetSetting.Range("D" & line) <> "" Then
       .Add item:=sheetSetting.Range("E" & line), Key:=sheetSetting.Range("D" & line)
      End If
    Next
  End With
  
  'ショートカットキー設定--------------------------------------------------------------------------
  With setVal
    For line = 3 To sheetSetting.Cells(Rows.count, 7).End(xlUp).row
      .Add item:=sheetSetting.Range(setVal("cell_ShortcutKey") & line), Key:=sheetSetting.Range(setVal("cell_ShortcutFuncName") & line)
    Next
  End With
  

  '担当者色読み込み--------------------------------------------------------------------------------
  For line = 3 To sheetSetting.Cells(Rows.count, 11).End(xlUp).row
    If sheetSetting.Range("K" & line).Value <> "" Then
      memberColor.Add sheetSetting.Range("K" & line).Value, sheetSetting.Range("K" & line).Interior.Color
    End If
  Next line


  'ファイル関連設定--------------------------------------------------------------------------------
  logFile = ThisBook.Path & "\ExcelMacro.log"
  
  
  'レジストリ関連設定------------------------------------------------------------------------------
  RegistryRibbonName = "RP_" & ActiveWorkbook.Name
  
  
  
  
'  If reCheckFlg = True Then
'    Call Check.項目列チェック
'    reCheckFlg = False
'    Call clearSetting
'
'    GoTo Label_reset
'  End If
  
  Call 名前定義
  Exit Function
  
'エラー発生時--------------------------------------------------------------------------------------
catchError:
  logFile = ""
'  Call Library.showNotice(Err.Number, Err.Description, True)
  
  GoTo Label_reset
  
End Function

'**************************************************************************************************
' * 休日設定
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function chkHollyday(chkDate As Date, HollydayName As String)
  Dim line As Long, endLine As Long
  Dim strFilMessage() As Date
  
  '休日判定
  Call GetHollyday(CDate(chkDate), HollydayName)
  
  '土日を判定
  If HollydayName = "" Then
    If Weekday(chkDate) = vbSunday Then
      HollydayName = "Sunday"
    ElseIf Weekday(chkDate) = vbSaturday Then
      HollydayName = "Saturday"
    End If
  End If
End Function


'**************************************************************************************************
' * 名前定義
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function 名前定義()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  Dim Name As Object
  
  On Error GoTo catchError

  For Each Name In Names
    If Name.Visible = False Then
      Name.Visible = True
    End If
    If Not Name.Name Like "*!Print_Area" And Not Name.Name Like "*!Print_Titles" Then
      Name.Delete
    End If
  Next
  
  'VBA用の設定
  For line = 3 To sheetSetting.Range("B5")
    If sheetSetting.Range("A" & line) <> "" Then
      sheetSetting.Range(setVal("cell_LevelInfo") & line).Name = sheetSetting.Range("A" & line)
    End If
  Next
  
  'ショートカットキーの設定
  endLine = sheetSetting.Cells(Rows.count, Library.getColumnNo(setVal("cell_ShortcutFuncName"))).End(xlUp).row
  For line = 3 To endLine
    If sheetSetting.Range(setVal("cell_ShortcutFuncName") & line) <> "" Then
      sheetSetting.Range(setVal("cell_ShortcutKey") & line).Name = sheetSetting.Range(setVal("cell_ShortcutFuncName") & line)
    End If
  Next
  
  
  endLine = sheetSetting.Cells(Rows.count, 11).End(xlUp).row
  sheetSetting.Range(setVal("cell_AssignorList") & "3:" & setVal("cell_AssignorList") & endLine).Name = "担当者"

  endLine = sheetSetting.Cells(Rows.count, 17).End(xlUp).row
  sheetSetting.Range(setVal("cell_CompanyHoliday") & "3:" & setVal("cell_CompanyHoliday") & endLine).Name = "休日リスト"

  Exit Function
'エラー発生時--------------------------------------------------------------------------------------
catchError:
  Call Library.showNotice(Err.Number, Err.Description, True)
  
End Function


'**************************************************************************************************
' * シートの表示/非表示
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function noDispSheet()

  Call init.setting
  tmpSheet.Visible = xlSheetVeryHidden
  sheetNotice.Visible = xlSheetVeryHidden
  Worksheets("サンプル").Visible = xlSheetVeryHidden
  sheetTeamsPlanner.Visible = xlSheetVeryHidden
  
  Worksheets(sheetMainName).Select
End Function



Function dispSheet()

  Call init.setting
  Worksheets("Help").Visible = True
  Worksheets("Tmp").Visible = True
  Worksheets("Notice").Visible = True
  Worksheets("設定").Visible = True
  Worksheets("サンプル").Visible = True
  
  Worksheets(sheetTeamsPlannerName).Visible = True
  Worksheets(sheetMainName).Visible = True
  
  Worksheets(sheetMainName).Select
  
End Function

