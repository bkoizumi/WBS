Attribute VB_Name = "init"
'ワークブック用変数------------------------------
Public ThisBook As Workbook

'ワークシート用変数------------------------------
Public noticeCodeSheet As Worksheet
Public helpSheet As Worksheet
Public setSheet As Worksheet
Public tmpSheet As Worksheet
Public mainSheet As Worksheet
Public TeamsPlannerSheet As Worksheet



'グローバル変数----------------------------------
Public Const thisAppName = "Excel for Work Breakdown Structure"



Public setVal As Collection
Public getVal As Collection
Public memberColor As Object

Public mainSheetName As String
Public TeamsPlannerSheetName As String

'レジストリ登録用サブキー
Public Const RegistrySubKey As String = "WBS"

'ログファイル
Public logFile As String

'ガントチャート選択
Public selectShapesName(0) As Variant
Public changeShapesName As String


'***********************************************************************************************************************************************
' * 設定クリア
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'***********************************************************************************************************************************************
Function clearSetting()
  Set helpSheet = Nothing
  Set noticeCodeSheet = Nothing
  Set setSheet = Nothing
  Set mainSheet = Nothing
  Set tmpSheet = Nothing
  Set TeamsPlannerSheet = Nothing
  
  Set setVal = Nothing
  Set memberColor = Nothing

  
End Function
'***********************************************************************************************************************************************
' * 設定
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'***********************************************************************************************************************************************
Function setting(Optional reCheckFlg As Boolean)
  Dim line As Long
  
  
  On Error GoTo catchError

  If logFile <> "" And setVal("debugMode") = setSheet.Range("B3") And reCheckFlg <> True Then
    Exit Function
    Set setVal = Nothing
  End If

Label_reset:
  
  'ブックの設定
  Set ThisBook = ThisWorkbook
  ThisBook.Activate
  
  'ワークシート名の設定
  mainSheetName = "メイン"
  TeamsPlannerSheetName = "チームプランナー"
  Set helpSheet = ThisBook.Worksheets("Help")
  Set noticeCodeSheet = ThisBook.Worksheets("Notice")
  Set setSheet = ThisBook.Worksheets("設定")
  Set mainSheet = ThisBook.Worksheets(mainSheetName)
  Set tmpSheet = ThisBook.Worksheets("Tmp")
  Set TeamsPlannerSheet = ThisBook.Worksheets(TeamsPlannerSheetName)
  
  Set setVal = New Collection
  Set memberColor = CreateObject("Scripting.Dictionary")
  
  
  '期間、基準日が未入力時の初期値
  Select Case True
    Case setSheet.Range("B7") = ""
      setSheet.Range("B7") = Format(DateAdd("d", 0, Date), "yyyy/mm/dd")
    
    Case setSheet.Range("B8") = ""
      setSheet.Range("A8") = Format(DateAdd("d", 60, Date), "yyyy/mm/dd")
    
    Case setSheet.Range("B9") = ""
      setSheet.Range("B9") = Format(DateAdd("d", 0, Date), "yyyy/mm/dd")
  End Select
  
  '設定値の読み込み
  With setVal
    For line = 3 To setSheet.Cells(Rows.count, 1).End(xlUp).row
      If setSheet.Range("A" & line) <> "" Then
       .Add item:=setSheet.Range("B" & line), Key:=setSheet.Range("A" & line)
      End If
    Next
    For line = 3 To setSheet.Cells(Rows.count, 4).End(xlUp).row
      If setSheet.Range("D" & line) <> "" Then
       .Add item:=setSheet.Range("E" & line), Key:=setSheet.Range("D" & line)
      End If
    Next
  End With
  
  'ショートカットキーの設定追加
  With setVal
    For line = 3 To setSheet.Cells(Rows.count, 7).End(xlUp).row
      .Add item:=setSheet.Range("I" & line), Key:=setSheet.Range("H" & line)
    Next
  End With
  

  '担当者色読み込み
  For line = 3 To setSheet.Cells(Rows.count, 11).End(xlUp).row
    If setSheet.Range("K" & line).Value <> "" Then
      memberColor.Add setSheet.Range("K" & line).Value, setSheet.Range("K" & line).Interior.Color
    End If
  Next line

'  lineColor = setSheet.Range("E3").Interior.Color
'  SaturdayColor = setSheet.Range("E4").Interior.Color
'  SundayColor = setSheet.Range("E5").Interior.Color
'  CompanyHolidayColor = setSheet.Range("E6").Interior.Color
'
'  taskLevel1Color = setSheet.Range("E7").Interior.Color
'  taskLevel2Color = setSheet.Range("E8").Interior.Color
'  taskLevel3Color = setSheet.Range("E9").Interior.Color
  
  logFile = ThisBook.Path & "\ExcelMacro.log"
  
  If reCheckFlg = True Then
    Call Check.項目列チェック
    reCheckFlg = False
    Call clearSetting
    
    GoTo Label_reset
  End If
  
  Call 名前定義
  Exit Function
  
'エラー発生時=====================================================================================
catchError:
  logFile = ""
'  Set setVal = Nothing
'  Set setVal = New Collection
'
'  With setVal
'    .Add item:="ABC", Key:="debugMode"
'  End With

'  Call Library.showNotice(Err.Number, Err.Description, True)
  
  GoTo Label_reset
  
End Function

'***********************************************************************************************************************************************
' * 休日設定
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'***********************************************************************************************************************************************
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
  
'  On Error GoTo catchError

  For Each Name In Names
    If Name.Visible = False Then
      Name.Visible = True
    End If
    If Not Name.Name Like "*!Print_Area" And Not Name.Name Like "*!Print_Titles" Then
      Name.Delete
    End If
  Next
  
  For line = 3 To setSheet.Range("B5")
    If setSheet.Range("A" & line) <> "" Then
      setSheet.Range("B" & line).Name = setSheet.Range("A" & line)
    End If
  Next
  endLine = setSheet.Cells(Rows.count, 11).End(xlUp).row
  setSheet.Range(setVal("cell_AssignorList") & "3:" & setVal("cell_AssignorList") & endLine).Name = "担当者"

  endLine = setSheet.Cells(Rows.count, 17).End(xlUp).row
  setSheet.Range(setVal("cell_CompanyHoliday") & "3:" & setVal("cell_CompanyHoliday") & endLine).Name = "休日リスト"

  Exit Function
'エラー発生時=====================================================================================
catchError:
  Call Library.showNotice(Err.Number, Err.Description, True)
  
End Function


'***********************************************************************************************************************************************
' * シートの表示/非表示
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'***********************************************************************************************************************************************
Function noDispSheet()

  Call init.setting
  Worksheets("Help").Visible = xlSheetVeryHidden
  Worksheets("Tmp").Visible = xlSheetVeryHidden
  Worksheets("Notice").Visible = xlSheetVeryHidden
'  Worksheets("設定").Visible = xlSheetVeryHidden
  
  Worksheets(mainSheetName).Select
End Function



Function dispSheet()

  Call init.setting
  Worksheets("Help").Visible = True
  Worksheets("Tmp").Visible = True
  Worksheets("Notice").Visible = True
  Worksheets("設定").Visible = True
  
  Worksheets(TeamsPlannerSheetName).Visible = True
  Worksheets(mainSheetName).Visible = True
  
  Worksheets(mainSheetName).Select
  
End Function




































