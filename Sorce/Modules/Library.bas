Attribute VB_Name = "Library"
'**************************************************************************************************
' * 参照設定、定数宣言
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
' 利用する参照設定まとめ
' Microsoft Office 14.0 Object Library
' Microsoft DAO 3.6 Objects Library
' Microsoft Scripting Runtime (WSH, FileSystemObject)
' Microsoft ActiveX Data Objects 2.8 Library
' UIAutomationClient

' Windows APIの利用--------------------------------------------------------------------------------
' ディスプレイの解像度取得用
' Sleep関数の利用
' クリップボード関数の利用
#If VBA7 And Win64 Then
  Private Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
  Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As LongPtr)
  Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hWnd As LongPtr) As Long
  Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
  Private Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
#Else
  Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
  Private Declare Sub Sleep Lib "kernel32" (ByVal ms As Long)
  Declare Function OpenClipboard Lib "user32" (ByVal hWnd As Long) As Long
  Declare Function CloseClipboard Lib "user32" () As Long
  Declare Function EmptyClipboard Lib "user32" () As Long
#End If



'ワークブック用変数------------------------------
'ワークシート用変数------------------------------
'グローバル変数----------------------------------
Public LibDAO As String
Public LibADOX As String
Public LibADO As String
Public LibScript As String

'アクティブセルの取得
Dim SelectionCell As String

' PC、Office等の情報取得用連想配列
Public MachineInfo As Object

' Selenium用設定
Public Const HalfWidthDigit = "1234567890"
Public Const HalfWidthCharacters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
Public Const SymbolCharacters = "!""#$%&'()=~|@[`{;:]+*},./\<>?_-^\"

'Public Const JapaneseCharacters = "あいうえおかきくけこさしすせそたちつてとなにぬねのはひふへほまみむめもやゆよらりるれろわをんがぎぐげござじずぜぞだぢづでどばびぶべぼぱぴぷぺぽ"
'Public Const JapaneseCharactersCommonUse = "雨学空金青林画岩京国姉知長直店東歩妹明門夜委育泳岸苦具幸始使事実者昔取受所注定波板表服物放味命油和英果芽官季泣協径固刷参治周松卒底的典毒念府法牧例易往価河居券効妻枝舎述承招性制版肥非武沿延拡供呼刻若宗垂担宙忠届乳拝並宝枚依押奇祈拠況屈肩刺沼征姓拓抵到突杯泊拍迫彼怖抱肪茂炎欧殴"
'Public Const MachineDependentCharacters = "①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ㍉纊褜鍈銈蓜俉炻昱棈鋹曻彅丨仡仼伀伃伹佖侒侊侚侔俍偀倢俿倞偆偰偂傔"


Public ThisBook As Workbook


'**************************************************************************************************
' * アドオンを閉じる
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function addinClose()

  Workbooks(ThisWorkbook.Name).Close
End Function


'**************************************************************************************************
' * エラー時の処理
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function errorHandle(funcName As String, ByRef objErr As Object)
  
  Dim message As String
  Dim runTime As Date
  Dim endLine As Long
  
  runTime = Format(Now(), "yyyy/mm/dd hh:nn:ss")
  message = funcName & vbCrLf & objErr.Description

  '音声認識発話
  Application.Speech.Speak Text:="エラーが発生しました", SpeakAsync:=True
  message = Application.WorksheetFunction.VLookup(objErr.Number, noticeCodeSheet.Range("A2:B" & endLine), 2, False)
  
  Call MsgBox(message, vbCritical)
  Call endScript
  Call ProgressBar.showEnd
  
  Debug.Print objErr.Number
  Debug.Print objErr.Description
  Call outputLog(runTime & vbTab & objErr.Number & vbTab & objErr.Description)
  End
  
  
  
End Function

'**************************************************************************************************
' * 画面描写制御開始
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function startScript()

  'アクティブセルの取得
  If TypeName(Selection) = "Range" Then
    SelectionCell = Selection.Address
  End If

  'マクロ動作でシートやウィンドウが切り替わるのを見せないようにします
  Application.ScreenUpdating = False

  'マクロ動作自体で別のイベントが生成されるのを抑制する
  Application.EnableEvents = False

  'マクロ動作でセルItemNameなどが変わる時自動計算が処理を遅くするのを避ける
  Application.Calculation = xlCalculationManual

  'マクロ動作中に一切のキーやマウス操作を制限する
  'Application.Interactive = False

  'マクロ動作中はマウスカーソルを「砂時計」にする
  'Application.Cursor = xlWait

  '確認メッセージを出さない
  Application.DisplayAlerts = False

End Function

'**************************************************************************************************
' * 画面描写制御終了
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function endScript(Optional flg As Boolean = False)

  'マクロ動作でシートやウィンドウが切り替わるのを見せないようにします
  Application.ScreenUpdating = True

  'マクロ動作自体で別のイベントが生成されるのを抑制する
  Application.EnableEvents = True

  'マクロ動作でセルItemNameなどが変わる時自動計算が処理を遅くするのを避ける
  Application.Calculation = xlCalculationAutomatic

  'マクロ動作中に一切のキーやマウス操作を制限する
  'Application.Interactive = True

  'マクロ動作終了後はマウスカーソルを「デフォルト」にもどす
  'Application.Cursor = xlDefault

  'マクロ動作終了後はステータスバーを「デフォルト」にもどす
  Application.StatusBar = False
  'Application.StatusBar = "基準日:" & setVal("baseDay")

  '確認メッセージを出さない
  Application.DisplayAlerts = True

  '強制的に再計算させる
  Application.CalculateFull

 'アクティブセルの選択
  If SelectionCell <> "" And flg = True Then
    Range(SelectionCell).Select
  End If
  
  Call unsetClipboard
End Function



'**************************************************************************************************
' * シートの存在確認
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function chkSheetName(sheetName) As Boolean

  Dim tempSheet As Object
  Dim Result As Boolean

  Result = False
  For Each tempSheet In Sheets
    If LCase(sheetName) = LCase(tempSheet.Name) Then
      Result = True
      Exit For
    End If
  Next
  chkSheetName = Result

End Function


'**************************************************************************************************
' * 除外シート判定
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function chkExcludeSheet(sheetName As String, colNo As Long)

  Dim endBookRowLine As Long
  Dim rowLine As Long
  Dim Result As Boolean

  Result = True

  ' 最終行取得
  endBookRowLine = Sheets("設定").Cells(Rows.count, colNo).End(xlUp).row
  For rowLine = 3 To endBookRowLine
    If sheetName = Sheets("設定").Cells(rowLine, colNo) Then
      Result = False
      Exit For
    End If
  Next
  CheckExcludeSheet = Result
End Function


'**************************************************************************************************
' * ブックが開かれているかチェック
' *
' * @Link https://www.moug.net/tech/exvba/0060042.html
'**************************************************************************************************
Function chkBookOpened(chkFile) As Boolean

  Dim myChkBook As Workbook
  On Error Resume Next

  Set myChkBook = Workbooks(chkFile)

  If Err.Number > 0 Then
    chkBookOpened = False
  Else
    chkBookOpened = True
  End If
End Function

'**************************************************************************************************
' * ファイルの存在確認
' *
' * @Link http://officetanaka.net/excel/vba/filesystemobject/filesystemobject10.htm
'**************************************************************************************************
Function chkFileExists(targetPath As String)
  Dim FSO As Object
  
  Set FSO = CreateObject("Scripting.FileSystemObject")
  
  If FSO.FileExists(targetPath) Then
    chkFileExists = True
  Else
    chkFileExists = False
  End If
  Set FSO = Nothing
 
End Function


'**************************************************************************************************
' * ディレクトリの存在確認
' *
' * @Link http://officetanaka.net/excel/vba/filesystemobject/filesystemobject10.htm
'**************************************************************************************************
Function chkDirExists(targetPath As String)
  Dim FSO As Object
  
  Set FSO = CreateObject("Scripting.FileSystemObject")
  
  If FSO.FolderExists(targetPath) Then
    chkDirExists = True
  Else
    chkDirExists = False
  End If
  Set FSO = Nothing
 
End Function

'**************************************************************************************************
' * キャメルケースをスネークケースに変換
' *
' * @Link https://ameblo.jp/i-devdev-beginner/entry-12225328059.html
'**************************************************************************************************
Function covCamelToSnake(ByVal val As String, Optional ByVal isUpper As Boolean = False) As String
  Dim ret As String
  Dim i      As Long, length As Long
  
  length = Len(val)

  For i = 1 To length
    If UCase(Mid(val, i, 1)) = Mid(val, i, 1) Then
      If i = 1 Then
        ret = ret & Mid(val, i, 1)
      ElseIf i > 1 And UCase(Mid(val, i - 1, 1)) = Mid(val, i - 1, 1) Then
        ret = ret & Mid(val, i, 1)
      Else
        ret = ret & "_" & Mid(val, i, 1)
      End If
    Else
      ret = ret & Mid(val, i, 1)
    End If
  Next
  
  If isUpper Then
    covCamelToSnake = UCase(ret)
  Else
    covCamelToSnake = LCase(ret)
  End If
End Function


'**************************************************************************************************
' * スネークケースをキャメルケースに変換
' *
' * @Link https://ameblo.jp/i-devdev-beginner/entry-12225328059.html
'**************************************************************************************************
Function convSnakeToCamel(ByVal val As String, Optional ByVal isFirstUpper As Boolean = False) As String

    Dim ret As String
    Dim i   As Long
    Dim snakeSplit As Variant

    snakeSplit = Split(val, "_")

    For i = LBound(snakeSplit) To UBound(snakeSplit)
      ret = ret & UCase(Mid(snakeSplit(i), 1, 1)) & Mid(snakeSplit(i), 2, Len(snakeSplit(i)))
    Next

    If isFirstUpper Then
      convSnakeToCamel = ret
    Else
      convSnakeToCamel = LCase(Mid(ret, 1, 1)) & Mid(ret, 2, Len(ret))
    End If
End Function


'**************************************************************************************************
' * 半角のカタカナを全角のカタカナに変換する(ただし英数字は半角にする)
' *
' * @link   http://officetanaka.net/excel/function/tips/tips45.htm
'**************************************************************************************************
Function convHan2Zen(Text As String) As String
  Dim i As Long, buf As String

  Dim c As Range
  Dim rData As Variant, ansData As Variant

  For i = 1 To Len(Text)
    DoEvents
    rData = StrConv(Text, vbWide)
    If Mid(rData, i, 1) Like "[Ａ-ｚ]" Or Mid(rData, i, 1) Like "[０-９]" Or Mid(rData, i, 1) Like "－" Then
      ansData = ansData & StrConv(Mid(rData, i, 1), vbNarrow)
    Else
      ansData = ansData & Mid(rData, i, 1)
    End If
  Next i
  convHan2Zen = ansData
End Function


'**************************************************************************************************
' * パイプをカンマに変換
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function convPipe2Comma(strText As String) As String

  Dim covString As String

  tmp = Split(strText, "|")
  covString = ""
  For i = 0 To UBound(tmp)
    If i = 0 Then
      covString = tmp(i)
    Else
      covString = covString & "," & tmp(i)
    End If
  Next i
  convPipe2Comma = covString

End Function


'**************************************************************************************************
' * Base64エンコード(ファイル)
' *
' * @link   http://www.ka-net.org/blog/?p=4524
'**************************************************************************************************
Function convBase64EncodeForFile(ByVal filePath As String) As String
  Dim elm As Object
  Dim ret As String
  Const adTypeBinary = 1
  Const adReadAll = -1

  ret = "" '初期化
  On Error Resume Next
  Set elm = CreateObject("MSXML2.DOMDocument").createElement("base64")
  With CreateObject("ADODB.Stream")
    .Type = adTypeBinary
    .Open
    .LoadFromFile filePath
    elm.dataType = "bin.base64"
    elm.nodeTypedValue = .Read(adReadAll)
    ret = elm.Text
    .Close
  End With
  On Error GoTo 0
  convBase64EncodeForFile = ret
End Function


'**************************************************************************************************
' * Base64エンコード(文字列)
' *
' * @link   http://www.ka-net.org/blog/?p=4524
'**************************************************************************************************
Function convBase64EncodeForString(ByVal str As String) As String

  Dim ret As String
  Dim d() As Byte

  Const adTypeBinary = 1
  Const adTypeText = 2

  ret = "" '初期化
  On Error Resume Next
  With CreateObject("ADODB.Stream")
    .Open
    .Type = adTypeText
    .Charset = "UTF-8"
    .WriteText str
    .Position = 0
    .Type = adTypeBinary
    .Position = 3
    d = .Read()
    .Close
  End With
  With CreateObject("MSXML2.DOMDocument").createElement("base64")
    .dataType = "bin.base64"
    .nodeTypedValue = d
    ret = .Text
  End With
  On Error GoTo 0
  convBase64EncodeForString = ret
End Function


'**************************************************************************************************
' * URL-safe Base64エンコード
' *
' * @link   http://www.ka-net.org/blog/?p=4524
'**************************************************************************************************
Function convURLSafeBase64Encode(ByVal str As String) As String

  str = convBase64EncodeForString(str)
  str = Replace(str, "+", "-")
  str = Replace(str, "/", "_")

  convURLSafeBase64Encode = str
End Function


'**************************************************************************************************
' * URLエンコード
' *
' * @link   http://www.ka-net.org/blog/?p=4524
'**************************************************************************************************
Function convURLEncode(ByVal str As String) As String

  With CreateObject("ScriptControl")
    .Language = "JScript"
    EncodeURL = .codeobject.encodeURIComponent(str)
  End With
End Function


'**************************************************************************************************
' * 先頭１文字目を大文字化
' *
' * @Link http://bekkou68.hatenablog.com/entry/20090414/1239685179
'**************************************************************************************************
Function convFirstCharConvert(ByVal strTarget As String) As String
    Dim strFirst As String
    Dim strExceptFirst As String

    strFirst = UCase(Left$(strTarget, 1))
    strExceptFirst = Mid$(strTarget, 2, Len(strTarget))
    convFirstCharConvert = strFirst & strExceptFirst
End Function


'**************************************************************************************************
' * 文字列の左側から指定文字数削除する関数
' *
' * @Link   https://vbabeginner.net/vbaで文字列の右側や左側から指定文字数削除する/
'**************************************************************************************************
Function cutLeft(s, i As Long) As String
    Dim iLen    As Long     '文字列長

    '文字列ではない場合
    If VarType(s) <> vbString Then
        cutLeft = s & "文字列ではない"
        Exit Function
    End If

    iLen = Len(s)

    '文字列長より指定文字数が大きい場合
    If iLen < i Then
        cutLeft = s & "文字列長より指定文字数が大きい"
        Exit Function
    End If

    '指定文字数を削除して返す
    cutLeft = Right(s, iLen - i)
End Function


'**************************************************************************************************
' * 文字列の右側から指定文字数削除する関数
' *
' * @Link   https://vbabeginner.net/vbaで文字列の右側や左側から指定文字数削除する/
'**************************************************************************************************
Function cutRight(s, i As Long) As String
    Dim iLen    As Long     '文字列長

    '文字列ではない場合
    If VarType(s) <> vbString Then
      cutRight = s & "文字列ではない"
      Exit Function
    End If

    iLen = Len(s)

    '文字列長より指定文字数が大きい場合
    If iLen < i Then
      cutRight = s & "文字列長より指定文字数が大きい"
      Exit Function
    End If

    '指定文字数を削除して返す
    cutRight = Left(s, iLen - i)
End Function


'**************************************************************************************************
' * 連続改行の削除
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function delMultipleLine(targetValue As String)
  With CreateObject("VBScript.RegExp")
    .Global = True
    .Pattern = "(\r\n)+"
    combineMultipleLine = .Replace(targetValue, vbCrLf)
  End With
End Function

'**************************************************************************************************
' *レジストリから情報取得
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function delRegistry(registryName As String)
  Dim regVal As String
  
  On Error Resume Next
  
  If registryName = "" Then
    DeleteSetting "ExcelHelp", RegistrySubKey
  Else
    DeleteSetting "ExcelHelp", RegistrySubKey, registryName
  End If
  
End Function
'**************************************************************************************************
' * シート削除
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function delSheetData(Optional line = 1)

  If IsNumeric(line) Then
    Rows(line & ":" & Rows.count).Delete Shift:=xlUp
  Else
    Cells.Delete Shift:=xlUp
  End If
  DoEvents
  Cells.NumberFormatLocal = "G/標準"

  Application.Goto Reference:=Range("A1"), Scroll:=True
End Function


'**************************************************************************************************
' * 選択範囲の画像削除
' *
' * @Link https://www.relief.jp/docs/018407.html
'**************************************************************************************************
Function delImage()
  Dim rng As Range
  Dim shp As Shape

  If TypeName(Selection) <> "Range" Then
    Exit Function
  End If

  For Each shp In ActiveSheet.Shapes
    Set rng = Range(shp.TopLeftCell, shp.BottomRightCell)

    If Not (Intersect(rng, Selection) Is Nothing) Then
      shp.Delete
    End If
  Next
End Function


'**************************************************************************************************
' * セルの名称設定削除
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function delVisibleNames()
  Dim Name As Object

  On Error Resume Next

  For Each Name In Names
    If Name.Visible = False Then
      Name.Visible = True
    End If
    If Not Name.Name Like "*!Print_Area" And Not Name.Name Like "*!Print_Titles" Then
      Name.Delete
    End If
  Next

End Function


'**************************************************************************************************
' * テーブルデータ削除
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function delTableData()
  Dim endLine As Long

  On Error Resume Next

  endLine = Cells(Rows.count, 1).End(xlUp).row
  Rows("3:" & endLine).Select
  Selection.Delete Shift:=xlUp

  Rows("2:3").Select
  Selection.SpecialCells(xlCellTypeConstants, 23).ClearContents

  Cells.Select
  Selection.NumberFormatLocal = "G/標準"

  Application.Goto Reference:=Range("A1"), Scroll:=True
End Function


'**************************************************************************************************
' * PC、Office等の情報取得
' * 連想配列を利用しているので、Microsoft Scripting Runtimeが必須
' * MachineInfo.Item ("Excel") で呼び出し
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function getMachineInfo() As Object
  Dim WshNetworkObject As Object

  Set MachineInfo = CreateObject("Scripting.Dictionary")
  Set WshNetworkObject = CreateObject("WScript.Network")

  ' OSのバージョン取得-----------------------------------------------------------------------------
  Select Case Application.OperatingSystem

    Case "Windows (64-bit) NT 6.01"
        MachineInfo.Add "OS", "Windows7-64"

    Case "Windows (32-bit) NT 6.01"
        MachineInfo.Add "OS", "Windows7-32"

    Case "Windows (32-bit) NT 5.01"
        MachineInfo.Add "OS", "WindowsXP-32"

    Case "Windows (64-bit) NT 5.01"
        MachineInfo.Add "OS", "WindowsXP-64"

    Case Else
       MachineInfo.Add "OS", Application.OperatingSystem
  End Select

  ' Excelのバージョン取得--------------------------------------------------------------------------
  Select Case Application.Version
    Case "16.0"
        MachineInfo.Add "Excel", "2016"
    Case "14.0"
        MachineInfo.Add "Excel", "2010"
    Case "12.0"
        MachineInfo.Add "Excel", "2007"
    Case "11.0"
        MachineInfo.Add "Excel", "2003"
    Case "10.0"
        MachineInfo.Add "Excel", "2002"
    Case "9.0"
        MachineInfo.Add "Excel", "2000"
    Case Else
       MachineInfo.Add "Excel", Application.Version
  End Select

  'PCの情報----------------------------------------------------------------------------------------
  MachineInfo.Add "UserName", WshNetworkObject.UserName
  MachineInfo.Add "ComputerName", WshNetworkObject.ComputerName
  MachineInfo.Add "UserDomain", WshNetworkObject.UserDomain

  '画面の解像度等取得------------------------------------------------------------------------------
  MachineInfo.Add "monitors", GetSystemMetrics(80)
  MachineInfo.Add "displayX", GetSystemMetrics(0)
  MachineInfo.Add "displayY", GetSystemMetrics(1)
  MachineInfo.Add "displayVirtualX", GetSystemMetrics(78)
  MachineInfo.Add "displayVirtualY", GetSystemMetrics(79)
  MachineInfo.Add "appTop", ActiveWindow.top
  MachineInfo.Add "appLeft", ActiveWindow.Left
  MachineInfo.Add "appWidth", ActiveWindow.Width
  MachineInfo.Add "appHeight", ActiveWindow.Height
End Function


'**************************************************************************************************
' * 文字数カウント
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function getByteString(arryColumn As String, Optional line As Long) As Long
  Dim colLineName As Variant
  Dim count As Long
  
  count = 0
  For Each colLineName In Split(arryColumn, ",")
    If line > 0 Then
      count = count + LenB(Range(colLineName & line).Value)
    Else
      count = count + LenB(Range(colLineName).Value)
    End If
  Next colLineName

  getByteString = count
End Function


'**************************************************************************************************
' * 列名から列番号を求める
' *
' * @link   http://www.happy2-island.com/excelsmile/smile03/capter00717.shtml
'**************************************************************************************************
Function getColumnNo(targetCell As String) As Long

  getColumnNo = Range(targetCell & ":" & targetCell).Column
End Function


'**************************************************************************************************
' * 列番号から列名を求める
' *
' * @link   http://www.happy2-island.com/excelsmile/smile03/capter00717.shtml
'**************************************************************************************************
Function getColumnName(targetCell As Long) As String

  getColumnName = Split(Cells(, targetCell).Address, "$")(1)
End Function

'**************************************************************************************************
' * カラーパレットを表示し、色コードを取得
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function getColor(colorValue As Long)
  Dim Red As Long, Green As Long, Blue As Long
  Dim setColorValue As Long
  
  Call getRGB(colorValue, Red, Green, Blue)
  setColorValue = Application.Dialogs(xlDialogEditColor).Show(10, Red, Green, Blue)
  If setColorValue = False Then
    setColorValue = colorValue
  Else
    setColorValue = ThisWorkbook.Colors(10)
  End If
  
  getColor = setColorValue

End Function


'**************************************************************************************************
' * IndentLevel値取得
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Public Function getIndentLevel(targetRange As Range)
  Dim thisTargetSheet As Worksheet
  
  Application.Volatile
  getIndentLevel = ""

  If targetRange = "" Then
    getIndentLevel = ""
  Else
    getIndentLevel = targetRange.IndentLevel + 1
  End If
End Function


'**************************************************************************************************
' * RGB値取得
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function getRGB(colorValue As Long, Red As Long, Green As Long, Blue As Long)
  Red = colorValue Mod 256
  Green = Int(colorValue / 256) Mod 256
  Blue = Int(colorValue / 256 / 256)
End Function


'**************************************************************************************************
' *レジストリから情報取得
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function getRegistry(registryName As String)
  Dim regVal As String

  regVal = GetSetting("ExcelHelp", RegistrySubKey, registryName)
  getRegistry = regVal
End Function


'**************************************************************************************************
' * ディレクトリ選択ダイアログ表示
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function getDirPath(CurrentDirectory As String, Optional title As String)

  With Application.FileDialog(msoFileDialogFolderPicker)
    .InitialFileName = CurrentDirectory & "\"
    .AllowMultiSelect = False
    .title = title & "の保存場所を選択してください"
    If .Show = True Then
      getDirPath = .SelectedItems(1)
    Else
      getDirPath = ""
    End If
  End With
End Function


'**************************************************************************************************
' * ファイル保存ダイアログ表示
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function getSaveFilePath(CurrentDirectory As String, saveFileName As String, FileTypeNo As Long)

  Dim filePath As String
  Dim Result As Long

  Dim FileName As Variant
  FileName = Application.GetSaveAsFilename( _
      InitialFileName:=CurrentDirectory & "\" & saveFileName, _
      FileFilter:="Excelファイル,*.xlsx,Excel2003以前,*.xls,Excelマクロブック,*.xlsm", _
      FilterIndex:=FileTypeNo)

  getSaveFilePath = FileName
End Function

'**************************************************************************************************
' * ファイル選択ダイアログ表示
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function getFilePath(CurrentDirectory As String, saveFileName As String, title As String, FileTypeNo As Long)

  Dim filePath As String
  Dim Result As Long

  With Application.FileDialog(msoFileDialogFilePicker)

    ' ファイルの種類を設定
    .Filters.Clear
    .Filters.Add "Excelブック", "*.xls; *.xlsx; *.xlsm"
    .Filters.Add "CSVファイル", "*.csv"
    .Filters.Add "SQLファイル", "*.sql"
    .Filters.Add "テキストファイル", "*.txt"
    .Filters.Add "JSONファイル", "*.json"
    .Filters.Add "Accesssデータベース", "*.mdb"
    .Filters.Add "すべてのファイル", "*.*"

    .FilterIndex = FileTypeNo

    '表示するフォルダ
    .InitialFileName = CurrentDirectory & "\" & saveFileName

    '表示形式の設定
    .InitialView = msoFileDialogViewWebView

    'ダイアログ ボックスのタイトル設定
    .title = title
    
    
    If .Show = -1 Then
      filePath = .SelectedItems(1)
    Else
      filePath = ""
    End If
  End With

  getFilePath = filePath

End Function


'**************************************************************************************************
' * 複数ファイル選択ダイアログ表示
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function getFilesPath(CurrentDirectory As String, saveFileName As String, title As String, FileTypeNo As Long)

  Dim filePath() As Variant
  Dim Result As Long

  With Application.FileDialog(msoFileDialogFilePicker)
    '複数選択を許可
    .AllowMultiSelect = True
    
    ' ファイルの種類を設定
    .Filters.Clear
    .Filters.Add "Excelブック", "*.xls; *.xlsx; *.xlsm"
    .Filters.Add "CSVファイル", "*.csv"
    .Filters.Add "SQLファイル", "*.sql"
    .Filters.Add "テキストファイル", "*.txt"
    .Filters.Add "JSONファイル", "*.json"
    .Filters.Add "Accesssデータベース", "*.mdb"
    .Filters.Add "すべてのファイル", "*.*"

    .FilterIndex = FileTypeNo

    '表示するフォルダ
    .InitialFileName = CurrentDirectory & "\" & saveFileName

    '表示形式の設定
    .InitialView = msoFileDialogViewWebView
    
    'ダイアログ ボックスのタイトル設定
    .title = title


    If .Show = -1 Then
      ReDim Preserve filePath(.SelectedItems.count - 1)
      For i = 1 To .SelectedItems.count
        filePath(i - 1) = .SelectedItems(i)
      Next i
    Else
      ReDim Preserve filePath(0)
      filePath(0) = ""
    End If
  End With

  getFilesPath = filePath

End Function

'**************************************************************************************************
' * ディレクトリ内のファイル一覧取得
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function getFileList(Path As String, FileName As String)
  Dim f As Object, cnt As Long
  Dim list() As String
  
  cnt = 0
  With CreateObject("Scripting.FileSystemObject")
    For Each f In .GetFolder(Path).Files
      If f.Name Like FileName Then
        ReDim Preserve list(cnt)
        list(cnt) = f.Name
        cnt = cnt + 1
      End If
    Next f
  End With
  
  getFileList = list
End Function




'**************************************************************************************************
' * 指定バイト数の固定長データ作成(文字列処理)
' *
' * @Link http://www.asahi-net.or.jp/~ef2o-inue/vba_o/sub05_110_055.html
'**************************************************************************************************
Function getFixlng(strInText As String, lngFixBytes As Long) As String
    Dim lngKeta As Long
    Dim lngByte As Long, lngByte2 As Long, lngByte3 As Long
    Dim IX As Long
    Dim intCHAR As Long
    Dim strOutText As String

    lngKeta = Len(strInText)
    strOutText = strInText
    ' バイト数判定
    For IX = 1 To lngKeta
        ' 1文字ずつ半角/全角を判断
        intCHAR = Asc(Mid(strInText, IX, 1))
        ' 全角と判断される場合はバイト数に1を加える
        If ((intCHAR < 0) Or (intCHAR > 255)) Then
            lngByte2 = 2        ' 全角
        Else
            lngByte2 = 1        ' 半角
        End If
        ' 桁あふれ判定(右切り捨て)
        lngByte3 = lngByte + lngByte2
        If lngByte3 >= lngFixBytes Then
            If lngByte3 > lngFixBytes Then
                strOutText = Left(strInText, IX - 1)
            Else
                strOutText = Left(strInText, IX)
                lngByte = lngByte3
            End If
            Exit For
        End If
        lngByte = lngByte3
    Next IX
    ' 桁不足判定(空白文字追加)
    If lngByte < lngFixBytes Then
        strOutText = strOutText & Space(lngFixBytes - lngByte)
    End If
    getFixlng = strOutText
End Function


'**************************************************************************************************
' * シートリスト取得
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function getSheetList(ColumnName As String)

  Dim i As Long
  Dim sheetName As Object

  i = 3
  If ColumnName = "" Then
    ColumnName = "E"
  End If

  On Error GoTo GetSheetListError:
  Call startScript

  '現設定値のクリア
  Worksheets("設定").Range(ColumnName & "3:" & ColumnName & "100").Select
  Selection.Borders(xlDiagonalDown).LineStyle = xlNone
  Selection.Borders(xlDiagonalUp).LineStyle = xlNone
  Selection.Borders(xlEdgeLeft).LineStyle = xlNone
  Selection.Borders(xlEdgeTop).LineStyle = xlNone
  Selection.Borders(xlEdgeBottom).LineStyle = xlNone
  Selection.Borders(xlEdgeRight).LineStyle = xlNone
  Selection.Borders(xlInsideVertical).LineStyle = xlNone
  Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
  With Selection.Interior
    .Pattern = xlSolid
    .PatternColorIndex = xlAutomatic
    .Color = xlNone
    .TintAndShade = 0
    .PatternTintAndShade = 0
  End With

  For Each sheetName In ActiveWorkbook.Sheets

    'シート名の設定
    Worksheets("設定").Range(ColumnName & i).Select
    Worksheets("設定").Range(ColumnName & i) = sheetName.Name

    ' セルの背景色解除
    With Worksheets("設定").Range(ColumnName & i).Interior
      .Pattern = xlPatternNone
      .Color = xlNone
    End With

    ' シート色と同じ色をセルに設定
    If Worksheets(sheetName.Name).Tab.Color Then
      With Worksheets("設定").Range(ColumnName & i).Interior
        .Pattern = xlPatternNone
        .Color = Worksheets(sheetName.Name).Tab.Color
      End With
    End If

    '罫線の設定
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
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
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone

    i = i + 1
  Next

  Worksheets("設定").Range(ColumnName & "3").Select
  Call endScript
  Exit Function
'--------------------------------------------------------------------------------------------------
'エラー発生時の処理
'--------------------------------------------------------------------------------------------------
GetSheetListError:

  ' 画面描写制御終了
  Call endScript
  Call errorHandle("シートリスト取得", Err)

End Function


'**************************************************************************************************
' * 選択セルの拡大表示呼出
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function showExpansionForm(Text As String, SetSelectTargetRows As String)
  With ExpansionForm
    .StartUpPosition = 0
    .top = Application.top + (ActiveWindow.Width / 10)
    .Left = Application.Left + (ActiveWindow.Height / 5)
    .TextBox = Text
    .TextBox.MultiLine = True
    .TextBox.MultiLine = True
    .TextBox.EnterKeyBehavior = True
    .Caption = SetSelectTargetRows
  End With
  ExpansionForm.Show vbModeless
End Function


'**************************************************************************************************
' * 選択セルの拡大表示終了
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function showExpansionFormClose(Text As String, SetSelectTargetRows As String)
  Range(SetSelectTargetRows).Value = Text
  Call endScript
End Function


'**************************************************************************************************
' * デバッグ用画面表示
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function showDebugForm(meg1 As String, Optional meg2 As String)
  Dim runTime As Date
  Dim StartUpPosition As Long
  
  runTime = Format(Now(), "yyyy/mm/dd hh:nn:ss")
  
  If setVal("debugMode") = "none" Then
    Exit Function
  End If

  If meg1 <> "" And Len(meg1) < 10 Then
    
    meg1 = meg1 & String(10 - Len(meg1), "　")
  End If
  
  Select Case setVal("debugMode")
    Case "file"
      If meg1 <> "" Then
        Call outputLog(runTime & vbTab & meg1 & vbTab & meg2)
      End If
      GoTo label_end
    Case "form"
      GoTo label_showForm
    Case "all"
      If meg1 <> "" Then
        Call outputLog(runTime & vbTab & meg1 & vbTab & meg2)
      End If
      GoTo label_showForm
      
    Case Else
      Exit Function
  End Select

label_showForm:
  If meg2 = "処理開始" Then
    With debugForm
      .Caption = "処理情報"
      .ListBox1.Clear
      .ListBox1.AddItem runTime & vbTab & meg1 & vbTab & meg2

    End With
  Else
    With debugForm
      .Caption = "処理情報"
      .ListBox1.AddItem runTime & vbTab & meg1 & vbTab & meg2
      .ListBox1.ListIndex = .ListBox1.ListCount - 1
    End With
  End If
  
  If (debugForm.Visible = True) Then
  debugForm.StartUpPosition = 0
  Else
  debugForm.StartUpPosition = 1
  End If
  debugForm.Show vbModeless


label_end:

  DoEvents
End Function

'**************************************************************************************************
' * 処理情報通知
' *
' * Worksheets("Notice").Visible = True
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function showNotice(code As Long, Optional process As String, Optional runEndflg As Boolean)
'  Dim noticeCodeSheet As Worksheet
'  Set noticeCodeSheet = ThisBook.Worksheets("NoticeCode")
  
  Dim message As String
  Dim runTime As Date
  Dim endLine As Long
  
  runTime = Format(Now(), "yyyy/mm/dd hh:nn:ss")
  
  endLine = noticeCodeSheet.Cells(Rows.count, 1).End(xlUp).row
  message = Application.WorksheetFunction.VLookup(code, noticeCodeSheet.Range("A2:B" & endLine), 2, False)
  
  If process <> "" Then
    message = Replace(message, "%%", process)
  End If

  If setVal("debugMode") = "speak" Or setVal("debugMode") = "all" Then
    Application.Speech.Speak Text:=message, SpeakAsync:=True, SpeakXML:=True
  End If
  
  Select Case code
    Case 0 To 399
      Call MsgBox(message, vbInformation, thisAppName)
    
    Case 400 To 499
      Call MsgBox(message, vbCritical, thisAppName)
    
    Case 500 To 599
      Call MsgBox(message, vbExclamation, thisAppName)

    Case Else
      Call MsgBox(message, vbCritical, thisAppName)
  End Select
'  Stop

  '画面描写制御終了処理
  If runEndflg = True Then
    Call endScript
    Call ProgressBar.showEnd
    End
  End If
End Function


'**************************************************************************************************
' * ランダム
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function makeRandomString(ByVal setString As String, ByVal setStringCnt As Integer) As String

 For i = 1 To setStringCnt
    '乱数ジェネレータを初期化
    Randomize
    n = Int((Len(setString) - 1 + 1) * Rnd + 1)
    str1 = str1 + Mid(setString, n, 1)
  Next i

  makeRandomString = str1

End Function

Function makeRandomNo(minNo As Long, maxNo As Long) As String

  '乱数ジェネレータを初期化
  Randomize
  makeRandomNo = Application.RoundDown(Int((maxNo - minNo + 1) * Rnd + minNo), -2)

End Function

'**************************************************************************************************
' * ログ出力
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function outputLog(message As String)
  Dim fileTimestamp As Date
  
  If chkFileExists(logFile) Then
    fileTimestamp = FileDateTime(logFile)
  Else
      fileTimestamp = DateAdd("d", -1, Date)
  End If
  
  If Format(Date, "yyyymmdd") = Format(fileTimestamp, "yyyymmdd") Then
    Open logFile For Append As #1
  Else
    Open logFile For Output As #1
  End If
  
  
  'Print #1, "[" & Format(Now, "YYYY/MM/DD hh:mm:ss") & "] " & Replace(Message, vbLf, " ")
  Print #1, Replace(message, vbLf, " ")
  
  Close #1
End Function

Function outputText(message As String, outputFilePath)

  Open outputFilePath For Output As #1
  Print #1, message
  Close #1

End Function

'**************************************************************************************************
' * CSVインポート
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
' * @link   https://www.tipsfound.com/vba/18014
'**************************************************************************************************
Function importCsv(filePath As String, Optional readLine As Long, Optional TextFormat As Variant)

  Dim ws As Worksheet
  Dim qt As QueryTable
  Dim count As Long, line As Long, endLine As Long

  endLine = Cells(Rows.count, 1).End(xlUp).row
  If endLine = 1 Then
    endLine = 1
  Else
    endLine = endLine + 1
  End If
  
  If readLine < 1 Then
    readLine = 1
  End If
  
  Set ws = ActiveSheet
  Set qt = ws.QueryTables.Add(Connection:="TEXT;" & filePath, Destination:=ws.Range("A" & endLine))
  With qt
    .TextFilePlatform = 932          ' Shift-JIS を開く
    .TextFileParseType = xlDelimited ' 文字で区切った形式
    .TextFileCommaDelimiter = True   ' 区切り文字はカンマ
    .TextFileStartRow = readLine     ' 1 行目から読み込み
    .AdjustColumnWidth = False       ' 列幅を自動調整しない
    .RefreshStyle = xlOverwriteCells '上書きを指定
    .TextFileTextQualifier = xlTextQualifierDoubleQuote ' 引用符の指定
    
    If IsArray(TextFormat) Then
      .TextFileColumnDataTypes = TextFormat
    End If
    
    .Refresh
    DoEvents
    .Delete
  End With
  Set qt = Nothing
  Set ws = Nothing
  
  Call Library.startScript
End Function


'**************************************************************************************************
' * Excelファイルのインポート
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function importXlsx(filePath As String, targeSheet As String, targeArea As String, dictSheet As Worksheet, Optional passWord As String)

  If passWord <> "" Then
    Workbooks.Open FileName:=filePath, ReadOnly:=True, passWord:=passWord
  Else
    Workbooks.Open FileName:=filePath, ReadOnly:=True
  End If
  
  If Worksheets(targeSheet).Visible = False Then
    Worksheets(targeSheet).Visible = True
  End If
  Sheets(targeSheet).Select

  ActiveWorkbook.Sheets(targeSheet).Rows.Hidden = False
  ActiveWorkbook.Sheets(targeSheet).Columns.Hidden = False
  
  If ActiveSheet.FilterMode Then ActiveSheet.ShowAllData
  
  ActiveWorkbook.Sheets(targeSheet).Range(targeArea).Copy
  dictSheet.Range("A1").PasteSpecial xlPasteValues
  
  Application.CutCopyMode = False
  ActiveWorkbook.Close SaveChanges:=False
  dictSheet.Range("A1").Select
  DoEvents
  Call unsetClipboard
  Call Library.startScript

End Function


'**************************************************************************************************
' * MkDirで階層の深いフォルダーを作る
' *
' * @link https://www.relief.jp/docs/excel-vba-mkdir-folder-structure.html
'**************************************************************************************************
Function makeDir(fullPath As String)
  Dim tmpPath As String, arr() As String
  Dim i As Long

  If chkDirExists(fullPath) Then
    Exit Function
  End If
  
  arr = Split(fullPath, "\")
  tmpPath = arr(0)  ' ドライブ名の代入

  For i = 1 To UBound(arr)
    tmpPath = tmpPath & "\" & arr(i)
    If Dir(tmpPath, vbDirectory) = "" Then
      MkDir tmpPath
    End If
  Next i

End Function


'**************************************************************************************************
' * パスワード生成
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function makePasswd() As String
  Dim halfChar As String, str1 As String

  halfChar = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz!#$%&"

  For i = 1 To 12
    '乱数ジェネレータを初期化
    Randomize
    n = Int((Len(halfChar) - 1 + 1) * Rnd + 1)
    str1 = str1 + Mid(halfChar, n, 1)
  Next i
  makePasswd = str1
End Function


'**************************************************************************************************
' * 文字列分割
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function splitString(targetString As String, separator As String, count As Integer) As String
  Dim tmp As Variant

  If targetString <> "" Then
    tmp = Split(targetString, separator)
    splitString = tmp(count)
  Else
    splitString = ""
  End If
End Function


'**************************************************************************************************
' * 配列の最後に追加する
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function setArrayPush(arrName As Variant, str As Variant)
  Dim i As Long
  
  i = UBound(arrName)
  If i = 0 Then
  
  Else
    i = i + 1
    ReDim Preserve arrName(i)
  End If
  arrName(i) = str

End Function


'**************************************************************************************************
' * フォントカラー設定
' *
' * @Link https://vbabeginner.net/vbaでセルの指定文字列の色や太さを変更する/
'**************************************************************************************************
Function setFontClor(a_sSearch, a_lColor, a_bBold)
  Dim f   As Font     'Fontオブジェクト
  Dim i               '引数文字列のセルの位置
  Dim iLen            '引数文字列の文字数
  Dim r   As Range    'セル範囲の１セル

  iLen = Len(a_sSearch)
  i = 1

  For Each r In Selection
    Do
      i = InStr(i, r.Value, a_sSearch)
      If (i = 0) Then
        i = 1
        Exit Do
      End If
      Set f = r.Characters(i, iLen).Font
      f.Color = a_lColor
      f.Bold = a_bBold
      i = i + 1
    Loop
  Next
End Function


'**************************************************************************************************
' * レジストリに情報登録
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function setRegistry(registryName As String, setVal As Variant)
  
  If getRegistry(registryName) <> setVal Then
    SaveSetting "ExcelHelp", RegistrySubKey, registryName, setVal
  End If
End Function


'**************************************************************************************************
' * 参照設定を自動で行う
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function setReferences(BookType As String)

  On Error GoTo Err_SetReferences:

  'Microsoft Scripting Runtime (Windows Script Host / FileSystemObject)----------------------------
    LibScript = "C:\Windows\System32\scrrun.dll"
    If Dir(LibScript) <> "" Then
      ActiveWorkbook.VBProject.References.AddFromFile (LibScript)
    Else
      MsgBox ("Microsoft Scripting Runtimeを利用できません。" & vbLf & "利用できない機能があります")
    End If
    
  'Microsoft ActiveX Data Objects Library 6.1 (ADO)------------------------------------------------
  If BookType = "DataBase" Then
    LibADO = "C:\Program Files\Common Files\System\Ado\msado15.dll"
    If Dir(LibADO) <> "" Then
      ActiveWorkbook.VBProject.References.AddFromFile (LibADO)
    Else
      MsgBox ("Microsoft ActiveX Data Objectsを利用できません" & vbLf & "利用できない機能があります")
    End If

  'Microsoft DAO 3.6 Objects Library (Database Access Object)--------------------------------------
  LibDAO = "C:\Program Files\Common Files\Microsoft Shared\DAO\dao360.dll"
    If Dir(LibDAO) <> "" Then
      ActiveWorkbook.VBProject.References.AddFromFile (LibDAO)
    Else
      LibDAO = "C:\Program Files (x86)\Common Files\microsoft shared\DAO\dao360.dll"
      If Dir(LibDAO) <> "" Then
        ActiveWorkbook.VBProject.References.AddFromFile (LibDAO)
      Else
        MsgBox ("Microsoft DAO 3.6 Objects Libraryを利用できません" & vbLf & "DBへの接続機能が利用できません")
      End If
    End If
  End If

  'Microsoft DAO 3.6 Objects Library (Database Access Object)--------------------------------------
  If BookType = "" Then
    LibDAO = "C:\Program Files\Common Files\Microsoft Shared\DAO\dao360.dll"
    If Dir(LibDAO) <> "" Then
      ActiveWorkbook.VBProject.References.AddFromFile (LibDAO)
    Else
      LibDAO = "C:\Program Files (x86)\Common Files\microsoft shared\DAO\dao360.dll"
      If Dir(LibDAO) <> "" Then
        ActiveWorkbook.VBProject.References.AddFromFile (LibDAO)
      Else
        MsgBox ("Microsoft DAO 3.6 Objects Libraryを利用できません" & vbLf & "DBへの接続機能が利用できません")
      End If
    End If
  End If


Func_Exit:
  Set Ref = Nothing
  Exit Function

Err_SetReferences:
  If Err.Number = 32813 Then
    Resume Next
  ElseIf Err.Number = 1004 Then
    MsgBox ("「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」に変更を！")
  Else
    MsgBox "Error Number : " & Err.Number & vbCrLf & Err.Description
    GoTo Func_Exit:
  End If
End Function


'**************************************************************************************************
' * 選択セルの行背景設定
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function setLineColor(SetArea As String, DisType As Boolean, SetColor As String)

  Range(SetArea).Select

  '条件付き書式をクリア
  Selection.FormatConditions.Delete

  If DisType = False Then
    '行だけ設定
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:="=CELL(""row"")=ROW()"
  Else
    '行と列に設定
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:="=OR(CELL(""row"")=ROW(), CELL(""col"")=COLUMN())"
  End If

  Selection.FormatConditions(Selection.FormatConditions.count).SetFirstPriority
  With Selection.FormatConditions(1)
    .Interior.PatternColorIndex = xlAutomatic
    .Interior.Color = SetColor
    .Interior.TintAndShade = 0
    .Font.ColorIndex = 1
  End With
  Selection.FormatConditions(1).StopIfTrue = False
'  Application.GoTo Reference:=Range("A1"), Scroll:=True
End Function


'**************************************************************************************************
' * 最初のシートを選択
' *
' * @Link https://www.relief.jp/docs/excel-vba-select-1st-visible%20-sheet.html
'**************************************************************************************************
Function setFirstsheet()
  Dim i As Long

  For i = 1 To Sheets.count
    If Sheets(i).Visible = xlSheetVisible Then
      Sheets(i).Select
      Exit Function
    End If
  Next i
End Function


'**************************************************************************************************
' * ファイル全体の文字列置換
' *
' * @Link   https://www.moug.net/tech/acvba/0090005.html
'**************************************************************************************************
Function replaceFromFile(FileName As String, TargetText As String, Optional NewText As String = "")

 Dim FSO         As FileSystemObject 'ファイルシステムオブジェクト
 Dim Txt         As TextStream       'テキストストリームオブジェクト
 Dim buf_strTxt  As String           '読み込みバッファ

 On Error GoTo Func_Err:

 'オブジェクト作成
 Set FSO = CreateObject("Scripting.FileSystemObject")
 Set Txt = FSO.OpenTextFile(FileName, ForReading)

 '全文読み込み
  buf_strTxt = Txt.ReadAll
  Txt.Close

  '元ファイルをリネームして、テンポラリファイル作成
  Name FileName As FileName & "_"

  '置換処理
   buf_strTxt = Replace(buf_strTxt, TargetText, NewText, , , vbBinaryCompare)

  '書込み用テキストファイル作成
   Set Txt = FSO.CreateTextFile(FileName, True)
  '書込み
  Txt.Write buf_strTxt
  Txt.Close

  'テンポラリファイルを削除
  FSO.DeleteFile FileName & "_"

'終了処理
Func_Exit:
    Set Txt = Nothing
    Set FSO = Nothing
    Exit Function

Func_Err:
    MsgBox "Error Number : " & Err.Number & vbCrLf & Err.Description
    GoTo Func_Exit:
End Function


'**************************************************************************************************
' * VBAでExcelのコメントを一括で自動サイズにしてカッコよくする
' *
' * @Link   http://techoh.net/customize-excel-comment-by-vba/
'**************************************************************************************************
Function resetComment()
    Dim cl As Range
    Dim count As Long

    count = 0
    For Each cl In Selection
      count = count + 1
      DoEvents
      If Not cl.Comment Is Nothing Then
        With cl.Comment.Shape
          ' サイズ自動設定
          .TextFrame.AutoSize = True
          .TextFrame.Characters.Font.Size = 9

          ' 形状を角丸四角形に変更
          .AutoShapeType = msoShapeRectangle
          ' 塗り色・線色 変更
          .line.ForeColor.RGB = RGB(128, 128, 128)
          .Fill.ForeColor.RGB = RGB(240, 240, 240)
          ' 影 透過率 30%、オフセット量 x:1px,y:1px
          .Shadow.Transparency = 0.3
          .Shadow.OffsetX = 1
          .Shadow.OffsetY = 1
          ' 太字解除、中央揃え
          .TextFrame.Characters.Font.Bold = False
          .TextFrame.HorizontalAlignment = xlLeft
          ' セルに合わせて移動する
          .Placement = xlMove
        End With
      End If
    Next cl
    Application.Goto Reference:=Range("A1"), Scroll:=True
End Function



'**************************************************************************************************
' * クリップボードクリア
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function unsetClipboard()
'  OpenClipboard 0
'  EmptyClipboard
'  CloseClipboard
End Function


'**************************************************************************************************
' * 選択セルの行背景解除
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function unsetLineColor(SetArea As String)
  ActiveSheet.Range(SetArea).Select

  '条件付き書式をクリア
  Selection.FormatConditions.Delete
'  Application.GoTo Reference:=Range("A1"), Scroll:=True
End Function


'**************************************************************************************************
' * リンク解除
' *
' * @Link   https://excel-excellent-technics.com/excel-vba-breaklinks-1019
'**************************************************************************************************
Function unsetLink()
  Dim wb          As Workbook
  Dim vntLink     As Variant
  Dim i           As Integer

  Set wb = ActiveWorkbook
  vntLink = wb.LinkSources(xlLinkTypeExcelLinks) 'ブックの中にあるリンク

  If IsArray(vntLink) Then
    For i = 1 To UBound(vntLink)
      wb.BreakLink vntLink(i), xlLinkTypeExcelLinks 'リンク解除
    Next i
  End If
End Function


'**************************************************************************************************
' * スリープ処理
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function waitTime(timeVal As Long)
  Sleep timeVal
  DoEvents
End Function


'**************************************************************************************************
' * TEXTJOIN関数
' *
' * @Link   https://www.excelspeedup.com/textjoin2/
'**************************************************************************************************
Public Function TEXTJOIN(Delim, Ignore As Boolean, ParamArray par())
  Dim i As Integer
  Dim tR As Range

  TEXTJOIN = ""
  For i = LBound(par) To UBound(par)
    If TypeName(par(i)) = "Range" Then
      For Each tR In par(i)
        If tR.Value <> "" Or Ignore = False Then
          TEXTJOIN = TEXTJOIN & Delim & tR.Value2
        End If
      Next
    Else
      If par(i) <> "" Or Ignore = False Then
        TEXTJOIN = TEXTJOIN & Delim & par(i)
      End If
    End If
  Next

  TEXTJOIN = Mid(TEXTJOIN, Len(Delim) + 1)

End Function


'**************************************************************************************************
' * リボンタブの選択
' *
' * @link https://www.ka-net.org/blog/?p=4624
'**************************************************************************************************
Function selectRibbonTab(ByVal TabName As String)

  Dim uiAuto As UIAutomationClient.CUIAutomation
  Dim elmRibbon As UIAutomationClient.IUIAutomationElement
  Dim elmRibbonTab As UIAutomationClient.IUIAutomationElement
  Dim cndProperty As UIAutomationClient.IUIAutomationCondition
  Dim aryRibbonTab As UIAutomationClient.IUIAutomationElementArray
  Dim ptnAcc As UIAutomationClient.IUIAutomationLegacyIAccessiblePattern
  Dim accRibbon As Office.IAccessible
  Dim i As Long

  Set elmRibbonTab = Nothing '初期化
  Set uiAuto = New UIAutomationClient.CUIAutomation
  Set accRibbon = Application.CommandBars("Ribbon")
  Set elmRibbon = uiAuto.ElementFromIAccessible(accRibbon, 0)
  Set cndProperty = uiAuto.CreatePropertyCondition(UIA_ClassNamePropertyId, "NetUIRibbonTab")
  Set aryRibbonTab = elmRibbon.FindAll(TreeScope_Subtree, cndProperty)
  
  Sleep 500
  
  For i = 0 To aryRibbonTab.length
    Debug.Print aryRibbonTab.GetElement(i).CurrentName
    If aryRibbonTab.GetElement(i).CurrentName = TabName Then
      Set elmRibbonTab = aryRibbonTab.GetElement(i)
      Exit For
    End If
  Next
  If elmRibbonTab Is Nothing Then Exit Function
  Set ptnAcc = elmRibbonTab.GetCurrentPattern(UIA_LegacyIAccessiblePatternId)
  ptnAcc.DoDefaultAction
End Function



