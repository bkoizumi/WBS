Attribute VB_Name = "Library"
'**************************************************************************************************
' * QÆÝèAèé¾
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
' p·éQÆÝèÜÆß
' Microsoft Office 14.0 Object Library
' Microsoft DAO 3.6 Objects Library
' Microsoft Scripting Runtime (WSH, FileSystemObject)
' Microsoft ActiveX Data Objects 2.8 Library
' UIAutomationClient

' Windows APIÌp--------------------------------------------------------------------------------
' fBXvCÌðxæ¾p
' SleepÖÌp
' Nbv{[hÖÌp
#If VBA7 And Win64 Then
  'fBXvCÌðxæ¾p
  Private Declare PtrSafe Function getSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

  'SleepÖÌp
  Private Declare PtrSafe Function Sleep Lib "kernel32" (ByVal ms As LongPtr)

  'Nbv{[hÖA
  Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hWnd As LongPtr) As Long
  Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
  Private Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long

#Else
  'fBXvCÌðxæ¾p
  Private Declare Function getSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

  'SleepÖÌp
  Private Declare Function Sleep Lib "kernel32" (ByVal ms As Long)

  'Nbv{[hÖA
  Declare Function OpenClipboard Lib "user32" (ByVal hWnd As Long) As Long
  Declare Function CloseClipboard Lib "user32" () As Long
  Declare Function EmptyClipboard Lib "user32" () As Long


  'ShellÖÅN®µ½vOÌI¹ðÒÂ
  Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
  Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
  Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
  Private Const PROCESS_QUERY_INFORMATION = &H400&
  Private Const STILL_ACTIVE = &H103&

#End If



'[NubNpÏ------------------------------
'[NV[gpÏ------------------------------
'O[oÏ----------------------------------
Public LibDAO As String
Public LibADOX As String
Public LibADO As String
Public LibScript As String

'ANeBuZÌæ¾
Dim SelectionCell As String
Dim SelectionSheet As String

' PCAOfficeÌîñæ¾pAzzñ
Public MachineInfo As Object

' SeleniumpÝè
Public Const HalfWidthDigit = "1234567890"
Public Const HalfWidthCharacters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
Public Const SymbolCharacters = "!""#$%&'()=~|@[`{;:]+*},./\<>?_-^\"

'Public Const JapaneseCharacters = " ¢¤¦¨©«­¯±³µ·¹»½¿ÂÄÆÈÉÊËÌÍÐÓÖÙÜÝÞßàâäæçèéêëíðñª¬®°²´¶¸º¼¾ÀÃÅÇÎÑÔ×ÚÏÒÕØÛ"
'Public Const JapaneseCharactersCommonUse = "JwóàÂÑæâom·¼Xà¾åéÏçjÝêïKngÀÒÌæóègÂ\¨ú¡½ûapÊè¯G¦aÅüQ¡ü¼²êITÅO{@qáÕ¿ÍøÈ}Éq³µ«§ÅìñgÄá@SÍûqÀóËïFµü¨hÀª©ñïËtÞ|øbÎ¢£"
'Public Const MachineDependentCharacters = "@ABCDEFGHIJKLMNOPQRSTUVWXYZ[\]_ú\ú]ú^ú_ú`úaúbúcúdúeúfúgúhúiújúkúlúmúnúoúpúqúrúsútúuúvúwúxúyúzú{"


Public ThisBook As Workbook


'**************************************************************************************************
' * AhIðÂ¶é
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function addinClose()
  Workbooks(ThisWorkbook.Name).Close
End Function


'**************************************************************************************************
' * G[Ì
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function errorHandle(funcName As String, ByRef objErr As Object)

  Dim Message As String
  Dim runTime As Date
  Dim endLine As Long

  runTime = Format(Now(), "yyyy/mm/dd hh:nn:ss")
  Message = funcName & vbCrLf & objErr.Description

  '¹ºF¯­b
  Application.Speech.Speak Text:="G[ª­¶µÜµ½", SpeakAsync:=True
  Message = Application.WorksheetFunction.VLookup(objErr.Number, noticeCodeSheet.Range("A2:B" & endLine), 2, False)

  Call MsgBox(Message, vbCritical)
  Call endScript
  Call ctl_ProgressBar.showEnd

  Call outputLog(runTime, objErr.Number & vbTab & objErr.Description)
End Function


'**************************************************************************************************
' * æÊ`Ê§äJn
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function startScript()

'  Call Library.showDebugForm("startScript", "")

  'ANeBuZÌæ¾
  If TypeName(Selection) = "Range" Then
    SelectionCell = Selection.Address
    SelectionSheet = ActiveWorkbook.ActiveSheet.Name
  End If

  '}N®ìÅV[gâEBhEªØèÖíéÌð©¹È¢æ¤ÉµÜ·
  Application.ScreenUpdating = False

  '}N®ì©ÌÅÊÌCxgª¶¬³êéÌð}§·é
  Application.EnableEvents = False

  '}N®ìÅZItemNameÈÇªÏíé©®vZªðx­·éÌðð¯é
  Application.Calculation = xlCalculationManual

  '}N®ìÉêØÌL[â}EXìð§À·é
  'Application.Interactive = False

  '}N®ìÍ}EXJ[\ðu»vvÉ·é
'  Application.Cursor = xlWait

  'mFbZ[Wðo³È¢
  Application.DisplayAlerts = False

End Function


'**************************************************************************************************
' * æÊ`Ê§äI¹
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function endScript(Optional flg As Boolean = False)
  On Error Resume Next

  '­§IÉÄvZ³¹é
  Application.CalculateFull

 'ANeBuZÌIð
  If SelectionCell <> "" And flg = True Then
    ActiveWorkbook.Worksheets(SelectionSheet).Select
    ActiveWorkbook.Range(SelectionCell).Select
  End If
  Call unsetClipboard

  '}N®ìÅV[gâEBhEªØèÖíéÌð©¹È¢æ¤ÉµÜ·
  Application.ScreenUpdating = True

  '}N®ì©ÌÅÊÌCxgª¶¬³êéÌð}§·é
  Application.EnableEvents = True

  '}N®ìÅZItemNameÈÇªÏíé©®vZªðx­·éÌðð¯é
  Application.Calculation = xlCalculationAutomatic

  '}N®ìÉêØÌL[â}EXìð§À·é
  'Application.Interactive = True

  '}N®ìI¹ãÍ}EXJ[\ðuftHgvÉàÇ·
  Application.Cursor = xlDefault

  '}N®ìI¹ãÍXe[^Xo[ðuftHgvÉàÇ·
  Application.StatusBar = False

  'mFbZ[Wðo³È¢
  Application.DisplayAlerts = True
End Function


'**************************************************************************************************
' * V[gÌ¶ÝmF
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function chkSheetExists(sheetName) As Boolean

  Dim tempSheet As Object
  Dim result As Boolean

  result = False
  For Each tempSheet In Sheets
    If LCase(sheetName) = LCase(tempSheet.Name) Then
      result = True
      Exit For
    End If
  Next
  chkSheetExists = result
End Function


'**************************************************************************************************
' * ®¹ÜÅÒ@
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function chkShellEnd(ProcessID As Long)
  Dim hProcess As Long
  Dim EndCode As Long
  Dim EndRet   As Long

  hProcess = OpenProcess(PROCESS_QUERY_INFORMATION, 1, ProcessID)
  Do
    EndRet = GetExitCodeProcess(hProcess, EndCode)
    DoEvents
  Loop While (EndCode = STILL_ACTIVE)
  EndRet = CloseHandle(hProcess)
End Function


'**************************************************************************************************
' * I[gVFCvÌ¶ÝmF
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function chkShapeName(ShapeName As String) As Boolean

  Dim objShp As Shape
  Dim result As Boolean

  result = False
  For Each objShp In ActiveSheet.Shapes
    If objShp.Name = ShapeName Then
      result = True
      Exit For
    End If
  Next
  chkShapeName = result
End Function


'**************************************************************************************************
' * OV[g»è
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function chkExcludeSheet(chkSheetName As String) As Boolean

 Dim result As Boolean
  Dim sheetName As Variant

  For Each sheetName In Range("ExcludeSheet")
    If sheetName = chkSheetName Then
      result = True
      Exit For
    Else
      result = False
    End If
  Next
  chkExcludeSheet = result
End Function


'**************************************************************************************************
' * zñªó©Ç¤©
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
 Function chkArrayEmpty(arrayTmp As Variant) As Boolean

On Error GoTo catchError

  If UBound(arrayTmp) >= 0 Then
    chkArrayEmpty = False
  Else
    chkArrayEmpty = True
  End If

  Exit Function

catchError:

  'G[ª­¶µ½ê
  chkArrayEmpty = True

End Function

'**************************************************************************************************
' * ubNªJ©êÄ¢é©`FbN
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
' * wb_[`FbN
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function chkHeader(baseNameArray As Variant, chkNameArray As Variant)
  Dim errMeg As String


On Error GoTo catchError
  errMeg = ""

  If UBound(baseNameArray) <> UBound(chkNameArray) Then
    errMeg = "ÂªÙÈèÜ·B"
    errMeg = errMeg & vbNewLine & UBound(baseNameArray) & "<=>" & UBound(chkNameArray) & vbNewLine
  Else
    For i = LBound(baseNameArray) To UBound(baseNameArray)
      If baseNameArray(i) <> chkNameArray(i) Then
        errMeg = errMeg & vbNewLine & i & ":" & baseNameArray(i) & "<=>" & chkNameArray(i)
      End If
    Next
  End If

  chkHeader = errMeg




  Exit Function

catchError:

  'G[ª­¶µ½ê
  chkHeader = "G[ª­¶µÜµ½"

End Function

'**************************************************************************************************
' * t@CÌ¶ÝmF
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
' * fBNgÌ¶ÝmF
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
' * Åè·¶ñÉÏ·
' *
' * @Link http://bekkou68.hatenablog.com/entry/20090414/1239685179
'**************************************************************************************************
Function convFixedLength(strTarget As String, lengs As Long, addString As String) As String
  Dim strFirst As String
  Dim strExceptFirst As String

  Do While LenB(strTarget) <= lengs
    strTarget = strTarget & addString
  Loop
  convFixedLength = strTarget
End Function


'**************************************************************************************************
' * LP[XðXl[NP[XÉÏ·
' *
' * @Link https://ameblo.jp/i-devdev-beginner/entry-12225328059.html
'**************************************************************************************************
Function covCamelToSnake(ByVal val As String, Optional ByVal isUpper As Boolean = False) As String
  Dim ret As String
  Dim i      As Long, Length As Long

  Length = Len(val)

  For i = 1 To Length
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
' * Xl[NP[XðLP[XÉÏ·
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
' * ¼pÌJ^JiðSpÌJ^JiÉÏ··é(½¾µpÍ¼pÉ·é)
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
    If Mid(rData, i, 1) Like "[`-]" Or Mid(rData, i, 1) Like "[O-X]" Or Mid(rData, i, 1) Like "[|Iij^]" Then
      ansData = ansData & StrConv(Mid(rData, i, 1), vbNarrow)
    Else
      ansData = ansData & Mid(rData, i, 1)
    End If
  Next i
  convHan2Zen = ansData
End Function


'**************************************************************************************************
' * pCvðJ}ÉÏ·
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
' * Base64GR[h(t@C)
' *
' * @link   http://www.ka-net.org/blog/?p=4524
'**************************************************************************************************
Function convBase64EncodeForFile(ByVal filePath As String) As String
  Dim elm As Object
  Dim ret As String
  Const adTypeBinary = 1
  Const adReadAll = -1

  ret = "" 'ú»
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
' * Base64GR[h(¶ñ)
' *
' * @link   http://www.ka-net.org/blog/?p=4524
'**************************************************************************************************
Function convBase64EncodeForString(ByVal str As String) As String

  Dim ret As String
  Dim d() As Byte

  Const adTypeBinary = 1
  Const adTypeText = 2

  ret = "" 'ú»
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
' * URL-safe Base64GR[h
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
' * URLGR[h
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
' * æªP¶Úðå¶»
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
' * ¶ñÌ¶¤©çwè¶í·éÖ
' *
' * @Link   https://vbabeginner.net/vbaÅ¶ñÌE¤â¶¤©çwè¶í·é/
'**************************************************************************************************
Function cutLeft(s, i As Long) As String
  Dim iLen    As Long

  '¶ñÅÍÈ¢ê
  If VarType(s) <> vbString Then
      cutLeft = s & "¶ñÅÍÈ¢"
      Exit Function
  End If

  iLen = Len(s)

  '¶ñ·æèwè¶ªå«¢ê
  If iLen < i Then
      cutLeft = s & "¶ñ·æèwè¶ªå«¢"
      Exit Function
  End If

  cutLeft = Right(s, iLen - i)
End Function


'**************************************************************************************************
' * ¶ñÌE¤©çwè¶í·éÖ
' *
' * @Link   https://vbabeginner.net/vbaÅ¶ñÌE¤â¶¤©çwè¶í·é/
'**************************************************************************************************
Function cutRight(s, i As Long) As String
  Dim iLen    As Long

  If VarType(s) <> vbString Then
    cutRight = s & "¶ñÅÍÈ¢"
    Exit Function
  End If

  iLen = Len(s)

  '¶ñ·æèwè¶ªå«¢ê
  If iLen < i Then
    cutRight = s & "¶ñ·æèwè¶ªå«¢"
    Exit Function
  End If

  cutRight = Left(s, iLen - i)
End Function


'**************************************************************************************************
' * A±üsÌí
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
' *WXg©çîñæ¾
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function delRegistry(registryName As String)
  Dim regVal As String

  On Error Resume Next

  If registryName = "" Then
    DeleteSetting RegistryKey, RegistrySubKey
  Else
    DeleteSetting RegistryKey, RegistrySubKey, registryName
  End If

End Function
'**************************************************************************************************
' * V[gí
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function delSheetData(Optional line As Long)

  If line <> 0 Then
    Rows(line & ":" & Rows.count).Delete Shift:=xlUp
    Rows(line & ":" & Rows.count).Select
    Rows(line & ":" & Rows.count).NumberFormatLocal = "G/W"
    Rows(line & ":" & Rows.count).Style = "Normal"
  Else
    Cells.Delete Shift:=xlUp
    Cells.NumberFormatLocal = "G/W"
    Cells.Style = "Normal"
  End If
  DoEvents

  Application.Goto Reference:=Range("A1"), Scroll:=True
End Function

'**************************************************************************************************
' * ZàÌüsí
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function delCellLinefeed(val As String)
  Dim stringVal As Variant
  Dim retVal As String
  Dim count As Integer

  retVal = ""
  count = 0
  For Each stringVal In Split(val, vbLf)
    If stringVal <> "" And count <= 1 Then
      retVal = retVal & stringVal & vbLf
      count = 0
    Else
      count = count + 1
    End If
  Next
  delCellLinefeed = retVal
End Function

'**************************************************************************************************
' * IðÍÍÌæí
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
' * ZÌ¼ÌÝèí
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
' * e[uf[^í
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
  Selection.NumberFormatLocal = "G/W"

  Application.Goto Reference:=Range("A1"), Scroll:=True
End Function


'**************************************************************************************************
' * PCAOfficeÌîñæ¾
' * AzzñðpµÄ¢éÌÅAMicrosoft Scripting RuntimeªK{
' * MachineInfo.Item ("Excel") ÅÄÑoµ
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function getMachineInfo() As Object
  Dim WshNetworkObject As Object

  Set MachineInfo = CreateObject("Scripting.Dictionary")
  Set WshNetworkObject = CreateObject("WScript.Network")

  ' OSÌo[Wæ¾-----------------------------------------------------------------------------
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

  ' ExcelÌo[Wæ¾--------------------------------------------------------------------------
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

  'PCÌîñ----------------------------------------------------------------------------------------
  MachineInfo.Add "UserName", WshNetworkObject.UserName
  MachineInfo.Add "ComputerName", WshNetworkObject.ComputerName
  MachineInfo.Add "UserDomain", WshNetworkObject.UserDomain

  'æÊÌðxæ¾------------------------------------------------------------------------------
'  MachineInfo.Add "monitors", getSystemMetrics(80)
'  MachineInfo.Add "displayX", getSystemMetrics(0)
'  MachineInfo.Add "displayY", getSystemMetrics(1)
'  MachineInfo.Add "displayVirtualX", getSystemMetrics(78)
'  MachineInfo.Add "displayVirtualY", getSystemMetrics(79)
'  MachineInfo.Add "appTop", ActiveWindow.Top
'  MachineInfo.Add "appLeft", ActiveWindow.Left
'  MachineInfo.Add "appWidth", ActiveWindow.Width
'  MachineInfo.Add "appHeight", ActiveWindow.Height
End Function


'**************************************************************************************************
' * ¶JEg
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
' * ñ¼©çñÔðßé
' *
' * @link   http://www.happy2-island.com/excelsmile/smile03/capter00717.shtml
'**************************************************************************************************
Function getColumnNo(targetCell As String) As Long

  getColumnNo = Range(targetCell & ":" & targetCell).Column
End Function


'**************************************************************************************************
' * ñÔ©çñ¼ðßé
' *
' * @link   http://www.happy2-island.com/excelsmile/smile03/capter00717.shtml
'**************************************************************************************************
Function getColumnName(targetCell As Long) As String

  getColumnName = Split(Cells(, targetCell).Address, "$")(1)
End Function

'**************************************************************************************************
' * J[pbgð\¦µAFR[hðæ¾
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
' * IndentLevellæ¾
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function getIndentLevel(targetRange As Range)
  Dim thisTargetSheet As Worksheet

  Application.Volatile

  If targetRange = "" Then
    getIndentLevel = ""
  Else
    getIndentLevel = targetRange.IndentLevel + 1
  End If
End Function


'**************************************************************************************************
' * RGBlæ¾
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function getRGB(colorValue As Long, Red As Long, Green As Long, Blue As Long)
  Red = colorValue Mod 256
  Green = Int(colorValue / 256) Mod 256
  Blue = Int(colorValue / 256 / 256)
End Function


'**************************************************************************************************
' *WXg©çîñæ¾
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function getRegistry(registryName As String, Optional SubKey As String)
  Dim regVal As String

  If SubKey = "" Then
    SubKey = RegistrySubKey
  End If
  
  If registryName <> "" Then
    regVal = GetSetting(RegistryKey, SubKey, registryName)
  End If
  getRegistry = regVal
End Function


'**************************************************************************************************
' * fBNgIð_CAO\¦
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function getDirPath(CurrentDirectory As String, Optional title As String)

  With Application.FileDialog(msoFileDialogFolderPicker)
    .InitialFileName = CurrentDirectory & "\"
    .AllowMultiSelect = False

    If title <> "" Then
      .title = title & "ÌêðIðµÄ­¾³¢"
    Else
      .title = "tH_[ðIðµÄ­¾³¢"
    End If

    If .Show = True Then
      getDirPath = .SelectedItems(1)
    Else
      getDirPath = ""
    End If
  End With
End Function


'**************************************************************************************************
' * t@CÛ¶_CAO\¦
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function getSaveFilePath(CurrentDirectory As String, saveFileName As String, FileTypeNo As Long)

  Dim filePath As String
  Dim result As Long

  Dim fileName As Variant
  fileName = Application.GetSaveAsFilename( _
      InitialFileName:=CurrentDirectory & "\" & saveFileName, _
      FileFilter:="Excelt@C,*.xlsx,Excel2003ÈO,*.xls,Excel}NubN,*.xlsm", _
      FilterIndex:=FileTypeNo)

  If fileName <> "False" Then
    getSaveFilePath = fileName
  Else
    getSaveFilePath = ""
  End If
End Function

'**************************************************************************************************
' * t@CIð_CAO\¦
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function getFilePath(CurrentDirectory As String, saveFileName As String, title As String, FileTypeNo As Long)

  Dim filePath As String
  Dim result As Long

  With Application.FileDialog(msoFileDialogFilePicker)

    ' t@CÌíÞðÝè
    .Filters.Clear
    .Filters.Add "ExcelubN", "*.xls; *.xlsx; *.xlsm"
    .Filters.Add "CSVt@C", "*.csv"
    .Filters.Add "SQLt@C", "*.sql"
    .Filters.Add "eLXgt@C", "*.txt"
    .Filters.Add "JSONt@C", "*.json"
    .Filters.Add "Accesssf[^x[X", "*.mdb"
    .Filters.Add "·×ÄÌt@C", "*.*"

    .FilterIndex = FileTypeNo

    '\¦·étH_
    .InitialFileName = CurrentDirectory & "\" & saveFileName

    '\¦`®ÌÝè
    .InitialView = msoFileDialogViewWebView

    '_CAO {bNXÌ^CgÝè
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
' * ¡t@CIð_CAO\¦
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function getFilesPath(CurrentDirectory As String, saveFileName As String, title As String, FileTypeNo As Long)

  Dim filePath() As Variant
  Dim result As Long

  With Application.FileDialog(msoFileDialogFilePicker)
    '¡IððÂ
    .AllowMultiSelect = True

    ' t@CÌíÞðÝè
    .Filters.Clear
    .Filters.Add "ExcelubN", "*.xls; *.xlsx; *.xlsm"
    .Filters.Add "CSVt@C", "*.csv"
    .Filters.Add "SQLt@C", "*.sql"
    .Filters.Add "eLXgt@C", "*.txt"
    .Filters.Add "JSONt@C", "*.json"
    .Filters.Add "Accesssf[^x[X", "*.mdb"
    .Filters.Add "·×ÄÌt@C", "*.*"

    .FilterIndex = FileTypeNo

    '\¦·étH_
    .InitialFileName = CurrentDirectory & "\" & saveFileName

    '\¦`®ÌÝè
    .InitialView = msoFileDialogViewWebView

    '_CAO {bNXÌ^CgÝè
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
' * fBNgàÌt@Cêæ¾
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function getFileList(Path As String, fileName As String)
  Dim f As Object, cnt As Long
  Dim list() As String

  cnt = 0
  With CreateObject("Scripting.FileSystemObject")
    For Each f In .GetFolder(Path).Files
      If f.Name Like fileName Then
        ReDim Preserve list(cnt)
        list(cnt) = f.Name
        cnt = cnt + 1
      End If
    Next f
  End With

  getFileList = list
End Function


'**************************************************************************************************
' * wèoCgÌÅè·f[^ì¬(¶ñ)
' *
' * @Link http://www.asahi-net.or.jp/~ef2o-inue/vba_o/function05_110_055.html
'**************************************************************************************************
Function getFixlng(strInText As String, lngFixBytes As Long) As String
    Dim lngKeta As Long
    Dim lngByte As Long, lngByte2 As Long, lngByte3 As Long
    Dim IX As Long
    Dim intCHAR As Long
    Dim strOutText As String

    lngKeta = Len(strInText)
    strOutText = strInText
    ' oCg»è
    For IX = 1 To lngKeta
        ' 1¶¸Â¼p/Spð»f
        intCHAR = Asc(Mid(strInText, IX, 1))
        ' SpÆ»f³êéêÍoCgÉ1ðÁ¦é
        If ((intCHAR < 0) Or (intCHAR > 255)) Then
            lngByte2 = 2        ' Sp
        Else
            lngByte2 = 1        ' ¼p
        End If
        '  Óê»è(EØèÌÄ)
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
    ' s«»è(ó¶ÇÁ)
    If lngByte < lngFixBytes Then
        strOutText = strOutText & Space(lngFixBytes - lngByte)
    End If
    getFixlng = strOutText
End Function


'**************************************************************************************************
' * V[gXgæ¾
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function getSheetList(columnName As String)

  Dim i As Long
  Dim sheetName As Object

  i = 3
  If columnName = "" Then
    columnName = "E"
  End If

  On Error GoTo GetSheetListError:
  Call startScript

  '»ÝèlÌNA
  Worksheets("Ýè").Range(columnName & "3:" & columnName & "100").Select
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

    'V[g¼ÌÝè
    Worksheets("Ýè").Range(columnName & i).Select
    Worksheets("Ýè").Range(columnName & i) = sheetName.Name

    ' ZÌwiFð
    With Worksheets("Ýè").Range(columnName & i).Interior
      .Pattern = xlPatternNone
      .Color = xlNone
    End With

    ' V[gFÆ¯¶FðZÉÝè
    If Worksheets(sheetName.Name).Tab.Color Then
      With Worksheets("Ýè").Range(columnName & i).Interior
        .Pattern = xlPatternNone
        .Color = Worksheets(sheetName.Name).Tab.Color
      End With
    End If

    'rüÌÝè
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

  Worksheets("Ýè").Range(columnName & "3").Select
  Call endScript
  Exit Function
'--------------------------------------------------------------------------------------------------
'G[­¶Ì
'--------------------------------------------------------------------------------------------------
GetSheetListError:

  ' æÊ`Ê§äI¹
  Call endScript
  Call errorHandle("V[gXgæ¾", Err)

End Function


'**************************************************************************************************
' * IðZÌgå\¦Äo
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
' * IðZÌgå\¦I¹
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function showExpansionFormClose(Text As String, SetSelectTargetRows As String)
  Range(SetSelectTargetRows).Value = Text
  Call endScript
End Function


'**************************************************************************************************
' * fobOpæÊ\¦
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function showDebugForm(meg1 As String, Optional meg2 As String)
  Dim runTime As Date
  Dim StartUpPosition As Long

  On Error GoTo catchError

  runTime = Format(Now(), "yyyy/mm/dd hh:nn:ss")

  If setVal("debugMode") = "none" Then
    Exit Function
  End If

'  If StopTime <> 0 Then
'    meg1 = meg1 & vbNewLine & "<ÔF" & StopTime & ">"
'  End If
  meg1 = Replace(meg1, vbNewLine, " ")


  Select Case setVal("debugMode")
    Case "file"
      If meg1 <> "" Then
        Call outputLog(runTime, meg1)
      End If
      GoTo label_end

    Case "form"
      GoTo label_showForm

    Case "all"
      If meg1 <> "" Then
        Call outputLog(runTime, meg1)
      End If
      GoTo label_showForm

    Case "develop"
      If meg1 <> "" Then
        Call outputLog(runTime, meg1)
        Debug.Print runTime & vbTab & meg1
      End If
      'GoTo label_showForm
      GoTo label_end

    Case Else
      Exit Function
  End Select

label_showForm:
  If meg1 Like "JnF*" Then

    With Frm_debug
      .Caption = "îñ"
      .ListBox1.Clear
      .ListBox1.AddItem runTime & vbTab & meg1
    End With
  Else
    With Frm_debug
      .Caption = "îñ"
      .ListBox1.AddItem runTime & vbTab & meg1
      .ListBox1.ListIndex = .ListBox1.ListCount - 1
    End With
  End If

  If (Frm_debug.Visible = True) Then
    Frm_debug.StartUpPosition = 0
  Else
    Frm_debug.StartUpPosition = 1
  End If
  Frm_debug.Show vbModeless


label_end:

  DoEvents
  Exit Function

'G[­¶=====================================================================================
catchError:
  Exit Function
End Function

'**************************************************************************************************
' * îñÊm
' *
' * Worksheets("Notice").Visible = True
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function showNotice(Code As Long, Optional process As String, Optional runEndflg As Boolean)


  Dim Message As String
  Dim runTime As Date
  Dim endLine As Long

  runTime = Format(Now(), "yyyy/mm/dd hh:nn:ss")

  endLine = sheetNotice.Cells(Rows.count, 1).End(xlUp).row
  Message = Application.WorksheetFunction.VLookup(Code, sheetNotice.Range("A2:B" & endLine), 2, False)

  If process <> "" Then
    Message = Replace(Message, "%%", process)
  End If
  If runEndflg = True Then
    Message = Message & vbNewLine & "ð~µÜ·"
  End If

  If StopTime <> 0 Then
    Message = Message & vbNewLine & "<ÔF" & StopTime & ">"
  End If

  If setVal("debugMode") = "speak" Or setVal("debugMode") = "develop" Or setVal("debugMode") = "all" Then
    Application.Speech.Speak Text:=Message, SpeakAsync:=True, SpeakXML:=True
  Else
'    Call outputLog(runTime, Message)
  End If



  Select Case Code
    Case 0 To 399
      Call MsgBox(Message, vbInformation, thisAppName)

    Case 400 To 499
      Call MsgBox(Message, vbCritical, thisAppName)

    Case 500 To 599
      Call MsgBox(Message, vbExclamation, thisAppName)

    Case Else
      Call MsgBox(Message, vbCritical, thisAppName)
  End Select
'  Stop

  'æÊ`Ê§äI¹
  If runEndflg = True Then
    Call endScript
    Call ctl_ProgressBar.showEnd
    End
  Else
    Call Library.showDebugForm(Message)
  End If
End Function


'**************************************************************************************************
' * _
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function makeRandomString(ByVal setString As String, ByVal setStringCnt As Integer) As String

 For i = 1 To setStringCnt
    'WFl[^ðú»
    Randomize
    n = Int((Len(setString) - 1 + 1) * Rnd + 1)
    str1 = str1 + Mid(setString, n, 1)
  Next i

  makeRandomString = str1

End Function

Function makeRandomNo(minNo As Long, maxNo As Long) As String

  'WFl[^ðú»
  Randomize
  makeRandomNo = Int((maxNo - minNo + 1) * Rnd + minNo)

End Function

'**************************************************************************************************
' * OoÍ
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function outputLog(runTime As Date, Message As String)
  Dim fileTimestamp As Date

  If chkFileExists(logFile) Then
    fileTimestamp = FileDateTime(logFile)
  Else
      fileTimestamp = DateAdd("d", -1, Date)
  End If

  With CreateObject("ADODB.Stream")
    .Charset = "UTF-8"
    .Open
    If Format(Date, "yyyymmdd") = Format(fileTimestamp, "yyyymmdd") Then
      .LoadFromFile logFile
      .Position = .Size
    End If
    .WriteText "<p><span class='time'>" & runTime & "</span>@<span class='message'>" & Message & "</span></p>", 1
    .SaveToFile logFile, 2
    .Close
  End With

End Function

'==================================================================================================
Function outputText(Message As String, outputFilePath)

  Open outputFilePath For Output As #1
  Print #1, Message
  Close #1

End Function

'**************************************************************************************************
' * CSVC|[g
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
    .TextFilePlatform = 932          ' Shift-JIS ðJ­
    .TextFileParseType = xlDelimited ' ¶ÅæØÁ½`®
    .TextFileCommaDelimiter = True   ' æØè¶ÍJ}
    .TextFileStartRow = readLine     ' 1 sÚ©çÇÝÝ
    .AdjustColumnWidth = False       ' ñð©®²®µÈ¢
    .RefreshStyle = xlOverwriteCells 'ã«ðwè
    .TextFileTextQualifier = xlTextQualifierDoubleQuote ' øpÌwè

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
' * Excelt@CÌC|[g
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function importXlsx(filePath As String, targeSheet As String, targeArea As String, dictSheet As Worksheet, Optional passWord As String)

  If passWord <> "" Then
    Workbooks.Open fileName:=filePath, ReadOnly:=True, passWord:=passWord
  Else
    Workbooks.Open fileName:=filePath, ReadOnly:=True
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
' * MkDirÅKwÌ[¢tH_[ðìé
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
  tmpPath = arr(0)  ' hCu¼Ìãü

  For i = 1 To UBound(arr)
    tmpPath = tmpPath & "\" & arr(i)
    If Dir(tmpPath, vbDirectory) = "" Then
      MkDir tmpPath
    End If
  Next i

End Function


'**************************************************************************************************
' * pX[h¶¬
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function makePasswd() As String
  Dim halfChar As String, str1 As String

  halfChar = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz!#$%&"

  For i = 1 To 12
    'WFl[^ðú»
    Randomize
    n = Int((Len(halfChar) - 1 + 1) * Rnd + 1)
    str1 = str1 + Mid(halfChar, n, 1)
  Next i
  makePasswd = str1
End Function


'**************************************************************************************************
' * ¶ñª
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
' * zñÌÅãÉÇÁ·é
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
' * tHgJ[Ýè
' *
' * @Link https://vbabeginner.net/vbaÅZÌwè¶ñÌFâ¾³ðÏX·é/
'**************************************************************************************************
Function setFontClor(a_sSearch, a_lColor, a_bBold)
  Dim f   As Font     'FontIuWFNg
  Dim i               'ø¶ñÌZÌÊu
  Dim iLen            'ø¶ñÌ¶
  Dim R   As Range    'ZÍÍÌPZ

  iLen = Len(a_sSearch)
  i = 1

  For Each R In Selection
    Do
      i = InStr(i, R.Value, a_sSearch)
      If (i = 0) Then
        i = 1
        Exit Do
      End If
      Set f = R.Characters(i, iLen).Font
      f.Color = a_lColor
      f.Bold = a_bBold
      i = i + 1
    Loop
  Next
End Function


'**************************************************************************************************
' * WXgÉîño^
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function setRegistry(registryName As String, setVal As Variant, Optional SubKey As String)

  If SubKey = "" Then
    SubKey = RegistrySubKey
  End If

  If getRegistry(registryName, SubKey) <> setVal And registryName <> "" Then
    Call SaveSetting(RegistryKey, SubKey, registryName, setVal)
  End If
End Function


'**************************************************************************************************
' * QÆÝèð©®Ås¤
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
      MsgBox ("Microsoft Scripting RuntimeðpÅ«Ü¹ñB" & vbLf & "pÅ«È¢@\ª èÜ·")
    End If

  'Microsoft ActiveX Data Objects Library 6.1 (ADO)------------------------------------------------
  If BookType = "DataBase" Then
    LibADO = "C:\Program Files\Common Files\System\Ado\msado15.dll"
    If Dir(LibADO) <> "" Then
      ActiveWorkbook.VBProject.References.AddFromFile (LibADO)
    Else
      MsgBox ("Microsoft ActiveX Data ObjectsðpÅ«Ü¹ñ" & vbLf & "pÅ«È¢@\ª èÜ·")
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
        MsgBox ("Microsoft DAO 3.6 Objects LibraryðpÅ«Ü¹ñ" & vbLf & "DBÖÌÚ±@\ªpÅ«Ü¹ñ")
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
        MsgBox ("Microsoft DAO 3.6 Objects LibraryðpÅ«Ü¹ñ" & vbLf & "DBÖÌÚ±@\ªpÅ«Ü¹ñ")
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
    MsgBox ("uVBA vWFNg IuWFNg fÖÌANZXðM·évÉÏXðI")
  Else
    MsgBox "Error Number : " & Err.Number & vbCrLf & Err.Description
    GoTo Func_Exit:
  End If
End Function


'**************************************************************************************************
' * IðZÌswiÝè
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function setLineColor(setArea As String, DisType As Boolean, SetColor As String)

  Range(setArea).Select

  'ðt«®ðNA
  Selection.FormatConditions.Delete

  If DisType = False Then
    's¾¯Ýè
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:="=CELL(""row"")=ROW()"
  Else
    'sÆñÉÝè
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:="=OR(CELL(""row"")=ROW(), CELL(""col"")=COLUMN())"
  End If

  Selection.FormatConditions(Selection.FormatConditions.count).SetFirstPriority
  With Selection.FormatConditions(1)
    .Interior.PatternColorIndex = xlAutomatic
    .Interior.Color = SetColor
'    .Interior.TintAndShade = 0
'    .Font.ColorIndex = 1
  End With
  Selection.FormatConditions(1).StopIfTrue = False
End Function


'**************************************************************************************************
' * ÅÌV[gðIð
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
' * t@CSÌÌ¶ñu·
' *
' * @Link   https://www.moug.net/tech/acvba/0090005.html
'**************************************************************************************************
Function replaceFromFile(fileName As String, TargetText As String, Optional NewText As String = "")

 Dim FSO         As FileSystemObject 't@CVXeIuWFNg
 Dim Txt         As TextStream       'eLXgXg[IuWFNg
 Dim buf_strTxt  As String           'ÇÝÝobt@

 On Error GoTo Func_Err:

 'IuWFNgì¬
 Set FSO = CreateObject("Scripting.FileSystemObject")
 Set Txt = FSO.OpenTextFile(fileName, ForReading)

 'S¶ÇÝÝ
  buf_strTxt = Txt.ReadAll
  Txt.Close

  '³t@Cðl[µÄAe|t@Cì¬
  Name fileName As fileName & "_"

  'u·
   buf_strTxt = Replace(buf_strTxt, TargetText, NewText, , , vbBinaryCompare)

  'ÝpeLXgt@Cì¬
   Set Txt = FSO.CreateTextFile(fileName, True)
  'Ý
  Txt.Write buf_strTxt
  Txt.Close

  'e|t@Cðí
  FSO.DeleteFile fileName & "_"

'I¹
Func_Exit:
    Set Txt = Nothing
    Set FSO = Nothing
    Exit Function

Func_Err:
    MsgBox "Error Number : " & Err.Number & vbCrLf & Err.Description
    GoTo Func_Exit:
End Function


'**************************************************************************************************
' * VBAÅExcelÌRgðêÅ©®TCYÉµÄJbRæ­·é
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
          ' TCY©®Ýè
          .TextFrame.AutoSize = True
          .TextFrame.Characters.Font.Size = 9

          ' `óðpÛlp`ÉÏX
          .AutoShapeType = msoShapeRectangle
          ' hèFEüF ÏX
          .line.ForeColor.RGB = RGB(128, 128, 128)
          .Fill.ForeColor.RGB = RGB(240, 240, 240)
          ' e §ß¦ 30%AItZbgÊ x:1px,y:1px
          .Shadow.Transparency = 0.3
          .Shadow.OffsetX = 1
          .Shadow.OffsetY = 1
          ' ¾ðAµ¦
          .TextFrame.Characters.Font.Bold = False
          .TextFrame.HorizontalAlignment = xlLeft
          ' ZÉí¹ÄÚ®·é
          .Placement = xlMove
        End With
      End If
    Next cl
    Application.Goto Reference:=Range("A1"), Scroll:=True
End Function



'**************************************************************************************************
' * Nbv{[hNA
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function unsetClipboard()
'  OpenClipboard 0
'  EmptyClipboard
'  CloseClipboard
End Function


'**************************************************************************************************
' * IðZÌswið
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function unsetLineColor(setArea As String)
  ActiveWorkbook.ActiveSheet.Range(setArea).Select

  'ðt«®ðNA
  Selection.FormatConditions.Delete
'  Application.GoTo Reference:=Range("A1"), Scroll:=True
End Function


'**************************************************************************************************
' * Nð
' *
' * @Link   https://excel-excellent-technics.com/excel-vba-breaklinks-1019
'**************************************************************************************************
Function unsetLink()
  Dim wb          As Workbook
  Dim vntLink     As Variant
  Dim i           As Integer

  Set wb = ActiveWorkbook
  vntLink = wb.LinkSources(xlLinkTypeExcelLinks) 'ubNÌÉ éN

  If IsArray(vntLink) Then
    For i = 1 To UBound(vntLink)
      wb.BreakLink vntLink(i), xlLinkTypeExcelLinks 'Nð
    Next i
  End If
End Function


'**************************************************************************************************
' * X[v
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function waitTime(timeVal As Long)
  DoEvents
  'Sleep timeVal
  
  Application.Wait [Now()] + timeVal / 86400000
  DoEvents
End Function


'**************************************************************************************************
' * TEXTJOINÖ
' *
' * @Link   https://www.excelspeedup.com/textjoin2/
'**************************************************************************************************
Function TEXTJOIN(Delim, Ignore As Boolean, ParamArray par())
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
      If (par(i) <> "" And par(i) <> "<>") Or Ignore = False Then
        TEXTJOIN = TEXTJOIN & Delim & par(i)
      End If
    End If
  Next

  TEXTJOIN = Mid(TEXTJOIN, Len(Delim) + 1)

End Function


'**************************************************************************************************
' * IðZÌgå\¦Äo
' *
' * @author Bunpei.Koizumi<koizumi.bunpei@trans-cosmos.co.jp>
'**************************************************************************************************
Function KOETOL_ExpansionFormStart(Text As String, SetSelectTargetRows As String)
  Dim colLineName As Variant
  Dim count As Integer

  With KOETOL_ExpansionForm
    .StartUpPosition = 2
'    .Top = Application.Top + (ActiveWindow.Width / 4)
'    .Left = Application.Left + (ActiveWindow.Height / 4)
    .TextBox = Text
'    .TextBox.ForeColor = Color
    .TextBox.MultiLine = True
    .TextBox.MultiLine = True
    .TextBox.EnterKeyBehavior = True
    .Caption = SetSelectTargetRows

    'j[YÌ`FbN{bNXÌÝè
    .needs01.Caption = Range("J3").Value
    .needs02.Caption = Range("K3").Value
    .needs03.Caption = Range("L3").Value
    .needs04.Caption = Range("M3").Value
    .needs05.Caption = Range("N3").Value
    .needs06.Caption = Range("O3").Value
    .needs07.Caption = Range("P3").Value
    .needs08.Caption = Range("Q3").Value
    .needs09.Caption = Range("R3").Value
    .needs10.Caption = Range("S3").Value

    '|WeBuÌ`FbN{bNXÌÝè
    .positive01.Caption = Range("T3").Value
    .positive02.Caption = Range("U3").Value
    .positive03.Caption = Range("V3").Value
    .positive04.Caption = Range("W3").Value
    .positive05.Caption = Range("X3").Value
    .positive06.Caption = Range("Y3").Value
    .positive07.Caption = Range("Z3").Value
    .positive08.Caption = Range("AA3").Value
    .positive09.Caption = Range("AB3").Value
    .positive10.Caption = Range("AC3").Value
    .positive11.Caption = Range("AD3").Value
    .positive12.Caption = Range("AE3").Value
    .positive13.Caption = Range("AF3").Value
    .positive14.Caption = Range("AG3").Value
    .positive15.Caption = Range("AH3").Value

    'lKeBuÌ`FbN{bNXÌÝè
    .negative01.Caption = Range("AI3").Value
    .negative02.Caption = Range("AJ3").Value
    .negative03.Caption = Range("AK3").Value
    .negative04.Caption = Range("AL3").Value
    .negative05.Caption = Range("AM3").Value
    .negative06.Caption = Range("AN3").Value
    .negative07.Caption = Range("AO3").Value
    .negative08.Caption = Range("AP3").Value
    .negative09.Caption = Range("AQ3").Value
    .negative10.Caption = Range("AR3").Value
    .negative11.Caption = Range("AS3").Value
    .negative12.Caption = Range("AT3").Value
    .negative13.Caption = Range("AU3").Value
    .negative14.Caption = Range("AV3").Value
    .negative15.Caption = Range("AW3").Value

    'j[YÌ`FbN{bNXÌl
    .needs01.Value = IIf(Range("J" & SetSelectTargetRows).Value, True, False)
    .needs02.Value = IIf(Range("K" & SetSelectTargetRows).Value, True, False)
    .needs03.Value = IIf(Range("L" & SetSelectTargetRows).Value, True, False)
    .needs04.Value = IIf(Range("M" & SetSelectTargetRows).Value, True, False)
    .needs05.Value = IIf(Range("N" & SetSelectTargetRows).Value, True, False)
    .needs06.Value = IIf(Range("O" & SetSelectTargetRows).Value, True, False)
    .needs07.Value = IIf(Range("P" & SetSelectTargetRows).Value, True, False)
    .needs08.Value = IIf(Range("Q" & SetSelectTargetRows).Value, True, False)
    .needs09.Value = IIf(Range("R" & SetSelectTargetRows).Value, True, False)
    .needs10.Value = IIf(Range("S" & SetSelectTargetRows).Value, True, False)

    '|WeBuÌ`FbN{bNXÌl
    .positive01.Value = IIf(Range("T" & SetSelectTargetRows).Value, True, False)
    .positive02.Value = IIf(Range("U" & SetSelectTargetRows).Value, True, False)
    .positive03.Value = IIf(Range("V" & SetSelectTargetRows).Value, True, False)
    .positive04.Value = IIf(Range("W" & SetSelectTargetRows).Value, True, False)
    .positive05.Value = IIf(Range("X" & SetSelectTargetRows).Value, True, False)
    .positive06.Value = IIf(Range("Y" & SetSelectTargetRows).Value, True, False)
    .positive07.Value = IIf(Range("Z" & SetSelectTargetRows).Value, True, False)
    .positive08.Value = IIf(Range("AA" & SetSelectTargetRows).Value, True, False)
    .positive09.Value = IIf(Range("AB" & SetSelectTargetRows).Value, True, False)
    .positive10.Value = IIf(Range("AC" & SetSelectTargetRows).Value, True, False)
    .positive11.Value = IIf(Range("AD" & SetSelectTargetRows).Value, True, False)
    .positive12.Value = IIf(Range("AE" & SetSelectTargetRows).Value, True, False)
    .positive13.Value = IIf(Range("AF" & SetSelectTargetRows).Value, True, False)
    .positive14.Value = IIf(Range("AG" & SetSelectTargetRows).Value, True, False)
    .positive15.Value = IIf(Range("AH" & SetSelectTargetRows).Value, True, False)

    'lKeBuÌ`FbN{bNXÌl
    .negative01.Value = IIf(Range("AI" & SetSelectTargetRows).Value, True, False)
    .negative02.Value = IIf(Range("AJ" & SetSelectTargetRows).Value, True, False)
    .negative03.Value = IIf(Range("AK" & SetSelectTargetRows).Value, True, False)
    .negative04.Value = IIf(Range("AL" & SetSelectTargetRows).Value, True, False)
    .negative05.Value = IIf(Range("AM" & SetSelectTargetRows).Value, True, False)
    .negative06.Value = IIf(Range("AN" & SetSelectTargetRows).Value, True, False)
    .negative07.Value = IIf(Range("AO" & SetSelectTargetRows).Value, True, False)
    .negative08.Value = IIf(Range("AP" & SetSelectTargetRows).Value, True, False)
    .negative09.Value = IIf(Range("AQ" & SetSelectTargetRows).Value, True, False)
    .negative10.Value = IIf(Range("AR" & SetSelectTargetRows).Value, True, False)
    .negative11.Value = IIf(Range("AS" & SetSelectTargetRows).Value, True, False)
    .negative12.Value = IIf(Range("AT" & SetSelectTargetRows).Value, True, False)
    .negative13.Value = IIf(Range("AU" & SetSelectTargetRows).Value, True, False)
    .negative14.Value = IIf(Range("AV" & SetSelectTargetRows).Value, True, False)
    .negative15.Value = IIf(Range("AW" & SetSelectTargetRows).Value, True, False)

  End With

  If (KOETOL_ExpansionForm.Visible = True) Then
    KOETOL_ExpansionForm.StartUpPosition = 0
  Else
    KOETOL_ExpansionForm.StartUpPosition = 2
  End If

  KOETOL_ExpansionForm.Show vbModeless

End Function


'**************************************************************************************************
' * IðZÌgå\¦I¹
' *
' * @author Bunpei.Koizumi<koizumi.bunpei@trans-cosmos.co.jp>
'**************************************************************************************************
Function KOETOL_ExpansionFormEnd()
  Call init.setting


  If setVal("HighLightFlg") = False Then
    SetActiveCell = Selection.Address
    endRowLine = sheetKoetol.Cells(Rows.count, 3).End(xlUp).row

    Call Library.startScript
    Call Library.unsetLineColor("C5:AZ" & endRowLine)
    Call Library.endScript

    Range(SetActiveCell).Select
  End If

End Function




'**************************************************************************************************
' * rü
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
'--------------------------------------------------------------------------------------------------
Function rü_NA(Optional setArea As Range)
  If TypeName(setArea) = "Range" Then
    With setArea
      .Borders(xlInsideVertical).LineStyle = xlNone
      .Borders(xlInsideHorizontal).LineStyle = xlNone
      .Borders(xlEdgeLeft).LineStyle = xlNone
      .Borders(xlEdgeRight).LineStyle = xlNone
      .Borders(xlEdgeTop).LineStyle = xlNone
      .Borders(xlEdgeBottom).LineStyle = xlNone
      .Borders(xlInsideVertical).LineStyle = xlNone
      .Borders(xlInsideHorizontal).LineStyle = xlNone
    End With
  Else
    With Selection
      .Borders(xlInsideVertical).LineStyle = xlNone
      .Borders(xlInsideHorizontal).LineStyle = xlNone
      .Borders(xlEdgeLeft).LineStyle = xlNone
      .Borders(xlEdgeRight).LineStyle = xlNone
      .Borders(xlEdgeTop).LineStyle = xlNone
      .Borders(xlEdgeBottom).LineStyle = xlNone
      .Borders(xlInsideVertical).LineStyle = xlNone
      .Borders(xlInsideHorizontal).LineStyle = xlNone
    End With
  End If
End Function

'--------------------------------------------------------------------------------------------------
Function rü_jü_ÍÝ(Optional setArea As Range, Optional lineColor As Long, Optional WeightVal = xlHairline)
  Dim Red As Long, Green As Long, Blue As Long

  Call Library.getRGB(lineColor, Red, Green, Blue)

  If TypeName(setArea) = "Range" Then
    With setArea
      .Borders(xlEdgeLeft).LineStyle = xlDash
      .Borders(xlEdgeRight).LineStyle = xlDash
      .Borders(xlEdgeTop).LineStyle = xlDash
      .Borders(xlEdgeBottom).LineStyle = xlDash

      .Borders(xlEdgeLeft).Weight = WeightVal
      .Borders(xlEdgeRight).Weight = WeightVal
      .Borders(xlEdgeTop).Weight = WeightVal
      .Borders(xlEdgeBottom).Weight = WeightVal
      If Not (IsMissing(Red)) Then
        .Borders(xlEdgeLeft).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeRight).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeTop).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeBottom).Color = RGB(Red, Green, Blue)
      End If
    End With
  Else
    With Selection
      .Borders(xlEdgeLeft).LineStyle = xlDash
      .Borders(xlEdgeRight).LineStyle = xlDash
      .Borders(xlEdgeTop).LineStyle = xlDash
      .Borders(xlEdgeBottom).LineStyle = xlDash

      .Borders(xlEdgeLeft).Weight = WeightVal
      .Borders(xlEdgeRight).Weight = WeightVal
      .Borders(xlEdgeTop).Weight = WeightVal
      .Borders(xlEdgeBottom).Weight = WeightVal
      If Not (IsMissing(Red)) Then
        .Borders(xlEdgeLeft).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeRight).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeTop).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeBottom).Color = RGB(Red, Green, Blue)
      End If
    End With
  End If
End Function


'--------------------------------------------------------------------------------------------------
Function rü_jü_iq(Optional setArea As Range, Optional lineColor As Long, Optional WeightVal = xlHairline)
  Dim Red As Long, Green As Long, Blue As Long

  Call Library.getRGB(lineColor, Red, Green, Blue)

  If TypeName(setArea) = "Range" Then
    With setArea
      .Borders(xlEdgeLeft).LineStyle = xlDash
      .Borders(xlEdgeRight).LineStyle = xlDash
      .Borders(xlEdgeTop).LineStyle = xlDash
      .Borders(xlEdgeBottom).LineStyle = xlDash
      .Borders(xlInsideVertical).LineStyle = xlDash
      .Borders(xlInsideHorizontal).LineStyle = xlDash

      .Borders(xlEdgeLeft).Weight = WeightVal
      .Borders(xlEdgeRight).Weight = WeightVal
      .Borders(xlEdgeTop).Weight = WeightVal
      .Borders(xlEdgeBottom).Weight = WeightVal
      .Borders(xlInsideVertical).Weight = WeightVal
      .Borders(xlInsideHorizontal).Weight = WeightVal

      If Not (IsMissing(Red)) Then
        .Borders(xlEdgeLeft).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeRight).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeTop).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeBottom).Color = RGB(Red, Green, Blue)
      End If
    End With
  Else
    With Selection
      .Borders(xlEdgeLeft).LineStyle = xlDash
      .Borders(xlEdgeRight).LineStyle = xlDash
      .Borders(xlEdgeTop).LineStyle = xlDash
      .Borders(xlEdgeBottom).LineStyle = xlDash
      .Borders(xlInsideVertical).LineStyle = xlDash
      .Borders(xlInsideHorizontal).LineStyle = xlDash

      .Borders(xlEdgeLeft).Weight = WeightVal
      .Borders(xlEdgeRight).Weight = WeightVal
      .Borders(xlEdgeTop).Weight = WeightVal
      .Borders(xlEdgeBottom).Weight = WeightVal
      .Borders(xlInsideVertical).Weight = WeightVal
      .Borders(xlInsideHorizontal).Weight = WeightVal

      If Not (IsMissing(Red)) Then
        .Borders(xlEdgeLeft).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeRight).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeTop).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeBottom).Color = RGB(Red, Green, Blue)
      End If
    End With
  End If
End Function


'--------------------------------------------------------------------------------------------------
Function rü_jü_¶E(Optional setArea As Range, Optional lineColor As Long, Optional WeightVal = xlHairline)
  Dim Red As Long, Green As Long, Blue As Long

  Call Library.getRGB(lineColor, Red, Green, Blue)

  If TypeName(setArea) = "Range" Then
    With setArea
      .Borders(xlEdgeLeft).LineStyle = xlDash
      .Borders(xlEdgeRight).LineStyle = xlDash
      
      .Borders(xlEdgeLeft).Weight = WeightVal
      .Borders(xlEdgeRight).Weight = WeightVal

      If Not (IsMissing(Red)) Then
        .Borders(xlEdgeLeft).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeRight).Color = RGB(Red, Green, Blue)
      End If
     End With
  Else

    With Selection
      .Borders(xlEdgeLeft).LineStyle = xlDash
      .Borders(xlEdgeRight).LineStyle = xlDash
      
      .Borders(xlEdgeLeft).Weight = WeightVal
      .Borders(xlEdgeRight).Weight = WeightVal

      If Not (IsMissing(Red)) Then
        .Borders(xlEdgeLeft).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeRight).Color = RGB(Red, Green, Blue)
      End If
     End With
  End If
End Function

'--------------------------------------------------------------------------------------------------
Function rü_jü_ãº(Optional setArea As Range, Optional lineColor As Long, Optional WeightVal = xlHairline)
  Dim Red As Long, Green As Long, Blue As Long

  Call Library.getRGB(lineColor, Red, Green, Blue)

  If TypeName(setArea) = "Range" Then
    With setArea
      .Borders(xlEdgeTop).LineStyle = xlDash
      .Borders(xlEdgeBottom).LineStyle = xlDash

      .Borders(xlEdgeTop).Weight = WeightVal
      .Borders(xlEdgeBottom).Weight = WeightVal

      If Not (IsMissing(Red)) Then
        .Borders(xlEdgeTop).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeBottom).Color = RGB(Red, Green, Blue)
      End If
    End With
  Else
    With Selection
      .Borders(xlEdgeTop).LineStyle = xlDash
      .Borders(xlEdgeBottom).LineStyle = xlDash

      .Borders(xlEdgeTop).Weight = WeightVal
      .Borders(xlEdgeBottom).Weight = WeightVal

      If Not (IsMissing(Red)) Then
        .Borders(xlEdgeTop).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeBottom).Color = RGB(Red, Green, Blue)
      End If
    End With
  End If
End Function


'--------------------------------------------------------------------------------------------------
Function rü_jü_¼(Optional setArea As Range, Optional lineColor As Long, Optional WeightVal = xlHairline)
  Dim Red As Long, Green As Long, Blue As Long

  Call Library.getRGB(lineColor, Red, Green, Blue)

  If TypeName(setArea) = "Range" Then
    With setArea
      .Borders(xlInsideVertical).LineStyle = xlDash
      .Borders(xlInsideVertical).Weight = WeightVal
      If Not (IsMissing(Red)) Then
        .Borders(xlInsideVertical).Color = RGB(Red, Green, Blue)
      End If
    End With
  Else
    With Selection
      .Borders(xlInsideVertical).LineStyle = xlDash
      .Borders(xlInsideVertical).Weight = WeightVal
      If Not (IsMissing(Red)) Then
        .Borders(xlInsideVertical).Color = RGB(Red, Green, Blue)
      End If
    End With
  End If
End Function


'--------------------------------------------------------------------------------------------------
Function rü_jü_½(Optional setArea As Range, Optional lineColor As Long, Optional WeightVal = xlHairline)
  Dim Red As Long, Green As Long, Blue As Long

  Call Library.getRGB(lineColor, Red, Green, Blue)

  If TypeName(setArea) = "Range" Then
    With setArea
      .Borders(xlInsideHorizontal).LineStyle = xlDash
      .Borders(xlInsideHorizontal).Weight = WeightVal

      If Not (IsMissing(Red)) Then
        .Borders(xlInsideHorizontal).Color = RGB(Red, Green, Blue)
      End If
     End With
  Else

    With Selection
      .Borders(xlInsideHorizontal).LineStyle = xlDash
      .Borders(xlInsideHorizontal).Weight = WeightVal

      If Not (IsMissing(Red)) Then
        .Borders(xlInsideHorizontal).Color = RGB(Red, Green, Blue)
      End If
    End With

  End If
End Function





'--------------------------------------------------------------------------------------------------
Function rü_Àü_ÍÝ(Optional setArea As Range, Optional lineColor As Long, Optional WeightVal = xlThin)
  Dim Red As Long, Green As Long, Blue As Long

  Call Library.getRGB(lineColor, Red, Green, Blue)

  If TypeName(setArea) = "Range" Then
    With setArea
      .Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Borders(xlEdgeRight).LineStyle = xlContinuous
      .Borders(xlEdgeTop).LineStyle = xlContinuous
      .Borders(xlEdgeBottom).LineStyle = xlContinuous

      .Borders(xlEdgeLeft).Weight = WeightVal
      .Borders(xlEdgeRight).Weight = WeightVal
      .Borders(xlEdgeTop).Weight = WeightVal
      .Borders(xlEdgeBottom).Weight = WeightVal

      If Not (IsMissing(Red)) Then
        .Borders(xlEdgeLeft).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeRight).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeTop).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeBottom).Color = RGB(Red, Green, Blue)
      End If
    End With
  Else
    With Selection
      .Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Borders(xlEdgeRight).LineStyle = xlContinuous
      .Borders(xlEdgeTop).LineStyle = xlContinuous
      .Borders(xlEdgeBottom).LineStyle = xlContinuous

      .Borders(xlEdgeLeft).Weight = WeightVal
      .Borders(xlEdgeRight).Weight = WeightVal
      .Borders(xlEdgeTop).Weight = WeightVal
      .Borders(xlEdgeBottom).Weight = WeightVal

      If Not (IsMissing(Red)) Then
        .Borders(xlEdgeLeft).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeRight).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeTop).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeBottom).Color = RGB(Red, Green, Blue)
      End If
    End With
  End If
End Function

'--------------------------------------------------------------------------------------------------
Function rü_Àü_iq(Optional setArea As Range, Optional lineColor As Long, Optional WeightVal = xlThin)
  Dim Red As Long, Green As Long, Blue As Long

  Call Library.getRGB(lineColor, Red, Green, Blue)

  If TypeName(setArea) = "Range" Then
    With setArea
      .Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Borders(xlEdgeRight).LineStyle = xlContinuous
      .Borders(xlEdgeTop).LineStyle = xlContinuous
      .Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Borders(xlInsideVertical).LineStyle = xlContinuous
      .Borders(xlInsideHorizontal).LineStyle = xlContinuous

      .Borders(xlEdgeLeft).Weight = WeightVal
      .Borders(xlEdgeRight).Weight = WeightVal
      .Borders(xlEdgeTop).Weight = WeightVal
      .Borders(xlEdgeBottom).Weight = WeightVal
      .Borders(xlInsideVertical).Weight = WeightVal
      .Borders(xlInsideHorizontal).Weight = WeightVal

      If Not (IsMissing(Red)) Then
        .Borders(xlEdgeLeft).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeRight).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeTop).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeBottom).Color = RGB(Red, Green, Blue)
        .Borders(xlInsideVertical).Color = RGB(Red, Green, Blue)
        .Borders(xlInsideHorizontal).Color = RGB(Red, Green, Blue)
      End If
    End With
  Else
    With Selection
      .Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Borders(xlEdgeRight).LineStyle = xlContinuous
      .Borders(xlEdgeTop).LineStyle = xlContinuous
      .Borders(xlEdgeBottom).LineStyle = xlContinuous
      .Borders(xlInsideVertical).LineStyle = xlContinuous
      .Borders(xlInsideHorizontal).LineStyle = xlContinuous

      .Borders(xlEdgeLeft).Weight = WeightVal
      .Borders(xlEdgeRight).Weight = WeightVal
      .Borders(xlEdgeTop).Weight = WeightVal
      .Borders(xlEdgeBottom).Weight = WeightVal
      .Borders(xlInsideVertical).Weight = WeightVal
      .Borders(xlInsideHorizontal).Weight = WeightVal

      If Not (IsMissing(Red)) Then
        .Borders(xlEdgeLeft).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeRight).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeTop).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeBottom).Color = RGB(Red, Green, Blue)
        .Borders(xlInsideVertical).Color = RGB(Red, Green, Blue)
        .Borders(xlInsideHorizontal).Color = RGB(Red, Green, Blue)
      End If
    End With
  End If
End Function

'--------------------------------------------------------------------------------------------------
Function rü_Àü_¶E(Optional setArea As Range, Optional lineColor As Long, Optional WeightVal = xlThin)
  Dim Red As Long, Green As Long, Blue As Long

  Call Library.getRGB(lineColor, Red, Green, Blue)

  If TypeName(setArea) = "Range" Then
    With setArea
      .Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Borders(xlEdgeRight).LineStyle = xlContinuous
      
      .Borders(xlEdgeLeft).Weight = WeightVal
      .Borders(xlEdgeRight).Weight = WeightVal

      If Not (IsMissing(Red)) Then
        .Borders(xlEdgeLeft).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeRight).Color = RGB(Red, Green, Blue)
      End If
     End With
  Else
    With Selection
      .Borders(xlEdgeLeft).LineStyle = xlContinuous
      .Borders(xlEdgeRight).LineStyle = xlContinuous
      
      .Borders(xlInsideHorizontal).Weight = WeightVal
      .Borders(xlEdgeRight).Weight = WeightVal

      If Not (IsMissing(Red)) Then
        .Borders(xlEdgeLeft).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeRight).Color = RGB(Red, Green, Blue)
      End If
     End With

  End If
End Function

'--------------------------------------------------------------------------------------------------
Function rü_Àü_ãº(Optional setArea As Range, Optional lineColor As Long, Optional WeightVal = xlThin)
  Dim Red As Long, Green As Long, Blue As Long

  Call Library.getRGB(lineColor, Red, Green, Blue)

  If TypeName(setArea) = "Range" Then
    With setArea
      .Borders(xlEdgeTop).LineStyle = xlContinuous
      .Borders(xlEdgeBottom).LineStyle = xlContinuous

      .Borders(xlEdgeTop).Weight = WeightVal
      .Borders(xlEdgeBottom).Weight = WeightVal

      If Not (IsMissing(Red)) Then
        .Borders(xlEdgeTop).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeBottom).Color = RGB(Red, Green, Blue)
      End If
    End With
  Else
    With Selection
      .Borders(xlEdgeTop).LineStyle = xlContinuous
      .Borders(xlEdgeBottom).LineStyle = xlContinuous

      .Borders(xlEdgeTop).Weight = WeightVal
      .Borders(xlEdgeBottom).Weight = WeightVal

      If Not (IsMissing(Red)) Then
        .Borders(xlEdgeTop).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeBottom).Color = RGB(Red, Green, Blue)
      End If
    End With
  End If
End Function

'--------------------------------------------------------------------------------------------------
Function rü_Àü_¼(Optional setArea As Range, Optional lineColor As Long, Optional WeightVal = xlThin)
  Dim Red As Long, Green As Long, Blue As Long

  Call Library.getRGB(lineColor, Red, Green, Blue)

  If TypeName(setArea) = "Range" Then
    With setArea
      .Borders(xlInsideVertical).LineStyle = xlContinuous
      .Borders(xlInsideVertical).Weight = WeightVal
      
      If Not (IsMissing(Red)) Then
        .Borders(xlInsideVertical).Color = RGB(Red, Green, Blue)
      End If
    End With
  Else
    With Selection
      .Borders(xlInsideVertical).LineStyle = xlDash
      .Borders(xlInsideVertical).Weight = WeightVal
      
      If Not (IsMissing(Red)) Then
        .Borders(xlInsideVertical).Color = RGB(Red, Green, Blue)
      End If
    End With
  End If
End Function

'--------------------------------------------------------------------------------------------------
Function rü_Àü_½(Optional setArea As Range, Optional lineColor As Long, Optional WeightVal = xlThin)
  Dim Red As Long, Green As Long, Blue As Long

  Call Library.getRGB(lineColor, Red, Green, Blue)

  If TypeName(setArea) = "Range" Then
    With setArea
      .Borders(xlInsideHorizontal).LineStyle = xlContinuous
      .Borders(xlInsideHorizontal).Weight = WeightVal

      If Not (IsMissing(Red)) Then
        .Borders(xlInsideHorizontal).Color = RGB(Red, Green, Blue)
      End If
     End With
  Else

    With Selection
      .Borders(xlInsideHorizontal).LineStyle = xlDash
      .Borders(xlInsideHorizontal).Weight = WeightVal

      If Not (IsMissing(Red)) Then
        .Borders(xlInsideHorizontal).Color = RGB(Red, Green, Blue)
      End If
    End With

  End If
End Function


'--------------------------------------------------------------------------------------------------
Function rü_ñdü_¶(Optional setArea As Range, Optional lineColor As Long, Optional WeightVal = xlThin)
  Dim Red As Long, Green As Long, Blue As Long

  Call Library.getRGB(lineColor, Red, Green, Blue)

  If TypeName(setArea) = "Range" Then
    With setArea
      .Borders(xlEdgeLeft).LineStyle = xlDouble
      
      If Not (IsMissing(Red)) Then
        .Borders(xlEdgeLeft).Color = RGB(Red, Green, Blue)
      End If
    End With
  Else
    With Selection
      .Borders(xlEdgeLeft).LineStyle = xlDouble
  
      If Not (IsMissing(Red)) Then
        .Borders(xlEdgeLeft).Color = RGB(Red, Green, Blue)
      End If
    End With
  End If
End Function

'--------------------------------------------------------------------------------------------------
Function rü_ñdü_¶E(Optional setArea As Range, Optional lineColor As Long, Optional WeightVal = xlThin)
  Dim Red As Long, Green As Long, Blue As Long

  Call Library.getRGB(lineColor, Red, Green, Blue)

  If TypeName(setArea) = "Range" Then
    With setArea
      .Borders(xlEdgeLeft).LineStyle = xlDouble
      .Borders(xlEdgeRight).LineStyle = xlDouble
      
      If Not (IsMissing(Red)) Then
        .Borders(xlEdgeLeft).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeRight).Color = RGB(Red, Green, Blue)
      End If
    End With
  Else
    With Selection
      .Borders(xlEdgeLeft).LineStyle = xlDouble
      .Borders(xlEdgeRight).LineStyle = xlDouble
  
      If Not (IsMissing(Red)) Then
        .Borders(xlEdgeLeft).Color = RGB(Red, Green, Blue)
        .Borders(xlEdgeRight).Color = RGB(Red, Green, Blue)
      End If
    End With
  End If
End Function


'--------------------------------------------------------------------------------------------------
Function rü_ñdü_º(Optional setArea As Range, Optional lineColor As Long, Optional WeightVal = xlThin)
  Dim Red As Long, Green As Long, Blue As Long

  Call Library.getRGB(lineColor, Red, Green, Blue)

  If TypeName(setArea) = "Range" Then
    With setArea
      .Borders(xlEdgeBottom).LineStyle = xlDouble
  
      If Not (IsMissing(Red)) Then
        .Borders(xlEdgeBottom).Color = RGB(Red, Green, Blue)
      End If
    End With
  Else
    With Selection
      .Borders(xlEdgeBottom).LineStyle = xlDouble
  
      If Not (IsMissing(Red)) Then
        .Borders(xlEdgeBottom).Color = RGB(Red, Green, Blue)
      End If
    End With
  End If
End Function





'--------------------------------------------------------------------------------------------------
Function rü_jü_tL(Optional setArea As Range, Optional lineColor As Long, Optional WeightVal = xlThin)
  Dim Red As Long, Green As Long, Blue As Long
  
  Call rü_jü_ÍÝ(setArea, lineColor, WeightVal)
  Call Library.getRGB(lineColor, Red, Green, Blue)

  If TypeName(setArea) = "Range" Then
    Set setArea = setArea.Offset(1, 1).Resize(setArea.Rows.count - 1, setArea.Columns.count - 1)
    Call rü_jü_½(setArea, lineColor, WeightVal)
    Call rü_jü_ÍÝ(setArea, lineColor, WeightVal)
  Else
    setArea.Offset(1, 1).Resize(setArea.Rows.count - 1, setArea.Columns.count - 1).Select
    Call rü_jü_½(setArea, lineColor, WeightVal)
    Call rü_jü_ÍÝ(setArea, lineColor, WeightVal)
  
  End If
End Function

