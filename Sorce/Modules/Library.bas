Attribute VB_Name = "Library"
'**************************************************************************************************
' * �Q�Ɛݒ�A�萔�錾
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
' ���p����Q�Ɛݒ�܂Ƃ�
' Microsoft Office 14.0 Object Library
' Microsoft DAO 3.6 Objects Library
' Microsoft Scripting Runtime (WSH, FileSystemObject)
' Microsoft ActiveX Data Objects 2.8 Library
' UIAutomationClient

' Windows API�̗��p--------------------------------------------------------------------------------
' �f�B�X�v���C�̉𑜓x�擾�p
' Sleep�֐��̗��p
' �N���b�v�{�[�h�֐��̗��p
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



'���[�N�u�b�N�p�ϐ�------------------------------
'���[�N�V�[�g�p�ϐ�------------------------------
'�O���[�o���ϐ�----------------------------------
Public LibDAO As String
Public LibADOX As String
Public LibADO As String
Public LibScript As String

'�A�N�e�B�u�Z���̎擾
Dim SelectionCell As String

' PC�AOffice���̏��擾�p�A�z�z��
Public MachineInfo As Object

' Selenium�p�ݒ�
Public Const HalfWidthDigit = "1234567890"
Public Const HalfWidthCharacters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
Public Const SymbolCharacters = "!""#$%&'()=~|@[`{;:]+*},./\<>?_-^\"

'Public Const JapaneseCharacters = "�����������������������������������ĂƂȂɂʂ˂̂͂Ђӂւق܂݂ނ߂�������������񂪂����������������������Âłǂ΂тԂׂڂς҂Ղ؂�"
'Public Const JapaneseCharactersCommonUse = "�J�w����щ�⋞���o�m�����X�����������ψ�j�݋��K�n�g�����Ґ̎�󏊒���g�\�������������a�p�ʉ芯�G�����a�ō��Q����������I�T�ŔO�{�@�q��Չ����͋������Ȏ}�ɏq���������Ŕ�񕐉����g���č���@���S�����͓��q���󖇈ˉ���F���������h���������˔t�������ޕ|���b�Ή�����"
'Public Const MachineDependentCharacters = "�@�A�B�C�D�E�F�G�H�I�J�K�L�M�N�O�P�Q�R�S�T�U�V�W�X�Y�Z�[�\�]�_�\�]�^�_�`�a�b�c�d�e�f�g�h�i�j�k�l�m�n�o�p�q�r�s�t�u�v�w�x�y�z�{"


Public ThisBook As Workbook


'**************************************************************************************************
' * �A�h�I�������
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function addinClose()

  Workbooks(ThisWorkbook.Name).Close
End Function


'**************************************************************************************************
' * �G���[���̏���
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function errorHandle(funcName As String, ByRef objErr As Object)
  
  Dim message As String
  Dim runTime As Date
  Dim endLine As Long
  
  runTime = Format(Now(), "yyyy/mm/dd hh:nn:ss")
  message = funcName & vbCrLf & objErr.Description

  '�����F�����b
  Application.Speech.Speak Text:="�G���[���������܂���", SpeakAsync:=True
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
' * ��ʕ`�ʐ���J�n
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function startScript()

  '�A�N�e�B�u�Z���̎擾
  If TypeName(Selection) = "Range" Then
    SelectionCell = Selection.Address
  End If

  '�}�N������ŃV�[�g��E�B���h�E���؂�ւ��̂������Ȃ��悤�ɂ��܂�
  Application.ScreenUpdating = False

  '�}�N�����쎩�̂ŕʂ̃C�x���g�����������̂�}������
  Application.EnableEvents = False

  '�}�N������ŃZ��ItemName�Ȃǂ��ς�鎞�����v�Z��������x������̂������
  Application.Calculation = xlCalculationManual

  '�}�N�����쒆�Ɉ�؂̃L�[��}�E�X����𐧌�����
  'Application.Interactive = False

  '�}�N�����쒆�̓}�E�X�J�[�\�����u�����v�v�ɂ���
  'Application.Cursor = xlWait

  '�m�F���b�Z�[�W���o���Ȃ�
  Application.DisplayAlerts = False

End Function

'**************************************************************************************************
' * ��ʕ`�ʐ���I��
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function endScript(Optional flg As Boolean = False)

  '�}�N������ŃV�[�g��E�B���h�E���؂�ւ��̂������Ȃ��悤�ɂ��܂�
  Application.ScreenUpdating = True

  '�}�N�����쎩�̂ŕʂ̃C�x���g�����������̂�}������
  Application.EnableEvents = True

  '�}�N������ŃZ��ItemName�Ȃǂ��ς�鎞�����v�Z��������x������̂������
  Application.Calculation = xlCalculationAutomatic

  '�}�N�����쒆�Ɉ�؂̃L�[��}�E�X����𐧌�����
  'Application.Interactive = True

  '�}�N������I����̓}�E�X�J�[�\�����u�f�t�H���g�v�ɂ��ǂ�
  'Application.Cursor = xlDefault

  '�}�N������I����̓X�e�[�^�X�o�[���u�f�t�H���g�v�ɂ��ǂ�
  Application.StatusBar = False
  'Application.StatusBar = "���:" & setVal("baseDay")

  '�m�F���b�Z�[�W���o���Ȃ�
  Application.DisplayAlerts = True

  '�����I�ɍČv�Z������
  Application.CalculateFull

 '�A�N�e�B�u�Z���̑I��
  If SelectionCell <> "" And flg = True Then
    Range(SelectionCell).Select
  End If
  
  Call unsetClipboard
End Function



'**************************************************************************************************
' * �V�[�g�̑��݊m�F
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
' * ���O�V�[�g����
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function chkExcludeSheet(sheetName As String, colNo As Long)

  Dim endBookRowLine As Long
  Dim rowLine As Long
  Dim Result As Boolean

  Result = True

  ' �ŏI�s�擾
  endBookRowLine = Sheets("�ݒ�").Cells(Rows.count, colNo).End(xlUp).row
  For rowLine = 3 To endBookRowLine
    If sheetName = Sheets("�ݒ�").Cells(rowLine, colNo) Then
      Result = False
      Exit For
    End If
  Next
  CheckExcludeSheet = Result
End Function


'**************************************************************************************************
' * �u�b�N���J����Ă��邩�`�F�b�N
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
' * �t�@�C���̑��݊m�F
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
' * �f�B���N�g���̑��݊m�F
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
' * �L�������P�[�X���X�l�[�N�P�[�X�ɕϊ�
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
' * �X�l�[�N�P�[�X���L�������P�[�X�ɕϊ�
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
' * ���p�̃J�^�J�i��S�p�̃J�^�J�i�ɕϊ�����(�������p�����͔��p�ɂ���)
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
    If Mid(rData, i, 1) Like "[�`-��]" Or Mid(rData, i, 1) Like "[�O-�X]" Or Mid(rData, i, 1) Like "�|" Then
      ansData = ansData & StrConv(Mid(rData, i, 1), vbNarrow)
    Else
      ansData = ansData & Mid(rData, i, 1)
    End If
  Next i
  convHan2Zen = ansData
End Function


'**************************************************************************************************
' * �p�C�v���J���}�ɕϊ�
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
' * Base64�G���R�[�h(�t�@�C��)
' *
' * @link   http://www.ka-net.org/blog/?p=4524
'**************************************************************************************************
Function convBase64EncodeForFile(ByVal filePath As String) As String
  Dim elm As Object
  Dim ret As String
  Const adTypeBinary = 1
  Const adReadAll = -1

  ret = "" '������
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
' * Base64�G���R�[�h(������)
' *
' * @link   http://www.ka-net.org/blog/?p=4524
'**************************************************************************************************
Function convBase64EncodeForString(ByVal str As String) As String

  Dim ret As String
  Dim d() As Byte

  Const adTypeBinary = 1
  Const adTypeText = 2

  ret = "" '������
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
' * URL-safe Base64�G���R�[�h
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
' * URL�G���R�[�h
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
' * �擪�P�����ڂ�啶����
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
' * ������̍�������w�蕶�����폜����֐�
' *
' * @Link   https://vbabeginner.net/vba�ŕ�����̉E���⍶������w�蕶�����폜����/
'**************************************************************************************************
Function cutLeft(s, i As Long) As String
    Dim iLen    As Long     '������

    '������ł͂Ȃ��ꍇ
    If VarType(s) <> vbString Then
        cutLeft = s & "������ł͂Ȃ�"
        Exit Function
    End If

    iLen = Len(s)

    '�����񒷂��w�蕶�������傫���ꍇ
    If iLen < i Then
        cutLeft = s & "�����񒷂��w�蕶�������傫��"
        Exit Function
    End If

    '�w�蕶�������폜���ĕԂ�
    cutLeft = Right(s, iLen - i)
End Function


'**************************************************************************************************
' * ������̉E������w�蕶�����폜����֐�
' *
' * @Link   https://vbabeginner.net/vba�ŕ�����̉E���⍶������w�蕶�����폜����/
'**************************************************************************************************
Function cutRight(s, i As Long) As String
    Dim iLen    As Long     '������

    '������ł͂Ȃ��ꍇ
    If VarType(s) <> vbString Then
      cutRight = s & "������ł͂Ȃ�"
      Exit Function
    End If

    iLen = Len(s)

    '�����񒷂��w�蕶�������傫���ꍇ
    If iLen < i Then
      cutRight = s & "�����񒷂��w�蕶�������傫��"
      Exit Function
    End If

    '�w�蕶�������폜���ĕԂ�
    cutRight = Left(s, iLen - i)
End Function


'**************************************************************************************************
' * �A�����s�̍폜
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
' *���W�X�g��������擾
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
' * �V�[�g�폜
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
  Cells.NumberFormatLocal = "G/�W��"

  Application.Goto Reference:=Range("A1"), Scroll:=True
End Function


'**************************************************************************************************
' * �I��͈͂̉摜�폜
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
' * �Z���̖��̐ݒ�폜
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
' * �e�[�u���f�[�^�폜
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
  Selection.NumberFormatLocal = "G/�W��"

  Application.Goto Reference:=Range("A1"), Scroll:=True
End Function


'**************************************************************************************************
' * PC�AOffice���̏��擾
' * �A�z�z��𗘗p���Ă���̂ŁAMicrosoft Scripting Runtime���K�{
' * MachineInfo.Item ("Excel") �ŌĂяo��
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function getMachineInfo() As Object
  Dim WshNetworkObject As Object

  Set MachineInfo = CreateObject("Scripting.Dictionary")
  Set WshNetworkObject = CreateObject("WScript.Network")

  ' OS�̃o�[�W�����擾-----------------------------------------------------------------------------
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

  ' Excel�̃o�[�W�����擾--------------------------------------------------------------------------
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

  'PC�̏��----------------------------------------------------------------------------------------
  MachineInfo.Add "UserName", WshNetworkObject.UserName
  MachineInfo.Add "ComputerName", WshNetworkObject.ComputerName
  MachineInfo.Add "UserDomain", WshNetworkObject.UserDomain

  '��ʂ̉𑜓x���擾------------------------------------------------------------------------------
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
' * �������J�E���g
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
' * �񖼂����ԍ������߂�
' *
' * @link   http://www.happy2-island.com/excelsmile/smile03/capter00717.shtml
'**************************************************************************************************
Function getColumnNo(targetCell As String) As Long

  getColumnNo = Range(targetCell & ":" & targetCell).Column
End Function


'**************************************************************************************************
' * ��ԍ�����񖼂����߂�
' *
' * @link   http://www.happy2-island.com/excelsmile/smile03/capter00717.shtml
'**************************************************************************************************
Function getColumnName(targetCell As Long) As String

  getColumnName = Split(Cells(, targetCell).Address, "$")(1)
End Function

'**************************************************************************************************
' * �J���[�p���b�g��\�����A�F�R�[�h���擾
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
' * IndentLevel�l�擾
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
' * RGB�l�擾
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function getRGB(colorValue As Long, Red As Long, Green As Long, Blue As Long)
  Red = colorValue Mod 256
  Green = Int(colorValue / 256) Mod 256
  Blue = Int(colorValue / 256 / 256)
End Function


'**************************************************************************************************
' *���W�X�g��������擾
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function getRegistry(registryName As String)
  Dim regVal As String

  regVal = GetSetting("ExcelHelp", RegistrySubKey, registryName)
  getRegistry = regVal
End Function


'**************************************************************************************************
' * �f�B���N�g���I���_�C�A���O�\��
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function getDirPath(CurrentDirectory As String, Optional title As String)

  With Application.FileDialog(msoFileDialogFolderPicker)
    .InitialFileName = CurrentDirectory & "\"
    .AllowMultiSelect = False
    .title = title & "�̕ۑ��ꏊ��I�����Ă�������"
    If .Show = True Then
      getDirPath = .SelectedItems(1)
    Else
      getDirPath = ""
    End If
  End With
End Function


'**************************************************************************************************
' * �t�@�C���ۑ��_�C�A���O�\��
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function getSaveFilePath(CurrentDirectory As String, saveFileName As String, FileTypeNo As Long)

  Dim filePath As String
  Dim Result As Long

  Dim FileName As Variant
  FileName = Application.GetSaveAsFilename( _
      InitialFileName:=CurrentDirectory & "\" & saveFileName, _
      FileFilter:="Excel�t�@�C��,*.xlsx,Excel2003�ȑO,*.xls,Excel�}�N���u�b�N,*.xlsm", _
      FilterIndex:=FileTypeNo)

  getSaveFilePath = FileName
End Function

'**************************************************************************************************
' * �t�@�C���I���_�C�A���O�\��
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function getFilePath(CurrentDirectory As String, saveFileName As String, title As String, FileTypeNo As Long)

  Dim filePath As String
  Dim Result As Long

  With Application.FileDialog(msoFileDialogFilePicker)

    ' �t�@�C���̎�ނ�ݒ�
    .Filters.Clear
    .Filters.Add "Excel�u�b�N", "*.xls; *.xlsx; *.xlsm"
    .Filters.Add "CSV�t�@�C��", "*.csv"
    .Filters.Add "SQL�t�@�C��", "*.sql"
    .Filters.Add "�e�L�X�g�t�@�C��", "*.txt"
    .Filters.Add "JSON�t�@�C��", "*.json"
    .Filters.Add "Accesss�f�[�^�x�[�X", "*.mdb"
    .Filters.Add "���ׂẴt�@�C��", "*.*"

    .FilterIndex = FileTypeNo

    '�\������t�H���_
    .InitialFileName = CurrentDirectory & "\" & saveFileName

    '�\���`���̐ݒ�
    .InitialView = msoFileDialogViewWebView

    '�_�C�A���O �{�b�N�X�̃^�C�g���ݒ�
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
' * �����t�@�C���I���_�C�A���O�\��
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function getFilesPath(CurrentDirectory As String, saveFileName As String, title As String, FileTypeNo As Long)

  Dim filePath() As Variant
  Dim Result As Long

  With Application.FileDialog(msoFileDialogFilePicker)
    '�����I��������
    .AllowMultiSelect = True
    
    ' �t�@�C���̎�ނ�ݒ�
    .Filters.Clear
    .Filters.Add "Excel�u�b�N", "*.xls; *.xlsx; *.xlsm"
    .Filters.Add "CSV�t�@�C��", "*.csv"
    .Filters.Add "SQL�t�@�C��", "*.sql"
    .Filters.Add "�e�L�X�g�t�@�C��", "*.txt"
    .Filters.Add "JSON�t�@�C��", "*.json"
    .Filters.Add "Accesss�f�[�^�x�[�X", "*.mdb"
    .Filters.Add "���ׂẴt�@�C��", "*.*"

    .FilterIndex = FileTypeNo

    '�\������t�H���_
    .InitialFileName = CurrentDirectory & "\" & saveFileName

    '�\���`���̐ݒ�
    .InitialView = msoFileDialogViewWebView
    
    '�_�C�A���O �{�b�N�X�̃^�C�g���ݒ�
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
' * �f�B���N�g�����̃t�@�C���ꗗ�擾
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
' * �w��o�C�g���̌Œ蒷�f�[�^�쐬(�����񏈗�)
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
    ' �o�C�g������
    For IX = 1 To lngKeta
        ' 1���������p/�S�p�𔻒f
        intCHAR = Asc(Mid(strInText, IX, 1))
        ' �S�p�Ɣ��f�����ꍇ�̓o�C�g����1��������
        If ((intCHAR < 0) Or (intCHAR > 255)) Then
            lngByte2 = 2        ' �S�p
        Else
            lngByte2 = 1        ' ���p
        End If
        ' �����ӂꔻ��(�E�؂�̂�)
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
    ' ���s������(�󔒕����ǉ�)
    If lngByte < lngFixBytes Then
        strOutText = strOutText & Space(lngFixBytes - lngByte)
    End If
    getFixlng = strOutText
End Function


'**************************************************************************************************
' * �V�[�g���X�g�擾
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

  '���ݒ�l�̃N���A
  Worksheets("�ݒ�").Range(ColumnName & "3:" & ColumnName & "100").Select
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

    '�V�[�g���̐ݒ�
    Worksheets("�ݒ�").Range(ColumnName & i).Select
    Worksheets("�ݒ�").Range(ColumnName & i) = sheetName.Name

    ' �Z���̔w�i�F����
    With Worksheets("�ݒ�").Range(ColumnName & i).Interior
      .Pattern = xlPatternNone
      .Color = xlNone
    End With

    ' �V�[�g�F�Ɠ����F���Z���ɐݒ�
    If Worksheets(sheetName.Name).Tab.Color Then
      With Worksheets("�ݒ�").Range(ColumnName & i).Interior
        .Pattern = xlPatternNone
        .Color = Worksheets(sheetName.Name).Tab.Color
      End With
    End If

    '�r���̐ݒ�
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

  Worksheets("�ݒ�").Range(ColumnName & "3").Select
  Call endScript
  Exit Function
'--------------------------------------------------------------------------------------------------
'�G���[�������̏���
'--------------------------------------------------------------------------------------------------
GetSheetListError:

  ' ��ʕ`�ʐ���I��
  Call endScript
  Call errorHandle("�V�[�g���X�g�擾", Err)

End Function


'**************************************************************************************************
' * �I���Z���̊g��\���ďo
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
' * �I���Z���̊g��\���I��
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function showExpansionFormClose(Text As String, SetSelectTargetRows As String)
  Range(SetSelectTargetRows).Value = Text
  Call endScript
End Function


'**************************************************************************************************
' * �f�o�b�O�p��ʕ\��
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
    
    meg1 = meg1 & String(10 - Len(meg1), "�@")
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
  If meg2 = "�����J�n" Then
    With debugForm
      .Caption = "�������"
      .ListBox1.Clear
      .ListBox1.AddItem runTime & vbTab & meg1 & vbTab & meg2

    End With
  Else
    With debugForm
      .Caption = "�������"
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
' * �������ʒm
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

  '��ʕ`�ʐ���I������
  If runEndflg = True Then
    Call endScript
    Call ProgressBar.showEnd
    End
  End If
End Function


'**************************************************************************************************
' * �����_��
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function makeRandomString(ByVal setString As String, ByVal setStringCnt As Integer) As String

 For i = 1 To setStringCnt
    '�����W�F�l���[�^��������
    Randomize
    n = Int((Len(setString) - 1 + 1) * Rnd + 1)
    str1 = str1 + Mid(setString, n, 1)
  Next i

  makeRandomString = str1

End Function

Function makeRandomNo(minNo As Long, maxNo As Long) As String

  '�����W�F�l���[�^��������
  Randomize
  makeRandomNo = Application.RoundDown(Int((maxNo - minNo + 1) * Rnd + minNo), -2)

End Function

'**************************************************************************************************
' * ���O�o��
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
' * CSV�C���|�[�g
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
    .TextFilePlatform = 932          ' Shift-JIS ���J��
    .TextFileParseType = xlDelimited ' �����ŋ�؂����`��
    .TextFileCommaDelimiter = True   ' ��؂蕶���̓J���}
    .TextFileStartRow = readLine     ' 1 �s�ڂ���ǂݍ���
    .AdjustColumnWidth = False       ' �񕝂������������Ȃ�
    .RefreshStyle = xlOverwriteCells '�㏑�����w��
    .TextFileTextQualifier = xlTextQualifierDoubleQuote ' ���p���̎w��
    
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
' * Excel�t�@�C���̃C���|�[�g
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
' * MkDir�ŊK�w�̐[���t�H���_�[�����
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
  tmpPath = arr(0)  ' �h���C�u���̑��

  For i = 1 To UBound(arr)
    tmpPath = tmpPath & "\" & arr(i)
    If Dir(tmpPath, vbDirectory) = "" Then
      MkDir tmpPath
    End If
  Next i

End Function


'**************************************************************************************************
' * �p�X���[�h����
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function makePasswd() As String
  Dim halfChar As String, str1 As String

  halfChar = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz!#$%&"

  For i = 1 To 12
    '�����W�F�l���[�^��������
    Randomize
    n = Int((Len(halfChar) - 1 + 1) * Rnd + 1)
    str1 = str1 + Mid(halfChar, n, 1)
  Next i
  makePasswd = str1
End Function


'**************************************************************************************************
' * �����񕪊�
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
' * �z��̍Ō�ɒǉ�����
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
' * �t�H���g�J���[�ݒ�
' *
' * @Link https://vbabeginner.net/vba�ŃZ���̎w�蕶����̐F�⑾����ύX����/
'**************************************************************************************************
Function setFontClor(a_sSearch, a_lColor, a_bBold)
  Dim f   As Font     'Font�I�u�W�F�N�g
  Dim i               '����������̃Z���̈ʒu
  Dim iLen            '����������̕�����
  Dim r   As Range    '�Z���͈͂̂P�Z��

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
' * ���W�X�g���ɏ��o�^
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function setRegistry(registryName As String, setVal As Variant)
  
  If getRegistry(registryName) <> setVal Then
    SaveSetting "ExcelHelp", RegistrySubKey, registryName, setVal
  End If
End Function


'**************************************************************************************************
' * �Q�Ɛݒ�������ōs��
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
      MsgBox ("Microsoft Scripting Runtime�𗘗p�ł��܂���B" & vbLf & "���p�ł��Ȃ��@�\������܂�")
    End If
    
  'Microsoft ActiveX Data Objects Library 6.1 (ADO)------------------------------------------------
  If BookType = "DataBase" Then
    LibADO = "C:\Program Files\Common Files\System\Ado\msado15.dll"
    If Dir(LibADO) <> "" Then
      ActiveWorkbook.VBProject.References.AddFromFile (LibADO)
    Else
      MsgBox ("Microsoft ActiveX Data Objects�𗘗p�ł��܂���" & vbLf & "���p�ł��Ȃ��@�\������܂�")
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
        MsgBox ("Microsoft DAO 3.6 Objects Library�𗘗p�ł��܂���" & vbLf & "DB�ւ̐ڑ��@�\�����p�ł��܂���")
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
        MsgBox ("Microsoft DAO 3.6 Objects Library�𗘗p�ł��܂���" & vbLf & "DB�ւ̐ڑ��@�\�����p�ł��܂���")
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
    MsgBox ("�uVBA �v���W�F�N�g �I�u�W�F�N�g ���f���ւ̃A�N�Z�X��M������v�ɕύX���I")
  Else
    MsgBox "Error Number : " & Err.Number & vbCrLf & Err.Description
    GoTo Func_Exit:
  End If
End Function


'**************************************************************************************************
' * �I���Z���̍s�w�i�ݒ�
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function setLineColor(SetArea As String, DisType As Boolean, SetColor As String)

  Range(SetArea).Select

  '�����t���������N���A
  Selection.FormatConditions.Delete

  If DisType = False Then
    '�s�����ݒ�
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:="=CELL(""row"")=ROW()"
  Else
    '�s�Ɨ�ɐݒ�
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
' * �ŏ��̃V�[�g��I��
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
' * �t�@�C���S�̂̕�����u��
' *
' * @Link   https://www.moug.net/tech/acvba/0090005.html
'**************************************************************************************************
Function replaceFromFile(FileName As String, TargetText As String, Optional NewText As String = "")

 Dim FSO         As FileSystemObject '�t�@�C���V�X�e���I�u�W�F�N�g
 Dim Txt         As TextStream       '�e�L�X�g�X�g���[���I�u�W�F�N�g
 Dim buf_strTxt  As String           '�ǂݍ��݃o�b�t�@

 On Error GoTo Func_Err:

 '�I�u�W�F�N�g�쐬
 Set FSO = CreateObject("Scripting.FileSystemObject")
 Set Txt = FSO.OpenTextFile(FileName, ForReading)

 '�S���ǂݍ���
  buf_strTxt = Txt.ReadAll
  Txt.Close

  '���t�@�C�������l�[�����āA�e���|�����t�@�C���쐬
  Name FileName As FileName & "_"

  '�u������
   buf_strTxt = Replace(buf_strTxt, TargetText, NewText, , , vbBinaryCompare)

  '�����ݗp�e�L�X�g�t�@�C���쐬
   Set Txt = FSO.CreateTextFile(FileName, True)
  '������
  Txt.Write buf_strTxt
  Txt.Close

  '�e���|�����t�@�C�����폜
  FSO.DeleteFile FileName & "_"

'�I������
Func_Exit:
    Set Txt = Nothing
    Set FSO = Nothing
    Exit Function

Func_Err:
    MsgBox "Error Number : " & Err.Number & vbCrLf & Err.Description
    GoTo Func_Exit:
End Function


'**************************************************************************************************
' * VBA��Excel�̃R�����g���ꊇ�Ŏ����T�C�Y�ɂ��ăJ�b�R�悭����
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
          ' �T�C�Y�����ݒ�
          .TextFrame.AutoSize = True
          .TextFrame.Characters.Font.Size = 9

          ' �`����p�ێl�p�`�ɕύX
          .AutoShapeType = msoShapeRectangle
          ' �h��F�E���F �ύX
          .line.ForeColor.RGB = RGB(128, 128, 128)
          .Fill.ForeColor.RGB = RGB(240, 240, 240)
          ' �e ���ߗ� 30%�A�I�t�Z�b�g�� x:1px,y:1px
          .Shadow.Transparency = 0.3
          .Shadow.OffsetX = 1
          .Shadow.OffsetY = 1
          ' ���������A��������
          .TextFrame.Characters.Font.Bold = False
          .TextFrame.HorizontalAlignment = xlLeft
          ' �Z���ɍ��킹�Ĉړ�����
          .Placement = xlMove
        End With
      End If
    Next cl
    Application.Goto Reference:=Range("A1"), Scroll:=True
End Function



'**************************************************************************************************
' * �N���b�v�{�[�h�N���A
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function unsetClipboard()
'  OpenClipboard 0
'  EmptyClipboard
'  CloseClipboard
End Function


'**************************************************************************************************
' * �I���Z���̍s�w�i����
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function unsetLineColor(SetArea As String)
  ActiveSheet.Range(SetArea).Select

  '�����t���������N���A
  Selection.FormatConditions.Delete
'  Application.GoTo Reference:=Range("A1"), Scroll:=True
End Function


'**************************************************************************************************
' * �����N����
' *
' * @Link   https://excel-excellent-technics.com/excel-vba-breaklinks-1019
'**************************************************************************************************
Function unsetLink()
  Dim wb          As Workbook
  Dim vntLink     As Variant
  Dim i           As Integer

  Set wb = ActiveWorkbook
  vntLink = wb.LinkSources(xlLinkTypeExcelLinks) '�u�b�N�̒��ɂ��郊���N

  If IsArray(vntLink) Then
    For i = 1 To UBound(vntLink)
      wb.BreakLink vntLink(i), xlLinkTypeExcelLinks '�����N����
    Next i
  End If
End Function


'**************************************************************************************************
' * �X���[�v����
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function waitTime(timeVal As Long)
  Sleep timeVal
  DoEvents
End Function


'**************************************************************************************************
' * TEXTJOIN�֐�
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
' * ���{���^�u�̑I��
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

  Set elmRibbonTab = Nothing '������
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



