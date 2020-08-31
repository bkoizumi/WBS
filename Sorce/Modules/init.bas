Attribute VB_Name = "init"
'���[�N�u�b�N�p�ϐ�------------------------------
Public ThisBook As Workbook

'���[�N�V�[�g�p�ϐ�------------------------------
Public noticeCodeSheet As Worksheet
Public helpSheet As Worksheet
Public setSheet As Worksheet
Public mainSheet As Worksheet
Public tmpSheet As Worksheet



'�O���[�o���ϐ�----------------------------------
Public setVal As Collection
Public getVal As Collection
Public memberColor As Object
Public debugMode As String

'Public lineColor As String
'Public SaturdayColor As String
'Public SundayColor As String
'Public CompanyHolidayColor As String
'
'Public taskLevel1Color As String
'Public taskLevel2Color As String
'Public taskLevel3Color As String

'���W�X�g���o�^�p�T�u�L�[
Public Const RegistrySubKey As String = "WBS"

'���O�t�@�C��
Public logFile As String

'�K���g�`���[�g�I��
Public selectShapesName(0) As Variant
Public changeShapesName As String





'***********************************************************************************************************************************************
' * �ݒ�
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'***********************************************************************************************************************************************
Function setting()
  Dim line As Long

  If debugMode <> "" Then
    Set setVal = Nothing
  End If
  
  Set setVal = New Collection
  Set memberColor = CreateObject("Scripting.Dictionary")
  
  '�u�b�N�̐ݒ�
  Set ThisBook = ThisWorkbook
  ThisBook.Activate
  
  '���[�N�V�[�g���̐ݒ�
  Set helpSheet = ThisBook.Worksheets("Help")
  Set noticeCodeSheet = ThisBook.Worksheets("Notice")
  Set setSheet = ThisBook.Worksheets("�ݒ�")
  Set mainSheet = ThisBook.Worksheets("WBS")
  Set tmpSheet = ThisBook.Worksheets("Tmp")
  
  '���ԁA����������͎��̏����l
  Select Case True
    Case setSheet.Range("B7") = ""
      setSheet.Range("B7") = Format(DateAdd("d", 0, Date), "yyyy/mm/dd")
    
    Case setSheet.Range("B8") = ""
      setSheet.Range("A8") = Format(DateAdd("d", 30, Date), "yyyy/mm/dd")
    
    Case setSheet.Range("B9") = ""
      setSheet.Range("B9") = Format(DateAdd("d", 0, Date), "yyyy/mm/dd")
  End Select
  
  '�ݒ�l�̓ǂݍ���
  With setVal
    For line = 3 To setSheet.Cells(Rows.count, 1).End(xlUp).row
      If setSheet.Range("A" & line) <> "" Then
       .Add item:=setSheet.Range("B" & line), Key:=setSheet.Range("A" & line)
      End If
    Next
  End With
  debugMode = setVal("debugMode")
  
  '�V���[�g�J�b�g�L�[�̐ݒ�ǉ�
  With setVal
    For line = 3 To setSheet.Cells(Rows.count, 7).End(xlUp).row
      .Add item:=setSheet.Range("I" & line), Key:=setSheet.Range("H" & line)
    Next
  End With
  

  '�S���ҐF�ǂݍ���
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
  
  Call ���O��`
End Function

'***********************************************************************************************************************************************
' * �x���ݒ�
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'***********************************************************************************************************************************************
Function chkHollyday(chkDate As Date, HollydayName As String, flg As Boolean)
  Dim line As Long, endLine As Long
  Dim strFilMessage() As Date
  
  '�x������
  Call GetHollyday(CDate(chkDate), HollydayName)
  
  '�y���𔻒�
  If Weekday(chkDate) = vbSunday Then
    HollydayName = "Sunday"
  ElseIf Weekday(chkDate) = vbSaturday Then
    HollydayName = "Saturday"
  End If
  
  If flg = True Then
    For line = 3 To setSheet.Cells(Rows.count, 12).End(xlUp).row
      If setSheet.Range("M" & line) = chkDate Then
        HollydayName = "��Ўw��x��"
        Exit For
      End If
    Next
  End If
  
End Function


'**************************************************************************************************
' * ���O��`
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function ���O��`()
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
  Set setSheet = ThisWorkbook.Worksheets("�ݒ�")
  
  For line = 3 To 20
    If setSheet.Range("A" & line) <> "" Then
      setSheet.Range("B" & line).Name = setSheet.Range("A" & line)
    End If
  Next
  endLine = setSheet.Cells(Rows.count, 11).End(xlUp).row
  setSheet.Range("K3:K" & endLine).Name = "�S����"

  endLine = setSheet.Cells(Rows.count, 15).End(xlUp).row
  setSheet.Range("O3:O" & endLine).Name = "�x�����X�g"

  Exit Function
'�G���[������=====================================================================================
catchError:

End Function
