Attribute VB_Name = "Task"
'**************************************************************************************************
' * �^�X�N�����o
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function �^�X�N�����o(taskList As Collection)
  Dim line As Long, endLine As Long, count As Long

'  On Error GoTo catchError

  Call init.setting
  Set taskList = New Collection
  count = 1
  
  endLine = setSheet.Cells(Rows.count, Library.getColumnNo(setVal("cell_DataExtract"))).End(xlUp).row
  count = count + 1
  For line = 3 To endLine
    If setSheet.Range(setVal("cell_DataExtract") & line) <> "" Then
      With taskList
        .Add item:=setSheet.Range(setVal("cell_DataExtract") & line).Value, Key:=str(count)
      End With
      count = count + 1
    End If
  Next
  Exit Function
  
'�G���[������=====================================================================================
catchError:

End Function


'**************************************************************************************************
' * �S���Ғ��o
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function �S���Ғ��o(memberList As Collection)
  Dim line As Long, endLine As Long, count As Long
  Dim assignor As String
  
  
'  On Error GoTo catchError

  Call init.setting
  Set memberList = New Collection
  count = 1
  
  endLine = mainSheet.Cells(Rows.count, 1).End(xlUp).row
  With memberList
    .Add item:="�H��", Key:=str(count)
  End With
  count = count + 1

  For line = 6 To endLine
    assignor = mainSheet.Range(setVal("cell_Assign") & line).Value
    If assignor <> "" Then
        If isCollection(memberList, assignor) = False Then
          With memberList
            .Add item:=assignor, Key:=str(count)
          End With
          count = count + 1
        End If
    End If
  Next





'  For line = 6 To endLine
'    If mainSheet.Range(setVal("cell_Assign") & line).Value <> "" Then
'      For Each assignName In Split(mainSheet.Range(setVal("cell_Assign") & line).Value, ",")
'        assignor = assignName
'        If assignor <> "" And isCollection(memberList, assignor) = False Then
'          With memberList
'            .Add item:=assignor, Key:=str(count)
'          End With
'          count = count + 1
'        End If
'      Next
'    End If
'  Next
  Exit Function
'�G���[������=====================================================================================
catchError:

End Function



Function isCollection(col As Collection, query) As Boolean
  Dim item
  
  For Each item In col
    If item = query Then
      isCollection = True
      Exit Function
    End If
  Next
  isCollection = False
End Function


'**************************************************************************************************
' * �S���҃t�B���^�[
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function �S���҃t�B���^�[(filterName As String)
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long

'  On Error GoTo catchError

  Unload FilterForm
  Call Library.startScript
  Call ProgressBar.showStart
  Call init.setting
  
  mainSheet.Select
  Cells.EntireRow.Hidden = False
  endLine = Cells(Rows.count, 1).End(xlUp).row
  
  For line = 6 To endLine
    Call ProgressBar.showCount("�S���҃t�B���^�[", line, endLine, "")
    
    If Range(setVal("cell_Assign") & line).Text = filterName Or Range(setVal("cell_Assign") & line).Text = filterName Then
    Else
      Rows(line & ":" & line).EntireRow.Hidden = True
    End If
  Next
  Call ProgressBar.showEnd
  Call Library.endScript
  Exit Function
'�G���[������=====================================================================================
catchError:

End Function
  
'**************************************************************************************************
' * �^�X�N���t�B���^�[
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function �^�X�N���t�B���^�[(filterName As String)
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long

'  On Error GoTo catchError

  Unload FilterForm
  Call Library.startScript
  Call ProgressBar.showStart
  Call init.setting
  
  mainSheet.Select
  Cells.EntireRow.Hidden = False
  endLine = Cells(Rows.count, 1).End(xlUp).row
  
  For line = 6 To endLine
    Call ProgressBar.showCount("�^�X�N���t�B���^�[", line, endLine, "")
    
    If Range(setVal("cell_TaskArea") & line) Like "*" & filterName & "*" Then
    Else
      Rows(line & ":" & line).EntireRow.Hidden = True
    End If
  Next
  Call ProgressBar.showEnd
  Call Library.endScript
  Exit Function
'�G���[������=====================================================================================
catchError:

End Function


'**************************************************************************************************
' * �i���R�s�[
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function �i���R�s�[()
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  On Error GoTo catchError
  
  Call Library.startScript
  Call init.setting
  mainSheet.Select
 
  endLine = Cells(Rows.count, 1).End(xlUp).row
  
  Range(setVal("cell_Progress") & 6 & ":" & setVal("cell_Progress") & endLine).Copy
  Range(setVal("cell_ProgressLast") & 6).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
  Application.CutCopyMode = False
  
  Call Library.endScript(True)
  Exit Function
'�G���[������=====================================================================================
catchError:

  Call Library.showNotice(Err.Number, Err.Description, True)
End Function


'**************************************************************************************************
' * �i�����ݒ�
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function �i�����ݒ�(progress As Long)
  Dim line As Long, endLine As Long, colLine As Long, endColLine As Long
  On Error GoTo catchError
  
  Call Library.startScript
  Call init.setting
  mainSheet.Select
 
  Range(setVal("cell_Progress") & ActiveCell.row) = progress
  
  
  Call Library.endScript(True)
  Exit Function
'�G���[������=====================================================================================
catchError:

  Call Library.showNotice(Err.Number, Err.Description, True)
End Function


'**************************************************************************************************
' * �^�X�N�̃����N�ݒ�
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function taskLink()
  Dim line As Long, oldLine As Long
  Dim selectedCells As Range
  Dim targetCell As Range
  Dim newStartDay As Date, newEndDay As Date
  Dim HollydayName As String
    
'  On Error GoTo catchError
  Call Library.startScript
  Call init.setting
  mainSheet.Select
   
  oldLine = 0
  Set selectedCells = Selection
  
  For Each targetCell In selectedCells
    If oldLine <> 0 Then
      If Range(setVal("cell_Task") & targetCell.row) = "" Then
        Range(setVal("cell_Task") & targetCell.row) = oldLine
      Else
        Range(setVal("cell_Task") & targetCell.row) = Range(setVal("cell_Task") & targetCell.row) & "," & oldLine
      End If
      
      '��s�^�X�N�̏I����+1���J�n���ɐݒ�
      newStartDay = Range(setVal("cell_PlanEnd") & oldLine) + 1
      Call init.chkHollyday(newStartDay, HollydayName)
      Do While HollydayName <> ""
        newStartDay = newStartDay + 1
        Call init.chkHollyday(newStartDay, HollydayName)
      Loop
      Range(setVal("cell_PlanStart") & targetCell.row) = newStartDay
      
      '�I�������Đݒ�
      newEndDay = Range(setVal("cell_PlanEnd") & targetCell.row) + Range(setVal("cell_WorkLoadP") & targetCell.row)
      Call init.chkHollyday(newEndDay, HollydayName)
      Do While HollydayName <> ""
        newEndDay = newEndDay + 1
        Call init.chkHollyday(newEndDay, HollydayName)
      Loop
      Range(setVal("cell_PlanEnd") & targetCell.row) = newEndDay
      
      Range(setVal("cell_Info") & targetCell.row) = "��"
    End If
    oldLine = targetCell.row
  Next

  Exit Function
'�G���[������=====================================================================================
catchError:
  Call Library.showNotice(Err.Number, Err.Description, True)
End Function


'**************************************************************************************************
' * �^�X�N�̃����N����
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function taskUnlink()
  Dim line As Long, oldLine As Long
  Dim selectedCells As Range
  Dim targetCell As Range
    
'  On Error GoTo catchError
  Call Library.startScript
  Call init.setting
  mainSheet.Select
   
  oldLine = 0
  Set selectedCells = Selection
  
  For Each targetCell In selectedCells
    Range(setVal("cell_Task") & targetCell.row) = ""
  Next


  Exit Function
'�G���[������=====================================================================================
catchError:
  Call Library.showNotice(Err.Number, Err.Description, True)
End Function


'**************************************************************************************************
' * xxxxxxxxxx
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function rTaskInsert()
  Dim taskLevelRange As Range
  On Error GoTo catchError
  
  Call Library.startScript
  Rows("4:4").Copy
  Rows(Selection.row & ":" & Selection.row).Insert Shift:=xlDown
  Range("A" & Selection.row).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
  Application.CutCopyMode = False
  Range(setVal("cell_Info") & Selection.row & ":XFD" & Selection.row).ClearContents
  Range(setVal("cell_Info") & Selection.row & ":XFD" & Selection.row).ClearComments
  
  Range("A" & Selection.row).FormulaR1C1 = "=ROW()-5"
 
  Set taskLevelRange = Range(setVal("cell_TaskArea") & Selection.row)
  Range("B" & line).FormulaR1C1 = "=getIndentLevel(" & taskLevelRange.Address(ReferenceStyle:=xlR1C1) & ")"
  Set taskLevelRange = Nothing

  Call Library.endScript(True)

  Exit Function
'�G���[������=====================================================================================
catchError:
  Call Library.showNotice(Err.Number, Err.Description, True)
End Function


'**************************************************************************************************
' * xxxxxxxxxx
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Function rTaskDell()
  Dim selectedCells As Range

'  On Error GoTo catchError
  Call Library.startScript
  Call init.setting
  mainSheet.Select


  Rows(Selection(1).row & ":" & Selection(Selection.count).row).Delete Shift:=xlUp

  Call Library.endScript(True)

  Exit Function
'�G���[������=====================================================================================
catchError:
  Call Library.showNotice(Err.Number, Err.Description, True)
End Function





