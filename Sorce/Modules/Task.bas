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
  
  endLine = setSheet.Cells(Rows.count, 18).End(xlUp).row
  count = count + 1
  For line = 3 To endLine
    If setSheet.Range("R" & line) <> "" Then
      With taskList
        .Add item:=setSheet.Range("R" & line).Value, Key:=str(count)
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

'  On Error GoTo catchError

  Call init.setting
  Set memberList = New Collection
  count = 1
  
  endLine = Cells(Rows.count, 1).End(xlUp).row
  With memberList
    .Add item:="�H��", Key:=str(count)
  End With
  count = count + 1
  
    For line = 6 To endLine
      If mainSheet.Range(setVal("cell_AssignP") & line).Value <> "" And isCollection(memberList, mainSheet.Range(setVal("cell_AssignP") & line).Value) = False Then
        With memberList
          .Add item:=mainSheet.Range(setVal("cell_AssignP") & line).Value, Key:=str(count)
        End With
        count = count + 1
      End If
    Next
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
    
    If Range(setVal("cell_AssignP") & line).Text = filterName Or Range(setVal("cell_AssignA") & line).Text = filterName Then
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
    
    If Library.TEXTJOIN(" ", True, Range("C" & line & ":" & setVal("cell_TaskAreaEnd") & line)) Like "*" & filterName & "*" Then
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















