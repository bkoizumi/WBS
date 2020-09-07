VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AssignorForm 
   Caption         =   "担当者"
   ClientHeight    =   5205
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8820
   OleObjectBlob   =   "AssignorForm.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "AssignorForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
#If Win64 Then
  Private Declare PtrSafe Function GetForegroundWindow Lib "user32" () As LongPtr
  Private Declare PtrSafe Function SetWindowPos Lib "user32" (ByVal hWnd As LongPtr, ByVal hWndInsertAfter As LongPtr, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
#Else
  Private Declare Function GetForegroundWindow Lib "user32" () As Long
  Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
#End If
Private Const HWND_TOPMOST As Long = -1
Private Const SWP_NOSIZE As Long = &H1&
Private Const SWP_NOMOVE As Long = &H2&

Private Sub UserForm_Activate()
    Call SetWindowPos(GetForegroundWindow, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
    Me.StartUpPosition = 1
End Sub


'**************************************************************************************************
' * 初期設定
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Private Sub UserForm_Initialize()
  Dim line As Long, endLine As Long
  Dim item As Variant, list As Variant
  Dim assignors As Collection, assignorRate As Collection
  
  Dim assignor01Default As Integer, assignor02Default As Integer, assignor03Default As Integer, assignor04Default As Integer, assignor05Default As Integer
  Dim count As Integer
  
 
  On Error Resume Next
  
  
  '担当者リストの取得
  Call init.setting
  
  '設定済みの値取得
  Set assignors = New Collection
  
  count = 1
  If Range(setVal("cell_TaskAllocation") & ActiveCell.row) <> "" Then
    For Each strAssignor In Split(Range(setVal("cell_TaskAllocation") & ActiveCell.row), ",")
      tmp = Split(strAssignor, "<>")
      assignors.Add item:=CStr(strAssignor), Key:=CStr(count)
      count = count + 1
    Next
  End If
  
  
  
  endLine = setSheet.Cells(Rows.count, Library.getColumnNo(setVal("cell_AssignorList"))).End(xlUp).row
  
  For line = 4 To endLine
    With Assignor01
      .AddItem setSheet.Range(setVal("cell_AssignorList") & line)
      If assignors("1") Like setSheet.Range(setVal("cell_AssignorList") & line) & "*" Then
        assignor01Default = line - 4
        tmp1 = Split(assignors("1"), "<>")
        taskAllocation01.Text = tmp1(1)
      End If
    End With
    With Assignor02
      .AddItem setSheet.Range(setVal("cell_AssignorList") & line)
      If assignors("2") Like setSheet.Range(setVal("cell_AssignorList") & line) & "*" Then
        assignor02Default = line - 4
        tmp2 = Split(assignors("2"), "<>")
        taskAllocation02.Text = tmp2(1)
      End If
    
    End With
    With Assignor03
      .AddItem setSheet.Range(setVal("cell_AssignorList") & line)
      
      If assignors("3") Like setSheet.Range(setVal("cell_AssignorList") & line) & "*" Then
        assignor03Default = line - 4
        tmp3 = Split(assignors("3"), "<>")
        taskAllocation03.Text = tmp3(1)
      End If
      
    End With
    With Assignor04
      .AddItem setSheet.Range(setVal("cell_AssignorList") & line)
      
      If assignors("4") Like setSheet.Range(setVal("cell_AssignorList") & line) & "*" Then
        assignor04Default = line - 4
        tmp4 = Split(assignors("4"), "<>")
        taskAllocation04.Text = tmp4(1)
      End If
      
    End With
    With Assignor05
      .AddItem setSheet.Range(setVal("cell_AssignorList") & line)

      If assignors("5") Like setSheet.Range(setVal("cell_AssignorList") & line) & "*" Then
        assignor05Default = line - 4
        tmp5 = Split(assignors("5"), "<>")
        taskAllocation05.Text = tmp5(1)
      End If
      
    End With
  
  Next
  
  If assignor01Default <> 34 Then
    Assignor01.ListIndex = assignor01Default
  End If
  If Assignor02 <> 34 Then
    Assignor02.ListIndex = assignor02Default
  End If
  If assignor03Default <> 34 Then
    Assignor03.ListIndex = assignor03Default
  End If
  If assignor04Default <> 34 Then
    Assignor04.ListIndex = assignor04Default
  End If
  If assignor05Default <> 34 Then
    Assignor05.ListIndex = assignor05Default
  End If


End Sub


'**************************************************************************************************
' * 処理実行
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Private Sub run_Click()
  
  ActiveCell.Value = Library.TEXTJOIN(",", True, Assignor01.Text, Assignor02.Text, Assignor03.Text, Assignor04.Text, Assignor05.Text)
  Range(setVal("cell_TaskAllocation") & ActiveCell.row) = Library.TEXTJOIN(",", True, _
                                                          Assignor01.Text & "<>" & taskAllocation01.Text, _
                                                          Assignor02.Text & "<>" & taskAllocation02.Text, _
                                                          Assignor03.Text & "<>" & taskAllocation03.Text, _
                                                          Assignor04.Text & "<>" & taskAllocation04.Text, _
                                                          Assignor05.Text & "<>" & taskAllocation05.Text _
                                                          )
    
  Unload Me
End Sub

'**************************************************************************************************
' * キャンセル
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Private Sub Cancel_Click()
  Unload Me
End Sub




















