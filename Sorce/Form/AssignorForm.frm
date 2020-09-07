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
  Dim col As Collection
  
  '担当者リストの取得
  Call init.setting
  endLine = setSheet.Cells(Rows.count, Library.getColumnNo(setVal("cell_AssignorList"))).End(xlUp).row
  
  For line = 4 To endLine
    With Assignor01
      .AddItem setSheet.Range(setVal("cell_AssignorList") & line)
    End With
    With Assignor02
      .AddItem setSheet.Range(setVal("cell_AssignorList") & line)
    End With
    With Assignor03
      .AddItem setSheet.Range(setVal("cell_AssignorList") & line)
    End With
    With Assignor04
      .AddItem setSheet.Range(setVal("cell_AssignorList") & line)
    End With
    With Assignor05
      .AddItem setSheet.Range(setVal("cell_AssignorList") & line)
    End With
  
  Next
End Sub


'**************************************************************************************************
' * 処理実行
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Private Sub run_Click()
  
  ActiveCell.Value = Library.TEXTJOIN(",", True, Assignor01.Text, Assignor02.Text, Assignor03.Text, Assignor04.Text, Assignor05.Text)
  
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

