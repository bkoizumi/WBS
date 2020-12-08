Attribute VB_Name = "ctl_ProgressBar"

'**************************************************************************************************
' * プログレスバー表示制御
' *
' * http://www.ne.jp/asahi/fuji/lake/excel/progress_01.html
'**************************************************************************************************
'Option Explicit

Public mypbProgCnt As Long       'Progress カウンター変数
Public mypbSCount As Long        '処理回数

Dim myJobCnt As Long             '現在進行中の回数
Dim myBarSize As Long            'プログレスバーサイズ


'**************************************************************************************************
' * プログレスバー表示開始
' *
'**************************************************************************************************
Public Sub showStart()
  Dim myMsg1 As String
  
  myMsg1 = " 待機中"

  'ダイアログへ表示
  With Frm_Progress
    .StartUpPosition = 0
    If ThisWorkbook.Worksheets("設定").Range("B3") = "develop" Then
      .top = Application.top + 50
    Else
      .top = Application.top + (ActiveWindow.Width / 4)
    End If
    .Left = Application.Left + (ActiveWindow.Height / 2)
    .Caption = myMsg1
    
    'プログレスバーの枠の部分
    With .Label1
      .BorderStyle = fmBorderStyleSingle       '枠あり
      .SpecialEffect = fmSpecialEffectSunken
      .Height = 15
      .Left = 12
      .Width = 250
      .top = 30
    End With

    'プログレスバーのバーの部分
    With .Label2
      .BackColor = RGB(90, 248, 82)
'        .BorderStyle = fmBorderStyleSingle       '枠あり
      .SpecialEffect = fmSpecialEffectRaised
      .Height = 13
      .Left = 13
      .Width = 0
      .top = 31
    End With

    '進捗状況表示の部分 ( % )
    With .Label3
      .TextAlign = fmTextAlignCenter
      .Caption = "0%"
      .BackStyle = 0
      .Height = 14
      .Left = 12
      .Width = 250
      .top = 32
      .Font.Size = 10
      .Font.Bold = False
    End With
    
    
    'メッセージ表示の部分
    With .Label4
      '.TextAlign = fmTextAlignCenter
      '.SpecialEffect = fmSpecialEffectEtched   '枠が沈む
      '.SpecialEffect = fmSpecialEffectRaised   '浮き上がる
      '.SpecialEffect = fmSpecialEffectBump
      .Caption = "待機中"
      .Height = 14
      .Left = 12
      .Width = 250
      .top = 9
      .Font.Size = 9
      .Font.Bold = False
    End With

    myBarSize = .Label3.Width
  End With

  Frm_Progress.Show vbModeless
End Sub


'**************************************************************************************************
' * プログレスバー表示更新
' *
'**************************************************************************************************
Public Sub showCount(ProgressBarTitle As String, mypbProgCnt As Long, mypbSCount As Long, myMsg1 As String)
  Dim myMsg2 As String
  
  If mypbProgCnt > 0 Then
    myJobCnt = Int(mypbProgCnt / mypbSCount * 100)
    myMsg2 = mypbProgCnt & "/" & mypbSCount & " (" & Int(myJobCnt) & "%)"
    
    With Frm_Progress
      .Caption = ProgressBarTitle
      .Label2.Width = Int(myBarSize * myJobCnt / 100)
      .Label3.Caption = myMsg2
      .Label4.Caption = " 処理中 …　" & myMsg1
    End With
  ElseIf mypbProgCnt = 0 Then
    myJobCnt = Int(mypbProgCnt / mypbSCount * 100)
    myMsg2 = ""
    
    With Frm_Progress
      .Caption = ProgressBarTitle
      .Label2.Width = Int(myBarSize * myJobCnt / 100)
      .Label3.Caption = myMsg2
      .Label4.Caption = " 準備中 …　" & myMsg1
    End With
  
  End If
  If setVal("debugMode") = "develop" And myMsg1 <> "" Then
    Call Library.showDebugForm(ProgressBarTitle, myMsg1)
  End If
  

  DoEvents
  
End Sub


'**************************************************************************************************
' * プログレスバー表示終了
' *
'**************************************************************************************************
Public Sub showEnd()
  
  'ダイアログを閉じる
  Unload Frm_Progress
  
End Sub

