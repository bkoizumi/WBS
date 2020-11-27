VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UF_Calendar5R 
   Caption         =   "日付選択"
   ClientHeight    =   2475
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2580
   OleObjectBlob   =   "UF_Calendar5R.frx":0000
   StartUpPosition =   2  '画面の中央
End
Attribute VB_Name = "UF_Calendar5R"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'***************************************************************************************************
'   カレンダーフォーム5(日付入力部品)    ※ユーザーフォーム(改⑤)   UF_Calendar5R(UserForm)
'
'   作成者:井上治  URL:http://www.ne.jp/asahi/excel/inoue/ [Excelでお仕事!]
'***************************************************************************************************
'   ※祝日法改正時は｢clsAboutCalendar2R｣の｢GP_MakeHoliParameter｣を修正する
'***************************************************************************************************
'変更日付 Rev  変更履歴内容========================================================================>
'18/02/20(4.30)祝日判定関連を｢modAboutCalendar2｣に移行する対応
'18/02/20(4.30)親フォームとの受け渡しをプロパティに変更
'18/02/21(4.31)カレンダー初週の表示不具合修正(1日が表示されない件)、モードレス表示対応
'18/02/21(4.31)結果日付を返さずに本フォーム内でセル(又はTextBox)に日付を書き込むように仕様変更
'18/02/22(4.32)結合セルから値を取り出す時の不具合対応
'18/09/17(1.70)祝日パラメータシートを使わない方式としてクラス化(チェック処理廃止⇒Subプロシージャ化)
'18/11/26(1.71)カレンダーフォーム表示幅調整(Windows10対応)
'18/11/28(1.80)カレンダーフォーム上の各日付ラベルをクラス化(WithEvents)させる対応
'19/12/08(1.91)UserForm_QueryCloseを追加(Close⇒Hide)
'***************************************************************************************************
Option Explicit
'===================================================================================================
' [起算曜日] ※カレンダーを月曜開始(曜日左端)にする場合は｢2｣に変更して下さい。
Private Const g_cnsStartYobi = 1                            ' 1=日曜日,2=月曜日(他は不可)
'===================================================================================================
' [年の表示限度(From/To)]
Private Const g_cnsYearFrom = 1947                          ' 祝日法施行
Private Const g_cnsYearToAdd = 3                            ' システム日の年+n年までの指定
'===================================================================================================
' フォーム上の色指定等の定数
Private Const cnsBC_Select = &HFFCC33                       ' 選択日付の背景色
Private Const cnsBC_Other = &HE0E0E0                        ' 当月以外の背景色
Private Const cnsBC_Sunday = &HFFDDFF                       ' 日曜の背景色
Private Const cnsBC_Saturday = &HDDFFDD                     ' 土曜の背景色
Private Const cnsBC_Month = &HFFFFFF                        ' 当月土日以外の背景色
Private Const cnsFC_Hori = &HFF                             ' 祝日の文字色
Private Const cnsFC_Normal = &HC00000                       ' 祝日以外の文字色
Private Const cnsDefaultGuide = "矢印キーで操作できます。"
Private Const g_cnsDateFormat = "YYYY/MM/DD"                ' デフォルトの日付Format
'===================================================================================================
' 呼び元との受け渡し変数
Private g_FormDate1 As Date                                 ' 現在日付
Private g_intEntMode As Integer                             ' 入力モード(0=セル、1=TextBox)
Private g_objRange As Range                                 ' 対象セル
Private g_objTextBox As MSForms.TextBox                     ' 対象TextBox
Private g_strDateFormat As String                           ' 日付フォーマット(TextBox時)
'===================================================================================================
' フォーム表示中に保持するモジュール変数
Private g_tblYobi As Variant                                ' 曜日テーブル
Private g_tblDateLabel(44) As New clsUF_Cal5Label1R         ' 日付ラベルイベントクラステーブル
Private g_tblFixLabel(11) As New clsUF_Cal5Label2R          ' 固定ラベルイベントクラステーブル
Private g_intCurYear As Integer                             ' 現在表示年
Private g_intCurMonth As Integer                            ' 現在表示月
Private g_CurPos As Long                                    ' 現在日付位置
Private g_CurPosF As Long                                   ' 当月月初日位置
Private g_CurPosT As Long                                   ' 当月月末日位置
Private g_swBatch As Boolean                                ' イベント抑制SW
Private g_VisibleYear As Boolean                            ' Conboの年表示スイッチ
Private g_VisibleMonth As Boolean                           ' Comboの月表示スイッチ

'***************************************************************************************************
' ■フォーム上のイベント
'***************************************************************************************************
'* 処理名　：CBO_MONTH_Change
'* 機能　　：｢月｣コンボの操作イベント
'===================================================================================================
'* 返り値　：(なし)
'* 引数　　：(なし)
'===================================================================================================
'* 作成日　：2010年01月13日
'* 作成者　：井上　治
'* 更新日　：2010年01月13日
'* 更新者　：井上　治
'* 機能説明：
'* 注意事項：
'***************************************************************************************************
Private Sub CBO_MONTH_Change()
    '==============================================================================================-
    Dim intMonth As Integer
    If g_swBatch Then Exit Sub
    intMonth = CInt(CBO_MONTH.Text)
    g_FormDate1 = DateSerial(g_intCurYear, intMonth, 1)
    ' 年月コンボの非表示化
    Call GP_EraseYearMonth
    ' カレンダー作成
    Call GP_MakeCalendar
End Sub

'***************************************************************************************************
'* 処理名　：CBO_YEAR_Change
'* 機能　　：｢年｣コンボの操作イベント
'===================================================================================================
'* 返り値　：(なし)
'* 引数　　：(なし)
'===================================================================================================
'* 作成日　：2010年01月13日
'* 作成者　：井上　治
'* 更新日　：2010年01月13日
'* 更新者　：井上　治
'* 機能説明：
'* 注意事項：
'***************************************************************************************************
Private Sub CBO_YEAR_Change()
    '==============================================================================================-
    Dim intYear As Integer
    If g_swBatch Then Exit Sub
    intYear = CInt(CBO_YEAR.Text)
    g_FormDate1 = DateSerial(intYear, g_intCurMonth, 1)
    ' 年月コンボの非表示化
    Call GP_EraseYearMonth
    ' カレンダー作成
    Call GP_MakeCalendar
End Sub

'***************************************************************************************************
'* 処理名　：LBL_PREV_Click
'* 機能　　：「←(前月)」Clickイベント
'===================================================================================================
'* 返り値　：(なし)
'* 引数　　：(なし)
'===================================================================================================
'* 作成日　：2010年01月13日
'* 作成者　：井上　治
'* 更新日　：2010年01月13日
'* 更新者　：井上　治
'* 機能説明：
'* 注意事項：
'***************************************************************************************************
Private Sub LBL_PREV_Click()
    '==============================================================================================-
    ' 年月コンボの非表示化
    Call GP_EraseYearMonth
    ' 前月月初日を設定
    g_FormDate1 = DateSerial(g_intCurYear, g_intCurMonth - 1, 1)
    ' カレンダー作成
    Call GP_MakeCalendar
End Sub

'***************************************************************************************************
'* 処理名　：LBL_NEXT_Click
'* 機能　　：「→(翌月)」Clickイベント
'===================================================================================================
'* 返り値　：(なし)
'* 引数　　：(なし)
'===================================================================================================
'* 作成日　：2010年01月13日
'* 作成者　：井上　治
'* 更新日　：2010年01月13日
'* 更新者　：井上　治
'* 機能説明：
'* 注意事項：
'***************************************************************************************************
Private Sub LBL_NEXT_Click()
    '==============================================================================================-
    ' 年月コンボの非表示化
    Call GP_EraseYearMonth
    ' 翌月月初日を設定
    g_FormDate1 = DateSerial(g_intCurYear, g_intCurMonth + 1, 1)
    ' カレンダー作成
    Call GP_MakeCalendar
End Sub

'***************************************************************************************************
'* 処理名　：LBL_MONTH_Click
'* 機能　　：「月」Clickイベント
'===================================================================================================
'* 返り値　：(なし)
'* 引数　　：(なし)
'===================================================================================================
'* 作成日　：2010年01月13日
'* 作成者　：井上　治
'* 更新日　：2010年01月13日
'* 更新者　：井上　治
'* 機能説明：
'* 注意事項：
'***************************************************************************************************
Private Sub LBL_MONTH_Click()
    '==============================================================================================-
    Dim intMonth As Integer
    Dim IX As Long, CUR As Long
    ' 年コンボが表示されていたら消去
    Call GP_EraseYear
    ' 月コンボの表示
    g_swBatch = True
    With CBO_MONTH
        .Clear
        For intMonth = 1 To 12
            .AddItem Format(intMonth, "00")
            If intMonth = g_intCurMonth Then CUR = IX
            IX = IX + 1
        Next intMonth
        .ListIndex = CUR
        .Visible = True
        g_VisibleMonth = True
    End With
    g_swBatch = False
End Sub

'***************************************************************************************************
'* 処理名　：LBL_YEAR_Click
'* 機能　　：「年」Clickイベント
'===================================================================================================
'* 返り値　：(なし)
'* 引数　　：(なし)
'===================================================================================================
'* 作成日　：2010年01月13日
'* 作成者　：井上　治
'* 更新日　：2010年01月13日
'* 更新者　：井上　治
'* 機能説明：
'* 注意事項：
'***************************************************************************************************
Private Sub LBL_YEAR_Click()
    '==============================================================================================-
    Dim intYear As Integer, intYearSTR As Integer, intYearEND As Integer
    Dim IX As Long, CUR As Long
    ' 月コンボが表示されていたら消去
    Call GP_EraseMonth
    ' 年コンボの表示
    g_swBatch = True
    With CBO_YEAR
        .Clear
        intYearSTR = g_intCurYear - 10
        If intYearSTR < g_cnsYearFrom Then intYearSTR = g_cnsYearFrom
        intYearEND = g_intCurYear + 10
        intYear = Year(Date) + g_cnsYearToAdd
        If intYearEND > intYear Then intYearEND = intYear
        For intYear = intYearSTR To intYearEND
            .AddItem CStr(intYear)
            If intYear = g_intCurYear Then CUR = IX
            IX = IX + 1
        Next intYear
        .ListIndex = CUR
        .Visible = True
        g_VisibleYear = True
    End With
    g_swBatch = False
End Sub

'***************************************************************************************************
'* 処理名　：UserForm_Activate
'* 機能　　：フォーム表示(繰り返し表示の場合はHideのみのためInitializeは起きない)
'===================================================================================================
'* 返り値　：(なし)
'* 引数　　：(なし)
'===================================================================================================
'* 作成日　：2010年01月13日
'* 作成者　：井上　治
'* 更新日　：2018年02月22日
'* 更新者　：井上　治
'* 機能説明：
'* 注意事項：
'***************************************************************************************************
Private Sub UserForm_Activate()
    '==============================================================================================-
    ' コンボは非表示
    CBO_YEAR.Visible = False
    CBO_MONTH.Visible = False
    g_VisibleYear = False
    g_VisibleMonth = False
    g_FormDate1 = 0
    '==============================================================================================-
    ' 入力モード(0=セル、1=TextBox)
    If g_intEntMode = 1 Then
        ' 1=TextBox
        ' 元となる日付をテキストボックスから取得
        If IsDate(Trim(g_objTextBox.Text)) Then
            g_FormDate1 = CDate(Trim(g_objTextBox.Text))
        End If
    Else
        ' 0=セル
        ' 元となる日付をセルから取得
        On Error Resume Next
        If IsDate(Trim(g_objRange.Cells(1).Value)) Then
            g_FormDate1 = CDate(Trim(g_objRange.Cells(1).Value))
        End If
        On Error GoTo 0
    End If
    ' 受け取れない場合は当日をセット
    If g_FormDate1 = 0 Then g_FormDate1 = Date
    '==============================================================================================-
    ' カレンダー作成
    Call GP_MakeCalendar
    LBL_GUIDE.Caption = cnsDefaultGuide             ' ガイド表示
    ' 表示位置をマニュアルに変更
    If Me.StartUpPosition <> 0 Then Me.StartUpPosition = 0
End Sub

'***************************************************************************************************
'* 処理名　：UserForm_Deactivate
'* 機能　　：フォーム非アクティブ状態
'===================================================================================================
'* 返り値　：(なし)
'* 引数　　：(なし)
'===================================================================================================
'* 作成日　：2018年02月21日
'* 作成者　：井上　治
'* 更新日　：2018年02月21日
'* 更新者　：井上　治
'* 機能説明：
'* 注意事項：UserForm上からの利用の対応のためモードレスにできないことからこのイベントは発生しない
'***************************************************************************************************
Private Sub UserForm_Deactivate()
    '==============================================================================================-
    Me.Hide
End Sub

'***************************************************************************************************
'* 処理名　：UserForm_Initialize
'* 機能　　：フォーム初期化(繰り返し表示の場合はHideのみのためInitializeは起きない)
'===================================================================================================
'* 返り値　：(なし)
'* 引数　　：(なし)
'===================================================================================================
'* 作成日　：2010年01月13日
'* 作成者　：井上　治
'* 更新日　：2018年11月28日
'* 更新者　：井上　治
'* 機能説明：
'* 注意事項：
'***************************************************************************************************
Private Sub UserForm_Initialize()
    '==============================================================================================-
    Dim lngIx As Long                                               ' テーブルINDEX
    Dim lngIx2 As Long                                              ' テーブルINDEX
    Dim lngIxC As Long                                              ' カレンダーテーブルINDEX
    Dim dteDate As Date                                             ' 日付WORK
    Dim tblTodayNm As Variant                                       ' 昨日､今日､明日の名称
    Dim tblCalendar() As g_typAboutCalendar2                        ' カレンダーテーブル
    '==============================================================================================-
    ' 各日付ラベルイベントクラスの初期化
    For lngIx = 0 To 44
        Call g_tblDateLabel(lngIx).NewClass(Me.Controls("LBL_" & Format((lngIx + 1), "00")), lngIx)
    Next lngIx
    ' 各固定ラベルイベントクラスの初期化
    Call g_tblFixLabel(0).NewClass(LBL_SUN, cnsDefaultGuide)        ' 日
    Call g_tblFixLabel(1).NewClass(LBL_MON, cnsDefaultGuide)        ' 月
    Call g_tblFixLabel(2).NewClass(LBL_TUE, cnsDefaultGuide)        ' 火
    Call g_tblFixLabel(3).NewClass(LBL_WED, cnsDefaultGuide)        ' 水
    Call g_tblFixLabel(4).NewClass(LBL_THU, cnsDefaultGuide)        ' 木
    Call g_tblFixLabel(5).NewClass(LBL_FRI, cnsDefaultGuide)        ' 金
    Call g_tblFixLabel(6).NewClass(LBL_SAT, cnsDefaultGuide)        ' 土
    Call g_tblFixLabel(7).NewClass(LBL_PREV, "前月に戻ります(PageUp)") ' ←
    Call g_tblFixLabel(8).NewClass(LBL_NEXT, "翌月に進みます(PageDown)") ' →
    Call g_tblFixLabel(9).NewClass(LBL_YM, "年か月を選択します。")  ' 年月
    Call g_tblFixLabel(10).NewClass(LBL_YEAR, "年が選択できます。") ' 年
    Call g_tblFixLabel(11).NewClass(LBL_MONTH, "月が選択できます。") ' 月
    '==============================================================================================-
    ' 起算曜日による曜日見出しの位置修正
    If g_cnsStartYobi = 2 Then
        ' 月曜起算
        LBL_MON.Left = 3
        LBL_MON.Width = 16.75
        LBL_TUE.Left = 20.5
        LBL_WED.Left = 38
        LBL_WED.Width = 17
        LBL_THU.Left = 55.5
        LBL_THU.Width = 16.75
        LBL_FRI.Left = 73
        LBL_SAT.Left = 90.5
        LBL_SUN.Left = 108
        LBL_SUN.Width = 17
    Else
        ' 日曜起算
        LBL_SUN.Left = 3
        LBL_SUN.Width = 16.75
        LBL_MON.Left = 20.5
        LBL_TUE.Left = 38
        LBL_TUE.Width = 17
        LBL_WED.Left = 55.5
        LBL_WED.Width = 16.75
        LBL_THU.Left = 73
        LBL_FRI.Left = 90.5
        LBL_SAT.Left = 108
        LBL_SAT.Width = 17
    End If
    g_tblYobi = Array("(日)", "(月)", "(火)", "(水)", "(木)", "(金)", "(土)")
    '==============================================================================================-
    ' 昨日､今日､明日の処理
    dteDate = Date              ' 今日
    ' カレンダーテーブル作成(当月+前後の3ヶ月用)
    Call modAboutCalendar2R.GP_GetCalendarTable3(Year(dteDate), Month(dteDate), tblCalendar)
    ' 昨日に戻す
    dteDate = dteDate - 1
    lngIxC = 0
    ' カレンダー上の昨日の位置を判定
    Do While lngIxC <= UBound(tblCalendar)
        ' 日付発見は抜ける
        If tblCalendar(lngIxC).Hiduke = dteDate Then Exit Do
        ' 次の日へ
        lngIxC = lngIxC + 1
    Loop
    tblTodayNm = Array("[昨日]", "[今日]", "[明日]")
    lngIx2 = 0
    ' 昨日､今日､明日のセット
    For lngIx = 42 To 44
        dteDate = tblCalendar(lngIxC).Hiduke
        ' 日付コントロール情報テーブルのセット
        With g_tblDateLabel(lngIx)
            .Hiduke = dteDate
            .Yobi = tblCalendar(lngIxC).Yobi
            .Syuku = tblCalendar(lngIxC).Syuku
            .StsGuide = tblTodayNm(lngIx2) & Format(dteDate, g_cnsDateFormat) & g_tblYobi(.Yobi)
            ' 祝日か
            If .Syuku <> 0 Then
                .StsGuide = .StsGuide & " " & tblCalendar(lngIxC).SyukuNm
            End If
        End With
        ' カレンダーテーブル位置を翌日に移動
        lngIxC = lngIxC + 1
        lngIx2 = lngIx2 + 1
    Next lngIx
End Sub

'***************************************************************************************************
'* 処理名　：UserForm_KeyDown
'* 機能　　：フォーム上キーボード処理
'===================================================================================================
'* 返り値　：(なし)
'* 引数　　：(既定)
'===================================================================================================
'* 作成日　：2010年01月13日
'* 作成者　：井上　治
'* 更新日　：2010年01月13日
'* 更新者　：井上　治
'* 機能説明：
'* 注意事項：
'***************************************************************************************************
Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, _
                             ByVal Shift As Integer)
    '==============================================================================================-
    ' KeyCode(Shift併用)による制御
    Select Case KeyCode
        Case vbKeyReturn, vbKeyExecute, vbKeySeparator  ' Enter(決定)
            Call GP_ClickCalendar(g_CurPos)
        Case vbKeyCancel, vbKeyEscape                   ' Cancel, Esc(終了)
            Me.Hide
        Case vbKeyPageDown                              ' PageDown(次月)
            Call LBL_NEXT_Click
        Case vbKeyPageUp                                ' KeyPageUp(前月)
            Call LBL_PREV_Click
        Case vbKeyRight, vbKeyNumpad6, vbKeyAdd         ' →(翌日)
            Call GP_MoveDay(1)
        Case vbKeyLeft, vbKeyNumpad4, vbKeySubtract     ' ←(前日)
            Call GP_MoveDay(-1)
        Case vbKeyUp, vbKeyNumpad8                      ' －(7日後)
            Call GP_MoveDay(-7)
        Case vbKeyDown, vbKeyNumpad2                    ' ＋(7日前)
            Call GP_MoveDay(7)
        Case vbKeyHome                                  ' Home(月初)
            Call GP_MoveDay(g_CurPosF - g_CurPos)
        Case vbKeyEnd                                   ' End(月末)
            Call GP_MoveDay(g_CurPosT - g_CurPos)
        Case vbKeyTab                                   ' Tab(Shiftによる)
            If Shift = 1 Then
                Call GP_MoveDay(-1)                     ' 前日
            Else
                Call GP_MoveDay(1)                      ' 翌日
            End If
        Case vbKeyF11                                   ' F11(前年)
            g_FormDate1 = DateSerial(g_intCurYear - 1, g_intCurMonth, 1)
            ' 年月コンボの非表示化
            Call GP_EraseYearMonth
            ' カレンダー作成
            Call GP_MakeCalendar
        Case vbKeyF12                                   ' F11(翌年)
            g_FormDate1 = DateSerial(g_intCurYear + 1, g_intCurMonth, 1)
            ' 年月コンボの非表示化
            Call GP_EraseYearMonth
            ' カレンダー作成
            Call GP_MakeCalendar
    End Select
End Sub

'***************************************************************************************************
'* 処理名　：UserForm_MouseMove
'* 機能　　：フォーム上マウス移動
'===================================================================================================
'* 返り値　：(なし)
'* 引数　　：(既定)
'===================================================================================================
'* 作成日　：2010年01月13日
'* 作成者　：井上　治
'* 更新日　：2010年01月13日
'* 更新者　：井上　治
'* 機能説明：
'* 注意事項：
'***************************************************************************************************
Private Sub UserForm_MouseMove(ByVal Button As Integer, _
                               ByVal Shift As Integer, _
                               ByVal x As Single, ByVal y As Single)
    '==============================================================================================-
    Me.LBL_GUIDE.Caption = cnsDefaultGuide
End Sub

'***************************************************************************************************
'* 処理名　：UserForm_QueryClose
'* 機能　　：フォームイベント(QueryClose)
'===================================================================================================
'* 返り値　：(なし)
'* 引数　　：(既定)
'===================================================================================================
'* 作成日　：2019年12月08日
'* 作成者　：井上　治
'* 更新日　：2019年12月08日
'* 更新者　：井上　治
'* 機能説明：
'* 注意事項：
'***************************************************************************************************
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    '==============================================================================================-
    ' CloseをHideに置き換える
    If CloseMode = vbFormControlMenu Then
        Me.Hide
        Cancel = True
    End If
End Sub

'***************************************************************************************************
' ■共通サブ処理(Private)
'***************************************************************************************************
'* 処理名　：GP_MakeCalendar
'* 機能　　：カレンダー表示
'===================================================================================================
'* 返り値　：(なし)
'* 引数　　：(なし)
'===================================================================================================
'* 作成日　：2010年01月13日
'* 作成者　：井上　治
'* 更新日　：2018年11月28日
'* 更新者　：井上　治
'* 機能説明：
'* 注意事項：
'***************************************************************************************************
Private Sub GP_MakeCalendar()
    '==============================================================================================-
    Dim dteDate As Date                                             ' 日付WORK
    Dim intYobi As Integer                                          ' 曜日位置INDEX
    Dim intYear As Integer                                          ' 指定年
    Dim lngCurStrIx As Long                                         ' 当月開始INDEX
    Dim lngCurEndIx As Long                                         ' 当月終了INDEX
    Dim lngIx As Long                                               ' フォーム上位置INDEX
    Dim lngIx2 As Long                                              ' テーブルINDEX(Work)
    Dim lngIxC As Long                                              ' カレンダーINDEX
    Dim lngIxCMax As Long                                           ' カレンダーINDEX上限
    Dim lngCurPos As Long                                           ' 選択日の位置
    Dim tblCalendar() As g_typAboutCalendar2                        ' カレンダーテーブル
    '==============================================================================================-
    intYear = Year(g_FormDate1)                                 ' 指定年
    ' 指定年月が利用可能かチェック
    If ((intYear < g_cnsYearFrom) Or _
        (intYear > (Year(Date) + g_cnsYearToAdd))) Then
        MsgBox "祝日計算範囲を超えています。", vbExclamation, Me.Caption
        g_FormDate1 = g_tblDateLabel(g_CurPos).Hiduke
    End If
    g_intCurYear = Year(g_FormDate1)                            ' 指定年
    g_intCurMonth = Month(g_FormDate1)                          ' 指定月
    LBL_YM.Caption = g_intCurYear & "年" & Format(g_intCurMonth, "00") & "月"
    '==============================================================================================-
    ' カレンダーテーブル作成(当月+前後の3ヶ月用)
    Call modAboutCalendar2R.GP_GetCalendarTable3R(g_intCurYear, _
                                                  g_intCurMonth, _
                                                  tblCalendar, _
                                                  lngCurStrIx, _
                                                  lngCurEndIx)
    lngIxCMax = UBound(tblCalendar)
    ' カレンダーテーブル先頭位置
    lngIxC = lngCurStrIx
    ' 指定日付から一旦、前週の最終日(日曜日)に戻す
    lngIxC = lngIxC - tblCalendar(lngCurStrIx).Yobi
    ' 月曜始まり時の調整
    If g_cnsStartYobi = 2 Then
        lngIxC = lngIxC + 1
        ' 2日始まりになってしまう時は1週戻す
        If lngIxC > lngCurStrIx Then lngIxC = lngIxC - 7
    End If
    intYobi = 0
    lngCurPos = -1
    '==============================================================================================-
    ' フォーム上の日付セット(7曜×6週=42件固定⇒0始まり)
    For lngIx = 0 To 41
        ' 当位置の日付、曜日を算出
        intYobi = intYobi + 1
        If intYobi > 7 Then intYobi = 1
        dteDate = tblCalendar(lngIxC).Hiduke
        ' 現在選択日か
        If dteDate = g_FormDate1 Then
            lngCurPos = lngIx
        End If
        ' 日付コントロール情報テーブルのセット
        With g_tblDateLabel(lngIx)
            .Hiduke = dteDate
            .Yobi = tblCalendar(lngIxC).Yobi
            .Syuku = tblCalendar(lngIxC).Syuku
            .StsGuide = Format(dteDate, g_cnsDateFormat) & g_tblYobi(.Yobi)
            ' 祝日か
            If .Syuku <> 0 Then
                .StsGuide = .StsGuide & " " & tblCalendar(lngIxC).SyukuNm
            End If
        End With
        ' 月初日、月末日の位置を取得
        If lngIxC = lngCurStrIx Then
            ' 当月初日
            g_CurPosF = lngIx
        ElseIf lngIxC = lngCurEndIx Then
            ' 当月末日
            g_CurPosT = lngIx
        End If
        ' ラベルコントロールを配列化した変数
        g_tblDateLabel(lngIx).Label.Caption = day(dteDate)
        ' 文字色、背景色のセット
        Call GP_SetForeColor(lngIx, g_FormDate1)
        ' カレンダーテーブル位置を翌日に移動
        lngIxC = lngIxC + 1
    Next lngIx
    LBL_GUIDE.Caption = g_tblDateLabel(g_CurPos).StsGuide   ' ガイド表示
End Sub

'***************************************************************************************************
'* 処理名　：GP_MoveDay
'* 機能　　：カレンダー上の選択位置移動
'===================================================================================================
'* 返り値　：(なし)
'* 引数　　：Arg1 = 移動量(Long)                ※マイナスあり
'===================================================================================================
'* 作成日　：2010年01月13日
'* 作成者　：井上　治
'* 更新日　：2018年11月28日
'* 更新者　：井上　治
'* 機能説明：
'* 注意事項：
'***************************************************************************************************
Private Sub GP_MoveDay(lngMove As Long)
    '==============================================================================================-
    Dim lngPos As Long                                              ' テーブル位置INDEX
    Dim dteDate As Date                                             ' 日付Work
    ' 年月コンボの非表示化
    Call GP_EraseYearMonth
    ' 移動後の位置,日付を算出
    lngPos = g_CurPos + lngMove                                 ' 移動後位置
    dteDate = g_FormDate1 + lngMove                             ' 移動後日付
    ' コントロールテーブル外か
    If ((lngPos < 0) Or (lngPos > 41)) Then
        ' 前月又は翌月に移動
        g_FormDate1 = dteDate
        Call GP_MakeCalendar
        Exit Sub
    End If
    '==============================================================================================-
    ' 以前の位置の日付ラベルの背景色を元に戻す
    Call GP_SetForeColor(g_CurPos, dteDate)
    '==============================================================================================-
    ' 現在日付(退避)を更新
    g_FormDate1 = dteDate
    g_CurPos = lngPos
    ' 今回の位置の日付ラベルの背景色を選択状態に変更
    Call GP_SetForeColor(g_CurPos, g_FormDate1)
    LBL_GUIDE.Caption = g_tblDateLabel(g_CurPos).StsGuide   ' ガイド表示
End Sub

'***************************************************************************************************
'* 処理名　：GP_SetForeColor
'* 機能　　：文字色、背景色のセット
'===================================================================================================
'* 返り値　：(なし)
'* 引数　　：Arg1 = 現在日付位置INDEX(Long)
'* 　　　　　Arg2 = 選択日付(Date)
'===================================================================================================
'* 作成日　：2018年02月20日
'* 作成者　：井上　治
'* 更新日　：2018年11月28日
'* 更新者　：井上　治
'* 機能説明：
'* 注意事項：位置INDEXはラベルコントロール配列上の位置
'***************************************************************************************************
Private Sub GP_SetForeColor(ByVal lngPos As Long, ByVal dteCurDate As Date)
    '==============================================================================================-
    Dim dteDate As Date                                             ' 日付Work
    Dim lngYear As Long                                             ' 現在年
    Dim lngMonth As Long                                            ' 現在月
    With g_tblDateLabel(lngPos)
        dteDate = .Hiduke
        lngYear = Year(dteDate)
        lngMonth = Month(dteDate)
        ' 月度、曜日によりラベルの書式をセット
        .Label.Font.Bold = False
        .Label.ForeColor = cnsFC_Normal
        ' 現在選択日か
        If dteDate = dteCurDate Then
            ' 初期選択日付
            .Label.BackColor = cnsBC_Select
            g_CurPos = lngPos
        ElseIf ((lngYear = g_intCurYear) And (lngMonth = g_intCurMonth)) Then   ' 当月内か
            ' 当月
            Select Case .Yobi
                Case 0                  ' 日曜日
                    .Label.BackColor = cnsBC_Sunday
                Case 6                  ' 土曜日
                    .Label.BackColor = cnsBC_Saturday
                Case Else
                    .Label.BackColor = cnsBC_Month
            End Select
        Else
            ' 当月以外
            .Label.BackColor = cnsBC_Other
        End If
        ' 祝日(含振替休日)の判定
        If .Syuku <> 0 Then
            ' 文字色を赤とする
            .Label.ForeColor = cnsFC_Hori
            ' 当月は太字
            If ((lngYear = g_intCurYear) And (lngMonth = g_intCurMonth)) Then .Label.Font.Bold = True
        End If
    End With
End Sub

'***************************************************************************************************
'* 処理名　：GP_ShowGuide
'* 機能　　：ステータスガイド表示
'===================================================================================================
'* 返り値　：(なし)
'* 引数　　：Arg1 = 現在日付位置INDEX(Long)
'===================================================================================================
'* 作成日　：2018年02月20日
'* 作成者　：井上　治
'* 更新日　：2018年11月28日
'* 更新者　：井上　治
'* 機能説明：
'* 注意事項：位置INDEXはラベルコントロール配列上の位置
'***************************************************************************************************
Friend Sub GP_ShowGuide(ByVal lngPos As Long)
    '==============================================================================================-
    LBL_GUIDE.Caption = g_tblDateLabel(lngPos).StsGuide
End Sub

'***************************************************************************************************
'* 処理名　：GP_ClickCalendar
'* 機能　　：カレンダークリック
'===================================================================================================
'* 返り値　：(なし)
'* 引数　　：Arg1 = 日付コントロール情報テーブルINDEX(Long)
'===================================================================================================
'* 作成日　：2010年01月13日
'* 作成者　：井上　治
'* 更新日　：2018年11月28日
'* 更新者　：井上　治
'* 機能説明：
'* 注意事項：
'***************************************************************************************************
Friend Sub GP_ClickCalendar(ByVal lngPos As Long)
    '==============================================================================================-
    g_FormDate1 = g_tblDateLabel(lngPos).Hiduke
    ' 年月コンボの非表示化
    Call GP_EraseYearMonth
    ' 入力モード(0=セル、1=TextBox)
    If g_intEntMode = 1 Then
        ' 1=TextBox
        g_objTextBox.Text = Format(g_FormDate1, g_strDateFormat)
    Else
        ' 0=セル
        g_objRange.Value = g_FormDate1
    End If
    Me.Hide
End Sub

'***************************************************************************************************
'* 処理名　：GP_EraseYearMonth
'* 機能　　：｢年｣｢月｣コンボの非表示化
'===================================================================================================
'* 返り値　：(なし)
'* 引数　　：(なし)
'===================================================================================================
'* 作成日　：2010年01月13日
'* 作成者　：井上　治
'* 更新日　：2010年01月13日
'* 更新者　：井上　治
'* 機能説明：
'* 注意事項：
'***************************************************************************************************
Private Sub GP_EraseYearMonth()
    '==============================================================================================-
    Call GP_EraseYear
    Call GP_EraseMonth
End Sub

'***************************************************************************************************
'* 処理名　：GP_EraseYear
'* 機能　　：｢年｣コンボの非表示化
'===================================================================================================
'* 返り値　：(なし)
'* 引数　　：(なし)
'===================================================================================================
'* 作成日　：2010年01月13日
'* 作成者　：井上　治
'* 更新日　：2010年01月13日
'* 更新者　：井上　治
'* 機能説明：
'* 注意事項：
'***************************************************************************************************
Private Sub GP_EraseYear()
    '==============================================================================================-
    If g_VisibleYear Then
        CBO_YEAR.Visible = False
        g_VisibleYear = False
    End If
End Sub

'***************************************************************************************************
'* 処理名　：GP_EraseMonth
'* 機能　　：｢月｣コンボの非表示化
'===================================================================================================
'* 返り値　：(なし)
'* 引数　　：(なし)
'===================================================================================================
'* 作成日　：2010年01月13日
'* 作成者　：井上　治
'* 更新日　：2010年01月13日
'* 更新者　：井上　治
'* 機能説明：
'* 注意事項：
'***************************************************************************************************
Private Sub GP_EraseMonth()
    '==============================================================================================-
    If g_VisibleMonth Then
        CBO_MONTH.Visible = False
        g_VisibleMonth = False
    End If
End Sub

'***************************************************************************************************
' ■■■ プロパティ ■■■
'***************************************************************************************************
' タイトル
'===================================================================================================
Friend Property Let prpTitle(ByVal strTitle As String)
    Me.Caption = strTitle
End Property

'===================================================================================================
' 入力モード(1=セル、2=TextBox)
'===================================================================================================
Friend Property Let prpEntMode(ByVal intValue As Integer)
    g_intEntMode = intValue
End Property

'===================================================================================================
' 対象セル(Object)
'===================================================================================================
Friend Property Set prpRange(ByRef objValue As Range)
    Set g_objRange = objValue
End Property

'===================================================================================================
' 対象TextBox(Object)
'===================================================================================================
Friend Property Set prpTextBox(ByRef objValue As MSForms.TextBox)
    Set g_objTextBox = objValue
End Property

'===================================================================================================
' 日付フォーマット(TextBox時)
'===================================================================================================
Friend Property Let prpDateFormat(ByVal strFormat As String)
    g_strDateFormat = strFormat
End Property

'==========================================<< End of Source >>=====================================-

