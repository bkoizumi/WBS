VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAboutCalendar2R"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************************************
'   カレンダー及び日付処理(祝日含む)関連関数クラス              clsAboutCalendar2R(Class)
'
'   ※本クラスの操作は全てカレンダー関連モジュール(modAboutCalendar2R)で行ないます
'   本モジュールは祝日パラメータテーブルを作成してその祝日情報を元にカレンダーテーブルを作成します
'   ※祝日法改正時は下の方にある｢祝日パラメータテーブルの作成(GP_MakeHoliParameter)｣の変更が必要です
'
'   作成者:井上治  URL:http://www.ne.jp/asahi/excel/inoue/ [Excelでお仕事!]
'***************************************************************************************************
'変更日付 Rev  変更履歴内容------------------------------------------------------------------------>
'18/02/07(1.00)新規作成
'18/02/08(1.10)カレンダーテーブルに振替区分(0=通常､1=振替休日を行なわない)を追加する対応
'18/02/08(1.10)｢カレンダーテーブル作成(当月+前後の3ヶ月用)｣を追加
'18/02/10(1.20)クラス化移行のため、本モジュールはユーザー定義のみとする
'18/02/18(1.30)動作改善のため再作成(クラス化は廃止)
'18/02/21(1.40)GetCalendarTable3に当月開始・終了INDEXを返す引数(Option)を追加
'18/02/21(1.40)名称変更(g_typCalendar2⇒g_typAboutCalendar2) ※他との重複対応
'18/02/28(1.50)会社休日適用時に会社休日が振替休日や国民の休日と重なる時は法制休日を優先させる対応
'18/03/07(1.51)1日が日曜日だと週数が2始まりになってしまう件を修正
'18/03/07(1.60)各関数にチェック処理をスキップするオプションを追加(事前チェックが済んでいる前提です)
'18/09/17(1.70)｢週数｣のカウント誤りを再修正
'18/09/17(1.70)祝日パラメータシートを使わない方式としてクラス化(チェック処理廃止⇒Subプロシージャ化)
'18/09/23(1.71)クラス初期化判定を行なわないプロシージャを追加(連続呼び出し用、クラス側変更無し)
'18/10/15(1.71)コード、コメント整理(バージョン値更新無し)
'18/10/23(1.72)祝日パラメータテーブル要素追加の引数INDEXを廃止する対応(内部Static変数に変更)
'18/12/08(1.73)新天皇即位の日(5/1)、即位礼正殿の儀(10/22)を2019年限定で追加(名称は暫定)
'19/12/22(1.74)春分の日・秋分の日算出の処理タイトルコメントの引数説明誤りを修正(処理変更無し)
'***************************************************************************************************
Option Explicit
'===================================================================================================
'  ＋＋＋ バージョン値、バージョン更新日は本クラスの変更を行なう時に必ず更新すること ＋＋＋
Private Const g_cnsCalendarVersion As String = "1.74"           ' 本クラスのバージョン
Private Const g_cnsCalendarVerUpdDate As Date = #12/22/2019#    ' バージョン更新日
'  －－－ バージョン値、バージョン更新日は本クラスの変更を行なう時に必ず更新すること －－－
'===================================================================================================
Private Const g_cnsFuri As String = "(振替休日)"
Private Const g_cnsKyu2 As String = "国民の休日"
Private Const g_cnsKyu3 As String = "(会社休日)"
'---------------------------------------------------------------------------------------------------
' 祝日パラメータテーブル用ユーザー定義
Private Type g_typParamater
    Getsu As Long                                   ' 月
    SyoriKbn As Long                                ' 処理区分(0=固定日､1=HappyMonday､2=会社休日､9=特殊)
    Hiduke As Long                                  ' [固定日]日(Date型ではない)
    FurikaeKbn As Long                              ' [固定日]振替区分(0=通常､1=振替休日を行なわない)
    HmSyusu As Long                                 ' [HM]週数(第n週)
    HmYobi As Long                                  ' [HM]曜日(0=日､1=月､2=火～6=土)
    SyukuNm As String                               ' 祝日名
    StrYear As Long                                 ' 開始年(西暦)
    EndYear As Long                                 ' 終了年(西暦)
End Type
'---------------------------------------------------------------------------------------------------
' 祝日パラメータ月別配置テーブル用ユーザー定義
Private Type g_typParamMMPos
    StrIx As Long                                   ' 当月開始位置INDEX(祝日無しは-1)
    EndIx As Long                                   ' 当月終了位置INDEX(祝日無しは-1)
End Type
'---------------------------------------------------------------------------------------------------
' 振替休日・国民の休日判定対象テーブル用ユーザー定義
Private Type g_typFuriDIx
    CalIx As Long                                   ' カレンダーINDEX
    Syubetsu As Long                                ' 0=振替休日、1=国民の休日
End Type
'===================================================================================================
' 祝日パラメータテーブル(全件)
Private g_tblParamater() As g_typParamater          ' 祝日パラメータテーブル
'---------------------------------------------------------------------------------------------------
' 祝日パラメータ配置テーブル(このテーブルは月をINDEXとするため1オリジン)
Private g_tblParamMMPos(1 To 12) As g_typParamMMPos ' 祝日パラメータ月別配置テーブル

'***************************************************************************************************
'   ■■■ プロパティ(Friend) ■■■
'***************************************************************************************************
'   本クラスのバージョン(String)
'---------------------------------------------------------------------------------------------------
Friend Property Get Version() As String
    Version = g_cnsCalendarVersion
End Property

'===================================================================================================
'   バージョン更新日(Date)
'---------------------------------------------------------------------------------------------------
Friend Property Get VerUpdDate() As Date
    VerUpdDate = g_cnsCalendarVerUpdDate
End Property

'***************************************************************************************************
'   ■■■ 公開プロシージャ(Friend) ■■■
'***************************************************************************************************
'* 処理名　：GetCalendarTable1
'* 機能　　：カレンダーテーブル作成(当月1ヶ月用)
'---------------------------------------------------------------------------------------------------
'* 返り値　：(なし)
'* 引数　　：Arg1 = 年(Long)
'* 　　　　　Arg2 = 月(Long)
'* 　　　　　Arg3 = カレンダーテーブル(Array:Structure)         ※Ref参照
'---------------------------------------------------------------------------------------------------
'* 作成日　：2018年02月07日
'* 作成者　：井上　治
'* 更新日　：2018年09月17日
'* 更新者　：井上　治
'* 機能説明：
'* 注意事項：
'***************************************************************************************************
Friend Sub GetCalendarTable1(ByVal lngY As Long, _
                             ByVal lngM As Long, _
                             ByRef tblCalendar() As g_typAboutCalendar2)
    '-----------------------------------------------------------------------------------------------
    ReDim tblCalendar(0)
    ' カレンダーテーブル作成(当月分)
    Call GP_GetCalendarSub(lngY, lngM, -1, tblCalendar)
End Sub

'***************************************************************************************************
'* 処理名　：GetCalendarTable3
'* 機能　　：カレンダーテーブル作成(当月+前後の3ヶ月用)
'---------------------------------------------------------------------------------------------------
'* 返り値　：(なし)
'* 引数　　：Arg1 = 年(Long)
'* 　　　　　Arg2 = 月(Long)
'* 　　　　　Arg3 = カレンダーテーブル(Array:Structure)         ※Ref参照
'* 　　　　　Arg4 = カレンダーテーブル当月開始INDEX(Long)       ※Ref参照(Option)
'* 　　　　　Arg5 = カレンダーテーブル当月終了INDEX(Long)       ※Ref参照(Option)
'---------------------------------------------------------------------------------------------------
'* 作成日　：2018年02月08日
'* 作成者　：井上　治
'* 更新日　：2018年09月17日
'* 更新者　：井上　治
'* 機能説明：
'* 注意事項：
'***************************************************************************************************
Friend Sub GetCalendarTable3(ByVal lngYear As Long, _
                             ByVal lngMonth As Long, _
                             ByRef tblCalendar() As g_typAboutCalendar2, _
                             Optional ByRef lngCurrStartIx As Long = -1, _
                             Optional ByRef lngCurrEndIx As Long = -1)
    '-----------------------------------------------------------------------------------------------
    Dim lngCalIxMax As Long                                         ' カレンダーテーブル要素上限
    Dim lngY As Long                                                ' 年WORK
    Dim lngM As Long                                                ' 月WORK
    ReDim tblCalendar(0)
    lngCalIxMax = -1
    ' 前月の年月を算出
    If lngMonth = 1 Then
        lngY = lngYear - 1
        lngM = 12
    Else
        lngY = lngYear
        lngM = lngMonth - 1
    End If
    ' カレンダーテーブル作成(前月分)
    Call GP_GetCalendarSub(lngY, lngM, lngCalIxMax, tblCalendar)
    '-----------------------------------------------------------------------------------------------
    ' 当月開始INDEX
    lngCurrStartIx = lngCalIxMax + 1
    '-----------------------------------------------------------------------------------------------
    ' カレンダーテーブル作成(当月分)
    Call GP_GetCalendarSub(lngYear, lngMonth, lngCalIxMax, tblCalendar)
    '-----------------------------------------------------------------------------------------------
    ' 当月終了INDEX
    lngCurrEndIx = lngCalIxMax
    '-----------------------------------------------------------------------------------------------
    ' 翌月の年月を算出
    If lngMonth = 12 Then
        lngY = lngYear + 1
        lngM = 1
    Else
        lngY = lngYear
        lngM = lngMonth + 1
    End If
    ' カレンダーテーブル作成(翌月分)
    Call GP_GetCalendarSub(lngY, lngM, lngCalIxMax, tblCalendar)
End Sub

'***************************************************************************************************
'* 処理名　：SumEigyoNissu
'* 機能　　：営業日数算出(土日祝日を除外)
'---------------------------------------------------------------------------------------------------
'* 返り値　：(なし)
'* 引数　　：Arg1 = 期間開始日(Date)
'* 　　　　　Arg2 = 期間終了日(Date)
'* 　　　　　Arg3 = 営業日数(Long)                              ※Ref参照
'* 　　　　　Arg4 = 歴日数(Long)                                ※Ref参照(Option)
'---------------------------------------------------------------------------------------------------
'* 作成日　：2018年02月08日
'* 作成者　：井上　治
'* 更新日　：2018年09月17日
'* 更新者　：井上　治
'* 機能説明：
'* 注意事項：開始日､終了日自体も営業日判断に適用される
'***************************************************************************************************
Friend Sub SumEigyoNissu(ByVal dteDateF As Date, _
                         ByVal dteDateT As Date, _
                         ByRef lngCntEigyo As Long, _
                         Optional ByRef lngCntReki As Long = 0)
    '-----------------------------------------------------------------------------------------------
    lngCntEigyo = 0
    lngCntReki = 0
    ' 時刻の除去
    dteDateF = DateSerial(Year(dteDateF), Month(dteDateF), day(dteDateF))
    dteDateT = DateSerial(Year(dteDateT), Month(dteDateT), day(dteDateT))
    ' 開始日>終了日(逆転)の場合は処理なし
    If dteDateF > dteDateF Then Exit Sub
    '-----------------------------------------------------------------------------------------------
    Dim dteTmp As Date                                              ' 作業日
    Dim currY As Long                                               ' 現在年
    Dim currM As Long                                               ' 現在月
    Dim prevY As Long                                               ' 直前年
    Dim prevM As Long                                               ' 直前月
    Dim swSyuku As Boolean                                          ' 祝日判定
    Dim lngIxS As Long                                              ' カレンダーテーブルINDEX
    Dim lngIxSMax As Long                                           ' カレンダーテーブルINDEX上限
    Dim tblCalendar() As g_typAboutCalendar2                        ' カレンダーテーブル
    dteTmp = dteDateF
    currY = Year(dteTmp)
    currM = Month(dteTmp)
    '===============================================================================================
    ' 終了日までループ
    Do While dteTmp <= dteDateT
        '-------------------------------------------------------------------------------------------
        ' ■月あたり前処理
        lngIxS = 0
        prevY = currY
        prevM = currM
        lngIxSMax = -1
        ReDim tblCalendar(0)
        ' カレンダーテーブル作成(当月1ヶ月用)
        Call GP_GetCalendarSub(prevY, prevM, lngIxSMax, tblCalendar)
        '-------------------------------------------------------------------------------------------
        ' ■月あたり主処理(月替わりか終了日までループ)
        Do While ((dteTmp <= dteDateT) And (currY = prevY) And (currM = prevM))
            '----------------------------------------------
            ' 1日単位処理
            swSyuku = False
            lngCntReki = lngCntReki + 1                         ' 歴日数を加算
            '----------------------------------------------
            ' カレンダーテーブルをサーチ
            Do While lngIxS <= lngIxSMax
                ' 日付発見か
                If tblCalendar(lngIxS).Hiduke = dteTmp Then
                    ' 祝日か判定
                    swSyuku = tblCalendar(lngIxS).Syuku <> 0
                    Exit Do
                End If
                ' 次へ
                lngIxS = lngIxS + 1
            Loop
            '----------------------------------------------
            ' 土日祝日でなければ営業日数を加算
            If ((Weekday(dteTmp) <> vbSunday) And _
                (Weekday(dteTmp) <> vbSaturday) And _
                (Not swSyuku)) Then
                lngCntEigyo = lngCntEigyo + 1                   ' 営業日数を加算
            End If
            '----------------------------------------------
            ' 翌日を設定
            dteTmp = dteTmp + 1
            currY = Year(dteTmp)
            currM = Month(dteTmp)
        Loop
    Loop
End Sub

'***************************************************************************************************
'* 処理名　：SumEigyoBi
'* 機能　　：営業日数経過後営業日算出(土日祝日を除外)
'---------------------------------------------------------------------------------------------------
'* 返り値　：(なし)
'* 引数　　：Arg1 = 起算日(Date)
'* 　　　　　Arg2 = 経過日数(Long)                              ※±可能
'* 　　　　　Arg3 = 営業日数経過後営業日(Date)                  ※Ref参照(Option)
'---------------------------------------------------------------------------------------------------
'* 作成日　：2018年02月08日
'* 作成者　：井上　治
'* 更新日　：2018年09月17日
'* 更新者　：井上　治
'* 機能説明：経過日数は翌日(翌営業日)を｢1｣として算出される(前営業日は｢-1｣)
'* 注意事項：経過日数がゼロの場合は起算日をそのまま返す(土日祝判断なし)
'***************************************************************************************************
Friend Sub SumEigyoBi(ByVal dteDateF As Date, _
                      ByVal lngCntKeika As Long, _
                      ByRef dteDateT As Date)
    '-----------------------------------------------------------------------------------------------
    dteDateT = dteDateF                                         ' 一旦、起算日とする
    ' 経過日数がゼロなら非処理 ⇒ 起算日をそのまま返す(正常処理)
    If lngCntKeika = 0 Then Exit Sub
    '-----------------------------------------------------------------------------------------------
    Dim lngSign As Long                                             ' ループ時の加算日数
    Dim lngCntKeikaABS As Long                                      ' 経過日数の絶対値
    Dim dteTmp As Date                                              ' 作業日
    Dim currY As Long                                               ' 現在年
    Dim currM As Long                                               ' 現在月
    Dim prevY As Long                                               ' 直前年
    Dim prevM As Long                                               ' 直前月
    Dim lngCntEigyo As Long                                         ' 営業日数
    Dim swSyuku As Boolean                                          ' 祝日判定
    Dim lngIxS As Long                                              ' カレンダーテーブルINDEX
    Dim lngIxSMax As Long                                           ' カレンダーテーブルINDEX上限
    Dim tblCalendar() As g_typAboutCalendar2                        ' カレンダーテーブル
    dteTmp = dteDateF
    ' 経過日数の絶対値及び符号判定
    If lngCntKeika > 0 Then
        ' プラスは後方に向かって判定
        lngSign = 1
        ' 翌日から判定を始める(起算日は0なので含めない)
        dteTmp = dteTmp + 1
    Else
        ' マイナスは前方に向かって判定
        lngSign = -1
        ' 前日から判定を始める(起算日は0なので含めない)
        dteTmp = dteTmp - 1
    End If
    lngCntKeikaABS = lngCntKeika * lngSign
    currY = Year(dteTmp)
    currM = Month(dteTmp)
    '===============================================================================================
    ' 経過日数が経過するまでループ
    Do While lngCntEigyo < lngCntKeikaABS
        '-------------------------------------------------------------------------------------------
        ' ■月あたり前処理
        prevY = currY
        prevM = currM
        lngIxSMax = -1
        ReDim tblCalendar(0)
        ' カレンダーテーブル作成(当月1ヶ月用)
        Call GP_GetCalendarSub(prevY, prevM, lngIxSMax, tblCalendar)
        ' 経過日数がプラスか
        If lngCntKeika > 0 Then
            lngIxS = 0
            '---------------------------------------------------------------------------------------
            ' ■月あたり主処理(月替わりか終了日までループ) ※プラス時
            Do While ((lngCntEigyo < lngCntKeikaABS) And (currY = prevY) And (currM = prevM))
                '----------------------------------------------
                ' 1日単位処理
                swSyuku = False
                '----------------------------------------------
                'カレンダーテーブルをサーチ
                Do While lngIxS <= lngIxSMax
                    ' 日付発見か
                    If tblCalendar(lngIxS).Hiduke = dteTmp Then
                        ' 祝日か判定
                        swSyuku = tblCalendar(lngIxS).Syuku <> 0
                        Exit Do
                    End If
                    ' 次へ
                    lngIxS = lngIxS + 1
                Loop
                '----------------------------------------------
                ' 土日祝日でなければ営業日数を加算
                If ((Weekday(dteTmp) <> vbSunday) And _
                    (Weekday(dteTmp) <> vbSaturday) And _
                    (Not swSyuku)) Then
                    lngCntEigyo = lngCntEigyo + 1               ' 営業日数を加算
                    dteDateT = dteTmp                           ' 現在日をセット
                End If
                '----------------------------------------------
                ' 翌日を設定
                dteTmp = dteTmp + 1
                currY = Year(dteTmp)
                currM = Month(dteTmp)
            Loop
        Else
            lngIxS = lngIxSMax
            '---------------------------------------------------------------------------------------
            ' ■月あたり主処理(月替わりか終了日までループ) ※マイナス時
            Do While ((lngCntEigyo < lngCntKeikaABS) And (currY = prevY) And (currM = prevM))
                '----------------------------------------------
                ' 1日単位処理
                swSyuku = False
                '----------------------------------------------
                'カレンダーテーブルをサーチ(後ろから)
                Do While lngIxS >= 0
                    ' 日付発見か
                    If tblCalendar(lngIxS).Hiduke = dteTmp Then
                        ' 祝日か判定
                        swSyuku = tblCalendar(lngIxS).Syuku <> 0
                        Exit Do
                    End If
                    ' 次へ
                    lngIxS = lngIxS - 1
                Loop
                '----------------------------------------------
                ' 土日祝日でなければ営業日数を加算
                If ((Weekday(dteTmp) <> vbSunday) And _
                    (Weekday(dteTmp) <> vbSaturday) And _
                    (Not swSyuku)) Then
                    lngCntEigyo = lngCntEigyo + 1               ' 営業日数を加算
                    dteDateT = dteTmp                           ' 現在日をセット
                End If
                '----------------------------------------------
                ' 前日を設定
                dteTmp = dteTmp - 1
                currY = Year(dteTmp)
                currM = Month(dteTmp)
            Loop
        End If
    Loop
End Sub

'***************************************************************************************************
'   ■■■ サブ処理(Private) ■■■
'***************************************************************************************************
'* 処理名　：Class_Initialize
'* 機能　　：クラス初期化(既定プロシージャ)
'---------------------------------------------------------------------------------------------------
'* 返り値　：(なし)
'* 引数　　：(なし)
'---------------------------------------------------------------------------------------------------
'* 作成日　：2018年09月17日
'* 作成者　：井上　治
'* 更新日　：2018年09月17日
'* 更新者　：井上　治
'* 機能説明：祝日パラメータテーブルの作成を行なう
'* 注意事項：
'***************************************************************************************************
Private Sub Class_Initialize()
    '-----------------------------------------------------------------------------------------------
    ' 祝日パラメータテーブルの作成
    Call GP_MakeHoliParameter
End Sub

'***************************************************************************************************
'* 処理名　：GP_GetCalendarSub
'* 機能　　：カレンダーテーブル作成(単月分)
'---------------------------------------------------------------------------------------------------
'* 返り値　：(なし)
'* 引数　　：Arg1 = 年(Long)
'* 　　　　　Arg2 = 月(Long)
'* 　　　　　Arg3 = カレンダーテーブル要素上限(Long)            ※Ref参照
'* 　　　　　Arg4 = カレンダーテーブル(Array:Structure)         ※Ref参照
'---------------------------------------------------------------------------------------------------
'* 作成日　：2018年02月07日
'* 作成者　：井上　治
'* 更新日　：2018年09月17日
'* 更新者　：井上　治
'* 機能説明：
'* 注意事項：本処理ではカレンダーテーブルの初期化は行なわずに要素が追加される
'***************************************************************************************************
Private Sub GP_GetCalendarSub(ByVal lngY As Long, _
                              ByVal lngM As Long, _
                              ByRef lngCalIxMax As Long, _
                              ByRef tblCalendar() As g_typAboutCalendar2)
    '-----------------------------------------------------------------------------------------------
    Dim lngCalStrIx As Long                                         ' 当月開始位置INDEX
    '-----------------------------------------------------------------------------------------------
    lngCalStrIx = lngCalIxMax + 1
    ' カレンダーテーブル作成(当月分：追加モード)
    Call GP_MakeCalendarTable(lngY, lngM, lngCalIxMax, tblCalendar)
    '-----------------------------------------------------------------------------------------------
    ' 当月に祝日がない時は終了
    If g_tblParamMMPos(lngM).EndIx < 0 Then Exit Sub
    '-----------------------------------------------------------------------------------------------
    Dim lngFuriDIxMax As Long                                       ' 振替休日判定日INDEXテーブル上限
    Dim tblFuriDIx() As g_typFuriDIx                                ' 振替休日判定日INDEXテーブル
    lngFuriDIxMax = -1
    ReDim tblFuriDIx(0)
    ' カレンダーテーブルに祝日情報を更新(この時点では振替休日の処置はしない)
    Call GP_UpdateHolidayToCalendar(lngY, _
                                    lngM, _
                                    lngCalStrIx, _
                                    lngFuriDIxMax, _
                                    tblFuriDIx, _
                                    tblCalendar)
    '-----------------------------------------------------------------------------------------------
    ' 振替休日判定が必要か
    If lngFuriDIxMax >= 0 Then
        ' 振替休日・国民の休日設定
        Call GP_UpdateFuriKokuminKyu(tblFuriDIx, tblCalendar)
    End If
End Sub

'***************************************************************************************************
'* 処理名　：GP_MakeCalendarTable
'* 機能　　：カレンダーテーブル作成(当月分：追加モード)
'---------------------------------------------------------------------------------------------------
'* 返り値　：(なし)
'* 引数　　：Arg1 = 年(Long)
'* 　　　　　Arg2 = 月(Long)
'* 　　　　　Arg3 = カレンダーテーブル要素上限(Long)            ※Ref参照
'* 　　　　　Arg4 = カレンダーテーブル(Array:Structure)         ※Ref参照
'---------------------------------------------------------------------------------------------------
'* 作成日　：2018年02月07日
'* 作成者　：井上　治
'* 更新日　：2018年03月07日
'* 更新者　：井上　治
'* 機能説明：カレンダーテーブルの日付要素の追加を行なう
'* 注意事項：本処理ではカレンダーテーブルの初期化は行なわずに要素が追加される
'***************************************************************************************************
Private Sub GP_MakeCalendarTable(ByVal lngY As Long, _
                                 ByVal lngM As Long, _
                                 ByRef lngCalIxMax As Long, _
                                 ByRef tblCalendar() As g_typAboutCalendar2)
    '-----------------------------------------------------------------------------------------------
    Dim dteStrDate As Date                                          ' 月初日
    Dim dteEndDate As Date                                          ' 月末日
    Dim dteDate As Date                                             ' 作業日付
    Dim lngSyusu As Long                                            ' 週数(1始まり､日曜日発見で1加算)
    Dim lngHmSyusu As Long                                          ' HM週数(0始まり､月曜日発見で1加算)
    Dim lngCalIxMax2 As Long                                        '
    dteStrDate = DateSerial(lngY, lngM, 1)
    dteEndDate = DateSerial(lngY, lngM + 1, 0)
    ' カレンダーテーブルの要素を追加(当月の日数分)
    lngCalIxMax2 = lngCalIxMax + day(dteEndDate)
    ReDim Preserve tblCalendar(lngCalIxMax2)
    dteDate = dteStrDate
    ' 1日が日曜日の時に週数が2にならないように対応
    If Weekday(dteStrDate, vbSunday) = vbSunday Then
        lngSyusu = 0
    Else
        lngSyusu = 1
    End If
    ' 月初日から月末日まで繰り返す
    Do While dteDate <= dteEndDate
        lngCalIxMax = lngCalIxMax + 1
        With tblCalendar(lngCalIxMax)
            .Hiduke = dteDate
            .Yobi = Weekday(dteDate, vbSunday) - 1      ' ←日曜日=0とする値(VBAのWeekdayは1になる)
            ' 日曜日なら週数を加算
            If .Yobi = 0 Then lngSyusu = lngSyusu + 1
            ' 月曜日ならHM週数を加算
            If .Yobi = 1 Then lngHmSyusu = lngHmSyusu + 1
            .Syusu = lngSyusu
            .HmSyusu = lngHmSyusu
        End With
        ' 次の日へ
        dteDate = dteDate + 1
    Loop
End Sub

'***************************************************************************************************
'* 処理名　：GP_UpdateHolidayToCalendar
'* 機能　　：カレンダーテーブルに祝日情報を更新
'---------------------------------------------------------------------------------------------------
'* 返り値　：(なし)
'* 引数　　：Arg1 = 年(Long)
'* 　　　　　Arg2 = 月(Long)
'* 　　　　　Arg3 = カレンダーテーブル当月開始位置INDEX(Long)
'* 　　　　　Arg4 = 振替休日判定日INDEXテーブル上限INDEX(Long)      ※Ref参照
'* 　　　　　Arg5 = 振替休日判定日INDEXテーブル(Array:Structure)    ※Ref参照
'* 　　　　　Arg6 = カレンダーテーブル(Array:Structure)             ※Ref参照
'---------------------------------------------------------------------------------------------------
'* 作成日　：2018年02月07日
'* 作成者　：井上　治
'* 更新日　：2018年02月21日
'* 更新者　：井上　治
'* 機能説明：
'* 注意事項：この時点では振替休日の処置はしない
'***************************************************************************************************
Private Sub GP_UpdateHolidayToCalendar(ByVal lngY As Long, _
                                       ByVal lngM As Long, _
                                       ByVal lngCalStrIx As Long, _
                                       ByRef lngFuriDIxMax As Long, _
                                       ByRef tblFuriDIx() As g_typFuriDIx, _
                                       ByRef tblCalendar() As g_typAboutCalendar2)
    '-----------------------------------------------------------------------------------------------
    Dim lngIxS_Str As Long                                          ' 祝日テーブル当月開始位置
    Dim lngIxS_End As Long                                          ' 祝日テーブル当月終了位置
    Dim lngIxSyuku As Long                                          ' 祝日テーブルテーブルINDEX
    ' 祝日パラメータテーブルの当月範囲を得る
    lngIxS_Str = g_tblParamMMPos(lngM).StrIx
    lngIxS_End = g_tblParamMMPos(lngM).EndIx
    ' カレンダーテーブルに祝日情報を更新(この時点では振替休日の判断はしない)
    For lngIxSyuku = lngIxS_Str To lngIxS_End
        With g_tblParamater(lngIxSyuku)
            ' 祝日パラメータの有効年範囲を判定
            If ((.StrYear <= lngY) And (.EndYear >= lngY)) Then
                ' 処理区分の判定
                Select Case .SyoriKbn
                    Case 0                          ' 固定日
                        ' 祝日情報のカレンダーテーブル更新(サブ処理：日付で判定)
                        Call GP_UpdateHolidaySUB1(.Hiduke, _
                                                  .SyukuNm, _
                                                  1, _
                                                  .FurikaeKbn, _
                                                  lngCalStrIx, _
                                                  lngIxSyuku, _
                                                  lngFuriDIxMax, _
                                                  tblFuriDIx, _
                                                  tblCalendar)
                    Case 1                          ' HappyMonday
                        ' 祝日情報のカレンダーテーブル更新(サブ処理：週数＋曜日で判定)
                        Call GP_UpdateHolidaySUB2(.HmSyusu, _
                                                  .HmYobi, _
                                                  .SyukuNm, _
                                                  1, _
                                                  .FurikaeKbn, _
                                                  lngCalStrIx, _
                                                  lngIxSyuku, _
                                                  lngFuriDIxMax, _
                                                  tblFuriDIx, _
                                                  tblCalendar)
                    Case 2                          ' 固定日(会社休日)
                        ' 祝日情報のカレンダーテーブル更新(サブ処理：日付で判定)
                        Call GP_UpdateHolidaySUB1(.Hiduke, _
                                                  .SyukuNm, _
                                                  3, _
                                                  .FurikaeKbn, _
                                                  lngCalStrIx, _
                                                  lngIxSyuku, _
                                                  lngFuriDIxMax, _
                                                  tblFuriDIx, _
                                                  tblCalendar)
                    Case 9                          ' 特殊処理
                        ' 3月か
                        If lngM = 3 Then
                            ' 春分の日の算出(簡易計算方式)
                            Call GP_GetSyunbun(lngY, _
                                               lngCalStrIx, _
                                               lngIxSyuku, _
                                               lngFuriDIxMax, _
                                               tblFuriDIx, _
                                               tblCalendar)
                        ElseIf lngM = 9 Then
                            ' 秋分の日の算出(簡易計算方式)
                            Call GP_GetSyuubun(lngY, _
                                               lngCalStrIx, _
                                               lngIxSyuku, _
                                               lngFuriDIxMax, _
                                               tblFuriDIx, _
                                               tblCalendar)
                        Else ' 他月は無視
                        End If
                End Select
            End If
        End With
    Next lngIxSyuku
End Sub

'***************************************************************************************************
'* 処理名　：GP_UpdateHolidaySUB1
'* 機能　　：祝日情報のカレンダーテーブル更新(サブ処理：日付で判定)
'---------------------------------------------------------------------------------------------------
'* 返り値　：(なし)
'* 引数　　：Arg1 = 日(Long:1～31)
'* 　　　　　Arg2 = 祝日名称(String)
'* 　　　　　Arg3 = 祝日判定(Long:0=通常､1=祝日､2=振替休日､3=会社休日)
'* 　　　　　Arg4 = 振替区分(Long:0=通常､1=振替休日を行なわない)
'* 　　　　　Arg5 = カレンダー当月開始位置INDEX(Long)
'* 　　　　　Arg6 = 祝日パラメータテーブルINDEX(Long)
'* 　　　　　Arg7 = 振替休日判定日INDEXテーブルINDEX(Long)          ※Ref参照
'* 　　　　　Arg8 = 振替休日判定日INDEXテーブル(Array:Structure)    ※Ref参照
'* 　　　　　Arg9 = カレンダーテーブル(Array:Structure)             ※Ref参照
'---------------------------------------------------------------------------------------------------
'* 作成日　：2018年02月07日
'* 作成者　：井上　治
'* 更新日　：2018年02月28日
'* 更新者　：井上　治
'* 機能説明：
'* 注意事項：
'***************************************************************************************************
Private Sub GP_UpdateHolidaySUB1(ByVal lngDay As Long, _
                                 ByVal strName As String, _
                                 ByVal lngSyuku As Long, _
                                 ByVal lngFurikaeKbn As Long, _
                                 ByVal lngCalStrIx As Long, _
                                 ByVal lngIxSyuku As Long, _
                                 ByRef lngFuriDIxMax As Long, _
                                 ByRef tblFuriDIx() As g_typFuriDIx, _
                                 ByRef tblCalendar() As g_typAboutCalendar2)
    '-----------------------------------------------------------------------------------------------
    Dim lngIx As Long                                           ' テーブルINDEX
    Dim lngIxMax As Long                                        ' テーブルINDEX上限
    Dim lngIxA1 As Long                                         ' テーブルINDEX(1日後)
    Dim lngIxA2 As Long                                         ' テーブルINDEX(2日後)
    Dim lngIxB1 As Long                                         ' テーブルINDEX(1日前)
    Dim lngIxB2 As Long                                         ' テーブルINDEX(2日前)
    Dim lngIxSD As Long                                         ' 振替休日判定日INDEXテーブルINDEX
    Dim lngIxSyukuN As Long                                     ' 次の祝日パラメータテーブルINDEX
    Dim lngYearN As Long                                        ' 次の年
    Dim lngMonthN As Long                                       ' 次の月
    ' ｢日｣から該当するカレンダーテーブル上のINDEXを算出
    lngIx = lngCalStrIx + lngDay - 1
    lngIxMax = UBound(tblCalendar)
    '-------------------------------------------------------
    ' 当該日が既に祝日(又は振替休日、国民の休日)だったら｢会社休日｣は適用しない
    ' ※この判定には｢会社休日｣の対する振替区分の判断は行なっていません
    If ((lngSyuku = 3) And (tblCalendar(lngIx).Syuku <> 0)) Then
        Exit Sub
    End If
    '-------------------------------------------------------
    ' 前後の日付位置INDEXを算出
    lngIxA1 = lngIx + 1
    lngIxA2 = lngIx + 2
    lngIxB1 = lngIx - 1
    lngIxB2 = lngIx - 2
    ' 祝日情報を反映
    With tblCalendar(lngIx)
        .Syuku = lngSyuku
        .SyukuNm = strName
        .FurikaeKbn = lngFurikaeKbn
    End With
    ' 振替休日を行なわない時は終了
    If lngFurikaeKbn = 1 Then Exit Sub
    '-----------------------------------------------------------------------------------------------
    ' 日曜日又は前後判定で｢国民の休日｣対象なら振替休日判定日INDEXテーブルに追加
    If tblCalendar(lngIx).Yobi = 0 Then
        lngFuriDIxMax = lngFuriDIxMax + 1
        ReDim Preserve tblFuriDIx(lngFuriDIxMax)
        tblFuriDIx(lngFuriDIxMax).CalIx = lngIx
        tblFuriDIx(lngFuriDIxMax).Syubetsu = 0
        Exit Sub
    End If
    ' 2日前が祝日で1日前が祝日でない⇒｢国民の休日｣対象
    If lngIxB2 >= 0 Then
        If ((tblCalendar(lngIxB2).Syuku <> 0) And _
            (tblCalendar(lngIxB1).Syuku = 0) And _
            (tblCalendar(lngIxB2).FurikaeKbn = 0)) Then
            lngFuriDIxMax = lngFuriDIxMax + 1
            ReDim Preserve tblFuriDIx(lngFuriDIxMax)
            tblFuriDIx(lngFuriDIxMax).CalIx = lngIxB2
            tblFuriDIx(lngFuriDIxMax).Syubetsu = 1
            Exit Sub
        End If
    End If
    ' 2日後が祝日で1日後が祝日でない⇒｢国民の休日｣対象
    If lngIxA2 <= lngIxMax Then
        If ((tblCalendar(lngIxA2).Syuku <> 0) And _
            (tblCalendar(lngIxA1).Syuku = 0) And _
            (tblCalendar(lngIxA2).FurikaeKbn = 0)) Then
            lngFuriDIxMax = lngFuriDIxMax + 1
            ReDim Preserve tblFuriDIx(lngFuriDIxMax)
            tblFuriDIx(lngFuriDIxMax).CalIx = lngIx
            tblFuriDIx(lngFuriDIxMax).Syubetsu = 1
            Exit Sub
        End If
    End If
    ' 月末日の1日前⇒翌月1日が祝日か判定
    If lngIx = lngIxMax - 1 Then
        ' 次の祝日パラメータテーブルINDEX
        lngIxSyukuN = lngIxSyuku + 1
        ' 次の祝日が1日か
        If ((g_tblParamater(lngIxSyukuN).Hiduke = 1) And _
            (g_tblParamater(lngIxSyukuN).FurikaeKbn = 0)) Then
            lngYearN = Year(tblCalendar(lngIx).Hiduke)
            lngMonthN = Month(tblCalendar(lngIx).Hiduke) + 1     ' 次の月
            ' 次の月が有効か
            If lngMonthN > 12 Then
                lngYearN = lngYearN + 1
                lngMonthN = 1
            End If
            With g_tblParamater(lngIxSyukuN)
                ' 有効な祝日か
                If ((.Getsu = lngMonthN) And _
                    (.StrYear <= lngYearN) And _
                    (.EndYear >= lngYearN)) Then
                    ' 月末日が｢国民の休日｣対象
                    lngFuriDIxMax = lngFuriDIxMax + 1
                    ReDim Preserve tblFuriDIx(lngFuriDIxMax)
                    tblFuriDIx(lngFuriDIxMax).CalIx = lngIx
                    tblFuriDIx(lngFuriDIxMax).Syubetsu = 1
                    Exit Sub
                End If
            End With
        End If
    End If
End Sub

'***************************************************************************************************
'* 処理名　：GP_UpdateHolidaySUB2
'* 機能　　：祝日情報のカレンダーテーブル更新(サブ処理：週数＋曜日で判定)
'---------------------------------------------------------------------------------------------------
'* 返り値　：(なし)
'* 引数　　：Arg1 = HM週数(Long)
'* 　　　　　Arg2 = 曜日(Long)
'* 　　　　　Arg3 = 祝日名称(String)
'* 　　　　　Arg4 = 祝日判定(0=通常､1=祝日､2=振替休日､3=会社休日)
'* 　　　　　Arg5 = 振替区分(0=通常､1=振替休日を行なわない)
'* 　　　　　Arg6 = カレンダー当月開始位置INDEX(Long)
'* 　　　　　Arg7 = 祝日パラメータテーブルINDEX(Long)
'* 　　　　　Arg8 = 振替休日判定日INDEXテーブルINDEX(Long)          ※Ref参照
'* 　　　　　Arg9 = 振替休日判定日INDEXテーブル(Array:Structure)    ※Ref参照
'* 　　　　　Arg10= カレンダーテーブル(Array:Structure)             ※Ref参照
'---------------------------------------------------------------------------------------------------
'* 作成日　：2018年02月07日
'* 作成者　：井上　治
'* 更新日　：2018年02月21日
'* 更新者　：井上　治
'* 機能説明：
'* 注意事項：
'***************************************************************************************************
Private Sub GP_UpdateHolidaySUB2(ByVal lngSyusu As Long, _
                                 ByVal lngYobi As Long, _
                                 ByVal strName As String, _
                                 ByVal lngHoli As Long, _
                                 ByVal lngFurikaeKbn As Long, _
                                 ByVal lngCalStrIx As Long, _
                                 ByVal lngIxSyuku As Long, _
                                 ByRef lngFuriDIxMax As Long, _
                                 ByRef tblFuriDIx() As g_typFuriDIx, _
                                 ByRef tblCalendar() As g_typAboutCalendar2)
    '-----------------------------------------------------------------------------------------------
    Dim lngIx As Long                                           ' テーブルINDEX
    lngIx = lngCalStrIx
    ' カレンダーテーブルから該当週数＋曜日を探す
    Do While lngIx <= UBound(tblCalendar)
        ' HM週数＋曜日発見か
        If ((tblCalendar(lngIx).HmSyusu = lngSyusu) And (tblCalendar(lngIx).Yobi = lngYobi)) Then
            ' 祝日情報のカレンダーテーブル更新(サブ処理：日付で判定)
            Call GP_UpdateHolidaySUB1(day(tblCalendar(lngIx).Hiduke), _
                                      strName, _
                                      lngHoli, _
                                      lngFurikaeKbn, _
                                      lngCalStrIx, _
                                      lngIxSyuku, _
                                      lngFuriDIxMax, _
                                      tblFuriDIx, _
                                      tblCalendar)
            Exit Do
        End If
        ' 次へ
        lngIx = lngIx + 1
    Loop
End Sub

'***************************************************************************************************
'* 処理名　：GP_GetSyunbun
'* 機能　　：春分の日の算出(簡易計算方式)
'---------------------------------------------------------------------------------------------------
'* 返り値　：(なし)
'* 引数　　：Arg1 = 年(Long)
'* 　　　　　Arg2 = カレンダー当月開始位置INDEX(Long)
'* 　　　　　Arg3 = 祝日パラメータテーブルINDEX(Long)
'* 　　　　　Arg4 = 振替休日判定日INDEXテーブルINDEX上限(Long)      ※Ref参照
'* 　　　　　Arg5 = 振替休日判定日INDEXテーブル(Array:Structure)    ※Ref参照
'* 　　　　　Arg6 = カレンダーテーブル(Array:Structure)             ※Ref参照
'---------------------------------------------------------------------------------------------------
'* 作成日　：2018年02月07日
'* 作成者　：井上　治
'* 更新日　：2019年12月22日
'* 更新者　：井上　治
'* 機能説明：本計算方法は1948～2150年まで算出可能
'* 注意事項：本処理内では振替休日判定は行なわない
'***************************************************************************************************
Private Sub GP_GetSyunbun(ByVal lngY As Long, _
                          ByVal lngCalStrIx As Long, _
                          ByVal lngIxSyuku As Long, _
                          ByRef lngFuriDIxMax As Long, _
                          ByRef tblFuriDIx() As g_typFuriDIx, _
                          ByRef tblCalendar() As g_typAboutCalendar2)
    '-----------------------------------------------------------------------------------------------
    Dim lngD As Long                                                ' 算出日
    Dim lngY2 As Long                                               ' 年算出WORK
    ' 祝日法施行(1947年)以前,2151年以降(簡易計算不可)は無視
    lngY2 = lngY - 1980
    Select Case lngY
        Case Is <= 1979
            lngD = Int(20.8357 + (0.242194 * lngY2) - Int(lngY2 / 4))
        Case Is <= 2099
            lngD = Int(20.8431 + (0.242194 * lngY2) - Int(lngY2 / 4))
        Case Else
            lngD = Int(21.851 + (0.242194 * lngY2) - Int(lngY2 / 4))
    End Select
    ' 祝日情報のカレンダーテーブル更新(サブ処理：日付で判定)
    Call GP_UpdateHolidaySUB1(lngD, _
                              "春分の日", _
                              1, _
                              0, _
                              lngCalStrIx, _
                              lngIxSyuku, _
                              lngFuriDIxMax, _
                              tblFuriDIx, _
                              tblCalendar)
End Sub

'***************************************************************************************************
'* 処理名　：GP_GetSyuubun
'* 機能　　：秋分の日の算出(簡易計算方式)
'---------------------------------------------------------------------------------------------------
'* 返り値　：(なし)
'* 引数　　：Arg1 = 年(Long)
'* 　　　　　Arg2 = カレンダー当月開始位置INDEX(Long)
'* 　　　　　Arg3 = 祝日パラメータテーブルINDEX(Long)
'* 　　　　　Arg4 = 振替休日判定日INDEXテーブルINDEX上限(Long)      ※Ref参照
'* 　　　　　Arg5 = 振替休日判定日INDEXテーブル(Array:Structure)    ※Ref参照
'* 　　　　　Arg6 = カレンダーテーブル(Array:Structure)             ※Ref参照
'---------------------------------------------------------------------------------------------------
'* 作成日　：2018年02月07日
'* 作成者　：井上　治
'* 更新日　：2019年12月22日
'* 更新者　：井上　治
'* 機能説明：本計算方法は1948～2150年まで算出可能
'* 注意事項：本処理内では振替休日判定は行なわない
'***************************************************************************************************
Private Sub GP_GetSyuubun(ByVal lngY As Long, _
                          ByVal lngCalStrIx As Long, _
                          ByVal lngIxSyuku As Long, _
                          ByRef lngFuriDIxMax As Long, _
                          ByRef tblFuriDIx() As g_typFuriDIx, _
                          ByRef tblCalendar() As g_typAboutCalendar2)
    '-----------------------------------------------------------------------------------------------
    Dim lngD As Long                                                ' 算出日
    Dim lngY2 As Long                                               ' 年算出WORK
    ' 祝日法施行(1947年)以前,2151年以降(簡易計算不可)は無視
    lngY2 = lngY - 1980
    Select Case lngY
        Case Is <= 1979
            lngD = Int(23.2588 + (0.242194 * lngY2) - Int(lngY2 / 4))
        Case Is <= 2099
            lngD = Int(23.2488 + (0.242194 * lngY2) - Int(lngY2 / 4))
        Case Else
            lngD = Int(24.2488 + (0.242194 * lngY2) - Int(lngY2 / 4))
    End Select
    ' 祝日情報のカレンダーテーブル更新(サブ処理：日付で判定)
    Call GP_UpdateHolidaySUB1(lngD, _
                              "秋分の日", _
                              1, _
                              0, _
                              lngCalStrIx, _
                              lngIxSyuku, _
                              lngFuriDIxMax, _
                              tblFuriDIx, _
                              tblCalendar)
End Sub

'***************************************************************************************************
'* 処理名　：GP_UpdateFuriKokuminKyu
'* 機能　　：振替休日・国民の休日設定
'---------------------------------------------------------------------------------------------------
'* 返り値　：(なし)
'* 引数　　：Arg1 = 振替休日判定日INDEXテーブル(Array:Structure)
'* 　　　　　Arg2 = カレンダーテーブル(Array:Structure)         ※Ref参照
'---------------------------------------------------------------------------------------------------
'* 作成日　：2018年02月07日
'* 作成者　：井上　治
'* 更新日　：2018年02月28日
'* 更新者　：井上　治
'* 機能説明：
'* 注意事項：
'***************************************************************************************************
Private Sub GP_UpdateFuriKokuminKyu(tblFuriDIx() As g_typFuriDIx, _
                                    ByRef tblCalendar() As g_typAboutCalendar2)
    '-----------------------------------------------------------------------------------------------
    Dim lngIxF As Long                                              ' 振替休日判定日テーブルINDEX
    Dim lngIxC As Long                                              ' カレンダーテーブルINDEX
    Dim lngIxFMax As Long                                           ' 振替休日判定日テーブルINDEX上限
    Dim lngIxCMax As Long                                           ' カレンダーテーブルINDEX上限
    lngIxFMax = UBound(tblFuriDIx)
    lngIxCMax = UBound(tblCalendar)
    '-----------------------------------------------------------------------------------------------
    Do While lngIxF <= lngIxFMax
        ' 翌日のカレンダー位置を算出
        lngIxC = tblFuriDIx(lngIxF).CalIx + 1
        ' 振替休日か
        If tblFuriDIx(lngIxF).Syubetsu = 0 Then
            ' 翌日以降の最初の平日(又は会社休日)を｢振替休日｣とする
            Do While lngIxC <= lngIxCMax
                ' 祝日以外(又は会社休日)を発見
                If ((tblCalendar(lngIxC).Syuku = 0) Or (tblCalendar(lngIxC).Syuku = 3)) Then
                    ' 振替休日とする(会社休日は振替休日に置き換える)
                    tblCalendar(lngIxC).Syuku = 2
                    tblCalendar(lngIxC).SyukuNm = g_cnsFuri         ' 振替休日
                    Exit Do
                ElseIf tblCalendar(lngIxC).SyukuNm = g_cnsKyu2 Then
                    ' 既に国民の休日がある場合は処置中止(国民の休日のままにする)
                    Exit Do
                End If
                ' 次の日へ
                lngIxC = lngIxC + 1
            Loop
        Else    ' 国民の休日
            ' 翌日を｢国民の休日｣とする
            tblCalendar(lngIxC).Syuku = 1
            tblCalendar(lngIxC).SyukuNm = g_cnsKyu2                 ' 国民の休日
        End If
        ' 次へ
        lngIxF = lngIxF + 1
    Loop
End Sub

'***************************************************************************************************
'* 処理名　：GP_MakeHoliParameter
'* 機能　　：祝日パラメータテーブルの作成
'---------------------------------------------------------------------------------------------------
'* 返り値　：(なし)
'* 引数　　：(なし)
'---------------------------------------------------------------------------------------------------
'* 作成日　：2018年09月17日
'* 作成者　：井上　治
'* 更新日　：2018年12月08日
'* 更新者　：井上　治
'* 機能説明：
'* 注意事項：本プロシージャは祝日法改正時に変更が必要！！
'***************************************************************************************************
Private Sub GP_MakeHoliParameter()
    '-----------------------------------------------------------------------------------------------
    Dim lngIxTbl As Long                                            ' テーブルINDEX
    ' 祝日パラ月別範囲テーブルを一旦初期化
    For lngIxTbl = 1 To 12
        With g_tblParamMMPos(lngIxTbl)
            .StrIx = -1
            .EndIx = -1
        End With
    Next lngIxTbl
    ReDim g_tblParamater(0)
    '===============================================================================================
    ' ※各月で呼び出している｢祝日パラメータテーブル要素追加(GP_ApendHoliParam)｣については
    ' 　この次のプロシージャのタイトルコメントを参照して下さい
    ' 　DBやExcelシートから祝日パラメータを参照してテーブル作成する機能を代替しています。
    '===============================================================================================
    ' 1月
    '-----------------------------------------------------------------------------------------------
    ' 元旦(1日)
    Call GP_ApendHoliParam(1, 0, 1, 0, 0, 0, "元旦", 0, 9999)
    ' ※引数の説明　　　　 － － － － － － －　　　－　－
    ' 　　　　　　　　　　 │ │ │ │ │ │ │　　　│　└───⑨終了年
    ' 　　　　　　　　　　 │ │ │ │ │ │ │　　　└────⑧開始年
    ' 　　　　　　　　　　 │ │ │ │ │ │ └───────⑦祝日名
    ' 　　　　　　　　　　 │ │ │ │ │ └─────── ⑥[HappyMonday]曜日(1=月曜)
    ' 　　　　　　　　　　 │ │ │ │ └────────⑤[HappyMonday]第n週
    ' 　　　　　　　　　　 │ │ │ └──────── ④[固定日]振替有無(1=振休無し)
    ' 　　　　　　　　　　 │ │ └─────────③[固定日]日付(1～31)
    ' 　　　　　　　　　　 │ └───────── ②処理区分(0=固定日､1=HappyMonday､2=会社休日､9=特殊)
    ' 　　　　　　　　　　 └──────────①月(1～12)
    ' 成人の日(1999年までは15日固定)
    Call GP_ApendHoliParam(1, 0, 15, 0, 0, 0, "成人の日", 0, 1999)
    ' 成人の日(2000年以降は第2月曜日(HappyMonday))
    Call GP_ApendHoliParam(1, 1, 0, 0, 2, 1, "成人の日", 2000, 9999)
    '===============================================================================================
    ' 2月
    '-----------------------------------------------------------------------------------------------
    ' 建国記念の日(11日)
    Call GP_ApendHoliParam(2, 0, 11, 0, 0, 0, "建国記念の日", 1967, 9999)
    ' 天皇誕生日(23日、徳仁親王:2020年より)
    Call GP_ApendHoliParam(2, 0, 23, 0, 0, 0, "天皇誕生日", 2020, 9999)
    '===============================================================================================
    ' 3月
    '-----------------------------------------------------------------------------------------------
    ' 春分の日(※特殊計算)
    Call GP_ApendHoliParam(3, 9, 0, 0, 0, 0, "春分の日", 0, 9999)
    '===============================================================================================
    ' 4月
    '-----------------------------------------------------------------------------------------------
    ' 天皇誕生日(29日、昭和天皇～1988年)
    Call GP_ApendHoliParam(4, 0, 29, 0, 0, 0, "天皇誕生日", 0, 1988)
    ' みどりの日(29日、1989～2006年)
    Call GP_ApendHoliParam(4, 0, 29, 0, 0, 0, "みどりの日", 1989, 2006)
    ' 昭和の日(29日、2007年～)
    Call GP_ApendHoliParam(4, 0, 29, 0, 0, 0, "昭和の日", 2007, 9999)
    '===============================================================================================
    ' 5月
    '-----------------------------------------------------------------------------------------------
    ' 新天皇即位の日(1日、2019年のみ、名称は暫定)
    Call GP_ApendHoliParam(5, 0, 1, 0, 0, 0, "新天皇即位の日", 2019, 2019)
    ' 憲法記念日(3日)
    Call GP_ApendHoliParam(5, 0, 3, 0, 0, 0, "憲法記念日", 0, 9999)
    ' 国民の休日(4日、2006年までは翌日が子供の日なので飛び越しの振替休日は無し)
    Call GP_ApendHoliParam(5, 0, 4, 1, 0, 0, "国民の休日", 1986, 2006)
    ' みどりの日(4日、2007年～、日曜日なら5/6に振り替え)
    Call GP_ApendHoliParam(5, 0, 4, 0, 0, 0, "みどりの日", 2007, 9999)
    ' こどもの日(5日)
    Call GP_ApendHoliParam(5, 0, 5, 0, 0, 0, "こどもの日", 0, 9999)
    '===============================================================================================
    ' 6月
    '-----------------------------------------------------------------------------------------------
    ' ※祝日無し
    '===============================================================================================
    ' 7月
    '-----------------------------------------------------------------------------------------------
    ' 海の日(2002年までは20日固定)
    Call GP_ApendHoliParam(7, 0, 20, 0, 0, 0, "海の日", 1996, 2002)
    ' 海の日(2003～2019年は第3月曜日(HappyMonday))
    Call GP_ApendHoliParam(7, 1, 0, 0, 3, 1, "海の日", 2003, 2019)
    ' 海の日(2020年のみ東京五輪特別措置法で23日)
    Call GP_ApendHoliParam(7, 0, 23, 0, 0, 0, "海の日", 2020, 2020)
    ' 海の日(2021年以降は第3月曜日(HappyMonday))
    Call GP_ApendHoliParam(7, 1, 0, 0, 3, 1, "海の日", 2021, 9999)
    ' スポーツの日(2020年のみ東京五輪特別措置法で24日、10月の｢体育の日｣の移動＋名称変更)
    Call GP_ApendHoliParam(7, 0, 24, 0, 0, 0, "スポーツの日", 2020, 2020)
    '===============================================================================================
    ' 8月
    '-----------------------------------------------------------------------------------------------
    ' 山の日(11日)
    Call GP_ApendHoliParam(8, 0, 11, 0, 0, 0, "山の日", 2016, 2019)
    ' 山の日(2020年のみ東京五輪特別措置法で10日)
    Call GP_ApendHoliParam(8, 0, 10, 0, 0, 0, "山の日", 2020, 2020)
    ' 山の日(11日)
    Call GP_ApendHoliParam(8, 0, 11, 0, 0, 0, "山の日", 2021, 9999)
    '===============================================================================================
    ' 9月
    '-----------------------------------------------------------------------------------------------
    ' 敬老の日(2002年までは15日固定)
    Call GP_ApendHoliParam(9, 0, 15, 0, 0, 0, "敬老の日", 0, 2002)
    ' 敬老の日(2003年以降は第3月曜日(HappyMonday))
    Call GP_ApendHoliParam(9, 1, 0, 0, 3, 1, "敬老の日", 2003, 9999)
    ' 秋分の日(※特殊計算)
    Call GP_ApendHoliParam(9, 9, 0, 0, 0, 0, "秋分の日", 0, 9999)
    '===============================================================================================
    ' 10月
    '-----------------------------------------------------------------------------------------------
    ' 体育の日(1999年までは10日固定)
    Call GP_ApendHoliParam(10, 0, 10, 0, 0, 0, "体育の日", 1966, 1999)
    ' 体育の日(2000年以降は第2月曜日(HappyMonday))
    Call GP_ApendHoliParam(10, 1, 0, 0, 2, 1, "体育の日", 2000, 2019)
    ' スポーツの日(名称変更(旧｢体育の日｣)、2020年は特別措置法で7月に移動)
    Call GP_ApendHoliParam(10, 1, 0, 0, 2, 1, "スポーツの日", 2021, 9999)
    ' 即位礼正殿の儀(22日、2019年のみ、名称は暫定)
    Call GP_ApendHoliParam(10, 0, 22, 0, 0, 0, "即位礼正殿の儀", 2019, 2019)
    '===============================================================================================
    ' 11月
    '-----------------------------------------------------------------------------------------------
    ' 文化の日(11日)
    Call GP_ApendHoliParam(11, 0, 3, 0, 0, 0, "文化の日", 0, 9999)
    ' 勤労感謝の日(23日)
    Call GP_ApendHoliParam(11, 0, 23, 0, 0, 0, "勤労感謝の日", 0, 9999)
    '===============================================================================================
    ' 12月
    '-----------------------------------------------------------------------------------------------
    ' 天皇誕生日(23日、平成天皇(明仁親王)～2018年)
    Call GP_ApendHoliParam(12, 0, 23, 0, 0, 0, "天皇誕生日", 1989, 2018)
End Sub

'***************************************************************************************************
'* 処理名　：GP_ApendHoliParam
'* 機能　　：祝日パラメータテーブル要素追加
'---------------------------------------------------------------------------------------------------
'* 返り値　：(なし)
'* 引数　　：Arg1 = 月(Long)
'* 　　　　　Arg2 = 処理区分(Long)              ※0=固定日､1=HappyMonday､2=会社休日､9=特殊
'* 　　　　　Arg3 = [固定日]日(Long)
'* 　　　　　Arg4 = [固定日]振替区分(Long)      ※0=通常､1=振替休日を行なわない
'* 　　　　　Arg5 = [HM]週数(Long)              ※第n週
'* 　　　　　Arg6 = [HM]曜日(Long)              ※0=日､1=月､2=火～6=土
'* 　　　　　Arg7 = 祝日名(String)
'* 　　　　　Arg8 = 開始年(西暦)(Long)
'* 　　　　　Arg9 = 終了年(西暦)(Long)
'---------------------------------------------------------------------------------------------------
'* 作成日　：2018年09月17日
'* 作成者　：井上　治
'* 更新日　：2018年10月23日
'* 更新者　：井上　治
'* 機能説明：
'* 注意事項：テーブルINDEXは本処理側で加算されます
'***************************************************************************************************
Private Sub GP_ApendHoliParam(ByVal lngGetsu As Long, _
                              ByVal lngSyoriKbn As Long, _
                              ByVal lngHiduke As Long, _
                              ByVal lngFurikaeKbn As Long, _
                              ByVal lngHmSyusu As Long, _
                              ByVal lngHmYobi As Long, _
                              ByVal strSyukuNm As String, _
                              ByVal lngStrYear As Long, _
                              ByVal lngEndYear As Long)
    '-----------------------------------------------------------------------------------------------
    Static lngIxTbl As Long                                         ' テーブルINDEX
    ReDim Preserve g_tblParamater(lngIxTbl)
    ' 祝日パラ月別範囲テーブルを更新
    With g_tblParamMMPos(lngGetsu)
        ' 開始位置
        If .StrIx < 0 Then .StrIx = lngIxTbl
        .EndIx = lngIxTbl
    End With
    ' 祝日パラメータテーブル要素追加
    With g_tblParamater(lngIxTbl)
        .Getsu = lngGetsu
        .SyoriKbn = lngSyoriKbn
        .Hiduke = lngHiduke
        .FurikaeKbn = lngFurikaeKbn
        .HmSyusu = lngHmSyusu
        .HmYobi = lngHmYobi
        .SyukuNm = strSyukuNm
        .StrYear = lngStrYear
        .EndYear = lngEndYear
    End With
    ' テーブルINDEXを加算
    lngIxTbl = lngIxTbl + 1
End Sub

'----------------------------------------<< End of Source >>----------------------------------------
