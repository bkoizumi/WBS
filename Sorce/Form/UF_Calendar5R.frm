VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UF_Calendar5R 
   Caption         =   "���t�I��"
   ClientHeight    =   2475
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2580
   OleObjectBlob   =   "UF_Calendar5R.frx":0000
   StartUpPosition =   2  '��ʂ̒���
End
Attribute VB_Name = "UF_Calendar5R"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'***************************************************************************************************
'   �J�����_�[�t�H�[��5(���t���͕��i)    �����[�U�[�t�H�[��(���D)   UF_Calendar5R(UserForm)
'
'   �쐬��:��㎡  URL:http://www.ne.jp/asahi/excel/inoue/ [Excel�ł��d��!]
'***************************************************************************************************
'   ���j���@�������͢clsAboutCalendar2R��̢GP_MakeHoliParameter����C������
'***************************************************************************************************
'�ύX���t Rev  �ύX������e------------------------------------------------------------------------>
'18/02/20(4.30)�j������֘A��modAboutCalendar2��Ɉڍs����Ή�
'18/02/20(4.30)�e�t�H�[���Ƃ̎󂯓n�����v���p�e�B�ɕύX
'18/02/21(4.31)�J�����_�[���T�̕\���s��C��(1�����\������Ȃ���)�A���[�h���X�\���Ή�
'18/02/21(4.31)���ʓ��t��Ԃ����ɖ{�t�H�[�����ŃZ��(����TextBox)�ɓ��t���������ނ悤�Ɏd�l�ύX
'18/02/22(4.32)�����Z������l�����o�����̕s��Ή�
'18/09/17(1.70)�j���p�����[�^�V�[�g���g��Ȃ������Ƃ��ăN���X��(�`�F�b�N�����p�~��Sub�v���V�[�W����)
'18/11/26(1.71)�J�����_�[�t�H�[���\��������(Windows10�Ή�)
'18/11/28(1.80)�J�����_�[�t�H�[����̊e���t���x�����N���X��(WithEvents)������Ή�
'19/12/08(1.91)UserForm_QueryClose��ǉ�(Close��Hide)
'***************************************************************************************************
Option Explicit
'---------------------------------------------------------------------------------------------------
' [�N�Z�j��] ���J�����_�[�����j�J�n(�j�����[)�ɂ���ꍇ�͢2��ɕύX���ĉ������B
Private Const g_cnsStartYobi = 1                            ' 1=���j��,2=���j��(���͕s��)
'---------------------------------------------------------------------------------------------------
' [�N�̕\�����x(From/To)]
Private Const g_cnsYearFrom = 1947                          ' �j���@�{�s
Private Const g_cnsYearToAdd = 3                            ' �V�X�e�����̔N+n�N�܂ł̎w��
'---------------------------------------------------------------------------------------------------
' �t�H�[����̐F�w�蓙�̒萔
Private Const cnsBC_Select = &HFFCC33                       ' �I����t�̔w�i�F
Private Const cnsBC_Other = &HE0E0E0                        ' �����ȊO�̔w�i�F
Private Const cnsBC_Sunday = &HFFDDFF                       ' ���j�̔w�i�F
Private Const cnsBC_Saturday = &HDDFFDD                     ' �y�j�̔w�i�F
Private Const cnsBC_Month = &HFFFFFF                        ' �����y���ȊO�̔w�i�F
Private Const cnsFC_Hori = &HFF                             ' �j���̕����F
Private Const cnsFC_Normal = &HC00000                       ' �j���ȊO�̕����F
Private Const cnsDefaultGuide = "���L�[�ő���ł��܂��B"
Private Const g_cnsDateFormat = "YYYY/MM/DD"                ' �f�t�H���g�̓��tFormat
'---------------------------------------------------------------------------------------------------
' �Ăь��Ƃ̎󂯓n���ϐ�
Private g_FormDate1 As Date                                 ' ���ݓ��t
Private g_intEntMode As Integer                             ' ���̓��[�h(0=�Z���A1=TextBox)
Private g_objRange As Range                                 ' �ΏۃZ��
Private g_objTextBox As MSForms.TextBox                     ' �Ώ�TextBox
Private g_strDateFormat As String                           ' ���t�t�H�[�}�b�g(TextBox��)
'---------------------------------------------------------------------------------------------------
' �t�H�[���\�����ɕێ����郂�W���[���ϐ�
Private g_tblYobi As Variant                                ' �j���e�[�u��
Private g_tblDateLabel(44) As New clsUF_Cal5Label1R         ' ���t���x���C�x���g�N���X�e�[�u��
Private g_tblFixLabel(11) As New clsUF_Cal5Label2R          ' �Œ胉�x���C�x���g�N���X�e�[�u��
Private g_intCurYear As Integer                             ' ���ݕ\���N
Private g_intCurMonth As Integer                            ' ���ݕ\����
Private g_CurPos As Long                                    ' ���ݓ��t�ʒu
Private g_CurPosF As Long                                   ' �����������ʒu
Private g_CurPosT As Long                                   ' �����������ʒu
Private g_swBatch As Boolean                                ' �C�x���g�}��SW
Private g_VisibleYear As Boolean                            ' Conbo�̔N�\���X�C�b�`
Private g_VisibleMonth As Boolean                           ' Combo�̌��\���X�C�b�`

'***************************************************************************************************
' ���t�H�[����̃C�x���g
'***************************************************************************************************
'* �������@�FCBO_MONTH_Change
'* �@�\�@�@�F�����R���{�̑���C�x���g
'---------------------------------------------------------------------------------------------------
'* �Ԃ�l�@�F(�Ȃ�)
'* �����@�@�F(�Ȃ�)
'---------------------------------------------------------------------------------------------------
'* �쐬���@�F2010�N01��13��
'* �쐬�ҁ@�F���@��
'* �X�V���@�F2010�N01��13��
'* �X�V�ҁ@�F���@��
'* �@�\�����F
'* ���ӎ����F
'***************************************************************************************************
Private Sub CBO_MONTH_Change()
    '-----------------------------------------------------------------------------------------------
    Dim intMonth As Integer
    If g_swBatch Then Exit Sub
    intMonth = CInt(CBO_MONTH.Text)
    g_FormDate1 = DateSerial(g_intCurYear, intMonth, 1)
    ' �N���R���{�̔�\����
    Call GP_EraseYearMonth
    ' �J�����_�[�쐬
    Call GP_MakeCalendar
End Sub

'***************************************************************************************************
'* �������@�FCBO_YEAR_Change
'* �@�\�@�@�F��N��R���{�̑���C�x���g
'---------------------------------------------------------------------------------------------------
'* �Ԃ�l�@�F(�Ȃ�)
'* �����@�@�F(�Ȃ�)
'---------------------------------------------------------------------------------------------------
'* �쐬���@�F2010�N01��13��
'* �쐬�ҁ@�F���@��
'* �X�V���@�F2010�N01��13��
'* �X�V�ҁ@�F���@��
'* �@�\�����F
'* ���ӎ����F
'***************************************************************************************************
Private Sub CBO_YEAR_Change()
    '-----------------------------------------------------------------------------------------------
    Dim intYear As Integer
    If g_swBatch Then Exit Sub
    intYear = CInt(CBO_YEAR.Text)
    g_FormDate1 = DateSerial(intYear, g_intCurMonth, 1)
    ' �N���R���{�̔�\����
    Call GP_EraseYearMonth
    ' �J�����_�[�쐬
    Call GP_MakeCalendar
End Sub

'***************************************************************************************************
'* �������@�FLBL_PREV_Click
'* �@�\�@�@�F�u��(�O��)�vClick�C�x���g
'---------------------------------------------------------------------------------------------------
'* �Ԃ�l�@�F(�Ȃ�)
'* �����@�@�F(�Ȃ�)
'---------------------------------------------------------------------------------------------------
'* �쐬���@�F2010�N01��13��
'* �쐬�ҁ@�F���@��
'* �X�V���@�F2010�N01��13��
'* �X�V�ҁ@�F���@��
'* �@�\�����F
'* ���ӎ����F
'***************************************************************************************************
Private Sub LBL_PREV_Click()
    '-----------------------------------------------------------------------------------------------
    ' �N���R���{�̔�\����
    Call GP_EraseYearMonth
    ' �O����������ݒ�
    g_FormDate1 = DateSerial(g_intCurYear, g_intCurMonth - 1, 1)
    ' �J�����_�[�쐬
    Call GP_MakeCalendar
End Sub

'***************************************************************************************************
'* �������@�FLBL_NEXT_Click
'* �@�\�@�@�F�u��(����)�vClick�C�x���g
'---------------------------------------------------------------------------------------------------
'* �Ԃ�l�@�F(�Ȃ�)
'* �����@�@�F(�Ȃ�)
'---------------------------------------------------------------------------------------------------
'* �쐬���@�F2010�N01��13��
'* �쐬�ҁ@�F���@��
'* �X�V���@�F2010�N01��13��
'* �X�V�ҁ@�F���@��
'* �@�\�����F
'* ���ӎ����F
'***************************************************************************************************
Private Sub LBL_NEXT_Click()
    '-----------------------------------------------------------------------------------------------
    ' �N���R���{�̔�\����
    Call GP_EraseYearMonth
    ' ������������ݒ�
    g_FormDate1 = DateSerial(g_intCurYear, g_intCurMonth + 1, 1)
    ' �J�����_�[�쐬
    Call GP_MakeCalendar
End Sub

'***************************************************************************************************
'* �������@�FLBL_MONTH_Click
'* �@�\�@�@�F�u���vClick�C�x���g
'---------------------------------------------------------------------------------------------------
'* �Ԃ�l�@�F(�Ȃ�)
'* �����@�@�F(�Ȃ�)
'---------------------------------------------------------------------------------------------------
'* �쐬���@�F2010�N01��13��
'* �쐬�ҁ@�F���@��
'* �X�V���@�F2010�N01��13��
'* �X�V�ҁ@�F���@��
'* �@�\�����F
'* ���ӎ����F
'***************************************************************************************************
Private Sub LBL_MONTH_Click()
    '-----------------------------------------------------------------------------------------------
    Dim intMonth As Integer
    Dim IX As Long, CUR As Long
    ' �N�R���{���\������Ă��������
    Call GP_EraseYear
    ' ���R���{�̕\��
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
'* �������@�FLBL_YEAR_Click
'* �@�\�@�@�F�u�N�vClick�C�x���g
'---------------------------------------------------------------------------------------------------
'* �Ԃ�l�@�F(�Ȃ�)
'* �����@�@�F(�Ȃ�)
'---------------------------------------------------------------------------------------------------
'* �쐬���@�F2010�N01��13��
'* �쐬�ҁ@�F���@��
'* �X�V���@�F2010�N01��13��
'* �X�V�ҁ@�F���@��
'* �@�\�����F
'* ���ӎ����F
'***************************************************************************************************
Private Sub LBL_YEAR_Click()
    '-----------------------------------------------------------------------------------------------
    Dim intYear As Integer, intYearSTR As Integer, intYearEND As Integer
    Dim IX As Long, CUR As Long
    ' ���R���{���\������Ă��������
    Call GP_EraseMonth
    ' �N�R���{�̕\��
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
'* �������@�FUserForm_Activate
'* �@�\�@�@�F�t�H�[���\��(�J��Ԃ��\���̏ꍇ��Hide�݂̂̂���Initialize�͋N���Ȃ�)
'---------------------------------------------------------------------------------------------------
'* �Ԃ�l�@�F(�Ȃ�)
'* �����@�@�F(�Ȃ�)
'---------------------------------------------------------------------------------------------------
'* �쐬���@�F2010�N01��13��
'* �쐬�ҁ@�F���@��
'* �X�V���@�F2018�N02��22��
'* �X�V�ҁ@�F���@��
'* �@�\�����F
'* ���ӎ����F
'***************************************************************************************************
Private Sub UserForm_Activate()
    '-----------------------------------------------------------------------------------------------
    ' �R���{�͔�\��
    CBO_YEAR.Visible = False
    CBO_MONTH.Visible = False
    g_VisibleYear = False
    g_VisibleMonth = False
    g_FormDate1 = 0
    '-----------------------------------------------------------------------------------------------
    ' ���̓��[�h(0=�Z���A1=TextBox)
    If g_intEntMode = 1 Then
        ' 1=TextBox
        ' ���ƂȂ���t���e�L�X�g�{�b�N�X����擾
        If IsDate(Trim(g_objTextBox.Text)) Then
            g_FormDate1 = CDate(Trim(g_objTextBox.Text))
        End If
    Else
        ' 0=�Z��
        ' ���ƂȂ���t���Z������擾
        On Error Resume Next
        If IsDate(Trim(g_objRange.Cells(1).Value)) Then
            g_FormDate1 = CDate(Trim(g_objRange.Cells(1).Value))
        End If
        On Error GoTo 0
    End If
    ' �󂯎��Ȃ��ꍇ�͓������Z�b�g
    If g_FormDate1 = 0 Then g_FormDate1 = Date
    '-----------------------------------------------------------------------------------------------
    ' �J�����_�[�쐬
    Call GP_MakeCalendar
    LBL_GUIDE.Caption = cnsDefaultGuide             ' �K�C�h�\��
    ' �\���ʒu���}�j���A���ɕύX
    If Me.StartUpPosition <> 0 Then Me.StartUpPosition = 0
End Sub

'***************************************************************************************************
'* �������@�FUserForm_Deactivate
'* �@�\�@�@�F�t�H�[����A�N�e�B�u���
'---------------------------------------------------------------------------------------------------
'* �Ԃ�l�@�F(�Ȃ�)
'* �����@�@�F(�Ȃ�)
'---------------------------------------------------------------------------------------------------
'* �쐬���@�F2018�N02��21��
'* �쐬�ҁ@�F���@��
'* �X�V���@�F2018�N02��21��
'* �X�V�ҁ@�F���@��
'* �@�\�����F
'* ���ӎ����FUserForm�ォ��̗��p�̑Ή��̂��߃��[�h���X�ɂł��Ȃ����Ƃ��炱�̃C�x���g�͔������Ȃ�
'***************************************************************************************************
Private Sub UserForm_Deactivate()
    '-----------------------------------------------------------------------------------------------
    Me.Hide
End Sub

'***************************************************************************************************
'* �������@�FUserForm_Initialize
'* �@�\�@�@�F�t�H�[��������(�J��Ԃ��\���̏ꍇ��Hide�݂̂̂���Initialize�͋N���Ȃ�)
'---------------------------------------------------------------------------------------------------
'* �Ԃ�l�@�F(�Ȃ�)
'* �����@�@�F(�Ȃ�)
'---------------------------------------------------------------------------------------------------
'* �쐬���@�F2010�N01��13��
'* �쐬�ҁ@�F���@��
'* �X�V���@�F2018�N11��28��
'* �X�V�ҁ@�F���@��
'* �@�\�����F
'* ���ӎ����F
'***************************************************************************************************
Private Sub UserForm_Initialize()
    '-----------------------------------------------------------------------------------------------
    Dim lngIx As Long                                               ' �e�[�u��INDEX
    Dim lngIx2 As Long                                              ' �e�[�u��INDEX
    Dim lngIxC As Long                                              ' �J�����_�[�e�[�u��INDEX
    Dim dteDate As Date                                             ' ���tWORK
    Dim tblTodayNm As Variant                                       ' ��������������̖���
    Dim tblCalendar() As g_typAboutCalendar2                        ' �J�����_�[�e�[�u��
    '-----------------------------------------------------------------------------------------------
    ' �e���t���x���C�x���g�N���X�̏�����
    For lngIx = 0 To 44
        Call g_tblDateLabel(lngIx).NewClass(Me.Controls("LBL_" & Format((lngIx + 1), "00")), lngIx)
    Next lngIx
    ' �e�Œ胉�x���C�x���g�N���X�̏�����
    Call g_tblFixLabel(0).NewClass(LBL_SUN, cnsDefaultGuide)        ' ��
    Call g_tblFixLabel(1).NewClass(LBL_MON, cnsDefaultGuide)        ' ��
    Call g_tblFixLabel(2).NewClass(LBL_TUE, cnsDefaultGuide)        ' ��
    Call g_tblFixLabel(3).NewClass(LBL_WED, cnsDefaultGuide)        ' ��
    Call g_tblFixLabel(4).NewClass(LBL_THU, cnsDefaultGuide)        ' ��
    Call g_tblFixLabel(5).NewClass(LBL_FRI, cnsDefaultGuide)        ' ��
    Call g_tblFixLabel(6).NewClass(LBL_SAT, cnsDefaultGuide)        ' �y
    Call g_tblFixLabel(7).NewClass(LBL_PREV, "�O���ɖ߂�܂�(PageUp)") ' ��
    Call g_tblFixLabel(8).NewClass(LBL_NEXT, "�����ɐi�݂܂�(PageDown)") ' ��
    Call g_tblFixLabel(9).NewClass(LBL_YM, "�N������I�����܂��B")  ' �N��
    Call g_tblFixLabel(10).NewClass(LBL_YEAR, "�N���I���ł��܂��B") ' �N
    Call g_tblFixLabel(11).NewClass(LBL_MONTH, "�����I���ł��܂��B") ' ��
    '-----------------------------------------------------------------------------------------------
    ' �N�Z�j���ɂ��j�����o���̈ʒu�C��
    If g_cnsStartYobi = 2 Then
        ' ���j�N�Z
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
        ' ���j�N�Z
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
    g_tblYobi = Array("(��)", "(��)", "(��)", "(��)", "(��)", "(��)", "(�y)")
    '-----------------------------------------------------------------------------------------------
    ' ��������������̏���
    dteDate = Date              ' ����
    ' �J�����_�[�e�[�u���쐬(����+�O���3�����p)
    Call modAboutCalendar2R.GP_GetCalendarTable3(Year(dteDate), Month(dteDate), tblCalendar)
    ' ����ɖ߂�
    dteDate = dteDate - 1
    lngIxC = 0
    ' �J�����_�[��̍���̈ʒu�𔻒�
    Do While lngIxC <= UBound(tblCalendar)
        ' ���t�����͔�����
        If tblCalendar(lngIxC).Hiduke = dteDate Then Exit Do
        ' ���̓���
        lngIxC = lngIxC + 1
    Loop
    tblTodayNm = Array("[���]", "[����]", "[����]")
    lngIx2 = 0
    ' ��������������̃Z�b�g
    For lngIx = 42 To 44
        dteDate = tblCalendar(lngIxC).Hiduke
        ' ���t�R���g���[�����e�[�u���̃Z�b�g
        With g_tblDateLabel(lngIx)
            .Hiduke = dteDate
            .Yobi = tblCalendar(lngIxC).Yobi
            .Syuku = tblCalendar(lngIxC).Syuku
            .StsGuide = tblTodayNm(lngIx2) & Format(dteDate, g_cnsDateFormat) & g_tblYobi(.Yobi)
            ' �j����
            If .Syuku <> 0 Then
                .StsGuide = .StsGuide & " " & tblCalendar(lngIxC).SyukuNm
            End If
        End With
        ' �J�����_�[�e�[�u���ʒu�𗂓��Ɉړ�
        lngIxC = lngIxC + 1
        lngIx2 = lngIx2 + 1
    Next lngIx
End Sub

'***************************************************************************************************
'* �������@�FUserForm_KeyDown
'* �@�\�@�@�F�t�H�[����L�[�{�[�h����
'---------------------------------------------------------------------------------------------------
'* �Ԃ�l�@�F(�Ȃ�)
'* �����@�@�F(����)
'---------------------------------------------------------------------------------------------------
'* �쐬���@�F2010�N01��13��
'* �쐬�ҁ@�F���@��
'* �X�V���@�F2010�N01��13��
'* �X�V�ҁ@�F���@��
'* �@�\�����F
'* ���ӎ����F
'***************************************************************************************************
Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, _
                             ByVal Shift As Integer)
    '-----------------------------------------------------------------------------------------------
    ' KeyCode(Shift���p)�ɂ�鐧��
    Select Case KeyCode
        Case vbKeyReturn, vbKeyExecute, vbKeySeparator  ' Enter(����)
            Call GP_ClickCalendar(g_CurPos)
        Case vbKeyCancel, vbKeyEscape                   ' Cancel, Esc(�I��)
            Me.Hide
        Case vbKeyPageDown                              ' PageDown(����)
            Call LBL_NEXT_Click
        Case vbKeyPageUp                                ' KeyPageUp(�O��)
            Call LBL_PREV_Click
        Case vbKeyRight, vbKeyNumpad6, vbKeyAdd         ' ��(����)
            Call GP_MoveDay(1)
        Case vbKeyLeft, vbKeyNumpad4, vbKeySubtract     ' ��(�O��)
            Call GP_MoveDay(-1)
        Case vbKeyUp, vbKeyNumpad8                      ' �|(7����)
            Call GP_MoveDay(-7)
        Case vbKeyDown, vbKeyNumpad2                    ' �{(7���O)
            Call GP_MoveDay(7)
        Case vbKeyHome                                  ' Home(����)
            Call GP_MoveDay(g_CurPosF - g_CurPos)
        Case vbKeyEnd                                   ' End(����)
            Call GP_MoveDay(g_CurPosT - g_CurPos)
        Case vbKeyTab                                   ' Tab(Shift�ɂ��)
            If Shift = 1 Then
                Call GP_MoveDay(-1)                     ' �O��
            Else
                Call GP_MoveDay(1)                      ' ����
            End If
        Case vbKeyF11                                   ' F11(�O�N)
            g_FormDate1 = DateSerial(g_intCurYear - 1, g_intCurMonth, 1)
            ' �N���R���{�̔�\����
            Call GP_EraseYearMonth
            ' �J�����_�[�쐬
            Call GP_MakeCalendar
        Case vbKeyF12                                   ' F11(���N)
            g_FormDate1 = DateSerial(g_intCurYear + 1, g_intCurMonth, 1)
            ' �N���R���{�̔�\����
            Call GP_EraseYearMonth
            ' �J�����_�[�쐬
            Call GP_MakeCalendar
    End Select
End Sub

'***************************************************************************************************
'* �������@�FUserForm_MouseMove
'* �@�\�@�@�F�t�H�[����}�E�X�ړ�
'---------------------------------------------------------------------------------------------------
'* �Ԃ�l�@�F(�Ȃ�)
'* �����@�@�F(����)
'---------------------------------------------------------------------------------------------------
'* �쐬���@�F2010�N01��13��
'* �쐬�ҁ@�F���@��
'* �X�V���@�F2010�N01��13��
'* �X�V�ҁ@�F���@��
'* �@�\�����F
'* ���ӎ����F
'***************************************************************************************************
Private Sub UserForm_MouseMove(ByVal Button As Integer, _
                               ByVal Shift As Integer, _
                               ByVal x As Single, ByVal y As Single)
    '-----------------------------------------------------------------------------------------------
    Me.LBL_GUIDE.Caption = cnsDefaultGuide
End Sub

'***************************************************************************************************
'* �������@�FUserForm_QueryClose
'* �@�\�@�@�F�t�H�[���C�x���g(QueryClose)
'---------------------------------------------------------------------------------------------------
'* �Ԃ�l�@�F(�Ȃ�)
'* �����@�@�F(����)
'---------------------------------------------------------------------------------------------------
'* �쐬���@�F2019�N12��08��
'* �쐬�ҁ@�F���@��
'* �X�V���@�F2019�N12��08��
'* �X�V�ҁ@�F���@��
'* �@�\�����F
'* ���ӎ����F
'***************************************************************************************************
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    '-----------------------------------------------------------------------------------------------
    ' Close��Hide�ɒu��������
    If CloseMode = vbFormControlMenu Then
        Me.Hide
        Cancel = True
    End If
End Sub

'***************************************************************************************************
' �����ʃT�u����(Private)
'***************************************************************************************************
'* �������@�FGP_MakeCalendar
'* �@�\�@�@�F�J�����_�[�\��
'---------------------------------------------------------------------------------------------------
'* �Ԃ�l�@�F(�Ȃ�)
'* �����@�@�F(�Ȃ�)
'---------------------------------------------------------------------------------------------------
'* �쐬���@�F2010�N01��13��
'* �쐬�ҁ@�F���@��
'* �X�V���@�F2018�N11��28��
'* �X�V�ҁ@�F���@��
'* �@�\�����F
'* ���ӎ����F
'***************************************************************************************************
Private Sub GP_MakeCalendar()
    '-----------------------------------------------------------------------------------------------
    Dim dteDate As Date                                             ' ���tWORK
    Dim intYobi As Integer                                          ' �j���ʒuINDEX
    Dim intYear As Integer                                          ' �w��N
    Dim lngCurStrIx As Long                                         ' �����J�nINDEX
    Dim lngCurEndIx As Long                                         ' �����I��INDEX
    Dim lngIx As Long                                               ' �t�H�[����ʒuINDEX
    Dim lngIx2 As Long                                              ' �e�[�u��INDEX(Work)
    Dim lngIxC As Long                                              ' �J�����_�[INDEX
    Dim lngIxCMax As Long                                           ' �J�����_�[INDEX���
    Dim lngCurPos As Long                                           ' �I����̈ʒu
    Dim tblCalendar() As g_typAboutCalendar2                        ' �J�����_�[�e�[�u��
    '-----------------------------------------------------------------------------------------------
    intYear = Year(g_FormDate1)                                 ' �w��N
    ' �w��N�������p�\���`�F�b�N
    If ((intYear < g_cnsYearFrom) Or _
        (intYear > (Year(Date) + g_cnsYearToAdd))) Then
        MsgBox "�j���v�Z�͈͂𒴂��Ă��܂��B", vbExclamation, Me.Caption
        g_FormDate1 = g_tblDateLabel(g_CurPos).Hiduke
    End If
    g_intCurYear = Year(g_FormDate1)                            ' �w��N
    g_intCurMonth = Month(g_FormDate1)                          ' �w�茎
    LBL_YM.Caption = g_intCurYear & "�N" & Format(g_intCurMonth, "00") & "��"
    '-----------------------------------------------------------------------------------------------
    ' �J�����_�[�e�[�u���쐬(����+�O���3�����p)
    Call modAboutCalendar2R.GP_GetCalendarTable3R(g_intCurYear, _
                                                  g_intCurMonth, _
                                                  tblCalendar, _
                                                  lngCurStrIx, _
                                                  lngCurEndIx)
    lngIxCMax = UBound(tblCalendar)
    ' �J�����_�[�e�[�u���擪�ʒu
    lngIxC = lngCurStrIx
    ' �w����t�����U�A�O�T�̍ŏI��(���j��)�ɖ߂�
    lngIxC = lngIxC - tblCalendar(lngCurStrIx).Yobi
    ' ���j�n�܂莞�̒���
    If g_cnsStartYobi = 2 Then
        lngIxC = lngIxC + 1
        ' 2���n�܂�ɂȂ��Ă��܂�����1�T�߂�
        If lngIxC > lngCurStrIx Then lngIxC = lngIxC - 7
    End If
    intYobi = 0
    lngCurPos = -1
    '-----------------------------------------------------------------------------------------------
    ' �t�H�[����̓��t�Z�b�g(7�j�~6�T=42���Œ��0�n�܂�)
    For lngIx = 0 To 41
        ' ���ʒu�̓��t�A�j�����Z�o
        intYobi = intYobi + 1
        If intYobi > 7 Then intYobi = 1
        dteDate = tblCalendar(lngIxC).Hiduke
        ' ���ݑI�����
        If dteDate = g_FormDate1 Then
            lngCurPos = lngIx
        End If
        ' ���t�R���g���[�����e�[�u���̃Z�b�g
        With g_tblDateLabel(lngIx)
            .Hiduke = dteDate
            .Yobi = tblCalendar(lngIxC).Yobi
            .Syuku = tblCalendar(lngIxC).Syuku
            .StsGuide = Format(dteDate, g_cnsDateFormat) & g_tblYobi(.Yobi)
            ' �j����
            If .Syuku <> 0 Then
                .StsGuide = .StsGuide & " " & tblCalendar(lngIxC).SyukuNm
            End If
        End With
        ' �������A�������̈ʒu���擾
        If lngIxC = lngCurStrIx Then
            ' ��������
            g_CurPosF = lngIx
        ElseIf lngIxC = lngCurEndIx Then
            ' ��������
            g_CurPosT = lngIx
        End If
        ' ���x���R���g���[����z�񉻂����ϐ�
        g_tblDateLabel(lngIx).Label.Caption = day(dteDate)
        ' �����F�A�w�i�F�̃Z�b�g
        Call GP_SetForeColor(lngIx, g_FormDate1)
        ' �J�����_�[�e�[�u���ʒu�𗂓��Ɉړ�
        lngIxC = lngIxC + 1
    Next lngIx
    LBL_GUIDE.Caption = g_tblDateLabel(g_CurPos).StsGuide   ' �K�C�h�\��
End Sub

'***************************************************************************************************
'* �������@�FGP_MoveDay
'* �@�\�@�@�F�J�����_�[��̑I���ʒu�ړ�
'---------------------------------------------------------------------------------------------------
'* �Ԃ�l�@�F(�Ȃ�)
'* �����@�@�FArg1 = �ړ���(Long)                ���}�C�i�X����
'---------------------------------------------------------------------------------------------------
'* �쐬���@�F2010�N01��13��
'* �쐬�ҁ@�F���@��
'* �X�V���@�F2018�N11��28��
'* �X�V�ҁ@�F���@��
'* �@�\�����F
'* ���ӎ����F
'***************************************************************************************************
Private Sub GP_MoveDay(lngMove As Long)
    '-----------------------------------------------------------------------------------------------
    Dim lngPos As Long                                              ' �e�[�u���ʒuINDEX
    Dim dteDate As Date                                             ' ���tWork
    ' �N���R���{�̔�\����
    Call GP_EraseYearMonth
    ' �ړ���̈ʒu,���t���Z�o
    lngPos = g_CurPos + lngMove                                 ' �ړ���ʒu
    dteDate = g_FormDate1 + lngMove                             ' �ړ�����t
    ' �R���g���[���e�[�u���O��
    If ((lngPos < 0) Or (lngPos > 41)) Then
        ' �O�����͗����Ɉړ�
        g_FormDate1 = dteDate
        Call GP_MakeCalendar
        Exit Sub
    End If
    '-----------------------------------------------------------------------------------------------
    ' �ȑO�̈ʒu�̓��t���x���̔w�i�F�����ɖ߂�
    Call GP_SetForeColor(g_CurPos, dteDate)
    '-----------------------------------------------------------------------------------------------
    ' ���ݓ��t(�ޔ�)���X�V
    g_FormDate1 = dteDate
    g_CurPos = lngPos
    ' ����̈ʒu�̓��t���x���̔w�i�F��I����ԂɕύX
    Call GP_SetForeColor(g_CurPos, g_FormDate1)
    LBL_GUIDE.Caption = g_tblDateLabel(g_CurPos).StsGuide   ' �K�C�h�\��
End Sub

'***************************************************************************************************
'* �������@�FGP_SetForeColor
'* �@�\�@�@�F�����F�A�w�i�F�̃Z�b�g
'---------------------------------------------------------------------------------------------------
'* �Ԃ�l�@�F(�Ȃ�)
'* �����@�@�FArg1 = ���ݓ��t�ʒuINDEX(Long)
'* �@�@�@�@�@Arg2 = �I����t(Date)
'---------------------------------------------------------------------------------------------------
'* �쐬���@�F2018�N02��20��
'* �쐬�ҁ@�F���@��
'* �X�V���@�F2018�N11��28��
'* �X�V�ҁ@�F���@��
'* �@�\�����F
'* ���ӎ����F�ʒuINDEX�̓��x���R���g���[���z���̈ʒu
'***************************************************************************************************
Private Sub GP_SetForeColor(ByVal lngPos As Long, ByVal dteCurDate As Date)
    '-----------------------------------------------------------------------------------------------
    Dim dteDate As Date                                             ' ���tWork
    Dim lngYear As Long                                             ' ���ݔN
    Dim lngMonth As Long                                            ' ���݌�
    With g_tblDateLabel(lngPos)
        dteDate = .Hiduke
        lngYear = Year(dteDate)
        lngMonth = Month(dteDate)
        ' ���x�A�j���ɂ�胉�x���̏������Z�b�g
        .Label.Font.Bold = False
        .Label.ForeColor = cnsFC_Normal
        ' ���ݑI�����
        If dteDate = dteCurDate Then
            ' �����I����t
            .Label.BackColor = cnsBC_Select
            g_CurPos = lngPos
        ElseIf ((lngYear = g_intCurYear) And (lngMonth = g_intCurMonth)) Then   ' ��������
            ' ����
            Select Case .Yobi
                Case 0                  ' ���j��
                    .Label.BackColor = cnsBC_Sunday
                Case 6                  ' �y�j��
                    .Label.BackColor = cnsBC_Saturday
                Case Else
                    .Label.BackColor = cnsBC_Month
            End Select
        Else
            ' �����ȊO
            .Label.BackColor = cnsBC_Other
        End If
        ' �j��(�ܐU�֋x��)�̔���
        If .Syuku <> 0 Then
            ' �����F��ԂƂ���
            .Label.ForeColor = cnsFC_Hori
            ' �����͑���
            If ((lngYear = g_intCurYear) And (lngMonth = g_intCurMonth)) Then .Label.Font.Bold = True
        End If
    End With
End Sub

'***************************************************************************************************
'* �������@�FGP_ShowGuide
'* �@�\�@�@�F�X�e�[�^�X�K�C�h�\��
'---------------------------------------------------------------------------------------------------
'* �Ԃ�l�@�F(�Ȃ�)
'* �����@�@�FArg1 = ���ݓ��t�ʒuINDEX(Long)
'---------------------------------------------------------------------------------------------------
'* �쐬���@�F2018�N02��20��
'* �쐬�ҁ@�F���@��
'* �X�V���@�F2018�N11��28��
'* �X�V�ҁ@�F���@��
'* �@�\�����F
'* ���ӎ����F�ʒuINDEX�̓��x���R���g���[���z���̈ʒu
'***************************************************************************************************
Friend Sub GP_ShowGuide(ByVal lngPos As Long)
    '-----------------------------------------------------------------------------------------------
    LBL_GUIDE.Caption = g_tblDateLabel(lngPos).StsGuide
End Sub

'***************************************************************************************************
'* �������@�FGP_ClickCalendar
'* �@�\�@�@�F�J�����_�[�N���b�N
'---------------------------------------------------------------------------------------------------
'* �Ԃ�l�@�F(�Ȃ�)
'* �����@�@�FArg1 = ���t�R���g���[�����e�[�u��INDEX(Long)
'---------------------------------------------------------------------------------------------------
'* �쐬���@�F2010�N01��13��
'* �쐬�ҁ@�F���@��
'* �X�V���@�F2018�N11��28��
'* �X�V�ҁ@�F���@��
'* �@�\�����F
'* ���ӎ����F
'***************************************************************************************************
Friend Sub GP_ClickCalendar(ByVal lngPos As Long)
    '-----------------------------------------------------------------------------------------------
    g_FormDate1 = g_tblDateLabel(lngPos).Hiduke
    ' �N���R���{�̔�\����
    Call GP_EraseYearMonth
    ' ���̓��[�h(0=�Z���A1=TextBox)
    If g_intEntMode = 1 Then
        ' 1=TextBox
        g_objTextBox.Text = Format(g_FormDate1, g_strDateFormat)
    Else
        ' 0=�Z��
        g_objRange.Value = g_FormDate1
    End If
    Me.Hide
End Sub

'***************************************************************************************************
'* �������@�FGP_EraseYearMonth
'* �@�\�@�@�F��N������R���{�̔�\����
'---------------------------------------------------------------------------------------------------
'* �Ԃ�l�@�F(�Ȃ�)
'* �����@�@�F(�Ȃ�)
'---------------------------------------------------------------------------------------------------
'* �쐬���@�F2010�N01��13��
'* �쐬�ҁ@�F���@��
'* �X�V���@�F2010�N01��13��
'* �X�V�ҁ@�F���@��
'* �@�\�����F
'* ���ӎ����F
'***************************************************************************************************
Private Sub GP_EraseYearMonth()
    '-----------------------------------------------------------------------------------------------
    Call GP_EraseYear
    Call GP_EraseMonth
End Sub

'***************************************************************************************************
'* �������@�FGP_EraseYear
'* �@�\�@�@�F��N��R���{�̔�\����
'---------------------------------------------------------------------------------------------------
'* �Ԃ�l�@�F(�Ȃ�)
'* �����@�@�F(�Ȃ�)
'---------------------------------------------------------------------------------------------------
'* �쐬���@�F2010�N01��13��
'* �쐬�ҁ@�F���@��
'* �X�V���@�F2010�N01��13��
'* �X�V�ҁ@�F���@��
'* �@�\�����F
'* ���ӎ����F
'***************************************************************************************************
Private Sub GP_EraseYear()
    '-----------------------------------------------------------------------------------------------
    If g_VisibleYear Then
        CBO_YEAR.Visible = False
        g_VisibleYear = False
    End If
End Sub

'***************************************************************************************************
'* �������@�FGP_EraseMonth
'* �@�\�@�@�F�����R���{�̔�\����
'---------------------------------------------------------------------------------------------------
'* �Ԃ�l�@�F(�Ȃ�)
'* �����@�@�F(�Ȃ�)
'---------------------------------------------------------------------------------------------------
'* �쐬���@�F2010�N01��13��
'* �쐬�ҁ@�F���@��
'* �X�V���@�F2010�N01��13��
'* �X�V�ҁ@�F���@��
'* �@�\�����F
'* ���ӎ����F
'***************************************************************************************************
Private Sub GP_EraseMonth()
    '-----------------------------------------------------------------------------------------------
    If g_VisibleMonth Then
        CBO_MONTH.Visible = False
        g_VisibleMonth = False
    End If
End Sub

'***************************************************************************************************
' ������ �v���p�e�B ������
'***************************************************************************************************
' �^�C�g��
'---------------------------------------------------------------------------------------------------
Friend Property Let prpTitle(ByVal strTitle As String)
    Me.Caption = strTitle
End Property

'===================================================================================================
' ���̓��[�h(1=�Z���A2=TextBox)
'---------------------------------------------------------------------------------------------------
Friend Property Let prpEntMode(ByVal intValue As Integer)
    g_intEntMode = intValue
End Property

'===================================================================================================
' �ΏۃZ��(Object)
'---------------------------------------------------------------------------------------------------
Friend Property Set prpRange(ByRef objValue As Range)
    Set g_objRange = objValue
End Property

'===================================================================================================
' �Ώ�TextBox(Object)
'---------------------------------------------------------------------------------------------------
Friend Property Set prpTextBox(ByRef objValue As MSForms.TextBox)
    Set g_objTextBox = objValue
End Property

'===================================================================================================
' ���t�t�H�[�}�b�g(TextBox��)
'---------------------------------------------------------------------------------------------------
Friend Property Let prpDateFormat(ByVal strFormat As String)
    g_strDateFormat = strFormat
End Property

'------------------------------------------<< End of Source >>--------------------------------------

