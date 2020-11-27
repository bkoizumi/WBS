Attribute VB_Name = "modCalendar5R"
'***************************************************************************************************
'   �J�����_�[�t�H�[��5(���t���͕��i)   ���Ăяo���v���V�[�W��    modCalendar5R(Module)
'
'   �쐬��:��㎡  URL:http://www.ne.jp/asahi/excel/inoue/ [Excel�ł��d��!]
'***************************************************************************************************
'�ύX���t Rev  �ύX������e========================================================================>
'18/02/21(1.00)�V�K�쐬
'18/11/28(1.80)�J�����_�[�t�H�[����̊e���t���x�����N���X��(WithEvents)������Ή�
'19/10/20(1.90)64�r�b�gWindows�Ή�
'19/12/08(1.95)�����X�N���[���Ή��A�X�N���[�����[�E�E�[����̂͂ݏo���Ή�
'***************************************************************************************************
Option Explicit
'===================================================================================================
Private Const g_cnsDateFormat = "YYYY/MM/DD"                    ' �f�t�H���g�̓��tFormat
Private Const g_cnsCaption = "���t�I��"                         ' �f�t�H���g��Caption
'===================================================================================================
' �t�H�[���ʒu����֘A
Private Const LOGPIXELSX As Long = 88
Private Const LOGPIXELSY As Long = 90
Private Const SM_CYSCREEN As Long = 1
Private Const SM_XVIRTUALSCREEN As Long = 76
Private Const SM_YVIRTUALSCREEN As Long = 77
Private Const SM_CXVIRTUALSCREEN As Long = 78
Private Const SM_CYVIRTUALSCREEN As Long = 79
Private Const SPI_GETWORKAREA As Long = 48
'===================================================================================================
' GetWindowRect�p���[�U�[��`
Private Type g_typRect
    Left As Long
    top As Long
    Right As Long
    Bottom As Long
End Type
' 64�r�b�g�Ŕ���
#If VBA7 Then
' ��GetDC(API)
Private Declare PtrSafe Function GetDC Lib "User32.dll" (ByVal hWnd As LongPtr) As LongPtr
' ��ReleaseDC(API)
Private Declare PtrSafe Function ReleaseDC Lib "User32.dll" _
    (ByVal hWnd As LongPtr, ByVal hdc As LongPtr) As Long
' ��GetDeviceCaps(API)
Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32.dll" _
    (ByVal hdc As LongPtr, ByVal nIndex As Long) As Long
' ��GetSystemMetrics(API)
Private Declare PtrSafe Function getSystemMetrics Lib "User32.dll" Alias "GetSystemMetrics" (ByVal nIndex As Long) As Long
' ��SystemParametersInfo(API)
Private Declare PtrSafe Function SystemParametersInfo Lib "User32.dll" _
    Alias "SystemParametersInfoA" ( _
    ByVal uAction As Long, _
    ByVal uParam As Long, _
    ByRef lpvParam As g_typRect, _
    ByVal fuWinIni As Long) As Long
#Else
' ��GetDC(API)
Private Declare Function GetDC Lib "User32.dll" (ByVal hWnd As Long) As Long
' ��ReleaseDC(API)
Private Declare Function ReleaseDC Lib "User32.dll" (ByVal hWnd As Long, ByVal hdc As Long) As Long
' ��GetDeviceCaps(API)
Private Declare Function GetDeviceCaps Lib "gdi32.dll" (ByVal hdc As Long, ByVal nIndex As Long) As Long
' ��GetSystemMetrics(API)
Private Declare Function getSystemMetrics Lib "User32.dll" Alias "GetSystemMetrics" (ByVal nIndex As Long) As Long
' ��SystemParametersInfo(API)
Private Declare Function SystemParametersInfo Lib "User32.dll" _
    Alias "SystemParametersInfoA" ( _
    ByVal uAction As Long, _
    ByVal uParam As Long, _
    ByRef lpvParam As g_typRect, _
    ByVal fuWinIni As Long) As Long
#End If

'***************************************************************************************************
'* �������@�FShowCalendarFromTextBox2
'* �@�\�@�@�F���[�U�[�t�H�[���̃e�L�X�g�{�b�N�X(MsForms.TextBox)����\��������
'===================================================================================================
'* �Ԃ�l�@�F(�Ȃ�)
'* �����@�@�FArg1 = �e�L�X�g�{�b�N�X(Object)
'* �@�@�@�@�@Arg2 = �J�����_�[�t�H�[���̕\���ʒu�F��(Long)  ��Option
'* �@�@�@�@�@Arg3 = �J�����_�[�t�H�[���̕\���ʒu�F�c(Long)  ��Option
'* �@�@�@�@�@Arg4 = �J�����_�[�t�H�[����Caption(String)     ��Option�A�f�t�H���g��"���t�I��"
'* �@�@�@�@�@Arg5 = �l��Ԃ�����Format(String)              ��Option�A�f�t�H���g��"YYYY/MM/DD"
'===================================================================================================
'* �쐬���@�F2018�N02��21��
'* �쐬�ҁ@�F���@��
'* �X�V���@�F2019�N12��08��
'* �X�V�ҁ@�F���@��
'* �@�\�����F
'* ���ӎ����F
'***************************************************************************************************
Public Sub ShowCalendarFromTextBox2(ByRef objTextBox As MSForms.TextBox, _
                                    Optional ByVal lngLeft As Long = 0, _
                                    Optional ByVal lngTop As Long = 0, _
                                    Optional ByVal strCaption As String = g_cnsCaption, _
                                    Optional ByVal strFormat As String = g_cnsDateFormat)
    '==============================================================================================-
    Dim lngFormWidth As Long                                        ' �J�����_�[�t�H�[���̕�
    Dim lngFormHeight As Long                                       ' �J�����_�[�t�H�[���̍���
    Dim lngScreenRight As Long                                      ' �X�N���[���E�[�ʒu
    Dim lngScreenBottom As Long                                     ' �X�N���[�����[�ʒu
    Dim lngDPIX As Long                                             ' Dots Per Inch(����)
    Dim lngDPIY As Long                                             ' Dots Per Inch(����)
    Dim lngPPI As Long                                              ' Pixels Per Inch
    '==============================================================================================-
    ' �J�����_�[�t�H�[���̑傫���擾
    lngFormWidth = UF_Calendar5R.Width
    lngFormHeight = UF_Calendar5R.Height
    ' ���ȉ���Excel2003�ȑO�ł͓��삵�Ȃ�
    lngDPIX = FP_GetDPIX
    lngDPIY = FP_GetDPIY
    lngPPI = FP_GetPPI
    '==============================================================================================-
    ' �X�N���[���T�C�Y�ʒu�̎擾
    Call GP_GetScreenPos(0, 0, lngScreenRight, lngScreenBottom)
    '==============================================================================================-
    ' �J�����_�[�t�H�[�����X�N���[������͂ݏo����(��)
    If (lngLeft + lngFormWidth) * (lngDPIX / lngPPI) > lngScreenRight Then
        ' �X�N���[���E�[�Ɉړ�(+3�͌덷�H)
        lngLeft = lngScreenRight * (lngPPI / lngDPIX) - lngFormWidth + 3
    End If
    ' �J�����_�[�t�H�[�����X�N���[������͂ݏo����(�c)
    If (lngTop + lngFormHeight) * (lngDPIY / lngPPI) > lngScreenBottom Then
        ' �Z����[�Ɉړ�(+3�͌덷�H)
        lngTop = lngTop - (objTextBox.Height + lngFormHeight) + 3
    End If
    '==============================================================================================-
    ' �J�����_�[�t�H�[��
    With UF_Calendar5R
        .prpTitle = strCaption                              ' �^�C�g��
        .prpEntMode = 1                                     ' ���̓��[�h(0=�Z���A1=TextBox)
        Set .prpTextBox = objTextBox                        ' �Ώ�TextBox
        .prpDateFormat = strFormat                          ' ���t�t�H�[�}�b�g
        ' �t�H�[���\���ʒu�̊m�F
        If ((lngLeft <> 0) Or (lngTop <> 0)) Then
            ' �w�肪����ꍇ�̓}�j���A���w��
            .StartUpPosition = 0
            .Left = lngLeft
            .top = lngTop
        Else
            ' �w�肪�Ȃ��ꍇ�̓X�N���[���̒���
            .StartUpPosition = 2
        End If
        ' �J�����_�[�t�H�[����\��
        .Show
    End With
End Sub

'***************************************************************************************************
'* �������@�FShowCalendarFromRange2
'* �@�\�@�@�F�Z��(Range)����\��������
'===================================================================================================
'* �Ԃ�l�@�F(�Ȃ�)
'* �����@�@�FArg1 = �Z��(Object) ���P��Z�����͌����������t�p�Z��
'* �@�@�@�@�@Arg2 = �J�����_�[�t�H�[����Caption(String)     ��Option�A�f�t�H���g��"���t�I��"
'===================================================================================================
'* �쐬���@�F2018�N02��21��
'* �쐬�ҁ@�F���@��
'* �X�V���@�F2019�N12��08��
'* �X�V�ҁ@�F���@��
'* �@�\�����F���Y�Z���̉��ɃJ�����_�[�t�H�[�����\�������
'* ���ӎ����F
'***************************************************************************************************
Public Sub ShowCalendarFromRange2(ByRef objRange As Range, _
                                  Optional ByVal strCaption As String = g_cnsCaption)
    '==============================================================================================-
    Dim lngLeft As Long                                             ' ���ʒu
    Dim lngTop As Long                                              ' �c�ʒu
    ' �񌋍��̃Z���͈͂�I�����Ă��鎞�͏������Ȃ�
    If objRange.count > 1 Then
        ' �P�ꌋ���Z����OK �Ƃ���
        If objRange.Address <> objRange.Cells(1).MergeArea.Address Then Exit Sub
    End If
    '==============================================================================================-
    ' ���[�U�[�t�H�[���\���ʒu�擾
    Call FP_GetFormPosition(objRange, UF_Calendar5R.Width, UF_Calendar5R.Height, lngLeft, lngTop)
    '==============================================================================================-
    ' �J�����_�[�t�H�[��
    With UF_Calendar5R
        .prpTitle = strCaption                              ' �^�C�g��
        .prpEntMode = 0                                     ' ���̓��[�h(0=�Z���A1=TextBox)
        Set .prpRange = objRange                            ' �ΏۃZ��
        ' �t�H�[���\���ʒu�̊m�F
        If ((lngLeft <> 0) Or (lngTop <> 0)) Then
            ' �w�肪����ꍇ�̓}�j���A���w��
            .StartUpPosition = 0
            .Left = lngLeft
            .top = lngTop
        Else
            ' �w�肪�Ȃ��ꍇ�̓X�N���[���̒���
            .StartUpPosition = 2
        End If
        ' �J�����_�[�t�H�[����\��
        .Show
    End With
End Sub

'***************************************************************************************************
'�@������ �T�u���� ������
'***************************************************************************************************
'* �������@�FFP_GetFormPosition
'* �@�\�@�@�F���[�U�[�t�H�[���\���ʒu�擾
'===================================================================================================
'* �Ԃ�l�@�F��������(Boolean)
'* �����@�@�FArg1 = �ΏۃZ��(Object)
'* �@�@�@�@�@Arg2 = ���[�U�[�t�H�[���̕�(Long)
'* �@�@�@�@�@Arg3 = ���[�U�[�t�H�[���̍���(Long)
'* �@�@�@�@�@Arg4 = �X�N���[����̉��ʒu(Long)          ��Ref�Q��
'* �@�@�@�@�@Arg5 = �X�N���[����̏c�ʒu(Long)          ��Ref�Q��
'===================================================================================================
'* �쐬���@�F2018�N02��21��
'* �쐬�ҁ@�F���@��
'* �X�V���@�F2019�N12��08��
'* �X�V�ҁ@�F���@��
'* �@�\�����F�Z���̐^���Ƀt�H�[����\��������ʒu���擾
'* ���ӎ����F�擾�ł��Ȃ����͉��ʒu��c�ʒu�Ƃ��[���ƂȂ�
'***************************************************************************************************
Private Function FP_GetFormPosition(ByRef objRange As Range, _
                                    ByVal lngFormWidth As Long, _
                                    ByVal lngFormHeight As Long, _
                                    ByRef lngFormLeft As Long, _
                                    ByRef lngFormTop As Long) As Boolean
    '==============================================================================================-
    Dim objTarget As Range                                          ' �ΏۃZ��(�擪�Z��)
    Dim objAW As Window                                             ' ActiveWindow
    Dim lngPaneIx As Long                                           ' PaneIndex(0�`4)
    Dim lngIx As Long                                               ' ���[�v�pINDEX(Work)
    Dim lngR1C1Left As Long                                         ' �N�_�Z�����[�ʒu
    Dim lngR1C1Top As Long                                          ' �N�_�Z����[�ʒu
    Dim lngTargetLeft As Long                                       ' �ΏۃZ�����[�ʒu
    Dim lngTargetTop As Long                                        ' �ΏۃZ����[�ʒu
    Dim lngScreenRight As Long                                      ' �X�N���[���E�[�ʒu
    Dim lngScreenBottom As Long                                     ' �X�N���[�����[�ʒu
    Dim lngDPIX As Long                                             ' Dots Per Inch(����)
    Dim lngDPIY As Long                                             ' Dots Per Inch(����)
    Dim lngPPI As Long                                              ' Pixels Per Inch
    FP_GetFormPosition = False
    lngFormLeft = 0
    lngFormTop = 0
    lngPaneIx = 0
    Set objTarget = objRange.Cells(1).MergeArea
    Set objAW = ActiveWindow
    '==============================================================================================-
    ' �E�B���h�E����������
    If Not objAW.FreezePanes And Not objAW.Split Then
        ' �\����O�͖���
        If Intersect(objAW.VisibleRange, objTarget) Is Nothing Then Exit Function
    Else            ' ��������
        ' �E�B���h�E�g�Œ肩
        If objAW.FreezePanes Then
            ' �ǂ̃E�B���h�E�ɑ����邩����
            For lngIx = 1 To objAW.Panes.count
                ' �����H
                If Not Intersect(objAW.Panes(lngIx).VisibleRange, objTarget) Is Nothing Then
                    lngPaneIx = lngIx
                    Exit For
                End If
            Next lngIx
            ' ������Ȃ���
            If lngPaneIx = 0 Then Exit Function
        Else
            ' �E�B���h�E�����̓A�N�e�B�u�y�C���̂ݔ���
            If Not Intersect(objAW.ActivePane.VisibleRange, objTarget) Is Nothing Then
                lngPaneIx = objAW.ActivePane.Index
            Else
                Exit Function
            End If
        End If
    End If
    '==============================================================================================-
    ' ���ȉ���Excel2003�ȑO�ł͓��삵�Ȃ�
    lngDPIX = FP_GetDPIX
    lngDPIY = FP_GetDPIY
    lngPPI = FP_GetPPI
    ' �E�B���h�E����������
    If lngPaneIx = 0 Then
        lngR1C1Left = objAW.PointsToScreenPixelsX(0)
        lngR1C1Top = objAW.PointsToScreenPixelsY(0)
    Else
        lngR1C1Left = objAW.Panes(lngPaneIx).PointsToScreenPixelsX(0)
        lngR1C1Top = objAW.Panes(lngPaneIx).PointsToScreenPixelsY(0)
    End If
    lngTargetLeft = ((objTarget.Left * (lngDPIX / lngPPI)) * (objAW.Zoom / 100)) + lngR1C1Left
    lngTargetTop = (((objTarget.top + objTarget.Height) * (lngDPIY / lngPPI)) * (objAW.Zoom / 100)) + lngR1C1Top
    lngFormLeft = lngTargetLeft * (lngPPI / lngDPIX)
    lngFormTop = lngTargetTop * (lngPPI / lngDPIY)
    '==============================================================================================-
    ' �X�N���[���T�C�Y�ʒu�̎擾
    Call GP_GetScreenPos(0, 0, lngScreenRight, lngScreenBottom)
    '==============================================================================================-
    ' ���[�U�[�t�H�[�����X�N���[������͂ݏo����(��)
    If (lngFormLeft + lngFormWidth) * (lngDPIX / lngPPI) > lngScreenRight Then
        ' �X�N���[���E�[�Ɉړ�(+3�͌덷�H)
        lngFormLeft = lngScreenRight * (lngPPI / lngDPIX) - lngFormWidth + 3
    End If
    ' ���[�U�[�t�H�[�����X�N���[������͂ݏo����(�c)
    If (lngFormTop + lngFormHeight) * (lngDPIY / lngPPI) > lngScreenBottom Then
        ' �Z����[�Ɉړ�
        lngFormTop = lngFormTop - (objRange.Height + lngFormHeight)
    End If
    FP_GetFormPosition = True
End Function

'***************************************************************************************************
'* �������@�FFP_GetPPI
'* �@�\�@�@�FPPI(Pixels Per Inch)�擾
'===================================================================================================
'* �Ԃ�l�@�FPPI�l(Long)
'* �����@�@�F(�Ȃ�)
'===================================================================================================
'* �쐬���@�F2018�N02��21��
'* �쐬�ҁ@�F���@��
'* �X�V���@�F2018�N02��21��
'* �X�V�ҁ@�F���@��
'* �@�\�����F
'* ���ӎ����F
'***************************************************************************************************
Private Function FP_GetPPI() As Long
    '==============================================================================================-
    FP_GetPPI = Application.InchesToPoints(1)
End Function

'***************************************************************************************************
'* �������@�FFP_GetDPIX
'* �@�\�@�@�FDPI(Dots Per Inch)�擾(��������)
'===================================================================================================
'* �Ԃ�l�@�FDPI�l(Long)
'* �����@�@�F(�Ȃ�)
'===================================================================================================
'* �쐬���@�F2018�N02��21��
'* �쐬�ҁ@�F���@��
'* �X�V���@�F2018�N02��21��
'* �X�V�ҁ@�F���@��
'* �@�\�����F
'* ���ӎ����F
'***************************************************************************************************
Private Function FP_GetDPIX() As Long
    '==============================================================================================-
    FP_GetDPIX = FP_GetDPI(LOGPIXELSX)
End Function

'***************************************************************************************************
'* �������@�FFP_GetDPIY
'* �@�\�@�@�FDPI(Dots Per Inch)�擾(��������)
'===================================================================================================
'* �Ԃ�l�@�FDPI�l(Long)
'* �����@�@�F(�Ȃ�)
'===================================================================================================
'* �쐬���@�F2018�N02��21��
'* �쐬�ҁ@�F���@��
'* �X�V���@�F2018�N02��21��
'* �X�V�ҁ@�F���@��
'* �@�\�����F
'* ���ӎ����F
'***************************************************************************************************
Private Function FP_GetDPIY() As Long
    '==============================================================================================-
    FP_GetDPIY = FP_GetDPI(LOGPIXELSY)
End Function

'***************************************************************************************************
'* �������@�FFP_GetDPI
'* �@�\�@�@�FDPI(Dots Per Inch)�擾(API)
'===================================================================================================
'* �Ԃ�l�@�FDPI�l(Long)
'* �����@�@�FArg1 = nFlag(Long)
'===================================================================================================
'* �쐬���@�F2018�N02��21��
'* �쐬�ҁ@�F���@��
'* �X�V���@�F2019�N10��20��
'* �X�V�ҁ@�F���@��
'* �@�\�����F
'* ���ӎ����F
'***************************************************************************************************
Private Function FP_GetDPI(ByVal nFlag As Long) As Long
    '==============================================================================================-
#If VBA7 Then
    Dim lngHdc As LongPtr                                           ' �E�B���h�E�n���h����DC
#Else
    Dim lngHdc As Long                                              ' �E�B���h�E�n���h����DC
#End If
    lngHdc = GetDC(Application.hWnd)
    FP_GetDPI = GetDeviceCaps(lngHdc, nFlag)
    Call ReleaseDC(&H0, lngHdc)
End Function

'***************************************************************************************************
'* �������@�FGP_GetScreenPos
'* �@�\�@�@�F�X�N���[���ʒu�̎擾
'===================================================================================================
'* �Ԃ�l�@�F(�Ȃ�)
'* �����@�@�FArg1 = �X�N���[�����[�ʒu(Long)              ��Ref�Q��
'* �@�@�@�@�@Arg2 = �X�N���[����[�ʒu(Long)              ��Ref�Q��
'* �@�@�@�@�@Arg3 = �X�N���[���E�[�ʒu(Long)              ��Ref�Q��
'* �@�@�@�@�@Arg4 = �X�N���[�����[�ʒu(Long)              ��Ref�Q��
'===================================================================================================
'* �쐬���@�F2019�N12��08��
'* �쐬�ҁ@�F���@��
'* �X�V���@�F2019�N12��08��
'* �X�V�ҁ@�F���@��
'* �@�\�����F�����X�N���[���S�̎l���̈ʒu���擾
'* ���ӎ����F
'***************************************************************************************************
Private Sub GP_GetScreenPos(ByRef lngScreenLeft As Long, _
                            ByRef lngScreenTop As Long, _
                            ByRef lngScreenRight As Long, _
                            ByRef lngScreenBottom As Long)
    '==============================================================================================-
    Dim lngWidth As Long                                            ' �X�N���[���̕�
    Dim lngHeight As Long                                           ' �X�N���[���̍����@
    Dim lngHeight2 As Long                                          ' �X�N���[���̍����A
    Dim lngHeight3 As Long                                          ' �X�N���[���̍����B
    Dim objRect As g_typRect                                        ' Rect
    ' �X�N���[���̍��[���[���������̎擾(�����X�N���[���Ή�)
    lngScreenLeft = getSystemMetrics(SM_XVIRTUALSCREEN)         ' ���[
    lngScreenTop = getSystemMetrics(SM_YVIRTUALSCREEN)          ' ��[
    lngWidth = getSystemMetrics(SM_CXVIRTUALSCREEN)             ' ��(���z�X�N���[����)
    lngHeight = getSystemMetrics(SM_CYVIRTUALSCREEN)            ' ����(���z�X�N���[����)
    lngHeight2 = getSystemMetrics(SM_CYSCREEN)                  ' ����(���C���X�N���[���̂�)
    ' �^�X�N�o�[�������X�N���[���̑傫���擾(���C���X�N���[���̂�)
    Call SystemParametersInfo(SPI_GETWORKAREA, 0, objRect, 0)
    lngHeight3 = objRect.Bottom - objRect.top                   ' ����(���C���̃^�X�N�o�[�ȊO�̕�)
    ' �^�X�N�o�[�����C���X�N���[���̉��[�ɂ�����̂Ƃ��A���̍�������������
    lngHeight = lngHeight - (lngHeight2 - lngHeight3)
    ' �E�[�̎Z�o
    lngScreenRight = lngWidth - lngScreenLeft
    ' ���[�̎Z�o
    lngScreenBottom = lngHeight - lngScreenTop
End Sub

'========================================<< End of Source >>========================================
