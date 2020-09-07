VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} OptionForm 
   Caption         =   "�I�v�V���� - WBS"
   ClientHeight    =   8745
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11970
   OleObjectBlob   =   "optionForm.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "optionForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************************************
'   �J�����_�[�t�H�[��4(���t���͕��i)   ���e�X�g�p�t�H�[��      UserForm1(UserForm)
'
'   �쐬��:��㎡  URL:http://www.ne.jp/asahi/excel/inoue/ [Excel�ł��d��!]
'***************************************************************************************************
'�ύX���t Rev  �ύX������e------------------------------------------------------------------------>
'18/02/21(1.00)�V�K�쐬
'18/11/28(1.80)�J�����_�[�t�H�[�����A���W���[�����ύX
'***************************************************************************************************
Option Explicit
'---------------------------------------------------------------------------------------------------
Private Const g_cnsAddLeft As Long = 3                          ' Left�����l
Private Const g_cnsAddTop As Long = 19                          ' Top�����l
Private Const g_cnsAddLeft2 As Long = 4                         ' Left�����l(�t���[���p)
Private Const g_cnsAddTop2 As Long = 25                         ' Top�����l(�t���[���p)
' �������̒����l��Windows10���_�̉�ʂœK���Ɍ��U�����l�ł��B
' �@��d�Ƀt���[�����d�Ȃ������̏ꍇ�͕ʓr�������K�v�ł��B


#If Win64 Then
  Private Declare PtrSafe Function GetAsyncKeyState Lib "User32.dll" (ByVal vKey As Long) As Integer
  Private Declare PtrSafe Function GetForegroundWindow Lib "user32" () As LongPtr
  Private Declare PtrSafe Function SetWindowPos Lib "user32" (ByVal hWnd As LongPtr, ByVal hWndInsertAfter As LongPtr, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
#Else
  Private Declare Function GetAsyncKeyState Lib "User32.dll" (ByVal vKey As Long) As Integer
  Private Declare Function GetForegroundWindow Lib "user32" () As Long
  Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
#End If


Private Const HWND_TOPMOST As Long = -1
Private Const SWP_NOSIZE As Long = &H1&
Private Const SWP_NOMOVE As Long = &H2&

Public KeyPressFlg As Boolean




'Private Sub UserForm_Activate()
'    Call SetWindowPos(GetForegroundWindow, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
'    Me.StartUpPosition = 1
'End Sub

'***************************************************************************************************
'* �������@�FUserForm_Initialize
'* �@�\�@�@�F���[�U�[�t�H�[���̏�����
'---------------------------------------------------------------------------------------------------
'* �Ԃ�l�@�F(�Ȃ�)
'* �����@�@�F(�Ȃ�)
'---------------------------------------------------------------------------------------------------
'* �쐬���@�F2018�N02��21��
'* �쐬�ҁ@�F���@��
'* �X�V���@�F2018�N02��21��
'* �X�V�ҁ@�F���@��
'* �@�\�����F
'* ���ӎ����F
'***************************************************************************************************
Private Sub UserForm_Initialize()
    ' �e�L�X�g�{�b�N�X�Ɂ��{�^����\��������
    startDay.ShowDropButtonWhen = fmShowDropButtonWhenAlways
    endDay.ShowDropButtonWhen = fmShowDropButtonWhenAlways
    baseDay.ShowDropButtonWhen = fmShowDropButtonWhenAlways
End Sub


'***************************************************************************************************
' ������ �t�H�[���C�x���g ������
'***************************************************************************************************
'* �������@�FstartDay_DropButtonClick
'* �@�\�@�@�F�t�H�[����̃e�L�X�g�{�b�N�X�C�x���g(DropButtonClick)
'---------------------------------------------------------------------------------------------------
'* �Ԃ�l�@�F(�Ȃ�)
'* �����@�@�F(�Ȃ�)
'---------------------------------------------------------------------------------------------------
'* �쐬���@�F2018�N02��21��
'* �쐬�ҁ@�F���@��
'* �X�V���@�F2018�N11��28��
'* �X�V�ҁ@�F���@��
'* �@�\�����F
'* ���ӎ����F
'***************************************************************************************************
Private Sub startDay_DropButtonClick()
    '-----------------------------------------------------------------------------------------------
    Dim lngLeft As Long                                             ' �ʒu(������)
    Dim lngTop As Long                                              ' �ʒu(�c����)
    ' �t�H�[��+�e�L�X�g�{�b�N�X��Left,Top�l����ʒu�𔻒�
    lngLeft = Me.Left + startDay.Left + g_cnsAddLeft
    lngTop = Me.top + startDay.top + startDay.Height + g_cnsAddTop
    '-----------------------------------------------------------------------------------------------
    ' �J�����_�[�t�H�[�����N������
    Call modCalendar5R.ShowCalendarFromTextBox2(startDay, lngLeft, lngTop, "�J�n���I��")
End Sub

Private Sub endDay_DropButtonClick()
    '-----------------------------------------------------------------------------------------------
    Dim lngLeft As Long                                             ' �ʒu(������)
    Dim lngTop As Long                                              ' �ʒu(�c����)
    ' �t�H�[��+�e�L�X�g�{�b�N�X��Left,Top�l����ʒu�𔻒�
    lngLeft = Me.Left + endDay.Left + g_cnsAddLeft
    lngTop = Me.top + endDay.top + endDay.Height + g_cnsAddTop
    '-----------------------------------------------------------------------------------------------
    ' �J�����_�[�t�H�[�����N������
    Call modCalendar5R.ShowCalendarFromTextBox2(endDay, lngLeft, lngTop, "�I�����I��")
End Sub

Private Sub baseDay_DropButtonClick()
    '-----------------------------------------------------------------------------------------------
    Dim lngLeft As Long                                             ' �ʒu(������)
    Dim lngTop As Long                                              ' �ʒu(�c����)
    ' �t�H�[��+�e�L�X�g�{�b�N�X��Left,Top�l����ʒu�𔻒�
    lngLeft = Me.Left + baseDay.Left + g_cnsAddLeft
    lngTop = Me.top + baseDay.top + baseDay.Height + g_cnsAddTop
    '-----------------------------------------------------------------------------------------------
    ' �J�����_�[�t�H�[�����N������
    Call modCalendar5R.ShowCalendarFromTextBox2(baseDay, lngLeft, lngTop, "����I��")
End Sub


'***************************************************************************************************
'* �������@�FGP_GakuEnter
'* �@�\�@�@�F���z���ړ��͗p�ҏW
'---------------------------------------------------------------------------------------------------
'* �Ԃ�l�@�F(�Ȃ�)
'* �����@�@�FArg1 = �e�L�X�g�{�b�N�X(Object)
'---------------------------------------------------------------------------------------------------
'* �쐬���@�F2003�N07��25��
'* �쐬�ҁ@�F���@��
'* �X�V���@�F2020�N02��24��
'* �X�V�ҁ@�F���@��
'* �@�\�����F
'* ���ӎ����F
'***************************************************************************************************
Private Sub GP_GakuEnter(objTextBox As MSForms.TextBox)
    '-----------------------------------------------------------------------------------------------
    Dim strGaku As String                                           ' ���̓e�L�X�g
    Dim crnGaku As Currency                                         ' ���z�l
    strGaku = Trim(objTextBox.Text)
    ' ���l��
    If IsNumeric(strGaku) Then
        crnGaku = CCur(strGaku)
        ' 3���J���}�����ŕҏW
        objTextBox.Text = Format(crnGaku, "0")
        ' �S���I��
        Call GP_AllSelect(objTextBox)
    End If
End Sub
'***************************************************************************************************
'* �������@�FFP_GakuExit
'* �@�\�@�@�F���z���ڕ\���p�ҏW
'---------------------------------------------------------------------------------------------------
'* �Ԃ�l�@�F�`�F�b�N����(Boolean)
'* �����@�@�FArg1 = �e�L�X�g�{�b�N�X(Object)
'---------------------------------------------------------------------------------------------------
'* �쐬���@�F2003�N07��25��
'* �쐬�ҁ@�F���@��
'* �X�V���@�F2020�N02��24��
'* �X�V�ҁ@�F���@��
'* �@�\�����F
'* ���ӎ����F
'***************************************************************************************************
Private Function FP_GakuExit(objTextBox As MSForms.TextBox) As Boolean
    '-----------------------------------------------------------------------------------------------
    Dim strGaku As String                                           ' ���̓e�L�X�g
    Dim crnGaku As Currency                                         ' ���z�l
    FP_GakuExit = False
    strGaku = Trim(objTextBox.Text)
    ' ���l��
    If IsNumeric(strGaku) Then
        crnGaku = CCur(strGaku)
        ' 3���J���}�t���ŕҏW
        objTextBox.Text = Format(crnGaku, "#,##0")
        FP_GakuExit = True
    ElseIf strGaku = "" Then
    Else
        MsgBox "�����ł͂���܂���B", vbExclamation
        ' �S���I��
        Call GP_AllSelect(objTextBox)
    End If
End Function
'***************************************************************************************************
'* �������@�FGP_AllSelect
'* �@�\�@�@�F�S���I��
'---------------------------------------------------------------------------------------------------
'* �Ԃ�l�@�F(�Ȃ�)
'* �����@�@�FArg1 = �e�L�X�g�{�b�N�X(Object)
'---------------------------------------------------------------------------------------------------
'* �쐬���@�F2003�N07��25��
'* �쐬�ҁ@�F���@��
'* �X�V���@�F2020�N02��24��
'* �X�V�ҁ@�F���@��
'* �@�\�����F
'* ���ӎ����F
'***************************************************************************************************
Private Sub GP_AllSelect(objTextBox As MSForms.TextBox)
    '-----------------------------------------------------------------------------------------------
    With objTextBox
      .SetFocus
      .SelStart = 0
      .SelLength = Len(.Text)
    End With
End Sub

Function chkScope(minVal As MSForms.TextBox, maxVal As MSForms.TextBox)

  If minVal.Text <= maxVal.Text Then
    chkScope = True
  ElseIf maxVal.Text = 0 Then
    chkScope = True
  Else
    message.Caption = "�͈͐ݒ肪�������܂���"
    chkScope = False
  End If
  
End Function

'**************************************************************************************************
' * �S���ҏ��̃N���A
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Private Sub clearAssignor_Click()
    Assign01.Text = ""
    Assign02.Text = ""
    Assign03.Text = ""
    Assign04.Text = ""
    Assign05.Text = ""
    Assign06.Text = ""
    Assign07.Text = ""
    Assign08.Text = ""
    Assign09.Text = ""
    Assign10.Text = ""
    Assign11.Text = ""
    Assign12.Text = ""
    Assign13.Text = ""
    Assign14.Text = ""
    Assign15.Text = ""
    Assign16.Text = ""
    Assign17.Text = ""
    Assign18.Text = ""
    Assign19.Text = ""
    Assign20.Text = ""
    Assign21.Text = ""
    Assign22.Text = ""
    Assign23.Text = ""
    Assign24.Text = ""
'    Assign25.Text = ""
'    Assign26.Text = ""
'    Assign27.Text = ""
'    Assign28.Text = ""
'    Assign29.Text = ""
'    Assign30.Text = ""
'    Assign31.Text = ""
'    Assign32.Text = ""
'    Assign33.Text = ""
'    Assign34.Text = ""
'    Assign35.Text = ""
    
    
    AssignColor01.BackColor = 16777215
    AssignColor02.BackColor = 16777215
    AssignColor03.BackColor = 16777215
    AssignColor04.BackColor = 16777215
    AssignColor05.BackColor = 16777215
    AssignColor06.BackColor = 16777215
    AssignColor07.BackColor = 16777215
    AssignColor08.BackColor = 16777215
    AssignColor09.BackColor = 16777215
    AssignColor10.BackColor = 16777215
    AssignColor11.BackColor = 16777215
    AssignColor12.BackColor = 16777215
    AssignColor13.BackColor = 16777215
    AssignColor14.BackColor = 16777215
    AssignColor15.BackColor = 16777215
    AssignColor16.BackColor = 16777215
    AssignColor17.BackColor = 16777215
    AssignColor18.BackColor = 16777215
    AssignColor19.BackColor = 16777215
    AssignColor20.BackColor = 16777215
    AssignColor21.BackColor = 16777215
    AssignColor22.BackColor = 16777215
    AssignColor23.BackColor = 16777215
    AssignColor24.BackColor = 16777215
'    AssignColor25.BackColor = 16777215
'    AssignColor26.BackColor = 16777215
'    AssignColor27.BackColor = 16777215
'    AssignColor28.BackColor = 16777215
'    AssignColor29.BackColor = 16777215
'    AssignColor30.BackColor = 16777215
'    AssignColor31.BackColor = 16777215
'    AssignColor32.BackColor = 16777215
'    AssignColor33.BackColor = 16777215
'    AssignColor34.BackColor = 16777215
'    AssignColor35.BackColor = 16777215



End Sub


'**************************************************************************************************
' * �������s
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Private Sub run_Click()
  Set getVal = New Collection
  
  With getVal
    .Add item:=startDay.Text, Key:="startDay"
    .Add item:=endDay.Text, Key:="endDay"
    .Add item:=baseDay.Text, Key:="baseDay"
    .Add item:=setLightning.Value, Key:="setLightning"
    .Add item:=setDispProgress100.Value, Key:="setDispProgress100"
    .Add item:=CompanyHoliday.Value, Key:="CompanyHoliday"
    
    '�X�^�C���֘A
    .Add item:=lineColor.BackColor, Key:="lineColor"
    
    .Add item:=SaturdayColor.BackColor, Key:="SaturdayColor"
    .Add item:=SundayColor.BackColor, Key:="SundayColor"
    .Add item:=CompanyHolidayColor.BackColor, Key:="CompanyHolidayColor"

    .Add item:=lineColor_Plan.BackColor, Key:="lineColor_Plan"
    .Add item:=lineColor_Achievement.BackColor, Key:="lineColor_Achievement"
    .Add item:=lineColor_Lightning.BackColor, Key:="lineColor_Lightning"
    .Add item:=lineColor_TaskLevel1.BackColor, Key:="lineColor_TaskLevel1"
    .Add item:=lineColor_TaskLevel2.BackColor, Key:="lineColor_TaskLevel2"
    .Add item:=lineColor_TaskLevel3.BackColor, Key:="lineColor_TaskLevel3"

    '�S����
    .Add item:=Assign01.Text, Key:="Assign01"
    .Add item:=Assign02.Text, Key:="Assign02"
    .Add item:=Assign03.Text, Key:="Assign03"
    .Add item:=Assign04.Text, Key:="Assign04"
    .Add item:=Assign05.Text, Key:="Assign05"
    .Add item:=Assign06.Text, Key:="Assign06"
    .Add item:=Assign07.Text, Key:="Assign07"
    .Add item:=Assign08.Text, Key:="Assign08"
    .Add item:=Assign09.Text, Key:="Assign09"
    .Add item:=Assign10.Text, Key:="Assign10"
    .Add item:=Assign11.Text, Key:="Assign11"
    .Add item:=Assign12.Text, Key:="Assign12"
    .Add item:=Assign13.Text, Key:="Assign13"
    .Add item:=Assign14.Text, Key:="Assign14"
    .Add item:=Assign15.Text, Key:="Assign15"
    .Add item:=Assign16.Text, Key:="Assign16"
    .Add item:=Assign17.Text, Key:="Assign17"
    .Add item:=Assign18.Text, Key:="Assign18"
    .Add item:=Assign19.Text, Key:="Assign19"
    .Add item:=Assign20.Text, Key:="Assign20"
    .Add item:=Assign21.Text, Key:="Assign21"
    .Add item:=Assign22.Text, Key:="Assign22"
    .Add item:=Assign23.Text, Key:="Assign23"
    .Add item:=Assign24.Text, Key:="Assign24"
    .Add item:=Assign25.Text, Key:="Assign25"
    .Add item:=Assign26.Text, Key:="Assign26"
    .Add item:=Assign27.Text, Key:="Assign27"
    .Add item:=Assign28.Text, Key:="Assign28"
    .Add item:=Assign29.Text, Key:="Assign29"
    .Add item:=Assign30.Text, Key:="Assign30"
    .Add item:=Assign31.Text, Key:="Assign31"
    .Add item:=Assign32.Text, Key:="Assign32"
    .Add item:=Assign33.Text, Key:="Assign33"
    .Add item:=Assign34.Text, Key:="Assign34"
    .Add item:=Assign35.Text, Key:="Assign35"

    '�S����Color
    .Add item:=AssignColor01.BackColor, Key:="AssignColor01"
    .Add item:=AssignColor02.BackColor, Key:="AssignColor02"
    .Add item:=AssignColor03.BackColor, Key:="AssignColor03"
    .Add item:=AssignColor04.BackColor, Key:="AssignColor04"
    .Add item:=AssignColor05.BackColor, Key:="AssignColor05"
    .Add item:=AssignColor06.BackColor, Key:="AssignColor06"
    .Add item:=AssignColor07.BackColor, Key:="AssignColor07"
    .Add item:=AssignColor08.BackColor, Key:="AssignColor08"
    .Add item:=AssignColor09.BackColor, Key:="AssignColor09"
    .Add item:=AssignColor10.BackColor, Key:="AssignColor10"
    .Add item:=AssignColor11.BackColor, Key:="AssignColor11"
    .Add item:=AssignColor12.BackColor, Key:="AssignColor12"
    .Add item:=AssignColor13.BackColor, Key:="AssignColor13"
    .Add item:=AssignColor14.BackColor, Key:="AssignColor14"
    .Add item:=AssignColor15.BackColor, Key:="AssignColor15"
    .Add item:=AssignColor16.BackColor, Key:="AssignColor16"
    .Add item:=AssignColor17.BackColor, Key:="AssignColor17"
    .Add item:=AssignColor18.BackColor, Key:="AssignColor18"
    .Add item:=AssignColor19.BackColor, Key:="AssignColor19"
    .Add item:=AssignColor20.BackColor, Key:="AssignColor20"
    .Add item:=AssignColor21.BackColor, Key:="AssignColor21"
    .Add item:=AssignColor22.BackColor, Key:="AssignColor22"
    .Add item:=AssignColor23.BackColor, Key:="AssignColor23"
    .Add item:=AssignColor24.BackColor, Key:="AssignColor24"
    .Add item:=AssignColor25.BackColor, Key:="AssignColor25"
    .Add item:=AssignColor26.BackColor, Key:="AssignColor26"
    .Add item:=AssignColor27.BackColor, Key:="AssignColor27"
    .Add item:=AssignColor28.BackColor, Key:="AssignColor28"
    .Add item:=AssignColor29.BackColor, Key:="AssignColor29"
    .Add item:=AssignColor30.BackColor, Key:="AssignColor30"
    .Add item:=AssignColor31.BackColor, Key:="AssignColor31"
    .Add item:=AssignColor32.BackColor, Key:="AssignColor32"
    .Add item:=AssignColor33.BackColor, Key:="AssignColor33"
    .Add item:=AssignColor34.BackColor, Key:="AssignColor34"
    .Add item:=AssignColor35.BackColor, Key:="AssignColor35"
    
    
    '�\���ݒ�
    .Add item:=view_Plan.Value, Key:="view_Plan"
    .Add item:=view_Assign.Value, Key:="view_Assign"
    .Add item:=view_Progress.Value, Key:="view_Progress"
    .Add item:=view_Achievement.Value, Key:="view_Achievement"
    .Add item:=view_Task.Value, Key:="view_Task"
    .Add item:=view_TaskInfo.Value, Key:="view_TaskInfo"
    .Add item:=View_TaskAllocation.Value, Key:="view_TaskAllocation"
    
    
    .Add item:=view_WorkLoad.Value, Key:="view_WorkLoad"
    .Add item:=view_LateOrEarly.Value, Key:="view_LateOrEarly"
    .Add item:=view_Note.Value, Key:="view_Note"
    .Add item:=viewGant_TaskName.Value, Key:="viewGant_TaskName"
    .Add item:=viewGant_Assignor.Value, Key:="viewGant_Assignor"

  End With
  Unload Me
  
  
  Call WBS_Option.�I�v�V�����ݒ�l�i�[

End Sub


'**************************************************************************************************
' * xxxxxxxxxx
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Private Sub Cancel_Click()
  Unload Me
End Sub

'**************************************************************************************************
' * �X�^�C���֘A
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************
Private Sub lineColor_Click()
  Dim colorValue As Long
  colorValue = Library.getColor(Me.lineColor.BackColor)
  Me.lineColor.BackColor = colorValue
End Sub
Private Sub SaturdayColor_Click()
  Dim colorValue As Long
    colorValue = Library.getColor(Me.SaturdayColor.BackColor)
  Me.SaturdayColor.BackColor = colorValue
End Sub
Private Sub SundayColor_Click()
  Dim colorValue As Long
    colorValue = Library.getColor(Me.SundayColor.BackColor)
  Me.SundayColor.BackColor = colorValue
End Sub
Private Sub CompanyHolidayColor_Click()
  Dim colorValue As Long
    colorValue = Library.getColor(Me.CompanyHolidayColor.BackColor)
  Me.CompanyHolidayColor.BackColor = colorValue
End Sub
Private Sub lineColor_Plan_Click()
  Dim colorValue As Long
    colorValue = Library.getColor(Me.lineColor_Plan.BackColor)
  Me.lineColor_Plan.BackColor = colorValue
End Sub
Private Sub lineColor_Achievement_Click()
  Dim colorValue As Long
    colorValue = Library.getColor(Me.lineColor_Achievement.BackColor)
  Me.lineColor_Achievement.BackColor = colorValue
End Sub
Private Sub lineColor_Lightning_Click()
  Dim colorValue As Long
    colorValue = Library.getColor(Me.lineColor_Lightning.BackColor)
  Me.lineColor_Lightning.BackColor = colorValue
End Sub
Private Sub lineColor_TaskLevel1_Click()
  Dim colorValue As Long
    colorValue = Library.getColor(Me.lineColor_TaskLevel1.BackColor)
  Me.lineColor_TaskLevel1.BackColor = colorValue
End Sub
Private Sub lineColor_TaskLevel2_Click()
  Dim colorValue As Long
    colorValue = Library.getColor(Me.lineColor_TaskLevel2.BackColor)
  Me.lineColor_TaskLevel2.BackColor = colorValue
End Sub
Private Sub lineColor_TaskLevel3_Click()
  Dim colorValue As Long
    colorValue = Library.getColor(Me.lineColor_TaskLevel3.BackColor)
  Me.lineColor_TaskLevel3.BackColor = colorValue
End Sub


'**************************************************************************************************
' * �S����
' *
' * @author Bunpei.Koizumi<bunpei.koizumi@gmail.com>
'**************************************************************************************************

Private Sub AssignColor01_Click()
  Dim colorValue As Long
  colorValue = Library.getColor(Me.AssignColor01.BackColor)
  Me.AssignColor01.BackColor = colorValue
End Sub

Private Sub AssignColor02_Click()
  Dim colorValue As Long
  colorValue = Library.getColor(Me.AssignColor02.BackColor)
  Me.AssignColor02.BackColor = colorValue
End Sub

Private Sub AssignColor03_Click()
  Dim colorValue As Long
  colorValue = Library.getColor(Me.AssignColor03.BackColor)
  Me.AssignColor03.BackColor = colorValue
End Sub

Private Sub AssignColor04_Click()
  Dim colorValue As Long
  colorValue = Library.getColor(Me.AssignColor04.BackColor)
  Me.AssignColor04.BackColor = colorValue
End Sub

Private Sub AssignColor05_Click()
  Dim colorValue As Long
  colorValue = Library.getColor(Me.AssignColor05.BackColor)
  Me.AssignColor05.BackColor = colorValue
End Sub

Private Sub AssignColor06_Click()
  Dim colorValue As Long
  colorValue = Library.getColor(Me.AssignColor06.BackColor)
  Me.AssignColor06.BackColor = colorValue
End Sub

Private Sub AssignColor07_Click()
  Dim colorValue As Long
  colorValue = Library.getColor(Me.AssignColor07.BackColor)
  Me.AssignColor07.BackColor = colorValue
End Sub

Private Sub AssignColor08_Click()
  Dim colorValue As Long
  colorValue = Library.getColor(Me.AssignColor08.BackColor)
  Me.AssignColor08.BackColor = colorValue
End Sub

Private Sub AssignColor09_Click()
  Dim colorValue As Long
  colorValue = Library.getColor(Me.AssignColor09.BackColor)
  Me.AssignColor09.BackColor = colorValue
End Sub

Private Sub AssignColor10_Click()
  Dim colorValue As Long
  colorValue = Library.getColor(Me.AssignColor10.BackColor)
  Me.AssignColor10.BackColor = colorValue
End Sub

Private Sub AssignColor11_Click()
  Dim colorValue As Long
  colorValue = Library.getColor(Me.AssignColor11.BackColor)
  Me.AssignColor11.BackColor = colorValue
End Sub

Private Sub AssignColor12_Click()
  Dim colorValue As Long
  colorValue = Library.getColor(Me.AssignColor12.BackColor)
  Me.AssignColor12.BackColor = colorValue
End Sub

Private Sub AssignColor13_Click()
  Dim colorValue As Long
  colorValue = Library.getColor(Me.AssignColor13.BackColor)
  Me.AssignColor13.BackColor = colorValue
End Sub

Private Sub AssignColor14_Click()
  Dim colorValue As Long
  colorValue = Library.getColor(Me.AssignColor14.BackColor)
  Me.AssignColor14.BackColor = colorValue
End Sub

Private Sub AssignColor15_Click()
  Dim colorValue As Long
  colorValue = Library.getColor(Me.AssignColor15.BackColor)
  Me.AssignColor15.BackColor = colorValue
End Sub

Private Sub AssignColor16_Click()
  Dim colorValue As Long
  colorValue = Library.getColor(Me.AssignColor16.BackColor)
  Me.AssignColor16.BackColor = colorValue
End Sub

Private Sub AssignColor17_Click()
  Dim colorValue As Long
  colorValue = Library.getColor(Me.AssignColor17.BackColor)
  Me.AssignColor17.BackColor = colorValue
End Sub

Private Sub AssignColor18_Click()
  Dim colorValue As Long
  colorValue = Library.getColor(Me.AssignColor18.BackColor)
  Me.AssignColor18.BackColor = colorValue
End Sub

Private Sub AssignColor19_Click()
  Dim colorValue As Long
  colorValue = Library.getColor(Me.AssignColor19.BackColor)
  Me.AssignColor19.BackColor = colorValue
End Sub

Private Sub AssignColor20_Click()
  Dim colorValue As Long
  colorValue = Library.getColor(Me.AssignColor20.BackColor)
  Me.AssignColor20.BackColor = colorValue
End Sub

Private Sub AssignColor21_Click()
  Dim colorValue As Long
  colorValue = Library.getColor(Me.AssignColor21.BackColor)
  Me.AssignColor21.BackColor = colorValue
End Sub

Private Sub AssignColor22_Click()
  Dim colorValue As Long
  colorValue = Library.getColor(Me.AssignColor22.BackColor)
  Me.AssignColor22.BackColor = colorValue
End Sub

Private Sub AssignColor23_Click()
  Dim colorValue As Long
  colorValue = Library.getColor(Me.AssignColor23.BackColor)
  Me.AssignColor23.BackColor = colorValue
End Sub

Private Sub AssignColor24_Click()
  Dim colorValue As Long
  colorValue = Library.getColor(Me.AssignColor24.BackColor)
  Me.AssignColor24.BackColor = colorValue
End Sub

Private Sub AssignColor25_Click()
  Dim colorValue As Long
  colorValue = Library.getColor(Me.AssignColor25.BackColor)
  Me.AssignColor25.BackColor = colorValue
End Sub

Private Sub AssignColor26_Click()
  Dim colorValue As Long
  colorValue = Library.getColor(Me.AssignColor26.BackColor)
  Me.AssignColor26.BackColor = colorValue
End Sub

Private Sub AssignColor27_Click()
  Dim colorValue As Long
  colorValue = Library.getColor(Me.AssignColor27.BackColor)
  Me.AssignColor27.BackColor = colorValue
End Sub

Private Sub AssignColor28_Click()
  Dim colorValue As Long
  colorValue = Library.getColor(Me.AssignColor28.BackColor)
  Me.AssignColor28.BackColor = colorValue
End Sub

Private Sub AssignColor29_Click()
  Dim colorValue As Long
  colorValue = Library.getColor(Me.AssignColor29.BackColor)
  Me.AssignColor29.BackColor = colorValue
End Sub

Private Sub AssignColor30_Click()
  Dim colorValue As Long
  colorValue = Library.getColor(Me.AssignColor30.BackColor)
  Me.AssignColor30.BackColor = colorValue
End Sub

Private Sub AssignColor31_Click()
  Dim colorValue As Long
  colorValue = Library.getColor(Me.AssignColor31.BackColor)
  Me.AssignColor31.BackColor = colorValue
End Sub

Private Sub AssignColor32_Click()
  Dim colorValue As Long
  colorValue = Library.getColor(Me.AssignColor32.BackColor)
  Me.AssignColor32.BackColor = colorValue
End Sub

Private Sub AssignColor33_Click()
  Dim colorValue As Long
  colorValue = Library.getColor(Me.AssignColor33.BackColor)
  Me.AssignColor33.BackColor = colorValue
End Sub

Private Sub AssignColor34_Click()
  Dim colorValue As Long
  colorValue = Library.getColor(Me.AssignColor34.BackColor)
  Me.AssignColor34.BackColor = colorValue
End Sub

Private Sub AssignColor35_Click()
  Dim colorValue As Long
  colorValue = Library.getColor(Me.AssignColor35.BackColor)
  Me.AssignColor35.BackColor = colorValue
End Sub

