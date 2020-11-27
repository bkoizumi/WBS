Attribute VB_Name = "Hollyday"
'******************************************************************************
' ��������j���ł��邩�H���̏ꍇ�ǂ̏j�����H�𒲂ׂ�֐��B
' http://www.excelio.jp/LABORATORY/EXCEL_CALENDER.html
'******************************************************************************
Public Function GetHollyday(targetdate As Date, HollydayName As String) As Boolean
    kaerichi = False
    HollydayName = ""
    kaerichi = NationalHollydays(targetdate, HollydayName)
    If kaerichi = True Then
        GetHollyday = True
    Else
        HollydayName = ""
        kaerichi = FurikaeKyujitsu(targetdate, HollydayName)
        If kaerichi = True Then
            GetHollyday = True
        Else
            HollydayName = ""
            kaerichi = KokuminnoKyujitsu(targetdate, HollydayName)
            If kaerichi = True Then
                GetHollyday = True
            Else
                HollydayName = ""
                kaerichi = TokubetsunaKyujitsu(targetdate, HollydayName)
                If kaerichi = True Then
                    GetHollyday = True
                Else
                    GetHollyday = False
                End If
            End If
        End If
    End If
End Function
'******************************************************************************
' �j������֐�
'******************************************************************************
Public Function NationalHollydays(targetdate As Date, HollydayName As String) As Boolean
Dim targetyear As Integer
Dim targetmonth As Integer
Dim targetday As Integer
    targetyear = CInt(Format(targetdate, "yyyy"))
    targetmonth = CInt(Format(targetdate, "m"))
    targetday = CInt(Format(targetdate, "d"))
    hantei = False
    Select Case targetmonth
    Case 1
        If targetyear > 1948 And targetday = 1 Then
                hantei = True
                HollydayName = "���U"
        End If
        If targetyear > 1948 Then
            If targetyear < 2000 Then
                If targetday = 15 Then
                    hantei = True
                    HollydayName = "���l�̓�"
                End If
            ElseIf CInt(Format(DaiXYoubi(targetyear, 1, 2, 1), "d")) = targetday Then
                    hantei = True
                    HollydayName = "���l�̓�"
            End If
        End If
    Case 2
        If targetyear > 1966 Then
                If targetday = 11 Then
                    hantei = True
                    HollydayName = "�����L�O�̓�"
                End If
         End If
    Case 3
        If targetyear > 1948 Then
            If targetday = Syunbun(targetyear) Then
                hantei = True
                HollydayName = "�t���̓�"
            End If
        End If
    Case 4
        If targetday = 29 Then
            If targetyear > 1948 Then
                If 1989 > targetyear Then
                    hantei = True
                    HollydayName = "�V�c�a����"
            ElseIf 2007 > targetyear And targetyear > 1988 Then
                    hantei = True
                    HollydayName = "�݂ǂ�̓�"
                Else
                    hantei = True
                    HollydayName = "���a�̓�"
                End If
            End If
        End If
    Case 5
        If targetyear > 1948 Then
            If targetday = 3 Then
                    hantei = True
                    HollydayName = "���@�L�O��"
            End If
            If targetday = 5 Then
                    hantei = True
                    HollydayName = "���ǂ��̓�"
            End If
            If targetday = 4 Then
                If targetyear > 2006 Then
                    hantei = True
                    HollydayName = "�݂ǂ�̓�"
                End If
            End If
        End If
    Case 7
        If targetyear > 1995 And targetyear <> 2020 Then
            If 2004 > targetyear Then
                If targetday = 20 Then
                    hantei = True
                    HollydayName = "�C�̓�"
                End If
            Else
                If CInt(Format(DaiXYoubi(targetyear, 7, 3, 0), "d")) = targetday Then
                    hantei = True
                    HollydayName = "�C�̓�"
                End If
            End If
        End If
    Case 8
            If targetyear >= 2016 And targetyear <> 2020 Then
                If targetday = 11 Then
                    hantei = True
                    HollydayName = "�R�̓�"
                End If
             End If
    Case 9
        If targetyear > 1965 Then
            If 2004 > targetyear Then
                If targetday = 15 Then
                    hantei = True
                    HollydayName = "�h�V�̓�"
                End If
            Else
                If targetyear > 2003 And CInt(Format(DaiXYoubi(targetyear, 9, 3, 1), "d")) = targetday Then
                    hantei = True
                    HollydayName = "�h�V�̓�"
                End If
            End If
        End If
        If targetyear > 1947 Then
            If targetday = Syuubun(targetyear) Then
                hantei = True
                HollydayName = "�H���̓�"
            End If
        End If
    Case 10
        If targetyear > 1965 Then
            If 2000 > targetyear Then
                If targetday = 10 Then
                    hantei = True
                    HollydayName = "�̈�̓�"
                End If
            ElseIf targetyear > 1999 Then
                If CInt(Format(DaiXYoubi(targetyear, 10, 2, 1), "d")) = targetday Then
                    hantei = True
                    HollydayName = "�̈�̓�"
                End If
            End If
        End If
    Case 11
        If targetyear > 1947 Then
            If targetday = 3 Then
                hantei = True
                HollydayName = "�����̓�"
            ElseIf targetday = 23 Then
                hantei = True
                HollydayName = "�ΘJ���ӂ̓�"
            End If
        End If
    Case 12
        If targetyear > 1988 Then
            If targetday = 23 Then
                hantei = True
                HollydayName = "�V�c�a����"
            End If
        End If
    End Select
    If hantei = True Then
        NationalHollydays = True
    Else
        NationalHollydays = False
    End If
End Function
'******************************************************************************
' �t���̓������߂�
'******************************************************************************
Public Function Syunbun(Nen As Integer) As Integer
    syubun = 0
    If (1899 >= Nen And Nen >= 1851) Then
        Syunbun = Int(19.8277 + 0.242194 * (Nen - 1980) - Int((Nen - 1983) / 4))
    End If
    If (1979 >= Nen And Nen >= 1900) Then
        Syunbun = Int(20.8357 + 0.242194 * (Nen - 1980) - Int((Nen - 1983) / 4))
    End If
    If 2099 >= Nen And Nen >= 1980 Then
        Syunbun = Int(20.8431 + 0.242194 * (Nen - 1980) - Int((Nen - 1980) / 4))
    End If
    If (2150 >= Nen And Nen >= 2100) Then
        Syunbun = Int(21.851 + 0.242194 * (Nen - 1980) - Int((Nen - 1980) / 4))
    End If

End Function
'******************************************************************************
' �H���̓������߂�
'******************************************************************************
Public Function Syuubun(Nen As Integer) As Integer
    Syuubun = 0
    If (1899 >= Nen And Nen >= 1851) Then
        Syuubun = Int(22.2588 + 0.242194 * (Nen - 1980) - Int((Nen - 1983) / 4))
    End If
    If (1979 >= Nen And Nen >= 1900) Then
        Syuubun = Int(23.2588 + 0.242194 * (Nen - 1980) - Int((Nen - 1983) / 4))
    End If
    If (2099 >= Nen And Nen >= 1980) Then
        Syuubun = Int(23.2488 + 0.242194 * (Nen - 1980) - Int((Nen - 1980) / 4))
    End If
    If (2150 >= Nen And Nen >= 2100) Then
        Syuubun = Int(24.2488 + 0.242194 * (Nen - 1980) - Int((Nen - 1980) / 4))
    End If
End Function
'******************************************************************************
' ���錎�̑恛���j���������ł��邩�𒲂ׂ�֐��B
'******************************************************************************
Public Function DaiXYoubi(y, m, n, Yobi As Integer) As String
    DaiXYoubi = ((9 - Weekday(DateSerial(y, m, 0))) + (n - 1) * 7 + 1)
End Function
'******************************************************************************
' �U�֋x�����𒲂ׂ�֐��B
'******************************************************************************
Public Function FurikaeKyujitsu(targetdate As Date, HollydayName As String) As Boolean
Dim lastsunday  As Date
Dim days As Integer
    HollydayName = ""
    hantei = False
    lastsunday = DateAdd("d", 1 - (Weekday(targetdate)), targetdate)
    days = (Weekday(targetdate) - 1)
    If targetdate > "1973/04/11" Then
        If NationalHollydays(targetdate, HollydayName) = False Then
            If targetyear < 2007 Then
                If NationalHollydays(DateAdd("d", -1, targetdate), HollydayName) = True And Weekday(targetdate) = 2 Then
                    HollydayName = "�U�֋x��"
                    FurikaeKyujitsu = True
                Else
                    HollydayName = ""
                    FurikaeKyujitsu = False
                End If
            Else
                If NationalHollydays(lastsunday, HollydayName) = True Then
                    For i = 0 To (days - 1)
                        If NationalHollydays(DateAdd("d", i, lastsunday), HollydayName) = False Then
                            FurikaeKyujitsu = False
                            HollydayName = ""
                            Exit Function
                        End If
                    Next i
                HollydayName = "�U�֋x��"
                FurikaeKyujitsu = True
                Else
                    FurikaeKyujitsu = False
                    HollydayName = ""
                End If
            End If
        End If
    End If
End Function
'******************************************************************************
' �����̋x�����𒲂ׂ�֐��B
'******************************************************************************
Public Function KokuminnoKyujitsu(targetdate As Date, HollydayName As String) As Boolean
    HollydayName = ""
    If targetdate > "1985/12/26" Then
        If NationalHollydays(targetdate, HollydayName) = False Then
            If targetyear < 2007 Then
                If FurikaeKyujitsu(targetdate, HollydayName) = False And Weekday(targetdate) <> 1 Then
                    If NationalHollydays(DateAdd("d", -1, targetdate), HollydayName) = True And NationalHollydays(DateAdd("d", 1, targetdate), HollydayName) = True Then
                        HollydayName = "�����̋x��"
                        KokuminnoKyujitsu = True
                    Else
                        HollydayName = ""
                        KokuminnoKyujitsu = False
                    End If
                Else
                    HollydayName = ""
                    KokuminnoKyujitsu = False
                End If
            Else
                If NationalHollydays(targetdate, HollydayName) = False Then
                    If NationalHollydays(DateAdd("d", -1, targetdate), HollydayName) = True And NationalHollydays(DateAdd("d", 1, targetdate), HollydayName) = True Then
                        HollydayName = "�����̋x��"
                        KokuminnoKyujitsu = True
                    Else
                        HollydayName = ""
                        KokuminnoKyujitsu = False
                    End If
                Else
                    HollydayName = ""
                    KokuminnoKyujitsu = False
                End If
            End If
        End If
    End If
End Function
'******************************************************************************
' ���ʂȋx��
'******************************************************************************
Public Function TokubetsunaKyujitsu(targetdate As Date, HollydayName As String) As Boolean
    Dim line As Long, endLine As Long
    
    TokubetsunaKyujitsu = False
    If targetdate = "1959/04/10" Then
        HollydayName = "���m�e�������̋V"
        TokubetsunaKyujitsu = True
    End If
    If targetdate = "1989/02/24" Then
        HollydayName = "���a�V�c��r�̗�"
        TokubetsunaKyujitsu = True
    End If
    If targetdate = "1990/11/12" Then
        HollydayName = "���ʗ琳�a�̋V"
        TokubetsunaKyujitsu = True
    End If
    If targetdate = "1993/06/09" Then
        HollydayName = "���m�e�������̋V"
        TokubetsunaKyujitsu = True
    End If
    
    If targetdate = "2020/07/23" Then
        HollydayName = "�C�̓�"
        TokubetsunaKyujitsu = True
    End If
    If targetdate = "2020/07/24" Then
        HollydayName = "�X�|�[�c�̓�"
        TokubetsunaKyujitsu = True
    End If
    If targetdate = "2020/08/10" Then
        HollydayName = "�R�̓�"
        TokubetsunaKyujitsu = True
    End If
    
    
    '��Ўw��x���̐ݒ�
    endLine = sheetSetting.Cells(Rows.count, Library.getColumnNo(setVal("cell_CompanyHoliday"))).End(xlUp).row
    For line = 3 To endLine
      If targetdate = sheetSetting.Range(setVal("cell_CompanyHoliday") & line) Then
          HollydayName = "��Ўw��x��"
          TokubetsunaKyujitsu = True
      End If
    Next
    
    
    
    
    
End Function
