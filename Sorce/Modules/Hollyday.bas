Attribute VB_Name = "Hollyday"
'******************************************************************************
' ‚ ‚é“ú‚ªj“ú‚Å‚ ‚é‚©H‚»‚Ìê‡‚Ç‚Ìj“ú‚©H‚ğ’²‚×‚éŠÖ”B
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
' j“ú”»’èŠÖ”
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
                HollydayName = "Œ³’U"
        End If
        If targetyear > 1948 Then
            If targetyear < 2000 Then
                If targetday = 15 Then
                    hantei = True
                    HollydayName = "¬l‚Ì“ú"
                End If
            ElseIf CInt(Format(DaiXYoubi(targetyear, 1, 2, 1), "d")) = targetday Then
                    hantei = True
                    HollydayName = "¬l‚Ì“ú"
            End If
        End If
    Case 2
        If targetyear > 1966 Then
                If targetday = 11 Then
                    hantei = True
                    HollydayName = "Œš‘‹L”O‚Ì“ú"
                End If
         End If
    Case 3
        If targetyear > 1948 Then
            If targetday = Syunbun(targetyear) Then
                hantei = True
                HollydayName = "t•ª‚Ì“ú"
            End If
        End If
    Case 4
        If targetday = 29 Then
            If targetyear > 1948 Then
                If 1989 > targetyear Then
                    hantei = True
                    HollydayName = "“Vc’a¶“ú"
            ElseIf 2007 > targetyear And targetyear > 1988 Then
                    hantei = True
                    HollydayName = "‚İ‚Ç‚è‚Ì“ú"
                Else
                    hantei = True
                    HollydayName = "º˜a‚Ì“ú"
                End If
            End If
        End If
    Case 5
        If targetyear > 1948 Then
            If targetday = 3 Then
                    hantei = True
                    HollydayName = "Œ›–@‹L”O“ú"
            End If
            If targetday = 5 Then
                    hantei = True
                    HollydayName = "‚±‚Ç‚à‚Ì“ú"
            End If
            If targetday = 4 Then
                If targetyear > 2006 Then
                    hantei = True
                    HollydayName = "‚İ‚Ç‚è‚Ì“ú"
                End If
            End If
        End If
    Case 7
        If targetyear > 1995 And targetyear <> 2020 Then
            If 2004 > targetyear Then
                If targetday = 20 Then
                    hantei = True
                    HollydayName = "ŠC‚Ì“ú"
                End If
            Else
                If CInt(Format(DaiXYoubi(targetyear, 7, 3, 0), "d")) = targetday Then
                    hantei = True
                    HollydayName = "ŠC‚Ì“ú"
                End If
            End If
        End If
    Case 8
            If targetyear >= 2016 And targetyear <> 2020 Then
                If targetday = 11 Then
                    hantei = True
                    HollydayName = "R‚Ì“ú"
                End If
             End If
    Case 9
        If targetyear > 1965 Then
            If 2004 > targetyear Then
                If targetday = 15 Then
                    hantei = True
                    HollydayName = "Œh˜V‚Ì“ú"
                End If
            Else
                If targetyear > 2003 And CInt(Format(DaiXYoubi(targetyear, 9, 3, 1), "d")) = targetday Then
                    hantei = True
                    HollydayName = "Œh˜V‚Ì“ú"
                End If
            End If
        End If
        If targetyear > 1947 Then
            If targetday = Syuubun(targetyear) Then
                hantei = True
                HollydayName = "H•ª‚Ì“ú"
            End If
        End If
    Case 10
        If targetyear > 1965 Then
            If 2000 > targetyear Then
                If targetday = 10 Then
                    hantei = True
                    HollydayName = "‘Ìˆç‚Ì“ú"
                End If
            ElseIf targetyear > 1999 Then
                If CInt(Format(DaiXYoubi(targetyear, 10, 2, 1), "d")) = targetday Then
                    hantei = True
                    HollydayName = "‘Ìˆç‚Ì“ú"
                End If
            End If
        End If
    Case 11
        If targetyear > 1947 Then
            If targetday = 3 Then
                hantei = True
                HollydayName = "•¶‰»‚Ì“ú"
            ElseIf targetday = 23 Then
                hantei = True
                HollydayName = "‹Î˜JŠ´Ó‚Ì“ú"
            End If
        End If
    Case 12
        If targetyear > 1988 Then
            If targetday = 23 Then
                hantei = True
                HollydayName = "“Vc’a¶“ú"
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
' t•ª‚Ì“ú‚ğ‹‚ß‚é
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
' H•ª‚Ì“ú‚ğ‹‚ß‚é
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
' ‚ ‚éŒ‚Ì‘æ›¢—j“ú‚ª “ú‚Å‚ ‚é‚©‚ğ’²‚×‚éŠÖ”B
'******************************************************************************
Public Function DaiXYoubi(y, m, n, Yobi As Integer) As String
    DaiXYoubi = ((9 - Weekday(DateSerial(y, m, 0))) + (n - 1) * 7 + 1)
End Function
'******************************************************************************
' U‘Ö‹x“ú‚©‚ğ’²‚×‚éŠÖ”B
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
                    HollydayName = "U‘Ö‹x“ú"
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
                HollydayName = "U‘Ö‹x“ú"
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
' ‘–¯‚Ì‹x“ú‚©‚ğ’²‚×‚éŠÖ”B
'******************************************************************************
Public Function KokuminnoKyujitsu(targetdate As Date, HollydayName As String) As Boolean
    HollydayName = ""
    If targetdate > "1985/12/26" Then
        If NationalHollydays(targetdate, HollydayName) = False Then
            If targetyear < 2007 Then
                If FurikaeKyujitsu(targetdate, HollydayName) = False And Weekday(targetdate) <> 1 Then
                    If NationalHollydays(DateAdd("d", -1, targetdate), HollydayName) = True And NationalHollydays(DateAdd("d", 1, targetdate), HollydayName) = True Then
                        HollydayName = "‘–¯‚Ì‹x“ú"
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
                        HollydayName = "‘–¯‚Ì‹x“ú"
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
' “Á•Ê‚È‹x“ú
'******************************************************************************
Public Function TokubetsunaKyujitsu(targetdate As Date, HollydayName As String) As Boolean
    Dim line As Long, endLine As Long
    
    TokubetsunaKyujitsu = False
    If targetdate = "1959/04/10" Then
        HollydayName = "–¾me‰¤Œ‹¥‚Ì‹V"
        TokubetsunaKyujitsu = True
    End If
    If targetdate = "1989/02/24" Then
        HollydayName = "º˜a“Vc‘å‘r‚Ì—ç"
        TokubetsunaKyujitsu = True
    End If
    If targetdate = "1990/11/12" Then
        HollydayName = "‘¦ˆÊ—ç³“a‚Ì‹V"
        TokubetsunaKyujitsu = True
    End If
    If targetdate = "1993/06/09" Then
        HollydayName = "“¿me‰¤Œ‹¥‚Ì‹V"
        TokubetsunaKyujitsu = True
    End If
    
    If targetdate = "2020/07/23" Then
        HollydayName = "ŠC‚Ì“ú"
        TokubetsunaKyujitsu = True
    End If
    If targetdate = "2020/07/24" Then
        HollydayName = "ƒXƒ|[ƒc‚Ì“ú"
        TokubetsunaKyujitsu = True
    End If
    If targetdate = "2020/08/10" Then
        HollydayName = "R‚Ì“ú"
        TokubetsunaKyujitsu = True
    End If
    
    
    '‰ïĞw’è‹x“ú‚Ìİ’è
    endLine = setSheet.Cells(Rows.count, 13).End(xlUp).row
    For line = 3 To endLine
      If targetdate = setSheet.Range("M" & line) Then
          HollydayName = "‰ïĞw’è‹x“ú"
          TokubetsunaKyujitsu = True
      End If
    Next
    
    
    
    
    
End Function
