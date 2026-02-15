Attribute VB_Name = "Module1"
'---------------------------------------------
' Accurate Persian (Jalali) to Gregorian Date Conversion (Excel 2013 Compatible)
'---------------------------------------------
Function JalaliToGregorianFixed(jy As Long, jm As Long, jd As Long) As Date
    Dim gy As Long, gm As Long, gd As Long
    Dim g_d_m(12) As Long
    g_d_m(1) = 0: g_d_m(2) = 31: g_d_m(3) = 59: g_d_m(4) = 90: g_d_m(5) = 120
    g_d_m(6) = 151: g_d_m(7) = 181: g_d_m(8) = 212: g_d_m(9) = 243
    g_d_m(10) = 273: g_d_m(11) = 304: g_d_m(12) = 334

    Dim jy2, jm2, jd2 As Long
    jy2 = jy - 979
    jm2 = jm - 1
    jd2 = jd - 1

    Dim j_day_no As Long
    j_day_no = 365 * jy2 + (jy2 \ 33) * 8 + ((jy2 Mod 33) + 3) \ 4

    Dim i As Long
    For i = 0 To jm2 - 1
        If i < 6 Then
            j_day_no = j_day_no + 31
        Else
            j_day_no = j_day_no + 30
        End If
    Next i

    j_day_no = j_day_no + jd2

    Dim g_day_no As Long
    g_day_no = j_day_no + 79

    gy = 1600 + 400 * (g_day_no \ 146097)
    g_day_no = g_day_no Mod 146097

    Dim leap As Boolean
    leap = True

    If g_day_no >= 36525 Then
        g_day_no = g_day_no - 1
        gy = gy + 100 * (g_day_no \ 36524)
        g_day_no = g_day_no Mod 36524
        If g_day_no >= 365 Then
            g_day_no = g_day_no + 1
        Else
            leap = False
        End If
    End If

    gy = gy + 4 * (g_day_no \ 1461)
    g_day_no = g_day_no Mod 1461

    If g_day_no >= 366 Then
        leap = False
        g_day_no = g_day_no - 1
        gy = gy + g_day_no \ 365
        g_day_no = g_day_no Mod 365
    End If

    For gm = 1 To 12
        Dim v As Long
        v = g_d_m(gm + 1) - g_d_m(gm)
        If leap And gm = 2 Then v = v + 1
        If g_day_no < v Then Exit For
        g_day_no = g_day_no - v
    Next gm

    gd = g_day_no + 1
    JalaliToGregorianFixed = DateSerial(gy, gm, gd)
End Function

'---------------------------------------------
' Macro for Automatic Conversion
'---------------------------------------------
Sub ConvertShamsiToGregorian_Fixed()
    Dim ws As Worksheet
    Dim LastRow As Long, i As Long
    Dim ShamsiDate As String
    Dim YearSh As Long, MonthSh As Long, DaySh As Long
    Dim GregDate As Date

    Set ws = ActiveSheet
    LastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ws.Cells(1, "P").Value = "Gregorian_Date"

    For i = 2 To LastRow
        ShamsiDate = Trim(ws.Cells(i, "A").Value)

        If Len(ShamsiDate) = 10 And InStr(ShamsiDate, "/") > 0 Then
            YearSh = CLng(Split(ShamsiDate, "/")(0))
            MonthSh = CLng(Split(ShamsiDate, "/")(1))
            DaySh = CLng(Split(ShamsiDate, "/")(2))

            GregDate = JalaliToGregorianFixed(YearSh, MonthSh, DaySh)
            ws.Cells(i, "P").Value = GregDate
        End If
    Next i

    ws.Columns("P").NumberFormat = "yyyy-mm-dd"
    MsgBox "Date conversion completed successfully!", vbInformation
End Sub

