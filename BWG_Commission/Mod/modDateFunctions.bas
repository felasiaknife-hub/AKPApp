Attribute VB_Name = "modDateFunctions"
Option Explicit
Enum DueDateType
    EndMonth = 1
    RoundMounth = 2
    StartNextMonth = 3
End Enum
Public TodayDate As String

Public Function CalPaymentDueDate(StartDate As String, CreditDay As Integer, calType As DueDateType) As String  'yyyy/mm/dd
Dim tmpSDate As Date
Dim tmpDueDate As Date
Dim y$, m$, d$
    y = Left(StartDate, 4)
    m = Mid(StartDate, 6, 2)
    d = Right(StartDate, 2)
    If IsNumeric(y) And IsNumeric(m) And IsNumeric(d) Then
        If CInt(y) > 2500 Then
            tmpSDate = DateSerial(CInt(y) - 543, CInt(m), CInt(d))
        Else
            tmpSDate = DateSerial(CInt(y), CInt(m), CInt(d))
        End If
    Else
        CalPaymentDueDate = ""
        Exit Function
    End If
    If calType = EndMonth Then
        If Year(tmpSDate) > 2500 Then
            tmpDueDate = DateSerial(Year(tmpSDate) - 543, Month(tmpSDate), Day(tmpSDate) + CreditDay)
        Else
            tmpDueDate = DateSerial(Year(tmpSDate), Month(tmpSDate), Day(tmpSDate) + CreditDay)
        End If
        tmpDueDate = DateSerial(Year(tmpDueDate), Month(tmpDueDate) + 1, 0)
    End If
    If Year(tmpDueDate) > 2500 Then
        CalPaymentDueDate = Year(tmpDueDate) & "/" & AddZero(Month(tmpDueDate)) & "/" & AddZero(Day(tmpDueDate))
    Else
        CalPaymentDueDate = Year(tmpDueDate) + 543 & "/" & AddZero(Month(tmpDueDate)) & "/" & AddZero(Day(tmpDueDate))
    End If
End Function

Private Function AddZero(str As String) As String
    AddZero = Format(str, "00")
End Function

Public Function FormatYMD_Thai(sDay, sMonth, sYear, Optional dDelim As String = "") As String
Dim strD$, strM$, strY$
    strD = AddZero("" & sDay)
    strM = AddZero("" & sMonth)
    If sYear > 2500 Then
        strY = "" & sYear
    Else
        strY = CStr(sYear + 543)
    End If
    
    FormatYMD_Thai = Trim(strY) & dDelim & Trim(strM) & dDelim & Trim(strD)
End Function

Public Function FormatYMD_Eng(sDay%, sMonth%, sYear%, Optional dDelim As String = "") As String
Dim strD$, strM$, strY$
    strD = AddZero("" & sDay)
    strM = AddZero("" & sMonth)
    If sYear > 2500 Then
        strY = "" & (sYear - 543)
    Else
        strY = "" & sYear
    End If
    
    FormatYMD_Eng = Trim(strY) & dDelim & Trim(strM) & dDelim & Trim(strD)
End Function

Public Function FormatYMD_to_DMY(sDate1 As String, Optional dDelim1 As String = "/", Optional dDelim2 As String = "/") As String
Dim tmpArr
Dim tmpStr$
Dim d$, m$, y$
    If InStr(1, sDate1, dDelim1) = 0 Then
        FormatYMD_to_DMY = ""
        Exit Function
    End If
    
    tmpArr = Split(sDate1, dDelim1)
    
    d = AddZero(Trim("" & tmpArr(2)))
    m = AddZero(Trim("" & tmpArr(1)))
    y = Trim("" & tmpArr(0))
    
    FormatYMD_to_DMY = d & dDelim2 & m & dDelim2 & y
End Function

Public Function FormatDMY_to_YMD(sDate1 As String, dDelim1 As String, dDelim2 As String) As String
Dim tmpArr
Dim tmpStr$
Dim d$, m$, y$
    If InStr(1, sDate1, dDelim1) = 0 Then
        FormatDMY_to_YMD = ""
        Exit Function
    End If
    
    tmpArr = Split(sDate1, dDelim1)
    
    d = AddZero(Trim("" & tmpArr(0)))
    m = AddZero(Trim("" & tmpArr(1)))
    y = Trim("" & tmpArr(2))
    
    FormatDMY_to_YMD = y & dDelim2 & m & dDelim2 & d
End Function

Public Function isDMYDate(selDate As String, delimStr As String) As Boolean
Dim d$, m$, y$
Dim splitStr
Dim date1 As Date
Dim date2 As Date
    isDMYDate = False
    If InStr(1, selDate, delimStr) = 0 Then Exit Function
    splitStr = Split(selDate, delimStr)
    d = splitStr(0)
    m = splitStr(1)
    y = splitStr(2)
    If Not IsNumeric(d) Then Exit Function
    If Not IsNumeric(m) Then Exit Function
    If Not IsNumeric(y) Then Exit Function
    If CInt(d) > 31 Then Exit Function
    If CInt(m) > 12 Then Exit Function
    date1 = DateSerial(CInt(y), CInt(m), CInt(d))
    date2 = DateSerial(CInt(y), CInt(m), 1)
    If Month(date1) <> Month(date2) Or Year(date1) <> Year(date2) Then Exit Function
    isDMYDate = True
End Function

