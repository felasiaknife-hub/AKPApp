Attribute VB_Name = "modUtil"
Option Explicit
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const SWP_SHOWWINDOW = &H40
Public Const SWP_NOACTIVATE = &H10
Public Const HWND_NOTOPMOST = -2
Public Const HWND_TOPMOST = -1
Public ProductID As String
Public ProductPassword As String

Private Type GUID_
  Data1 As Long
  Data2 As Integer
  Data3 As Integer
  Data4(7) As Byte
End Type

Private Declare Function CoCreateGuid Lib "ole32.dll" (guid As GUID_) As Long
Private Declare Function StringFromGUID2 Lib "ole32.dll" (guid As GUID_, ByVal str As String, ByVal cchMax As Long) As Long
Public TimeInterval As Long

Public Function GetGUID() As String
  Dim g As GUID_, s As String
  s = String(76, vbNullChar)
  CoCreateGuid g
  StringFromGUID2 g, s, Len(s)
  s = StrConv(s, vbFromUnicode)
' GetGUID = Trim(Replace(Replace(Replace(s, "}", ""), "{", ""), "-", ""))
 Dim sqlStr As String
 Dim tmpRec As ADODB.Recordset
 Set tmpRec = New ADODB.Recordset

    sqlStr = " exec [dbo].[sp_GetID] "
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    GetGUID = tmpRec(0)

End Function
Public Function getListString(i As Integer, str As String) As String
Dim tmpStr As String
    If Trim(str) = vbNullString Then
        Exit Function
    End If
    If InStr(1, str, " - <") = 0 Then Exit Function
    
    If i = 0 Then
        tmpStr = Trim(Split(str, " - <")(0))
    Else
        tmpStr = Trim(Split(str, " - <")(1))
        tmpStr = Mid(tmpStr, 1, Len(tmpStr) - 1)
    End If
    getListString = tmpStr
End Function

Public Function getNewProductID(strPre As String) As String
Dim dt
    dt = Now
    getNewProductID = Trim(strPre) & Year(dt) & Format(Month(dt), "00") & Format(Day(dt), "00") & Format(Hour(dt), "00") & Format(Minute(dt), "00") & Format(Second(dt), "00")
End Function

Public Function isExist(tmpCol As Object, strKey As String) As Boolean
On Error GoTo genErr:
Dim tmpObj As Object
    Set tmpObj = tmpCol(strKey)
    isExist = True
    Exit Function
genErr:
    isExist = False
End Function

Public Function NullToZero(sVal)
    If IsNull(sVal) Then
        NullToZero = 0
    Else
        NullToZero = sVal
    End If
End Function

Public Function AmountToString(selAmt As Double) As String
Dim amt1 As Double
Dim amt2 As Double
Dim cc As String
    cc = Trim(Format(selAmt, "############0.00"))
    If IsNumeric(Split(cc, ".")(0)) Then amt1 = CDbl(Split(cc, ".")(0))
    If IsNumeric(Split(cc, ".")(1)) Then amt2 = CDbl(Split(cc, ".")(1))

    If amt2 > 0 Then
        AmountToString = ReadAmountToString(amt1) & "บาท" & ReadAmountToString(amt2) & "สตางค์"
    Else
        AmountToString = ReadAmountToString(amt1) & "บาทถ้วน"
    End If
End Function

Private Function ReadAmountToString(selAmt As Double) As String
Dim x, Number(0 To 9), Position(1 To 7)
Dim a As String
Dim cc As String
    a = ""
    cc = Trim(Format(selAmt, "############0"))
    If Len(cc) > 7 Then
        a = ReadAmountToString(CDbl(Left(cc, Len(cc) - 6))) & "ล้าน"
        cc = Mid(cc, Len(cc) - 5)
    End If
    Number(0) = "ศูนย์"
    Number(1) = "หนึ่ง"
    Number(2) = "สอง"
    Number(3) = "สาม"
    Number(4) = "สี่"
    Number(5) = "ห้า"
    Number(6) = "หก"
    Number(7) = "เจ็ด"
    Number(8) = "แปด"
    Number(9) = "เก้า"
    Position(1) = ""
    Position(2) = "สิบ"
    Position(3) = "ร้อย"
    Position(4) = "พัน"
    Position(5) = "หมื่น"
    Position(6) = "แสน"
    Position(7) = "ล้าน"
    
    For x = 1 To Len(Trim(cc))
         If x = Len(Trim(cc)) - 1 And Mid(Trim(cc), x, 1) = "1" And Len(Trim(cc)) <> 1 Then
            a = a + Position(Len(Trim(cc)) - x + 1)
        ElseIf x = Len(Trim(cc)) - 1 And Mid(Trim(cc), x, 1) = "2" And Len(Trim(cc)) <> 1 Then
            a = a + "ยี่สิบ" ' + Position(Len(Trim(cc)) - x + 1)
        ElseIf x = Len(Trim(cc)) And Mid(Trim(cc), x, 1) = "1" And Len(Trim(cc)) <> 1 Then
            If Right(a, 3) = "สิบ" Then
                a = a + "เอ็ด" + Position(Len(Trim(cc)) - x + 1)
            Else
                a = a + Number(Val(Mid(Trim(cc), x, 1))) + Position(Len(Trim(cc)) - x + 1)
            End If
        ElseIf Mid(Trim(cc), x, 1) <> "0" Then
            a = a + Number(Val(Mid(Trim(cc), x, 1))) + Position(Len(Trim(cc)) - x + 1)
        End If
    Next
    ReadAmountToString = Trim(a)
End Function

Public Function ClearEscString(selStr As String) As String
Dim tmpStr As String
    tmpStr = Replace(selStr, Chr$(10), vbNullString)
    tmpStr = Replace(tmpStr, vbCrLf, vbNullString)
    tmpStr = Replace(tmpStr, vbCr, vbNullString)
    tmpStr = Replace(tmpStr, vbLf, vbNullString)
    ClearEscString = Trim(tmpStr)
End Function

Public Function TrimString(oriStr As String, maxLen As Long) As String
    If Len(Trim(oriStr)) > maxLen Then
        TrimString = Trim(Left(Trim(oriStr), maxLen)) & "..."
    Else
        TrimString = Trim(oriStr)
    End If
End Function

Public Function ClearSgQuote(selStr As String) As String
    ClearSgQuote = Replace(selStr, "'", "|")
End Function

Public Function ReturnSgQuote(selStr As String) As String
    ReturnSgQuote = Replace(selStr, "|", "'")
End Function

Public Function isFormLoaded(frmName As String) As Boolean
Dim x%
    isFormLoaded = False
    For x = 0 To Forms.Count - 1
        If LCase(Forms(x).Name) = LCase(frmName) Then
            isFormLoaded = True
            Exit For
        End If
    Next
End Function

Public Sub SetTextToCombo(selTxt As String, selCbo As ComboBox)
On Error GoTo setErr:
    selCbo.Text = Trim(selTxt)
    Exit Sub
setErr:
    Exit Sub
End Sub

Public Function GetVolumeID() As String
Dim myDrive As Drive
Dim objFile As New FileSystemObject
    Set myDrive = objFile.Drives("C")
    GetVolumeID = myDrive.SerialNumber
    Set myDrive = Nothing
    Set objFile = Nothing
End Function

Public Function getPassword(p_id As String) As String
Dim x, y, z, u, v, uu, ran, p1, p2, p3
aaa:
ran = Int((59.4 - 4.3 + 1) * Rnd + 4.3) / 59.43
    For y = 1 To Int(ran * 5.2349)
        For x = 1 To Len(p_id)
            y = Int(ran * y * 21.95) + Trim(str(Val(Mid(Trim(str(p_id)), x, 1)) And Int(Rnd(p_id / 15 * x * ran * 10) * 51.56 * ran))) '+ Chr(Rnd(x) * 10 + Val(Mid(Trim(Str(P_id)), x, 1)) + 64)
        Next
        u = u + Trim(str(y))
    Next
If Len(Trim(u)) < 15 Then
    GoTo aaa:
End If
ProductID = Left(Trim(u), 5) + "-" + Mid(Trim(u), 6, Len(u) - 10) + "-" + Right(Trim(u), 5)
p1 = Left(Trim(u), 5)
p2 = Right(Trim(u), 5)
p3 = Mid(Trim(u), 5, 5)
p2 = Val(p2) And 57295
p1 = Val(p1) And 17190
p3 = Val(p3) And 30171
ProductPassword = Trim(str(p2)) + "-" + Trim(str(p3)) + "-" + Trim(str(p1))
End Function

Public Sub PopulateProvince(selCbo As ComboBox)
    selCbo.Clear
    With selCbo
        .AddItem "เชียงใหม่"
        .AddItem "เชียงราย"
        .AddItem "เพชรบุรี"
        .AddItem "เพชรบูรณ์"
        .AddItem "เลย"
        .AddItem "แพร่"
        .AddItem "แม่ฮ่องสอน"
        .AddItem "กระบี่"
        .AddItem "กรุงเทพมหานคร"
        .AddItem "กาญจนบุรี"
        .AddItem "กาฬสินธุ์"
        .AddItem "กำแพงเพชร"
        .AddItem "ขอนแก่น"
        .AddItem "จันทบุรี"
        .AddItem "ฉะเชิงเทรา"
        .AddItem "ชลบุรี"
        .AddItem "ชัยนาท"
        .AddItem "ชัยภูมิ"
        .AddItem "ชุมพร"
        .AddItem "ตรัง"
        .AddItem "ตราด"
        .AddItem "ตาก"
        .AddItem "นครนายก"
        .AddItem "นครปฐม"
        .AddItem "นครพนม"
        .AddItem "นครราชสีมา"
        .AddItem "นครศรีธรรมราช"
        .AddItem "นครสวรรค์"
        .AddItem "นนทบุรี"
        .AddItem "นราธิวาส"
        .AddItem "น่าน"
        .AddItem "บุรีรัมย์"
        .AddItem "ปทุมธานี"
        .AddItem "ประจวบคีรีขันธ์"
        .AddItem "ปราจีนบุรี"
        .AddItem "ปัตตานี"
        .AddItem "พระนครศรีอยุธยา"
        .AddItem "พะเยา"
        .AddItem "พังงา"
        .AddItem "พัทลุง"
        .AddItem "พิจิตร"
        .AddItem "พิษณุโลก"
        .AddItem "ภูเก็ต"
        .AddItem "มหาสารคาม"
        .AddItem "มุกดาหาร"
        .AddItem "ยโสธร"
        .AddItem "ยะลา"
        .AddItem "ร้อยเอ็ด"
        .AddItem "ระนอง"
        .AddItem "ระยอง"
        .AddItem "ราชบุรี"
        .AddItem "ลพบุรี"
        .AddItem "ลำปาง"
        .AddItem "ลำพูน"
        .AddItem "ศรีสะเกษ"
        .AddItem "สกลนคร"
        .AddItem "สงขลา"
        .AddItem "สตูล"
        .AddItem "สมุทรปราการ"
        .AddItem "สมุทรสงคราม"
        .AddItem "สมุทรสาคร"
        .AddItem "สระแก้ว"
        .AddItem "สระบุรี"
        .AddItem "สิงห์บุรี"
        .AddItem "สุโขทัย"
        .AddItem "สุพรรณบุรี"
        .AddItem "สุราษฎร์ธานี"
        .AddItem "สุรินทร์"
        .AddItem "หนองคาย"
        .AddItem "หนองบัวลำภู"
        .AddItem "อ่างทอง"
        .AddItem "อำนาจเจริญ"
        .AddItem "อุดรธานี"
        .AddItem "อุตรดิตถ์"
        .AddItem "อุทัยธานี"
        .AddItem "อุบลราชธานี"
    End With
End Sub

Public Function getFileName(selStr As String) As String
Dim x%
Dim tmpStr As String
    tmpStr = Trim(selStr)
    x = 1
    Do Until InStr(x, tmpStr, "\") = 0
        tmpStr = Mid(tmpStr, InStr(x, tmpStr, "\") + 1)
    Loop
    getFileName = Trim(tmpStr)
End Function

Public Function TF_To_YN(selVal As Boolean) As String
    If selVal Then
        TF_To_YN = "Y"
    Else
        TF_To_YN = "N"
    End If
End Function

Public Sub SetTextFocus(selTxt As TextBox)
    selTxt.SelStart = 0
    selTxt.SelLength = Len(selTxt.Text)
End Sub

Public Function Format2Decimal(selNum)
Dim leftStr As String
Dim rightStr As String
    If Not IsNumeric(selNum) Then
        Format2Decimal = 0
    Else
        If InStr(1, Trim(CStr(selNum)), ".") > 0 Then
            leftStr = Split("" & selNum, ".")(0)
            rightStr = Split("" & selNum, ".")(1)
            If Len(rightStr) > 2 Then
                'Format2Decimal = Format(leftStr, "###,###,##") & "." & Left(rightStr, 2)
                Format2Decimal = Format(selNum, "###,###,##0.00")
            Else
                Format2Decimal = Format(selNum, "###,###,##0.00")
            End If
        Else
            Format2Decimal = Format(selNum, "###,###,##0.00")
        End If
    End If
End Function

Public Function ThaiMonthName(selMonth As Integer)
    Select Case selMonth
        Case 1
            ThaiMonthName = "ม.ค."
        Case 2
            ThaiMonthName = "ก.พ."
        Case 3
            ThaiMonthName = "มี.ค."
        Case 4
            ThaiMonthName = "เม.ย."
        Case 5
            ThaiMonthName = "พ.ค."
        Case 6
            ThaiMonthName = "มิ.ย."
        Case 7
            ThaiMonthName = "ก.ค."
        Case 8
            ThaiMonthName = "ส.ค."
        Case 9
            ThaiMonthName = "ก.ย."
        Case 10
            ThaiMonthName = "ต.ค."
        Case 11
            ThaiMonthName = "พ.ย."
        Case 12
            ThaiMonthName = "ธ.ค."
    End Select
End Function

Public Function TrueFalseToYN(selVal As Boolean) As String
    If selVal Then
        TrueFalseToYN = "Y"
    Else
        TrueFalseToYN = "N"
    End If
End Function

Public Function getNumberValue(selStr As String) As String
    If IsNumeric(selStr) Then
        getNumberValue = selStr
    Else
        getNumberValue = 0
    End If
End Function
