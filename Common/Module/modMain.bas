Attribute VB_Name = "modMain"
Option Explicit

Public Sub Main()
On Error GoTo ErrD
Dim tmpDate As Date
Dim tmpCol As Collection, tmpType As BWGWasteType
Dim RS As New ADODB.Recordset, tmpRec As New ADODB.Recordset
Dim AttacheFiles As String, tSql As String

Dim oFS As FileSystemObject
Dim sSystemFolder As String
Dim sTempFolder As String
Dim sWindowsFolder As String

    Dim iccex As tagInitCommonControlsEx
    With iccex
        .lngSize = LenB(iccex)
        .lngICC = ICC_USEREX_CLASSES
    End With
    InitCommonControlsEx iccex
   
    App.Title = "Waste Management System"
    AppVersion = App.Major & "." & App.Minor & "." & App.Revision

    TodayDate = FormatYMD_Thai(Day(Date), Month(Date), Year(Date), "/")
    tmpDate = DateSerial(Year(Date), Month(Date) - 3, Day(Date))
    
    DefaultPath = AttacheFiles
    If Trim(ComName) <> "" Then ComName = Replace(Replace(Replace(Replace(ComName, "-", ""), " ", ""), "/", ""), "\", "")
    
    Set oFS = New FileSystemObject
    sSystemFolder = oFS.GetSpecialFolder(SystemFolder)
    sTempFolder = App.Path & "\SecureLogin.dll"
    
    If oFS.FileExists(App.Path & "\OCX\DateCtl.ocx") = True Then
        If AppNewVersion(sSystemFolder, App.Path & "\OCX", "\DateCtl.ocx") = True Then
            oFS.CopyFile App.Path & "\OCX\DateCtl.ocx", sSystemFolder & "\DateCtl.ocx", True
        End If
        Shell sSystemFolder & "\regsvr32 /s " & sSystemFolder & "\DateCtl.ocx", vbHide
    End If
    
    If oFS.FileExists(App.Path & "\OCX\MD5Crypt.dll") = True Then
        If oFS.FileExists(sSystemFolder & "\MD5Crypt.dll") = False Then
'            oFS.CopyFile App.Path & "\OCX\MD5Crypt.dll", sSystemFolder & "\MD5Crypt.dll", True
'            Shell sSystemFolder & "\regsvr32 /s " & sSystemFolder & "\MD5Crypt.dll", vbHide
            Shell sSystemFolder & "\regsvr32 /s " & App.Path & "\OCX\MD5Crypt.dll", vbHide
        End If
    End If
    
    If oFS.FileExists(sTempFolder) = False Then
        frmLogin.Show
    Else
        If AppNewVersion(sSystemFolder, App.Path, "SecureLogin.dll") = True Then
            MsgBox "ตรวจพบการอัพเดทโปรแกรมกรุณาปิดแล้วเปิดโปรแกรมใหม่", vbInformation
            oFS.CopyFile sTempFolder, sSystemFolder & "\SecureLogin.dll", True
        End If
        Shell sSystemFolder & "\regsvr32 /s " & sSystemFolder & "\SecureLogin.dll", vbHide
        SetLogin frmMain
    End If
    
    If EndApplication = True Then End
    '*** Check Application Version
    '*** ย้ายไป Form Login
'    Dim i As Integer, CheckVersion As Boolean
'    Dim UserGrp As New AWSUserGroup
'    CheckVersion = SkipCheckVersion
'    If Not CurrentUser Is Nothing Then
'        For i = 1 To CurrentUser.getUserGroups.Count
'            Set UserGrp = CurrentUser.getUserGroups(i)
'            If Trim(UserGrp.Name) = "Administrator" Or Trim(UserGrp.Name) = "IT" Then
'                If SkipCheckVersion <> CheckVersion Then SkipCheckVersion = CheckVersion
'                Exit For
'            Else
'                SkipCheckVersion = False
'            End If
'        Next i
'    End If
'
'    If CurrentUser Is Nothing Then Exit Sub
'    If SkipCheckVersion = False Then
'        tSql = "SELECT ApplicationName, vApp FROM ApplicationTB WHERE ApplicationName = '" & App.EXEName & "' OR ApplicationName = '" & App.ProductName & "'"
'        Set RS = GetRS(tSql, adOpenStatic, adLockOptimistic)
'        If RS.RecordCount > 0 Then
'            If AppVersion < RS!vApp Then
'                MsgBox "โปรแกรม Version ไม่ตรงกับ Server กรุณาแจ้งฝ่าย IT เพื่ออัพเดทโปรแกรม" & vbCrLf & _
'                "App Version " & AppVersion & " = Server Version " & RS!vApp
'                End '*** End Application
'            End If
'        Else
'            MsgBox "โปรแกรมที่ท่านใช้เก่าเกินไป กรุณาแจ้งฝ่าย IT เพื่ออัพเดทโปรแกรม" & vbCrLf & _
'            "Application Name " & App.EXEName & " " & AppVersion
'            End '*** End Application
'        End If
'    End If
    '*** End check version
    
'    ShellExecute 0, vbNullString, App.Path & "\BWGUpdate.exe", vbNullString, vbNullString, vbNormalFocus
    
    Exit Sub
ErrD:
    MsgBox err.Description & vbCrLf & "เกิดข้อผิดพลาด, กรุณาลองใหม่", vbCritical, "Error"
    End
    Exit Sub
End Sub

Public Sub GenerateFile(srcFile As String, tgtFile As String)
Dim src As Byte
Dim aadByte(2) As Byte
Dim ref1 As Long
Dim ref2 As Long
Dim x As Long
    ref1 = FreeFile
    Open srcFile For Binary As ref1
    ref2 = FreeFile
    Open tgtFile For Binary As ref2
    Get ref1, , aadByte
    Do Until EOF(ref1)
        Get ref1, , src
        Put ref2, , src
    Loop
    Close ref1
    Close ref2
End Sub

Public Function CountSellRecord() As Long
On Error GoTo ErrD
Dim tmpRec As ADODB.Recordset
Dim recCount As Long
    Set tmpRec = New ADODB.Recordset
    Set tmpRec = GetRS("Select count(RequestID) as rCount from tbSellHeader", adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then recCount = 0
    If Not tmpRec.EOF Then
        recCount = tmpRec("rCount")
    End If
    tmpRec.Close
    Set tmpRec = Nothing
    CountSellRecord = recCount
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function

Function validHex(x)
Dim blnValid
    blnValid = False
    If x >= 48 And x <= 57 Then
        blnValid = True
    End If
    If x >= 65 And x <= 70 Then
        blnValid = True
    End If
    validHex = blnValid
End Function

Public Function getPriceCode(selPrice As Double) As String
Dim tmpStr As String
Dim x%
    tmpStr = Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Format(selPrice, "###0"), "1", "I"), "2", "Z"), "3", "E"), "4", "C"), "5", "B"), "6", "J"), "7", "T"), "8", "S"), "9", "D"), "0", "O")
    tmpStr = Replace(tmpStr, "II", "IM")
    tmpStr = Replace(tmpStr, "ZZ", "ZM")
    
    tmpStr = Replace(tmpStr, "ZZ", "ZM")
    tmpStr = Replace(tmpStr, "EE", "EM")
    tmpStr = Replace(tmpStr, "CC", "CM")
    tmpStr = Replace(tmpStr, "BB", "BM")
    tmpStr = Replace(tmpStr, "JJ", "JM")
    tmpStr = Replace(tmpStr, "TT", "TM")
    tmpStr = Replace(tmpStr, "SS", "SM")
    tmpStr = Replace(tmpStr, "DD", "DM")
    tmpStr = Replace(tmpStr, "OO", "OM")
    getPriceCode = Trim(tmpStr)
End Function

Public Function CurrentDayName() As String
On Error GoTo ErrD
Dim tmpRec As New ADODB.Recordset
    Set tmpRec = GetRS("SELECT DATENAME(dw, getdate()) AS 'DayName'", adOpenForwardOnly, adLockReadOnly)
    If tmpRec.EOF = False Then CurrentDayName = LCase(Trim(tmpRec("DayName")))
    tmpRec.Close
    Set tmpRec = Nothing
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function

Function RoundNumber(selNum As String) As Double
Dim tmpDigit As String, addNumber As Integer
Dim tmpStr As String
    If InStr(1, selNum, ".") > 0 Then
        tmpStr = Trim(Split(selNum, ".")(0))
    Else
        tmpStr = Trim(selNum)
    End If
    tmpDigit = Right(tmpStr, 1)
    If IsNumeric(tmpDigit) = True Then
        If CInt(tmpDigit) >= 5 Then
            addNumber = 10 - CInt(tmpDigit)
        Else
            addNumber = -1 * CInt(tmpDigit)
        End If
        RoundNumber = CDbl(tmpStr) + CDbl(addNumber)
    Else
        RoundNumber = selNum
    End If
End Function

Public Function getDueDate(StartDate As String, CreditDay As Double) As String
Dim tmpYear%
Dim tmpDate As Date
Dim newDate As Date
Dim tmpStr
    tmpDate = DateSerial(CInt(Left(StartDate, 4)), CInt(Mid(StartDate, 6, 2)), CInt(Right(StartDate, 2)))
    newDate = DateAdd("d", CreditDay, tmpDate)
    If Year(newDate) > 2500 Then
        tmpStr = Year(newDate) & "/" & Format(Month(newDate), "00") & "/" & Format(Day(newDate), "00")
    Else
        tmpStr = (Year(newDate) + 543) & "/" & Format(Month(newDate), "00") & "/" & Format(Day(newDate), "00")
    End If
    getDueDate = Trim(tmpStr)
End Function

Public Function getLastMonth() As String
Dim m%, y%
Dim tmpArr
    tmpArr = Split(TodayDate, "/")
    m = Trim(tmpArr(1))
    y = Trim(tmpArr(0))
    If m = 1 Then
        getLastMonth = Trim(CStr((y - 1))) & "/12"
    Else
        getLastMonth = Trim(CStr(y)) & "/" & Format(m - 1, "00")
    End If
End Function

Public Function FailTypeDesc(selType As String) As String
    Select Case selType
        Case "A"
            FailTypeDesc = "เป็นความผิดพลาดของลูกค้า"
        Case "B"
            FailTypeDesc = "เป็นความผิดพลาดของบริษัท"
        Case "C"
            FailTypeDesc = "เปลี่ยนกากของเสีย"
    End Select
End Function

Public Function getColorFromError(selType As String) As Long
    Select Case selType
        Case "A"
            getColorFromError = &H80FFFF
        Case "B"
            getColorFromError = &H8080FF
        Case "C"
            getColorFromError = &HE0E0E0
    End Select
End Function

Public Sub ClearInvReport(selRpt As CrystalReport)
Dim x%, y%
    With selRpt
        For x = 1 To 30
            .Formulas(20 + (x - 1) * 4) = "Col1_" & x & "=" & Chr(34) & "" & Chr(34)
            .Formulas(21 + (x - 1) * 4) = "Col2_" & x & "=" & Chr(34) & "" & Chr(34)
            .Formulas(22 + (x - 1) * 4) = "Col3_" & x & "=" & Chr(34) & "" & Chr(34)
            .Formulas(23 + (x - 1) * 4) = "Col4_" & x & "=" & Chr(34) & "" & Chr(34)
        Next
        .Formulas(0) = "InvNo=" & Chr(34) & "" & Chr(34)
        .Formulas(1) = "InvDate=" & Chr(34) & "" & Chr(34)
        .Formulas(2) = "CustName=" & Chr(34) & "" & Chr(34)
        .Formulas(3) = "CustAddr=" & Chr(34) & "" & Chr(34)
        .Formulas(4) = "TotalAmt=" & Chr(34) & "" & Chr(34)
        .Formulas(5) = "VATAmt=" & Chr(34) & "" & Chr(34)
        .Formulas(6) = "GrandTotalAmt=" & Chr(34) & "" & Chr(34)
        .Formulas(7) = "AMTWord=" & Chr(34) & "" & Chr(34)
        .Formulas(8) = "DueDay=" & Chr(34) & "" & Chr(34)
        .Formulas(9) = "DueMonth=" & Chr(34) & "" & Chr(34)
        .Formulas(10) = "DueYear=" & Chr(34) & "" & Chr(34)
        .Formulas(11) = "WTaxAmt=" & Chr(34) & "" & Chr(34)
    End With
End Sub

Public Sub ClearReceiptReport(selRpt As CrystalReport)
Dim x%, y%
    With selRpt
        For x = 1 To 30
            .Formulas(20 + (x - 1) * 4) = "Col1_" & x & "=" & Chr(34) & "" & Chr(34)
            .Formulas(21 + (x - 1) * 4) = "Col2_" & x & "=" & Chr(34) & "" & Chr(34)
            .Formulas(22 + (x - 1) * 4) = "Col3_" & x & "=" & Chr(34) & "" & Chr(34)
            .Formulas(23 + (x - 1) * 4) = "Col4_" & x & "=" & Chr(34) & "" & Chr(34)
        Next
        .Formulas(0) = "InvNo=" & Chr(34) & "" & Chr(34)
        .Formulas(1) = "InvDate=" & Chr(34) & "" & Chr(34)
        .Formulas(2) = "CustName=" & Chr(34) & "" & Chr(34)
        .Formulas(3) = "CustAddr=" & Chr(34) & "" & Chr(34)
        .Formulas(4) = "TotalAmt=" & Chr(34) & "" & Chr(34)
        .Formulas(5) = "VATAmt=" & Chr(34) & "" & Chr(34)
        .Formulas(6) = "GrandTotalAmt=" & Chr(34) & "" & Chr(34)
        .Formulas(7) = "AMTWord=" & Chr(34) & "" & Chr(34)
        .Formulas(8) = "WTaxAmt=" & Chr(34) & "" & Chr(34)
    End With
End Sub

Public Sub VisibleImage(Optional ByVal bool = True)
    frmMain.VisibleImage bool
    frmMain.ActivateForm
End Sub
