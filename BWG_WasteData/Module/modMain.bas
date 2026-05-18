Attribute VB_Name = "modMain"
Option Explicit
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public MDB_PATH As String

Public MainURL$
Public DefaultShrNo$
Public StrGrp As String
Public strSql As String
Public sqlStr As String

Public AttacheFiles As String

Public UserName As String
Public isAkp As Boolean
Public isBWG As Boolean
Public isPST As Boolean
Public isBWC As Boolean
Public RptTypeOrd05 As Integer

Public Sub Main()
Dim tmpDate As Date
Dim tmpCol As Collection
Dim tmpS As New BWGSearchManager
Dim AttacheFiles As String

Dim oFS As FileSystemObject
Dim sSystemFolder As String
Dim sTempFolder As String
Dim sWindowsFolder As String

    App.Title = "Waste Data"
    AppVersion = App.Major & "." & App.Minor & "." & App.Revision
    
    TodayDate = FormatYMD_Thai(Day(Date), Month(Date), Year(Date), "/")
    tmpDate = DateSerial(Year(Date), Month(Date) - 3, Day(Date))
    
    Set DBManager = New AWSDBManager
       
    GetCnString
    DSN_NAME = "AKP_AWSDSN"
'    DBManager.ConnectionString = DBConString
    DBConn.ConnectionString = DBConString
    
    Set SystemConfig = New BWGSystemConfig
'    SystemConfig.GetSystemConfigFromDB
    
'    Call WasteType
    DefaultPath = AttacheFiles
    
    Set oFS = New FileSystemObject
    sSystemFolder = oFS.GetSpecialFolder(SystemFolder)
'    sTempFolder = App.Path & "\SecureLogin.dll"
    
    If oFS.FileExists(sTempFolder) = False Then
        frmLogin.Show
    Else
        If AppNewVersion(sSystemFolder, App.Path, "SecureLogin.dll") = True Then
'            MsgBox "ตรวจพบการอัพเดทโปรแกรมกรุณาปิดแล้วเปิดโปรแกรมใหม่", vbInformation
            oFS.CopyFile sTempFolder, sSystemFolder & "\SecureLogin.dll", True
        End If
        Shell sSystemFolder & "\regsvr32 /s " & sSystemFolder & "\SecureLogin.dll", vbHide
        SetLogin frmMain
    End If
    
'    ShellExecute 0, vbNullString, App.Path & "\BWGUpdate.exe", vbNullString, vbNullString, vbNormalFocus
    
'    Set tmpCol = tmpS.UserSearch("StaffCode='user'")
'    Set CurrentUser = tmpCol(1)
'    Set tmpCol = Nothing
'    Set tmpS = Nothing
    
'    frmMDIMain.Show
End Sub

Public Function ProcessLogin_Old(UserName As String, pwd As String, Optional companyID As String = "") As BWGCompanyStaff
On Error GoTo ErrD
Dim tmpUser As BWGCompanyStaff
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpPwd As String
Dim EncObj As RC4Crypt
Dim admID As String
    Set tmpRec = New ADODB.Recordset
    Set tmpRec = GetRS("Select * From vw_UserSearch Where LoginName<>'' And isActive<>'N' and isDeleted<>'Y'", adOpenForwardOnly, adLockReadOnly)
'    tmpRec.Open "Select * From vw_UserSearch Where LoginName<>'' And isActive<>'N' and isDeleted<>'Y'", DBManager.GetDBConnection, adOpenKeyset, adLockOptimistic
    If tmpRec.EOF Then
        Set ProcessLogin_Old = New BWGCompanyStaff
        tmpRec.Close

' This part will be un-remarked when complete
        Set tmpRec = GetRS("Select count(*) as RecCount from tbUserGroup Where isDeleted='N'", adOpenForwardOnly, adLockReadOnly)
'        tmpRec.Open "Select count(*) as RecCount from tbUserGroup Where isDeleted='N'", DBManager.GetDBConnection, adOpenKeyset, adLockOptimistic
        If tmpRec("RecCount") = 0 Then
            'admID = getNewProductID(modPrefix.UserGroupPrefix)
            admID = GetGUID
            DBManager.GetDBConnection.Execute "Insert into tbUserGroup (GroupID,GroupName,GroupDesc,isActive,isSystem) Values ('" & _
                                admID & "','Administrator','Administrator','Y','Y')"
            PopulateModuleList admID 'This function will depend on each project, written in modMain
            'Code will be like this--------------------------------------
                'Public Sub PopulateModuleList(AdminID As String)
                '    With DBManager.GetDBConnection
                '        .Execute "Insert into tbFormList (FormName,FormDesc) Values ('frAlert','ระบบนัดหมาย/เตือน')"
                '        .Execute "Insert into tbAccessRole (GroupID,FormName,CanAccess,CanAdd,CanUpdate,CanDelete,CanPrint) values ('" & AdminID & "','frAlert','Y','Y','Y','Y','Y')"
                'End Sub
                '---------------------------------------------------------
        End If
        Set tmpRec = Nothing
        Exit Function
    End If
    tmpRec.Close

    If Trim(companyID) = "" Then
        sqlStr = "Select * from vw_UserSearch Where LoginName='" & UserName & "' And isActive='Y' and isDeleted<>'Y'"
    Else
        sqlStr = "Select * from vw_UserSearch Where LoginName='" & UserName & "' And isActive='Y' and isDeleted<>'Y' And StaffID in (Select StaffID from tbStaffInCompany Where CompanyID='" & companyID & "')"
    End If
    Set EncObj = New RC4Crypt
    If Trim(pwd) = "" Then
        tmpPwd = ""
    Else
        tmpPwd = EncObj.EncryptString(pwd, ENC_KEY)
    End If
    'sqlStr = sqlStr & " And LoginPwd='" & Trim(tmpPwd) & "'"
    Set EncObj = Nothing
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
'    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec.EOF Then
        Set tmpUser = Nothing
    Else
        If Trim(tmpRec("LoginPwd")) = Trim(tmpPwd) Then
            Set tmpUser = New BWGCompanyStaff
            tmpUser.SetDataFromRecordset tmpRec
        Else
            Set tmpUser = Nothing
        End If
    End If
    tmpRec.Close
    Set tmpRec = Nothing
    Set ProcessLogin_Old = tmpUser
    Exit Function
ErrD:
    MsgBox Err.Description & vbCrLf & "กรุณาลองใหม่หรือติดต่อผู้พัฒนาระบบ", vbCritical, "Error"
    Exit Function
End Function

Private Sub WasteType()
Dim tmpRec As New ADODB.Recordset
Dim tmpType As BWGWasteType
    Set tmpRec = GetRS("Select * from tbWasteType Order By PriorityLevel", adOpenForwardOnly, adLockReadOnly)
'    tmpRec.Open "Select * from tbWasteType Order By PriorityLevel", DBManager.GetDBConnection, adOpenKeyset, adLockOptimistic
    Set WasteTypes = New Collection
    Do Until tmpRec.EOF
        Set tmpType = New BWGWasteType
        tmpType.SetDataFromRecordset tmpRec
        WasteTypes.Add tmpType, "'" & tmpType.WasteType & "'"
        Set tmpType = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
End Sub

Public Function CloseAllForms() As Integer
Dim x%
Dim lastForm As String
    Do Until Forms.Count = 1
        If LCase(Trim(Forms(x).Name)) <> LCase(Trim("frMain")) Then
            lastForm = Trim(Forms(x).Name)
            Unload Forms(x)
            If isFormLoaded(lastForm) Then
                CloseAllForms = 1
                Exit Function
            End If
            x = 0
        Else
            x = x + 1
        End If
    Loop
    CloseAllForms = 0
End Function

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
Dim tmpRec As ADODB.Recordset
Dim recCount As Long
    Set tmpRec = New ADODB.Recordset
    Set tmpRec = GetRS("Select count(RequestID) as rCount from tbSellHeader", adOpenForwardOnly, adLockReadOnly)
'    tmpRec.Open "Select count(RequestID) as rCount from tbSellHeader", DBManager.GetDBConnection, adOpenStatic, adLockReadOnly
    If Not tmpRec.EOF Then
        recCount = tmpRec("rCount")
    End If
    tmpRec.Close
    Set tmpRec = Nothing
    CountSellRecord = recCount
End Function

Public Sub SetLogin_Old()
'Dim SLogin As New SecureLogin.classMain
Dim str As String, tPwd As String, UserPwd As String
Dim EnCryptObj As New RC4Crypt
Dim TrustLogin As Boolean
Dim SLogin As Object
    Set SLogin = CreateObject("securelogin.classmain")
    Call GetCnString
    With SLogin
        .DBType "MSSQL"
        .dbServer strSvr
        .DBPwd strPwd
        .DBUser strUid
        .dbName strDb
        .DBTable "tbCompanyStaff"
        .UserField "LoginName"
        .PwdField "LoginPwd"
        .Login
        If .TrustLogin = False Then
            TrustLogin = False
        Else
            TrustLogin = True
            UserName = .strUser
            UserPwd = .strPwd
        End If
    End With
    SLogin.Terminate
    If TrustLogin = False Then End
    tPwd = EnCryptObj.EncryptString(UserPwd, ENC_KEY)
    Set CurrentUser = ProcessLogin(UserName, tPwd)
    If Not CurrentUser Is Nothing Then
        frmMain.Show
    End If
End Sub

Private Sub PopulateModuleList(AdminID As String)
Dim tmpArr(20) As String
Dim x%
Dim tmpStr1$, tmpStr2$
    tmpArr(0) = "frCompanyConfig,ระบบจัดการข้อมูลบริษัท"
    tmpArr(1) = "frCustomer,ระบบจัดการข้อมูลลูกค้า"
    
    With DBManager.GetDBConnection
        For x = 0 To UBound(tmpArr)
            If Trim(tmpArr(x)) <> "" Then
                tmpStr1 = Trim(Split(tmpArr(x), ",")(0))
                tmpStr2 = Trim(Split(tmpArr(x), ",")(1))
                .Execute "Insert into tbFormList (FormName,FormDesc) Values ('" & tmpStr1 & "','" & tmpStr2 & "')"
                .Execute "Insert into tbAccessRole (GroupID,FormName,CanAccess,CanAdd,CanUpdate,CanDelete,CanPrint) values ('" & AdminID & "','" & tmpStr1 & "','Y','Y','Y','Y','Y')"
            End If
        Next
    End With
End Sub

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
Dim tmpRec As New ADODB.Recordset
    Set tmpRec = GetRS("SELECT DATENAME(dw, getdate()) AS 'DayName'", adOpenForwardOnly, adLockReadOnly)
'    tmpRec.Open "SELECT DATENAME(dw, getdate()) AS 'DayName'", DBManager.GetDBConnection, adOpenStatic, adLockReadOnly
    If tmpRec.EOF = False Then CurrentDayName = LCase(Trim(tmpRec("DayName")))
    tmpRec.Close
    Set tmpRec = Nothing
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
Dim M%, Y%
Dim tmpArr
    tmpArr = Split(TodayDate, "/")
    M = Trim(tmpArr(1))
    Y = Trim(tmpArr(0))
    If M = 1 Then
        getLastMonth = Trim(CStr((Y - 1))) & "/12"
    Else
        getLastMonth = Trim(CStr(Y)) & "/" & Format(M - 1, "00")
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
Dim x%, Y%
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
    End With
End Sub

Public Sub ClearReceiptReport(selRpt As CrystalReport)
Dim x%, Y%
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
    End With
End Sub

