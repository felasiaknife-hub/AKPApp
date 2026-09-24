Attribute VB_Name = "modConfiguration"
Option Explicit
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public DSN_NAME As String
Public MDB_PATH As String
Public Const ENC_KEY = "39835443558332234"
Public CurrentUser As BWGCompanyStaff
Public SystemConfig As BWGSystemConfig

Public MainURL$
Public DefaultShrNo$
Public DP$

Public sqlStr As String
Public ConnectState As Boolean
Public WasteReferent As String
Public AppVersion As String
Public SkipCheckVersion As Boolean
Public ConnectionType As String
Public ConnectionPort As String
Public ConnectTimeOut As String
Public ComName As String

Public strUid As String
Public strPwd As String
Public strDb As String
Public DBStock As String
Public strSvr As String
Public CnStr As String
Public AttacheFiles As String

Public EndApplication As Boolean
Public ConnectServer As Integer

Public CurrentCulture As String

Public Type GetSystemDateTimeFunction
    SystemDate  As String
    SystemTime  As String
    SystemDateTime  As String
End Type

Public Function GetSystemConfig() As Boolean
    Call GetCnString
    Select Case ConnectServer
    Case 2
        DSN_NAME = "BWG_AWSDSN"
    Case Else
    DSN_NAME = "AKP_AWSDSN"
    End Select
'    DBConn.ConnectionString = DBConString
    If GetDBConnection = False Then GetSystemConfig = False: Exit Function
    GetSystemConfig = True
    Call SetODBC
    Call WasteType
    Call WasteCase
    Call GetSystemDateTime
    
    Set SystemConfig = New BWGSystemConfig
    SystemConfig.GetSystemConfigFromDB
End Function

Public Sub SetODBC(Optional ByVal UseStock As Boolean = False)
    DP = "C:\Windows\System32\sqlsrv32.dll"
    If SetDSN(strDb, "AKP_AWSDSN", "BWG DSN", "SQL Server", DP, strUid, strSvr, "") = False Then
        MsgBox "Set ODBC Fail, Please reopen application agian.", vbInformation, "Error"
        End
    End If
    If UseStock = False Then Exit Sub
    If SetDSN(DBStock, "AKP_STOCKDSN", "BWG STOCK DSN", "SQL Server", DP, strUid, strSvr, "") = False Then
        MsgBox "Set ODBC Fail " & DBStock & " , Please reopen application agian.", vbInformation, "Error"
        End
    End If
End Sub

Public Sub SetLogin(ByVal frm As Form)
'Dim SLogin As New SecureLogin.classMain
Dim STR As String, tPwd As String, UserPwd As String
Dim EnCryptObj As New RC4Crypt
Dim TrustLogin As Boolean
Dim SLogin As Object
Dim UserName As String

Dim oFS As FileSystemObject
Dim sSystemFolder As String
Dim sTempFolder As String
Dim sWindowsFolder As String
    
    Set oFS = New FileSystemObject
    sSystemFolder = oFS.GetSpecialFolder(SystemFolder)
    sTempFolder = App.Path & "\SecureLogin.dll"
    
    If oFS.FileExists(sTempFolder) = False Then
        frmLogin.Show
        frmMain.InitializeMenu
        frmMain.ActivateForm
        Exit Sub
    Else
        If AppNewVersion(sSystemFolder, App.Path, "SecureLogin.dll") = True Then
            MsgBox "ตรวจพบการอัพเดทโปรแกรมกรุณาปิดแล้วเปิดโปรแกรมใหม่", vbInformation
            oFS.CopyFile sTempFolder, sSystemFolder & "\SecureLogin.dll", True
        End If
        Shell sSystemFolder & "\regsvr32 /s " & sSystemFolder & "\SecureLogin.dll", vbHide
'        SetLogin frmMain
    
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
    End If

    tPwd = EnCryptObj.EncryptString(UserPwd, ENC_KEY)
    Set CurrentUser = ProcessLogin(UserName, tPwd)
    If Not CurrentUser Is Nothing Then
        frm.Show
        frmMain.InitializeMenu
        frmMain.ActivateForm
    End If
End Sub

Public Sub GetCnString()
    Dim ConfigFile As String
    Dim oFS As FileSystemObject
    Dim ret As String
    Dim Cnn As Long
    Set oFS = New FileSystemObject
        
        ConfigFile = App.Path & "\DBConfig.ini"
        If oFS.FileExists(ConfigFile) = False Then MsgBox "ไม่พบไฟล์ dbconfig.ini", vbCritical: End
        
        ret = String(255, 0)
        Cnn = GetPrivateProfileString("CTRUSER", "UserID", "", ret, 255, ConfigFile)
        strUid = Mid(ret, 1, Cnn)
        ret = String(255, 0)
        Cnn = GetPrivateProfileString("CTRUSER", "Password", "", ret, 255, ConfigFile)
        strPwd = Mid(ret, 1, Cnn)
        'StrPws = DecryDat(StrPws)
        ret = String(255, 0)
        Cnn = GetPrivateProfileString("CTRUSER", "DBVolumn", "", ret, 255, ConfigFile)
        strDb = Mid(ret, 1, Cnn)
        ret = String(255, 0)
        Cnn = GetPrivateProfileString("CTRUSER", "DBStock", "", ret, 255, ConfigFile)
        DBStock = Mid(ret, 1, Cnn)
        ret = String(255, 0)
        Cnn = GetPrivateProfileString("CTRUSER", "ServerName", "", ret, 255, ConfigFile)
        strSvr = Mid(ret, 1, Cnn)
                
        Select Case ConnectServer
            Case 0
                strSvr = "58.82.140.99"
          Case 1
                 strSvr = "58.82.140.99"
         Case 2
                strSvr = "223.27.196.35"
                    ret = String(255, 0)
                    Cnn = GetPrivateProfileString("CTRUSER", "PasswordBWG", "", ret, 255, ConfigFile)
                    strPwd = Mid(ret, 1, Cnn)
            'Case 4
             '   strSvr = "(local)"
        End Select
            
        ret = String(255, 0)
        Cnn = GetPrivateProfileString("CTRUSER", "ConnectTimeOut", "", ret, 255, ConfigFile)
        ConnectTimeOut = IIf(Mid(ret, 1, Cnn) = "", "", Mid(ret, 1, Cnn))
        
        ret = String(255, 0)
        Cnn = GetPrivateProfileString("CTRUSER", "Attachefiles", "", ret, 255, ConfigFile)
        AttacheFiles = "\\" & strSvr & "\" & Mid(ret, 1, Cnn)
        
        ret = String(255, 0)
        Cnn = GetPrivateProfileString("VAR", "ORD05", "", ret, 255, ConfigFile)
        RptTypeOrd05 = Mid(ret, 1, Cnn)
        
        ret = String(255, 0)
        Cnn = GetPrivateProfileString("OPTION", "SkipCheckVersion", "", ret, 255, ConfigFile)
        SkipCheckVersion = IIf(Mid(ret, 1, Cnn) = "", False, Mid(ret, 1, Cnn))
        
        ret = String(255, 0)
        Cnn = GetPrivateProfileString("OPTION", "ConnectionType", "", ret, 255, ConfigFile)
        ConnectionType = IIf(Mid(ret, 1, Cnn) = "", False, Mid(ret, 1, Cnn))
        
        CnStr = "Provider=SQLOLEDB.1;Persist Security Info=True;User ID=" & strUid & ";Password=" & strPwd & ";Initial Catalog=" & strDb & ";Data Source=" & strSvr & IIf(Trim(ConnectTimeOut) <> "", "; Connect TimeOut = " & ConnectTimeOut, "")
        DBConString = "Provider=SQLOLEDB.1;Persist Security Info=True;Data Source=" & strSvr & ";Initial Catalog=" & strDb & ";User ID=" & strUid & ";Password=" & strPwd & IIf(Trim(ConnectTimeOut) <> "", "; Connect TimeOut = " & ConnectTimeOut, "")
        DBConStock = "Provider=SQLOLEDB.1;Persist Security Info=True;Data Source=" & strSvr & ";Initial Catalog=" & DBStock & ";User ID=" & strUid & ";Password=" & strPwd & IIf(Trim(ConnectTimeOut) <> "", "; Connect TimeOut = " & ConnectTimeOut, "")
End Sub

Public Function CloseAllForms() As Integer
Dim x%
Dim lastForm As String
    Do Until Forms.Count = 1
        If LCase(Trim(Forms(x).Name)) <> LCase(Trim("frmMain")) Then
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

Public Function ProcessLogin(UserName As String, pwd As String, Optional companyID As String = "") As BWGCompanyStaff
On Error GoTo ErrD
Dim tmpUser As BWGCompanyStaff
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpPwd As String
Dim EncObj As New RC4Crypt
Dim objMD5Encrypt As New MD5Crypt.EncryptionTools
Dim admID As String
    sqlStr = "SELECT * FROM vw_UserSearch WHERE LoginName<>'' And isActive<>'N' and isDeleted<>'Y'"
    Set tmpRec = New ADODB.Recordset
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec.EOF Then
        Set ProcessLogin = New BWGCompanyStaff
        tmpRec.Close

' This part will be un-remarked when complete
        sqlStr = "Select count(*) as RecCount from tbUserGroup WHERE isDeleted='N'"
        Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
        If tmpRec("RecCount") = 0 Then
            'admID = getNewProductID(modPrefix.UserGroupPrefix)
            admID = GetGUID
            DBConnExc "Insert into tbUserGroup (GroupID,GroupName,GroupDesc,isActive,isSystem) Values ('" & _
                                admID & "','Administrator','Administrator','Y','Y')"
            PopulateModuleList admID 'This function will depend on each project, written in modMain
            'Code will be like this--------------------------------------
                'Public Sub PopulateModuleList(AdminID As String)
                '    With DBConn
                '        .Execute "Insert into tbFormList (FormName,FormDesc) Values ('frmAlert','ระบบนัดหมาย/เตือน')"
                '        .Execute "Insert into tbAccessRole (GroupID,FormName,CanAccess,CanAdd,CanUpdate,CanDelete,CanPrint) values ('" & AdminID & "','frmAlert','Y','Y','Y','Y','Y')"
                'End Sub
                '---------------------------------------------------------
        End If
        Set tmpRec = Nothing
        Exit Function
    End If
    tmpRec.Close

    If Trim(companyID) = "" Then
        sqlStr = "SELECT * FROM vw_UserSearch WHERE LoginName='" & UserName & "' And isActive='Y' and isDeleted<>'Y'"
    Else
        sqlStr = "SELECT * FROM vw_UserSearch WHERE LoginName='" & UserName & "' And isActive='Y' and isDeleted<>'Y' And StaffID in (Select StaffID from tbStaffInCompany WHERE CompanyID='" & companyID & "')"
    End If
    Set EncObj = New RC4Crypt
    Set objMD5Encrypt = New MD5Crypt.EncryptionTools
    If Trim(pwd) = "" Then
        tmpPwd = ""
    Else
        tmpPwd = EncObj.EncryptString(pwd, ENC_KEY)
        tmpPwd = objMD5Encrypt.MD5HashString(pwd)
    End If
    'sqlStr = sqlStr & " And LoginPwd='" & Trim(tmpPwd) & "'"
    Set EncObj = Nothing
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec.EOF Then
        Set tmpUser = Nothing
    Else
        'If Trim(tmpRec("LoginPwd")) = Trim(tmpPwd) Then
        If Trim(tmpRec("MD5Pwd")) = Trim(tmpPwd) Then
            Set tmpUser = New BWGCompanyStaff
            tmpUser.SetDataFromRecordset tmpRec
        Else
            Set tmpUser = Nothing
        End If
    End If
    tmpRec.Close
    Set tmpRec = Nothing
    Set ProcessLogin = tmpUser
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function

Private Sub PopulateModuleList(AdminID As String)
Dim tmpArr(20) As String
Dim x%
Dim tmpStr1$, tmpStr2$
    tmpArr(0) = "frmCompanyConfig,ระบบจัดการข้อมูลบริษัท"
    tmpArr(1) = "frmCustomer,ระบบจัดการข้อมูลลูกค้า"
    
        For x = 0 To UBound(tmpArr)
            If Trim(tmpArr(x)) <> "" Then
                tmpStr1 = Trim(Split(tmpArr(x), ",")(0))
                tmpStr2 = Trim(Split(tmpArr(x), ",")(1))
                DBConnExc "Insert into tbFormList (FormName,FormDesc) Values ('" & tmpStr1 & "','" & tmpStr2 & "')"
                DBConnExc "Insert into tbAccessRole (GroupID,FormName,CanAccess,CanAdd,CanUpdate,CanDelete,CanPrint) values ('" & AdminID & "','" & tmpStr1 & "','Y','Y','Y','Y','Y')"
            End If
        Next
        
End Sub

Public Function SelectWithCompanyID(ByVal strSql As String, Optional ByVal AddCondition As Boolean = True) As String
    Dim companyID As String, Sql As String
    Dim x As Integer
    SelectWithCompanyID = ""
    companyID = "("
    For x = 1 To CurrentUser.WorkCompanies.Count
        companyID = companyID & "'" & Trim(CurrentUser.WorkCompanies(x).ID) & "',"
    Next
    companyID = Mid(companyID, 1, Len(companyID) - 1)
    companyID = companyID & ")"
    
    If AddCondition = False Then
        If InStr(1, UCase(strSql), "WHERE") > 0 Then
            Sql = strSql & vbCrLf & "AND CompanyID IN " & companyID
        Else
            Sql = strSql & " WHERE CompanyID IN " & companyID
        End If
    Else
        Sql = strSql & vbCrLf & "AND CompanyID IN " & companyID
    End If
    SelectWithCompanyID = Sql
End Function

Public Function GetSystemDateTime() As GetSystemDateTimeFunction
    Dim rs As New ADODB.Recordset
    Dim tSql As String
    If DBConn.State = adStateClosed Then GetDBConnection
    tSql = "SELECT CONVERT(VARCHAR(10), GETDATE(), 111) AS SystemDate, CONVERT(VARCHAR(8), GETDATE(), 108) AS SystemTime, GETDATE() As SystemDateTime"
    rs.Open tSql, DBConn, adOpenForwardOnly, adLockOptimistic
    GetSystemDateTime.SystemDate = rs!SystemDate
    GetSystemDateTime.SystemTime = rs!SystemTime
    GetSystemDateTime.SystemDateTime = rs!SystemDateTime
End Function

Public Function GetSystemDateTimeThai() As GetSystemDateTimeFunction
    Dim rs As New ADODB.Recordset
    Dim tSql As String
    If DBConn.State = adStateClosed Then GetDBConnection
    tSql = "SELECT CONVERT(varchar(6), GETDATE(), 103) + CONVERT(varchar(4), YEAR(GETDATE()) + 543) AS SystemDate, CONVERT(VARCHAR(8), GETDATE(), 108) AS SystemTime" & vbCrLf
    tSql = tSql & ", CONVERT(varchar(6), GETDATE(), 103) + CONVERT(varchar(4), YEAR(GETDATE()) + 543) + ' ' + CONVERT(VARCHAR(8), GETDATE(), 108) AS SystemDateTime"
    rs.Open tSql, DBConn, adOpenForwardOnly, adLockOptimistic
    GetSystemDateTimeThai.SystemDate = Format(rs!SystemDate, "yyyy/MM/dd")
    GetSystemDateTimeThai.SystemTime = rs!SystemTime
    GetSystemDateTimeThai.SystemDateTime = Format(rs!SystemDateTime, "yyyy/MM/dd HH:mm:ss.nnn")
End Function

