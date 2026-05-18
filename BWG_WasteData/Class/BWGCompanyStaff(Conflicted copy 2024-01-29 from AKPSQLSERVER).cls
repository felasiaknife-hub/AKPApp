VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BWGCompanyStaff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim id_ As String
Dim companyID_ As String
Dim companyName_ As String
Dim staffCode_ As String
Dim staffFName_ As String
Dim staffLName_ As String
Dim nickName_ As String
Dim positionName_ As String
Dim telNo_ As String
Dim emailAddr_ As String
Dim loginName_ As String
Dim loginPwd_ As String
Dim isActive_ As String
Dim staffRem_ As String
Dim staffPic_
Dim addr1_ As String
Dim addr2_ As String
Dim addr3_ As String
Dim departmentID_ As String
Dim departmentName_ As String
Dim titleName_ As String
Dim isChanged_ As Boolean
'Security
Dim colGroup_ As Collection
Dim colWorkCompany_  As Collection
Dim colWasteCompany_ As Collection
Dim colUserCompany_ As Collection
Public Enum AWSRoleType
    AccessRole = 1
    AddNewRole = 2
    UpdateRole = 3
    DeleteRole = 4
    PrintRole = 5
    ApproveRole = 6
End Enum

Public Property Get ID() As String
    ID = id_
End Property
Public Property Let ID(ByVal newval As String)
    id_ = Trim(newval)
End Property

Public Property Get companyID() As String
    companyID = companyID_
End Property
Public Property Let companyID(ByVal newval As String)
    If Trim(companyID_) <> Trim(newval) Then
        companyID_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get TitleName() As String
    TitleName = titleName_
End Property
Public Property Let TitleName(ByVal newval As String)
    If Trim(titleName_) <> Trim(newval) Then
        titleName_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get CompanyName() As String
    CompanyName = companyName_
End Property
Public Property Let CompanyName(ByVal newval As String)
    If Trim(companyName_) <> Trim(newval) Then
        companyName_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get StaffCode() As String
    StaffCode = staffCode_
End Property
Public Property Let StaffCode(ByVal newval As String)
    If Trim(staffCode_) <> Trim(newval) Then
        staffCode_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get StaffFName() As String
    StaffFName = staffFName_
End Property
Public Property Let StaffFName(ByVal newval As String)
    If Trim(staffFName_) <> Trim(newval) Then
        staffFName_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get StaffLName() As String
    StaffLName = staffLName_
End Property
Public Property Let StaffLName(ByVal newval As String)
    If Trim(staffLName_) <> Trim(newval) Then
        staffLName_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get StaffNickName() As String
    StaffNickName = nickName_
End Property
Public Property Let StaffNickName(ByVal newval As String)
    If Trim(nickName_) <> Trim(newval) Then
        nickName_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get PositionName() As String
    PositionName = positionName_
End Property
Public Property Let PositionName(ByVal newval As String)
    If Trim(positionName_) <> Trim(newval) Then
        positionName_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get EmailAddr() As String
    EmailAddr = emailAddr_
End Property
Public Property Let EmailAddr(ByVal newval As String)
    If Trim(emailAddr_) <> Trim(newval) Then
        emailAddr_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get TelNo() As String
    TelNo = telNo_
End Property
Public Property Let TelNo(ByVal newval As String)
    If Trim(telNo_) <> Trim(newval) Then
        telNo_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get LoginName() As String
    LoginName = loginName_
End Property
Public Property Let LoginName(ByVal newval As String)
    If Trim(loginName_) <> Trim(newval) Then
        loginName_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get LoginPwd() As String
    LoginPwd = loginPwd_
End Property
Public Property Let LoginPwd(ByVal newval As String)
    If Trim(loginPwd_) <> Trim(newval) Then
        loginPwd_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get isActive() As Boolean
    If Trim(isActive_) = "" Then isActive_ = "Y"
    If isActive_ = "Y" Then
        isActive = True
    Else
        isActive = False
    End If
End Property
Public Property Let isActive(ByVal newval As Boolean)
Dim tmpStr As String
    If newval Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(isActive_) <> Trim(tmpStr) Then
        isActive_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property

Public Property Get StaffRemark() As String
    StaffRemark = staffRem_
End Property
Public Property Let StaffRemark(ByVal newval As String)
    If Trim(staffRem_) <> Trim(newval) Then
        staffRem_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get Addr1() As String
    Addr1 = addr1_
End Property
Public Property Let Addr1(ByVal newval As String)
    If Trim(addr1_) <> Trim(newval) Then
        addr1_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get Addr2() As String
    Addr2 = addr2_
End Property
Public Property Let Addr2(ByVal newval As String)
    If Trim(addr2_) <> Trim(newval) Then
        addr2_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get Addr3() As String
    Addr3 = addr3_
End Property
Public Property Let Addr3(ByVal newval As String)
    If Trim(addr3_) <> Trim(newval) Then
        addr3_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get DepartmentID() As String
    DepartmentID = departmentID_
End Property
Public Property Let DepartmentID(ByVal newval As String)
    If Trim(departmentID_) <> Trim(newval) Then
        departmentID_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get DepartmentName() As String
    DepartmentName = departmentName_
End Property
Public Property Let DepartmentName(ByVal newval As String)
    If Trim(departmentName_) <> Trim(newval) Then
        departmentName_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get Picture()
    Picture = staffPic_
End Property
Public Property Let Picture(newval)
    staffPic_ = newval
    isChanged_ = True
End Property

Public Property Get isChanged() As Boolean
    isChanged = isChanged_
End Property
Public Property Let isChanged(ByVal newval As Boolean)
    isChanged_ = newval
End Property

Public Sub Save()
End Sub

Public Function Delete() As String
On Error GoTo delErr:
    DBManager.GetDBConnection.Execute "Update tbCompanyStaff Set isDeleted='Y',DeletedDate='" & TodayDate & "' Where StaffID='" & id_ & "'"
    Delete = ""
    Exit Function
delErr:
    Delete = Err.Description
End Function

Public Sub SetDataFromRecordset(selRec As ADODB.Recordset)
    id_ = selRec("StaffID")
'    companyID_ = Trim("" & selRec("CompanyID"))
'    companyName_ = Trim("" & selRec("CompanyNameTH"))
    staffCode_ = Trim("" & selRec("StaffCode"))
    staffFName_ = Trim("" & selRec("StaffFName"))
    staffLName_ = Trim("" & selRec("StaffLName"))
    nickName_ = Trim("" & selRec("NickName"))
    positionName_ = Trim("" & selRec("PositionName"))
    telNo_ = Trim("" & selRec("TelNo"))
    emailAddr_ = Trim("" & selRec("EmailAddr"))
    loginName_ = Trim("" & selRec("LoginName"))
    loginPwd_ = Trim("" & selRec("LoginPwd"))
    isActive_ = Trim("" & selRec("isActive"))
    staffRem_ = Trim("" & selRec("StaffRem"))
    staffPic_ = Trim("" & selRec("StaffPic"))
    addr1_ = Trim("" & selRec("Addr1"))
    addr2_ = Trim("" & selRec("Addr2"))
    addr3_ = Trim("" & selRec("Addr3"))
    titleName_ = Trim("" & selRec("TitleDesc"))
    departmentID_ = Trim("" & selRec("DepartmentID"))
'    departmentName_ = Trim("" & selRec("DeptName"))
End Sub

Public Function VerifyAccessRole(frmName As String, roleType As AWSRoleType) As String
End Function

Public Function getUserGroups() As Collection
On Error GoTo ErrD
Dim tmpRec As ADODB.Recordset
Dim tmpGroup As AWSUserGroup
    If colGroup_ Is Nothing Then
        Set colGroup_ = New Collection
    End If
    If colGroup_.Count = 0 Then
        Set tmpRec = New ADODB.Recordset
        Set tmpRec = GetRS("Select G.* From tbUserInGroup N, tbUserGroup G Where G.isDeleted='N' And N.GroupID=G.GroupID And N.StaffID='" & id_ & "'", adOpenForwardOnly, adLockReadOnly)
'        tmpRec.Open "Select G.* From tbUserInGroup N, tbUserGroup G Where G.isDeleted='N' And N.GroupID=G.GroupID And N.StaffID='" & id_ & "'", DBManager.GetDBConnection, adOpenKeyset, adLockOptimistic
        Do Until tmpRec.EOF
            Set tmpGroup = New AWSUserGroup
            tmpGroup.SetDataFromRecordset tmpRec
            colGroup_.Add tmpGroup, "'" & tmpGroup.ID & "'"
            Set tmpGroup = Nothing
            tmpRec.MoveNext
        Loop
        tmpRec.Close
        Set tmpRec = Nothing
    End If
    Set getUserGroups = colGroup_
    Exit Function
ErrD:
    MsgBox Err.Description & vbCrLf & "กรุณาลองใหม่หรือติดต่อผู้พัฒนาระบบ", vbCritical, "Error"
    Exit Function
End Function

Public Property Get WorkCompanies() As Collection
Dim tmpS As New BWGSearchManager
    If colWorkCompany_ Is Nothing Then
        Set colWorkCompany_ = tmpS.CompanySearch("CompanyID in (Select CompanyID from tbStaffInCompany Where StaffID='" & id_ & "')")
    End If
    Set tmpS = Nothing
    Set WorkCompanies = colWorkCompany_
End Property
Public Property Set WorkCompanies(newval As Collection)
    Set colWorkCompany_ = newval
    isChanged_ = True
End Property

Public Property Get WasteCompany() As Collection
Dim tmpS As New BWGSearchManager
    If colWasteCompany_ Is Nothing Then
        'Set colWasteCompany_ = tmpS.CompanySearch("CompanyID in ('776EDA5830F744A6B3551348D851BC22', '89F23B57E7DC4407B95953A38AC5B537','C451AF07C9A44B4282D6E08B14AC0195')")
        Set colWasteCompany_ = tmpS.CompanySearch("isDeleted <>'Y'")
    End If
    Set tmpS = Nothing
    Set WasteCompany = colWasteCompany_
End Property
Public Property Set WasteCompany(newval As Collection)
    Set colWasteCompany_ = newval
    isChanged_ = True
End Property

Public Sub RefreshUserGroup()
    Set colGroup_ = Nothing
End Sub

Public Function DuplicateUser() As BWGCompanyStaff
Dim tmpUser As BWGCompanyStaff
    Set tmpUser = New BWGCompanyStaff
    With tmpUser
        .ID = ID
        .StaffFName = StaffFName
        .StaffLName = StaffLName
        .StaffNickName = StaffNickName
        .Addr1 = Addr1
        .Addr2 = Addr2
        .Addr3 = Addr3
        .TelNo = TelNo
        .LoginName = LoginName
        .LoginPwd = LoginPwd
        .StaffCode = StaffCode
        .TitleName = TitleName
'        .isAdmin = isAdmin
'        .isTechnician = isTechnician
'        .isReception = isReception
'        .isSale = isSale
        .isActive = isActive
        .isChanged = isChanged
    End With
    Set DuplicateUser = tmpUser
    Set tmpUser = Nothing
End Function

Public Function ChangePassword(newPwd As String)
End Function

Public Function getUserCompany() As Collection
On Error GoTo ErrD
Dim tmpRec As ADODB.Recordset
Dim tmpUserCompany As BWGCompanyConfig
    If colUserCompany_ Is Nothing Then
        Set colUserCompany_ = New Collection
    End If
    If colUserCompany_.Count = 0 Then
        Set tmpRec = New ADODB.Recordset
        Set tmpRec = GetRS("SELECT     C.*, S.StaffID, T.CompanyName AS TSDFName FROM dbo.tbCompanyConfig AS C INNER JOIN dbo.tbStaffInCompany AS S ON C.CompanyID = S.CompanyID LEFT OUTER JOIN dbo.tbTSDFCompany AS T ON C.TSDFID = T.TSDFID WHERE StaffID='" & id_ & "' ORDER BY CompanyNameTH", adOpenForwardOnly, adLockReadOnly)
        Do Until tmpRec.EOF
            Set tmpUserCompany = New BWGCompanyConfig
            tmpUserCompany.SetDataFromRecordset tmpRec
            colUserCompany_.Add tmpUserCompany, "'" & tmpUserCompany.ID & "'"
            Set tmpUserCompany = Nothing
            tmpRec.MoveNext
        Loop
        tmpRec.Close
        Set tmpRec = Nothing
    End If
    Set getUserCompany = colUserCompany_
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function
