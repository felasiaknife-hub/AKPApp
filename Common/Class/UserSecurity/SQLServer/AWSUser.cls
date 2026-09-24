VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AWSUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim userID_ As String
Dim fName_ As String
Dim lName_ As String
Dim nickName_ As String
Dim addr1_ As String
Dim addr2_ As String
Dim addr3_ As String
Dim phoneNo_ As String
Dim loginName_ As String
Dim passWord_ As String
Dim isActive_ As String
Dim colGroup_ As Collection
Dim isChanged_ As Boolean
Public Enum AWSRoleType
    AccessRole = 1
    AddNewRole = 2
    UpdateRole = 3
    DeleteRole = 4
    PrintRole = 5
End Enum

Public Property Get ID() As String
    ID = userID_
End Property
Public Property Let ID(ByVal newVal As String)
    userID_ = Trim(newVal)
End Property

Public Property Get FirstName() As String
    FirstName = ReturnSgQuote(fName_)
End Property
Public Property Let FirstName(ByVal newVal As String)
Dim tmpStr As String
    tmpStr = ClearSgQuote(newVal)
    If Trim(fName_) <> Trim(tmpStr) Then
        fName_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property

Public Property Get LastName() As String
    LastName = ReturnSgQuote(lName_)
End Property
Public Property Let LastName(ByVal newVal As String)
Dim tmpStr As String
    tmpStr = ClearSgQuote(newVal)
    If Trim(lName_) <> Trim(tmpStr) Then
        lName_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property

Public Property Get NickName() As String
    NickName = ReturnSgQuote(nickName_)
End Property
Public Property Let NickName(ByVal newVal As String)
Dim tmpStr As String
    tmpStr = ClearSgQuote(newVal)
    If Trim(nickName_) <> Trim(tmpStr) Then
        nickName_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property

Public Property Get Address1() As String
    Address1 = ReturnSgQuote(addr1_)
End Property
Public Property Let Address1(ByVal newVal As String)
Dim tmpStr As String
    tmpStr = ClearSgQuote(newVal)
    If Trim(addr1_) <> Trim(tmpStr) Then
        addr1_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property

Public Property Get Address2() As String
    Address2 = addr2_
End Property
Public Property Let Address2(ByVal newVal As String)
Dim tmpStr As String
    tmpStr = ClearSgQuote(newVal)
    If Trim(addr2_) <> Trim(tmpStr) Then
        addr2_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property

Public Property Get Address3() As String
    Address3 = addr3_
End Property
Public Property Let Address3(ByVal newVal As String)
Dim tmpStr As String
    tmpStr = ClearSgQuote(newVal)
    If Trim(addr3_) <> Trim(tmpStr) Then
        addr3_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property

Public Property Get PhoneNumber() As String
    PhoneNumber = phoneNo_
End Property
Public Property Let PhoneNumber(ByVal newVal As String)
Dim tmpStr As String
    tmpStr = ClearSgQuote(newVal)
    If Trim(phoneNo_) <> Trim(tmpStr) Then
        phoneNo_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property

Public Property Get LoginName() As String
    LoginName = loginName_
End Property
Public Property Let LoginName(ByVal newVal As String)
Dim tmpStr As String
    tmpStr = ClearSgQuote(newVal)
    If Trim(loginName_) <> Trim(tmpStr) Then
        loginName_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property

Public Property Get Password() As String
    Password = passWord_
End Property
Public Property Let Password(ByVal newVal As String)
Dim tmpStr As String
    tmpStr = ClearSgQuote(newVal)
    If Trim(passWord_) <> Trim(tmpStr) Then
        passWord_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property

Public Property Get isActive() As Boolean
    If Trim(isActive_) = "Y" Then
        isActive = True
    Else
        isActive = False
    End If
End Property
Public Property Let isActive(ByVal newVal As Boolean)
Dim tmpStr As String
    If newVal Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(isActive_) <> Trim(tmpStr) Then
        isActive_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property

Public Property Get isChanged() As Boolean
    isChanged = isChanged_
End Property
Public Property Let isChanged(ByVal newVal As Boolean)
    isChanged_ = newVal
End Property

Public Function Save()
Dim recCount%, x%
Dim tmpCon As ADODB.Connection
Dim encPwd As String
Dim encObj As RC4Crypt
Dim tmpRec As ADODB.Recordset
    Set encObj = New RC4Crypt
    encPwd = encObj.EncryptString(Trim(passWord_), ENC_KEY)
    Set encObj = Nothing
    Set tmpCon = DBManager.GetDBConnection
    Set tmpRec = New ADODB.Recordset
    tmpRec.Open "Select * From tbStaff Where StaffID='" & userID_ & "'", tmpCon, adOpenDynamic, adLockOptimistic
    If tmpRec.EOF Then
        tmpRec.AddNew
        tmpRec("StaffID") = userID_
        tmpRec("LoginName") = loginName_
        tmpRec("LoginPwd") = encPwd
    End If
    tmpRec("FName") = fName_
    tmpRec("LName") = lName_
    tmpRec("NickName") = nickName_
    tmpRec("Addr1") = addr1_
    tmpRec("Addr2") = addr2_
    tmpRec("Addr3") = addr3_
    tmpRec("PhoneNo") = phoneNo_
    tmpRec("isActive") = isActive_
    tmpRec.Update
    tmpRec.Close
    If Not colGroup_ Is Nothing Then
        tmpCon.Execute "Delete tbUserInGroup Where StaffID='" & userID_ & "'"
        For x = 1 To colGroup_.Count
            tmpCon.Execute "Insert into tbUserInGroup (GroupID,StaffID) Values ('" & colGroup_(x).ID & "','" & userID_ & "')"
        Next
    End If
    Set tmpRec = Nothing
    Set tmpCon = Nothing
    isChanged_ = False
End Function

Public Function Delete()
    DBManager.GetDBConnection.Execute "Delete tbUserInGroup Where StaffID='" & userID_ & "'"
    DBManager.GetDBConnection.Execute "Delete tbStaff Where StaffID='" & userID_ & "'"
End Function

Public Function ChangePassword(newPwd As String)
Dim encObj As RC4Crypt
Dim encPwd As String
Dim tmpRec As ADODB.Recordset
    Set encObj = New RC4Crypt
    encPwd = encObj.EncryptString(Trim(newPwd), ENC_KEY)
    Set encObj = Nothing
    
    Set tmpRec = New ADODB.Recordset
    tmpRec.Open "Select * From tbStaff Where StaffID='" & userID_ & "'", DBManager.GetDBConnection, adOpenDynamic, adLockOptimistic
    If tmpRec.EOF = False Then
        tmpRec("LoginPwd") = encPwd
        tmpRec.Update
    End If
    tmpRec.Close
    Set tmpRec = Nothing
End Function

Public Sub setEncPWD(encPwd As String)
Dim encObj As RC4Crypt
    Set encObj = New RC4Crypt
    Password = encObj.DecryptString(Trim(encPwd), ENC_KEY)
    Set encObj = Nothing
End Sub

Public Sub SetDataFromRecordset(selRec As ADODB.Recordset)
    userID_ = Trim("" & selRec("StaffID"))
    fName_ = Trim("" & selRec("FName"))
    lName_ = Trim("" & selRec("LName"))
    nickName_ = Trim("" & selRec("NickName"))
    addr1_ = Trim("" & selRec("Addr1"))
    addr2_ = Trim("" & selRec("Addr2"))
    addr3_ = Trim("" & selRec("Addr3"))
    phoneNo_ = Trim("" & selRec("PhoneNo"))
    loginName_ = Trim("" & selRec("LoginName"))
    passWord_ = Trim("" & selRec("LoginPwd"))
'    isAdmin_ = Trim("" & selRec("isAdmin"))
'    isTech_ = Trim("" & selRec("isTechnic"))
'    isRecp_ = Trim("" & selRec("isRecp"))
'    isSale_ = Trim("" & selRec("isSale"))
    isActive_ = Trim("" & selRec("isActive"))
End Sub

Public Function DuplicateUser() As AWSUser
Dim tmpUser As AWSUser
    Set tmpUser = New AWSUser
    With tmpUser
        .ID = ID
        .FirstName = FirstName
        .LastName = LastName
        .NickName = NickName
        .Address1 = Address1
        .Address2 = Address2
        .Address3 = Address3
        .PhoneNumber = PhoneNumber
        .LoginName = LoginName
        .Password = Password
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

Private Sub Class_Initialize()
    isActive_ = "Y"
End Sub

Public Function getUserGroups() As Collection
Dim tmpRec As ADODB.Recordset
Dim tmpGroup As AWSUserGroup
    If colGroup_ Is Nothing Then
        Set colGroup_ = New Collection
    End If
    If colGroup_.Count = 0 Then
        Set tmpRec = New ADODB.Recordset
        tmpRec.Open "Select G.* From tbUserInGroup N, tbUserGroup G Where G.isDeleted='N' And N.GroupID=G.GroupID And N.StaffID='" & userID_ & "'", DBManager.GetDBConnection, adOpenStatic, adLockReadOnly
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
End Function

Public Function VerifyAccessRole(frmName As String, roleType As AWSRoleType) As Boolean
Dim x%
Dim tmpRole As AWSAccessRole
Dim tmpGroup As AWSUserGroup
    VerifyAccessRole = False
    If colGroup_ Is Nothing Then
        getUserGroups
    End If
    For x = 1 To colGroup_.Count
        Set tmpGroup = colGroup_(x)
        If isExist(tmpGroup.getAccessRoles, "'" & frmName & "'") Then
            Set tmpRole = tmpGroup.getAccessRoles("'" & frmName & "'")
            If (tmpRole.CanAccess And roleType = AccessRole) Or _
                (tmpRole.CanAddNew And roleType = AddNewRole) Or _
                (tmpRole.CanUpdate And roleType = UpdateRole) Or _
                (tmpRole.CanDelete And roleType = DeleteRole) Or _
                (tmpRole.CanPrint And roleType = PrintRole) Then
                VerifyAccessRole = True
                Exit For
            End If
        End If
    Next
    Set tmpRole = Nothing
    Set tmpGroup = Nothing
End Function

Public Sub RefreshUserGroup()
    Set colGroup_ = Nothing
End Sub
