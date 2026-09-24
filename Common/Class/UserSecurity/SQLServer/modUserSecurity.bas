Attribute VB_Name = "modUserSecurity"
Option Explicit
'Public CurrentUser As scmteacher
Public ColFormList As Collection
Public ColUserGroup As Collection
'Copy these lines to modPrefix of each project
'Public Const UserPrefix As String = "UR"
'Public Const UserGroupPrefix As String = "GR"
'Public Const RoleAccessPrefix As String = "RA"

'Public Function ProcessLogin(userName As String, pwd As String) As scmteacher
'Dim tmpUser As scmteacher
'Dim tmpRec As ADODB.Recordset
'Dim sqlStr As String
'Dim tmpPwd As String
'Dim encObj As RC4Crypt
'Dim tmpCon As ADODB.Recordset
'Dim admID As String
'    Set tmpCon = DBManager
'    Set tmpRec = New ADODB.Recordset
'    tmpRec.Open "Select * From tbUser"', tmpCon, adOpenStatic, adLockReadOnly
'    If tmpRec.EOF Then
'        Set ProcessLogin = New scmteacher
'        'tmpRec.Close
'        tmpRec.Open "Select count(*) as RecCount from tbUserGroup"', tmpCon, adOpenStatic, adLockReadOnly
'        If tmpRec("RecCount") = 0 Then
'            admID = getNewProductID(modPrefix.UserGroupPrefix)
'            tmpCon.Execute "Insert into tbUserGroup (GroupID,GroupName,GroupDesc,isActive,isSystem) Values ('" & _
'                                admID & "','Administrator','Administrator','Y','Y')"
'            PopulateModuleList admID 'This function will depend on each project, written in modMain
'            'Code will be like this--------------------------------------
'                'Public Sub PopulateModuleList(AdminID As String)
'                '    With DBManager
'                '        .Execute "Insert into tbFormList (FormName,FormDesc) Values ('frAlert','ระบบนัดหมาย/เตือน')"
'                '        .Execute "Insert into tbAccessRole (GroupID,FormName,CanAccess,CanAdd,CanUpdate,CanDelete,CanPrint) values ('" & AdminID & "','frAlert','Y','Y','Y','Y','Y')"
'                'End Sub
'                '---------------------------------------------------------
'        End If
'        Set tmpCon = Nothing
'        Set tmpRec = Nothing
'        Exit Function
'    End If
'    'tmpRec.Close
'
'    sqlStr = "Select * from tbTeacherData Where LoginName='" & userName & "'"
'    Set encObj = New RC4Crypt
'    If Trim(pwd) = "" Then
'        tmpPwd = ""
'    Else
'        tmpPwd = encObj.EncryptString(pwd, ENC_KEY)
'    End If
'    'sqlStr = sqlStr & " And LoginPwd='" & Trim(tmpPwd) & "'"
'    Set encObj = Nothing
'
'
'    tmpRec.Open sqlStr', tmpCon, adOpenStatic, adLockReadOnly
'    If tmpRec.EOF Then
'        Set tmpUser = Nothing
'    Else
'        If Trim(tmpRec("LoginPwd")) = Trim(tmpPwd) Then
'            Set tmpUser = New scmteacher
'            tmpUser.SetDataFromRecordset tmpRec
'        Else
'            Set tmpUser = Nothing
'        End If
'    End If
'    'tmpRec.Close
'    Set tmpRec = Nothing
'    Set ProcessLogin = tmpUser
'    Set tmpCon = Nothing
'End Function

'Public Function UserSearch(Optional ID As String = "", Optional strKey As String = "", _
'                                Optional inAddr As String = "", Optional strCond As String = "") As Collection
'Dim tmpRec As ADODB.Recordset
'Dim tmpCon As ADODB.Recordset
'Dim sqlStr As String
'Dim tmpUser As scmteacher
'Dim tmpCol As Collection
'Dim tmpStr As String
'    Set tmpCon = DBManager
'    sqlStr = "Select * From tbUser"
'    If Trim(ID) <> "" Then
'        sqlStr = sqlStr & " Where StaffID='" & ID & "'"
'    ElseIf Trim(strKey) <> "" Then
'        tmpStr = Replace(strKey, "*", "%")
'        If Trim(inAddr) = "" Then
'            sqlStr = sqlStr & " Where FName Like '%" & tmpStr & "%' Or LName Like '%" & tmpStr & "%'"
'        Else
'            sqlStr = sqlStr & " Where FName Like '%" & tmpStr & "%' Or LName Like '%" & tmpStr & "%'" & _
'                        " Or Addr1 Like '%" & tmpStr & "%' Or Addr2 Like '%" & tmpStr & "%' Or Addr3 Like '%" & _
'                        tmpStr & "%'"
'        End If
'    ElseIf Trim(strCond) <> "" Then
'        sqlStr = sqlStr & " Where " & strCond
'    End If
'    Set tmpRec = New ADODB.Recordset
'    tmpRec.Open sqlStr', tmpCon, adOpenStatic, adLockReadOnly
'    Set tmpCol = New Collection
'    Do Until tmpRec.EOF
'        Set tmpUser = New scmteacher
'        tmpUser.SetDataFromRecordset tmpRec
'        tmpCol.Add tmpUser, "'" & tmpUser.ID & "'"
'        Set tmpUser = Nothing
'        tmpRec.MoveNext
'    Loop
'    Set UserSearch = tmpCol
'    'tmpRec.Close
'    Set tmpRec = Nothing
'    Set tmpCol = Nothing
'    Set tmpCon = Nothing
'End Function

Public Sub PopulateFormList()
Dim tmpRec As ADODB.Recordset
Dim tmpCol As Collection
    If ColFormList Is Nothing Then
        Set tmpRec = New ADODB.Recordset
        tmpRec.Open "Select * From tbFormList Order By FormDesc" ', DBManager, adOpenDynamic, adLockOptimistic
        Set tmpCol = New Collection
        Do Until tmpRec.EOF
            tmpCol.Add Trim("" & tmpRec("FormName")) & "_" & Trim("" & tmpRec("FormDesc")), "'" & Trim("" & tmpRec("FormName")) & "'"
            tmpRec.MoveNext
        Loop
        'tmpRec.Close
        Set tmpRec = Nothing
        Set ColFormList = tmpCol
        Set tmpCol = Nothing
    End If
End Sub

Public Sub PopulateUserGroup()
Dim tmpRec As ADODB.Recordset
Dim tmpCol As Collection
Dim tmpGroup As AWSUserGroup
    If ColUserGroup Is Nothing Then
        Set tmpRec = New ADODB.Recordset
        tmpRec.Open "Select * From tbUserGroup Where IsDeleted='N' Order By GroupName" ', DBManager, adOpenStatic, adLockReadOnly
        Set tmpCol = New Collection
        Do Until tmpRec.EOF
            Set tmpGroup = New AWSUserGroup
            tmpGroup.SetDataFromRecordset tmpRec
            tmpCol.Add tmpGroup, "'" & tmpGroup.ID & "'"
            Set tmpGroup = Nothing
            tmpRec.MoveNext
        Loop
        'tmpRec.Close
        Set tmpRec = Nothing
        Set ColUserGroup = tmpCol
        Set tmpCol = Nothing
    End If
End Sub
