Attribute VB_Name = "modUserSecurity"
Option Explicit
'Public CurrentUser As scmteacher
Public CanAccess As Boolean, CanUpdate As Boolean
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
'    tmpRec.Open "SELECT * FROM tbUser"', tmpCon, adOpenKeyset, adLockOptimistic
'    If tmpRec.EOF Then
'        Set ProcessLogin = New scmteacher
'        'tmpRec.Close
'        tmpRec.Open "Select count(*) as RecCount from tbUserGroup"', tmpCon, adOpenKeyset, adLockOptimistic
'        If tmpRec("RecCount") = 0 Then
'            admID = getNewProductID(modPrefix.UserGroupPrefix)
'            DBConnExc "Insert into tbUserGroup (GroupID,GroupName,GroupDesc,isActive,isSystem) Values ('" & _
'                                admID & "','Administrator','Administrator','Y','Y')"
'            PopulateModuleList admID 'This function will depend on each project, written in modMain
'            'Code will be like this--------------------------------------
'                'Public Sub PopulateModuleList(AdminID As String)
'                '    With DBManager
'                '        .Execute "Insert into tbFormList (FormName,FormDesc) Values ('frmAlert','ระบบนัดหมาย/เตือน')"
'                '        .Execute "Insert into tbAccessRole (GroupID,FormName,CanAccess,CanAdd,CanUpdate,CanDelete,CanPrint) values ('" & AdminID & "','frmAlert','Y','Y','Y','Y','Y')"
'                'End Sub
'                '---------------------------------------------------------
'        End If
'        Set tmpCon = Nothing
'        Set tmpRec = Nothing
'        Exit Function
'    End If
'    'tmpRec.Close
'
'    sqlStr = "SELECT * FROM tbTeacherData WHERE LoginName='" & userName & "'"
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
'    tmpRec.Open sqlStr', tmpCon, adOpenKeyset, adLockOptimistic
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
'    sqlStr = "SELECT * FROM tbUser"
'    If Trim(ID) <> "" Then
'        sqlStr = sqlStr & " WHERE StaffID='" & ID & "'"
'    ElseIf Trim(strKey) <> "" Then
'        tmpStr = Replace(strKey, "*", "%")
'        If Trim(inAddr) = "" Then
'            sqlStr = sqlStr & " WHERE FName Like '%" & tmpStr & "%' Or LName Like '%" & tmpStr & "%'"
'        Else
'            sqlStr = sqlStr & " WHERE FName Like '%" & tmpStr & "%' Or LName Like '%" & tmpStr & "%'" & _
'                        " Or Addr1 Like '%" & tmpStr & "%' Or Addr2 Like '%" & tmpStr & "%' Or Addr3 Like '%" & _
'                        tmpStr & "%'"
'        End If
'    ElseIf Trim(strCond) <> "" Then
'        sqlStr = sqlStr & " WHERE " & strCond
'    End If
'    Set tmpRec = New ADODB.Recordset
'    tmpRec.Open sqlStr', tmpCon, adOpenKeyset, adLockOptimistic
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

Public Sub SetPermission(ByVal frm As Form, ByRef CanAccess As Boolean, ByRef CanUpdate As Boolean)
On Error Resume Next
Dim UserGrp As Collection
Dim Group As New AWSUserGroup
Dim i As Integer
Dim tSql As String
Dim RS As ADODB.Recordset
Dim CanAdd As Boolean, CanDelete As Boolean, CanPrint As Boolean, CanApprove As Boolean
Dim ctl As Object

    Set UserGrp = CurrentUser.getUserGroups
    For i = 1 To UserGrp.Count
        Set Group = UserGrp(i)
        tSql = "SELECT * FROM tbAccessRole WHERE GroupID = '" & Group.ID & "' AND FormName = '" & frm.Name & "'"
        Set RS = New ADODB.Recordset
        
        Set RS = GetRS(tSql, adOpenForwardOnly, adLockReadOnly)
        If Not RS.EOF Then
            If CanAccess = False Then CanAccess = IIf(RS!CanAccess = "Y", True, False)
            If CanAdd = False Then CanAdd = IIf(RS!CanAdd = "Y", True, False)
            If CanUpdate = False Then CanUpdate = IIf(RS!CanUpdate = "Y", True, False)
            If CanDelete = False Then CanDelete = IIf(RS!CanDelete = "Y", True, False)
            If CanPrint = False Then CanPrint = IIf(RS!CanPrint = "Y", True, False)
            If CanApprove = False Then CanApprove = IIf(RS!CanApprove = "Y", True, False)
        End If
        Set RS = Nothing
    Next
    Set RS = Nothing
    For Each ctl In frm
'        Debug.Print Ctl.Name
        If LCase(ctl.Name) = "cmdsave" Then ctl.Enabled = CanAdd
        If LCase(ctl.Name) = "cmddel" Then ctl.Enabled = CanDelete
        If LCase(ctl.Name) = "cmddelete" Then ctl.Enabled = CanDelete
        If LCase(ctl.Name) = "cmdcancel" Then ctl.Enabled = CanDelete
        If LCase(ctl.Name) = "cmdapprove" Then ctl.Enabled = CanApprove
        If LCase(ctl.Name) = "cmdprint" Then ctl.Enabled = CanPrint
        If LCase(ctl.Name) = "cmdpreview" Then ctl.Enabled = CanAccess
    Next
End Sub


