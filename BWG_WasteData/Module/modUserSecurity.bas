Attribute VB_Name = "modUserSecurity"
Option Explicit

Public Sub SetPermission(ByVal frm As Form, ByRef CanAccess As Boolean, ByRef CanUpdate As Boolean)
On Error Resume Next
Dim UserGrp As Collection
Dim Group As New AWSUserGroup
Dim i As Integer
Dim tSql As String
Dim RS As ADODB.Recordset
Dim CanAdd As Boolean, CanDelete As Boolean, CanPrint As Boolean, CanApprove As Boolean
Dim Ctl As Object

    Set UserGrp = CurrentUser.getUserGroups
    For i = 1 To UserGrp.Count
        Set Group = UserGrp(i)
        tSql = "SELECT * FROM tbAccessRole WHERE GroupID = '" & Group.ID & "' AND FormName = '" & frm.Name & "'"
        Set RS = New ADODB.Recordset
        Set RS = GetRS(tSql, adOpenForwardOnly, adLockReadOnly)
'        Set RS = GetRS(tSql, adOpenForwardOnly, adLockReadOnly)
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
    For Each Ctl In frm
'        Debug.Print Ctl.Name
        If LCase(Ctl.Name) = "cmdsave" Then Ctl.Enabled = CanAdd
        If LCase(Ctl.Name) = "cmddel" Then Ctl.Enabled = CanDelete
        If LCase(Ctl.Name) = "cmddelete" Then Ctl.Enabled = CanDelete
        If LCase(Ctl.Name) = "cmdcancel" Then Ctl.Enabled = CanDelete
        If LCase(Ctl.Name) = "cmdapprove" Then Ctl.Enabled = CanApprove
        If LCase(Ctl.Name) = "cmdprint" Then Ctl.Enabled = CanPrint
        If LCase(Ctl.Name) = "cmdpreview" Then Ctl.Enabled = CanPrint
    Next
End Sub
