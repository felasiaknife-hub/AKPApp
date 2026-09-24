VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AWSUserGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim id_ As String
Dim name_ As String
Dim desc_ As String
Dim isActive_ As String
Dim colAccessRole_ As Collection
Dim isSystem_ As String
Dim isChanged_ As Boolean

Public Property Get ID() As String
    ID = Trim(id_)
End Property
Public Property Let ID(ByVal newVal As String)
    id_ = Trim(newVal)
End Property

Public Property Get Name() As String
    Name = Trim(name_)
End Property
Public Property Let Name(ByVal newVal As String)
    If Trim(name_) <> Trim(newVal) Then
        name_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get Description() As String
    Description = Trim(desc_)
End Property
Public Property Let Description(ByVal newVal As String)
    If Trim(desc_) <> Trim(newVal) Then
        desc_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get isActive() As Boolean
    If isActive_ = "Y" Then
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

Public Property Get isSystem() As Boolean
    If isSystem_ = "Y" Then
        isSystem = True
    Else
        isSystem = False
    End If
End Property
Public Property Let isSystem(ByVal newVal As Boolean)
Dim tmpStr As String
    If newVal Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(isSystem_) <> Trim(tmpStr) Then
        isSystem_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property

Public Property Get isChanged() As Boolean
    isChanged = isChanged_
End Property
Public Property Let isChanged(ByVal newVal As Boolean)
    isChanged_ = newVal
End Property

Public Sub Save()
Dim tmpRec As ADODB.Recordset
Dim x%
Dim tmpRole As AWSAccessRole
    Set tmpRec = New ADODB.Recordset
    tmpRec.Open "Select * From tbUserGroup Where GroupID='" & id_ & "'", DBManager.GetDBConnection, adOpenDynamic, adLockOptimistic
    If tmpRec.EOF Then
        tmpRec.AddNew
        tmpRec("GroupID") = Trim(id_)
    End If
    tmpRec("GroupName") = Trim(name_)
    tmpRec("GroupDesc") = Trim(desc_)
    tmpRec("isActive") = Trim(isActive_)
    tmpRec.Update
    tmpRec.Close
    Set tmpRec = Nothing
    For x = 1 To colAccessRole_.Count
        Set tmpRole = colAccessRole_(x)
        tmpRole.Save
        Set tmpRole = Nothing
    Next
    isChanged_ = False
End Sub

Public Sub Delete()
    DBManager.GetDBConnection.Execute "Update tbUserGroup Set isDeleted='Y',DeletedDate='" & TodayDate & "' Where GroupID='" & id_ & "'"
End Sub

Public Sub SetDataFromRecordset(selRec As ADODB.Recordset)
    id_ = Trim("" & selRec("GroupID"))
    name_ = Trim("" & selRec("GroupName"))
    desc_ = Trim("" & selRec("GroupDesc"))
    isActive_ = Trim("" & selRec("isActive"))
End Sub

Public Function getAccessRoles() As Collection
Dim tmpRec As ADODB.Recordset
Dim tmpRole As AWSAccessRole
    If colAccessRole_ Is Nothing Then
        Set colAccessRole_ = New Collection
        Set tmpRec = New ADODB.Recordset
        tmpRec.Open "Select * From tbAccessRole Where GroupID='" & id_ & "' Order By FormName", DBManager.GetDBConnection, adOpenStatic, adLockReadOnly
        Do Until tmpRec.EOF
            Set tmpRole = New AWSAccessRole
            tmpRole.SetDataFromRecordset tmpRec
            colAccessRole_.Add tmpRole, "'" & tmpRole.FormName & "'"
            Set tmpRole = Nothing
            tmpRec.MoveNext
        Loop
    End If
    Set getAccessRoles = colAccessRole_
End Function
