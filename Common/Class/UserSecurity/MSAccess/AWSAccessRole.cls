VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AWSAccessRole"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim id_ As String
Dim groupID_ As String
Dim formName_ As String
Dim canAccess_ As String
Dim canAdd_ As String
Dim canUpdate_ As String
Dim canDelete_ As String
Dim canPrint_ As String
Dim isChanged_ As Boolean

Public Property Get ID() As String
    ID = Trim(id_)
End Property
Public Property Let ID(ByVal newVal As String)
    id_ = Trim(newVal)
End Property

Public Property Get GroupID() As String
    GroupID = Trim(groupID_)
End Property
Public Property Let GroupID(ByVal newVal As String)
    If Trim(groupID_) <> Trim(newVal) Then
        groupID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get FormName() As String
    FormName = Trim(formName_)
End Property
Public Property Let FormName(ByVal newVal As String)
    If Trim(formName_) <> Trim(newVal) Then
        formName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get CanAccess() As Boolean
    If canAccess_ = "Y" Then
        CanAccess = True
    Else
        CanAccess = False
    End If
End Property
Public Property Let CanAccess(ByVal newVal As Boolean)
Dim tmpStr As String
    If newVal Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(canAccess_) <> Trim(tmpStr) Then
        canAccess_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property

Public Property Get CanAddNew() As Boolean
    If canAdd_ = "Y" Then
        CanAddNew = True
    Else
        CanAddNew = False
    End If
End Property
Public Property Let CanAddNew(ByVal newVal As Boolean)
Dim tmpStr As String
    If newVal Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(canAdd_) <> Trim(tmpStr) Then
        canAdd_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property

Public Property Get CanUpdate() As Boolean
    If canUpdate_ = "Y" Then
        CanUpdate = True
    Else
        CanUpdate = False
    End If
End Property
Public Property Let CanUpdate(ByVal newVal As Boolean)
Dim tmpStr As String
    If newVal Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(canUpdate_) <> Trim(tmpStr) Then
        canUpdate_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property

Public Property Get CanDelete() As Boolean
    If canDelete_ = "Y" Then
        CanDelete = True
    Else
        CanDelete = False
    End If
End Property
Public Property Let CanDelete(ByVal newVal As Boolean)
Dim tmpStr As String
    If newVal Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(canDelete_) <> Trim(tmpStr) Then
        canDelete_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property

Public Property Get CanPrint() As Boolean
    If canPrint_ = "Y" Then
        CanPrint = True
    Else
        CanPrint = False
    End If
End Property
Public Property Let CanPrint(ByVal newVal As Boolean)
Dim tmpStr As String
    If newVal Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(canPrint_) <> Trim(tmpStr) Then
        canPrint_ = Trim(tmpStr)
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
    Set tmpRec = New ADODB.Recordset
    tmpRec.Open "Select * From tbAccessRole Where RoleAccessID='" & id_ & "'", DBManager.GetDBConnection, adOpenDynamic, adLockOptimistic
    If tmpRec.EOF Then
        tmpRec.AddNew
        tmpRec("RoleAccessID") = Trim(id_)
    End If
    tmpRec("GroupID") = Trim(groupID_)
    tmpRec("FormName") = Trim(formName_)
    tmpRec("CanAccess") = Trim(canAccess_)
    tmpRec("CanAdd") = Trim(canAdd_)
    tmpRec("CanUpdate") = Trim(canUpdate_)
    tmpRec("CanDelete") = Trim(canDelete_)
    tmpRec("CanPrint") = Trim(canPrint_)
    tmpRec.Update
    tmpRec.Close
    Set tmpRec = Nothing
    isChanged_ = False
End Sub

Public Sub SetDataFromRecordset(selRec As ADODB.Recordset)
    id_ = Trim("" & selRec("RoleAccessID"))
    groupID_ = Trim("" & selRec("GroupID"))
    formName_ = Trim("" & selRec("FormName"))
    canAccess_ = Trim("" & selRec("CanAccess"))
    canAdd_ = Trim("" & selRec("CanAdd"))
    canUpdate_ = Trim("" & selRec("CanUpdate"))
    canDelete_ = Trim("" & selRec("CanDelete"))
    canPrint_ = Trim("" & selRec("CanPrint"))
End Sub
