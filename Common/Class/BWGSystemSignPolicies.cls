VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BWGSystemSignPolicies"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim id_ As String
Dim companyID_ As String
Dim SignName1_ As String
Dim SignPosition1_ As String
Dim isActive_ As String
Dim isChanged_ As Boolean

'Security
Dim colSignPolicies As Collection

Public Property Get ID() As String
    ID = id_
End Property

Public Property Let ID(ByVal newVal As String)
    id_ = Trim(newVal)
End Property

Public Property Get CompanyID() As String
    CompanyID = companyID_
End Property

Public Property Let CompanyID(ByVal newVal As String)
    If Trim(companyID_) <> Trim(newVal) Then
        companyID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get SignName1() As String
    SignName1 = SignName1_
End Property
Public Property Let SignName1(ByVal newVal As String)
    If Trim(SignName1_) <> Trim(newVal) Then
        SignName1_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get SignPosition1() As String
    SignPosition1 = SignPosition1_
End Property
Public Property Let SignPosition1(ByVal newVal As String)
    If Trim(SignPosition1_) <> Trim(newVal) Then
        SignPosition1_ = Trim(newVal)
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

Public Sub Save()
On Error GoTo ErrD
Dim tmpRec As ADODB.Recordset
Dim tmpCon As ADODB.Connection
Dim sqlStr As String, x%

    GetDBConnection
    sqlStr = "SELECT * FROM tbSystemSignPolicies WHERE SignStaffID='" & id_ & "'"
    Set tmpRec = New ADODB.Recordset
'    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    tmpRec.open sqlStr, DBConn, adOpenDynamic, adLockOptimistic
    If tmpRec.EOF Then
        tmpRec.AddNew
        tmpRec("SignStaffID") = id_
    End If
    tmpRec("CompanyID") = Trim(companyID_)
    tmpRec("SignName1") = Trim(SignName1_)
    tmpRec("SignPosition1") = Trim(SignPosition1_)
    If Trim(isActive_) = "" Then
        isActive_ = "Y"
    End If
    tmpRec("isActive") = Trim(isActive_)

    tmpRec.Update
    tmpRec.Close
    GetDBConnection
    
    Set tmpRec = Nothing
    isChanged_ = False
    Exit Sub
ErrD:
    GetMsg "", vbCritical
    Exit Sub
End Sub

Public Function DeActivate() As String
On Error GoTo delErr:
    DBConnExc "UPDATE tbSystemSignPolicies SET isActive = 'Y' WHERE SignStaffID='" & id_ & "'"
    DeActivate = ""
    Exit Function
delErr:
    DeActivate = err.Description
End Function

Public Sub SetDataFromRecordset(selRec As ADODB.Recordset)
    id_ = selRec("SignStaffID")
    companyID_ = Trim("" & selRec!CompanyID)
    SignName1_ = Trim("" & selRec("SignName1"))
    SignPosition1_ = Trim("" & selRec("SignPosition1"))
    isActive_ = Trim("" & selRec("isActive"))
End Sub

Public Property Get SystemSignPolicies() As Collection
Dim tmpS As New BWGSearchManager
    If colSignPolicies Is Nothing Then
        Set colSignPolicies = tmpS.UserSearch("StaffID IN (SELECT SignStaffID FROM tbSystemSignPolicies WHERE SignStaffID = '" & id_ & "')")
    End If
    Set tmpS = Nothing
    Set SystemSignPolicies = colSignPolicies
    Set colSignPolicies = Nothing
End Property

