VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BWGTruckSubType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim id_ As String
Dim truckTypeID_ As String
Dim subTypeDesc_ As String
Dim typeDesc_ As String
Dim isChargeByWeight_ As String
Dim isChanged_ As Boolean

Public Property Get ID() As String
    ID = id_
End Property
Public Property Let ID(ByVal newVal As String)
    id_ = Trim(newVal)
End Property

Public Property Get TruckTypeID() As String
    TruckTypeID = truckTypeID_
End Property
Public Property Let TruckTypeID(ByVal newVal As String)
    If Trim(truckTypeID_) <> Trim(newVal) Then
        truckTypeID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get SubTypeDesc() As String
    SubTypeDesc = subTypeDesc_
End Property
Public Property Let SubTypeDesc(ByVal newVal As String)
    If Trim(subTypeDesc_) <> Trim(newVal) Then
        subTypeDesc_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get TypeDesc() As String
    TypeDesc = typeDesc_
End Property
Public Property Let TypeDesc(ByVal newVal As String)
    If Trim(typeDesc_) <> Trim(newVal) Then
        typeDesc_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get isChargeByWeight() As Boolean
    If Trim(isChargeByWeight_) = "Y" Then
        isChargeByWeight = True
    Else
        isChargeByWeight = False
    End If
End Property
Public Property Let isChargeByWeight(ByVal newVal As Boolean)
Dim tmpStr As String
    If newVal Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(isChargeByWeight_) <> Trim(tmpStr) Then
        isChargeByWeight_ = Trim(tmpStr)
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
Dim sqlStr As String
    sqlStr = "SELECT *, CONVERT(DATETIME, GETDATE(), 103) AS GETDATE FROM tbTruckSubType WHERE SubTypeID='" & id_ & "'"
    Set tmpRec = New ADODB.Recordset
    GetDBConnection
    tmpRec.open sqlStr, DBConn, adOpenDynamic, adLockOptimistic
    If tmpRec.EOF Then
        tmpRec.AddNew
        tmpRec("SubTypeID") = id_
    End If
    tmpRec("TruckTypeID") = Trim(truckTypeID_)
    tmpRec("SubTypeDesc") = Trim(subTypeDesc_)
    tmpRec("isChargeByWeight") = Trim(isChargeByWeight_)
    
    tmpRec!LastUser = Trim(Trim(CurrentUser.StaffFName) & " " & Trim(CurrentUser.StaffLName))
    tmpRec!LastUpdate = GetSystemDateTime.SystemDateTime
    
    tmpRec.Update
'    tmpRec.Close
    Set tmpRec = Nothing
    isChanged_ = False
    CloseDBConnection
    Exit Sub
ErrD:
    GetMsg "", vbCritical
    Exit Sub
End Sub

Public Function Delete() As String
On Error GoTo delErr:
    DBConnExc "Delete tbTruckSubType WHERE SubTypeID='" & id_ & "'"
    Delete = ""
    Exit Function
delErr:
    Delete = err.Description
End Function

Public Sub SetDataFromRecordset(selRec As ADODB.Recordset)
    id_ = selRec("SubTypeID")
    truckTypeID_ = Trim("" & selRec("TruckTypeID"))
    subTypeDesc_ = Trim("" & selRec("SubTypeDesc"))
    typeDesc_ = Trim("" & selRec("TruckTypeDesc"))
    isChargeByWeight_ = Trim("" & selRec("isChargeByWeight"))
End Sub
