VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BWGZoneArea"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim id_ As String
Dim zoneName_ As String
Dim zoneDesc_ As String
Dim isChanged_ As Boolean

Public Property Get ID() As String
    ID = id_
End Property
Public Property Let ID(ByVal newVal As String)
    id_ = Trim(newVal)
End Property

Public Property Get ZoneName() As String
    ZoneName = zoneName_
End Property
Public Property Let ZoneName(ByVal newVal As String)
    If Trim(zoneName_) <> Trim(newVal) Then
        zoneName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get Description() As String
    Description = zoneDesc_
End Property
Public Property Let Description(ByVal newVal As String)
    If Trim(zoneDesc_) <> Trim(newVal) Then
        zoneDesc_ = Trim(newVal)
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
    Set tmpRec = New ADODB.Recordset
    sqlStr = "SELECT * FROM tbZoneArea WHERE ZoneID='" & id_ & "'"
    GetDBConnection
    tmpRec.open sqlStr, DBConn, adOpenDynamic, adLockOptimistic
    If tmpRec.EOF Then
        tmpRec.AddNew
        tmpRec("ZoneID") = id_
    End If
    tmpRec("ZoneName") = Trim(zoneName_)
    tmpRec("ZoneDesc") = Trim(zoneDesc_)
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
    DBConnExc "Delete tbZoneArea WHERE ZoneID='" & id_ & "'"
    Delete = ""
    Exit Function
delErr:
    Delete = err.Description
End Function

Public Sub SetDataFromRecordset(selRec As ADODB.Recordset)
    id_ = selRec("ZoneID")
    zoneName_ = Trim("" & selRec("ZoneName"))
    zoneDesc_ = Trim("" & selRec("ZoneDesc"))
End Sub

