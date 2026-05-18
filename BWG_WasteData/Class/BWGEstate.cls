VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BWGEstate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim id_ As String
Dim EstateName_ As String
Dim isChanged_ As Boolean
Dim ProvID_ As String
Public Property Get ID() As String
    ID = id_
End Property
Public Property Let ID(ByVal newval As String)
    id_ = Trim(newval)
End Property

Public Property Get EstateName() As String
    EstateName = EstateName_
End Property
Public Property Let EstateName(ByVal newval As String)
    If Trim(EstateName_) <> Trim(newval) Then
        EstateName_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get ProvID_() As String
    ProvID = ProvID_
End Property
Public Property Let ProvID(ByVal newval As String)
    If Trim(ProvID_) <> Trim(newval) Then
        ProvID_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get isChanged() As Boolean
    isChanged = isChanged_
End Property
Public Property Let isChanged(ByVal newval As Boolean)
    isChanged_ = newval
End Property

Public Sub SetDataFromRecordset(selRec As ADODB.Recordset)
    id_ = selRec("EstateID")
    EstateName_ = Trim("" & selRec("EstateName"))
    ProvID_ = selRec("ProvinceID")
End Sub


