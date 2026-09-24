VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IMITumbol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim id_ As String
Dim tumbolName_ As String
Dim aumphurID_ As String
Dim aumphurName_ As String
Dim provinceID_ As String
Dim provinceName_ As String
Dim isChanged_ As Boolean
Dim Adesc_ As String
Public Property Get ID() As String
    ID = id_
End Property
Public Property Let ID(ByVal newVal As String)
    id_ = newVal
End Property

Public Property Get Name() As String
    Name = tumbolName_
End Property
Public Property Let Name(ByVal newVal As String)
    If Trim(tumbolName_) <> Trim(newVal) Then
        tumbolName_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get AumphurID() As String
    AumphurID = aumphurID_
End Property
Public Property Let AumphurID(ByVal newVal As String)
    If Trim(aumphurID_) <> Trim(newVal) Then
        aumphurID_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get AumphurName() As String
    AumphurName = aumphurName_
End Property
Public Property Let AumphurName(ByVal newVal As String)
    If Trim(aumphurName_) <> Trim(newVal) Then
        aumphurName_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get ProvinceID() As String
    ProvinceID = provinceID_
End Property
Public Property Let ProvinceID(ByVal newVal As String)
    If Trim(provinceID_) <> Trim(newVal) Then
        provinceID_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get ProvinceName() As String
    ProvinceName = provinceName_
End Property
Public Property Let ProvinceName(ByVal newVal As String)
    If Trim(provinceName_) <> Trim(newVal) Then
        provinceName_ = newVal
        isChanged_ = True
    End If
End Property
Public Property Get ADesc() As String
    ADesc = Adesc_
End Property
Public Property Let ADesc(ByVal newVal As String)
    If Trim(Adesc_) <> Trim(newVal) Then
        Adesc_ = newVal
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
Dim tmpRec As New ADODB.Recordset
Dim sqlStr As String
    sqlStr = "SELECT * FROM tbTumbol WHERE TumbolID='" & id_ & "'"
    GetDBConnection
    tmpRec.open sqlStr, DBConn, adOpenDynamic, adLockOptimistic
    If tmpRec.EOF Then
        tmpRec.AddNew
        tmpRec("TumbolID") = id_
    End If
    tmpRec("TumbolName") = Trim(tumbolName_)
    tmpRec("AumphurID") = Trim(aumphurID_)
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
    DBConnExc "Delete tbTumbol WHERE TumbolID='" & id_ & "'"
    Delete = ""
    Exit Function
delErr:
    Delete = err.Description
End Function

Public Sub SetDataFromRecordset(selRec As ADODB.Recordset)
    id_ = Trim("" & selRec("TumbolID"))
    tumbolName_ = Trim("" & selRec("TumbolName"))
    aumphurID_ = Trim("" & selRec("AumphurID"))
    aumphurName_ = Trim("" & selRec("AumphurName"))
    provinceID_ = Trim("" & selRec("ProvinceID"))
    provinceName_ = Trim("" & selRec("ProvinceName"))
    Adesc_ = Trim("" & selRec("ADesc"))
End Sub

