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
Dim ADesc_ As String
Public Property Get ID() As String
    ID = id_
End Property
Public Property Let ID(ByVal newval As String)
    id_ = newval
End Property

Public Property Get Name() As String
    Name = tumbolName_
End Property
Public Property Let Name(ByVal newval As String)
    If Trim(tumbolName_) <> Trim(newval) Then
        tumbolName_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get AumphurID() As String
    AumphurID = aumphurID_
End Property
Public Property Let AumphurID(ByVal newval As String)
    If Trim(aumphurID_) <> Trim(newval) Then
        aumphurID_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get AumphurName() As String
    AumphurName = aumphurName_
End Property
Public Property Let AumphurName(ByVal newval As String)
    If Trim(aumphurName_) <> Trim(newval) Then
        aumphurName_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get ProvinceID() As String
    ProvinceID = provinceID_
End Property
Public Property Let ProvinceID(ByVal newval As String)
    If Trim(provinceID_) <> Trim(newval) Then
        provinceID_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get ProvinceName() As String
    ProvinceName = provinceName_
End Property
Public Property Let ProvinceName(ByVal newval As String)
    If Trim(provinceName_) <> Trim(newval) Then
        provinceName_ = newval
        isChanged_ = True
    End If
End Property
Public Property Get ADesc() As String
    ADesc = ADesc_
End Property
Public Property Let ADesc(ByVal newval As String)
    If Trim(ADesc_) <> Trim(newval) Then
        ADesc_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get isChanged() As Boolean
    isChanged = isChanged_
End Property
Public Property Let isChanged(ByVal newval As Boolean)
    isChanged_ = newval
End Property

Public Sub Save()
On Error GoTo ErrD
Dim tmpRec As New ADODB.Recordset
    tmpRec.Open "Select * from tbTumbol Where TumbolID='" & id_ & "'", DBManager.GetDBConnection, adOpenDynamic, adLockOptimistic
    If tmpRec.EOF Then
        tmpRec.AddNew
        tmpRec("TumbolID") = id_
    End If
    tmpRec("TumbolName") = Trim(tumbolName_)
    tmpRec("AumphurID") = Trim(aumphurID_)
    tmpRec.Update
    tmpRec.Close
    Set tmpRec = Nothing
    isChanged_ = False
    Exit Sub
ErrD:
    MsgBox Err.Description & vbCrLf & "กรุณาลองใหม่หรือติดต่อผู้พัฒนาระบบ", vbCritical, "Error"
    Exit Sub
End Sub

Public Function Delete() As String
On Error GoTo delErr:
    DBManager.GetDBConnection.Execute "Delete tbTumbol Where TumbolID='" & id_ & "'"
    Delete = ""
    Exit Function
delErr:
    Delete = Err.Description
End Function

Public Sub SetDataFromRecordset(selRec As ADODB.Recordset)
    id_ = Trim("" & selRec("TumbolID"))
    tumbolName_ = Trim("" & selRec("TumbolName"))
    aumphurID_ = Trim("" & selRec("AumphurID"))
    aumphurName_ = Trim("" & selRec("AumphurName"))
    provinceID_ = Trim("" & selRec("ProvinceID"))
    provinceName_ = Trim("" & selRec("ProvinceName"))
    ADesc_ = Trim("" & selRec("Adesc"))
End Sub

