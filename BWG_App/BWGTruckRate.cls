VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BWGTruckRate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim id_ As String
Dim TypeDesc_ As String
Dim canBeTrailer_ As String
Dim isFactoryCar_ As String
Dim isChanged_ As Boolean
Dim Rate_ As Double
Dim tumbolID_ As String
Dim truckTypeID_ As String
Dim TumbolName_ As String
Dim aumphurID_ As String
Dim AumphurName_ As String
Dim provinceID_ As String
Dim ProvinceName_ As String

Public Property Get id() As String
    id = id_
End Property
Public Property Let id(ByVal newval As String)
    id_ = Trim(newval)
End Property

Public Property Get TypeDesc() As String
    TypeDesc = TypeDesc_
End Property
Public Property Let TypeDesc(ByVal newval As String)
    If Trim(TypeDesc_) <> Trim(newval) Then
        TypeDesc_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get TruckTypeID() As String
    TruckTypeID = truckTypeID_
End Property
Public Property Let TruckTypeID(ByVal newval As String)
    If Trim(truckTypeID_) <> Trim(newval) Then
        TruckTypeID = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get CanBeTrailer() As Boolean
    If canBeTrailer_ = "Y" Then
        CanBeTrailer = True
    Else
        CanBeTrailer = False
    End If
End Property
Public Property Let CanBeTrailer(ByVal newval As Boolean)
Dim tmpStr As String
    If newval Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(canBeTrailer_) <> Trim(tmpStr) Then
        canBeTrailer_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property

Public Property Get isFactoryCar() As Boolean
    If isFactoryCar_ = "Y" Then
        isFactoryCar = True
    Else
        isFactoryCar = False
    End If
End Property
Public Property Let isFactoryCar(ByVal newval As Boolean)
Dim tmpStr As String
    If newval Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(isFactoryCar_) <> Trim(tmpStr) Then
        isFactoryCar_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property

Public Property Get isChanged() As Boolean
    isChanged = isChanged_
End Property
Public Property Let isChanged(ByVal newval As Boolean)
    isChanged_ = newval
End Property

Public Property Get Rate() As Double
    Rate = Rate_
End Property
Public Property Let Rate(ByVal newval As Double)
     If Trim(Rate_) <> Trim(newval) Then
        Rate_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get TumbolID() As String
    TumbolID = tumbolID_
End Property
Public Property Let TumbolID(ByVal newval As String)
     If Trim(tumbolID_) <> Trim(newval) Then
        tumbolID_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get TumbolName() As String
    TumbolName = TumbolName_
End Property
Public Property Let TumbolName(ByVal newval As String)
     If Trim(TumbolName_) <> Trim(newval) Then
        TumbolName_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get AumphurID() As String
    AumphurID = aumphurID_
End Property
Public Property Let AumphurID(ByVal newval As String)
     If Trim(aumphurID_) <> Trim(newval) Then
        aumphurID_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get AumphurName() As String
    AumphurName = AumphurName_
End Property
Public Property Let AumphurName(ByVal newval As String)
     If Trim(AumphurName_) <> Trim(newval) Then
        AumphurName_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get ProvinceID() As String
    ProvinceID = provinceID_
End Property
Public Property Let ProvinceID(ByVal newval As String)
     If Trim(provinceID_) <> Trim(newval) Then
        provinceID_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get ProvinceName() As String
    ProvinceName = ProvinceName_
End Property
Public Property Let ProvinceName(ByVal newval As String)
     If Trim(ProvinceName_) <> Trim(newval) Then
        ProvinceName_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Sub Save()
On Error GoTo ErrD
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
    sqlStr = "SELECT *, CONVERT(DATETIME, GETDATE(), 103) AS GETDATE FROM tbTruckRate WHERE TruckRateID='" & id_ & "'"
    Set tmpRec = New ADODB.Recordset
    GetDBConnection
    tmpRec.open sqlStr, DBConn, adOpenDynamic, adLockOptimistic
    If tmpRec.EOF Then
        tmpRec.AddNew
        tmpRec("TruckRateID") = id_
    End If
    tmpRec("TruckTypeID") = Trim(truckTypeID_)
    tmpRec("TumbolID") = Trim(tumbolID_)
    tmpRec("Rate") = Trim(Rate_)
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
    DBConnExc "Delete tbTruckRate WHERE TruckRateID='" & id_ & "'"
    Delete = ""
    Exit Function
delErr:
    Delete = err.Description
End Function
Public Sub SetDataFromRecordset(selRec As ADODB.Recordset)
    id_ = selRec("TruckRateID")
    TypeDesc_ = Trim("" & selRec("TruckTypeDesc"))
    canBeTrailer_ = Trim("" & selRec("CanBeTrailer"))
    isFactoryCar_ = Trim("" & selRec("isFactoryCar"))
    Rate_ = IIf(IsNull(selRec!Rate), 0, selRec!Rate)
    tumbolID_ = Trim("" & selRec("TumbolID"))
    TumbolName_ = Trim("" & selRec("TumbolName"))
    aumphurID_ = Trim("" & selRec("AumphurID"))
    AumphurName = Trim("" & selRec("AumphurName"))
    provinceID_ = Trim("" & selRec("ProvinceID"))
    ProvinceName = Trim("" & selRec("ProvinceName"))
End Sub


