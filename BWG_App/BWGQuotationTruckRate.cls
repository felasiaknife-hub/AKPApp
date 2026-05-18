VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BWGQuotationTruckRate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim id_ As String
Dim TruckRateID_ As String
Dim TruckTypeID_ As String

Dim QuotationID_ As String
Dim TypeDesc_ As String
Dim SubTypeID_ As String
Dim SubTypeDesc_ As String
Dim TumbolName_  As String
Dim AumphurName_  As String
Dim ProvinceName_  As String

Dim TumbolID_  As String
Dim AumphurID_  As String
Dim ProvinceID_  As String
Dim isChanged_ As Boolean
Dim Rate_ As Double

Public Property Get ProvinceID() As String
    ProvinceID = ProvinceID_
End Property
Public Property Let ProvinceID(ByVal newval As String)
    If Trim(ProvinceID_) <> Trim(newval) Then
        ProvinceID_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get AumphurID() As String
    AumphurID = AumphurID_
End Property
Public Property Let AumphurID(ByVal newval As String)
    If Trim(AumphurID_) <> Trim(newval) Then
        AumphurID_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get TumbolID() As String
    TumbolID = TumbolID_
End Property
Public Property Let TumbolID(ByVal newval As String)
    If Trim(TumbolID_) <> Trim(newval) Then
        TumbolID_ = Trim(newval)
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
Public Property Get AumphurName() As String
    AumphurName = AumphurName_
End Property
Public Property Let AumphurName(ByVal newval As String)
    If Trim(AumphurName_) <> Trim(newval) Then
        AumphurName_ = Trim(newval)
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

Public Property Get id() As String
    id = id_
End Property
Public Property Let id(ByVal newval As String)
    id_ = Trim(newval)
End Property

Public Property Get TruckTypeID() As String
    TruckTypeID = TruckTypeID_
End Property
Public Property Let TruckTypeID(ByVal newval As String)
    If Trim(TypeDesc_) <> Trim(newval) Then
        TruckTypeID_ = Trim(newval)
        isChanged_ = True
    End If
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
Public Property Get SubTypeDesc() As String
    SubTypeDesc = SubTypeDesc_
End Property
Public Property Let SubTypeDesc(ByVal newval As String)
    If Trim(SubTypeDesc_) <> Trim(newval) Then
        SubTypeDesc_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get SubTypeID() As String
    SubTypeID = SubTypeID_
End Property
Public Property Let SubTypeID(ByVal newval As String)
    If Trim(SubTypeID_) <> Trim(newval) Then
        SubTypeID_ = Trim(newval)
        isChanged_ = True
    End If
End Property


Public Property Get TruckRateID() As String
    TruckRateID = TruckRateID_
End Property
Public Property Let TruckRateID(ByVal newval As String)
    If Trim(TruckRateID_) <> Trim(newval) Then
        TruckRateID_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get QuotationID() As String
    QuotationID = QuotationID_
End Property
Public Property Let QuotationID(ByVal newval As String)
    If Trim(QuotationID_) <> Trim(newval) Then
        QuotationID_ = Trim(newval)
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

Public Sub Save()
On Error GoTo ErrD
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
    sqlStr = "SELECT *, CONVERT(DATETIME, GETDATE(), 103) AS GETDATE FROM tbQuotationTruckRate WHERE QuotationTruckRateID='" & id_ & "'"
    Set tmpRec = New ADODB.Recordset
    GetDBConnection
    tmpRec.open sqlStr, DBConn, adOpenDynamic, adLockOptimistic
    If tmpRec.EOF Then
        tmpRec.AddNew
        tmpRec("QuotationTruckRateID") = id_
    End If
    tmpRec("QuotationID") = Trim(QuotationID_)
    tmpRec("TruckRateID") = Trim(TruckRateID_)
    tmpRec("Rate") = Trim(Rate_)
    tmpRec("isDelete") = ""
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
    DBConnExc "Delete tbQuotationTruckRate WHERE QuotationTruckRateID='" & id_ & "'"
    Delete = ""
    Exit Function
delErr:
    Delete = err.Description
End Function
Public Sub SetDataFromRecordset(selRec As ADODB.Recordset)
    id_ = selRec("QuotationTruckRateID")
    TypeDesc_ = Trim("" & selRec("TruckTypeDesc"))
    TruckRateID = Trim("" & selRec("TruckRateID"))
    SubTypeID_ = Trim("" & selRec("SubTypeID"))
    SubTypeDesc_ = Trim("" & selRec("SubTypeDesc"))
    QuotationID_ = Trim("" & selRec("QuotationID"))
    Rate_ = IIf(IsNull(selRec!Rate), 0, selRec!Rate)
    TumbolName_ = Trim("" & selRec("TumbolName"))
    AumphurName_ = Trim("" & selRec("AumphurName"))
    ProvinceName_ = Trim("" & selRec("ProvinceName"))
    TumbolID_ = Trim("" & selRec("TumbolID"))
    AumphurID_ = Trim("" & selRec("AumphurID"))
    ProvinceID_ = Trim("" & selRec("ProvinceID"))
End Sub
Public Sub SetDataFromRecordset2(selRec As ADODB.Recordset)
    id_ = selRec("SubTypeID")
    TypeDesc_ = Trim("" & selRec("TruckTypeDesc"))
    TruckRateID = Trim("" & selRec("TruckRateID"))
    SubTypeID_ = Trim("" & selRec("SubTypeID"))
    SubTypeDesc_ = Trim("" & selRec("SubTypeDesc"))
    QuotationID_ = Trim("" & selRec("QuotationID"))
    Rate_ = IIf(IsNull(selRec!Rate), 0, selRec!Rate)
    TruckTypeID_ = IIf(IsNull(selRec!TruckTypeID), 0, selRec!TruckTypeID)
End Sub




