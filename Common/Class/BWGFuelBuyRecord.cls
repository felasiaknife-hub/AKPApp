VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BWGFuelBuyRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim id_ As String
Dim vendorCarID_ As String
Dim vendorName_ As String
Dim CarRegisID_ As String
Dim carTypeDesc_ As String
Dim buyDate_ As String
Dim DriverName_ As String
Dim fuelQty_ As Double
Dim unitPrice_ As Double
Dim isPaid_ As String
Dim remDesc_ As String
Dim isChanged_ As Boolean

Public Property Get ID() As String
    ID = id_
End Property
Public Property Let ID(ByVal newVal As String)
    id_ = Trim(newVal)
End Property

Public Property Get VendorCarID() As String
    VendorCarID = vendorCarID_
End Property
Public Property Let VendorCarID(ByVal newVal As String)
    If Trim(vendorCarID_) <> Trim(newVal) Then
        vendorCarID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get VendorName() As String
    VendorName = vendorName_
End Property
Public Property Let VendorName(ByVal newVal As String)
    If Trim(vendorName_) <> Trim(newVal) Then
        vendorName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get CarRegisID() As String
    CarRegisID = CarRegisID_
End Property
Public Property Let CarRegisID(ByVal newVal As String)
    If Trim(CarRegisID_) <> Trim(newVal) Then
        CarRegisID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get CarTypeDesc() As String
    CarTypeDesc = carTypeDesc_
End Property
Public Property Let CarTypeDesc(ByVal newVal As String)
    If Trim(carTypeDesc_) <> Trim(newVal) Then
        carTypeDesc_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get BuyDate() As String
    BuyDate = buyDate_
End Property
Public Property Let BuyDate(ByVal newVal As String)
    If Trim(buyDate_) <> Trim(newVal) Then
        buyDate_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get DriverName() As String
    DriverName = DriverName_
End Property
Public Property Let DriverName(ByVal newVal As String)
    If Trim(DriverName_) <> Trim(newVal) Then
        DriverName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get FuelQty() As Double
    FuelQty = fuelQty_
End Property
Public Property Let FuelQty(ByVal newVal As Double)
    If fuelQty_ <> newVal Then
        fuelQty_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get UnitPrice() As Double
    UnitPrice = unitPrice_
End Property
Public Property Let UnitPrice(ByVal newVal As Double)
    If unitPrice_ <> newVal Then
        unitPrice_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get isPaid() As Boolean
    If Trim(isPaid_) = "Y" Then
        isPaid = True
    Else
        isPaid = False
    End If
End Property
Public Property Let isPaid(ByVal newVal As Boolean)
Dim tmpStr As String
    If newVal Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If isPaid_ <> Trim(tmpStr) Then
        isPaid_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property

Public Property Get RemDesc() As String
    RemDesc = remDesc_
End Property
Public Property Let RemDesc(ByVal newVal As String)
    If Trim(remDesc_) <> Trim(newVal) Then
        remDesc_ = Trim(newVal)
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
    sqlStr = "SELECT * FROM tbFuelBuyRecord WHERE BuyRecordID='" & id_ & "'"
    Set tmpRec = New ADODB.Recordset
    GetDBConnection
    tmpRec.open sqlStr, DBConn, adOpenDynamic, adLockOptimistic
    If tmpRec.EOF Then
        tmpRec.AddNew
        tmpRec("BuyRecordID") = id_
    End If
    tmpRec("VendorCarID") = Trim(vendorCarID_)
    tmpRec("BuyDate") = Trim(buyDate_)
    tmpRec("DriverName") = Trim(DriverName_)
    tmpRec("FuelQty") = fuelQty_
    tmpRec("UnitPrice") = unitPrice_
    tmpRec("isPaid") = Trim(isPaid_)
    tmpRec("RemDesc") = Trim(remDesc_)
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
    DBConnExc "Delete tbFuelBuyRecord WHERE BuyRecordID='" & id_ & "'"
    Delete = ""
    Exit Function
delErr:
    Delete = err.Description
End Function

Public Sub SetDataFromRecordset(selRec As ADODB.Recordset)
    id_ = selRec("BuyRecordID")
    vendorCarID_ = Trim("" & selRec("VendorCarID"))
    vendorName_ = Trim("" & selRec("NameTH"))
    CarRegisID_ = Trim("" & selRec("CarRegisID"))
    carTypeDesc_ = Trim("" & selRec("TruckTypeDesc") & ", " & selRec("SubTypeDesc"))
    buyDate_ = Trim("" & selRec("BuyDate"))
    DriverName_ = Trim("" & selRec("DriverName"))
    If IsNull(selRec("FuelQty")) = False Then fuelQty_ = selRec("FuelQty")
    If IsNull(selRec("UnitPrice")) = False Then unitPrice_ = selRec("UnitPrice")
    isPaid_ = Trim("" & selRec("isPaid"))
    remDesc_ = Trim("" & selRec("RemDesc"))
End Sub
