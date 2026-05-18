VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BWGCustomerBillAddr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim id_ As String
Dim customerID_ As String
Dim billName_ As String
Dim addr1_ As String
Dim tumbolID_ As String
Dim tumbolName_ As String
Dim aumphurID_ As String
Dim aumphurName_ As String
Dim provinceID_ As String
Dim provinceName_ As String
Dim zipCode_ As String
Dim telNo_ As String
Dim faxNo_ As String
Dim contactName_ As String
Dim isChanged_ As Boolean


Public Property Get ID() As String
    ID = id_
End Property
Public Property Let ID(ByVal newVal As String)
    id_ = Trim(newVal)
End Property

Public Property Get CustomerID() As String
    CustomerID = customerID_
End Property
Public Property Let CustomerID(ByVal newVal As String)
    If Trim(customerID_) <> Trim(newVal) Then
        customerID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get BillName() As String
    BillName = billName_
End Property
Public Property Let BillName(ByVal newVal As String)
    If Trim(billName_) <> Trim(newVal) Then
        billName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get Address() As String
    Address = addr1_
End Property
Public Property Let Address(ByVal newVal As String)
    If Trim(addr1_) <> Trim(newVal) Then
        addr1_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get TumbolID() As String
    TumbolID = tumbolID_
End Property
Public Property Let TumbolID(ByVal newVal As String)
    If Trim(tumbolID_) <> Trim(newVal) Then
        tumbolID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get TumbolName() As String
    TumbolName = tumbolName_
End Property
Public Property Let TumbolName(ByVal newVal As String)
    If Trim(tumbolName_) <> Trim(newVal) Then
        tumbolName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get AumphurID() As String
    AumphurID = aumphurID_
End Property
Public Property Let AumphurID(ByVal newVal As String)
    If Trim(aumphurID_) <> Trim(newVal) Then
        aumphurID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get AumphurName() As String
    AumphurName = aumphurName_
End Property
Public Property Let AumphurName(ByVal newVal As String)
    If Trim(aumphurName_) <> Trim(newVal) Then
        aumphurName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get ProvinceID() As String
    ProvinceID = provinceID_
End Property
Public Property Let ProvinceID(ByVal newVal As String)
    If Trim(provinceID_) <> Trim(newVal) Then
        provinceID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get ProvinceName() As String
    ProvinceName = provinceName_
End Property
Public Property Let ProvinceName(ByVal newVal As String)
    If Trim(provinceName_) <> Trim(newVal) Then
        provinceName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get ZipCode() As String
    ZipCode = zipCode_
End Property
Public Property Let ZipCode(ByVal newVal As String)
    If Trim(zipCode_) <> Trim(newVal) Then
        zipCode_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get TelNo() As String
    TelNo = telNo_
End Property
Public Property Let TelNo(ByVal newVal As String)
    If Trim(telNo_) <> Trim(newVal) Then
        telNo_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get FaxNo() As String
    FaxNo = faxNo_
End Property
Public Property Let FaxNo(ByVal newVal As String)
    If Trim(faxNo_) <> Trim(newVal) Then
        faxNo_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get ContactName() As String
    ContactName = contactName_
End Property
Public Property Let ContactName(ByVal newVal As String)
    If Trim(contactName_) <> Trim(newVal) Then
        contactName_ = Trim(newVal)
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
    tmpRec.Open "Select * from tbCustomerBillAddr Where BillAddrID='" & id_ & "'", DBManager.GetDBConnection, adOpenDynamic, adLockOptimistic
    If tmpRec.EOF Then
        tmpRec.AddNew
        tmpRec("BillAddrID") = id_
    End If
    tmpRec("CustomerID") = Trim(customerID_)
    tmpRec("BillName") = Trim(billName_)
    tmpRec("Addr1") = Trim(addr1_)
    tmpRec("TumbolID") = Trim(tumbolID_)
    tmpRec("ZipCode") = Trim(zipCode_)
    tmpRec("TelNo") = Trim(telNo_)
    tmpRec("FaxNo") = Trim(faxNo_)
    tmpRec("ContactName") = Trim(contactName_)
    tmpRec.Update
    tmpRec.Close
    Set tmpRec = Nothing
    isChanged_ = False
End Sub

Public Function Delete() As String
On Error GoTo delErr:
    DBManager.GetDBConnection.Execute "Delete tbCustomerBillAddr Where BillAddrID='" & id_ & "'"
    Delete = ""
    Exit Function
delErr:
    Delete = Err.Description
End Function

Public Sub SetDataFromRecordset(selRec As ADODB.Recordset)
    id_ = Trim("" & selRec("BillAddrID"))
    customerID_ = Trim("" & selRec("CustomerID"))
    billName_ = Trim("" & selRec("BillName"))
    addr1_ = Trim("" & selRec("Addr1"))
    tumbolID_ = Trim("" & selRec("TumbolID"))
    tumbolName_ = Trim("" & selRec("TumbolName"))
    aumphurID_ = Trim("" & selRec("AumphurID"))
    aumphurName_ = Trim("" & selRec("AumphurName"))
    provinceID_ = Trim("" & selRec("ProvinceID"))
    provinceName_ = Trim("" & selRec("ProvinceName"))
    zipCode_ = Trim("" & selRec("ZipCode"))
    telNo_ = Trim("" & selRec("TelNo"))
    faxNo_ = Trim("" & selRec("FaxNo"))
    contactName_ = Trim("" & selRec("ContactName"))
End Sub
