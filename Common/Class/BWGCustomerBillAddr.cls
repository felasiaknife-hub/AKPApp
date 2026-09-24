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
Dim CustomerID_ As String
Dim billName_ As String
Dim addr1_ As String
Dim tumbolID_ As String
Dim tumbolName_ As String
Dim aumphurID_ As String
Dim aumphurName_ As String
Dim provinceID_ As String
Dim provinceName_ As String
Dim zipCode_ As String
Dim Country_ As String
Dim TelNo_ As String
Dim faxNo_ As String
Dim ContactName_ As String
Dim isChanged_ As Boolean
Dim TaxNum_ As String

Dim LastUser_ As String
Dim LastUpdate_ As String

Public Property Get ID() As String
    ID = id_
End Property
Public Property Let ID(ByVal newVal As String)
    id_ = Trim(newVal)
End Property

Public Property Get CustomerID() As String
    CustomerID = CustomerID_
End Property
Public Property Let CustomerID(ByVal newVal As String)
    If Trim(CustomerID_) <> Trim(newVal) Then
        CustomerID_ = Trim(newVal)
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

Public Property Get Country() As String
    Country = Country_
End Property
Public Property Let Country(ByVal newVal As String)
    If Trim(Country_) <> Trim(newVal) Then
        Country_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get TelNo() As String
    TelNo = TelNo_
End Property
Public Property Let TelNo(ByVal newVal As String)
    If Trim(TelNo_) <> Trim(newVal) Then
        TelNo_ = Trim(newVal)
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
    ContactName = ContactName_
End Property
Public Property Let ContactName(ByVal newVal As String)
    If Trim(ContactName_) <> Trim(newVal) Then
        ContactName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get LastUser() As String
    LastUser = LastUser_
End Property
Public Property Let LastUser(ByVal newVal As String)
    If Trim(LastUser_) <> Trim(newVal) Then
        LastUser_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get LastUpdate() As String
    LastUpdate = LastUpdate_
End Property
Public Property Let LastUpdate(ByVal newVal As String)
    If Trim(LastUpdate_) <> Trim(newVal) Then
        LastUpdate_ = Trim(newVal)
        isChanged_ = True
    End If
End Property
Public Property Get TaxNum() As String
    TaxNum = TaxNum_
End Property
Public Property Let TaxNum(ByVal newVal As String)
    If Trim(TaxNum_) <> Trim(newVal) Then
        TaxNum_ = Trim(newVal)
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
    sqlStr = "SELECT *, CONVERT(DATETIME, GETDATE(), 103) AS GETDATE FROM tbCustomerBillAddr WHERE BillAddrID='" & id_ & "'"
    Set tmpRec = New ADODB.Recordset
    GetDBConnection
    tmpRec.open sqlStr, DBConn, adOpenDynamic, adLockOptimistic
    If tmpRec.EOF Then
        tmpRec.AddNew
        tmpRec("BillAddrID") = id_
    End If
    tmpRec("CustomerID") = Trim(CustomerID_)
    tmpRec("BillName") = Trim(billName_)
    tmpRec("Addr1") = Trim(addr1_)
    tmpRec("TumbolID") = Trim(tumbolID_)
    tmpRec("ZipCode") = Trim(zipCode_)
    tmpRec("Country") = Trim(Country_)
    tmpRec("TelNo") = Trim(TelNo_)
    tmpRec("FaxNo") = Trim(faxNo_)
    tmpRec("ContactName") = Trim(ContactName_)
    tmpRec("TaxNum") = Trim(TaxNum_)
    
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
    DBConnExc "Delete tbCustomerBillAddr WHERE BillAddrID='" & id_ & "'"
    Delete = ""
    Exit Function
delErr:
    Delete = err.Description
End Function

Public Sub SetDataFromRecordset(selRec As ADODB.Recordset)
    id_ = Trim("" & selRec("BillAddrID"))
    CustomerID_ = Trim("" & selRec("CustomerID"))
    billName_ = Trim("" & selRec("BillName"))
    addr1_ = Trim("" & selRec("Addr1"))
    tumbolID_ = Trim("" & selRec("TumbolID"))
    tumbolName_ = Trim("" & selRec("TumbolName"))
    aumphurID_ = Trim("" & selRec("AumphurID"))
    aumphurName_ = Trim("" & selRec("AumphurName"))
    provinceID_ = Trim("" & selRec("ProvinceID"))
    provinceName_ = Trim("" & selRec("ProvinceName"))
    zipCode_ = Trim("" & selRec("ZipCode"))
    Country_ = Trim("" & selRec("Country"))
    TelNo_ = Trim("" & selRec("TelNo"))
    faxNo_ = Trim("" & selRec("FaxNo"))
    ContactName_ = Trim("" & selRec("ContactName"))
    TaxNum_ = Trim("" & selRec("TaxNum"))
    If isExist(selRec, "LastUser") Then LastUser_ = Trim("" & selRec!LastUser)
    If isExist(selRec, "LastUpdate") Then LastUpdate_ = Trim("" & selRec!LastUpdate)
End Sub
