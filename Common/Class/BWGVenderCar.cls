VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BWGVenderCar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim id_ As String
Dim venderID_ As String
Dim vendorName_ As String
Dim vendorLicenseNo_ As String
Dim vendorWorkerCharge_ As Double
Dim subTypeID_ As String
Dim SubTypeDesc_ As String
Dim truckTypeID_ As String
Dim TruckTypeDesc_ As String
Dim canBeTrailer_ As String
Dim optionDesc_ As String
Dim truckCount_ As Long
Dim contactID_ As String
Dim ContactName_ As String
Dim CarRegisID_ As String
Dim licenseType_ As String
Dim carRemDesc_ As String
Dim headOrTrail_ As String
Dim IsSell_ As String
Dim colTruckSubType_ As Collection
Dim coltruckType_ As Collection
Dim colSupplierContact_ As Collection
Dim isChanged_ As Boolean

Public Property Get ID() As String
    ID = id_
End Property
Public Property Let ID(ByVal newval As String)
    id_ = Trim(newval)
End Property

Public Property Get VendorID() As String
    VendorID = venderID_
End Property
Public Property Let VendorID(ByVal newval As String)
    If Trim(venderID_) <> Trim(newval) Then
        venderID_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get VendorName() As String
    VendorName = vendorName_
End Property
Public Property Let VendorName(ByVal newval As String)
    If Trim(vendorName_) <> Trim(newval) Then
        vendorName_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get VendorLicenseNo() As String
    VendorLicenseNo = vendorLicenseNo_
End Property
Public Property Let VendorLicenseNo(ByVal newval As String)
    If Trim(vendorLicenseNo_) <> Trim(newval) Then
        vendorLicenseNo_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get VendorWorkerCharge() As Double
    VendorWorkerCharge = vendorWorkerCharge_
End Property
Public Property Let VendorWorkerCharge(ByVal newval As Double)
    If vendorWorkerCharge_ <> newval Then
        vendorWorkerCharge_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get SubTypeID() As String
    SubTypeID = subTypeID_
End Property
Public Property Let SubTypeID(ByVal newval As String)
    If Trim(subTypeID_) <> Trim(newval) Then
        subTypeID_ = Trim(newval)
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

Public Property Get TruckTypeID() As String
    TruckTypeID = truckTypeID_
End Property
Public Property Let TruckTypeID(ByVal newval As String)
    If Trim(truckTypeID_) <> Trim(newval) Then
        truckTypeID_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get TruckTypeDesc() As String
    TruckTypeDesc = TruckTypeDesc_
End Property
Public Property Let TruckTypeDesc(ByVal newval As String)
    If Trim(TruckTypeDesc_) <> Trim(newval) Then
        TruckTypeDesc_ = Trim(newval)
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

Public Property Get OptionDesc() As String
    OptionDesc = optionDesc_
End Property
Public Property Let OptionDesc(ByVal newval As String)
    If Trim(optionDesc_) <> Trim(newval) Then
        optionDesc_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get TruckCount() As Long
    TruckCount = truckCount_
End Property
Public Property Let TruckCount(ByVal newval As Long)
    If truckCount_ <> newval Then
        truckCount_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get ContactID() As String
    ContactID = contactID_
End Property
Public Property Let ContactID(ByVal newval As String)
    If Trim(contactID_) <> Trim(newval) Then
        contactID_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get ContactName() As String
    ContactName = ContactName_
End Property
Public Property Let ContactName(ByVal newval As String)
    If Trim(ContactName_) <> Trim(newval) Then
        ContactName_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get CarRegisID() As String
    CarRegisID = CarRegisID_
End Property
Public Property Let CarRegisID(ByVal newval As String)
    If Trim(CarRegisID_) <> Trim(newval) Then
        CarRegisID_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get LicenseType() As String
    LicenseType = licenseType_
End Property
Public Property Let LicenseType(ByVal newval As String)
    If Trim(licenseType_) <> Trim(newval) Then
        licenseType_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get Remark() As String
    Remark = carRemDesc_
End Property
Public Property Let Remark(ByVal newval As String)
    If Trim(carRemDesc_) <> Trim(newval) Then
        carRemDesc_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get headOrTrail() As String
    headOrTrail = headOrTrail_
End Property
Public Property Let headOrTrail(ByVal newval As String)
    If Trim(headOrTrail_) <> Trim(newval) Then
        headOrTrail_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get IsSell() As String
    IsSell = IsSell_
End Property
Public Property Let IsSell(ByVal newval As String)
    If Trim(IsSell_) <> Trim(newval) Then
        IsSell_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get SupplierContact() As Collection
Dim tmpS As New BWGSearchManager
    If colSupplierContact_ Is Nothing Then Set colSupplierContact_ = tmpS.SupplierContactSearch("ContactID='" & id_ & "'")
    Set SupplierContact = colSupplierContact_
End Property
Public Property Set SupplierContact(ByVal newval As Collection)
    Set colSupplierContact_ = newval
    isChanged_ = True
End Property
Public Sub RequerySupplierContact()
    Set colSupplierContact_ = Nothing
End Sub

Public Property Get TruckSubType() As Collection
Dim tmpS As New BWGSearchManager
    If colTruckSubType_ Is Nothing Then Set colTruckSubType_ = tmpS.TruckSubTypeSearch("SubTypeID='" & id_ & "'")
    Set SupplierContact = colTruckSubType_
End Property
Public Property Set TruckSubType(ByVal newval As Collection)
    Set colTruckSubType_ = newval
    isChanged_ = True
End Property
Public Sub RequeryTruckSubType()
    Set colTruckSubType_ = Nothing
End Sub

Public Property Get isChanged() As Boolean
    isChanged = isChanged_
End Property
Public Property Let isChanged(ByVal newval As Boolean)
    isChanged_ = newval
End Property

Public Sub Save()
On Error GoTo ErrD
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
    sqlStr = "SELECT *, CONVERT(DATETIME, GETDATE(), 103) AS GETDATE FROM tbVendorCar WHERE  VendorTruckID='" & id_ & "'"
    Set tmpRec = New ADODB.Recordset
    GetDBConnection
    tmpRec.open sqlStr, DBConn, adOpenDynamic, adLockOptimistic
    If tmpRec.EOF Then
        tmpRec.AddNew
        tmpRec("VendorTruckID") = id_
    End If
    tmpRec("VendorID") = Trim(venderID_)
    tmpRec("SubTypeID") = Trim(subTypeID_)
    tmpRec("OptionDesc") = Trim(optionDesc_)
    tmpRec("TruckCount") = Trim(truckCount_)
    tmpRec("ContactID") = Trim(contactID_)
    tmpRec("CarRegisID") = Trim(CarRegisID_)
    tmpRec("LicenseType") = Trim(licenseType_)
    tmpRec("CarRemDesc") = Trim(carRemDesc_)
    tmpRec("HeadOrTrail") = Trim(headOrTrail_)
    tmpRec("IsSell") = Trim(IsSell_)
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
Dim tmpRec As New ADODB.Recordset
Dim sqlStr As String
    sqlStr = "SELECT * FROM tbJobDataTimeTable WHERE VendorTruckID in (Select VendorTruckID from tbVendorCar WHERE VendorID='" & id_ & "')"

    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec.EOF Then
        DBConnExc "Delete tbVendorCar WHERE VendorID='" & id_ & "'"
        Delete = ""
    Else
        Delete = "ไม่สามารถลบข้อมูล Vendor นี้ได้เพราะมีการใช้งานรถของ Vendor นี้แล้ว"
    End If
    tmpRec.Close
    Set tmpRec = Nothing
    
    Exit Function
delErr:
    Delete = err.Description
    If tmpRec.State = adStateOpen Then tmpRec.Close
    Set tmpRec = Nothing
End Function
', , ,

Public Sub SetDataFromRecordset(selRec As ADODB.Recordset)
    id_ = selRec("VendorTruckID")
    venderID_ = Trim("" & selRec("VendorID"))
    vendorName_ = Trim("" & selRec("NameTH"))
    vendorLicenseNo_ = Trim("" & selRec("TransLicenseNo"))
    If IsNull(selRec("WorkerCharge")) = False Then vendorWorkerCharge_ = Trim("" & selRec("WorkerCharge"))
    subTypeID_ = Trim("" & selRec("SubTypeID"))
    optionDesc_ = Trim("" & selRec("OptionDesc"))
    truckCount_ = Trim("" & selRec("TruckCount"))
    contactID_ = Trim("" & selRec("ContactID"))
    CarRegisID_ = Trim("" & selRec("CarRegisID"))
    SubTypeDesc_ = Trim("" & selRec("SubTypeDesc"))
    truckTypeID_ = Trim("" & selRec("TruckTypeID"))
    TruckTypeDesc_ = Trim("" & selRec("TruckTypeDesc"))
    canBeTrailer_ = Trim("" & selRec("CanBeTrailer"))
    contactID_ = Trim("" & selRec("ContactID"))
    ContactName_ = Trim("" & selRec("ContactName"))
    licenseType_ = Trim("" & selRec("LicenseType"))
    carRemDesc_ = Trim("" & selRec("CarRemDesc"))
    headOrTrail_ = Trim("" & selRec("HeadOrTrail"))
    IsSell_ = Trim("" & selRec("IsSell"))
End Sub

