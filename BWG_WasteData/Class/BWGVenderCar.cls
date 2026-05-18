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
Dim subTypeDesc_ As String
Dim truckTypeID_ As String
Dim truckTypeDesc_ As String
Dim canBeTrailer_ As String
Dim optionDesc_ As String
Dim truckCount_ As Long
Dim contactID_ As String
Dim contactName_ As String
Dim carRegisID_ As String
Dim licenseType_ As String
Dim carRemDesc_ As String
Dim headOrTrail_ As String
Dim colTruckSubType_ As Collection
Dim coltruckType_ As Collection
Dim colSupplierContact_ As Collection
Dim isChanged_ As Boolean

Public Property Get ID() As String
    ID = id_
End Property
Public Property Let ID(ByVal newVal As String)
    id_ = Trim(newVal)
End Property

Public Property Get VendorID() As String
    VendorID = venderID_
End Property
Public Property Let VendorID(ByVal newVal As String)
    If Trim(venderID_) <> Trim(newVal) Then
        venderID_ = Trim(newVal)
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

Public Property Get VendorLicenseNo() As String
    VendorLicenseNo = vendorLicenseNo_
End Property
Public Property Let VendorLicenseNo(ByVal newVal As String)
    If Trim(vendorLicenseNo_) <> Trim(newVal) Then
        vendorLicenseNo_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get VendorWorkerCharge() As Double
    VendorWorkerCharge = vendorWorkerCharge_
End Property
Public Property Let VendorWorkerCharge(ByVal newVal As Double)
    If vendorWorkerCharge_ <> newVal Then
        vendorWorkerCharge_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get SubTypeID() As String
    SubTypeID = subTypeID_
End Property
Public Property Let SubTypeID(ByVal newVal As String)
    If Trim(subTypeID_) <> Trim(newVal) Then
        subTypeID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get SubTypeDesc() As String
    SubTypeDesc = subTypeDesc_
End Property
Public Property Let SubTypeDesc(ByVal newVal As String)
    If Trim(subTypeDesc_) <> Trim(newVal) Then
        subTypeDesc_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get TruckTypeID() As String
    TruckTypeID = truckTypeID_
End Property
Public Property Let TruckTypeID(ByVal newVal As String)
    If Trim(truckTypeID_) <> Trim(newVal) Then
        truckTypeID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get TruckTypeDesc() As String
    TruckTypeDesc = truckTypeDesc_
End Property
Public Property Let TruckTypeDesc(ByVal newVal As String)
    If Trim(truckTypeDesc_) <> Trim(newVal) Then
        truckTypeDesc_ = Trim(newVal)
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
Public Property Let CanBeTrailer(ByVal newVal As Boolean)
Dim tmpStr As String
    If newVal Then
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
Public Property Let OptionDesc(ByVal newVal As String)
    If Trim(optionDesc_) <> Trim(newVal) Then
        optionDesc_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get TruckCount() As Long
    TruckCount = truckCount_
End Property
Public Property Let TruckCount(ByVal newVal As Long)
    If truckCount_ <> newVal Then
        truckCount_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get ContactID() As String
    ContactID = contactID_
End Property
Public Property Let ContactID(ByVal newVal As String)
    If Trim(contactID_) <> Trim(newVal) Then
        contactID_ = Trim(newVal)
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

Public Property Get CarRegisID() As String
    CarRegisID = carRegisID_
End Property
Public Property Let CarRegisID(ByVal newVal As String)
    If Trim(carRegisID_) <> Trim(newVal) Then
        carRegisID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get LicenseType() As String
    LicenseType = licenseType_
End Property
Public Property Let LicenseType(ByVal newVal As String)
    If Trim(licenseType_) <> Trim(newVal) Then
        licenseType_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get Remark() As String
    Remark = carRemDesc_
End Property
Public Property Let Remark(ByVal newVal As String)
    If Trim(carRemDesc_) <> Trim(newVal) Then
        carRemDesc_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get headOrTrail() As String
    headOrTrail = headOrTrail_
End Property
Public Property Let headOrTrail(ByVal newVal As String)
    If Trim(headOrTrail_) <> Trim(newVal) Then
        headOrTrail_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get SupplierContact() As Collection
Dim tmpS As New BWGSearchManager
    If colSupplierContact_ Is Nothing Then Set colSupplierContact_ = tmpS.SupplierContactSearch("ContactID='" & id_ & "'")
    Set SupplierContact = colSupplierContact_
End Property
Public Property Set SupplierContact(ByVal newVal As Collection)
    Set colSupplierContact_ = newVal
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
Public Property Set TruckSubType(ByVal newVal As Collection)
    Set colTruckSubType_ = newVal
    isChanged_ = True
End Property
Public Sub RequeryTruckSubType()
    Set colTruckSubType_ = Nothing
End Sub

Public Property Get isChanged() As Boolean
    isChanged = isChanged_
End Property
Public Property Let isChanged(ByVal newVal As Boolean)
    isChanged_ = newVal
End Property

Public Sub Save()
On Error GoTo ErrD
Dim tmpRec As ADODB.Recordset
    
    Set tmpRec = New ADODB.Recordset
    tmpRec.Open "Select * from tbVendorCar Where  VendorTruckID='" & id_ & "'", DBManager.GetDBConnection, adOpenDynamic, adLockOptimistic
    If tmpRec.EOF Then
        tmpRec.AddNew
        tmpRec("VendorTruckID") = id_
    End If
    tmpRec("VendorID") = Trim(venderID_)
    tmpRec("SubTypeID") = Trim(subTypeID_)
    tmpRec("OptionDesc") = Trim(optionDesc_)
    tmpRec("TruckCount") = Trim(truckCount_)
    tmpRec("ContactID") = Trim(contactID_)
    tmpRec("CarRegisID") = Trim(carRegisID_)
    tmpRec("LicenseType") = Trim(licenseType_)
    tmpRec("CarRemDesc") = Trim(carRemDesc_)
    tmpRec("HeadOrTrail") = Trim(headOrTrail_)
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
Dim tmpRec As New ADODB.Recordset
    With DBManager.GetDBConnection
        Set tmpRec = GetRS("Select * from tbJobDataTimeTable Where VendorTruckID in (Select VendorTruckID from tbVendorCar Where VendorID='" & id_ & "')", adOpenForwardOnly, adLockReadOnly)
'        tmpRec.Open "Select * from tbJobDataTimeTable Where VendorTruckID in (Select VendorTruckID from tbVendorCar Where VendorID='" & id_ & "')", DBManager.GetDBConnection, adOpenKeyset
        If tmpRec.EOF Then
            .Execute "Delete tbVendorCar Where VendorID='" & id_ & "'"
            Delete = ""
        Else
            Delete = "ไม่สามารถลบข้อมูล Vendor นี้ได้เพราะมีการใช้งานรถของ Vendor นี้แล้ว"
        End If
        tmpRec.Close
        Set tmpRec = Nothing
    End With
    Exit Function
delErr:
    Delete = Err.Description
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
    carRegisID_ = Trim("" & selRec("CarRegisID"))
    subTypeDesc_ = Trim("" & selRec("SubTypeDesc"))
    truckTypeID_ = Trim("" & selRec("TruckTypeID"))
    truckTypeDesc_ = Trim("" & selRec("TruckTypeDesc"))
    canBeTrailer_ = Trim("" & selRec("CanBeTrailer"))
    contactID_ = Trim("" & selRec("ContactID"))
    contactName_ = Trim("" & selRec("ContactName"))
    licenseType_ = Trim("" & selRec("LicenseType"))
    carRemDesc_ = Trim("" & selRec("CarRemDesc"))
    headOrTrail_ = Trim("" & selRec("HeadOrTrail"))
End Sub

