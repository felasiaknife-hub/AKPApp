VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BWGJobDataTimeTableBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim JobDataCarID_ As String
Dim TripNo_ As Integer
Dim startTime_ As String
Dim VendorTruckID_ As String
Dim vendorID_ As String
Dim vendorName_ As String
Dim vendorRegisNo_ As String
Dim vendorWorkerCharge_ As Double
Dim carRegisNo_ As String
Dim TrailerTruckNo_ As String
Dim DriverName_ As String
Dim MenifestID_ As String
Dim workerCount_ As Integer
Dim workDate_ As String
Dim TimeTableNo_ As String
Dim appointDate_ As String
Dim appointTime_ As String
Dim custInDate_ As String
Dim custInTime_ As String
Dim custOutDate_ As String
Dim custOutTime_ As String
Dim backDate_ As String
Dim backTime_ As String
Dim isCancel_ As String
Dim cancelDate_ As String
Dim cancelStaffID_ As String
Dim cancelStaffName_ As String
Dim trailTruckID_ As String
Dim custBillID_ As String
Dim BillingNo_ As String
Dim isTrailer_ As String
Dim TrailerType_ As String
'''''''''''''''''''''''
Dim tripTranFee_ As Double
Dim tripMoreWorker_ As Integer
Dim tripMoreWorkerFee_ As Double
Dim tripEquipFee_ As Double
Dim tripOtherDesc_ As String
Dim tripOtherFee_ As Double
Dim tripAllFee_ As Double
Dim tripTreatFee_ As Double
Dim IncTripTransFee_ As Double
Dim EditTypeTrip_ As String
''''''''''''''''''''''
Dim vdTranFee_ As Double
Dim vdWorker_ As Integer
Dim vdWorkerFee_ As Double
Dim vdEquipFee_ As Double
Dim vdOtherDesc_ As String
Dim vdOtherFee_ As Double
Dim EditTypeVD_ As String
''''''''''''''''''''''
Dim truckTypeID_ As String
Dim truckTypeDesc_ As String
Dim truckSubTypeID_ As String
Dim truckSubTypeDesc_ As String
Dim fromAumphurID_ As String
Dim fromAumphurName_ As String
Dim fromProvinceID_ As String
Dim fromProvinceName_ As String
Dim toAumphurID_ As String
Dim toAumphurName_ As String
Dim toProvinceID_ As String
Dim toProvinceName_ As String

''''''''''''''''''''''
Dim isMain_ As String
Dim BoxType_ As String
Dim BoxCount_ As Integer

Dim colJobDet_ As Collection
Dim colMenifest_ As Collection

Dim CompanyCode_ As String
Dim CustomerName_ As String
Dim SiteAddress_ As String
Dim TelNo_ As String
Dim ContactName_ As String
Dim SaleStaffName_ As String
Dim StaffTelNo_ As String
Dim CarRem_ As String

Dim isChanged_ As Boolean

Public Property Get CompanyCode() As String
    CompanyCode = CompanyCode_
End Property
Public Property Let CompanyCode(ByVal newVal As String)
    If CompanyCode_ <> Trim(newVal) Then
        CompanyCode_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get JobDataCarID() As String
    JobDataCarID = JobDataCarID_
End Property
Public Property Let JobDataCarID(ByVal newVal As String)
    If JobDataCarID_ <> Trim(newVal) Then
        JobDataCarID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get TripNo() As Integer
    TripNo = TripNo_
End Property
Public Property Let TripNo(ByVal newVal As Integer)
    If TripNo_ <> newVal Then
        TripNo_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get StartTime() As String
    StartTime = startTime_
End Property
Public Property Let StartTime(ByVal newVal As String)
    If startTime_ <> Trim(newVal) Then
        startTime_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get VendorTruckID() As String
    VendorTruckID = VendorTruckID_
End Property
Public Property Let VendorTruckID(ByVal newVal As String)
    If VendorTruckID_ <> Trim(newVal) Then
        VendorTruckID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get VendorID() As String
    VendorID = vendorID_
End Property
Public Property Let VendorID(ByVal newVal As String)
    If vendorID_ <> Trim(newVal) Then
        vendorID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get VendorName() As String
    VendorName = vendorName_
End Property
Public Property Let VendorName(ByVal newVal As String)
    If vendorName_ <> Trim(newVal) Then
        vendorName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get VendorRegisNo() As String
    VendorRegisNo = vendorRegisNo_
End Property
Public Property Let VendorRegisNo(ByVal newVal As String)
    If vendorRegisNo_ <> Trim(newVal) Then
        vendorRegisNo_ = Trim(newVal)
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

Public Property Get CarRegisNo() As String
    CarRegisNo = carRegisNo_
End Property
Public Property Let CarRegisNo(ByVal newVal As String)
    If carRegisNo_ <> Trim(newVal) Then
        carRegisNo_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get TrailerTruckNo() As String
    TrailerTruckNo = TrailerTruckNo_
End Property
Public Property Let TrailerTruckNo(ByVal newVal As String)
    If TrailerTruckNo_ <> Trim(newVal) Then
        TrailerTruckNo_ = Trim(newVal)
    End If
End Property

Public Property Get DriverName() As String
    DriverName = DriverName_
End Property
Public Property Let DriverName(ByVal newVal As String)
    If DriverName_ <> Trim(newVal) Then
        DriverName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get menifestID() As String
    menifestID = MenifestID_
End Property
Public Property Let menifestID(ByVal newVal As String)
    If MenifestID_ <> Trim(newVal) Then
        MenifestID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get WorkerCount() As Integer
    WorkerCount = workerCount_
End Property
Public Property Let WorkerCount(ByVal newVal As Integer)
    If workerCount_ <> newVal Then
        workerCount_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get WorkDate() As String
    WorkDate = workDate_
End Property
Public Property Let WorkDate(ByVal newVal As String)
    If workDate_ <> Trim(newVal) Then
        workDate_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get TimeTableNo() As String
    TimeTableNo = TimeTableNo_
End Property
Public Property Let TimeTableNo(ByVal newVal As String)
    If TimeTableNo_ <> Trim(newVal) Then
        TimeTableNo_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get AppointDate() As String
    AppointDate = appointDate_
End Property
Public Property Let AppointDate(ByVal newVal As String)
    If appointDate_ <> Trim(newVal) Then
        appointDate_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get AppointTime() As String
    AppointTime = appointTime_
End Property
Public Property Let AppointTime(ByVal newVal As String)
    If appointTime_ <> Trim(newVal) Then
        appointTime_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get CustInDate() As String
    CustInDate = custInDate_
End Property
Public Property Let CustInDate(ByVal newVal As String)
    If custInDate_ <> Trim(newVal) Then
        custInDate_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get CustInTime() As String
    CustInTime = custInTime_
End Property
Public Property Let CustInTime(ByVal newVal As String)
    If custInTime_ <> Trim(newVal) Then
        custInTime_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get CustOutDate() As String
    CustOutDate = custOutDate_
End Property
Public Property Let CustOutDate(ByVal newVal As String)
    If custOutDate_ <> Trim(newVal) Then
        custOutDate_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get CustOutTime() As String
    CustOutTime = custOutTime_
End Property
Public Property Let CustOutTime(ByVal newVal As String)
    If custOutTime_ <> Trim(newVal) Then
        custOutTime_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get backDate() As String
    backDate = backDate_
End Property
Public Property Let backDate(ByVal newVal As String)
    If backDate_ <> Trim(newVal) Then
        backDate_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get BackTime() As String
    BackTime = backTime_
End Property
Public Property Let BackTime(ByVal newVal As String)
    If backTime_ <> Trim(newVal) Then
        backTime_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get isCancel() As Boolean
    If Trim(isCancel_) = "Y" Then
        isCancel = True
    Else
        isCancel = False
    End If
End Property
Public Property Let isCancel(ByVal newVal As Boolean)
Dim tmpStr As String
    If newVal Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If isCancel_ <> Trim(tmpStr) Then
        isCancel_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property

Public Property Get CancelDate() As String
    CancelDate = cancelDate_
End Property
Public Property Let CancelDate(ByVal newVal As String)
    If cancelDate_ <> Trim(newVal) Then
        cancelDate_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get CancelStaffID() As String
    CancelStaffID = cancelStaffID_
End Property
Public Property Let CancelStaffID(ByVal newVal As String)
    If cancelStaffID_ <> Trim(newVal) Then
        cancelStaffID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get CancelStaffName() As String
    CancelStaffName = cancelStaffName_
End Property
Public Property Let CancelStaffName(ByVal newVal As String)
    If cancelStaffName_ <> Trim(newVal) Then
        cancelStaffName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get TrailTruckID() As String
    TrailTruckID = trailTruckID_
End Property
Public Property Let TrailTruckID(ByVal newVal As String)
    If trailTruckID_ <> Trim(newVal) Then
        trailTruckID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get CustBillID() As String
    CustBillID = custBillID_
End Property
Public Property Let CustBillID(ByVal newVal As String)
    If custBillID_ <> Trim(newVal) Then
        custBillID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get BillingNo() As String
    BillingNo = BillingNo_
End Property
Public Property Let BillingNo(ByVal newVal As String)
    If BillingNo_ <> Trim(newVal) Then
        BillingNo_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get isTrailer() As Boolean
    If Trim(isTrailer_) = "Y" Then
        isTrailer = True
    Else
        isTrailer = False
    End If
End Property
Public Property Let isTrailer(ByVal newVal As Boolean)
Dim tmpStr As String
    If newVal Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If isTrailer_ <> Trim(tmpStr) Then
        isTrailer_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property

Public Property Get TrailerType() As String
    TrailerType = TrailerType_
End Property
Public Property Let TrailerType(ByVal newVal As String)
    If TrailerType_ <> newVal Then
        TrailerType_ = newVal
    End If
End Property

Public Property Get TripTranFee() As Double
    TripTranFee = tripTranFee_
End Property
Public Property Let TripTranFee(ByVal newVal As Double)
    If tripTranFee_ <> newVal Then
        tripTranFee_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get TripMoreWorker() As Double
    TripMoreWorker = tripMoreWorker_
End Property
Public Property Let TripMoreWorker(ByVal newVal As Double)
    If tripMoreWorker_ <> newVal Then
        tripMoreWorker_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get TripMoreWorkerFee() As Double
    TripMoreWorkerFee = tripMoreWorkerFee_
End Property
Public Property Let TripMoreWorkerFee(ByVal newVal As Double)
    If tripMoreWorkerFee_ <> newVal Then
        tripMoreWorkerFee_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get TripEquipFee() As Double
    TripEquipFee = tripEquipFee_
End Property
Public Property Let TripEquipFee(ByVal newVal As Double)
    If tripEquipFee_ <> newVal Then
        tripEquipFee_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get TripOtherDesc() As String
    TripOtherDesc = Trim(tripOtherDesc_)
End Property
Public Property Let TripOtherDesc(ByVal newVal As String)
    If Trim(tripOtherDesc_) <> Trim(newVal) Then
        tripOtherDesc_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get TripOtherFee() As Double
    TripOtherFee = tripOtherFee_
End Property
Public Property Let TripOtherFee(ByVal newVal As Double)
    If tripOtherFee_ <> newVal Then
        tripOtherFee_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get TripAllFee() As Double
    TripAllFee = tripAllFee_
End Property
Public Property Let TripAllFee(ByVal newVal As Double)
    If tripAllFee_ <> newVal Then
        tripAllFee_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get TripTreatFee() As Double
    TripTreatFee = tripTreatFee_
End Property
Public Property Let TripTreatFee(ByVal newVal As Double)
    If tripTreatFee_ <> newVal Then
        tripTreatFee_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get IncTripTransFee() As Double
    IncTripTransFee = IncTripTransFee_
End Property
Public Property Let IncTripTransFee(ByVal newVal As Double)
    If IncTripTransFee_ <> newVal Then
        IncTripTransFee_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get EditTypeTrip() As String
    EditTypeTrip = EditTypeTrip_
End Property
Public Property Let EditTypeTrip(ByVal newVal As String)
Dim tmpStr As String
        If newVal = "" Then
            tmpStr = "A"
        Else
            tmpStr = Trim(newVal)
        End If
    If Trim(EditTypeTrip_) <> Trim(tmpStr) Then
        EditTypeTrip_ = Trim(tmpStr)
    End If
End Property

Public Property Get VDTranFee() As Double
    VDTranFee = vdTranFee_
End Property
Public Property Let VDTranFee(ByVal newVal As Double)
    If vdTranFee_ <> newVal Then
        vdTranFee_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get VDWorker() As Double
    VDWorker = vdWorker_
End Property
Public Property Let VDWorker(ByVal newVal As Double)
    If vdWorker_ <> newVal Then
        vdWorker_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get VDWorkerFee() As Double
    VDWorkerFee = vdWorkerFee_
End Property
Public Property Let VDWorkerFee(ByVal newVal As Double)
    If vdWorkerFee_ <> newVal Then
        vdWorkerFee_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get VDEquipFee() As Double
    VDEquipFee = vdEquipFee_
End Property
Public Property Let VDEquipFee(ByVal newVal As Double)
    If vdEquipFee_ <> newVal Then
        vdEquipFee_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get VDOtherDesc() As String
    VDOtherDesc = Trim(vdOtherDesc_)
End Property
Public Property Let VDOtherDesc(ByVal newVal As String)
    If Trim(vdOtherDesc_) <> Trim(newVal) Then
        vdOtherDesc_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get VDOtherFee() As Double
    VDOtherFee = vdOtherFee_
End Property
Public Property Let VDOtherFee(ByVal newVal As Double)
    If vdOtherFee_ <> newVal Then
        vdOtherFee_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get EditTypeVD() As String
    EditTypeVD = EditTypeVD_
End Property
Public Property Let EditTypeVD(ByVal newVal As String)
Dim tmpStr As String
        If newVal = "" Then
            tmpStr = "A"
        Else
            tmpStr = Trim(newVal)
        End If
    If Trim(EditTypeVD_) <> Trim(tmpStr) Then
        EditTypeVD_ = Trim(tmpStr)
    End If
End Property

Public Property Get TruckTypeID() As String
    TruckTypeID = Trim(truckTypeID_)
End Property
Public Property Let TruckTypeID(ByVal newVal As String)
    If Trim(truckTypeID_) <> Trim(newVal) Then
        truckTypeID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get TruckTypeDesc() As String
    TruckTypeDesc = Trim(truckTypeDesc_)
End Property
Public Property Let TruckTypeDesc(ByVal newVal As String)
    If Trim(truckTypeDesc_) <> Trim(newVal) Then
        truckTypeDesc_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get TruckSubTypeID() As String
    TruckSubTypeID = Trim(truckSubTypeID_)
End Property
Public Property Let TruckSubTypeID(ByVal newVal As String)
    If Trim(truckSubTypeID_) <> Trim(newVal) Then
        truckSubTypeID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get TruckSubTypeDesc() As String
    TruckSubTypeDesc = Trim(truckSubTypeDesc_)
End Property
Public Property Let TruckSubTypeDesc(ByVal newVal As String)
    If Trim(truckSubTypeDesc_) <> Trim(newVal) Then
        truckSubTypeDesc_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get FromAumphurID() As String
    FromAumphurID = Trim(fromAumphurID_)
End Property
Public Property Let FromAumphurID(ByVal newVal As String)
    If Trim(fromAumphurID_) <> Trim(newVal) Then
        fromAumphurID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get FromAumphurName() As String
    FromAumphurName = Trim(fromAumphurName_)
End Property
Public Property Let FromAumphurName(ByVal newVal As String)
    If Trim(fromAumphurName_) <> Trim(newVal) Then
        fromAumphurName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get FromProvinceID() As String
    FromProvinceID = Trim(fromProvinceID_)
End Property
Public Property Let FromProvinceID(ByVal newVal As String)
    If Trim(fromProvinceID_) <> Trim(newVal) Then
        fromProvinceID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get FromProvinceName() As String
    FromProvinceName = Trim(fromProvinceName_)
End Property
Public Property Let FromProvinceName(ByVal newVal As String)
    If Trim(fromProvinceName_) <> Trim(newVal) Then
        fromProvinceName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get ToAumphurID() As String
    ToAumphurID = Trim(toAumphurID_)
End Property
Public Property Let ToAumphurID(ByVal newVal As String)
    If Trim(toAumphurID_) <> Trim(newVal) Then
        toAumphurID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get ToAumphurName() As String
    ToAumphurName = Trim(toAumphurName_)
End Property
Public Property Let ToAumphurName(ByVal newVal As String)
    If Trim(toAumphurName_) <> Trim(newVal) Then
        toAumphurName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get ToProvinceID() As String
    ToProvinceID = Trim(toProvinceID_)
End Property
Public Property Let ToProvinceID(ByVal newVal As String)
    If Trim(toProvinceID_) <> Trim(newVal) Then
        toProvinceID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get ToProvinceName() As String
    ToProvinceName = Trim(toProvinceName_)
End Property
Public Property Let ToProvinceName(ByVal newVal As String)
    If Trim(toProvinceName_) <> Trim(newVal) Then
        toProvinceName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get isMain() As Boolean
    If Trim(isMain_) = "Y" Then
        isMain = True
    Else
        isMain = False
    End If
End Property
Public Property Let isMain(ByVal newVal As Boolean)
Dim tmpStr As String
    If newVal Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(isMain_) <> Trim(tmpStr) Then
        isMain_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property

Public Property Get BoxType() As String
    BoxType = BoxType_
End Property
Public Property Let BoxType(ByVal newVal As String)
    If Trim(BoxType_) <> Trim(newVal) Then
        BoxType_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get BoxCount() As Integer
    BoxCount = BoxCount_
End Property
Public Property Let BoxCount(ByVal newVal As Integer)
    If BoxCount_ <> newVal Then
        BoxCount_ = newVal
        isChanged_ = True
    End If
End Property

'Public Property Get JobDetails() As Collection
'Dim tmpS As New BWGSearchManager
'    'If colJobDet_ Is Nothing Then Set colJobDet_ = tmpS.JobDataDetailSearch("JobID in (Select CreatedFromJobID from tbMenifestHeader WHERE JobDataCarID='" & jobDataCarID_ & "' And TripNo=" & tripNo_ & " and (isCanceled<>'Y' Or isCanceled is null))")
'    If colJobDet_ Is Nothing Then Set colJobDet_ = tmpS.JobDataDetailSearch("JobDetID in (Select JobDetID from vw_JobDetIDFromTrip WHERE JobDataCarID='" & JobDataCarID_ & "' And TripNo=" & TripNo_ & " and (isCanceled<>'Y' Or isCanceled is null))")
'    Set JobDetails = colJobDet_
'End Property
'Public Property Set JobDetails(ByVal newVal As Collection)
'    Set colJobDet_ = newVal
'    isChanged_ = True
'End Property

Public Property Get CustomerName() As String
    CustomerName = Trim(CustomerName_)
End Property
Public Property Let CustomerName(ByVal newVal As String)
    If Trim(CustomerName_) <> Trim(newVal) Then
        CustomerName_ = Trim(newVal)
    End If
End Property

Public Property Get SiteAddress() As String
    SiteAddress = Trim(SiteAddress_)
End Property
Public Property Let SiteAddress(ByVal newVal As String)
    If Trim(SiteAddress_) <> Trim(newVal) Then
        SiteAddress_ = Trim(newVal)
    End If
End Property

Public Property Get TelNo() As String
    TelNo = Trim(TelNo_)
End Property
Public Property Let TelNo(ByVal newVal As String)
    If Trim(TelNo_) <> Trim(newVal) Then
        TelNo_ = Trim(newVal)
    End If
End Property

Public Property Get ContactName() As String
    ContactName = Trim(ContactName_)
End Property
Public Property Let ContactName(ByVal newVal As String)
    If Trim(ContactName_) <> Trim(newVal) Then
        ContactName_ = Trim(newVal)
    End If
End Property

Public Property Get SaleStaffName() As String
    SaleStaffName = Trim(SaleStaffName_)
End Property
Public Property Let SaleStaffName(ByVal newVal As String)
    If Trim(SaleStaffName_) <> Trim(newVal) Then
        SaleStaffName_ = Trim(newVal)
    End If
End Property

Public Property Get StaffTelNo() As String
    StaffTelNo = Trim(StaffTelNo_)
End Property
Public Property Let StaffTelNo(ByVal newVal As String)
    If Trim(StaffTelNo_) <> Trim(newVal) Then
        StaffTelNo_ = Trim(newVal)
    End If
End Property

Public Property Get CarRem() As String
    CarRem = Trim(CarRem_)
End Property
Public Property Let CarRem(ByVal newVal As String)
    If Trim(CarRem_) <> Trim(newVal) Then
        CarRem_ = Trim(newVal)
    End If
End Property

Public Property Get isChanged() As Boolean
    isChanged = isChanged_
End Property
Public Property Let isChanged(ByVal newVal As Boolean)
    isChanged_ = newVal
End Property

Public Sub SetDataFromRecordset(ByVal selRec As ADODB.Recordset, Optional ByVal isTruck As Boolean = False)
    If isTruck = False Then JobDataCarID_ = Trim("" & selRec!JobDataCarID)
    If isTruck = True Then CompanyCode_ = Trim("" & selRec!CompanyCode)
    If IsNull(selRec("TripNo")) = False Then TripNo_ = selRec("TripNo")
    startTime_ = Trim("" & selRec("StartTime"))
    
    VendorTruckID_ = Trim("" & selRec("VendorTruckID"))
    vendorID_ = Trim("" & selRec("VendorID"))
    vendorName_ = Trim("" & selRec("NameTH"))
    carRegisNo_ = Trim("" & selRec("CarRegisID"))
    If isExist(selRec, "TrailerTruckNo") Then TrailerTruckNo_ = Trim("" & selRec("TrailerTruckNo"))
    DriverName_ = Trim("" & selRec("DriverName"))
    MenifestID_ = Trim("" & selRec("MenifestID"))
    vendorRegisNo_ = Trim("" & selRec("TransLicenseNo"))
    If IsNull(selRec("WorkerCharge")) = False Then vendorWorkerCharge_ = Trim("" & selRec("WorkerCharge"))
    
    If IsNull(selRec("WorkerCount")) = False Then workerCount_ = selRec("WorkerCount")
    workDate_ = Trim("" & selRec("WorkDate"))
    TimeTableNo_ = Trim("" & selRec("TimeTableNo"))
    If isExist(selRec, "AppointDate") Then appointDate_ = Trim("" & selRec("AppointDate"))
    If isExist(selRec, "AppointTime") Then appointTime_ = Trim("" & selRec("AppointTime"))
    If isExist(selRec, "CustInDate") Then custInDate_ = Trim("" & selRec("CustInDate"))
    If isExist(selRec, "CustInTime") Then custInTime_ = Trim("" & selRec("CustInTime"))
    If isExist(selRec, "CustOutDate") Then custOutDate_ = Trim("" & selRec("CustOutDate"))
    If isExist(selRec, "CustOutTime") Then custOutTime_ = Trim("" & selRec("CustOutTime"))
    If isExist(selRec, "BackDate") Then backDate_ = Trim("" & selRec("BackDate"))
    If isExist(selRec, "BackTime") Then backTime_ = Trim("" & selRec("BackTime"))
    
    isCancel_ = Trim("" & selRec("isCancel"))
'    cancelDate_ = Trim("" & selRec("CancelDate"))
'    cancelStaffID_ = Trim("" & selRec("CancelStaffID"))
    cancelStaffName_ = Trim("" & selRec("StaffFName")) & " " & Trim("" & selRec("StaffLName"))
    trailTruckID_ = Trim("" & selRec("TrailTruckID"))
    If isExist(selRec, "CustBillID") Then custBillID_ = Trim("" & selRec("CustBillID"))
    If isExist(selRec, "BillingNo") Then BillingNo_ = Trim("" & selRec("BillingNo"))
    If isExist(selRec, "isMain") Then isMain_ = Trim("" & selRec("isMain"))
    
    If isExist(selRec, "TripTranFee") Then If IsNull(selRec("TripTranFee")) = False Then tripTranFee_ = selRec("TripTranFee")
    If isExist(selRec, "TripMoreWorker") Then If IsNull(selRec("TripMoreWorker")) = False Then tripMoreWorker_ = selRec("TripMoreWorker")
    If isExist(selRec, "TripMoreWorkerFee") Then If IsNull(selRec("TripMoreWorkerFee")) = False Then tripMoreWorkerFee_ = selRec("TripMoreWorkerFee")
    If isExist(selRec, "TripEquipFee") Then tripEquipFee_ = IIf(IsNull(selRec!TripEquipFee), 0, selRec!TripEquipFee)
    If isExist(selRec, "TripOtherDesc") Then tripOtherDesc_ = IIf(IsNull(selRec!TripOtherDesc), 0, selRec!TripOtherDesc)
    If isExist(selRec, "TripOtherFee") Then tripOtherFee_ = IIf(IsNull(selRec!TripOtherFee), 0, selRec!TripOtherFee)
    If isExist(selRec, "TripAllFee") Then tripAllFee_ = IIf(IsNull(selRec!TripAllFee), 0, selRec!TripAllFee)
    If isExist(selRec, "TripTreatFee") Then tripTreatFee_ = IIf(IsNull(selRec!TripTreatFee), 0, selRec!TripTreatFee)
    If isExist(selRec, "TripTransFee_Inc") Then IncTripTransFee_ = IIf(IsNull(selRec!TripTransFee_Inc), 0, selRec!TripTransFee_Inc)
    
    If isExist(selRec, "VDTransFee") Then If IsNull(selRec("VDTransFee")) = False Then vdTranFee_ = selRec("VDTransFee")
    If isExist(selRec, "VDWorker") Then If IsNull(selRec("VDWorker")) = False Then vdWorker_ = selRec("VDWorker")
    If isExist(selRec, "VDWorkerFee") Then If IsNull(selRec("VDWorkerFee")) = False Then vdWorkerFee_ = selRec("VDWorkerFee")
    If isExist(selRec, "VDEquipFee") Then If IsNull(selRec("VDEquipFee")) = False Then vdEquipFee_ = selRec("VDEquipFee")
    If isExist(selRec, "VDOtherDesc") Then If IsNull(selRec("VDOtherDesc")) = False Then vdOtherDesc_ = Trim("" & selRec("VDOtherDesc"))
    If isExist(selRec, "VDOtherFee") Then If IsNull(selRec("VDOtherFee")) = False Then vdOtherFee_ = selRec("VDOtherFee")
    
    If isExist(selRec, "isTrailer") Then isTrailer_ = Trim(selRec("isTrailer"))
    If isExist(selRec, "HeadOrTrail") Then TrailerType_ = Trim("" & selRec("HeadOrTrail"))

    truckTypeID_ = Trim("" & selRec("TruckTypeID"))
    truckTypeDesc_ = Trim("" & selRec("TruckTypeDesc"))
    truckSubTypeID_ = Trim("" & selRec("SubTypeID"))
    truckSubTypeDesc_ = Trim("" & selRec("SubTypeDesc"))
    fromAumphurID_ = Trim("" & selRec("FromAumphurID"))
    fromAumphurName_ = Trim("" & selRec("FromAumphurName"))
    fromProvinceID_ = Trim("" & selRec("FromProvinceID"))
    fromProvinceName_ = Trim("" & selRec("FromProvinceName"))
'*** ไม่ต้องระบุเนื่องจากเป็นการวางบ็อกไม่ต้องระบุปลายทาง
'    toAumphurID_ = Trim("" & selRec("ToAumphurID"))
'    toAumphurName_ = Trim("" & selRec("ToAumphurName"))
'    toProvinceID_ = Trim("" & selRec("ToProvinceID"))
'    toProvinceName_ = Trim("" & selRec("ToProvinceName"))
    
    If isExist(selRec, "BoxType") Then BoxType_ = Trim("" & selRec!BoxType)
    If isExist(selRec, "BoxCount") Then BoxCount_ = IIf(IsNull(selRec!BoxCount), 0, selRec!BoxCount)
    
    If isExist(selRec, "EditTypeTrip") Then EditTypeTrip_ = Trim("" & IIf(IsNull(selRec!EditTypeTrip), "A", selRec!EditTypeTrip))
    If isExist(selRec, "EditTypeVD") Then EditTypeVD_ = Trim("" & IIf(IsNull(selRec!EditTypeVD), "A", selRec!EditTypeVD))
    
    If isExist(selRec, "CompanyName") Then CustomerName_ = Trim("" & selRec!CompanyName)
    SiteAddress_ = Trim("" & selRec!Addr1) & " " & Trim("" & selRec!FromTumbolName) & " " & Trim("" & selRec!FromAumphurName) & " " & Trim("" & selRec!FromProvinceName)
    TelNo_ = Trim("" & selRec!TelNo)
    ContactName_ = Trim("" & selRec!ContactName)
    SaleStaffName_ = Trim("" & selRec!SaleStaffName)
    StaffTelNo_ = Trim("" & selRec!SaleStaffTelNo)
    CarRem_ = Trim("" & selRec!CarRem)
End Sub

Public Sub Save(Optional ByVal WorkDate As String = "", Optional ByVal AutoCal As Boolean = False)
On Error GoTo ErrD
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
    sqlStr = "SELECT *, CONVERT(DATETIME, GETDATE(), 103) AS GETDATE FROM tbJobDataTimeTableBox WHERE JobDataCarID='" & JobDataCarID_ & "' And TripNo=" & TripNo_
    Set tmpRec = New ADODB.Recordset
    GetDBConnection
    tmpRec.open sqlStr, DBConn, adOpenDynamic, adLockOptimistic
    If tmpRec.EOF Then
        tmpRec.AddNew
        tmpRec("JobDataCarID") = JobDataCarID_
        tmpRec("TripNo") = TripNo_
    End If
    tmpRec!TimeTableNo = TimeTableNo_
    tmpRec("StartTime") = Trim(startTime_)
    tmpRec("VendorTruckID") = Trim(VendorTruckID_)
    tmpRec("DriverName") = Trim(DriverName_)
    tmpRec("MenifestID") = Trim(MenifestID_)
    tmpRec("WorkerCount") = Trim(workerCount_)
    If Trim(WorkDate) <> "" Then workDate_ = WorkDate
    tmpRec("WorkDate") = Trim(workDate_)
    tmpRec("AppointDate") = Trim(appointDate_)
    tmpRec("AppointTime") = Trim(appointTime_)
    tmpRec("CustInDate") = Trim(custInDate_)
    tmpRec("CustInTime") = Trim(custInTime_)
    tmpRec("CustOutDate") = Trim(custOutDate_)
    tmpRec("CustOutTime") = Trim(custOutTime_)
    tmpRec("BackDate") = Trim(backDate_)
    tmpRec("BackTime") = Trim(backTime_)
    If Trim(isTrailer_) <> "Y" Then
        isMain_ = "Y"
    Else
        If Trim(trailTruckID_) = "" Then
            isMain_ = "Y"
        Else
            isMain_ = "Y"
        End If
    End If
    tmpRec("isMain") = isMain_
    tmpRec("MarkForDelete") = "N"
    tmpRec("TrailTruckID") = Trim(trailTruckID_)
    
    tmpRec("TripTranFee") = tripTranFee_
    tmpRec("TripMoreWorker") = tripMoreWorker_
    tmpRec("TripMoreWorkerFee") = tripMoreWorkerFee_
    tmpRec("TripEquipFee") = tripEquipFee_
    tmpRec("TripOtherDesc") = tripOtherDesc_
    tmpRec("TripOtherFee") = tripOtherFee_
    tmpRec("TripAllFee") = tripAllFee_
    tmpRec("TripTreatFee") = tripTreatFee_
    tmpRec("CustBillID") = custBillID_
    
    tmpRec("VDTransFee") = vdTranFee_
    tmpRec("VDWorker") = vdWorker_
    tmpRec("VDWorkerFee") = vdWorkerFee_
    tmpRec("VDEquipFee") = vdEquipFee_
    tmpRec("VDOtherDesc") = vdOtherDesc_
    tmpRec("VDOtherFee") = vdOtherFee_
    
    tmpRec!BoxType = BoxType_
    tmpRec!BoxCount = BoxCount_
    
    tmpRec!TripTransFee_Inc = IncTripTransFee_
    
    If Trim("" & tmpRec!EditTypeTrip) <> "M" Then tmpRec!EditTypeTrip = EditTypeTrip_
    If Trim("" & tmpRec!EditTypeVD) <> "M" Then tmpRec!EditTypeVD = EditTypeVD_
    
    If AutoCal = False Then
        tmpRec!LastUser = Trim(Trim(CurrentUser.StaffFName) & " " & Trim(CurrentUser.StaffLName))
        tmpRec!LastUpdate = GetSystemDateTime.SystemDateTime
    End If
    
    tmpRec.Update
    tmpRec.Requery
    TimeTableNo_ = Trim("" & tmpRec("TimeTableNo"))
    '*** Update Worker
    sqlStr = "UPDATE tbJobDataTimeTableBox SET TripMoreWorker = " & tripMoreWorker_ & ", TripMoreWorkerFee = " & tripMoreWorkerFee_ & vbCrLf
    sqlStr = sqlStr & "WHERE TimeTableNo = '" & TimeTableNo_ & "'"
    DBConn.Execute sqlStr
    
    '*** Update Vendor
    sqlStr = "UPDATE tbJobDataTimeTableBox SET VDTransFee = " & vdTranFee_ & ", VDWorker = " & vdWorker_ & ", VDWorkerFee = " & vdWorkerFee_ & ", VDEquipFee = " & vdEquipFee_ & ", VDOtherDesc = '" & vdOtherDesc_ & "', VDOtherFee = " & vdOtherFee_ & "" & vbCrLf
    sqlStr = sqlStr & "WHERE TimeTableNo = '" & TimeTableNo_ & "'"
    DBConn.Execute sqlStr
'    tmpRec.Close
    Set tmpRec = Nothing
    isChanged_ = False
    CloseDBConnection
    Exit Sub
ErrD:
    GetMsg "", vbCritical
    Exit Sub
End Sub

Public Sub CancelBill()
    DBConnExc "Update tbJobDataTimeTableBox set isCancel='Y',CancelDate='" & TodayDate & "',CancelStaffID='" & CurrentUser.ID & "', BoxType='', BoxCount = 0 WHERE JobDataCarID='" & JobDataCarID_ & "' And TripNo=" & TripNo_
'    DBConnExc "Update tbMenifestHeader isCancel='Y',CancelDate='" & TodayDate & "',CancelStaffID='" & CurrentUser.ID & "' WHERE JobDataCarID='" & JobDataCarID_ & "' And TripNo=" & TripNo_
    isCancel_ = "Y"
    cancelDate_ = TodayDate
    cancelStaffID_ = CurrentUser.ID
End Sub

'Public Function GetMenifests(isAllOrOnlyThisCar As String) As Collection
'Dim tmpS As New BWGSearchManager
'    If isAllOrOnlyThisCar = "A" Then
'        Set GetMenifests = tmpS.MenifestSearch("(TimeTableNo='" & TimeTableNo_ & "') And (isCanceled<>'Y' Or isCanceled is null)")
'    Else
'        Set GetMenifests = tmpS.MenifestSearch("JobDataCarID='" & JobDataCarID_ & "' And TripNo=" & TripNo_ & " And (isCanceled<>'Y' Or isCanceled is null)")
'    End If
'    Set tmpS = Nothing
'End Function

Public Function GetQuotationForThisTrip() As Collection
Dim tmpS As New BWGSearchManager
    Set GetQuotationForThisTrip = tmpS.QuotationInTripSearch("JobDataCarID='" & JobDataCarID_ & "' And TripNo=" & TripNo_)
    Set tmpS = Nothing
End Function

Public Function getDuplicate() As BWGJobDataTimeTableBox
Dim tmpTime As New BWGJobDataTimeTableBox
    With tmpTime
        .JobDataCarID = JobDataCarID_
        .TripNo = TripNo_
        .StartTime = startTime_
        .VendorTruckID = VendorTruckID_
        .VendorID = vendorID_
        .VendorName = vendorName_
        .VendorRegisNo = vendorRegisNo_
        .VendorWorkerCharge = vendorWorkerCharge_
        .CarRegisNo = carRegisNo_
        .DriverName = DriverName_
        .menifestID = MenifestID_
        .WorkerCount = workerCount_
        .WorkDate = workDate_
        .TimeTableNo = TimeTableNo_
        .AppointDate = appointDate_
        .AppointTime = appointTime_
        .CustInDate = custInDate_
        .CustInTime = custInTime_
        .CustOutDate = custOutDate_
        .CustOutTime = custOutTime_
        .backDate = backDate_
        .BackTime = backTime_
        .CancelDate = cancelDate_
        .CancelStaffID = cancelStaffID_
        .CancelStaffName = cancelStaffName_
        .TrailTruckID = trailTruckID_
    
        '.TripTranFee = tripTranFee_
        '.TripMoreWorker = tripMoreWorker_
        '.TripMoreWorkerFee = tripMoreWorkerFee_
        '.TripEquipFee = tripEquipFee_
        '.TripOtherDesc = tripOtherDesc_
        '.TripOtherFee = tripOtherFee_
        '.TripAllFee = tripAllFee_

    End With
    Set getDuplicate = tmpTime
    Set tmpTime = Nothing
End Function

Public Sub UpdateCustBillNo()
    DBConnExc "Update tbJobDataTimeTableBox Set CustBillID='" & custBillID_ & "' WHERE JobDataCarID='" & JobDataCarID_ & "' And TripNo=" & TripNo_
End Sub

Public Sub UpdateBoxCanceled()
    DBConnExc "Update tbJobDataTimeTableBox Set BoxCanceled = BoxCount WHERE JobDataCarID='" & JobDataCarID_ & "' And VendorTruckID = '" & VendorTruckID_ & "' AND TripNo=" & TripNo_
    DBConnExc "Update tbJobDataTimeTableBox Set BoxCount = 0 WHERE JobDataCarID='" & JobDataCarID_ & "' And VendorTruckID = '" & VendorTruckID_ & "' AND TripNo=" & TripNo_
End Sub
