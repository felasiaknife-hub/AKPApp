VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BWGJobDataTimeTable"
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
Dim WorkDate_ As String
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
Dim isJobs_ As String
Dim TrailerType_ As String
Dim GPSSend_ As String
Dim SendDesc_ As String
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
Dim IsDiscountPriceTrans_ As String
''''''''''''''''''''''
Dim vdTranFee_ As Double
Dim vdWorker_ As Integer
Dim vdWorkerFee_ As Double
Dim vdEquipFee_ As Double
Dim vdOtherDesc_ As String
Dim vdOtherFee_ As Double
Dim EditTypeVD_ As String
''''''''''''''''''''''
Dim TruckTypeID_ As String
Dim TruckTypeDesc_ As String
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
Dim Allowance_ As Double
Dim Dream_ As Double

''''''''''''''''''''''
Dim isMain_ As String
Dim BoxType_ As String
Dim BoxCount_ As Integer

Dim TankType_ As String
Dim Tankcount_ As Integer
Dim TankTypeB_ As Integer
Dim TankTypeS_ As String
Dim TankTypeSCount_ As Integer
Dim TankTypeSB_ As Integer
Dim TankPrice_ As Integer
Dim TankPriceS_ As Integer
Dim TankPriceSV_ As Integer
Dim BagType_ As String
Dim BagCount_ As Integer
Dim BagPrice_ As Integer
Dim TankPriceV_ As Integer
Dim BagPriceV_ As Integer

Dim colJobDet_ As Collection
Dim colMenifest_ As Collection

Dim CompanyCode_ As String
Dim isChanged_ As Boolean

Dim AumphurName_ As String
Dim ProvinceName_ As String

Dim StTel_ As String
Dim StName_ As String
Dim OilUse_ As String


Public Property Get OilUse() As String
    OilUse = OilUse_
End Property
Public Property Let OilUse(ByVal newval As String)
    If OilUse_ <> Trim(newval) Then
        OilUse_ = Trim(newval)
        isChanged_ = True
    End If
End Property



Public Property Get AumphurName() As String
    AumphurName = AumphurName_
End Property
Public Property Let AumphurName(ByVal newval As String)
    If AumphurName_ <> Trim(newval) Then
        AumphurName_ = Trim(newval)
        isChanged_ = True
    End If
End Property


Public Property Get ProvinceName() As String
    ProvinceName = ProvinceName_
End Property
Public Property Let ProvinceName(ByVal newval As String)
    If ProvinceName_ <> Trim(newval) Then
        ProvinceName_ = Trim(newval)
        isChanged_ = True
    End If
End Property


Public Property Get CompanyCode() As String
    CompanyCode = CompanyCode_
End Property
Public Property Let CompanyCode(ByVal newval As String)
    If CompanyCode_ <> Trim(newval) Then
        CompanyCode_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get GPSSend() As String
    GPSSend = GPSSend_
End Property
Public Property Let GPSSend(ByVal newval As String)
    If GPSSend_ <> Trim(newval) Then
        GPSSend_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get SendDesc() As String
    SendDesc = SendDesc_
End Property
Public Property Let SendDesc(ByVal newval As String)
    If SendDesc_ <> Trim(newval) Then
        SendDesc_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get JobDataCarID() As String
    JobDataCarID = JobDataCarID_
End Property
Public Property Let JobDataCarID(ByVal newval As String)
    If JobDataCarID_ <> Trim(newval) Then
        JobDataCarID_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get TripNo() As Integer
    TripNo = TripNo_
End Property
Public Property Let TripNo(ByVal newval As Integer)
    If TripNo_ <> newval Then
        TripNo_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get StartTime() As String
    StartTime = startTime_
End Property
Public Property Let StartTime(ByVal newval As String)
    If startTime_ <> Trim(newval) Then
        startTime_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get VendorTruckID() As String
    VendorTruckID = VendorTruckID_
End Property
Public Property Let VendorTruckID(ByVal newval As String)
    If VendorTruckID_ <> Trim(newval) Then
        VendorTruckID_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get VendorID() As String
    VendorID = vendorID_
End Property
Public Property Let VendorID(ByVal newval As String)
    If vendorID_ <> Trim(newval) Then
        vendorID_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get VendorName() As String
    VendorName = vendorName_
End Property
Public Property Let VendorName(ByVal newval As String)
    If vendorName_ <> Trim(newval) Then
        vendorName_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get VendorRegisNo() As String
    VendorRegisNo = vendorRegisNo_
End Property
Public Property Let VendorRegisNo(ByVal newval As String)
    If vendorRegisNo_ <> Trim(newval) Then
        vendorRegisNo_ = Trim(newval)
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

Public Property Get CarRegisNo() As String
    CarRegisNo = carRegisNo_
End Property
Public Property Let CarRegisNo(ByVal newval As String)
    If carRegisNo_ <> Trim(newval) Then
        carRegisNo_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get TrailerTruckNo() As String
    TrailerTruckNo = TrailerTruckNo_
End Property
Public Property Let TrailerTruckNo(ByVal newval As String)
    If TrailerTruckNo_ <> Trim(newval) Then
        TrailerTruckNo_ = Trim(newval)
    End If
End Property

Public Property Get DriverName() As String
    DriverName = DriverName_
End Property
Public Property Let DriverName(ByVal newval As String)
    If DriverName_ <> Trim(newval) Then
        DriverName_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get menifestID() As String
    menifestID = MenifestID_
End Property
Public Property Let menifestID(ByVal newval As String)
    If MenifestID_ <> Trim(newval) Then
        MenifestID_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get WorkerCount() As Integer
    WorkerCount = workerCount_
End Property
Public Property Let WorkerCount(ByVal newval As Integer)
    If workerCount_ <> newval Then
        workerCount_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get WorkDate() As String
    WorkDate = WorkDate_
End Property
Public Property Let WorkDate(ByVal newval As String)
    If WorkDate_ <> Trim(newval) Then
        WorkDate_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get TimeTableNo() As String
    TimeTableNo = TimeTableNo_
End Property
Public Property Let TimeTableNo(ByVal newval As String)
    If TimeTableNo_ <> Trim(newval) Then
        TimeTableNo_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get AppointDate() As String
    AppointDate = appointDate_
End Property
Public Property Let AppointDate(ByVal newval As String)
    If appointDate_ <> Trim(newval) Then
        appointDate_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get AppointTime() As String
    AppointTime = appointTime_
End Property
Public Property Let AppointTime(ByVal newval As String)
    If appointTime_ <> Trim(newval) Then
        appointTime_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get CustInDate() As String
    CustInDate = custInDate_
End Property
Public Property Let CustInDate(ByVal newval As String)
    If custInDate_ <> Trim(newval) Then
        custInDate_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get CustInTime() As String
    CustInTime = custInTime_
End Property
Public Property Let CustInTime(ByVal newval As String)
    If custInTime_ <> Trim(newval) Then
        custInTime_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get CustOutDate() As String
    CustOutDate = custOutDate_
End Property
Public Property Let CustOutDate(ByVal newval As String)
    If custOutDate_ <> Trim(newval) Then
        custOutDate_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get CustOutTime() As String
    CustOutTime = custOutTime_
End Property
Public Property Let CustOutTime(ByVal newval As String)
    If custOutTime_ <> Trim(newval) Then
        custOutTime_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get backDate() As String
    backDate = backDate_
End Property
Public Property Let backDate(ByVal newval As String)
    If backDate_ <> Trim(newval) Then
        backDate_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get BackTime() As String
    BackTime = backTime_
End Property
Public Property Let BackTime(ByVal newval As String)
    If backTime_ <> Trim(newval) Then
        backTime_ = Trim(newval)
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
Public Property Let isCancel(ByVal newval As Boolean)
Dim tmpStr As String
    If newval Then
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
Public Property Let CancelDate(ByVal newval As String)
    If cancelDate_ <> Trim(newval) Then
        cancelDate_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get CancelStaffID() As String
    CancelStaffID = cancelStaffID_
End Property
Public Property Let CancelStaffID(ByVal newval As String)
    If cancelStaffID_ <> Trim(newval) Then
        cancelStaffID_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get CancelStaffName() As String
    CancelStaffName = cancelStaffName_
End Property
Public Property Let CancelStaffName(ByVal newval As String)
    If cancelStaffName_ <> Trim(newval) Then
        cancelStaffName_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get TrailTruckID() As String
    TrailTruckID = trailTruckID_
End Property
Public Property Let TrailTruckID(ByVal newval As String)
    If trailTruckID_ <> Trim(newval) Then
        trailTruckID_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get CustBillID() As String
    CustBillID = custBillID_
End Property
Public Property Let CustBillID(ByVal newval As String)
    If custBillID_ <> Trim(newval) Then
        custBillID_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get BillingNo() As String
    BillingNo = BillingNo_
End Property
Public Property Let BillingNo(ByVal newval As String)
    If BillingNo_ <> Trim(newval) Then
        BillingNo_ = Trim(newval)
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
Public Property Let isTrailer(ByVal newval As Boolean)
Dim tmpStr As String
    If newval Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If isTrailer_ <> Trim(tmpStr) Then
        isTrailer_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property
Public Property Get isJobs() As String
    isJobs = isJobs_
End Property
Public Property Let isJobs(ByVal newval As String)
    If isJobs_ <> newval Then
        isJobs_ = newval
    End If
End Property

Public Property Get TrailerType() As String
    TrailerType = TrailerType_
End Property
Public Property Let TrailerType(ByVal newval As String)
    If TrailerType_ <> newval Then
        TrailerType_ = newval
    End If
End Property

Public Property Get TripTranFee() As Double
    TripTranFee = tripTranFee_
End Property
Public Property Let TripTranFee(ByVal newval As Double)
    If tripTranFee_ <> newval Then
        tripTranFee_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get TripMoreWorker() As Double
    TripMoreWorker = tripMoreWorker_
End Property
Public Property Let TripMoreWorker(ByVal newval As Double)
    If tripMoreWorker_ <> newval Then
        tripMoreWorker_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get TripMoreWorkerFee() As Double
    TripMoreWorkerFee = tripMoreWorkerFee_
End Property
Public Property Let TripMoreWorkerFee(ByVal newval As Double)
    If tripMoreWorkerFee_ <> newval Then
        tripMoreWorkerFee_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get TripEquipFee() As Double
    TripEquipFee = tripEquipFee_
End Property
Public Property Let TripEquipFee(ByVal newval As Double)
    If tripEquipFee_ <> newval Then
        tripEquipFee_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get TripOtherDesc() As String
    TripOtherDesc = Trim(tripOtherDesc_)
End Property
Public Property Let TripOtherDesc(ByVal newval As String)
    If Trim(tripOtherDesc_) <> Trim(newval) Then
        tripOtherDesc_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get TripOtherFee() As Double
    TripOtherFee = tripOtherFee_
End Property
Public Property Let TripOtherFee(ByVal newval As Double)
    If tripOtherFee_ <> newval Then
        tripOtherFee_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get TripAllFee() As Double
    TripAllFee = tripAllFee_
End Property
Public Property Let TripAllFee(ByVal newval As Double)
    If tripAllFee_ <> newval Then
        tripAllFee_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get TripTreatFee() As Double
    TripTreatFee = tripTreatFee_
End Property
Public Property Let TripTreatFee(ByVal newval As Double)
    If tripTreatFee_ <> newval Then
        tripTreatFee_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get IncTripTransFee() As Double
    IncTripTransFee = IncTripTransFee_
End Property
Public Property Let IncTripTransFee(ByVal newval As Double)
    If IncTripTransFee_ <> newval Then
        IncTripTransFee_ = newval
        isChanged_ = True
    End If
End Property
Public Property Get IsDiscountPriceTrans() As Boolean
    If Trim(IsDiscountPriceTrans_) = "Y" Then
        IsDiscountPriceTrans = True
    Else
        IsDiscountPriceTrans = False
    End If
End Property
Public Property Let IsDiscountPriceTrans(ByVal newval As Boolean)
Dim tmpStr As String
    If newval Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(IsDiscountPriceTrans_) <> Trim(tmpStr) Then
        IsDiscountPriceTrans_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property

Public Property Get EditTypeTrip() As String
    EditTypeTrip = EditTypeTrip_
End Property
Public Property Let EditTypeTrip(ByVal newval As String)
Dim tmpStr As String
        If newval = "" Then
            tmpStr = "A"
        Else
            tmpStr = Trim(newval)
        End If
    If Trim(EditTypeTrip_) <> Trim(tmpStr) Then
        EditTypeTrip_ = Trim(tmpStr)
    End If
End Property

Public Property Get VDTranFee() As Double
    VDTranFee = vdTranFee_
End Property
Public Property Let VDTranFee(ByVal newval As Double)
    If vdTranFee_ <> newval Then
        vdTranFee_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get VDWorker() As Double
    VDWorker = vdWorker_
End Property
Public Property Let VDWorker(ByVal newval As Double)
    If vdWorker_ <> newval Then
        vdWorker_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get VDWorkerFee() As Double
    VDWorkerFee = vdWorkerFee_
End Property
Public Property Let VDWorkerFee(ByVal newval As Double)
    If vdWorkerFee_ <> newval Then
        vdWorkerFee_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get VDEquipFee() As Double
    VDEquipFee = vdEquipFee_
End Property
Public Property Let VDEquipFee(ByVal newval As Double)
    If vdEquipFee_ <> newval Then
        vdEquipFee_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get VDOtherDesc() As String
    VDOtherDesc = Trim(vdOtherDesc_)
End Property
Public Property Let VDOtherDesc(ByVal newval As String)
    If Trim(vdOtherDesc_) <> Trim(newval) Then
        vdOtherDesc_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get VDOtherFee() As Double
    VDOtherFee = vdOtherFee_
End Property
Public Property Let VDOtherFee(ByVal newval As Double)
    If vdOtherFee_ <> newval Then
        vdOtherFee_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get EditTypeVD() As String
    EditTypeVD = EditTypeVD_
End Property
Public Property Let EditTypeVD(ByVal newval As String)
Dim tmpStr As String
        If newval = "" Then
            tmpStr = "A"
        Else
            tmpStr = Trim(newval)
        End If
    If Trim(EditTypeVD_) <> Trim(tmpStr) Then
        EditTypeVD_ = Trim(tmpStr)
    End If
End Property

Public Property Get TruckTypeID() As String
    TruckTypeID = Trim(TruckTypeID_)
End Property
Public Property Let TruckTypeID(ByVal newval As String)
    If Trim(TruckTypeID_) <> Trim(newval) Then
        TruckTypeID_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get TruckTypeDesc() As String
    TruckTypeDesc = Trim(TruckTypeDesc_)
End Property
Public Property Let TruckTypeDesc(ByVal newval As String)
    If Trim(TruckTypeDesc_) <> Trim(newval) Then
        TruckTypeDesc_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get TruckSubTypeID() As String
    TruckSubTypeID = Trim(truckSubTypeID_)
End Property
Public Property Let TruckSubTypeID(ByVal newval As String)
    If Trim(truckSubTypeID_) <> Trim(newval) Then
        truckSubTypeID_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get TruckSubTypeDesc() As String
    TruckSubTypeDesc = Trim(truckSubTypeDesc_)
End Property
Public Property Let TruckSubTypeDesc(ByVal newval As String)
    If Trim(truckSubTypeDesc_) <> Trim(newval) Then
        truckSubTypeDesc_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get FromAumphurID() As String
    FromAumphurID = Trim(fromAumphurID_)
End Property
Public Property Let FromAumphurID(ByVal newval As String)
    If Trim(fromAumphurID_) <> Trim(newval) Then
        fromAumphurID_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get FromAumphurName() As String
    FromAumphurName = Trim(fromAumphurName_)
End Property
Public Property Let FromAumphurName(ByVal newval As String)
    If Trim(fromAumphurName_) <> Trim(newval) Then
        fromAumphurName_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get FromProvinceID() As String
    FromProvinceID = Trim(fromProvinceID_)
End Property
Public Property Let FromProvinceID(ByVal newval As String)
    If Trim(fromProvinceID_) <> Trim(newval) Then
        fromProvinceID_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get FromProvinceName() As String
    FromProvinceName = Trim(fromProvinceName_)
End Property
Public Property Let FromProvinceName(ByVal newval As String)
    If Trim(fromProvinceName_) <> Trim(newval) Then
        fromProvinceName_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get ToAumphurID() As String
    ToAumphurID = Trim(toAumphurID_)
End Property
Public Property Let ToAumphurID(ByVal newval As String)
    If Trim(toAumphurID_) <> Trim(newval) Then
        toAumphurID_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get ToAumphurName() As String
    ToAumphurName = Trim(toAumphurName_)
End Property
Public Property Let ToAumphurName(ByVal newval As String)
    If Trim(toAumphurName_) <> Trim(newval) Then
        toAumphurName_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get ToProvinceID() As String
    ToProvinceID = Trim(toProvinceID_)
End Property
Public Property Let ToProvinceID(ByVal newval As String)
    If Trim(toProvinceID_) <> Trim(newval) Then
        toProvinceID_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get ToProvinceName() As String
    ToProvinceName = Trim(toProvinceName_)
End Property
Public Property Let ToProvinceName(ByVal newval As String)
    If Trim(toProvinceName_) <> Trim(newval) Then
        toProvinceName_ = Trim(newval)
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
Public Property Let isMain(ByVal newval As Boolean)
Dim tmpStr As String
    If newval Then
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
Public Property Let BoxType(ByVal newval As String)
    If Trim(BoxType_) <> Trim(newval) Then
        BoxType_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get BoxCount() As Integer
    BoxCount = BoxCount_
End Property
Public Property Let BoxCount(ByVal newval As Integer)
    If BoxCount_ <> newval Then
        BoxCount_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get JobDetails() As Collection
Dim tmpS As New BWGSearchManager
    'If colJobDet_ Is Nothing Then Set colJobDet_ = tmpS.JobDataDetailSearch("JobID in (Select CreatedFromJobID from tbMenifestHeader WHERE JobDataCarID='" & jobDataCarID_ & "' And TripNo=" & tripNo_ & " and (isCanceled<>'Y' Or isCanceled is null))")
    If colJobDet_ Is Nothing Then Set colJobDet_ = tmpS.JobDataDetailSearch("JobDetID in (Select JobDetID from vw_JobDetIDFromTrip WHERE JobDataCarID='" & JobDataCarID_ & "' And TripNo=" & TripNo_ & " and (isCanceled<>'Y' Or isCanceled is null))")
    Set JobDetails = colJobDet_
End Property
Public Property Set JobDetails(ByVal newval As Collection)
    Set colJobDet_ = newval
    isChanged_ = True
End Property

Public Property Get isChanged() As Boolean
    isChanged = isChanged_
End Property
Public Property Let isChanged(ByVal newval As Boolean)
    isChanged_ = newval
End Property

Public Property Get Allowance() As Double
      Allowance = Allowance_
End Property
Public Property Let Allowance(ByVal newval As Double)
    If Allowance_ <> newval Then
        Allowance_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get Dream() As Double
      Dream = Dream_
End Property
Public Property Let Dream(ByVal newval As Double)
      If Dream_ <> newval Then
        Dream_ = newval
        isChanged_ = True
    End If
End Property
Public Property Let TankType(ByVal newval As String)
    If TankType_ <> newval Then
        TankType_ = newval
        isChanged_ = True
    End If
End Property
Public Property Get Tankcount() As Integer
    Tankcount = Tankcount_
End Property
Public Property Let Tankcount(ByVal newval As Integer)
    If Tankcount_ <> newval Then
        Tankcount_ = newval
        isChanged_ = True
    End If
End Property
Public Property Get TankTypeB() As Integer
    TankTypeB = TankTypeB_
End Property
Public Property Let TankTypeB(ByVal newval As Integer)
    If TankTypeB_ <> newval Then
        TankTypeB_ = newval
        isChanged_ = True
    End If
End Property
Public Property Get TankTypeS() As String
    TankTypeS = TankTypeS_
End Property
Public Property Let TankTypeS(ByVal newval As String)
    If TankTypeS_ <> newval Then
        TankTypeS_ = newval
        isChanged_ = True
    End If
End Property
Public Property Get TankTypeSCount() As Integer
    TankTypeSCount = TankTypeSCount_
End Property
Public Property Let TankTypeSCount(ByVal newval As Integer)
    If TankTypeSCount_ <> newval Then
        TankTypeSCount_ = newval
        isChanged_ = True
    End If
End Property
Public Property Get TankTypeSB() As Integer
    TankTypeSB = TankTypeSB_
End Property
Public Property Let TankTypeSB(ByVal newval As Integer)
    If TankTypeSB_ <> newval Then
        TankTypeSB_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get TankPrice() As Integer
    TankPrice = TankPrice_
End Property
Public Property Let TankPrice(ByVal newval As Integer)
    If TankPrice_ <> newval Then
        TankPrice_ = newval
        isChanged_ = True
    End If
End Property
Public Property Get TankPriceS() As Integer
    TankPriceS = TankPriceS_
End Property
Public Property Let TankPriceS(ByVal newval As Integer)
    If TankPriceS_ <> newval Then
        TankPriceS_ = newval
        isChanged_ = True
    End If
End Property
Public Property Get TankPriceSV() As Integer
    TankPriceSV = TankPriceSV_
End Property
Public Property Let TankPriceSV(ByVal newval As Integer)
    If TankPriceSV_ <> newval Then
        TankPriceSV_ = newval
        isChanged_ = True
    End If
End Property
Public Property Get BagType() As String
    BagType = BagType_
End Property
Public Property Let BagType(ByVal newval As String)
    If BagType_ <> newval Then
        BagType_ = newval
        isChanged_ = True
    End If
End Property
Public Property Get BagCount() As Integer
    BagCount = BagCount_
End Property
Public Property Let BagCount(ByVal newval As Integer)
    If BagCount_ <> newval Then
        BagCount_ = newval
        isChanged_ = True
    End If
End Property
Public Property Get BagPrice() As Integer
    BagPrice = BagPrice_
End Property
Public Property Let BagPrice(ByVal newval As Integer)
    If BagPrice_ <> newval Then
        BagPrice_ = newval
        isChanged_ = True
    End If
End Property
Public Property Get TankPriceV() As Integer
    TankPriceV = TankPriceV_
End Property
Public Property Let TankPriceV(ByVal newval As Integer)
    If TankPriceV_ <> newval Then
        TankPriceV_ = newval
        isChanged_ = True
    End If
End Property
Public Property Get BagPriceV() As Integer
    BagPriceV = BagPriceV_
End Property
Public Property Let BagPriceV(ByVal newval As Integer)
    If BagPriceV_ <> newval Then
        BagPriceV_ = newval
        isChanged_ = True
    End If
End Property
Public Property Get StName() As String
    StName = StName_
End Property
Public Property Let StName(ByVal newval As String)
    If StName_ <> newval Then
        StName_ = newval
        isChanged_ = True
    End If
End Property
Public Property Get StTel() As String
    StTel = StTel_
End Property
Public Property Let StTel(ByVal newval As String)
    If StTel_ <> newval Then
        StTel_ = newval
        isChanged_ = True
    End If
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
    WorkDate_ = Trim("" & selRec("WorkDate"))
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
    cancelDate_ = Trim("" & selRec("CancelDate"))
    cancelStaffID_ = Trim("" & selRec("CancelStaffID"))
    cancelStaffName_ = Trim("" & selRec("StaffFName")) & " " & Trim("" & selRec("StaffLName"))
    trailTruckID_ = Trim("" & selRec("TrailTruckID"))
    If isExist(selRec, "CustBillID") Then custBillID_ = Trim("" & selRec("CustBillID"))
    If isExist(selRec, "BillingNo") Then BillingNo_ = Trim("" & selRec("BillingNo"))
    If isExist(selRec, "isMain") Then isMain_ = Trim("" & selRec("isMain"))
    IsDiscountPriceTrans_ = IIf(IsNull(selRec!IsDiscountPriceTrans), "N", selRec!IsDiscountPriceTrans)
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
    If isExist(selRec, "Allowance") Then If IsNull(selRec("Allowance")) = False Then Allowance_ = selRec("Allowance")
    If isExist(selRec, "Dream") Then If IsNull(selRec("Dream")) = False Then Dream_ = selRec("Dream")
    If isExist(selRec, "StName") Then If IsNull(selRec("StName")) = False Then StName_ = selRec("StName")
    If isExist(selRec, "TelNo") Then If IsNull(selRec("TelNo")) = False Then StTel_ = selRec("TelNo")
      
    If isExist(selRec, "isTrailer") Then isTrailer_ = Trim(selRec("isTrailer"))
    If isExist(selRec, "isJobs") Then isJobs_ = Trim(selRec("isJobs"))
    If isExist(selRec, "HeadOrTrail") Then TrailerType_ = Trim("" & selRec("HeadOrTrail"))

    TruckTypeID_ = Trim("" & selRec("TruckTypeID"))
    TruckTypeDesc_ = Trim("" & selRec("TruckTypeDesc"))
    truckSubTypeID_ = Trim("" & selRec("SubTypeID"))
    truckSubTypeDesc_ = Trim("" & selRec("SubTypeDesc"))
    fromAumphurID_ = Trim("" & selRec("FromAumphurID"))
    fromAumphurName_ = Trim("" & selRec("FromAumphurName"))
    fromProvinceID_ = Trim("" & selRec("FromProvinceID"))
    fromProvinceName_ = Trim("" & selRec("FromProvinceName"))
    toAumphurID_ = Trim("" & selRec("ToAumphurID"))
    toAumphurName_ = Trim("" & selRec("ToAumphurName"))
    toProvinceID_ = Trim("" & selRec("ToProvinceID"))
    toProvinceName_ = Trim("" & selRec("ToProvinceName"))
'    GPSSend_ = Trim("" & selRec("GPSSend"))
    AumphurName_ = Trim("" & selRec("AumphurName"))
    ProvinceName_ = Trim("" & selRec("ProvinceName"))
    OilUse_ = Trim("" & selRec("OilUse"))
    If isExist(selRec, "BoxType") Then BoxType_ = Trim("" & selRec!BoxType)
    If isExist(selRec, "BoxCount") Then BoxCount_ = IIf(IsNull(selRec!BoxCount), 0, selRec!BoxCount)
    
    If isExist(selRec, "EditTypeTrip") Then EditTypeTrip_ = Trim("" & IIf(IsNull(selRec!EditTypeTrip), "A", selRec!EditTypeTrip))
    If isExist(selRec, "EditTypeVD") Then EditTypeVD_ = Trim("" & IIf(IsNull(selRec!EditTypeVD), "A", selRec!EditTypeVD))
    
End Sub
Public Sub SetDataFromRecordsetBill(ByVal selRec As ADODB.Recordset, Optional ByVal isTruck As Boolean = False)
    If isTruck = False Then JobDataCarID_ = Trim("" & selRec!JobDataCarID)
    If isTruck = True Then CompanyCode_ = Trim("" & selRec!CompanyCode)
    If IsNull(selRec("TripNo")) = False Then TripNo_ = selRec("TripNo")
    WorkDate_ = Trim("" & selRec("WorkDate"))
    TimeTableNo_ = Trim("" & selRec("TimeTableNo"))
    carRegisNo_ = Trim("" & selRec("CarRegisID"))
    DriverName_ = Trim("" & selRec("DriverName"))
    If isExist(selRec, "VDTransFee") Then If IsNull(selRec("VDTransFee")) = False Then vdTranFee_ = selRec("VDTransFee")
    TruckTypeDesc_ = Trim("" & selRec("TruckTypeDesc"))
    fromAumphurName_ = Trim("" & selRec("FromAumphurName"))
    fromProvinceName_ = Trim("" & selRec("FromProvinceName"))
    toAumphurName_ = Trim("" & selRec("ToAumphurName"))
    toProvinceName_ = Trim("" & selRec("ToProvinceName"))
    If IsNull(selRec("TankType")) = False Then TankType_ = Trim("" & selRec("TankType"))
    If IsNull(selRec("Tankcount")) = False Then Tankcount_ = Trim("" & selRec("Tankcount"))
    If IsNull(selRec("TankTypeB")) = False Then TankTypeB_ = Trim("" & selRec("TankTypeB"))
    If IsNull(selRec("TankTypeS")) = False Then TankTypeS_ = Trim("" & selRec("TankTypeS"))
    If IsNull(selRec("TankTypeSCount")) = False Then TankTypeSCount_ = Trim("" & selRec("TankTypeSCount"))
    If IsNull(selRec("TankTypeSB")) = False Then TankTypeSB_ = Trim("" & selRec("TankTypeSB"))
    If IsNull(selRec("TankPrice")) = False Then TankPrice_ = Trim("" & selRec("TankPrice"))
    If IsNull(selRec("TankPriceS")) = False Then TankPriceS_ = Trim("" & selRec("TankPriceS"))
    If IsNull(selRec("TankPriceSV")) = False Then TankPriceSV_ = Trim("" & selRec("TankPriceSV"))
    If IsNull(selRec("BagType")) = False Then BagType_ = Trim("" & selRec("BagType"))
    If IsNull(selRec("BagCount")) = False Then BagCount_ = Trim("" & selRec("BagCount"))
    If IsNull(selRec("BagPrice")) = False Then BagPrice_ = Trim("" & selRec("BagPrice"))
    If IsNull(selRec("TankPriceV")) = False Then TankPriceV_ = Trim("" & selRec("TankPriceV"))
    If IsNull(selRec("BagPriceV")) = False Then BagPriceV_ = Trim("" & selRec("BagPriceV"))
    AumphurName_ = Trim("" & selRec("AumphurName"))
    ProvinceName_ = Trim("" & selRec("ProvinceName"))
    OilUse_ = Trim("" & selRec("OilUse_"))
'    If isExist(selRec, "Allowance") Then If IsNull(selRec("Allowance")) = False Then Allowance_ = selRec("Allowance")
'    If isExist(selRec, "Dream") Then If IsNull(selRec("Dream")) = False Then Dream_ = selRec("Dream")
'
    
End Sub


Public Sub Save(Optional ByVal WorkDate As String = "", Optional ByVal AutoCal As Boolean = False)
On Error GoTo ErrD
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
    sqlStr = "SELECT *, CONVERT(DATETIME, GETDATE(), 103) AS GETDATE FROM tbJobDataTimeTable WHERE JobDataCarID='" & JobDataCarID_ & "' And TripNo=" & TripNo_
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
    If Trim(WorkDate) <> "" Then WorkDate_ = WorkDate
    tmpRec("WorkDate") = Trim(WorkDate_)
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
    tmpRec("Allowance") = Allowance_
    tmpRec("Dream") = Dream_
    tmpRec!BoxType = BoxType_
    tmpRec!BoxCount = BoxCount_
    tmpRec("TankType") = TankType_
    tmpRec("Tankcount") = Tankcount_
    tmpRec("TankTypeB") = TankTypeB_
    tmpRec("TankTypeS") = TankTypeS_
    tmpRec("TankTypeSCount") = TankTypeSCount_
    tmpRec("TankTypeSB") = TankTypeSB_
    tmpRec("TankPrice") = TankPrice_
    tmpRec("TankPriceS") = TankPriceS_
    tmpRec("TankPriceSV") = TankPriceSV_
    tmpRec("BagType") = BagType_
    tmpRec("BagCount") = BagCount_
    tmpRec("BagPrice") = BagPrice_
    tmpRec("TankPriceV") = TankPriceV_
    tmpRec("BagPriceV") = BagPriceV_
    
    
    tmpRec("IsDiscountPriceTrans") = IsDiscountPriceTrans_
    tmpRec("OilUse") = OilUse_
    tmpRec!TripTransFee_Inc = IncTripTransFee_
'    tmpRec("GPSSend") = GPSSend_
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
'    sqlStr = "UPDATE tbJobDataTimeTable SET TripMoreWorker = " & tripMoreWorker_ & ", TripMoreWorkerFee = " & tripMoreWorkerFee_ & vbCrLf
'    sqlStr = sqlStr & "WHERE TimeTableNo = '" & TimeTableNo_ & "'"
'    DBConn.Execute sqlStr
    
    '*** Update Vendor
    sqlStr = "UPDATE tbJobDataTimeTable SET VDTransFee = " & vdTranFee_ & ", VDWorker = " & vdWorker_ & ", VDWorkerFee = " & vdWorkerFee_ & ", VDEquipFee = " & vdEquipFee_ & ", VDOtherDesc = '" & vdOtherDesc_ & "', VDOtherFee = " & vdOtherFee_ & " ," & vbCrLf
    sqlStr = sqlStr & " TripMoreWorker = " & tripMoreWorker_ & ", TripMoreWorkerFee = " & tripMoreWorkerFee_ & ",  Allowance = " & Allowance_ & ",Dream = " & Dream_ & "  " & vbCrLf
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
    DBConnExc "Update tbJobDataTimeTable set isCancel='Y',CancelDate='" & TodayDate & "',CancelStaffID='" & CurrentUser.id & "', BoxType='', BoxCount = 0 WHERE JobDataCarID='" & JobDataCarID_ & "' And TripNo=" & TripNo_
    DBConnExc "Update tbMenifestHeader isCancel='Y',CancelDate='" & TodayDate & "',CancelStaffID='" & CurrentUser.id & "' WHERE JobDataCarID='" & JobDataCarID_ & "' And TripNo=" & TripNo_
    isCancel_ = "Y"
    cancelDate_ = TodayDate
    cancelStaffID_ = CurrentUser.id
End Sub

Public Function GetMenifests(isAllOrOnlyThisCar As String) As Collection
Dim tmpS As New BWGSearchManager
    If isAllOrOnlyThisCar = "A" Then
        Set GetMenifests = tmpS.MenifestSearch("(TimeTableNo='" & TimeTableNo_ & "') And (isCanceled<>'Y' Or isCanceled is null)")
    Else
        Set GetMenifests = tmpS.MenifestSearch("JobDataCarID='" & JobDataCarID_ & "' And TripNo=" & TripNo_ & " And (isCanceled<>'Y' Or isCanceled is null)")
    End If
    Set tmpS = Nothing
End Function

Public Function GetQuotationForThisTrip() As Collection
Dim tmpS As New BWGSearchManager
    Set GetQuotationForThisTrip = tmpS.QuotationInTripSearch("JobDataCarID='" & JobDataCarID_ & "' And TripNo=" & TripNo_)
    Set tmpS = Nothing
End Function

Public Function getDuplicate() As BWGJobDataTimeTable
Dim tmpTime As New BWGJobDataTimeTable
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
        .WorkDate = WorkDate_
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
    DBConnExc "Update tbJobDataTimeTable Set CustBillID='" & custBillID_ & "' WHERE JobDataCarID='" & JobDataCarID_ & "' And TripNo=" & TripNo_
End Sub

Public Sub UpdateBoxCanceled()
    DBConnExc "Update tbJobDataTimeTable Set BoxCanceled = BoxCount WHERE JobDataCarID='" & JobDataCarID_ & "' And VendorTruckID = '" & VendorTruckID_ & "' AND TripNo=" & TripNo_
    DBConnExc "Update tbJobDataTimeTable Set BoxCount = 0 WHERE JobDataCarID='" & JobDataCarID_ & "' And VendorTruckID = '" & VendorTruckID_ & "' AND TripNo=" & TripNo_
End Sub
