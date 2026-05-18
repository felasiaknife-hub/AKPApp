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

Dim jobDataCarID_ As String
Dim tripNo_ As Integer
Dim startTime_ As String
Dim vendorTruckID_ As String
Dim vendorID_ As String
Dim vendorName_ As String
Dim vendorRegisNo_ As String
Dim vendorWorkerCharge_ As Double
Dim carRegisNo_ As String
Dim driverName_ As String
Dim menifestID_ As String
Dim workerCount_ As Integer
Dim workDate_ As String
Dim timeTableNo_ As String
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
Dim billingNo_ As String
Dim isTrailer_ As String
'''''''''''''''''''''''
Dim tripTranFee_ As Double
Dim tripMoreWorker_ As Integer
Dim tripMoreWorkerFee_ As Double
Dim tripEquipFee_ As Double
Dim tripOtherDesc_ As String
Dim tripOtherFee_ As Double
Dim tripAllFee_ As Double
Dim tripTreatFee_ As Double
''''''''''''''''''''''
Dim vdTranFee_ As Double
Dim vdWorker_ As Integer
Dim vdWorkerFee_ As Double
Dim vdEquipFee_ As Double
Dim vdOtherDesc_ As String
Dim vdOtherFee_ As Double
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

Dim colJobDet_ As Collection
Dim colMenifest_ As Collection
Dim isChanged_ As Boolean

Public Property Get JobDataCarID() As String
    JobDataCarID = jobDataCarID_
End Property
Public Property Let JobDataCarID(ByVal newVal As String)
    If jobDataCarID_ <> Trim(newVal) Then
        jobDataCarID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get tripNo() As Integer
    tripNo = tripNo_
End Property
Public Property Let tripNo(ByVal newVal As Integer)
    If tripNo_ <> newVal Then
        tripNo_ = newVal
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
    VendorTruckID = vendorTruckID_
End Property
Public Property Let VendorTruckID(ByVal newVal As String)
    If vendorTruckID_ <> Trim(newVal) Then
        vendorTruckID_ = Trim(newVal)
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

Public Property Get DriverName() As String
    DriverName = driverName_
End Property
Public Property Let DriverName(ByVal newVal As String)
    If driverName_ <> Trim(newVal) Then
        driverName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get MenifestID() As String
    MenifestID = menifestID_
End Property
Public Property Let MenifestID(ByVal newVal As String)
    If menifestID_ <> Trim(newVal) Then
        menifestID_ = Trim(newVal)
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
    TimeTableNo = timeTableNo_
End Property
Public Property Let TimeTableNo(ByVal newVal As String)
    If timeTableNo_ <> Trim(newVal) Then
        timeTableNo_ = Trim(newVal)
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
    BillingNo = billingNo_
End Property
Public Property Let BillingNo(ByVal newVal As String)
    If billingNo_ <> Trim(newVal) Then
        billingNo_ = Trim(newVal)
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

Public Property Get JobDetails() As Collection
Dim tmpS As New BWGSearchManager
    'If colJobDet_ Is Nothing Then Set colJobDet_ = tmpS.JobDataDetailSearch("JobID in (Select CreatedFromJobID from tbMenifestHeader Where JobDataCarID='" & jobDataCarID_ & "' And TripNo=" & tripNo_ & " and (isCanceled<>'Y' Or isCanceled is null))")
    If colJobDet_ Is Nothing Then Set colJobDet_ = tmpS.JobDataDetailSearch("JobDetID in (Select JobDetID from vw_JobDetIDFromTrip Where JobDataCarID='" & jobDataCarID_ & "' And TripNo=" & tripNo_ & " and (isCanceled<>'Y' Or isCanceled is null))")
    Set JobDetails = colJobDet_
End Property
Public Property Set JobDetails(ByVal newVal As Collection)
    Set colJobDet_ = newVal
    isChanged_ = True
End Property

Public Property Get isChanged() As Boolean
    isChanged = isChanged_
End Property
Public Property Let isChanged(ByVal newVal As Boolean)
    isChanged_ = newVal
End Property

Public Sub SetDataFromRecordset(selRec As ADODB.Recordset)
    jobDataCarID_ = Trim("" & selRec("JobDataCarID"))
    If IsNull(selRec("TripNo")) = False Then tripNo_ = selRec("TripNo")
    startTime_ = Trim("" & selRec("StartTime"))
    
    vendorTruckID_ = Trim("" & selRec("VendorTruckID"))
    vendorID_ = Trim("" & selRec("VendorID"))
    vendorName_ = Trim("" & selRec("NameTH"))
    carRegisNo_ = Trim("" & selRec("CarRegisID"))
    driverName_ = Trim("" & selRec("DriverName"))
    menifestID_ = Trim("" & selRec("MenifestID"))
    vendorRegisNo_ = Trim("" & selRec("TransLicenseNo"))
    If IsNull(selRec("WorkerCharge")) = False Then vendorWorkerCharge_ = Trim("" & selRec("WorkerCharge"))
    
    If IsNull(selRec("WorkerCount")) = False Then workerCount_ = selRec("WorkerCount")
    workDate_ = Trim("" & selRec("WorkDate"))
    timeTableNo_ = Trim("" & selRec("TimeTableNo"))
    appointDate_ = Trim("" & selRec("AppointDate"))
    appointTime_ = Trim("" & selRec("AppointTime"))
    custInDate_ = Trim("" & selRec("CustInDate"))
    custInTime_ = Trim("" & selRec("CustInTime"))
    custOutDate_ = Trim("" & selRec("CustOutDate"))
    custOutTime_ = Trim("" & selRec("CustOutTime"))
    backDate_ = Trim("" & selRec("BackDate"))
    backTime_ = Trim("" & selRec("BackTime"))
    
    isCancel_ = Trim("" & selRec("isCancel"))
    cancelDate_ = Trim("" & selRec("CancelDate"))
    cancelStaffID_ = Trim("" & selRec("CancelStaffID"))
    cancelStaffName_ = Trim("" & selRec("StaffFName")) & " " & Trim("" & selRec("StaffLName"))
    trailTruckID_ = Trim("" & selRec("TrailTruckID"))
    custBillID_ = Trim("" & selRec("CustBillID"))
    billingNo_ = Trim("" & selRec("BillingNo"))
    isMain_ = Trim("" & selRec("isMain"))
    
    If IsNull(selRec("TripTranFee")) = False Then tripTranFee_ = selRec("TripTranFee")
    If IsNull(selRec("TripMoreWorker")) = False Then tripMoreWorker_ = selRec("TripMoreWorker")
    If IsNull(selRec("TripMoreWorkerFee")) = False Then tripMoreWorkerFee_ = selRec("TripMoreWorkerFee")
    If IsNull(selRec("TripEquipFee")) = False Then tripEquipFee_ = selRec("TripEquipFee")
    If IsNull(selRec("TripOtherDesc")) = False Then tripOtherDesc_ = Trim("" & selRec("TripOtherDesc"))
    If IsNull(selRec("TripOtherFee")) = False Then tripOtherFee_ = selRec("TripOtherFee")
    If IsNull(selRec("TripAllFee")) = False Then tripAllFee_ = selRec("TripAllFee")
    If IsNull(selRec("TripTreatFee")) = False Then tripTreatFee_ = selRec("TripTreatFee")
    
    If IsNull(selRec("VDTransFee")) = False Then vdTranFee_ = selRec("VDTransFee")
    If IsNull(selRec("VDWorker")) = False Then vdWorker_ = selRec("VDWorker")
    If IsNull(selRec("VDWorkerFee")) = False Then vdWorkerFee_ = selRec("VDWorkerFee")
    If IsNull(selRec("VDEquipFee")) = False Then vdEquipFee_ = selRec("VDEquipFee")
    If IsNull(selRec("VDOtherDesc")) = False Then vdOtherDesc_ = Trim("" & selRec("VDOtherDesc"))
    If IsNull(selRec("VDOtherFee")) = False Then vdOtherFee_ = selRec("VDOtherFee")
    
    If isExist(selRec, "isTrailer") Then isTrailer_ = Trim(selRec("isTrailer"))
    
    truckTypeID_ = Trim("" & selRec("TruckTypeID"))
    truckTypeDesc_ = Trim("" & selRec("TruckTypeDesc"))
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
End Sub

Public Sub Save(Optional WorkDate As String = "")
On Error GoTo ErrD
Dim tmpRec As ADODB.Recordset
    Set tmpRec = New ADODB.Recordset
    tmpRec.Open "Select * from tbJobDataTimeTable Where JobDataCarID='" & jobDataCarID_ & "' And TripNo=" & tripNo_, DBManager.GetDBConnection, adOpenDynamic, adLockOptimistic
    If tmpRec.EOF Then
        tmpRec.AddNew
        tmpRec("JobDataCarID") = jobDataCarID_
        tmpRec("TripNo") = tripNo_
    End If
    tmpRec("StartTime") = Trim(startTime_)
    tmpRec("VendorTruckID") = Trim(vendorTruckID_)
    tmpRec("DriverName") = Trim(driverName_)
    tmpRec("MenifestID") = Trim(menifestID_)
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
            isMain_ = "N"
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
        
    tmpRec.Update
    tmpRec.Requery
    timeTableNo_ = Trim("" & tmpRec("TimeTableNo"))
    tmpRec.Close
    Set tmpRec = Nothing
    isChanged_ = False
    Exit Sub
ErrD:
    MsgBox Err.Description & vbCrLf & "กรุณาลองใหม่หรือติดต่อผู้พัฒนาระบบ", vbCritical, "Error"
    Exit Sub
End Sub

Public Sub CancelBill()
    With DBManager.GetDBConnection
        .Execute "Update tbJobDataTimeTable set isCancel='Y',CancelDate='" & TodayDate & "',CancelStaffID='" & CurrentUser.ID & "' Where JobDataCarID='" & jobDataCarID_ & "' And TripNo=" & tripNo_
        .Execute "Update tbMenifestHeader isCancel='Y',CancelDate='" & TodayDate & "',CancelStaffID='" & CurrentUser.ID & "' Where JobDataCarID='" & jobDataCarID_ & "' And TripNo=" & tripNo_
    End With
    isCancel_ = "Y"
    cancelDate_ = TodayDate
    cancelStaffID_ = CurrentUser.ID
End Sub

Public Function GetMenifests(isAllOrOnlyThisCar As String) As Collection
Dim tmpS As New BWGSearchManager
    If isAllOrOnlyThisCar = "A" Then
        Set GetMenifests = tmpS.MenifestSearch("(TimeTableNo='" & timeTableNo_ & "') And (isCanceled<>'Y' Or isCanceled is null)")
    Else
        Set GetMenifests = tmpS.MenifestSearch("JobDataCarID='" & jobDataCarID_ & "' And TripNo=" & tripNo_ & " And (isCanceled<>'Y' Or isCanceled is null)")
    End If
    Set tmpS = Nothing
End Function

Public Function GetQuotationForThisTrip() As Collection
Dim tmpS As New BWGSearchManager
    Set GetQuotationForThisTrip = tmpS.QuotationInTripSearch("JobDataCarID='" & jobDataCarID_ & "' And TripNo=" & tripNo_)
    Set tmpS = Nothing
End Function

Public Function getDuplicate() As BWGJobDataTimeTable
Dim tmpTime As New BWGJobDataTimeTable
    With tmpTime
        .JobDataCarID = jobDataCarID_
        .tripNo = tripNo_
        .StartTime = startTime_
        .VendorTruckID = vendorTruckID_
        .VendorID = vendorID_
        .VendorName = vendorName_
        .VendorRegisNo = vendorRegisNo_
        .VendorWorkerCharge = vendorWorkerCharge_
        .CarRegisNo = carRegisNo_
        .DriverName = driverName_
        .MenifestID = menifestID_
        .WorkerCount = workerCount_
        .WorkDate = workDate_
        .TimeTableNo = timeTableNo_
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
    DBManager.GetDBConnection.Execute "Update tbJobDataTimeTable Set CustBillID='" & custBillID_ & "' Where JobDataCarID='" & jobDataCarID_ & "' And TripNo=" & tripNo_
End Sub
