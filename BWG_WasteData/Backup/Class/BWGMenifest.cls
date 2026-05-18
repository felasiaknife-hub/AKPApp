VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BWGMenifest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim id_ As String
Dim createdFromJobID_ As String
Dim companyID_ As String
Dim companyName_ As String
Dim docNo_ As String
Dim menifestNo_ As String
Dim customerID_ As String
Dim customerName_ As String
Dim custSiteID_ As String
Dim custSiteName_ As String
Dim custSiteAumphurID_ As String
Dim custSiteAddr_ As String
Dim wasteGenNo_ As String
Dim telNo_ As String
Dim faxNo_ As String
Dim erContact_ As String
Dim transportName1_ As String
Dim transRegisNo1_ As String
Dim transportName2_ As String
Dim transRegisNo2_ As String
Dim tsdfName1_ As String
Dim tsdfRegisNo1_ As String
Dim tsdfName2_ As String
Dim tsdfRegisNo2_ As String
Dim totalLiquid_ As Double
Dim liquidUnitID_ As String
Dim liquidUnitName_ As String
Dim totalSolid_ As Double
Dim solidUnitID_ As String
Dim solidUnitName_ As String
Dim specialComment_ As String
Dim transTelNo1_ As String
Dim transFaxNo1_ As String
Dim transErContact1_ As String
Dim truckSubTypeID_ As String
Dim truckSubTypeName_ As String
Dim truckTypeID_ As String
Dim truckTypeName_ As String
Dim carRegisID_ As String
Dim fromProvinceID_ As String
Dim fromProvinceName_ As String
Dim toProvinceID_ As String
Dim toProvinceName_ As String
Dim timeSpending_ As Double
Dim dayOrHour_ As String
Dim tsdfAddr_ As String
Dim tsdfTumbolID_ As String
Dim tsdfTumbolName_ As String
Dim tsdfAumphurID_ As String
Dim tsdfAumphurName_ As String
Dim tsdfProvinceName_ As String
Dim tsdfTelNo_ As String
Dim tsdfFaxNo_ As String
Dim tsdfErContact_ As String
Dim treatmentPeriod_ As Double
Dim periodUnit_ As String
Dim actualQty_ As Double
Dim qtyUnitID_ As String
Dim qtyUnitName_ As String
Dim openStaffID_ As String
Dim openStaffName_ As String
Dim openDate_ As String
Dim workDate_ As String
Dim closeStaffID_ As String
Dim closeStaffName_ As String
Dim closeDate_ As String
Dim closeTime_ As String
Dim isClosed_ As String
Dim isCanceled_ As String
Dim cancelDate_ As String
Dim cancelStaffID_ As String
Dim cancelStaffName_ As String
Dim driverName_ As String
Dim isPrinted_ As String
Dim custMenifestNo_ As String
Dim nonOrHz_ As String
Dim transport2ID_ As String
Dim tsdf1ID_ As String
Dim tsdf2ID_ As String
Dim docJobRecord_ As String
Dim docControl_ As String
Dim docSiteWeight_ As String
Dim docCustWeight_ As String
Dim docCustControl_ As String
Dim isDocCompleted_ As String

Dim inCollectFee_ As Double
Dim inTransFee_ As Double
Dim inWorkerFee_ As Double
Dim inEquipFee_ As Double
Dim inOtherDesc_ As String
Dim inOtherFee_ As Double
Dim outTransFee_ As Double
Dim outWorkerCount_ As Integer
Dim outWorkerFee_ As Double
Dim outEquipFee_ As Double
Dim outOtherDesc_ As String
Dim outOtherFee_ As Double
Dim accStaffID_ As String
Dim accStaffName_ As String
Dim accLastUpdate_ As String

Dim signName1_ As String
Dim signDate1_ As String
Dim signName2_ As String
Dim signDate2_ As String
Dim signName3_ As String
Dim signDate3_ As String

Dim timeOutSite_ As String
Dim dateArriveCust_ As String
Dim timeArriveCust_ As String
Dim dateOutCust_ As String
Dim timeOutCust_ As String
Dim dateArriveSite_ As String
Dim timeArriveSite_ As String

Dim jobNo_ As String
Dim chargeWType_ As String
Dim colJobDet_ As Collection
Dim colDiscrep_ As Collection
Dim colWasteAdded_ As Collection

Dim billingID_ As String
Dim billingNo_ As String
Dim billPayID_ As String
Dim billPayNo_ As String

Dim jobDataCarID_ As String
Dim tripNo_ As Integer
Dim timeTableNo_ As String
Dim isMain_ As String
Dim underStaffID_ As String
Dim underStaffName_ As String
Dim startTime_ As String
Dim tripTranFee_ As Double
Dim tripMoreWorker_ As Integer
Dim tripMoreWorkerFee_ As Double
Dim tripEquipFee_ As Double
Dim tripOtherDesc_ As String
Dim tripOtherFee_ As Double
Dim tripAllFee_ As Double

Dim isManifestAddr_ As String
Dim custFullAddr_ As String
Dim custTypeFlag_ As String
Dim parentCustID_ As String
Dim parentCustName_ As String

Dim isChanged_ As Boolean
Dim isSave_ As Boolean

Public Property Get ID() As String
    ID = id_
End Property
Public Property Let ID(ByVal newVal As String)
    id_ = Trim(newVal)
End Property

Public Property Get CreatedFromJobID() As String
    CreatedFromJobID = createdFromJobID_
End Property
Public Property Let CreatedFromJobID(ByVal newVal As String)
    If Trim(createdFromJobID_) <> Trim(newVal) Then
        createdFromJobID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get companyID() As String
    companyID = companyID_
End Property
Public Property Let companyID(ByVal newVal As String)
    If Trim(companyID_) <> Trim(newVal) Then
        companyID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get CompanyName() As String
    CompanyName = companyName_
End Property
Public Property Let CompanyName(ByVal newVal As String)
    If Trim(companyName_) <> Trim(newVal) Then
        companyName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get docNo() As String
    docNo = docNo_
End Property
Public Property Let docNo(ByVal newVal As String)
    If Trim(docNo_) <> Trim(newVal) Then
        docNo_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get JobNo() As String
    JobNo = jobNo_
End Property
Public Property Let JobNo(ByVal newVal As String)
    If Trim(jobNo_) <> Trim(newVal) Then
        jobNo_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get ChargeWType() As String
    ChargeWType = chargeWType_
End Property
Public Property Let ChargeWType(ByVal newVal As String)
    If Trim(chargeWType_) <> Trim(newVal) Then
        chargeWType_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get JobDetails() As Collection
Dim tmpS As New BWGSearchManager
    If colJobDet_ Is Nothing Then Set colJobDet_ = tmpS.MenifestDetailSearch("MenifestID='" & id_ & "'")
    Set JobDetails = colJobDet_
End Property
Public Property Set JobDetails(ByVal newVal As Collection)
    Set colJobDet_ = newVal
    isChanged_ = True
End Property
Public Sub RequeryJobDetail()
    Set colJobDet_ = Nothing
End Sub

Public Property Get DiscrepWastes() As Collection
Dim tmpS As New BWGSearchManager
    If colDiscrep_ Is Nothing Then Set colDiscrep_ = tmpS.DiscrepWasteSearch("MenifestID='" & id_ & "'")
    Set DiscrepWastes = colDiscrep_
End Property
Public Property Set DiscrepWastes(ByVal newVal As Collection)
    Set colDiscrep_ = newVal
    isChanged_ = True
End Property
Public Sub RequeryDiscrepWaste()
    Set colDiscrep_ = Nothing
End Sub

Public Property Get AddedWastes() As Collection
Dim tmpS As New BWGSearchManager
    If colWasteAdded_ Is Nothing Then Set colWasteAdded_ = tmpS.MenifestWasteAddSearch("MenifestID='" & id_ & "'")
    Set AddedWastes = colWasteAdded_
End Property
Public Property Set AddedWastes(ByVal newVal As Collection)
    Set colWasteAdded_ = newVal
    isChanged_ = True
End Property
Public Sub RequeryAddedWaste()
    Set colWasteAdded_ = Nothing
End Sub

Public Property Get BillingID() As String
    BillingID = billingID_
End Property
Public Property Let BillingID(ByVal newVal As String)
    If Trim(billingID_) <> Trim(newVal) Then
        billingID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get BillingNo() As String
    BillingNo = billingNo_
End Property
Public Property Let BillingNo(ByVal newVal As String)
    If Trim(billingNo_) <> Trim(newVal) Then
        billingNo_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get BillPayID() As String
    BillPayID = billPayID_
End Property
Public Property Let BillPayID(ByVal newVal As String)
    If Trim(billPayID_) <> Trim(newVal) Then
        billPayID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get BillPayNo() As String
    BillPayNo = billPayNo_
End Property
Public Property Let BillPayNo(ByVal newVal As String)
    If Trim(billPayNo_) <> Trim(newVal) Then
        billPayNo_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get JobDataCarID() As String
    JobDataCarID = jobDataCarID_
End Property
Public Property Let JobDataCarID(ByVal newVal As String)
    If Trim(jobDataCarID_) <> Trim(newVal) Then
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

Public Property Get TimeTableNo() As String
    TimeTableNo = timeTableNo_
End Property
Public Property Let TimeTableNo(ByVal newVal As String)
    If Trim(timeTableNo_) <> Trim(newVal) Then
        timeTableNo_ = Trim(newVal)
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

Public Property Get UnderStaffID() As String
    UnderStaffID = underStaffID_
End Property
Public Property Let UnderStaffID(ByVal newVal As String)
    If Trim(underStaffID_) <> Trim(newVal) Then
        underStaffID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get UnderStaffName() As String
    UnderStaffName = underStaffName_
End Property
Public Property Let UnderStaffName(ByVal newVal As String)
    If Trim(underStaffName_) <> Trim(newVal) Then
        underStaffName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get StartTime() As String
    StartTime = startTime_
End Property
Public Property Let StartTime(ByVal newVal As String)
    If Trim(startTime_) <> Trim(newVal) Then
        startTime_ = Trim(newVal)
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

Public Property Get isManifestAddr() As Boolean
    If Trim(isManifestAddr_) = "Y" Then
        isManifestAddr = True
    Else
        isManifestAddr = False
    End If
End Property
Public Property Let isManifestAddr(ByVal newVal As Boolean)
Dim tmpStr As String
    If newVal Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(isManifestAddr_) <> Trim(tmpStr) Then
        isManifestAddr_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property

Public Property Get CustFullAddr() As String
    CustFullAddr = custFullAddr_
End Property
Public Property Let CustFullAddr(ByVal newVal As String)
    If Trim(custFullAddr_) <> Trim(newVal) Then
        custFullAddr_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get CustTypeFlag() As String
    CustTypeFlag = custTypeFlag_
End Property
Public Property Let CustTypeFlag(ByVal newVal As String)
    If Trim(custTypeFlag_) <> Trim(newVal) Then
        custTypeFlag_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get ParentCustID() As String
    ParentCustID = parentCustID_
End Property
Public Property Let ParentCustID(ByVal newVal As String)
    If Trim(parentCustID_) <> Trim(newVal) Then
        parentCustID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get ParentCustName() As String
    ParentCustName = parentCustName_
End Property
Public Property Let ParentCustName(ByVal newVal As String)
    If Trim(parentCustName_) <> Trim(newVal) Then
        parentCustName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get MenifestNo() As String
    MenifestNo = menifestNo_
End Property
Public Property Let MenifestNo(ByVal newVal As String)
    If Trim(menifestNo_) <> Trim(newVal) Then
        menifestNo_ = Trim(newVal)
        isChanged_ = True
    End If
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

Public Property Get CustomerName() As String
    CustomerName = customerName_
End Property
Public Property Let CustomerName(ByVal newVal As String)
    If Trim(customerName_) <> Trim(newVal) Then
        customerName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get CustSiteID() As String
    CustSiteID = custSiteID_
End Property
Public Property Let CustSiteID(ByVal newVal As String)
    If Trim(custSiteID_) <> Trim(newVal) Then
        custSiteID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get CustSiteName() As String
    CustSiteName = custSiteName_
End Property
Public Property Let CustSiteName(ByVal newVal As String)
    If Trim(custSiteName_) <> Trim(newVal) Then
        custSiteName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get CustSiteAumphurID() As String
    CustSiteAumphurID = custSiteAumphurID_
End Property
Public Property Let CustSiteAumphurID(ByVal newVal As String)
    If Trim(custSiteAumphurID_) <> Trim(newVal) Then
        custSiteAumphurID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get CustSiteAddr() As String
    CustSiteAddr = custSiteAddr_
End Property
Public Property Let CustSiteAddr(ByVal newVal As String)
    If Trim(custSiteAddr_) <> Trim(newVal) Then
        custSiteAddr_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get WasteGenNo() As String
    WasteGenNo = wasteGenNo_
End Property
Public Property Let WasteGenNo(ByVal newVal As String)
    If Trim(wasteGenNo_) <> Trim(newVal) Then
        wasteGenNo_ = Trim(newVal)
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

Public Property Get ErContact() As String
    ErContact = erContact_
End Property
Public Property Let ErContact(ByVal newVal As String)
    If Trim(erContact_) <> Trim(newVal) Then
        erContact_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get TransportName1() As String
    TransportName1 = transportName1_
End Property
Public Property Let TransportName1(ByVal newVal As String)
    If Trim(transportName1_) <> Trim(newVal) Then
        transportName1_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get TransRegisNo1() As String
    TransRegisNo1 = transRegisNo1_
End Property
Public Property Let TransRegisNo1(ByVal newVal As String)
    If Trim(transRegisNo1_) <> Trim(newVal) Then
        transRegisNo1_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get TransportName2() As String
    TransportName2 = transportName2_
End Property
Public Property Let TransportName2(ByVal newVal As String)
    If Trim(transportName2_) <> Trim(newVal) Then
        transportName2_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get TransRegisNo2() As String
    TransRegisNo2 = transRegisNo2_
End Property
Public Property Let TransRegisNo2(ByVal newVal As String)
    If Trim(transRegisNo2_) <> Trim(newVal) Then
        transRegisNo2_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get TSDFName1() As String
    TSDFName1 = tsdfName1_
End Property
Public Property Let TSDFName1(ByVal newVal As String)
    If Trim(tsdfName1_) <> Trim(newVal) Then
        tsdfName1_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get TSDFRegisNo1() As String
    TSDFRegisNo1 = tsdfRegisNo1_
End Property
Public Property Let TSDFRegisNo1(ByVal newVal As String)
    If Trim(tsdfRegisNo1_) <> Trim(newVal) Then
        tsdfRegisNo1_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get TSDFName2() As String
    TSDFName2 = tsdfName2_
End Property
Public Property Let TSDFName2(ByVal newVal As String)
    If Trim(tsdfName2_) <> Trim(newVal) Then
        tsdfName2_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get TSDFRegisNo2() As String
    TSDFRegisNo2 = tsdfRegisNo2_
End Property
Public Property Let TSDFRegisNo2(ByVal newVal As String)
    If Trim(tsdfRegisNo2_) <> Trim(newVal) Then
        tsdfRegisNo2_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get TotalLiquid() As Double
    TotalLiquid = totalLiquid_
End Property
Public Property Let TotalLiquid(ByVal newVal As Double)
    If totalLiquid_ <> newVal Then
        totalLiquid_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get LiquidUnitID() As String
    LiquidUnitID = liquidUnitID_
End Property
Public Property Let LiquidUnitID(ByVal newVal As String)
    If Trim(liquidUnitID_) <> Trim(newVal) Then
        liquidUnitID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get LiquidUnitName() As String
    LiquidUnitName = liquidUnitName_
End Property
Public Property Let LiquidUnitName(ByVal newVal As String)
    If Trim(liquidUnitName_) <> Trim(newVal) Then
        liquidUnitName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get TotalSolid() As Double
    TotalSolid = totalSolid_
End Property
Public Property Let TotalSolid(ByVal newVal As Double)
    If totalSolid_ <> newVal Then
        totalSolid_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get SolidUnitID() As String
    SolidUnitID = solidUnitID_
End Property
Public Property Let SolidUnitID(ByVal newVal As String)
    If Trim(solidUnitID_) <> Trim(newVal) Then
        solidUnitID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get SolidUnitName() As String
    SolidUnitName = solidUnitName_
End Property
Public Property Let SolidUnitName(ByVal newVal As String)
    If Trim(solidUnitName_) <> Trim(newVal) Then
        solidUnitName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get SpecialComment() As String
    SpecialComment = specialComment_
End Property
Public Property Let SpecialComment(ByVal newVal As String)
    If Trim(specialComment_) <> Trim(newVal) Then
        specialComment_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get TransTelNo1() As String
    TransTelNo1 = transTelNo1_
End Property
Public Property Let TransTelNo1(ByVal newVal As String)
    If Trim(transTelNo1_) <> Trim(newVal) Then
        transTelNo1_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get TransFaxNo1() As String
    TransFaxNo1 = transFaxNo1_
End Property
Public Property Let TransFaxNo1(ByVal newVal As String)
    If Trim(transFaxNo1_) <> Trim(newVal) Then
        transFaxNo1_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get TransErContact1() As String
    TransErContact1 = transErContact1_
End Property
Public Property Let TransErContact1(ByVal newVal As String)
    If Trim(transErContact1_) <> Trim(newVal) Then
        transErContact1_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get TruckSubTypeID() As String
    TruckSubTypeID = truckSubTypeID_
End Property
Public Property Let TruckSubTypeID(ByVal newVal As String)
    If Trim(truckSubTypeID_) <> Trim(newVal) Then
        truckSubTypeID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get TruckSubTypeName() As String
    TruckSubTypeName = truckSubTypeName_
End Property
Public Property Let TruckSubTypeName(ByVal newVal As String)
    If Trim(truckSubTypeName_) <> Trim(newVal) Then
        truckSubTypeName_ = Trim(newVal)
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

Public Property Get TruckTypeName() As String
    TruckTypeName = truckTypeName_
End Property
Public Property Let TruckTypeName(ByVal newVal As String)
    If Trim(truckTypeName_) <> Trim(newVal) Then
        truckTypeName_ = Trim(newVal)
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

Public Property Get FromProvinceID() As String
    FromProvinceID = fromProvinceID_
End Property
Public Property Let FromProvinceID(ByVal newVal As String)
    If Trim(fromProvinceID_) <> Trim(newVal) Then
        fromProvinceID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get FromProvinceName() As String
    FromProvinceName = fromProvinceName_
End Property
Public Property Let FromProvinceName(ByVal newVal As String)
    If Trim(fromProvinceName_) <> Trim(newVal) Then
        fromProvinceName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get ToProvinceID() As String
    ToProvinceID = toProvinceID_
End Property
Public Property Let ToProvinceID(ByVal newVal As String)
    If Trim(toProvinceID_) <> Trim(newVal) Then
        toProvinceID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get ToProvinceName() As String
    ToProvinceName = toProvinceName_
End Property
Public Property Let ToProvinceName(ByVal newVal As String)
    If Trim(toProvinceName_) <> Trim(newVal) Then
        toProvinceName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get TimeSpending() As Double
    TimeSpending = timeSpending_
End Property
Public Property Let TimeSpending(ByVal newVal As Double)
    If timeSpending_ <> newVal Then
        timeSpending_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get DayOrHour() As String
    DayOrHour = dayOrHour_
End Property
Public Property Let DayOrHour(ByVal newVal As String)
    If Trim(dayOrHour_) <> Trim(newVal) Then
        dayOrHour_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get TSDFAddr() As String
    TSDFAddr = tsdfAddr_
End Property
Public Property Let TSDFAddr(ByVal newVal As String)
    If Trim(tsdfAddr_) <> Trim(newVal) Then
        tsdfAddr_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get TSDFTumbolID() As String
    TSDFTumbolID = tsdfTumbolID_
End Property
Public Property Let TSDFTumbolID(ByVal newVal As String)
    If Trim(tsdfTumbolID_) <> Trim(newVal) Then
        tsdfTumbolID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get TSDFTumbolName() As String
    TSDFTumbolName = tsdfTumbolName_
End Property
Public Property Let TSDFTumbolName(ByVal newVal As String)
    If Trim(tsdfTumbolName_) <> Trim(newVal) Then
        tsdfTumbolName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get TSDFAumphurID() As String
    TSDFAumphurID = tsdfAumphurID_
End Property
Public Property Let TSDFAumphurID(ByVal newVal As String)
    If Trim(tsdfAumphurID_) <> Trim(newVal) Then
        tsdfAumphurID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get TSDFAumphurName() As String
    TSDFAumphurName = tsdfAumphurName_
End Property
Public Property Let TSDFAumphurName(ByVal newVal As String)
    If Trim(tsdfAumphurName_) <> Trim(newVal) Then
        tsdfAumphurName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get TSDFProvinceName() As String
    TSDFProvinceName = tsdfProvinceName_
End Property
Public Property Let TSDFProvinceName(ByVal newVal As String)
    If Trim(tsdfProvinceName_) <> Trim(newVal) Then
        tsdfProvinceName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get TSDFTelNo() As String
    TSDFTelNo = tsdfTelNo_
End Property
Public Property Let TSDFTelNo(ByVal newVal As String)
    If Trim(tsdfTelNo_) <> Trim(newVal) Then
        tsdfTelNo_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get TSDFFaxNo() As String
    TSDFFaxNo = tsdfFaxNo_
End Property
Public Property Let TSDFFaxNo(ByVal newVal As String)
    If Trim(tsdfFaxNo_) <> Trim(newVal) Then
        tsdfFaxNo_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get TSDFErContact() As String
    TSDFErContact = tsdfErContact_
End Property
Public Property Let TSDFErContact(ByVal newVal As String)
    If Trim(tsdfErContact_) <> Trim(newVal) Then
        tsdfErContact_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get TreatmentPeriod() As Double
    TreatmentPeriod = treatmentPeriod_
End Property
Public Property Let TreatmentPeriod(ByVal newVal As Double)
    If treatmentPeriod_ <> newVal Then
        treatmentPeriod_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get PeriodUnit() As String
    PeriodUnit = periodUnit_
End Property
Public Property Let PeriodUnit(ByVal newVal As String)
    If Trim(periodUnit_) <> Trim(newVal) Then
        periodUnit_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get ActualQty() As Double
    ActualQty = actualQty_
End Property
Public Property Let ActualQty(ByVal newVal As Double)
    If actualQty_ <> newVal Then
        actualQty_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get QtyUnitID() As String
    QtyUnitID = qtyUnitID_
End Property
Public Property Let QtyUnitID(ByVal newVal As String)
    If Trim(qtyUnitID_) <> Trim(newVal) Then
        qtyUnitID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get QtyUnitName() As String
    QtyUnitName = qtyUnitName_
End Property
Public Property Let QtyUnitName(ByVal newVal As String)
    If Trim(qtyUnitName_) <> Trim(newVal) Then
        qtyUnitName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get OpenStaffID() As String
    OpenStaffID = openStaffID_
End Property
Public Property Let OpenStaffID(ByVal newVal As String)
    If Trim(openStaffID_) <> Trim(newVal) Then
        openStaffID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get OpenStaffName() As String
    OpenStaffName = openStaffName_
End Property
Public Property Let OpenStaffName(ByVal newVal As String)
    If Trim(openStaffName_) <> Trim(newVal) Then
        openStaffName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get OpenDate() As String
    OpenDate = openDate_
End Property
Public Property Let OpenDate(ByVal newVal As String)
    If Trim(openDate_) <> Trim(newVal) Then
        openDate_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get WorkDate() As String
    WorkDate = workDate_
End Property
Public Property Let WorkDate(ByVal newVal As String)
    If Trim(workDate_) <> Trim(newVal) Then
        workDate_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get CloseStaffID() As String
    CloseStaffID = closeStaffID_
End Property
Public Property Let CloseStaffID(ByVal newVal As String)
    If Trim(closeStaffID_) <> Trim(newVal) Then
        closeStaffID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get CloseStaffName() As String
    CloseStaffName = closeStaffName_
End Property
Public Property Let CloseStaffName(ByVal newVal As String)
    If Trim(closeStaffName_) <> Trim(newVal) Then
        closeStaffName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get CloseDate() As String
    CloseDate = closeDate_
End Property
Public Property Let CloseDate(ByVal newVal As String)
    If Trim(closeDate_) <> Trim(newVal) Then
        closeDate_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get CloseTime() As String
    CloseTime = closeTime_
End Property
Public Property Let CloseTime(ByVal newVal As String)
    If Trim(closeTime_) <> Trim(newVal) Then
        closeTime_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get isClosed() As Boolean
    If Trim(isClosed_) = "Y" Then
        isClosed = True
    Else
        isClosed = False
    End If
End Property
Public Property Let isClosed(ByVal newVal As Boolean)
Dim tmpStr As String
    If newVal Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(isClosed_) <> Trim(tmpStr) Then
        isClosed_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property

Public Property Get isCanceled() As Boolean
    If Trim(isCanceled_) = "Y" Then
        isCanceled = True
    Else
        isCanceled = False
    End If
End Property
Public Property Let isCanceled(ByVal newVal As Boolean)
Dim tmpStr As String
    If newVal Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(isCanceled_) <> Trim(tmpStr) Then
        isCanceled_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property

Public Property Get CancelDate() As String
    CancelDate = cancelDate_
End Property
Public Property Let CancelDate(ByVal newVal As String)
    If Trim(cancelDate_) <> Trim(newVal) Then
        cancelDate_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get CancelStaffID() As String
    CancelStaffID = cancelStaffID_
End Property
Public Property Let CancelStaffID(ByVal newVal As String)
    If Trim(cancelStaffID_) <> Trim(newVal) Then
        cancelStaffID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get CancelStaffName() As String
    CancelStaffName = cancelStaffName_
End Property
Public Property Let CancelStaffName(ByVal newVal As String)
    If Trim(cancelStaffName_) <> Trim(newVal) Then
        cancelStaffName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get DriverName() As String
    DriverName = driverName_
End Property
Public Property Let DriverName(ByVal newVal As String)
    If Trim(driverName_) <> Trim(newVal) Then
        driverName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get isPrinted() As Boolean
    If isPrinted_ = "Y" Then
        isPrinted = True
    Else
        isPrinted = False
    End If
End Property
Public Property Let isPrinted(ByVal newVal As Boolean)
Dim tmpStr As String
    If newVal Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(isPrinted_) <> Trim(tmpStr) Then
        isPrinted_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property

Public Property Get CustMenifestNo() As String
    CustMenifestNo = custMenifestNo_
End Property
Public Property Let CustMenifestNo(ByVal newVal As String)
    If Trim(custMenifestNo_) <> Trim(newVal) Then
        custMenifestNo_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get NonOrHz() As String
    NonOrHz = nonOrHz_
End Property
Public Property Let NonOrHz(ByVal newVal As String)
    If Trim(nonOrHz_) <> Trim(newVal) Then
        nonOrHz_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get Transport2ID() As String
    Transport2ID = transport2ID_
End Property
Public Property Let Transport2ID(ByVal newVal As String)
    If Trim(transport2ID_) <> Trim(newVal) Then
        transport2ID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get TSDF1ID() As String
    TSDF1ID = tsdf1ID_
End Property
Public Property Let TSDF1ID(ByVal newVal As String)
    If Trim(tsdf1ID_) <> Trim(newVal) Then
        tsdf1ID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get TSDF2ID() As String
    TSDF2ID = tsdf2ID_
End Property
Public Property Let TSDF2ID(ByVal newVal As String)
    If Trim(tsdf2ID_) <> Trim(newVal) Then
        tsdf2ID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get DocJobRecord() As Boolean
    If Trim(docJobRecord_) = "Y" Then
        DocJobRecord = True
    Else
        DocJobRecord = False
    End If
End Property
Public Property Let DocJobRecord(ByVal newVal As Boolean)
Dim tmpStr As String
    If newVal Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(docJobRecord_) <> Trim(tmpStr) Then
        docJobRecord_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property

Public Property Get DocControl() As Boolean
    If Trim(docControl_) = "Y" Then
        DocControl = True
    Else
        DocControl = False
    End If
End Property
Public Property Let DocControl(ByVal newVal As Boolean)
Dim tmpStr As String
    If newVal Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(docControl_) <> Trim(tmpStr) Then
        docControl_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property

Public Property Get DocSiteWeight() As Boolean
    If Trim(docSiteWeight_) = "Y" Then
        DocSiteWeight = True
    Else
        DocSiteWeight = False
    End If
End Property
Public Property Let DocSiteWeight(ByVal newVal As Boolean)
Dim tmpStr As String
    If newVal Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(docSiteWeight_) <> Trim(tmpStr) Then
        docSiteWeight_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property

Public Property Get DocCustWeight() As Boolean
    If Trim(docCustWeight_) = "Y" Then
        DocCustWeight = True
    Else
        DocCustWeight = False
    End If
End Property
Public Property Let DocCustWeight(ByVal newVal As Boolean)
Dim tmpStr As String
    If newVal Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(docCustWeight_) <> Trim(tmpStr) Then
        docCustWeight_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property

Public Property Get DocCustControl() As Boolean
    If Trim(docCustControl_) = "Y" Then
        DocCustControl = True
    Else
        DocCustControl = False
    End If
End Property
Public Property Let DocCustControl(ByVal newVal As Boolean)
Dim tmpStr As String
    If newVal Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(docCustControl_) <> Trim(tmpStr) Then
        docCustControl_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property

Public Property Get isDocCompleted() As Boolean
    If Trim(isDocCompleted_) = "Y" Then
        isDocCompleted = True
    Else
        isDocCompleted = False
    End If
End Property
Public Property Let isDocCompleted(ByVal newVal As Boolean)
Dim tmpStr As String
    If newVal Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(isDocCompleted_) <> Trim(tmpStr) Then
        isDocCompleted_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property

Public Property Get InCollectFee() As Double
    InCollectFee = inCollectFee_
End Property
Public Property Let InCollectFee(ByVal newVal As Double)
    If inCollectFee_ <> newVal Then
        inCollectFee_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get IntransFee() As Double
    IntransFee = inTransFee_
End Property
Public Property Let IntransFee(ByVal newVal As Double)
    If inTransFee_ <> newVal Then
        inTransFee_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get InWorkerFee() As Double
    InWorkerFee = inWorkerFee_
End Property
Public Property Let InWorkerFee(ByVal newVal As Double)
    If inWorkerFee_ <> newVal Then
        inWorkerFee_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get InEquipFee() As Double
    InEquipFee = inEquipFee_
End Property
Public Property Let InEquipFee(ByVal newVal As Double)
    If inEquipFee_ <> newVal Then
        inEquipFee_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get InOtherDesc() As String
    InOtherDesc = Trim(inOtherDesc_)
End Property
Public Property Let InOtherDesc(ByVal newVal As String)
    If Trim(inOtherDesc_) <> Trim(newVal) Then
        inOtherDesc_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get InOtherFee() As Double
    InOtherFee = inOtherFee_
End Property
Public Property Let InOtherFee(ByVal newVal As Double)
    If inOtherFee_ <> newVal Then
        inOtherFee_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get OutTransFee() As Double
    OutTransFee = outTransFee_
End Property
Public Property Let OutTransFee(ByVal newVal As Double)
    If outTransFee_ <> newVal Then
        outTransFee_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get OutWorkerCount() As Integer
    OutWorkerCount = outWorkerCount_
End Property
Public Property Let OutWorkerCount(ByVal newVal As Integer)
    If outWorkerCount_ <> newVal Then
        outWorkerCount_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get OutWorkerFee() As Double
    OutWorkerFee = outWorkerFee_
End Property
Public Property Let OutWorkerFee(ByVal newVal As Double)
    If outWorkerFee_ <> newVal Then
        outWorkerFee_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get OutEquipFee() As Double
    OutEquipFee = outEquipFee_
End Property
Public Property Let OutEquipFee(ByVal newVal As Double)
    If outEquipFee_ <> newVal Then
        outEquipFee_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get OutOtherDesc() As String
    OutOtherDesc = outOtherDesc_
End Property
Public Property Let OutOtherDesc(ByVal newVal As String)
    If Trim(outOtherDesc_) <> Trim(newVal) Then
        outOtherDesc_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get OutOtherFee() As Double
    OutOtherFee = outOtherFee_
End Property
Public Property Let OutOtherFee(ByVal newVal As Double)
    If outOtherFee_ <> newVal Then
        outOtherFee_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get AccStaffID() As String
    AccStaffID = accStaffID_
End Property
Public Property Let AccStaffID(ByVal newVal As String)
    If Trim(accStaffID_) <> Trim(newVal) Then
        accStaffID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get AccStaffName() As String
    AccStaffName = accStaffName_
End Property
Public Property Let AccStaffName(ByVal newVal As String)
    If Trim(accStaffName_) <> Trim(newVal) Then
        accStaffName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get AccLastUpdate() As String
    AccLastUpdate = accLastUpdate_
End Property
Public Property Let AccLastUpdate(ByVal newVal As String)
    If Trim(accLastUpdate_) <> Trim(newVal) Then
        accLastUpdate_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get SignName1() As String
    SignName1 = signName1_
End Property
Public Property Let SignName1(ByVal newVal As String)
    If Trim(signName1_) <> Trim(newVal) Then
        signName1_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get SignDate1() As String
    SignDate1 = signDate1_
End Property
Public Property Let SignDate1(ByVal newVal As String)
    If Trim(signDate1_) <> Trim(newVal) Then
        signDate1_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get SignName2() As String
    SignName2 = signName2_
End Property
Public Property Let SignName2(ByVal newVal As String)
    If Trim(signName2_) <> Trim(newVal) Then
        signName2_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get SignDate2() As String
    SignDate2 = signDate2_
End Property
Public Property Let SignDate2(ByVal newVal As String)
    If Trim(signDate2_) <> Trim(newVal) Then
        signDate2_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get SignName3() As String
    SignName3 = signName3_
End Property
Public Property Let SignName3(ByVal newVal As String)
    If Trim(signName3_) <> Trim(newVal) Then
        signName3_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get SignDate3() As String
    SignDate3 = signDate3_
End Property
Public Property Let SignDate3(ByVal newVal As String)
    If Trim(signDate3_) <> Trim(newVal) Then
        signDate3_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get TimeOutSite() As String
    TimeOutSite = timeOutSite_
End Property
Public Property Let TimeOutSite(ByVal newVal As String)
    If Trim(timeOutSite_) <> Trim(newVal) Then
        timeOutSite_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get DateArriveCust() As String
    DateArriveCust = dateArriveCust_
End Property
Public Property Let DateArriveCust(ByVal newVal As String)
    If Trim(dateArriveCust_) <> Trim(newVal) Then
        dateArriveCust_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get TimeArriveCust() As String
    TimeArriveCust = timeArriveCust_
End Property
Public Property Let TimeArriveCust(ByVal newVal As String)
    If Trim(timeArriveCust_) <> Trim(newVal) Then
        timeArriveCust_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get DateOutCust() As String
    DateOutCust = dateOutCust_
End Property
Public Property Let DateOutCust(ByVal newVal As String)
    If Trim(dateOutCust_) <> Trim(newVal) Then
        dateOutCust_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get TimeOutCust() As String
    TimeOutCust = timeOutCust_
End Property
Public Property Let TimeOutCust(ByVal newVal As String)
    If Trim(timeOutCust_) <> Trim(newVal) Then
        timeOutCust_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get DateArriveSite() As String
    DateArriveSite = dateArriveSite_
End Property
Public Property Let DateArriveSite(ByVal newVal As String)
    If Trim(dateArriveSite_) <> Trim(newVal) Then
        dateArriveSite_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get TimeArriveSite() As String
    TimeArriveSite = timeArriveSite_
End Property
Public Property Let TimeArriveSite(ByVal newVal As String)
    If Trim(timeArriveSite_) <> Trim(newVal) Then
        timeArriveSite_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get isChanged() As Boolean
    isChanged = isChanged_
End Property
Public Property Let isChanged(ByVal newVal As Boolean)
    isChanged_ = newVal
End Property

Public Property Get isSave() As Boolean
    isSave = isSave_
End Property
Public Property Let isSave(ByVal newVal As Boolean)
    isSave_ = newVal
End Property

Public Sub Save()
On Error GoTo ErrD
Dim tmpRec As ADODB.Recordset
Dim x%, tmpJobDet As BWGJobDataDetail
    isSave_ = True
    Set tmpRec = New ADODB.Recordset
    tmpRec.Open "Select * from tbMenifestHeader Where MenifestID='" & id_ & "'", DBManager.GetDBConnection, adOpenDynamic, adLockOptimistic
    If tmpRec.EOF Then
        tmpRec.AddNew
        tmpRec("MenifestID") = id_
        tmpRec("OpenStaffID") = Trim(openStaffID_)
        tmpRec("OpenDate") = Trim(openDate_)
        tmpRec("CloseStaffID") = Trim(closeStaffID_)
        tmpRec("CloseDate") = Trim(closeDate_)
        tmpRec("CloseTime") = Trim(closeTime_)
        tmpRec("isClosed") = Trim(isClosed_)
        tmpRec("isCanceled") = Trim(isCanceled_)
        tmpRec("CancelDate") = Trim(cancelDate_)
        tmpRec("CancelStaffID") = Trim(cancelStaffID_)
        tmpRec("isPrinted") = Trim(isPrinted_)
    End If
    tmpRec("DocNo") = Trim(docNo_)
    tmpRec("MenifestNo") = Trim(menifestNo_)
    tmpRec("CustomerID") = Trim(customerID_)
    tmpRec("CustSiteID") = Trim(custSiteID_)
    tmpRec("WasteGenNo") = Trim(wasteGenNo_)
    tmpRec("TelNo") = Trim(telNo_)
    tmpRec("FaxNo") = Trim(faxNo_)
    tmpRec("ErContact") = Trim(erContact_)
    tmpRec("TransportName1") = Trim(transportName1_)
    tmpRec("TransRegisNo1") = Trim(transRegisNo1_)
    tmpRec("TransportName2") = Trim(transportName2_)
    tmpRec("TransRegisNo2") = Trim(transRegisNo2_)
    tmpRec("TSDFName1") = Trim(tsdfName1_)
    tmpRec("TSDFRegisNo1") = Trim(tsdfRegisNo1_)
    tmpRec("TSDFName2") = Trim(tsdfName2_)
    tmpRec("TSDFRegisNo2") = Trim(tsdfRegisNo2_)
    tmpRec("TotalLiquid") = Trim(totalLiquid_)
    tmpRec("LiquidUnitID") = Trim(liquidUnitID_)
    tmpRec("TotalSolid") = Trim(totalSolid_)
    tmpRec("SolidUnitID") = Trim(solidUnitID_)
    tmpRec("SpecialComment") = Trim(specialComment_)
    tmpRec("TransTelNo1") = Trim(transTelNo1_)
    tmpRec("TransFaxNo1") = Trim(transFaxNo1_)
    tmpRec("TransErContact1") = Trim(transErContact1_)
    tmpRec("TruckSubTypeID") = Trim(truckSubTypeID_)
    tmpRec("CarRegisID") = Trim(carRegisID_)
    tmpRec("FromProvinceID") = Trim(fromProvinceID_)
    tmpRec("ToProvinceID") = Trim(toProvinceID_)
    tmpRec("TimeSpending") = Trim(timeSpending_)
    tmpRec("DayOrHour") = Trim(dayOrHour_)
    tmpRec("TSDFAddr") = Trim(tsdfAddr_)
    tmpRec("TSDFTelNo") = Trim(tsdfTelNo_)
    tmpRec("TSDFFaxNo") = Trim(tsdfFaxNo_)
    tmpRec("TSDFErContact") = Trim(tsdfErContact_)
    tmpRec("TreatmentPeriod") = Trim(treatmentPeriod_)
    tmpRec("PeriodUnit") = Trim(periodUnit_)
    tmpRec("ActualQty") = Trim(actualQty_)
    tmpRec("QtyUnitID") = Trim(qtyUnitID_)
    tmpRec("CreatedFromJobID") = Trim(createdFromJobID_)
    tmpRec("CompanyID") = Trim(companyID_)
    tmpRec("WorkDate") = Trim(workDate_)
    tmpRec("TSDFTumbolID") = Trim(tsdfTumbolID_)
    tmpRec("DriverName") = Trim(driverName_)
    tmpRec("CustMenifestNo") = Trim(custMenifestNo_)
    tmpRec("NonOrHz") = Trim(nonOrHz_)
    tmpRec("Transport2ID") = Trim(transport2ID_)
    tmpRec("TSDF1ID") = Trim(tsdf1ID_)
    tmpRec("TSDF2ID") = Trim(tsdf2ID_)
    
    If Trim(docJobRecord_) = "" Then docJobRecord_ = "N"
    If Trim(docControl_) = "" Then docControl_ = "N"
    If Trim(docSiteWeight_) = "" Then docSiteWeight_ = "N"
    If Trim(docCustWeight_) = "" Then docCustWeight_ = "N"
    If Trim(docCustControl_) = "" Then docCustControl_ = "N"
    If Trim(isDocCompleted_) = "" Then isDocCompleted_ = "N"
    
    tmpRec("DocJobRecord") = Trim(docJobRecord_)
    tmpRec("DocControl") = Trim(docControl_)
    tmpRec("DocSiteWeight") = Trim(docSiteWeight_)
    tmpRec("DocCustWeight") = Trim(docCustWeight_)
    tmpRec("DocCustControl") = Trim(docCustControl_)
    tmpRec("isDocCompleted") = Trim(isDocCompleted_)
    
    tmpRec("SignName1") = Trim(signName1_)
    tmpRec("SignDate1") = Trim(signDate1_)
    tmpRec("SignName2") = Trim(signName2_)
    tmpRec("SignDate2") = Trim(signDate2_)
    tmpRec("SignName3") = Trim(signName3_)
    tmpRec("SignDate3") = Trim(signDate3_)
    
    tmpRec("InCollectFee") = inCollectFee_
    tmpRec("InTransFee") = inTransFee_
    tmpRec("InWorkerFee") = inWorkerFee_
    tmpRec("InEquipFee") = inEquipFee_
    tmpRec("InOtherDesc") = inOtherDesc_
    tmpRec("InOtherFee") = inOtherFee_
    tmpRec("OutTransFee") = outTransFee_
    tmpRec("OutWorkerCount") = outWorkerCount_
    tmpRec("OutWorkerFee") = outWorkerFee_
    tmpRec("OutEquipFee") = outEquipFee_
    tmpRec("OutOtherDesc") = outOtherDesc_
    tmpRec("OutOtherFee") = outOtherFee_
    tmpRec("AcctStaffID") = accStaffID_
    tmpRec("AccLastUpdate") = accLastUpdate_
    If Trim(chargeWType_) = "" Then chargeWType_ = "A"
    tmpRec("ChargeWType") = chargeWType_
    tmpRec("UnderSaleStaffID") = underStaffID_
    
    tmpRec("JobDataCarID") = jobDataCarID_
    tmpRec("TripNo") = tripNo_
   
    tmpRec.Update
    tmpRec.Requery
    docNo_ = Trim("" & tmpRec("DocNo"))
    menifestNo_ = Trim("" & tmpRec("MenifestNo"))
    tmpRec.Close
    
    DBManager.GetDBConnection.BeginTrans
    With DBManager.GetDBConnection
        If Not colJobDet_ Is Nothing Then
            .Execute "Delete tbMenifestDetail Where MenifestID='" & id_ & "'"
            For x = 1 To colJobDet_.Count
                Set tmpJobDet = colJobDet_(x)
                .Execute "Insert into tbMenifestDetail (MenifestID,JobDetID,ContainerCount,ContainerDesc,CustWeight,DisposerWeight,TreatmentCharge) Values ('" & id_ & "','" & tmpJobDet.ID & "'," & tmpJobDet.ContainerCount & ",'" & tmpJobDet.ContainerTypeDesc & "'," & tmpJobDet.EstWasteQty & "," & tmpJobDet.DisposerWeight & "," & tmpJobDet.TreatmentCharge & ")"
                tmpJobDet.Save
                .Execute "Delete tbJobFail Where MenifestID='" & id_ & "' And JobDetID='" & tmpJobDet.ID & "'"
                If Trim(tmpJobDet.FailType) <> "" Then
                    .Execute "Insert into tbJobFail (JobDetID,MenifestID,FailType,FailRemark) Values ('" & tmpJobDet.ID & "','" & id_ & "','" & tmpJobDet.FailType & "','" & tmpJobDet.FailRemark & "')"
                End If
                Set tmpJobDet = Nothing
            Next
            .Execute "Delete tbMenifestWasteAdded Where MenifestID='" & id_ & "'"
            If Not colWasteAdded_ Is Nothing Then
                For x = 1 To colWasteAdded_.Count
                    colWasteAdded_(x).Save
                Next
            End If
        End If
    End With
    DBManager.GetDBConnection.CommitTrans
    Set tmpRec = Nothing
    isChanged_ = False
    Exit Sub
ErrD:
    isSave_ = False
    DBManager.GetDBConnection.RollbackTrans
    MsgBox Err.Description & vbCrLf & "เกิดข้อผิดพลาด การบันทึกถูกยกเลิก โปรแกรมตรวจสอบข้อมูลที่บันทึกลงไป หลังจากปิดหน้าต่างนี้", vbCritical, "Error"
    Exit Sub
End Sub

Public Function Cancel() As String
On Error GoTo delErr:
    With DBManager.GetDBConnection
        .Execute "Update tbMenifestHeader Set isCanceled='Y',CancelDate='" & TodayDate & "',CancelStaffID='" & CurrentUser.ID & "' Where MenifestID='" & id_ & "'"
        .Execute "Update tbJobdatatimetable set menifestid='' Where MenifestID='" & id_ & "'"
        .Execute "Update tbJobDataHeader set isCompleted='N' Where JobID='" & createdFromJobID_ & "'"
    End With
    Cancel = ""
    isCanceled_ = "Y"
    cancelDate_ = TodayDate
    cancelStaffID_ = CurrentUser.ID
    cancelStaffName_ = CurrentUser.StaffFName & " " & CurrentUser.StaffLName
    Exit Function
delErr:
    Cancel = Err.Description
End Function

Public Sub SetDataFromRecordset(selRec As ADODB.Recordset)
    id_ = Trim("" & selRec("MenifestID"))
    createdFromJobID_ = Trim("" & selRec("CreatedFromJobID"))
    companyID_ = Trim("" & selRec("CompanyID"))
    companyName_ = Trim("" & selRec("CompanyNameTH"))
    docNo_ = Trim("" & selRec("DocNo"))
    menifestNo_ = Trim("" & selRec("MenifestNo"))
    customerID_ = Trim("" & selRec("CustomerID"))
    customerName_ = Trim("" & selRec("CustomerName"))
    custSiteID_ = Trim("" & selRec("CustSiteID"))
    custSiteName_ = Trim("" & selRec("SiteName"))
    If isExist(selRec, "SiteAumphurID") Then custSiteAumphurID_ = Trim("" & selRec("SiteAumphurID"))
    custSiteAddr_ = Replace(Trim(selRec("SiteAddr") & " ต." & selRec("SiteTumbolName") & " อ." & selRec("SiteAumphurName") & " จ." & selRec("SiteProvinceName") & " " & selRec("SiteZipCode")), vbCrLf, "")
    'custSiteAddr_ = Replace(Trim(selRec("SiteAddr") & " อ." & selRec("SiteAumphurName") & " จ." & selRec("SiteProvinceName") & " " & selRec("SiteZipCode")), vbCrLf, "")
    wasteGenNo_ = Trim("" & selRec("WasteGenNo"))
    jobNo_ = Trim("" & selRec("JobNo"))
    chargeWType_ = Trim("" & selRec("ChargeWType"))
    telNo_ = Trim("" & selRec("TelNo"))
    faxNo_ = Trim("" & selRec("FaxNo"))
    erContact_ = Trim("" & selRec("ErContact"))
    transportName1_ = Trim("" & selRec("TransportName1"))
    transRegisNo1_ = Trim("" & selRec("TransRegisNo1"))
    transportName2_ = Trim("" & selRec("TransportName2"))
    transRegisNo2_ = Trim("" & selRec("TransRegisNo2"))
    tsdfName1_ = Trim("" & selRec("TSDFName1"))
    tsdfRegisNo1_ = Trim("" & selRec("TSDFRegisNo1"))
    tsdfName2_ = Trim("" & selRec("TSDFName2"))
    tsdfRegisNo2_ = Trim("" & selRec("TSDFRegisNo2"))
    If IsNull(selRec("TotalLiquid")) = False Then totalLiquid_ = selRec("TotalLiquid")
    liquidUnitID_ = Trim("" & selRec("LiquidUnitID"))
    liquidUnitName_ = Trim("" & selRec("LiquidUnitName"))
    If IsNull(selRec("TotalSolid")) = False Then totalSolid_ = selRec("TotalSolid")
    solidUnitID_ = Trim("" & selRec("SolidUnitID"))
    solidUnitName_ = Trim("" & selRec("SolidUnitName"))
    specialComment_ = Trim("" & selRec("SpecialComment"))
    transTelNo1_ = Trim("" & selRec("TransTelNo1"))
    transFaxNo1_ = Trim("" & selRec("TransFaxNo1"))
    transErContact1_ = Trim("" & selRec("TransErContact1"))
    truckSubTypeID_ = Trim("" & selRec("TruckSubTypeID"))
    truckSubTypeName_ = Trim("" & selRec("SubTypeDesc"))
    truckTypeID_ = Trim("" & selRec("TruckTypeID"))
    truckTypeName_ = Trim("" & selRec("TruckTypeDesc"))
    carRegisID_ = Trim("" & selRec("CarRegisID"))
    fromProvinceID_ = Trim("" & selRec("FromProvinceID"))
    fromProvinceName_ = Trim("" & selRec("FromProvinceName"))
    toProvinceID_ = Trim("" & selRec("ToProvinceID"))
    toProvinceName_ = Trim("" & selRec("ToProvinceName"))
    If IsNull(selRec("TimeSpending")) = False Then timeSpending_ = selRec("TimeSpending")
    dayOrHour_ = Trim("" & selRec("DayOrHour"))
    tsdfAddr_ = Trim("" & selRec("TSDFAddr"))
    tsdfTumbolID_ = Trim("" & selRec("TSDFTumbolID"))
    tsdfTumbolName_ = Trim("" & selRec("TSDFTumbolName"))
    tsdfAumphurID_ = Trim("" & selRec("TSDFAumphurID"))
    tsdfAumphurName_ = Trim("" & selRec("TSDFAumphurName"))
    tsdfProvinceName_ = Trim("" & selRec("TSDFProvinceName"))
    tsdfTelNo_ = Trim("" & selRec("TSDFTelNo"))
    tsdfFaxNo_ = Trim("" & selRec("TSDFFaxNo"))
    tsdfErContact_ = Trim("" & selRec("TSDFErContact"))
    If IsNull(selRec("TreatmentPeriod")) = False Then treatmentPeriod_ = selRec("TreatmentPeriod")
    periodUnit_ = Trim("" & selRec("PeriodUnit"))
    If IsNull(selRec("ActualQty")) = False Then actualQty_ = selRec("ActualQty")
    qtyUnitID_ = Trim("" & selRec("QtyUnitID"))
    qtyUnitName_ = Trim("" & selRec("QtyUnitName"))
    openStaffID_ = Trim("" & selRec("OpenStaffID"))
    openStaffName_ = Trim("" & selRec("OpenStaffName"))
    openDate_ = Trim("" & selRec("OpenDate"))
    workDate_ = Trim("" & selRec("WorkDate"))
    closeStaffID_ = Trim("" & selRec("CloseStaffID"))
    closeStaffName_ = Trim("" & selRec("CloseStaffName"))
    closeDate_ = Trim("" & selRec("CloseDate"))
    closeTime_ = Trim("" & selRec("CloseTime"))
    isClosed_ = Trim("" & selRec("isClosed"))
    isCanceled_ = Trim("" & selRec("isCanceled"))
    cancelDate_ = Trim("" & selRec("CancelDate"))
    cancelStaffID_ = Trim("" & selRec("CancelStaffID"))
    cancelStaffName_ = Trim("" & selRec("CancelStaffName"))
    driverName_ = Trim("" & selRec("DriverName"))
    isPrinted_ = Trim("" & selRec("isPrinted"))
    custMenifestNo_ = Trim("" & selRec("CustMenifestNo"))
    nonOrHz_ = Trim("" & selRec("NonOrHz"))
    transport2ID_ = Trim("" & selRec("Transport2ID"))
    tsdf1ID_ = Trim("" & selRec("TSDF1ID"))
    tsdf2ID_ = Trim("" & selRec("TSDF2ID"))
    docJobRecord_ = Trim("" & selRec("DocJobRecord"))
    docControl_ = Trim("" & selRec("DocControl"))
    docSiteWeight_ = Trim("" & selRec("DocSiteWeight"))
    docCustWeight_ = Trim("" & selRec("DocCustWeight"))
    docCustControl_ = Trim("" & selRec("DocCustControl"))
    isDocCompleted_ = Trim("" & selRec("isDocCompleted"))
    
    signName1_ = Trim("" & selRec("SignName1"))
    signDate1_ = Trim("" & selRec("SignDate1"))
    signName2_ = Trim("" & selRec("SignName2"))
    signDate2_ = Trim("" & selRec("SignDate2"))
    signName3_ = Trim("" & selRec("SignName3"))
    signDate3_ = Trim("" & selRec("SignDate3"))

    If IsNull(selRec("InCollectFee")) = False Then inCollectFee_ = selRec("InCollectFee")
    If IsNull(selRec("InTransFee")) = False Then inTransFee_ = selRec("InTransFee")
    If IsNull(selRec("InWorkerFee")) = False Then inWorkerFee_ = selRec("InWorkerFee")
    If IsNull(selRec("InEquipFee")) = False Then inEquipFee_ = selRec("InEquipFee")
    inOtherDesc_ = Trim("" & selRec("InOtherDesc"))
    If IsNull(selRec("InOtherFee")) = False Then inOtherFee_ = selRec("InOtherFee")
    If IsNull(selRec("OutTransFee")) = False Then outTransFee_ = selRec("OutTransFee")
    If IsNull(selRec("OutWorkerCount")) = False Then outWorkerCount_ = selRec("OutWorkerCount")
    If IsNull(selRec("OutWorkerFee")) = False Then outWorkerFee_ = selRec("OutWorkerFee")
    If IsNull(selRec("OutEquipFee")) = False Then outEquipFee_ = selRec("OutEquipFee")
    outOtherDesc_ = Trim("" & selRec("OutOtherDesc"))
    If IsNull(selRec("OutOtherFee")) = False Then outOtherFee_ = selRec("OutOtherFee")
    accStaffID_ = Trim("" & selRec("AcctStaffID"))
    accStaffName_ = Trim("" & selRec("AccStaffFName") & " " & selRec("AccStaffLName"))
    accLastUpdate_ = Trim("" & selRec("AccLastUpdate"))
    
    timeOutSite_ = Trim("" & selRec("TimeOutSite"))
    dateArriveCust_ = Trim("" & selRec("DateArriveCust"))
    timeArriveCust_ = Trim("" & selRec("TimeArriveCust"))
    dateOutCust_ = Trim("" & selRec("DateOutCust"))
    timeOutCust_ = Trim("" & selRec("TimeOutCust"))
    dateArriveSite_ = Trim("" & selRec("DateArriveSite"))
    timeArriveSite_ = Trim("" & selRec("TimeArriveSite"))
    underStaffID_ = Trim("" & selRec("UnderSaleStaffID"))
    underStaffName_ = Trim("" & selRec("SaleFName")) & " " & Trim("" & selRec("SaleLName"))
    If isExist(selRec, "StartTime") Then startTime_ = Trim("" & selRec("StartTime"))
    
    If isExist(selRec, "TripTranFee") Then
        If IsNull(selRec("TripTranFee")) = False Then tripTranFee_ = selRec("TripTranFee")
    End If
    If isExist(selRec, "TripMoreWorker") Then
        If IsNull(selRec("TripMoreWorker")) = False Then tripMoreWorker_ = selRec("TripMoreWorker")
    End If
    If isExist(selRec, "TripMoreWorkerFee") Then
        If IsNull(selRec("TripMoreWorkerFee")) = False Then tripMoreWorkerFee_ = selRec("TripMoreWorkerFee")
    End If
    If isExist(selRec, "TripEquipFee") Then
        If IsNull(selRec("TripEquipFee")) = False Then tripEquipFee_ = selRec("TripEquipFee")
    End If
    If isExist(selRec, "TripOtherDesc") Then
        If IsNull(selRec("TripOtherDesc")) = False Then tripOtherDesc_ = selRec("TripOtherDesc")
    End If
    If isExist(selRec, "TripOtherFee") Then
        If IsNull(selRec("TripOtherFee")) = False Then tripOtherFee_ = selRec("TripOtherFee")
    End If
    If isExist(selRec, "TripAllFee") Then
        If IsNull(selRec("TripAllFee")) = False Then tripAllFee_ = selRec("TripAllFee")
    End If

    billingID_ = Trim("" & selRec("BillingHeaderID"))
    billingNo_ = Trim("" & selRec("BillingNo"))
    billPayID_ = Trim("" & selRec("BillPayHeaderID"))
    billPayNo_ = Trim("" & selRec("BillPayNo"))
    
    jobDataCarID_ = Trim("" & selRec("JobDataCarID"))
    If IsNull(selRec("TripNo")) = False Then tripNo_ = Trim("" & selRec("TripNo"))
    timeTableNo_ = Trim("" & selRec("TimeTableNo"))
    isMain_ = Trim("" & selRec("isMain"))
    
    isManifestAddr_ = Trim("" & selRec("isManifestAddr"))
    'custFullAddr_ = Replace(Trim(selRec("CustAddr1") & " อ." & selRec("CustAumphurName") & " จ." & selRec("CustProvinceName") & " " & selRec("CustZipCode")), vbCrLf, "")
    custFullAddr_ = Replace(Trim(selRec("CustAddr1") & " ต." & selRec("CustTumbolName") & " อ." & selRec("CustAumphurName") & " จ." & selRec("CustProvinceName") & " " & selRec("CustZipCode")), vbCrLf, "")
    custTypeFlag_ = Trim("" & selRec("CustTypeFlag"))
    parentCustID_ = Trim("" & selRec("ParentCustID"))
    parentCustName_ = Trim("" & selRec("ParentCustName"))
    'CustTumbolName
End Sub

Public Sub CloseMenifest()
Dim amt1 As Double, amt2 As Double, selTripControl As BWGJobDataTimeTable
Dim maxWorker%, maxTransFee As Double, tmpCol As Collection, tmpS As New BWGSearchManager
Dim x%, tmpChargeRate As BWGTransVendorChargeRate
'    timeOutSite_ = Trim("" & tmpRec("TimeOutSite"))
'    dateArriveCust_ = Trim("" & tmpRec("DateArriveCust"))
'    timeArriveCust_ = Trim("" & tmpRec("TimeArriveCust"))
'    dateOutCust_ = Trim("" & tmpRec("DateOutCust"))
'    timeOutCust_ = Trim("" & tmpRec("TimeOutCust"))
'    dateArriveSite_ = Trim("" & tmpRec("DateArriveSite"))
'    TimeArriveSite = Trim("" & tmpRec("TimeArriveSite"))
    
    DBManager.GetDBConnection.Execute "Update tbMenifestHeader Set CloseStaffID='" & CurrentUser.ID & "', CloseDate='" & TodayDate & "',CloseTime=left(convert(varchar(10),getdate(),108),5),isClosed='Y',TimeOutSite='" & timeOutSite_ & "',DateArriveCust='" & dateArriveCust_ & "',TimeArriveCust='" & timeArriveCust_ & "',DateOutCust='" & dateOutCust_ & "',TimeOutCust='" & timeOutCust_ & "',DateArriveSite='" & dateArriveSite_ & "',TimeArriveSite='" & timeArriveSite_ & "' Where MenifestID='" & id_ & "'"
    closeStaffID_ = CurrentUser.ID
    closeStaffName_ = CurrentUser.StaffFName & " " & CurrentUser.StaffLName
    closeDate_ = TodayDate
    closeTime_ = Format(Now, "hh:mm")
    isClosed_ = "Y"
    
    Set tmpCol = tmpS.JobDataTimeTableQuickSearch("TimeTableNo='" & timeTableNo_ & "' And isMain='Y' And isCancel ='N'")
    If tmpCol.Count > 0 Then
        amt1 = CalCustTreatFee(tmpCol(1))
        CalCustTripFee tmpCol(1), maxWorker, maxTransFee
        Set tmpChargeRate = CalTransFeeFromDistance(custSiteAumphurID_, tsdfAumphurID_, truckTypeID_)
        If Not tmpChargeRate Is Nothing Then
            If tmpCol(1).isTrailer Then
                amt2 = tmpChargeRate.TrailerRateTrip
            Else
                amt2 = tmpChargeRate.RatePerTrip
            End If
        End If
        DBManager.GetDBConnection.Execute "Update tbJobDataTimeTable Set TripTranFee=" & maxTransFee & ",TripTreatFee=" & amt1 & ",VDTransFee=" & amt2 & " Where TimeTableNo='" & timeTableNo_ & "' And isCancel='N'"
    End If
    Set tmpChargeRate = Nothing
    Set tmpCol = Nothing
    Set tmpS = Nothing
    isChanged_ = False
End Sub

Public Function getTripControl() As BWGJobDataTimeTable
Dim tmpCol As Collection, tmpS As New BWGSearchManager
    Set tmpCol = tmpS.JobDataTimeTableQuickSearch("JobDataCarID='" & jobDataCarID_ & "' And TripNo=" & tripNo_)
    If tmpCol.Count > 0 Then
        Set getTripControl = tmpCol(1)
    End If
End Function
