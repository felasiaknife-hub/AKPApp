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
Dim CompanyName_ As String
Dim DocNo_ As String
Dim DocNoRef_ As String
Dim MenifestNo_ As String
Dim CustomerID_ As String
Dim CustomerCode_ As String
Dim CustomerName_ As String
Dim CommentTimeNo_ As String
Dim custSiteID_ As String
Dim custSiteName_ As String
Dim custSiteAumphurID_ As String
Dim custSiteAddr_ As String
Dim WasteGenNo_ As String
Dim telNo_ As String
Dim faxNo_ As String
Dim erContact_ As String
Dim transportName1_ As String
Dim TransRegisNo1_ As String
Dim TransportName2_ As String
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
Dim TruckTypeID_ As String
Dim truckTypeName_ As String
Dim CarRegisID_ As String
Dim TrailerCarRegisID_ As String
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
Dim WorkDate_ As String
Dim closeStaffID_ As String
Dim closeStaffName_ As String
Dim closeDate_ As String
Dim closeTime_ As String
Dim isClosed_ As String

Dim approveStaffID_ As String
Dim approveStaffName_ As String
Dim approveDate_ As String
Dim ApproveTime_ As String
Dim isApproved_ As String

Dim IsCanceled_ As String
Dim cancelDate_ As String
Dim CancelDetail_ As String
Dim cancelStaffID_ As String
Dim cancelStaffName_ As String
Dim DriverName_ As String
Dim DriverTelNo_ As String
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
Dim IsPost_ As String
Dim PostDate_ As String
Dim EMSNo_ As String
Dim RefME_ As String

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

Dim SignName1_ As String
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
Dim ChargeWType_ As String
Dim colJobDet_ As Collection
Dim colDiscrep_ As Collection
Dim colWasteAdded_ As Collection
Dim tmpDataCar As Collection
Dim colTripControl As Collection
Dim colMenifestHeader As Collection

Dim billingID_ As String
Dim BillingNo_ As String
Dim billPayID_ As String
Dim billPayNo_ As String

Dim JobDataCarID_ As String
Dim TripNo_ As Integer
Dim TimeTableNo_ As String
Dim isMain_ As String
Dim GPSSend_ As String
Dim SendDesc_ As String
Dim CarRate_ As String

Dim saleDistrictID_ As String
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
Dim TripTransFeeInc_ As Double
Dim isPriceIncTrans_ As String
Dim IsDiscountPriceTrans_  As String


Dim isManifestAddr_ As String
Dim custFullAddr_ As String
Dim CustTypeFlag_ As String
Dim parentCustID_ As String
Dim parentCustName_ As String

Dim isChanged_ As Boolean
Dim isSave_ As Boolean

Dim StaffTelNo_ As String
Dim CarRem_ As String
Dim isJob_ As String
Dim CommentList_ As String
Dim TankType_ As String
Dim Tankcount_ As String
Dim TankcountS_ As String
Dim BagType_ As String
Dim BagCount_ As String
Dim TankTypeSCount_ As String
Dim BagBack_ As String
Dim Container_ As String
Dim Lack_ As String
Dim PaperBox_ As String
Dim BoutS_ As String
Dim LackNew_ As String
Dim CovidBin_ As String
Dim LackGas_ As String
Dim Plate_ As String
Dim Plastic_ As String
Dim BoutCage_ As String

Dim InvState_ As String
Dim REState_ As String
Dim InvDetail_ As String
Dim RemID_ As String
Dim RemComment_ As String
Dim DocCustDate_ As String
Dim PONo_  As String
'** New Variable for AP
Dim APManifestNo_ As String

Dim PhotoOnWeight_ As String
Dim PhotoOnBelt_ As String
Dim PhotoOnBoxfeed_ As String
Dim CustPhoto_ As String
Dim CustCon_ As String
Dim CloseWeight_ As String
Dim EmpTrain_ As String
Dim CarCheck_ As String
Dim takePhoto_ As String
Dim SaleDistricIDOrder_ As String
'***
Dim InWeight_ As Double
Dim Cost_ As Double
Dim LastUser_ As String
Dim LastUpdate_ As String
Dim NewCustJob_ As String

Dim AumphurName_ As String
Dim ProvinceName_ As String
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





Public Property Get Cost() As Double
    Cost = Cost_
End Property
Public Property Let Cost(ByVal newval As Double)
    If Cost_ <> newval Then
        Cost_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get InWeight() As Double
    InWeight = InWeight_
End Property
Public Property Let InWeight(ByVal newval As Double)
    If InWeight_ <> newval Then
        InWeight_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get id() As String
    id = id_
End Property
Public Property Let id(ByVal newval As String)
    id_ = Trim(newval)
End Property
Public Property Get GPSSend() As String
    GPSSend = GPSSend_
End Property
Public Property Let GPSSend(ByVal newval As String)
    If Trim(GPSSend_) <> Trim(newval) Then
        GPSSend_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get SendDesc() As String
    SendDesc = SendDesc_
End Property
Public Property Let SendDesc(ByVal newval As String)
    If Trim(SendDesc_) <> Trim(newval) Then
        SendDesc_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get CarRate() As String
    CarRate = CarRate_
End Property
Public Property Let CarRate(ByVal newval As String)
    If Trim(CarRate_) <> Trim(newval) Then
        CarRate_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get CreatedFromJobID() As String
    CreatedFromJobID = createdFromJobID_
End Property
Public Property Let CreatedFromJobID(ByVal newval As String)
    If Trim(createdFromJobID_) <> Trim(newval) Then
        createdFromJobID_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get CompanyID() As String
    CompanyID = companyID_
End Property
Public Property Let CompanyID(ByVal newval As String)
    If Trim(companyID_) <> Trim(newval) Then
        companyID_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get CompanyName() As String
    CompanyName = CompanyName_
End Property
Public Property Let CompanyName(ByVal newval As String)
    If Trim(CompanyName_) <> Trim(newval) Then
        CompanyName_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get DocNoRef() As String
    DocNoRef = DocNoRef_
End Property
Public Property Let DocNoRef(ByVal newval As String)
    If Trim(DocNoRef_) <> Trim(newval) Then
        DocNoRef_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get DocNo() As String
    DocNo = DocNo_
End Property
Public Property Let DocNo(ByVal newval As String)
    If Trim(DocNo_) <> Trim(newval) Then
        DocNo_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get JobNo() As String
    JobNo = jobNo_
End Property
Public Property Let JobNo(ByVal newval As String)
    If Trim(jobNo_) <> Trim(newval) Then
        jobNo_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get ChargeWType() As String
    ChargeWType = ChargeWType_
End Property
Public Property Let ChargeWType(ByVal newval As String)
    If Trim(ChargeWType_) <> Trim(newval) Then
        ChargeWType_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get JobDetails() As Collection
Dim tmpS As New BWGSearchManager
    If colJobDet_ Is Nothing Then Set colJobDet_ = tmpS.MenifestDetailSearch("MenifestID='" & id_ & "'")
    Set JobDetails = colJobDet_
End Property
Public Property Set JobDetails(ByVal newval As Collection)
    Set colJobDet_ = newval
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
Public Property Set DiscrepWastes(ByVal newval As Collection)
    Set colDiscrep_ = newval
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
Public Property Set AddedWastes(ByVal newval As Collection)
    Set colWasteAdded_ = newval
    isChanged_ = True
End Property
Public Sub RequeryAddedWaste()
    Set colWasteAdded_ = Nothing
End Sub

Public Property Get BillingID() As String
    BillingID = billingID_
End Property
Public Property Let BillingID(ByVal newval As String)
    If Trim(billingID_) <> Trim(newval) Then
        billingID_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get BillingNo() As String
    BillingNo = BillingNo_
End Property
Public Property Let BillingNo(ByVal newval As String)
    If Trim(BillingNo_) <> Trim(newval) Then
        BillingNo_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get BillPayID() As String
    BillPayID = billPayID_
End Property
Public Property Let BillPayID(ByVal newval As String)
    If Trim(billPayID_) <> Trim(newval) Then
        billPayID_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get BillPayNo() As String
    BillPayNo = billPayNo_
End Property
Public Property Let BillPayNo(ByVal newval As String)
    If Trim(billPayNo_) <> Trim(newval) Then
        billPayNo_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get JobDataCarID() As String
    JobDataCarID = JobDataCarID_
End Property
Public Property Let JobDataCarID(ByVal newval As String)
    If Trim(JobDataCarID_) <> Trim(newval) Then
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

Public Property Get TimeTableNo() As String
    TimeTableNo = TimeTableNo_
End Property
Public Property Let TimeTableNo(ByVal newval As String)
    If Trim(TimeTableNo_) <> Trim(newval) Then
        TimeTableNo_ = Trim(newval)
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

Public Property Get SaleDistrictID() As String
    SaleDistrictID = saleDistrictID_
End Property
Public Property Let SaleDistrictID(ByVal newval As String)
    If Trim(saleDistrictID_) <> Trim(newval) Then
        saleDistrictID_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get UnderStaffID() As String
    UnderStaffID = underStaffID_
End Property
Public Property Let UnderStaffID(ByVal newval As String)
    If Trim(underStaffID_) <> Trim(newval) Then
        underStaffID_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get UnderStaffName() As String
    UnderStaffName = underStaffName_
End Property
Public Property Let UnderStaffName(ByVal newval As String)
    If Trim(underStaffName_) <> Trim(newval) Then
        underStaffName_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get StartTime() As String
    StartTime = startTime_
End Property
Public Property Let StartTime(ByVal newval As String)
    If Trim(startTime_) <> Trim(newval) Then
        startTime_ = Trim(newval)
        isChanged_ = True
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

Public Property Get TripTransFeeInc() As Double
    TripTransFeeInc = TripTransFeeInc_
End Property
Public Property Let TripTransFeeInc(ByVal newval As Double)
    If TripTransFeeInc_ <> newval Then
        TripTransFeeInc_ = newval
    End If
End Property


Public Property Get IsDiscountPriceTrans() As Boolean
    If IsDiscountPriceTrans_ = "Y" Then
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
    If IsDiscountPriceTrans_ <> Trim(tmpStr) Then
        IsDiscountPriceTrans_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property


Public Property Get isPriceIncTrans() As Boolean
    If isPriceIncTrans_ = "Y" Then
        isPriceIncTrans = True
    Else
        isPriceIncTrans = False
    End If
End Property
Public Property Let isPriceIncTrans(ByVal newval As Boolean)
Dim tmpStr As String
    If newval Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If isPriceIncTrans_ <> Trim(tmpStr) Then
        isPriceIncTrans_ = Trim(tmpStr)
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
Public Property Let isManifestAddr(ByVal newval As Boolean)
Dim tmpStr As String
    If newval Then
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
Public Property Let CustFullAddr(ByVal newval As String)
    If Trim(custFullAddr_) <> Trim(newval) Then
        custFullAddr_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get CustTypeFlag() As String
    CustTypeFlag = CustTypeFlag_
End Property
Public Property Let CustTypeFlag(ByVal newval As String)
    If Trim(CustTypeFlag_) <> Trim(newval) Then
        CustTypeFlag_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get ParentCustID() As String
    ParentCustID = parentCustID_
End Property
Public Property Let ParentCustID(ByVal newval As String)
    If Trim(parentCustID_) <> Trim(newval) Then
        parentCustID_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get ParentCustName() As String
    ParentCustName = parentCustName_
End Property
Public Property Let ParentCustName(ByVal newval As String)
    If Trim(parentCustName_) <> Trim(newval) Then
        parentCustName_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get MenifestNo() As String
    MenifestNo = MenifestNo_
End Property
Public Property Let MenifestNo(ByVal newval As String)
    If Trim(MenifestNo_) <> Trim(newval) Then
        MenifestNo_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get CustomerID() As String
    CustomerID = CustomerID_
End Property
Public Property Let CustomerID(ByVal newval As String)
    If Trim(CustomerID_) <> Trim(newval) Then
        CustomerID_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get CustomerCode() As String
    CustomerName = CustomerName_
End Property
Public Property Let CustomerCode(ByVal newval As String)
If Trim(CustomerCode_) <> Trim(newval) Then
    CustomerCode_ = Trim(newval)
    isChanged_ = True
End If
End Property
Public Property Get CustomerName() As String
    CustomerName = CustomerName_
End Property
Public Property Let CustomerName(ByVal newval As String)
    If Trim(CustomerName_) <> Trim(newval) Then
        CustomerName_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get CommentTimeNo() As String
    CommentTimeNo = CommentTimeNo_
End Property
Public Property Let CommentTimeNo(ByVal newval As String)
    If Trim(CommentTimeNo_) <> Trim(newval) Then
        CommentTimeNo_ = Trim(newval)
        isChanged_ = True
    End If
End Property


Public Property Get CustSiteID() As String
    CustSiteID = custSiteID_
End Property
Public Property Let CustSiteID(ByVal newval As String)
    If Trim(custSiteID_) <> Trim(newval) Then
        custSiteID_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get CustSiteName() As String
    CustSiteName = custSiteName_
End Property
Public Property Let CustSiteName(ByVal newval As String)
    If Trim(custSiteName_) <> Trim(newval) Then
        custSiteName_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get CustSiteAumphurID() As String
    CustSiteAumphurID = custSiteAumphurID_
End Property
Public Property Let CustSiteAumphurID(ByVal newval As String)
    If Trim(custSiteAumphurID_) <> Trim(newval) Then
        custSiteAumphurID_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get CustSiteAddr() As String
    CustSiteAddr = custSiteAddr_
End Property
Public Property Let CustSiteAddr(ByVal newval As String)
    If Trim(custSiteAddr_) <> Trim(newval) Then
        custSiteAddr_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get WasteGenNo() As String
    WasteGenNo = WasteGenNo_
End Property
Public Property Let WasteGenNo(ByVal newval As String)
    If Trim(WasteGenNo_) <> Trim(newval) Then
        WasteGenNo_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get TelNo() As String
    TelNo = telNo_
End Property
Public Property Let TelNo(ByVal newval As String)
    If Trim(telNo_) <> Trim(newval) Then
        telNo_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get FaxNo() As String
    FaxNo = faxNo_
End Property
Public Property Let FaxNo(ByVal newval As String)
    If Trim(faxNo_) <> Trim(newval) Then
        faxNo_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get ErContact() As String
    ErContact = erContact_
End Property
Public Property Let ErContact(ByVal newval As String)
    If Trim(erContact_) <> Trim(newval) Then
        erContact_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get TransportName1() As String
    TransportName1 = transportName1_
End Property
Public Property Let TransportName1(ByVal newval As String)
    If Trim(transportName1_) <> Trim(newval) Then
        transportName1_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get TransRegisNo1() As String
    TransRegisNo1 = TransRegisNo1_
End Property
Public Property Let TransRegisNo1(ByVal newval As String)
    If Trim(TransRegisNo1_) <> Trim(newval) Then
        TransRegisNo1_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get TransportName2() As String
    TransportName2 = TransportName2_
End Property
Public Property Let TransportName2(ByVal newval As String)
    If Trim(TransportName2_) <> Trim(newval) Then
        TransportName2_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get TransRegisNo2() As String
    TransRegisNo2 = transRegisNo2_
End Property
Public Property Let TransRegisNo2(ByVal newval As String)
    If Trim(transRegisNo2_) <> Trim(newval) Then
        transRegisNo2_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get TSDFName1() As String
    TSDFName1 = tsdfName1_
End Property
Public Property Let TSDFName1(ByVal newval As String)
    If Trim(tsdfName1_) <> Trim(newval) Then
        tsdfName1_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get TSDFRegisNo1() As String
    TSDFRegisNo1 = tsdfRegisNo1_
End Property
Public Property Let TSDFRegisNo1(ByVal newval As String)
    If Trim(tsdfRegisNo1_) <> Trim(newval) Then
        tsdfRegisNo1_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get TSDFName2() As String
    TSDFName2 = tsdfName2_
End Property
Public Property Let TSDFName2(ByVal newval As String)
    If Trim(tsdfName2_) <> Trim(newval) Then
        tsdfName2_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get TSDFRegisNo2() As String
    TSDFRegisNo2 = tsdfRegisNo2_
End Property
Public Property Let TSDFRegisNo2(ByVal newval As String)
    If Trim(tsdfRegisNo2_) <> Trim(newval) Then
        tsdfRegisNo2_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get TotalLiquid() As Double
    TotalLiquid = totalLiquid_
End Property
Public Property Let TotalLiquid(ByVal newval As Double)
    If totalLiquid_ <> newval Then
        totalLiquid_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get LiquidUnitID() As String
    LiquidUnitID = liquidUnitID_
End Property
Public Property Let LiquidUnitID(ByVal newval As String)
    If Trim(liquidUnitID_) <> Trim(newval) Then
        liquidUnitID_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get LiquidUnitName() As String
    LiquidUnitName = liquidUnitName_
End Property
Public Property Let LiquidUnitName(ByVal newval As String)
    If Trim(liquidUnitName_) <> Trim(newval) Then
        liquidUnitName_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get TotalSolid() As Double
    TotalSolid = totalSolid_
End Property
Public Property Let TotalSolid(ByVal newval As Double)
    If totalSolid_ <> newval Then
        totalSolid_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get SolidUnitID() As String
    SolidUnitID = solidUnitID_
End Property
Public Property Let SolidUnitID(ByVal newval As String)
    If Trim(solidUnitID_) <> Trim(newval) Then
        solidUnitID_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get SolidUnitName() As String
    SolidUnitName = solidUnitName_
End Property
Public Property Let SolidUnitName(ByVal newval As String)
    If Trim(solidUnitName_) <> Trim(newval) Then
        solidUnitName_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get SpecialComment() As String
    SpecialComment = specialComment_
End Property
Public Property Let SpecialComment(ByVal newval As String)
    If Trim(specialComment_) <> Trim(newval) Then
        specialComment_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get TransTelNo1() As String
    TransTelNo1 = transTelNo1_
End Property
Public Property Let TransTelNo1(ByVal newval As String)
    If Trim(transTelNo1_) <> Trim(newval) Then
        transTelNo1_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get TransFaxNo1() As String
    TransFaxNo1 = transFaxNo1_
End Property
Public Property Let TransFaxNo1(ByVal newval As String)
    If Trim(transFaxNo1_) <> Trim(newval) Then
        transFaxNo1_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get TransErContact1() As String
    TransErContact1 = transErContact1_
End Property
Public Property Let TransErContact1(ByVal newval As String)
    If Trim(transErContact1_) <> Trim(newval) Then
        transErContact1_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get TruckSubTypeID() As String
    TruckSubTypeID = truckSubTypeID_
End Property
Public Property Let TruckSubTypeID(ByVal newval As String)
    If Trim(truckSubTypeID_) <> Trim(newval) Then
        truckSubTypeID_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get TruckSubTypeName() As String
    TruckSubTypeName = truckSubTypeName_
End Property
Public Property Let TruckSubTypeName(ByVal newval As String)
    If Trim(truckSubTypeName_) <> Trim(newval) Then
        truckSubTypeName_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get TruckTypeID() As String
    TruckTypeID = TruckTypeID_
End Property
Public Property Let TruckTypeID(ByVal newval As String)
    If Trim(TruckTypeID_) <> Trim(newval) Then
        TruckTypeID_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get TruckTypeName() As String
    TruckTypeName = truckTypeName_
End Property
Public Property Let TruckTypeName(ByVal newval As String)
    If Trim(truckTypeName_) <> Trim(newval) Then
        truckTypeName_ = Trim(newval)
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

Public Property Get TrailerCarRegisID() As String
    TrailerCarRegisID = TrailerCarRegisID_
End Property
Public Property Let TrailerCarRegisID(ByVal newval As String)
    If Trim(TrailerCarRegisID_) <> Trim(newval) Then
        TrailerCarRegisID_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get FromProvinceID() As String
    FromProvinceID = fromProvinceID_
End Property
Public Property Let FromProvinceID(ByVal newval As String)
    If Trim(fromProvinceID_) <> Trim(newval) Then
        fromProvinceID_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get FromProvinceName() As String
    FromProvinceName = fromProvinceName_
End Property
Public Property Let FromProvinceName(ByVal newval As String)
    If Trim(fromProvinceName_) <> Trim(newval) Then
        fromProvinceName_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get ToProvinceID() As String
    ToProvinceID = toProvinceID_
End Property
Public Property Let ToProvinceID(ByVal newval As String)
    If Trim(toProvinceID_) <> Trim(newval) Then
        toProvinceID_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get ToProvinceName() As String
    ToProvinceName = toProvinceName_
End Property
Public Property Let ToProvinceName(ByVal newval As String)
    If Trim(toProvinceName_) <> Trim(newval) Then
        toProvinceName_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get TimeSpending() As Double
    TimeSpending = timeSpending_
End Property
Public Property Let TimeSpending(ByVal newval As Double)
    If timeSpending_ <> newval Then
        timeSpending_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get DayOrHour() As String
    DayOrHour = dayOrHour_
End Property
Public Property Let DayOrHour(ByVal newval As String)
    If Trim(dayOrHour_) <> Trim(newval) Then
        dayOrHour_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get TSDFAddr() As String
    TSDFAddr = tsdfAddr_
End Property
Public Property Let TSDFAddr(ByVal newval As String)
    If Trim(tsdfAddr_) <> Trim(newval) Then
        tsdfAddr_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get TSDFTumbolID() As String
    TSDFTumbolID = tsdfTumbolID_
End Property
Public Property Let TSDFTumbolID(ByVal newval As String)
    If Trim(tsdfTumbolID_) <> Trim(newval) Then
        tsdfTumbolID_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get TSDFTumbolName() As String
    TSDFTumbolName = tsdfTumbolName_
End Property
Public Property Let TSDFTumbolName(ByVal newval As String)
    If Trim(tsdfTumbolName_) <> Trim(newval) Then
        tsdfTumbolName_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get TSDFAumphurID() As String
    TSDFAumphurID = tsdfAumphurID_
End Property
Public Property Let TSDFAumphurID(ByVal newval As String)
    If Trim(tsdfAumphurID_) <> Trim(newval) Then
        tsdfAumphurID_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get TSDFAumphurName() As String
    TSDFAumphurName = tsdfAumphurName_
End Property
Public Property Let TSDFAumphurName(ByVal newval As String)
    If Trim(tsdfAumphurName_) <> Trim(newval) Then
        tsdfAumphurName_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get TSDFProvinceName() As String
    TSDFProvinceName = tsdfProvinceName_
End Property
Public Property Let TSDFProvinceName(ByVal newval As String)
    If Trim(tsdfProvinceName_) <> Trim(newval) Then
        tsdfProvinceName_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get TSDFTelNo() As String
    TSDFTelNo = tsdfTelNo_
End Property
Public Property Let TSDFTelNo(ByVal newval As String)
    If Trim(tsdfTelNo_) <> Trim(newval) Then
        tsdfTelNo_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get TSDFFaxNo() As String
    TSDFFaxNo = tsdfFaxNo_
End Property
Public Property Let TSDFFaxNo(ByVal newval As String)
    If Trim(tsdfFaxNo_) <> Trim(newval) Then
        tsdfFaxNo_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get TSDFErContact() As String
    TSDFErContact = tsdfErContact_
End Property
Public Property Let TSDFErContact(ByVal newval As String)
    If Trim(tsdfErContact_) <> Trim(newval) Then
        tsdfErContact_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get TreatmentPeriod() As Double
    TreatmentPeriod = treatmentPeriod_
End Property
Public Property Let TreatmentPeriod(ByVal newval As Double)
    If treatmentPeriod_ <> newval Then
        treatmentPeriod_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get PeriodUnit() As String
    PeriodUnit = periodUnit_
End Property
Public Property Let PeriodUnit(ByVal newval As String)
    If Trim(periodUnit_) <> Trim(newval) Then
        periodUnit_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get ActualQty() As Double
    ActualQty = actualQty_
End Property
Public Property Let ActualQty(ByVal newval As Double)
    If actualQty_ <> newval Then
        actualQty_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get QtyUnitID() As String
    QtyUnitID = qtyUnitID_
End Property
Public Property Let QtyUnitID(ByVal newval As String)
    If Trim(qtyUnitID_) <> Trim(newval) Then
        qtyUnitID_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get QtyUnitName() As String
    QtyUnitName = qtyUnitName_
End Property
Public Property Let QtyUnitName(ByVal newval As String)
    If Trim(qtyUnitName_) <> Trim(newval) Then
        qtyUnitName_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get OpenStaffID() As String
    OpenStaffID = openStaffID_
End Property
Public Property Let OpenStaffID(ByVal newval As String)
    If Trim(openStaffID_) <> Trim(newval) Then
        openStaffID_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get OpenStaffName() As String
    OpenStaffName = openStaffName_
End Property
Public Property Let OpenStaffName(ByVal newval As String)
    If Trim(openStaffName_) <> Trim(newval) Then
        openStaffName_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get OpenDate() As String
    OpenDate = openDate_
End Property
Public Property Let OpenDate(ByVal newval As String)
    If Trim(openDate_) <> Trim(newval) Then
        openDate_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get WorkDate() As String
    WorkDate = WorkDate_
End Property
Public Property Let WorkDate(ByVal newval As String)
    If Trim(WorkDate_) <> Trim(newval) Then
        WorkDate_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get CloseStaffID() As String
    CloseStaffID = closeStaffID_
End Property
Public Property Let CloseStaffID(ByVal newval As String)
    If Trim(closeStaffID_) <> Trim(newval) Then
        closeStaffID_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get CloseStaffName() As String
    CloseStaffName = closeStaffName_
End Property
Public Property Let CloseStaffName(ByVal newval As String)
    If Trim(closeStaffName_) <> Trim(newval) Then
        closeStaffName_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get CloseDate() As String
    CloseDate = closeDate_
End Property
Public Property Let CloseDate(ByVal newval As String)
    If Trim(closeDate_) <> Trim(newval) Then
        closeDate_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get CloseTime() As String
    CloseTime = closeTime_
End Property
Public Property Let CloseTime(ByVal newval As String)
    If Trim(closeTime_) <> Trim(newval) Then
        closeTime_ = Trim(newval)
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
Public Property Let isClosed(ByVal newval As Boolean)
Dim tmpStr As String
    If newval Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(isClosed_) <> Trim(tmpStr) Then
        isClosed_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property

Public Property Get ApproveStaffID() As String
    ApproveStaffID = approveStaffID_
End Property
Public Property Let ApproveStaffID(ByVal newval As String)
    If Trim(approveStaffID_) <> Trim(newval) Then
        approveStaffID_ = Trim(newval)
'        isChanged_ = True
    End If
End Property

Public Property Get ApproveStaffName() As String
    ApproveStaffName = approveStaffName_
End Property
Public Property Let ApproveStaffName(ByVal newval As String)
    If Trim(approveStaffName_) <> Trim(newval) Then
        approveStaffName_ = Trim(newval)
'        isChanged_ = True
    End If
End Property

Public Property Get ApproveDate() As String
    ApproveDate = approveDate_
End Property
Public Property Let ApproveDate(ByVal newval As String)
    If Trim(approveDate_) <> Trim(newval) Then
        approveDate_ = Trim(newval)
'        isChanged_ = True
    End If
End Property

Public Property Get ApproveTime() As String
    ApproveTime = ApproveTime_
End Property
Public Property Let ApproveTime(ByVal newval As String)
    If Trim(ApproveTime_) <> Trim(newval) Then
        ApproveTime_ = Trim(newval)
'        isChanged_ = True
    End If
End Property

Public Property Get isApproved() As Boolean
    If Trim(isApproved_) = "Y" Then
        isApproved = True
    Else
        isApproved = False
    End If
End Property
Public Property Let isApproved(ByVal newval As Boolean)
Dim tmpStr As String
    If newval Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(isApproved_) <> Trim(tmpStr) Then
        isApproved_ = Trim(tmpStr)
'        isChanged_ = True
    End If
End Property

Public Property Get IsCanceled() As Boolean
    If Trim(IsCanceled_) = "Y" Then
        IsCanceled = True
    Else
        IsCanceled = False
    End If
End Property
Public Property Let IsCanceled(ByVal newval As Boolean)
Dim tmpStr As String
    If newval Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(IsCanceled_) <> Trim(tmpStr) Then
        IsCanceled_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property

Public Property Get CancelDate() As String
    CancelDate = cancelDate_
End Property
Public Property Let CancelDate(ByVal newval As String)
    If Trim(cancelDate_) <> Trim(newval) Then
        cancelDate_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get CancelDetail() As String
    CancelDetail = CancelDetail_
End Property
Public Property Let CancelDetail(ByVal newval As String)
    If Trim(CancelDetail_) <> Trim(newval) Then
        CancelDetail_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get CancelStaffID() As String
    CancelStaffID = cancelStaffID_
End Property
Public Property Let CancelStaffID(ByVal newval As String)
    If Trim(cancelStaffID_) <> Trim(newval) Then
        cancelStaffID_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get CancelStaffName() As String
    CancelStaffName = cancelStaffName_
End Property
Public Property Let CancelStaffName(ByVal newval As String)
    If Trim(cancelStaffName_) <> Trim(newval) Then
        cancelStaffName_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get DriverName() As String
    DriverName = DriverName_
End Property
Public Property Let DriverName(ByVal newval As String)
    If Trim(DriverName_) <> Trim(newval) Then
        DriverName_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get DriverTelNo() As String
    DriverTelNo = DriverTelNo_
End Property
Public Property Let DriverTelNo(ByVal newval As String)
    If Trim(DriverTelNo_) <> Trim(newval) Then
        DriverTelNo_ = Trim(newval)
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
Public Property Let isPrinted(ByVal newval As Boolean)
Dim tmpStr As String
    If newval Then
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
Public Property Let CustMenifestNo(ByVal newval As String)
    If Trim(custMenifestNo_) <> Trim(newval) Then
        custMenifestNo_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get NonOrHz() As String
    NonOrHz = nonOrHz_
End Property
Public Property Let NonOrHz(ByVal newval As String)
    If Trim(nonOrHz_) <> Trim(newval) Then
        nonOrHz_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get Transport2ID() As String
    Transport2ID = transport2ID_
End Property
Public Property Let Transport2ID(ByVal newval As String)
    If Trim(transport2ID_) <> Trim(newval) Then
        transport2ID_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get TSDF1ID() As String
    TSDF1ID = tsdf1ID_
End Property
Public Property Let TSDF1ID(ByVal newval As String)
    If Trim(tsdf1ID_) <> Trim(newval) Then
        tsdf1ID_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get TSDF2ID() As String
    TSDF2ID = tsdf2ID_
End Property
Public Property Let TSDF2ID(ByVal newval As String)
    If Trim(tsdf2ID_) <> Trim(newval) Then
        tsdf2ID_ = Trim(newval)
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
Public Property Let DocJobRecord(ByVal newval As Boolean)
Dim tmpStr As String
    If newval Then
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
Public Property Let DocControl(ByVal newval As Boolean)
Dim tmpStr As String
    If newval Then
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
Public Property Let DocSiteWeight(ByVal newval As Boolean)
Dim tmpStr As String
    If newval Then
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
Public Property Let DocCustWeight(ByVal newval As Boolean)
Dim tmpStr As String
    If newval Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(docCustWeight_) <> Trim(tmpStr) Then
        docCustWeight_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property
Public Property Get IsPost() As Boolean
    If Trim(IsPost_) = "Y" Then
        IsPost = True
    Else
        IsPost = False
    End If
End Property
Public Property Let IsPost(ByVal newval As Boolean)
Dim tmpStr As String
    If newval Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(IsPost_) <> Trim(tmpStr) Then
        IsPost_ = Trim(tmpStr)
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
Public Property Let DocCustControl(ByVal newval As Boolean)
Dim tmpStr As String
    If newval Then
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
Public Property Let isDocCompleted(ByVal newval As Boolean)
Dim tmpStr As String
    If newval Then
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
Public Property Let InCollectFee(ByVal newval As Double)
    If inCollectFee_ <> newval Then
        inCollectFee_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get IntransFee() As Double
    IntransFee = inTransFee_
End Property
Public Property Let IntransFee(ByVal newval As Double)
    If inTransFee_ <> newval Then
        inTransFee_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get InWorkerFee() As Double
    InWorkerFee = inWorkerFee_
End Property
Public Property Let InWorkerFee(ByVal newval As Double)
    If inWorkerFee_ <> newval Then
        inWorkerFee_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get InEquipFee() As Double
    InEquipFee = inEquipFee_
End Property
Public Property Let InEquipFee(ByVal newval As Double)
    If inEquipFee_ <> newval Then
        inEquipFee_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get InOtherDesc() As String
    InOtherDesc = Trim(inOtherDesc_)
End Property
Public Property Let InOtherDesc(ByVal newval As String)
    If Trim(inOtherDesc_) <> Trim(newval) Then
        inOtherDesc_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get InOtherFee() As Double
    InOtherFee = inOtherFee_
End Property
Public Property Let InOtherFee(ByVal newval As Double)
    If inOtherFee_ <> newval Then
        inOtherFee_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get OutTransFee() As Double
    OutTransFee = outTransFee_
End Property
Public Property Let OutTransFee(ByVal newval As Double)
    If outTransFee_ <> newval Then
        outTransFee_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get OutWorkerCount() As Integer
    OutWorkerCount = outWorkerCount_
End Property
Public Property Let OutWorkerCount(ByVal newval As Integer)
    If outWorkerCount_ <> newval Then
        outWorkerCount_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get OutWorkerFee() As Double
    OutWorkerFee = outWorkerFee_
End Property
Public Property Let OutWorkerFee(ByVal newval As Double)
    If outWorkerFee_ <> newval Then
        outWorkerFee_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get OutEquipFee() As Double
    OutEquipFee = outEquipFee_
End Property
Public Property Let OutEquipFee(ByVal newval As Double)
    If outEquipFee_ <> newval Then
        outEquipFee_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get OutOtherDesc() As String
    OutOtherDesc = outOtherDesc_
End Property
Public Property Let OutOtherDesc(ByVal newval As String)
    If Trim(outOtherDesc_) <> Trim(newval) Then
        outOtherDesc_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get OutOtherFee() As Double
    OutOtherFee = outOtherFee_
End Property
Public Property Let OutOtherFee(ByVal newval As Double)
    If outOtherFee_ <> newval Then
        outOtherFee_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get AccStaffID() As String
    AccStaffID = accStaffID_
End Property
Public Property Let AccStaffID(ByVal newval As String)
    If Trim(accStaffID_) <> Trim(newval) Then
        accStaffID_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get AccStaffName() As String
    AccStaffName = accStaffName_
End Property
Public Property Let AccStaffName(ByVal newval As String)
    If Trim(accStaffName_) <> Trim(newval) Then
        accStaffName_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get AccLastUpdate() As String
    AccLastUpdate = accLastUpdate_
End Property
Public Property Let AccLastUpdate(ByVal newval As String)
    If Trim(accLastUpdate_) <> Trim(newval) Then
        accLastUpdate_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get SignName1() As String
    SignName1 = SignName1_
End Property
Public Property Let SignName1(ByVal newval As String)
    If Trim(SignName1_) <> Trim(newval) Then
        SignName1_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get SignDate1() As String
    SignDate1 = signDate1_
End Property
Public Property Let SignDate1(ByVal newval As String)
    If Trim(signDate1_) <> Trim(newval) Then
        signDate1_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get SignName2() As String
    SignName2 = signName2_
End Property
Public Property Let SignName2(ByVal newval As String)
    If Trim(signName2_) <> Trim(newval) Then
        signName2_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get SignDate2() As String
    SignDate2 = signDate2_
End Property
Public Property Let SignDate2(ByVal newval As String)
    If Trim(signDate2_) <> Trim(newval) Then
        signDate2_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get SignName3() As String
    SignName3 = signName3_
End Property
Public Property Let SignName3(ByVal newval As String)
    If Trim(signName3_) <> Trim(newval) Then
        signName3_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get SignDate3() As String
    SignDate3 = signDate3_
End Property
Public Property Let SignDate3(ByVal newval As String)
    If Trim(signDate3_) <> Trim(newval) Then
        signDate3_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get TimeOutSite() As String
    TimeOutSite = timeOutSite_
End Property
Public Property Let TimeOutSite(ByVal newval As String)
    If Trim(timeOutSite_) <> Trim(newval) Then
        timeOutSite_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get DateArriveCust() As String
    DateArriveCust = dateArriveCust_
End Property
Public Property Let DateArriveCust(ByVal newval As String)
    If Trim(dateArriveCust_) <> Trim(newval) Then
        dateArriveCust_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get TimeArriveCust() As String
    TimeArriveCust = timeArriveCust_
End Property
Public Property Let TimeArriveCust(ByVal newval As String)
    If Trim(timeArriveCust_) <> Trim(newval) Then
        timeArriveCust_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get DateOutCust() As String
    DateOutCust = dateOutCust_
End Property
Public Property Let DateOutCust(ByVal newval As String)
    If Trim(dateOutCust_) <> Trim(newval) Then
        dateOutCust_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get TimeOutCust() As String
    TimeOutCust = timeOutCust_
End Property
Public Property Let TimeOutCust(ByVal newval As String)
    If Trim(timeOutCust_) <> Trim(newval) Then
        timeOutCust_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get DateArriveSite() As String
    DateArriveSite = dateArriveSite_
End Property
Public Property Let DateArriveSite(ByVal newval As String)
    If Trim(dateArriveSite_) <> Trim(newval) Then
        dateArriveSite_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get TimeArriveSite() As String
    TimeArriveSite = timeArriveSite_
End Property
Public Property Let TimeArriveSite(ByVal newval As String)
    If Trim(timeArriveSite_) <> Trim(newval) Then
        timeArriveSite_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get isChanged() As Boolean
    isChanged = isChanged_
End Property
Public Property Let isChanged(ByVal newval As Boolean)
    isChanged_ = newval
End Property

Public Property Get isSave() As Boolean
    isSave = isSave_
End Property
Public Property Let isSave(ByVal newval As Boolean)
    isSave_ = newval
End Property

Public Property Get StaffTelNo() As String
    StaffTelNo = StaffTelNo_
End Property
Public Property Let StaffTelNo(ByVal newval As String)
    If Trim(StaffTelNo_) <> Trim(newval) Then
        StaffTelNo_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get CarRem() As String
    CarRem = CarRem_
End Property
Public Property Let CarRem(ByVal newval As String)
    If Trim(CarRem_) <> Trim(newval) Then
        CarRem_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get isJob() As String
    isJob = isJob_
End Property
Public Property Let isJob(ByVal newval As String)
    If Trim(isJob_) <> Trim(newval) Then
        isJob_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get LastUser() As String
    LastUser = LastUser_
End Property
Public Property Let LastUser(ByVal newval As String)
    If Trim(LastUser_) <> Trim(newval) Then
        LastUser_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get LastUpdate() As String
    LastUpdate = LastUpdate_
End Property
Public Property Let LastUpdate(ByVal newval As String)
    If Trim(LastUpdate_) <> Trim(newval) Then
        LastUpdate_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get CommentList() As String
    CommentList = CommentList_
End Property
Public Property Let CommentList(ByVal newval As String)
    If Trim(CommentList_) <> Trim(newval) Then
        CommentList_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get REState() As String
    REState = REState_
End Property
Public Property Let REState(ByVal newval As String)
    If Trim(REState_) <> Trim(newval) Then
        REState_ = Trim(newval)
        isChanged_ = True
    End If
End Property


Public Property Get InvState() As String
    InvState = InvState_
End Property
Public Property Let InvState(ByVal newval As String)
    If Trim(InvState_) <> Trim(newval) Then
        InvState_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get InvDetail() As String
    InvDetail = InvDetail_
End Property
Public Property Let InvDetail(ByVal newval As String)
    If Trim(InvDetail_) <> Trim(newval) Then
        InvDetail_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get TankType() As String
    TankType = TankType_
End Property
Public Property Let TankType(ByVal newval As String)
    If Trim(TankType_) <> Trim(newval) Then
        TankType_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get Tankcount() As String
    Tankcount = Tankcount_
End Property
Public Property Let Tankcount(ByVal newval As String)
    If Trim(Tankcount_) <> Trim(newval) Then
        Tankcount_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get TankTypeSCount() As String
    TankTypeSCount = TankTypeSCount_
End Property
Public Property Let TankTypeSCount(ByVal newval As String)
    If Trim(TankTypeSCount_) <> Trim(newval) Then
        TankTypeSCount_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get BagBack() As String
    BagBack = BagBack_
End Property
Public Property Let BagBack(ByVal newval As String)
    If Trim(BagBack_) <> Trim(newval) Then
        BagBack_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get Container() As String
    Container = Container_
End Property
Public Property Let Container(ByVal newval As String)
    If Trim(Container_) <> Trim(newval) Then
        Container_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get Lack() As String
    Lack = Lack_
End Property
Public Property Let Lack(ByVal newval As String)
    If Trim(Lack_) <> Trim(newval) Then
        Lack_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get PaperBox() As String
    PaperBox = PaperBox_
End Property
Public Property Let PaperBox(ByVal newval As String)
    If Trim(PaperBox_) <> Trim(newval) Then
        PaperBox_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get BoutS() As String
    BoutS = BoutS_
End Property
Public Property Let BoutS(ByVal newval As String)
    If Trim(BoutS_) <> Trim(newval) Then
        BoutS_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get CovidBin() As String
     CovidBin = CovidBin_
End Property
Public Property Let CovidBin(ByVal newval As String)
    If Trim(CovidBin_) <> Trim(newval) Then
         CovidBin_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get LackNew() As String
     LackNew = LackNew_
End Property
Public Property Let LackNew(ByVal newval As String)
    If Trim(LackNew_) <> Trim(newval) Then
         LackNew_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get LackGas() As String
     LackGas = LackGas_
End Property
Public Property Let LackGas(ByVal newval As String)
    If Trim(LackGas_) <> Trim(newval) Then
         LackGas_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get Plate() As String
     Plate = Plate_
End Property
Public Property Let Plate(ByVal newval As String)
    If Trim(Plate_) <> Trim(newval) Then
         Plate_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get Plastic() As String
     Plastic = Plastic_
End Property
Public Property Let Plastic(ByVal newval As String)
    If Trim(Plastic_) <> Trim(newval) Then
         Plastic_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get BoutCage() As String
     BoutCage = BoutCage_
End Property
Public Property Let BoutCage(ByVal newval As String)
    If Trim(BoutCage_) <> Trim(newval) Then
         BoutCage_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get TankcountS() As String
    TankcountS = TankcountS_
End Property
Public Property Let TankcountS(ByVal newval As String)
    If Trim(TankcountS_) <> Trim(newval) Then
        TankcountS_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get BagType() As String
    BagType = BagType_
End Property
Public Property Let BagType(ByVal newval As String)
    If Trim(BagType_) <> Trim(newval) Then
        BagType_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get BagCount() As String
    BagCount = BagCount_
End Property
Public Property Let BagCount(ByVal newval As String)
    If Trim(BagCount_) <> Trim(newval) Then
        BagCount_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get RemID() As String
    RemID = RemID_
End Property
Public Property Let RemID(ByVal newval As String)
    If Trim(RemID_) <> Trim(newval) Then
        RemID_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get RemComment() As String
    RemComment = RemComment_
End Property
Public Property Let RemComment(ByVal newval As String)
    If Trim(RemComment_) <> Trim(newval) Then
        RemComment_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get DocCustDate() As String
    DocCustDate = DocCustDate_
End Property
Public Property Let DocCustDate(ByVal newval As String)
    If Trim(DocCustDate_) <> Trim(newval) Then
        DocCustDate_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get PostDate() As String
    PostDate = PostDate_
End Property
Public Property Let PostDate(ByVal newval As String)
    If Trim(PostDate_) <> Trim(newval) Then
        PostDate_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get EMSNo() As String
    EMSNo = EMSNo_
End Property
Public Property Let EMSNo(ByVal newval As String)
    If Trim(EMSNo_) <> Trim(newval) Then
        EMSNo_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get RefME() As String
    RefME = RefME_
End Property
Public Property Let RefME(ByVal newval As String)
    If Trim(RefME_) <> Trim(newval) Then
        RefME_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get PONo() As String
    PONo = PONo_
End Property
Public Property Let PONo(ByVal newval As String)
    If Trim(PONo_) <> Trim(newval) Then
        PONo_ = Trim(newval)
        isChanged_ = True
    End If
End Property
'*** New Variable for AP
Public Property Get APManifestNo() As String
    APManifestNo = APManifestNo_
End Property
Public Property Let APManifestNo(ByVal newval As String)
    If Trim(APManifestNo_) <> Trim(newval) Then
        APManifestNo_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get PhotoOnWeight() As String
    PhotoOnWeight = PhotoOnWeight_
End Property
Public Property Let PhotoOnWeight(ByVal newval As String)
    If Trim(PhotoOnWeight_) <> Trim(newval) Then
        PhotoOnWeight_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get PhotoOnBelt() As String
    PhotoOnBelt = PhotoOnBelt_
End Property
Public Property Let PhotoOnBelt(ByVal newval As String)
    If Trim(PhotoOnBelt_) <> Trim(newval) Then
        PhotoOnBelt_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get NewCustJob() As String
    NewCustJob = NewCustJob_
End Property
Public Property Let NewCustJob(ByVal newval As String)
    If Trim(NewCustJob_) <> Trim(newval) Then
        NewCustJob_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get PhotoOnBoxfeed() As String
    PhotoOnBoxfeed = PhotoOnBoxfeed_
End Property
Public Property Let PhotoOnBoxfeed(ByVal newval As String)
    If Trim(PhotoOnBoxfeed_) <> Trim(newval) Then
        PhotoOnBoxfeed_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get CustCon() As String
    CustCon = CustCon_
End Property
Public Property Let CustCon(ByVal newval As String)
    If Trim(CustCon_) <> Trim(newval) Then
        CustCon_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get CustPhoto() As String
    CustPhoto = CustPhoto_
End Property
Public Property Let CustPhoto(ByVal newval As String)
    If Trim(CustPhoto_) <> Trim(newval) Then
        CustPhoto_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get EmpTrain() As String
    EmpTrain = EmpTrain_
End Property
Public Property Let EmpTrain(ByVal newval As String)
    If Trim(EmpTrain_) <> Trim(newval) Then
        EmpTrain_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get CarCheck() As String
    CarCheck = CarCheck_
End Property
Public Property Let CarCheck(ByVal newval As String)
    If Trim(CarCheck_) <> Trim(newval) Then
        CarCheck_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get TakePhoto() As String
    TakePhoto = takePhoto_
End Property
Public Property Let TakePhoto(ByVal newval As String)
    If Trim(takePhoto_) <> Trim(newval) Then
        takePhoto_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get CloseWeight() As String
    CloseWeight = CloseWeight_
End Property
Public Property Let CloseWeight(ByVal newval As String)
    If Trim(CloseWeight_) <> Trim(newval) Then
        CloseWeight_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get SaleDistricIDOrder() As String
    SaleDistricIDOrder = SaleDistricIDOrder_
End Property
Public Property Let SaleDistricIDOrder(ByVal newval As String)
    If Trim(SaleDistricIDOrder_) <> Trim(newval) Then
        SaleDistricIDOrder_ = Trim(newval)
        isChanged_ = True
    End If
End Property
'*** End Variable for Ap

Public Sub Save()
On Error GoTo ErrD
Dim tmpRec As ADODB.Recordset
Dim x%, tmpJobDet As BWGJobDataDetail
Dim sqlStr As String
    isSave_ = True
    sqlStr = "SELECT *, CONVERT(DATETIME, GETDATE(), 103) AS GETDATE FROM tbMenifestHeader WHERE MenifestID='" & id_ & "'"
    Set tmpRec = New ADODB.Recordset
    GetDBConnection
    tmpRec.open sqlStr, DBConn, adOpenDynamic, adLockOptimistic
    If tmpRec.EOF Then
        tmpRec.AddNew
        tmpRec("MenifestID") = id_
        tmpRec("OpenStaffID") = Trim(openStaffID_)
        tmpRec("OpenDate") = Trim(openDate_)
        tmpRec("CloseStaffID") = Trim(closeStaffID_)
        tmpRec("CloseDate") = Trim(closeDate_)
        tmpRec("CloseTime") = Trim(closeTime_)
        tmpRec("isClosed") = Trim(isClosed_)
        tmpRec("ApproveStaffID") = Trim(approveStaffID_)
        tmpRec("ApproveDate") = Trim(approveDate_)
        tmpRec("ApproveTime") = Trim(ApproveTime_)
        tmpRec("isApproved") = Trim(isApproved_)
        tmpRec("isCanceled") = Trim(IsCanceled_)
        tmpRec("CancelDate") = Trim(cancelDate_)
        tmpRec("CancelStaffID") = Trim(cancelStaffID_)
        tmpRec("isPrinted") = Trim(isPrinted_)
    End If
    tmpRec("DocNo") = Trim(DocNo_)
    tmpRec("MenifestNo") = Trim(MenifestNo_)
    tmpRec("CustomerID") = Trim(CustomerID_)
    tmpRec("CustSiteID") = Trim(custSiteID_)
    tmpRec("WasteGenNo") = Trim(WasteGenNo_)
    tmpRec("TelNo") = Trim(telNo_)
    tmpRec("FaxNo") = Trim(faxNo_)
    tmpRec("ErContact") = Trim(erContact_)
    tmpRec("TransportName1") = Trim(transportName1_)
    tmpRec("TransRegisNo1") = Trim(TransRegisNo1_)
    tmpRec("TransportName2") = Trim(TransportName2_)
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
    tmpRec("CarRegisID") = Trim(CarRegisID_)
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
    tmpRec("WorkDate") = Trim(WorkDate_)
    tmpRec("TSDFTumbolID") = Trim(tsdfTumbolID_)
    tmpRec("DriverName") = Trim(DriverName_)
    tmpRec("CustMenifestNo") = Trim(custMenifestNo_)
    tmpRec("NonOrHz") = Trim(nonOrHz_)
    tmpRec("Transport2ID") = Trim(transport2ID_)
    tmpRec("TSDF1ID") = Trim(tsdf1ID_)
    tmpRec("TSDF2ID") = Trim(tsdf2ID_)
    tmpRec("PONo") = Trim(PONo_)
    
    If Trim(docJobRecord_) = "" Then docJobRecord_ = "N"
    If Trim(docControl_) = "" Then docControl_ = "N"
    If Trim(docSiteWeight_) = "" Then docSiteWeight_ = "N"
    If Trim(docCustWeight_) = "" Then docCustWeight_ = "N"
    If Trim(docCustControl_) = "" Then docCustControl_ = "N"
    If Trim(isDocCompleted_) = "" Then isDocCompleted_ = "N"
    
    If Trim(RefME_) <> "" Then tmpRec("RefME") = Trim(RefME_)

    tmpRec("DocJobRecord") = Trim(docJobRecord_)
    tmpRec("DocControl") = Trim(docControl_)
    tmpRec("DocSiteWeight") = Trim(docSiteWeight_)
    tmpRec("DocCustWeight") = Trim(docCustWeight_)
    tmpRec("DocCustControl") = Trim(docCustControl_)
    tmpRec("isDocCompleted") = Trim(isDocCompleted_)
    
    tmpRec("SignName1") = Trim(SignName1_)
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
    If Trim(ChargeWType_) = "" Then ChargeWType_ = "A"
    tmpRec("ChargeWType") = ChargeWType_
    tmpRec("SaleDistrictID") = saleDistrictID_
    tmpRec("UnderSaleStaffID") = underStaffID_
    
    tmpRec("JobDataCarID") = JobDataCarID_
    tmpRec("TripNo") = TripNo_
    
    tmpRec("InvState") = InvState_
    tmpRec("REState") = REState_
    tmpRec("InvDetail") = InvDetail_
    tmpRec("CommentList") = CommentList_
    
    tmpRec!ModifyUser = Trim(Trim(CurrentUser.StaffFName) & " " & Trim(CurrentUser.StaffLName))
    tmpRec!ModifyDate = tmpRec!GetDate
    tmpRec!RemID = RemID_
    tmpRec!RemComment = RemComment_
    
    tmpRec.Update
    tmpRec.Requery
    DocNo_ = Trim("" & tmpRec("DocNo"))
    'DocNoRef_ = Trim("" & tmpRec("DocNoRef_"))
    MenifestNo_ = Trim("" & tmpRec("MenifestNo"))
    APManifestNo_ = Trim("" & tmpRec("APManifestNo"))
'    tmpRec.Close
    
    '**** Update Vendor *******
    
    sqlStr = "UPDATE tbJobDataTimeTable SET gpssend='" & GPSSend_ & "' ,SendDesc='" & SendDesc_ & "',Carrate='" & CarRate_ & "' , VendorTruckID = (SELECT TOP(1) VendorTruckID FROM vw_VendorCarSearch " & vbCrLf
    sqlStr = sqlStr & "WHERE VendorID = '" & Trim(transport2ID_) & "' AND CarRegisID = '" & Trim(CarRegisID_) & "')" & vbCrLf
    sqlStr = sqlStr & "WHERE JobDataCarID = '" & JobDataCarID_ & "' AND TripNO = '" & TripNo_ & "'"
    DBConnExc sqlStr

    '*******************************
    
'    DBConn.BeginTrans
'    With DBConn
        If Not colJobDet_ Is Nothing Then
            DBConnExc "Delete tbMenifestDetail WHERE MenifestID='" & id_ & "'"
            For x = 1 To colJobDet_.Count
                Set tmpJobDet = colJobDet_(x)
                DBConnExc "Insert into tbMenifestDetail (MenifestID,JobDetID,ContainerCount,ContainerDesc,CustWeight,DisposerWeight,TreatmentCharge,InWeight) Values ('" & id_ & "','" & tmpJobDet.id & "'," & tmpJobDet.ContainerCount & ",'" & tmpJobDet.ContainerTypeDesc & "'," & tmpJobDet.EstWasteQty & "," & tmpJobDet.DisposerWeight & "," & tmpJobDet.TreatmentCharge & "," & tmpJobDet.InWeight & ")"
                tmpJobDet.Save
                DBConnExc "Delete tbJobFail WHERE MenifestID='" & id_ & "' And JobDetID='" & tmpJobDet.id & "'"
                If Trim(tmpJobDet.FailType) <> "" Then
                    DBConnExc "Insert into tbJobFail (JobDetID,MenifestID,FailType,FailRemark,FailCheck) Values ('" & tmpJobDet.id & "','" & id_ & "','" & tmpJobDet.FailType & "','" & tmpJobDet.FailRemark & "', " & tmpJobDet.FailCheck & ")"
                End If
                Set tmpJobDet = Nothing
            Next
            DBConnExc "Delete tbMenifestWasteAdded WHERE MenifestID='" & id_ & "'"
            If Not colWasteAdded_ Is Nothing Then
                For x = 1 To colWasteAdded_.Count
                    colWasteAdded_(x).Save
                Next
            End If
        End If
'    End With
'    DBConn.CommitTrans
    Set tmpRec = Nothing
    isChanged_ = False
    CloseDBConnection
    Exit Sub
ErrD:
    isSave_ = False
'    DBConn.RollbackTrans
    MsgBox err.Description & vbCrLf & "เกิดข้อผิดพลาด การบันทึกถูกยกเลิก โปรแกรมตรวจสอบข้อมูลที่บันทึกลงไป หลังจากปิดหน้าต่างนี้", vbCritical, "Error"
    Exit Sub
End Sub

Public Function Update(ByVal UpdateString As String) As String
On Error GoTo ErrD:
    DBConnExc UpdateString
    Exit Function
ErrD:
    Update = Trim(err.Description)
    Exit Function
End Function

Public Function UpdateDoc() As String
On Error GoTo ErrD
Dim tmpRec As ADODB.Recordset
Dim x%, tmpJobDet As BWGJobDataDetail
Dim sqlStr As String
    sqlStr = "SELECT *, CONVERT(DATETIME, GETDATE(), 103) AS GETDATE FROM tbMenifestHeader WHERE MenifestID='" & id_ & "'"
    Set tmpRec = New ADODB.Recordset
    GetDBConnection
    tmpRec.open sqlStr, DBConn, adOpenDynamic, adLockOptimistic
    If tmpRec.EOF = False Then
        If Trim(docJobRecord_) = "" Then docJobRecord_ = "N"
        If Trim(docControl_) = "" Then docControl_ = "N"
        If Trim(docSiteWeight_) = "" Then docSiteWeight_ = "N"
        If Trim(docCustWeight_) = "" Then docCustWeight_ = "N"
        If Trim(docCustControl_) = "" Then docCustControl_ = "N"
        If Trim(isDocCompleted_) = "" Then isDocCompleted_ = "N"
        If Trim(IsPost_) = "" Then IsPost_ = "N"
        
        tmpRec("DocJobRecord") = Trim(docJobRecord_)
        tmpRec("DocControl") = Trim(docControl_)
        tmpRec("DocSiteWeight") = Trim(docSiteWeight_)
        tmpRec("DocCustWeight") = Trim(docCustWeight_)
        tmpRec("DocCustControl") = Trim(docCustControl_)
        tmpRec("isDocCompleted") = Trim(isDocCompleted_)
        tmpRec("CommentList") = Trim(CommentList_)
        If Trim(DocCustDate_) <> "" Then
         tmpRec("DocCustDate") = Trim(DocCustDate_)
         End If
         tmpRec("IsPost") = Trim(IsPost_)
         tmpRec("PostDate") = Trim(PostDate_)
         tmpRec("EMSNo") = Trim(EMSNo_)
'        tmpRec!ModifyUser = Trim(Trim(CurrentUser.StaffFName) & " " & Trim(CurrentUser.StaffLName))
'        tmpRec!ModifyDate = tmpRec!GetDate
        
        tmpRec.Update
        UpdateDoc = "Update Completed"
    End If
    Set tmpRec = Nothing
    CloseDBConnection
    Exit Function
ErrD:
    UpdateDoc = Trim(err.Description)
    Exit Function
End Function
Public Function Cancel(ByVal CancelDetail As String) As String
On Error GoTo delErr:
    DBConnExc "Update tbMenifestHeader Set isCanceled='Y',CancelDate='" & TodayDate & "',CancelStaffID='" & CurrentUser.id & "', CancelDetail = '" & CancelDetail & "' WHERE MenifestID='" & id_ & "'"
    DBConnExc "Update tbJobdatatimetable set menifestid='' WHERE MenifestID='" & id_ & "'"
    DBConnExc "Update tbJobDataHeader set isCompleted='N' WHERE JobID='" & createdFromJobID_ & "'"
    Cancel = ""
    IsCanceled_ = "Y"
    cancelDate_ = TodayDate
    cancelStaffID_ = CurrentUser.id
    cancelStaffName_ = CurrentUser.StaffFName & " " & CurrentUser.StaffLName
    Exit Function
delErr:
    Cancel = err.Description
End Function

Public Sub SetDataFromRecordset(selRec As ADODB.Recordset)
    id_ = Trim("" & selRec("MenifestID"))
    createdFromJobID_ = Trim("" & selRec("CreatedFromJobID"))
    companyID_ = Trim("" & selRec("CompanyID"))
    CompanyName_ = Trim("" & selRec("CompanyNameTH"))
    DocNo_ = Trim("" & selRec("DocNo"))
    DocNoRef_ = Trim("" & selRec("DocNoRef"))
    MenifestNo_ = Trim("" & selRec("MenifestNo"))
    CustomerID_ = Trim("" & selRec("CustomerID"))
    CustomerCode_ = Trim("" & selRec("CustomerCode"))
    CustomerName_ = Trim("" & selRec("CustomerName"))
    CommentTimeNo_ = Trim("" & selRec("CommentTimeNo"))
    custSiteID_ = Trim("" & selRec("CustSiteID"))
    custSiteName_ = Trim("" & selRec("SiteName"))
    If isExist(selRec, "SiteAumphurID") Then custSiteAumphurID_ = Trim("" & selRec("SiteAumphurID"))

    WasteGenNo_ = Trim("" & selRec("WasteGenNo"))
    jobNo_ = Trim("" & selRec("JobNo"))
    ChargeWType_ = Trim("" & selRec("ChargeWType"))
    telNo_ = Trim("" & selRec("TelNo"))
    faxNo_ = Trim("" & selRec("FaxNo"))
    erContact_ = Trim("" & selRec("ErContact"))
    transportName1_ = Trim("" & selRec("TransportName1"))
    TransRegisNo1_ = Trim("" & selRec("TransRegisNo1"))
    TransportName2_ = Trim("" & selRec("TransportName2"))
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
    TruckTypeID_ = Trim("" & selRec("TruckTypeID"))
    truckTypeName_ = Trim("" & selRec("TruckTypeDesc"))
    CarRegisID_ = Trim("" & selRec("CarRegisID"))
    TrailerCarRegisID_ = Trim("" & selRec("TrailerCarRegisID"))
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
    WorkDate_ = Trim("" & selRec("WorkDate"))
    closeStaffID_ = Trim("" & selRec("CloseStaffID"))
    closeStaffName_ = Trim("" & selRec("CloseStaffName"))
    closeDate_ = Trim("" & selRec("CloseDate"))
    closeTime_ = Trim("" & selRec("CloseTime"))
    isClosed_ = Trim("" & selRec("isClosed"))
    
    approveStaffID_ = Trim("" & selRec("ApproveStaffID"))
    approveStaffName_ = Trim("" & selRec("ApproveStaffName"))
    approveDate_ = Trim("" & selRec("ApproveDate"))
    ApproveTime_ = Trim("" & selRec("ApproveTime"))
    isApproved_ = Trim("" & selRec("isApproved"))
    
    IsCanceled_ = Trim("" & selRec("isCanceled"))
    cancelDate_ = Trim("" & selRec("CancelDate"))
    If isExist(selRec, "CancelDetail") Then CancelDetail_ = Trim("" & selRec("CancelDetail"))
    cancelStaffID_ = Trim("" & selRec("CancelStaffID"))
    cancelStaffName_ = Trim("" & selRec("CancelStaffName"))
    DriverName_ = Trim("" & selRec("DriverName"))
    DriverTelNo_ = Trim("" & selRec("DriverTelNo"))
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
    
    SignName1_ = Trim("" & selRec("SignName1"))
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
    If IsNull(selRec("SaleDistrictIDOrder")) = False Then SaleDistricIDOrder_ = selRec("SaleDistrictIDOrder")
    If IsNull(selRec("RemID")) = False Then RemID_ = Trim("" & selRec("RemID"))
    If IsNull(selRec("RemComment")) = False Then RemComment_ = Trim("" & selRec("RemComment"))
    
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
    If IsNull(selRec!SaleFName) Then
        underStaffName_ = Trim("" & selRec!SaleStaffName)
    Else
        underStaffName_ = Trim("" & selRec("SaleFName")) & " " & Trim("" & selRec("SaleLName"))
    End If
    
    If IsNull(selRec!StaffTelNo) Then
        StaffTelNo_ = Trim("" & selRec!SaleStaffTelNo)
    Else
        StaffTelNo_ = Trim("" & selRec("StaffTelNo"))
    End If
    
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
    If isExist(selRec, "TripTransFee_Inc") Then
        If IsNull(selRec("TripTransFee_Inc")) = False Then TripTransFeeInc_ = selRec("TripTransFee_Inc")
    End If
    
    isPriceIncTrans_ = Trim("" & selRec("isPriceIncTrans"))
    IsDiscountPriceTrans_ = Trim("" & selRec("IsDiscountPriceTrans"))

    billingID_ = Trim("" & selRec("BillingHeaderID"))
    BillingNo_ = Trim("" & selRec("BillingNo"))
    billPayID_ = Trim("" & selRec("BillPayHeaderID"))
    billPayNo_ = Trim("" & selRec("BillPayNo"))
    
    JobDataCarID_ = Trim("" & selRec("JobDataCarID"))
    If IsNull(selRec("TripNo")) = False Then TripNo_ = Trim("" & selRec("TripNo"))
    TimeTableNo_ = Trim("" & selRec("TimeTableNo"))
    isMain_ = Trim("" & selRec("isMain"))
    GPSSend_ = Trim("" & selRec("GPSSend"))
    SendDesc_ = Trim("" & selRec("SendDesc"))
    CarRate_ = Trim("" & selRec("CarRate"))

    CarRem_ = Trim("" & selRec("CarRem"))
    TankType_ = Trim("" & selRec("TankType"))
    Tankcount_ = Trim("" & selRec("TankCount"))
    TankcountS_ = Trim("" & selRec("TankCountS"))
    BagType_ = Trim("" & selRec("BagType"))
    BagCount_ = Trim("" & selRec("Bagcount"))
    PONo_ = Trim("" & selRec("PONo"))
    
    isManifestAddr_ = Trim("" & selRec("isManifestAddr"))
    
    'custFullAddr_ = Replace(Trim(selRec("CustAddr1") & " อ." & selRec("CustAumphurName") & " จ." & selRec("CustProvinceName") & " " & selRec("CustZipCode")), vbCrLf, "")
    'custFullAddr_ = Replace(Trim(selRec("CustAddr1") & " ต." & selRec("CustTumbolName") & " อ." & selRec("CustAumphurName") & " จ." & selRec("CustProvinceName") & " " & selRec("CustZipCode")), vbCrLf, "")
    'custSiteAddr_ = Replace(Trim(selRec("SiteAddr") & " ต." & selRec("SiteTumbolName") & " อ." & selRec("SiteAumphurName") & " จ." & selRec("SiteProvinceName") & " " & selRec("SiteZipCode")), vbCrLf, "")
    'custSiteAddr_ = Replace(Trim(selRec("SiteAddr") & " อ." & selRec("SiteAumphurName") & " จ." & selRec("SiteProvinceName") & " " & selRec("SiteZipCode")), vbCrLf, "")
    custFullAddr_ = Replace(Trim(selRec("SiteAddr") & " ต." & selRec("TumbolName") & " อ." & selRec("AumphurName") & " จ." & selRec("ProvinceName") & " " & selRec("SiteZipCode")), vbCrLf, "")
    custSiteAddr_ = Replace(Trim(selRec("SiteAddr") & " ต." & selRec("TumbolName") & " อ." & selRec("AumphurName") & " จ." & selRec("ProvinceName") & " " & selRec("SiteZipCode")), vbCrLf, "")
    
    CustTypeFlag_ = Trim("" & selRec("CustTypeFlag"))
    parentCustID_ = Trim("" & selRec("ParentCustID"))
    parentCustName_ = Trim("" & selRec("ParentCustName"))
    'CustTumbolName
    InvState_ = Trim("" & selRec("InvState"))
    REState_ = Trim("" & selRec("REState"))
    InvDetail_ = Trim("" & selRec("InvDetail"))

    
    '*** New Variable for AP
    APManifestNo_ = Trim("" & selRec("APManifestNo"))
    LastUser_ = Trim("" & selRec!ModifyUser)
    LastUpdate_ = Trim("" & selRec!ModifyDate)
    RefME_ = Trim("" & selRec!RefME)
    PhotoOnBoxfeed_ = Trim("" & selRec!PhotoOnBoxfeed)
    PhotoOnBelt = Trim("" & selRec!PhotoOnBelt)
    PhotoOnWeight_ = Trim("" & selRec!PhotoOnWeight)
    CustPhoto_ = Trim("" & selRec!CustPhoto)
    CustCon_ = Trim("" & selRec!CustCon)
    CloseWeight_ = Trim("" & selRec!CloseWeight)
    CarCheck_ = Trim("" & selRec!CarCheck)
    EmpTrain_ = Trim("" & selRec!EmpTrain)
    If isExist(selRec, "NewCustJob") Then
        NewCustJob_ = Trim("" & selRec!NewCustJob)
    End If
    If isExist(selRec, "TakePhoto") Then
        takePhoto_ = Trim("" & selRec!TakePhoto)
    End If
    If isExist(selRec, "InWeight") Then
        If IsNull(selRec("InWeight")) = False Then InWeight_ = selRec("InWeight")
    End If
    If isExist(selRec, "TankTypeSCount") Then
        If IsNull(selRec("TankTypeSCount")) = False Then TankTypeSCount_ = selRec("TankTypeSCount")
    End If
    If isExist(selRec, "Container") Then
        If IsNull(selRec("Container")) = False Then Container_ = selRec("Container")
    End If
    If isExist(selRec, "Lack") Then
        If IsNull(selRec("Lack")) = False Then Lack_ = selRec("Lack")
    End If
    If isExist(selRec, "PaperBox") Then
        If IsNull(selRec("PaperBox")) = False Then PaperBox_ = selRec("PaperBox")
    End If
    If isExist(selRec, "BoutS") Then
        If IsNull(selRec("BoutS")) = False Then BoutS_ = selRec("BoutS")
    End If
    If isExist(selRec, "CovidBin") Then
        If IsNull(selRec("CovidBin")) = False Then CovidBin_ = selRec("CovidBin")
    End If
    If isExist(selRec, "LackNew") Then
        If IsNull(selRec("LackNew")) = False Then LackNew_ = selRec("LackNew")
    End If
    If isExist(selRec, "LackGas") Then
        If IsNull(selRec("LackGas")) = False Then LackGas_ = selRec("LackGas")
    End If
    If isExist(selRec, "Plate") Then
        If IsNull(selRec("Plate")) = False Then Plate_ = selRec("Plate")
    End If
    If isExist(selRec, "Plastic") Then
        If IsNull(selRec("Plastic")) = False Then Plastic_ = selRec("Plastic")
    End If
    If isExist(selRec, "BoutCage") Then
        If IsNull(selRec("BoutCage")) = False Then BoutCage_ = selRec("BoutCage")
    End If
    If isExist(selRec, "sumCost") Then
        If IsNull(selRec("sumCost")) = False Then Cost_ = selRec("sumCost")
    End If
    
    AumphurName_ = Trim("" & selRec("AumphurName"))
    ProvinceName_ = Trim("" & selRec("ProvinceName"))
    OilUse_ = Trim("" & selRec("OilUse"))

End Sub

Public Function CloseMenifest() As String
On Error GoTo ErrD:
    CloseMenifest = ""
    DBConnExc "Update tbMenifestHeader Set CloseStaffID='" & CurrentUser.id & "', CloseDate='" & TodayDate & "',CloseTime=left(convert(varchar(10),getdate(),108),5),isClosed='" & isClosed_ & "',TimeOutSite='" & timeOutSite_ & "',DateArriveCust='" & dateArriveCust_ & "',TimeArriveCust='" & timeArriveCust_ & "',DateOutCust='" & dateOutCust_ & "',TimeOutCust='" & timeOutCust_ & "',DateArriveSite='" & dateArriveSite_ & "',TimeArriveSite='" & timeArriveSite_ & "' WHERE MenifestID='" & id_ & "'"
    closeStaffID_ = CurrentUser.id
    closeStaffName_ = CurrentUser.StaffFName & " " & CurrentUser.StaffLName
    closeDate_ = TodayDate
    closeTime_ = Format(Now, "hh:mm")
'    isClosed_ = "Y"
    
    isChanged_ = False
    If isClosed_ = "Y" Then
        Call UpdateTransportFee(TimeTableNo_)
        Call UpdateWasteAnalys
    End If
    Exit Function
ErrD:
    CloseMenifest = Trim(err.Description)
    Exit Function
End Function

Public Function ApproveMenifest() As String
On Error GoTo ErrD:
    ApproveMenifest = ""
    DBConnExc "Update tbMenifestHeader Set ApproveStaffID = '" & CurrentUser.id & "', ApproveDate = CONVERT(VARCHAR(10), GETDATE(), 111), ApproveTime = CONVERT(VARCHAR(10), GETDATE(),108), isApproved = '" & isApproved_ & "' WHERE MenifestID = '" & id_ & "'"
    approveStaffID_ = CurrentUser.id
    approveStaffName_ = Trim(CurrentUser.StaffFName & " " & CurrentUser.StaffLName)
'    approveDate_ = TodayDate
'    ApproveTime_ = Format(Now, "HH:MM:SS")
    Exit Function
ErrD:
    ApproveMenifest = Trim(err.Description)
    Exit Function
End Function

Public Sub UpdateTransportFee(ByVal TimeTableNo As String)
Dim tmpS As New BWGSearchManager
Dim amt1 As Double, amt2 As Double
Dim maxWorker%, maxTransFee As Double, tmpCol As Collection
Dim DiscountTrans As String
Dim x%, tmpChargeRate As BWGTransVendorChargeRate
    Set tmpCol = tmpS.JobDataTimeTableQuickSearch("TimeTableNo='" & TimeTableNo & "' And isMain='Y' And isCancel ='N' AND ISNULL(isCanceled, 'N') <> 'Y' AND TripNo = " & TripNo_)
    If tmpCol.Count > 0 Then
        amt1 = CalCustTreatFee(tmpCol(1))
        CalCustTripFee tmpCol(1), maxWorker, maxTransFee, DiscountTrans
        Set tmpChargeRate = CalTransFeeFromDistance(custSiteAumphurID_, tsdfAumphurID_, TruckTypeID_, WorkDate_)
        If Not tmpChargeRate Is Nothing Then
            If tmpCol(1).isTrailer Then
                amt2 = tmpChargeRate.TrailerRateTrip
            Else
                amt2 = tmpChargeRate.RatePerTrip
            End If
        End If
        DBConnExc "Update tbJobDataTimeTable Set TripTranFee=" & maxTransFee & ",TripTreatFee=" & amt1 & ",VDTransFee=" & amt2 & ",IsDiscountPriceTrans='" & DiscountTrans & "'  WHERE TimeTableNo='" & TimeTableNo & "' And isCancel='N'"
    End If
    Set tmpChargeRate = Nothing
    Set tmpCol = Nothing
    Set tmpS = Nothing
End Sub

Public Sub UpdateWasteAnalys()
Dim rs As New ADODB.Recordset, RS2 As ADODB.Recordset
Dim tSql As String
    tSql = "SELECT  WasteDataID, MenifestID, CustomerID, DocNo, WasteName, WasteType AS [Type], GetDate() AS WasteAppoveDate, 1750 AS PriceLab" & vbCrLf
    tSql = tSql & "From vw_MenifestDetailSearch" & vbCrLf
    tSql = tSql & "WHERE MenifestID = '" & id_ & "'"
    Set rs = New ADODB.Recordset
    Set rs = GetRS(tSql, adOpenForwardOnly, adLockReadOnly)
    Do While Not rs.EOF
        tSql = "SELECT *, CONVERT(DATETIME, GETDATE(), 103) AS GETDATE FROM tbWasteDataLab_AP WHERE WasteDataID = '" & rs!wastedataID & "'"
        GetDBConnection
        Set RS2 = New ADODB.Recordset
        RS2.open tSql, DBConn, adOpenKeyset, adLockOptimistic
        If RS2.EOF Then
            RS2.AddNew
            RS2!wastedataID = rs!wastedataID
            RS2!menifestID = rs!menifestID
            RS2!CustomerID = rs!CustomerID
            RS2!DocNo = rs!DocNo
            RS2!WasteName = rs!WasteName
            RS2!Type = rs!Type
            RS2!WasteAppoveDate = Format(rs!WasteAppoveDate, "yyyy/mm/dd")
            RS2!PriceLab = rs!PriceLab
            
            RS2!LastUser = Trim(Trim(CurrentUser.StaffFName) & " " & Trim(CurrentUser.StaffLName))
            RS2!LastUpdate = RS2!GetDate
    
            RS2.Update
        End If
        Set RS2 = Nothing
        rs.MoveNext
    Loop
    Set rs = Nothing
End Sub

Public Function getTripControl() As BWGJobDataTimeTable
Dim tmpS As New BWGSearchManager
    If tmpDataCar Is Nothing Then Set tmpDataCar = tmpS.JobDataTimeTableQuickSearch("JobDataCarID='" & JobDataCarID_ & "' And TripNo=" & TripNo_)
    If tmpDataCar.Count > 0 Then
        Set getTripControl = tmpDataCar(1)
    End If
End Function

Public Function getTripCount() As Collection
Dim tmpS As New BWGSearchManager
    If colTripControl Is Nothing Then Set colTripControl = tmpS.JobDataTimeTableQuickSearch("JobDataCarID='" & JobDataCarID_ & "'")
    If colTripControl.Count > 0 Then
        Set getTripCount = colTripControl
    End If
End Function

Public Property Get getMenifestHeader() As Collection
Dim tmpS As New BWGSearchManager
    If colMenifestHeader Is Nothing Then Set colMenifestHeader = tmpS.MenifestSearch("TimeTableNo = '" & TimeTableNo_ & "'")
    Set getMenifestHeader = colMenifestHeader
End Property

