VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BWGWasteDataCR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim id_ As String
Dim WasteReferrent_ As String
Dim WasteYears_ As String
Dim CustomerID_ As String
Dim CustomerName_ As String
Dim companyID_ As String
Dim CompanyCode_ As String
Dim wasteNo_ As String
Dim WasteName_ As String
Dim euCode1ID_ As String
Dim euCode2ID_ As String
Dim euCode3ID_ As String
Dim euCode1_ As String
Dim euCode2_ As String
Dim euCode3_ As String
Dim wasteType_ As String
Dim mouD_ As String
Dim koR_ As String
Dim normalQty_ As Double
Dim normalUnit_ As String
Dim maxQty_ As Double
Dim maxUnit_ As String
Dim wasteRate_ As Double
Dim wasteRateUnit_ As String
Dim wasteRateTime_ As String
Dim isApproveByLAB_ As String
Dim marketingStaffID_ As String
Dim marketingStaffName_ As String
Dim enviStaffID_ As String
Dim enviStaffName_ As String
Dim isTreatAble_ As String
Dim isCancel_ As String
Dim wasteCreateNo_ As String
Dim FactoryRegisNo_ As String
Dim licenseExpDate_ As String
Dim licenseStartDate_ As String
Dim isLabResultNeeded_ As String
Dim industryTypeID_ As String
Dim industryTypeName_ As String
Dim isAcceptedByEnvi_ As String
Dim acceptedDate_ As String
Dim treatmentType_ As String
Dim contractObj_ As BWGContract
Dim colKorOrNo_ As Collection
Dim Approveddate_ As String
Dim staffCount_ As Integer
Dim staffCharge_ As Double
Dim carTypeID_ As String
Dim carTypeName_ As String
Dim isTrailer_ As String
Dim QuotationID_ As String
Dim QuotationNo_ As String
Dim BMEQuotation_ As String
Dim transFee_ As Double
Dim transUnitID_ As String
Dim transUnitName_ As String
Dim treatmentFee_ As Double
Dim minWeightPerCar_ As Double
Dim TreatmentUnitID_ As String
Dim treatmentUnitName_ As String
Dim MenifestID_ As String
Dim isPriceIncTrans_ As String
Dim IsDiscountPriceTrans_ As String
Dim nonMinRate_ As Double
Dim hazMinRate_ As Double
Dim nblMinRate_ As Double
Dim hblMinRate_ As Double
Dim MinRate_ As Double
Dim isChanged_ As Boolean

Dim AnalyseDate_  As String
Dim HeatValue_ As Double
Dim isSample_ As String

Dim isQuotationReq_ As String
Dim Status_ As String
Dim isJob_ As String
Dim TreatmentNo_ As String

Dim isNew_ As String
Dim isRequest_ As String
Dim RequestDate_ As String
Dim isReject_ As String
Dim RejectDate_ As String
Dim RejectDesc_ As String

Dim MaxWeightPerCar_ As Double
Dim MaxRate_ As Double
Dim MaxRateType_ As String

Dim tsdfID_ As String
Dim TSDFNo_ As String
Dim QuotationPrice_ As String
Dim RefME_ As String

Dim CanEvap_ As String
Dim EvapDesc_ As String
Dim Precipitation_ As String
Dim PrecipitationDesc_ As String
Dim Neutralization_ As String
Dim NeutralizationDesc_ As String

Dim Remark_ As String
Dim RemarkDesc_ As String
Dim CommentWaste_ As String
Dim CommentWasteDesc_ As String

Public Property Get CommentWaste() As Boolean
    If Trim(CommentWaste_) = "Y" Then
        CommentWaste = True
    Else
        CommentWaste = False
    End If
End Property
Public Property Let CommentWaste(ByVal newval As Boolean)
Dim tmpStr As String
    If newval Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(CommentWaste_) <> Trim(tmpStr) Then
        CommentWaste_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property

Public Property Get Remark() As Boolean
    If Trim(Remark_) = "Y" Then
        Remark = True
    Else
        Remark = False
    End If
End Property
Public Property Let Remark(ByVal newval As Boolean)
Dim tmpStr As String
    If newval Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(Remark_) <> Trim(tmpStr) Then
        Remark_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property

Public Property Get CommentWasteDesc() As String
    CommentWasteDesc = CommentWasteDesc_
End Property
Public Property Let CommentWasteDesc(ByVal newval As String)
    CommentWasteDesc_ = Trim(newval)
End Property
Public Property Get RemarkDesc() As String
    RemarkDesc = RemarkDesc_
End Property
Public Property Let RemarkDesc(ByVal newval As String)
    RemarkDesc_ = Trim(newval)
End Property


Public Property Get EvapDesc() As String
    EvapDesc = EvapDesc_
End Property
Public Property Let EvapDesc(ByVal newval As String)
    EvapDesc_ = Trim(newval)
End Property

Public Property Get PrecipitationDesc() As String
    PrecipitationDesc = PrecipitationDesc_
End Property
Public Property Let PrecipitationDesc(ByVal newval As String)
    PrecipitationDesc_ = Trim(newval)
End Property

Public Property Get NeutralizationDesc() As String
    NeutralizationDesc = NeutralizationDesc_
End Property
Public Property Let NeutralizationDesc(ByVal newval As String)
    NeutralizationDesc_ = Trim(newval)
End Property


Public Property Get CanEvap() As Boolean
    If Trim(CanEvap_) = "Y" Then
        CanEvap = True
    Else
        CanEvap = False
    End If
End Property
Public Property Let CanEvap(ByVal newval As Boolean)
Dim tmpStr As String
    If newval Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(CanEvap_) <> Trim(tmpStr) Then
        CanEvap_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property

Public Property Get Neutralization() As Boolean
    If Trim(Neutralization_) = "Y" Then
        Neutralization = True
    Else
        Neutralization = False
    End If
End Property
Public Property Let Neutralization(ByVal newval As Boolean)
Dim tmpStr As String
    If newval Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(Neutralization_) <> Trim(tmpStr) Then
        Neutralization_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property

Public Property Get Precipitation() As Boolean
    If Trim(Precipitation_) = "Y" Then
        Precipitation = True
    Else
        Precipitation = False
    End If
End Property
Public Property Let Precipitation(ByVal newval As Boolean)
Dim tmpStr As String
    If newval Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(Precipitation_) <> Trim(tmpStr) Then
        Precipitation_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property


Public Property Get id() As String
    id = id_
End Property
Public Property Let id(ByVal newval As String)
    id_ = Trim(newval)
End Property

Public Property Get WasteReferrent() As String
    WasteReferrent = WasteReferrent_
End Property
Public Property Let WasteReferrent(ByVal newval As String)
    WasteReferrent_ = Trim(newval)
End Property

Public Property Get WasteYears() As String
    WasteYears = WasteYears_
End Property
Public Property Let WasteYears(ByVal newval As String)
WasteYears_ = Trim(newval)
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

Public Property Get CustomerName() As String
    CustomerName = CustomerName_
End Property
Public Property Let CustomerName(ByVal newval As String)
    If Trim(CustomerName_) <> Trim(newval) Then
        CustomerName_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get Contract() As BWGContract
    If contractObj_ Is Nothing Then
        Set contractObj_ = New BWGContract
        contractObj_.id = GetGUID
    End If
    Set Contract = contractObj_
End Property
Public Property Set Contract(newval As BWGContract)
    Set contractObj_ = newval
    isChanged_ = True
End Property

Public Property Get KorOrNumbers() As Collection
Dim tmpS As New BWGSearchManager
    If colKorOrNo_ Is Nothing Then Set colKorOrNo_ = tmpS.KorOrNoSearch("WasteDataID='" & id_ & "'")
    Set tmpS = Nothing
    Set KorOrNumbers = colKorOrNo_
End Property
Public Property Set KorOrNumbers(newval As Collection)
    Set colKorOrNo_ = newval
    isChanged_ = True
End Property
Public Sub RequeryKorOrNumber()
    Set colKorOrNo_ = Nothing
End Sub

Public Property Get CompanyID() As String
    CompanyID = companyID_
End Property
Public Property Let CompanyID(ByVal newval As String)
    If Trim(companyID_) <> Trim(newval) Then
        companyID_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get CompanyCode() As String
    CompanyCode = CompanyCode_
End Property
Public Property Let CompanyCode(ByVal newval As String)
    If Trim(CompanyCode_) <> Trim(newval) Then
        CompanyCode_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get StaffCount() As Integer
    StaffCount = staffCount_
End Property
Public Property Let StaffCount(ByVal newval As Integer)
    If staffCount_ <> newval Then
        staffCount_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get StaffCharge() As Double
    StaffCharge = staffCharge_
End Property
Public Property Let StaffCharge(ByVal newval As Double)
    If staffCharge_ <> newval Then
        staffCharge_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get CarTypeID() As String
    CarTypeID = carTypeID_
End Property
Public Property Let CarTypeID(ByVal newval As String)
    If Trim(carTypeID_) <> Trim(newval) Then
        carTypeID_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get CarTypeName() As String
    CarTypeName = carTypeName_
End Property
Public Property Let CarTypeName(ByVal newval As String)
    If Trim(carTypeName_) <> Trim(newval) Then
        carTypeName_ = Trim(newval)
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
    If Trim(isTrailer_) <> Trim(tmpStr) Then
        isTrailer_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property

Public Property Get QuotationID() As String
    QuotationID = QuotationID_
End Property
Public Property Let QuotationID(ByVal newval As String)
    If Trim(QuotationID_) <> Trim(newval) Then
        QuotationID_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get WasteNo() As String
    WasteNo = wasteNo_
End Property
Public Property Let WasteNo(ByVal newval As String)
    If Trim(wasteNo_) <> Trim(newval) Then
        wasteNo_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get WasteName() As String
    WasteName = WasteName_
End Property
Public Property Let WasteName(ByVal newval As String)
    If Trim(WasteName_) <> Trim(newval) Then
        WasteName_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get EUCode1ID() As String
    EUCode1ID = euCode1ID_
End Property
Public Property Let EUCode1ID(ByVal newval As String)
    If Trim(euCode1ID_) <> Trim(newval) Then
        euCode1ID_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get EUCode2ID() As String
    EUCode2ID = euCode2ID_
End Property
Public Property Let EUCode2ID(ByVal newval As String)
    If Trim(euCode2ID_) <> Trim(newval) Then
        euCode2ID_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get EUCode3ID() As String
    EUCode3ID = euCode3ID_
End Property
Public Property Let EUCode3ID(ByVal newval As String)
    If Trim(euCode3ID_) <> Trim(newval) Then
        euCode3ID_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get EUCode1() As String
    EUCode1 = euCode1_
End Property
Public Property Let EUCode1(ByVal newval As String)
    If Trim(euCode1_) <> Trim(newval) Then
        euCode1_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get EUCode2() As String
    EUCode2 = euCode2_
End Property
Public Property Let EUCode2(ByVal newval As String)
    If Trim(euCode2_) <> Trim(newval) Then
        euCode2_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get EUCode3() As String
    EUCode3 = euCode3_
End Property
Public Property Let EUCode3(ByVal newval As String)
    If Trim(euCode3_) <> Trim(newval) Then
        euCode3_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get WasteType() As String
    WasteType = wasteType_
End Property
Public Property Let WasteType(ByVal newval As String)
    If Trim(wasteType_) <> Trim(newval) Then
        wasteType_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get TSDFID() As String
    TSDFID = tsdfID_
End Property
Public Property Let TSDFID(ByVal newval As String)
    If Trim(tsdfID_) <> Trim(newval) Then
        tsdfID_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get TSDFNo() As String
    TSDFNo = TSDFNo_
End Property
Public Property Let TSDFNo(ByVal newval As String)
    If Trim(TSDFNo_) <> Trim(newval) Then
        TSDFNo_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get Moud() As String
    Moud = mouD_
End Property
Public Property Let Moud(ByVal newval As String)
    If Trim(mouD_) <> Trim(newval) Then
        mouD_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get Kor() As String
    Kor = koR_
End Property
Public Property Let Kor(ByVal newval As String)
    If Trim(koR_) <> Trim(newval) Then
        koR_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get NormalQty() As Double
    NormalQty = normalQty_
End Property
Public Property Let NormalQty(ByVal newval As Double)
    If normalQty_ <> newval Then
        normalQty_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get NormalUnit() As String
    NormalUnit = normalUnit_
End Property
Public Property Let NormalUnit(ByVal newval As String)
    If Trim(normalUnit_) <> Trim(newval) Then
        normalUnit_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get MaxQty() As Double
    MaxQty = maxQty_
End Property
Public Property Let MaxQty(ByVal newval As Double)
    If maxQty_ <> newval Then
        maxQty_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get MaxUnit() As String
    MaxUnit = maxUnit_
End Property
Public Property Let MaxUnit(ByVal newval As String)
    If Trim(maxUnit_) <> Trim(newval) Then
        maxUnit_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get WasteRate() As Double
    WasteRate = wasteRate_
End Property
Public Property Let WasteRate(ByVal newval As Double)
    If wasteRate_ <> newval Then
        wasteRate_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get WasteRateUnit() As String
    WasteRateUnit = wasteRateUnit_
End Property
Public Property Let WasteRateUnit(ByVal newval As String)
    If Trim(wasteRateUnit_) <> Trim(newval) Then
        wasteRateUnit_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get WasteRateTime() As String
    WasteRateTime = wasteRateTime_
End Property
Public Property Let WasteRateTime(ByVal newval As String)
    If Trim(wasteRateTime_) <> Trim(newval) Then
        wasteRateTime_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get WasteApproveStatus() As String
    WasteApproveStatus = isApproveByLAB_
End Property
Public Property Let WasteApproveStatus(ByVal newval As String)
    If Trim(isApproveByLAB_) <> Trim(newval) Then
        isApproveByLAB_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get MarketingStaffID() As String
    MarketingStaffID = marketingStaffID_
End Property
Public Property Let MarketingStaffID(ByVal newval As String)
    If Trim(marketingStaffID_) <> Trim(newval) Then
        marketingStaffID_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get MarketingStaffName() As String
    MarketingStaffName = marketingStaffName_
End Property
Public Property Let MarketingStaffName(ByVal newval As String)
    If Trim(marketingStaffName_) <> Trim(newval) Then
        marketingStaffName_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get EnviStaffID() As String
    EnviStaffID = enviStaffID_
End Property
Public Property Let EnviStaffID(ByVal newval As String)
    If Trim(enviStaffID_) <> Trim(newval) Then
        enviStaffID_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get EnviStaffName() As String
    EnviStaffName = enviStaffName_
End Property
Public Property Let EnviStaffName(ByVal newval As String)
    If Trim(enviStaffName_) <> Trim(newval) Then
        enviStaffName_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get isTreatAble() As Boolean
    If isTreatAble_ = "Y" Then
        isTreatAble = True
    Else
        isTreatAble = False
    End If
End Property
Public Property Let isTreatAble(ByVal newval As Boolean)
Dim tmpStr As String
    If newval Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(isTreatAble_) <> Trim(tmpStr) Then
        isTreatAble_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property

Public Property Get isCancel() As Boolean
    If isCancel_ = "Y" Then
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
    If Trim(isCancel_) <> Trim(tmpStr) Then
        isCancel_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property

Public Property Get WasteCreateNo() As String
    WasteCreateNo = wasteCreateNo_
End Property
Public Property Let WasteCreateNo(ByVal newval As String)
    If Trim(wasteCreateNo_) <> Trim(newval) Then
        wasteCreateNo_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get FactoryRegisNo() As String
    FactoryRegisNo = FactoryRegisNo_
End Property
Public Property Let FactoryRegisNo(ByVal newval As String)
    If Trim(FactoryRegisNo_) <> Trim(newval) Then
        FactoryRegisNo_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get LicenseExpDate() As String
    LicenseExpDate = licenseExpDate_
End Property
Public Property Let LicenseExpDate(ByVal newval As String)
    If Trim(licenseExpDate_) <> Trim(newval) Then
        licenseExpDate_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get licenseStartDate() As String
    licenseStartDate = licenseStartDate_
End Property
Public Property Let licenseStartDate(ByVal newval As String)
    If Trim(licenseStartDate_) <> Trim(newval) Then
        licenseStartDate_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get isLabResultNeeded() As Boolean
    If Trim(isLabResultNeeded_) = "Y" Then
        isLabResultNeeded = True
    Else
        isLabResultNeeded = False
    End If
End Property
Public Property Let isLabResultNeeded(ByVal newval As Boolean)
Dim tmpStr As String
    If newval Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(isLabResultNeeded_) <> Trim(tmpStr) Then
        isLabResultNeeded_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property

Public Property Get IndustryTypeID() As String
    IndustryTypeID = industryTypeID_
End Property
Public Property Let IndustryTypeID(ByVal newval As String)
    If Trim(industryTypeID_) <> Trim(newval) Then
        industryTypeID_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get IndustryTypeName() As String
    IndustryTypeName = industryTypeName_
End Property
Public Property Let IndustryTypeName(ByVal newval As String)
    If Trim(industryTypeName_) <> Trim(newval) Then
        industryTypeName_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get QuotationNo() As String
    QuotationNo = QuotationNo_
End Property
Public Property Let QuotationNo(ByVal newval As String)
    If Trim(QuotationNo_) <> Trim(newval) Then
        QuotationNo_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get BMEQuotation() As String
    BMEQuotation = BMEQuotation_
End Property
Public Property Let BMEQuotation(ByVal newval As String)
    If Trim(BMEQuotation_) <> Trim(newval) Then
        BMEQuotation_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get TransportFee() As Double
    TransportFee = transFee_
End Property
Public Property Let TransportFee(ByVal newval As Double)
    If transFee_ <> newval Then
        transFee_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get TransUnitID() As String
    TransUnitID = transUnitID_
End Property
Public Property Let TransUnitID(ByVal newval As String)
    If Trim(transUnitID_) <> Trim(newval) Then
        transUnitID_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get TransUnitName() As String
    TransUnitName = transUnitName_
End Property
Public Property Let TransUnitName(ByVal newval As String)
    If Trim(transUnitName_) <> Trim(newval) Then
        transUnitName_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get TreatmentFee() As Double
    TreatmentFee = treatmentFee_
End Property
Public Property Let TreatmentFee(ByVal newval As Double)
    If treatmentFee_ <> newval Then
        treatmentFee_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get MinWeightPerCar() As Double
    MinWeightPerCar = minWeightPerCar_
End Property
Public Property Let MinWeightPerCar(ByVal newval As Double)
    If minWeightPerCar_ <> newval Then
        minWeightPerCar_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get MaxWeightPerCar() As Double
    MaxWeightPerCar = MaxWeightPerCar_
End Property
Public Property Let MaxWeightPerCar(ByVal newval As Double)
    If MaxWeightPerCar_ <> newval Then
        MaxWeightPerCar_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get MaxRate() As Double
    MaxRate = MaxRate_
End Property
Public Property Let MaxRate(ByVal newval As Double)
    If MaxRate_ <> newval Then
        MaxRate_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get MaxRateType() As String
    MaxRateType = MaxRateType_
End Property
Public Property Let MaxRateType(ByVal newval As String)
    If Trim(MaxRateType_) <> Trim(newval) Then
        MaxRateType_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get TreatmentUnitID() As String
    TreatmentUnitID = TreatmentUnitID_
End Property
Public Property Let TreatmentUnitID(ByVal newval As String)
    If Trim(TreatmentUnitID_) <> Trim(newval) Then
        TreatmentUnitID_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get TreatmentUnitName() As String
    TreatmentUnitName = treatmentUnitName_
End Property
Public Property Let TreatmentUnitName(ByVal newval As String)
    If Trim(treatmentUnitName_) <> Trim(newval) Then
        treatmentUnitName_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get menifestID() As String
    menifestID = MenifestID_
End Property
Public Property Let menifestID(ByVal newval As String)
    MenifestID_ = Trim(newval)
End Property

Public Property Get isPriceIncTrans() As Boolean
    If Trim(isPriceIncTrans_) = "Y" Then
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
    If Trim(isPriceIncTrans_) <> Trim(tmpStr) Then
        isPriceIncTrans_ = Trim(tmpStr)
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

Public Property Get NonMinRate() As Double
    NonMinRate = nonMinRate_
End Property
Public Property Let NonMinRate(ByVal newval As Double)
    If nonMinRate_ <> newval Then
        nonMinRate_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get HazMinRate() As Double
    HazMinRate = hazMinRate_
End Property
Public Property Let HazMinRate(ByVal newval As Double)
    If hazMinRate_ <> newval Then
        hazMinRate_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get NBLMinRate() As Double
    NBLMinRate = nblMinRate_
End Property
Public Property Let NBLMinRate(ByVal newval As Double)
    If nblMinRate_ <> newval Then
        nblMinRate_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get HBLMinRate() As Double
    HBLMinRate = hblMinRate_
End Property
Public Property Let HBLMinRate(ByVal newval As Double)
    If hblMinRate_ <> newval Then
        hblMinRate_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get MinRate() As Double
    MinRate = MinRate_
End Property
Public Property Let MinRate(ByVal newval As Double)
    If MinRate_ <> newval Then
        MinRate_ = newval
    End If
End Property

Public Property Get isAcceptedByEnvi() As Boolean
    If Trim(isAcceptedByEnvi_) = "Y" Then
        isAcceptedByEnvi = True
    Else
        isAcceptedByEnvi = False
    End If
End Property
Public Property Let isAcceptedByEnvi(ByVal newval As Boolean)
Dim tmpStr As String
    If newval Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(isAcceptedByEnvi_) <> Trim(tmpStr) Then
        isAcceptedByEnvi_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property

Public Property Get AcceptedDate() As String
    AcceptedDate = acceptedDate_
End Property
Public Property Let AcceptedDate(ByVal newval As String)
    If Trim(acceptedDate_) <> Trim(newval) Then
        acceptedDate_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get isReject() As Boolean
    If Trim(isReject_) = "Y" Then
        isReject = True
    Else
        isReject = False
    End If
End Property
Public Property Let isReject(ByVal newval As Boolean)
Dim tmpStr As String
    If newval Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(isReject_) <> Trim(tmpStr) Then
        isReject_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property

Public Property Get RejectDate() As String
    RejectDate = RejectDate_
End Property
Public Property Let RejectDate(ByVal newval As String)
    If Trim(RejectDate_) <> Trim(newval) Then
        RejectDate_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get RejectDesc() As String
    RejectDesc = RejectDesc_
End Property
Public Property Let RejectDesc(ByVal newval As String)
    If Trim(RejectDesc_) <> Trim(newval) Then
        RejectDesc_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get isRequest() As Boolean
    If Trim(isRequest_) = "Y" Then
        isRequest = True
    Else
        isRequest = False
    End If
End Property
Public Property Let isRequest(ByVal newval As Boolean)
Dim tmpStr As String
    If newval Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(isRequest_) <> Trim(tmpStr) Then
        isRequest_ = Trim(tmpStr)
    End If
End Property

Public Property Get RequestDate() As String
    RequestDate = RequestDate_
End Property
Public Property Let RequestDate(ByVal newval As String)
    If Trim(RequestDate_) <> Trim(newval) Then
        RequestDate_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get TreatmentType() As String
    TreatmentType = treatmentType_
End Property
Public Property Let TreatmentType(ByVal newval As String)
    If Trim(treatmentType_) <> Trim(newval) Then
        treatmentType_ = Trim(newval)
        isChanged_ = True
    End If
End Property

'For AP
Public Property Get HeatValue() As Double
    HeatValue = HeatValue_
End Property
Public Property Let HeatValue(ByVal newval As Double)
    If HeatValue_ <> newval Then
        HeatValue_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get AnalyseDate() As String
    AnalyseDate = AnalyseDate_
End Property
Public Property Let AnalyseDate(ByVal newval As String)
    If Trim(AnalyseDate_) <> Trim(newval) Then
        AnalyseDate_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get Approveddate() As String
    Approveddate = Approveddate_
End Property
Public Property Let Approveddate(ByVal newval As String)
    If Trim(Approveddate_) <> Trim(newval) Then
        Approveddate_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get isSample() As String
    isSample = isSample_
End Property
Public Property Let isSample(ByVal newval As String)
    If Trim(isSample_) <> Trim(newval) Then
        isSample_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get isChanged() As Boolean
    isChanged = isChanged_
End Property
Public Property Let isChanged(ByVal newval As Boolean)
    isChanged_ = newval
End Property

Public Property Get isNew() As Boolean
    If Trim(isNew_) = "Y" Then
        isNew = True
    Else
        isNew = False
    End If
End Property
Public Property Let isNew(ByVal newval As Boolean)
Dim tmpStr As String
    If newval Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(isNew_) <> Trim(tmpStr) Then
        isNew_ = Trim(tmpStr)
    End If
End Property

Public Property Get isQuotationReq() As Boolean
    If Trim(isQuotationReq_) = "Y" Then
        isQuotationReq = True
    Else
        isQuotationReq = False
    End If
End Property
Public Property Let isQuotationReq(ByVal newval As Boolean)
Dim tmpStr As String
    If newval Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(isQuotationReq_) <> Trim(tmpStr) Then
        isQuotationReq_ = Trim(tmpStr)
    End If
End Property

Public Property Get Status() As String
    Status = Status_
End Property
Public Property Let Status(ByVal newval As String)
    If Trim(Status_) <> Trim(newval) Then
        Status_ = Trim(newval)
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

Public Property Get TreatmentNo() As String
    TreatmentNo = TreatmentNo_
End Property
Public Property Let TreatmentNo(ByVal newval As String)
    TreatmentNo_ = Trim(newval)
End Property

Public Property Get QuotationPrice() As String
    QuotationPrice = QuotationPrice_
End Property
Public Property Let QuotationPrice(ByVal newval As String)
    QuotationPrice_ = Trim(newval)
End Property
Public Property Get RefME() As String
    RefME = RefME_
End Property
Public Property Let RefME(ByVal newval As String)
    RefME_ = Trim(newval)
End Property

Public Sub Save()
On Error GoTo ErrD
Dim tmpRec As ADODB.Recordset
Dim x%
Dim sqlStr As String
    sqlStr = "SELECT *, CONVERT(DATETIME, GETDATE(), 103) AS GETDATE FROM tbWasteDataCR WHERE WasteDataID='" & id_ & "'"
    Set tmpRec = New ADODB.Recordset
    GetDBConnection
    tmpRec.open sqlStr, DBConn, adOpenDynamic, adLockOptimistic
    If tmpRec.EOF Then
        tmpRec.AddNew
        tmpRec("WasteDataID") = id_
        tmpRec("CreatedDate") = TodayDate
        tmpRec("WasteYears") = Trim(WasteYears_)
    End If
    tmpRec("CustomerID") = Trim(CustomerID_)
    tmpRec("CompanyID") = Trim(companyID_)
    tmpRec("WasteNo") = Trim(wasteNo_)
    tmpRec("WasteName") = Trim(WasteName_)
    tmpRec("EUCode1ID") = Trim(euCode1ID_)
    tmpRec("EUCode2ID") = Trim(euCode2ID_)
    tmpRec("EUCode3ID") = Trim(euCode3ID_)
    tmpRec("WasteType") = Trim(wasteType_)
    tmpRec("Moud") = Trim(mouD_)
    tmpRec("Kor") = Trim(koR_)
    tmpRec("NormalQty") = Trim(normalQty_)
    tmpRec("NormalUnit") = Trim(normalUnit_)
    tmpRec("MaxQty") = Trim(maxQty_)
    tmpRec("MaxUnit") = Trim(maxUnit_)
    tmpRec("WasteRate") = Trim(wasteRate_)
    tmpRec("WasteRateUnit") = Trim(wasteRateUnit_)
    tmpRec("WasteRateTime") = Trim(wasteRateTime_)
    tmpRec("isApproveByLab") = Trim(isApproveByLAB_)
    tmpRec("MarketingStaffID") = Trim(marketingStaffID_)
    tmpRec("EnviStaffID") = Trim(enviStaffID_)
    tmpRec("isTreatAble") = Trim(isTreatAble_)
    If Trim(isCancel_) = "" Then isCancel_ = "N"
    tmpRec("IsCancel") = isCancel_
    tmpRec("WasteManageID") = Trim(wasteCreateNo_)
    tmpRec("FactoryRegisNo") = Trim(FactoryRegisNo_)
    tmpRec("LicenseExpDate") = Trim(licenseExpDate_)
    tmpRec("licenseStartDate") = Trim(licenseStartDate_)
    tmpRec("isLabResultNeeded") = Trim(isLabResultNeeded_)
    tmpRec("IndustryTypeID") = Trim(industryTypeID_)
    tmpRec("isAcceptedByEnvi") = Trim(isAcceptedByEnvi_)
    tmpRec("AcceptedDate") = Trim(acceptedDate_)
    tmpRec("isRequest") = Trim(isRequest_)
    tmpRec("TreatmentType") = Trim(treatmentType_)
    If Not contractObj_ Is Nothing Then tmpRec("ContractID") = contractObj_.id
    tmpRec!AnalyseDate = AnalyseDate_
    tmpRec!HeatingValue = HeatValue_
    tmpRec!isNew = isNew_
    tmpRec!isRequest = isRequest_
    tmpRec!isSample = isSample_
    tmpRec!isQuotationReq = isQuotationReq_
    tmpRec!Status = Trim(Status_)
    If Trim(RefME_) <> "" Then tmpRec("RefME") = Trim(RefME_)
    
    tmpRec!LastUser = Trim(Trim(CurrentUser.StaffFName) & " " & Trim(CurrentUser.StaffLName))
    tmpRec!LastUpdate = GetSystemDateTime.SystemDateTime
    
    tmpRec.Update
    tmpRec.Requery
    wasteNo_ = Trim("" & tmpRec("WasteNo"))
'    tmpRec.Close
    Set tmpRec = Nothing
    If Not colKorOrNo_ Is Nothing Then

        DBConnExc "Delete tbKorOrNumber WHERE WasteDataID='" & id_ & "'"
        For x = 1 To colKorOrNo_.Count
            If Trim(colKorOrNo_(x).KorOrNumber) <> "" Then
                DBConnExc "Insert into tbKorOrNumber (KorOrID,WasteDataID,KorOrNumber) Values ('" & colKorOrNo_(x).id & "','" & colKorOrNo_(x).wastedataID & "','" & colKorOrNo_(x).KorOrNumber & "')"
            End If
        Next
            
    End If
    isChanged_ = False
    CloseDBConnection
    Exit Sub
ErrD:
    GetMsg "", vbCritical
    Exit Sub
End Sub

Public Function Delete() As String
On Error GoTo delErr:
    DBConnExc "Delete tbOldTreatment WHERE WasteDataLabID in (Select WasteDataLabID from tbWasteDataLAB WHERE WasteDataID='" & id_ & "')"
    DBConnExc "Delete tbTreatment WHERE WasteDataLabID in (Select WasteDataLabID from tbWasteDataLAB WHERE WasteDataID='" & id_ & "')"
    DBConnExc "Delete tbWasteDataLAB WHERE WasteDataID='" & id_ & "'"
    DBConnExc "Delete tbWasteDataCR WHERE WasteDataID='" & id_ & "'"
    DBConnExc "Delete tbRunWasteNo WHERE WasteDataID='" & id_ & "'"
    DBConnExc "Update tbWasteData_New SET  isAcceptedByEnvi = 'N', AcceptedDate = '" & TodayDate & "', isNew = 'Y'"
    Delete = ""
    Exit Function
delErr:
    Delete = err.Description
End Function

Public Function DeleteNew(ByVal WasteID As String) As String
On Error GoTo delErr:
    DBConnExc "Delete tbOldTreatment_New WHERE WasteDataLabID in (Select WasteDataLabID from tbWasteDataLAB WHERE WasteDataID='" & WasteID & "')"
    DBConnExc "Delete tbTreatment_New WHERE WasteDataLabID in (Select WasteDataLabID from tbWasteDataLAB WHERE WasteDataID='" & WasteID & "')"
    DBConnExc "Delete tbWasteDataLAB_New WHERE WasteDataID='" & WasteID & "'"
    DBConnExc "Delete tbWasteData_New WHERE WasteDataID='" & WasteID & "'"
    DBConnExc "Delete tbRunWasteNo WHERE WasteDataID='" & WasteID & "'"
    DeleteNew = ""
    Exit Function
delErr:
    DeleteNew = err.Description
End Function

Public Sub SetDataFromRecordset(selRec As ADODB.Recordset)
    id_ = Trim("" & selRec("WasteDataID"))
    If isExist(selRec, "WasteYears") Then WasteYears_ = Trim("" & selRec!WasteYears & "")
    If isExist(selRec, "WasteReferrent") Then WasteReferrent_ = Trim("" & selRec!WasteReferrent)
    CustomerID_ = Trim("" & selRec("CustomerID"))
    CustomerName_ = Trim("" & selRec("CompanyName"))
    companyID_ = Trim("" & selRec("CompanyID"))
    CompanyCode_ = Trim("" & selRec("CompanyCode"))
    wasteNo_ = Trim("" & selRec("WasteNo"))
    WasteName_ = Trim("" & selRec("WasteName"))
    euCode1ID_ = Trim("" & selRec("EUCode1ID"))
    euCode2ID_ = Trim("" & selRec("EUCode2ID"))
    euCode3ID_ = Trim("" & selRec("EUCode3ID"))
    euCode1_ = Trim("" & selRec("EUCode1"))
    euCode2_ = Trim("" & selRec("EUCode2"))
    euCode3_ = Trim("" & selRec("EUCode3"))
    wasteType_ = Trim("" & selRec("WasteType"))
    mouD_ = Trim("" & selRec("Moud"))
    koR_ = Trim("" & selRec("Kor"))
    normalQty_ = selRec("NormalQty")
    normalUnit_ = Trim("" & selRec("NormalUnit"))
    isLabResultNeeded_ = Trim("" & selRec("isLabResultNeeded"))
    industryTypeID_ = Trim("" & selRec("IndustryTypeID"))
    industryTypeName_ = Trim("" & selRec("IndustryTypeName"))
    maxQty_ = selRec("MaxQty")
    maxUnit_ = Trim("" & selRec("MaxUnit"))
    wasteRate_ = selRec("WasteRate")
    wasteRateUnit_ = Trim("" & selRec("WasteRateUnit"))
    wasteRateTime_ = Trim("" & selRec("WasteRateTime"))
    isApproveByLAB_ = Trim("" & selRec("isApproveByLab"))
    marketingStaffID_ = Trim("" & selRec("MarketingStaffID"))
    marketingStaffName_ = Trim("" & selRec("CRStaffName"))
    enviStaffID_ = Trim("" & selRec("EnviStaffID"))
    enviStaffName_ = Trim("" & selRec("EnviStaffName"))
    isTreatAble_ = Trim("" & selRec("isTreatAble"))
    isCancel_ = Trim("" & selRec("isCancel"))
    wasteCreateNo_ = Trim("" & selRec("WasteManageID"))
    FactoryRegisNo_ = Trim("" & selRec("FactoryRegisNo"))
    licenseExpDate_ = Trim("" & selRec("LicenseExpDate"))
    licenseStartDate_ = Trim("" & selRec("licenseStartDate"))
    isAcceptedByEnvi_ = Trim("" & selRec("isAcceptedByEnvi"))
    acceptedDate_ = Trim("" & selRec("AcceptedDate"))
    treatmentType_ = Trim("" & selRec("TreatmentType"))
    CanEvap_ = Trim("" & selRec("CanEvap"))
    EvapDesc_ = Trim("" & selRec("EvapDesc"))
    Precipitation_ = Trim("" & selRec("Precipitation"))
    PrecipitationDesc_ = Trim("" & selRec("PrecipitationDesc"))
    Neutralization_ = Trim("" & selRec("Neutralization"))
    NeutralizationDesc_ = Trim("" & selRec("NeutralizationDesc"))
    
    CommentWasteDesc_ = Trim("" & selRec("CommentWasteDesc"))
        
    CommentWaste_ = Trim("" & selRec("CommentWaste"))
    RefME_ = Trim("" & selRec("RefME"))
    If isExist(selRec, "isSample") Then isSample_ = IIf(IsNull(selRec!isSample), "N", selRec!isSample)
    If isExist(selRec, "StaffCountPerCar") Then
        If IsNull(selRec("StaffCountPerCar")) = False Then staffCount_ = selRec("StaffCountPerCar")
    End If
    If isExist(selRec, "ExtraLaborFee") Then
        If IsNull(selRec("ExtraLaborFee")) = False Then staffCharge_ = selRec("ExtraLaborFee")
    End If
    If isExist(selRec, "TruckTypeID") Then carTypeID_ = Trim("" & selRec("TruckTypeID"))
    If isExist(selRec, "TruckTypeDesc") Then carTypeName_ = Trim("" & selRec("TruckTypeDesc"))
    If isExist(selRec, "QuotationID") Then QuotationID_ = Trim("" & selRec("QuotationID"))
    If isExist(selRec, "QuotationNo") Then QuotationNo_ = Trim("" & selRec("QuotationNo"))
    If isExist(selRec, "BMEQuotation") Then BMEQuotation_ = Trim("" & selRec("BMEQuotation"))
    If isExist(selRec, "isTrailer") Then isTrailer_ = Trim("" & selRec("isTrailer"))
    If isExist(selRec, "TransportFee") Then transFee_ = selRec("TransportFee")
    If isExist(selRec, "TransferUnitID") Then transUnitID_ = Trim("" & selRec("TransferUnitID"))
    If isExist(selRec, "TransUnitName") Then transUnitName_ = Trim("" & selRec("TransUnitName"))
    If isExist(selRec, "TreatmentRate") Then treatmentFee_ = selRec("TreatmentRate")
    If isExist(selRec, "MinWeightPerCar") Then minWeightPerCar_ = selRec("MinWeightPerCar")
    If isExist(selRec, "TreatmentUnitID") Then TreatmentUnitID_ = Trim("" & selRec("TreatmentUnitID"))
    If isExist(selRec, "TreatmentUnitName") Then treatmentUnitName_ = Trim("" & selRec("TreatmentUnitName"))
    If isExist(selRec, "isPriceIncTransport") Then isPriceIncTrans_ = Trim("" & selRec("isPriceIncTransport"))
    If isExist(selRec, "IsDiscountPriceTrans") Then IsDiscountPriceTrans_ = Trim("" & selRec("IsDiscountPriceTrans"))

    If isExist(selRec, "NonMinRate") Then nonMinRate_ = Trim("0" & selRec("NonMinRate"))
    If isExist(selRec, "HazMinRate") Then hazMinRate_ = Trim("0" & selRec("HazMinRate"))
    If isExist(selRec, "NBLMinRate") Then nblMinRate_ = Trim("0" & selRec("NBLMinRate"))
    If isExist(selRec, "HBLMinRate") Then hblMinRate_ = Trim("0" & selRec("HBLMinRate"))
    
    If isExist(selRec, "TSDFID") Then tsdfID_ = selRec("TSDFID")
    If isExist(selRec, "TSDFNo") Then TSDFNo_ = selRec("TSDFNo")
    
    If isExist(selRec, "StartDate") And isExist(selRec, "EndDate") And isExist(selRec, "DocNo") Then
        Set contractObj_ = New BWGContract
        With contractObj_
            .DocNo = Trim("" & selRec("DocNo"))
            .EndDate = Trim("" & selRec("EndDate"))
            .id = Trim("" & selRec("ContractID"))
            .Remdesc = Trim("" & selRec("RemDesc"))
            .isChanged = False
            .StartDate = Trim("" & selRec("StartDate"))
        End With
    End If
    
    If isExist(selRec, "AnalyseDate") Then AnalyseDate_ = IIf(IsNull(selRec!AnalyseDate), "", selRec!AnalyseDate)
    If isExist(selRec, "HeatingValue") Then HeatValue_ = IIf(IsNull(selRec!HeatingValue), 0, selRec!HeatingValue)
    If isExist(selRec, "isQuotationReq") Then
        isQuotationReq_ = IIf(IsNull(selRec!isQuotationReq), "N", selRec!isQuotationReq)
    Else
        isQuotationReq_ = "N"
    End If
    
    If isExist(selRec, "WasteStatus") Then Status_ = IIf(IsNull(selRec!WasteStatus), "D", selRec!WasteStatus)
    If isExist(selRec, "Status") Then Status_ = IIf(IsNull(selRec!Status), "D", Trim(selRec!Status))
    If isExist(selRec, "isJobs") Then isJob_ = IIf(IsNull(selRec!isJobs), "D", selRec!isJobs)
    If isExist(selRec, "TreatmentNo") Then TreatmentNo_ = IIf(IsNull(selRec!TreatmentNo), "", selRec!TreatmentNo)
    
    If isExist(selRec, "isRequest") Then isRequest_ = IIf(IsNull(selRec!isRequest), "N", selRec!isRequest)
    If isExist(selRec, "RequestDate") Then RequestDate_ = IIf(IsNull(selRec!RequestDate), "", selRec!RequestDate)
    If isExist(selRec, "isNew") Then isNew_ = IIf(IsNull(selRec!isNew), "N", selRec!isNew)
    If isExist(selRec, "QuoPrice") Then QuotationPrice_ = Trim("" & selRec("QuoPrice"))
    
    If isExist(selRec, "RejectDetail") Then
        If Mid(IIf(IsNull(selRec!RejectDetail), "N", selRec!RejectDetail), 1, 1) = "Y" Then
            isReject_ = "Y"
        Else
            isReject_ = "N"
        End If
    Else
        isReject_ = "N"
    End If
    
    If isExist(selRec, "RejectDetail") Then
        If IsNull(selRec!RejectDetail) = False Then
            If Mid(selRec!RejectDetail, 1, 1) = "Y" Or Mid(selRec!RejectDetail, 1, 1) = "N" Then RejectDesc_ = Mid(selRec!RejectDetail, 2)
        Else
            RejectDesc_ = ""
        End If
    End If

End Sub

Public Sub EnviAccepted()
    DBConnExc "Update tbWasteDataCR Set isAcceptedByEnvi='Y',AcceptedDate='" & TodayDate & "' WHERE WasteDataID='" & id_ & "'"
    isAcceptedByEnvi_ = "Y"
    acceptedDate_ = TodayDate
End Sub

Public Sub Request()
Dim tSql As String
Dim rs As ADODB.Recordset
    
    tSql = "SELECT * FROM tbWasteDataCR" & vbCrLf
    tSql = tSql & "WHERE WasteDataID = '" & id_ & "'"
    Set rs = New ADODB.Recordset
    Set rs = GetRS(tSql, adOpenForwardOnly, adLockReadOnly)
    If Not rs.EOF Then
        DBConnExc "Update tbWasteDataCR Set isNew = 'N', isRequest = 'Y', RequestDate='" & TodayDate & "' WHERE WasteDataID='" & id_ & "'"
        DBConnExc "Update tbWasteData_New Set isNew = 'N', isRequest = 'Y', RequestDate='" & TodayDate & "' WHERE WasteDataID='" & id_ & "'"
        If Status_ = "W" Then
           DBConnExc "Update tbWasteDataCR Set Status ='R', RequestDate='" & TodayDate & "' WHERE WasteDataID='" & id_ & "'"
           DBConnExc "Update tbWasteData_New Set Status ='R', RequestDate='" & TodayDate & "' WHERE WasteDataID='" & id_ & "'"
        Else
            DBConnExc "Update tbWasteDataCR Set Status ='R', RequestDate='" & TodayDate & "' WHERE WasteDataID='" & id_ & "'"
            DBConnExc "Update tbWasteData_New Set Status ='R', RequestDate='" & TodayDate & "' WHERE WasteDataID='" & id_ & "'"
        End If
      Else
        tSql = "SELECT * FROM tbWasteData_New" & vbCrLf
        tSql = tSql & "WHERE WasteDataID = '" & id_ & "'"
        Set rs = New ADODB.Recordset
        Set rs = GetRS(tSql, adOpenForwardOnly, adLockReadOnly)
        If Not rs.EOF Then
            DBConnExc "Update tbWasteData_New Set isNew = 'Y', isRequest = 'Y', RequestDate='" & TodayDate & "', isAcceptedByEnvi = '', AcceptedDate = '', RejectDetail = ''  WHERE WasteDataID='" & id_ & "'"
            DBConnExc "Update tbWasteData_New Set Status='R' WHERE WasteDataID='" & id_ & "'"
        End If
    End If
    isRequest_ = "Y"
    RequestDate_ = TodayDate
End Sub

Public Sub Reject(Optional ByVal RejectDetail As String = "ไม่ระบุเหตุผล")
Dim tSql As String
Dim rs As ADODB.Recordset
    
    tSql = "SELECT * FROM tbWasteData_New" & vbCrLf
    tSql = tSql & "WHERE WasteDataID = '" & id_ & "'"
    Set rs = New ADODB.Recordset
    Set rs = GetRS(tSql, adOpenForwardOnly, adLockReadOnly)
    If Not rs.EOF Then
        DBConnExc "Update tbWasteData_New Set isNew = 'Y', isRequest = 'N', RequestDate= '', isAcceptedByEnvi = 'N', AcceptedDate = '" & TodayDate & "', RejectDetail = 'Y" & Trim(RejectDetail) & "'  WHERE WasteDataID='" & id_ & "'"
        DBConnExc "Update tbWasteData_New Set Status = 'J', AcceptedDate = '" & TodayDate & "', RejectDetail = 'Y" & Trim(RejectDetail) & "'  WHERE WasteDataID='" & id_ & "'"
    End If
    tSql = "SELECT * FROM tbWasteDataCR" & vbCrLf
    tSql = tSql & "WHERE WasteDataID = '" & id_ & "'"
       Set rs = New ADODB.Recordset
    Set rs = GetRS(tSql, adOpenForwardOnly, adLockReadOnly)
    If Not rs.EOF Then
        DBConnExc "Update tbWasteDataCR Set Status = 'J', AcceptedDate = '" & TodayDate & "' WHERE WasteDataID='" & id_ & "'"
    End If
    isReject_ = "Y"
    RejectDate_ = TodayDate
End Sub
