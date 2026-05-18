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
Dim customerID_ As String
Dim customerName_ As String
Dim companyID_ As String
Dim wasteNo_ As String
Dim wasteName_ As String
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
Dim factoryRegisNo_ As String
Dim licenseStartDate_ As String
Dim licenseExpDate_ As String
Dim isLabResultNeeded_ As String
Dim industryTypeID_ As String
Dim industryTypeName_ As String
Dim industryTypeIDNew_ As String
Dim TypeDesc_ As String
Dim isAcceptedByEnvi_ As String
Dim acceptedDate_ As String
Dim treatmentType_ As String
Dim contractObj_ As BWGContract
Dim colKorOrNo_ As Collection
Dim staffCount_ As Integer
Dim staffCharge_ As Double
Dim carTypeID_ As String
Dim carTypeName_ As String
Dim isTrailer_ As String
Dim quotationID_ As String
Dim quotationNo_ As String
Dim transFee_ As Double
Dim transUnitID_ As String
Dim transUnitName_ As String
Dim treatmentFee_ As Double
Dim minWeightPerCar_ As Double
Dim treatmentUnitID_ As String
Dim treatmentUnitName_ As String
Dim menifestID_ As String
Dim isPriceIncTrans_ As String
Dim nonMinRate_ As Double
Dim hazMinRate_ As Double
Dim nblMinRate_ As Double
Dim hblMinRate_ As Double
Dim isChanged_ As Boolean
Dim isNew_ As String
Dim isRequest_ As String
Dim RequestDate_ As String
Dim isReject_ As String
Dim RejectDate_ As String
Dim RejectDesc_ As String
Dim isSample_ As String
Dim Status_  As String
Dim HeatValue_ As Double
Dim isQuotationReq_ As String
Dim EnviAcceptedDateAuto_ As String
Dim SaleSendDateAuto_ As String
Dim QuoPrice_  As String
Dim LicenseWeigh_ As Double
Dim LicenseWeighBWG_ As Double
Dim WasteCaseID_ As String
Dim isTruckTen_ As String
Dim isPickup_ As String
Dim isLugger_ As String
Dim isRollOff_ As String
Dim isHuerta_ As String
Dim isTrailerW_ As String
Dim isTankcar_ As String
Dim isFactoryCar_ As String
Dim isOrtherCar_ As String
Dim CarOrther_ As String

Dim FactoryCar_ As String
Dim AcceptStatusDate_ As String
Dim SDate_ As String
Dim EDate_ As String
Dim RefME_ As String
Dim Digit14_ As String
Dim CanEvap_ As String
Dim EvapDesc_ As String
Dim Precipitation_ As String
Dim PrecipitationDesc_ As String
Dim Neutralization_ As String
Dim NeutralizationDesc_ As String
Dim CommentWaste_ As String
Dim Remark_ As String
Dim RemarkDesc_ As String
Dim CommentWasteDesc_ As String
Dim colLab_ As Collection

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

Public Property Get SendLabdet() As Collection
Dim tmpS As New BWGSearchManager
    If colLab_ Is Nothing Then Set colLab_ = tmpS.SendSampleDetailSearch("WasteDataID='" & id_ & "'")
    Set tmpS = Nothing
    Set SendLabdet = colLab_
End Property
Public Property Set SendLabdet(newval As Collection)
    Set colLab_ = newval
End Property

Public Property Get FactoryCar() As String
    FactoryCar = FactoryCar_
End Property
Public Property Let FactoryCar(ByVal newval As String)
FactoryCar_ = Trim(newval)
End Property


Public Property Get CarOrther() As String
    CarOrther = CarOrther_
End Property
Public Property Let CarOrther(ByVal newval As String)
CarOrther_ = Trim(newval)
End Property

Public Property Get isOrtherCar() As Boolean
    If Trim(isOrtherCar_) = "Y" Then
        isOrtherCar = True
    Else
        isOrtherCar = False
    End If
End Property
Public Property Let isOrtherCar(ByVal newval As Boolean)
Dim tmpStr As String
    If newval Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(isOrtherCar_) <> Trim(tmpStr) Then
        isOrtherCar_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property
Public Property Get isFactoryCar() As Boolean
    If Trim(isFactoryCar_) = "Y" Then
        isFactoryCar = True
    Else
        isFactoryCar = False
    End If
End Property
Public Property Let isFactoryCar(ByVal newval As Boolean)
Dim tmpStr As String
    If newval Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(isFactoryCar_) <> Trim(tmpStr) Then
        isFactoryCar_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property
Public Property Get isTankcar() As Boolean
    If Trim(isTankcar_) = "Y" Then
        isTankcar = True
    Else
        isTankcar = False
    End If
End Property
Public Property Let isTankcar(ByVal newval As Boolean)
Dim tmpStr As String
    If newval Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(isTankcar_) <> Trim(tmpStr) Then
        isTankcar_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property
Public Property Get isTrailerW() As Boolean
    If Trim(isTrailerW_) = "Y" Then
        isTrailerW = True
    Else
        isTrailerW = False
    End If
End Property
Public Property Let isTrailerW(ByVal newval As Boolean)
Dim tmpStr As String
    If newval Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(isTrailerW_) <> Trim(tmpStr) Then
        isTrailerW_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property
Public Property Get isHuerta() As Boolean
    If Trim(isHuerta_) = "Y" Then
        isHuerta = True
    Else
        isHuerta = False
    End If
End Property
Public Property Let isHuerta(ByVal newval As Boolean)
Dim tmpStr As String
    If newval Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(isHuerta_) <> Trim(tmpStr) Then
        isHuerta_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property
Public Property Get isRollOff() As Boolean
    If Trim(isRollOff_) = "Y" Then
        isRollOff = True
    Else
        isRollOff = False
    End If
End Property
Public Property Let isRollOff(ByVal newval As Boolean)
Dim tmpStr As String
    If newval Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(isRollOff_) <> Trim(tmpStr) Then
        isRollOff_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property
Public Property Get isLugger() As Boolean
    If Trim(isLugger_) = "Y" Then
        isLugger = True
    Else
        isLugger = False
    End If
End Property
Public Property Let isLugger(ByVal newval As Boolean)
Dim tmpStr As String
    If newval Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(isLugger_) <> Trim(tmpStr) Then
        isLugger_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property
Public Property Get isPickup() As Boolean
    If Trim(isPickup_) = "Y" Then
        isPickup = True
    Else
        isPickup = False
    End If
End Property
Public Property Let isPickup(ByVal newval As Boolean)
Dim tmpStr As String
    If newval Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(isPickup_) <> Trim(tmpStr) Then
        isPickup_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property
Public Property Get isTruckTen() As Boolean
    If Trim(isTruckTen_) = "Y" Then
        isTruckTen = True
    Else
        isTruckTen = False
    End If
End Property
Public Property Let isTruckTen(ByVal newval As Boolean)
Dim tmpStr As String
    If newval Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(isTruckTen_) <> Trim(tmpStr) Then
        isTruckTen_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property
Public Property Get ID() As String
    ID = id_
End Property
Public Property Let ID(ByVal newval As String)
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
    CustomerID = customerID_
End Property
Public Property Let CustomerID(ByVal newval As String)
    If Trim(customerID_) <> Trim(newval) Then
        customerID_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get CustomerName() As String
    CustomerName = customerName_
End Property
Public Property Let CustomerName(ByVal newval As String)
    If Trim(customerName_) <> Trim(newval) Then
        customerName_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get Contract() As BWGContract
    If contractObj_ Is Nothing Then
        Set contractObj_ = New BWGContract
        contractObj_.ID = GetGUID
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
Public Property Get KorDate() As Collection
Dim tmpS As New BWGSearchManager
    If colKorOrNo_ Is Nothing Then Set colKorOrNo_ = tmpS.KorOrNoSearch("WasteDataID='" & id_ & "'")
    Set tmpS = Nothing
    Set KorDate = colKorOrNo_
End Property
Public Property Set KorDate(newval As Collection)
    Set colKorOrNo_ = newval
    isChanged_ = True
End Property
Public Sub RequeryKorOrNumber()
    Set colKorOrNo_ = Nothing
End Sub

Public Property Get companyID() As String
    companyID = companyID_
End Property
Public Property Let companyID(ByVal newval As String)
    If Trim(companyID_) <> Trim(newval) Then
        companyID_ = Trim(newval)
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
    QuotationID = quotationID_
End Property
Public Property Let QuotationID(ByVal newval As String)
    If Trim(quotationID_) <> Trim(newval) Then
        quotationID_ = Trim(newval)
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
    WasteName = wasteName_
End Property
Public Property Let WasteName(ByVal newval As String)
    If Trim(wasteName_) <> Trim(newval) Then
        wasteName_ = Trim(newval)
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
Public Property Get Digit14() As String
    Digit14 = Digit14_
End Property
Public Property Let Digit14(ByVal newval As String)
    If Trim(Digit14_) <> Trim(newval) Then
        Digit14_ = Trim(newval)
        isChanged_ = True
    End If
End Property

'Public Property Get CanEvap() As String
'    CanEvap = CanEvap_
'End Property
'Public Property Let CanEvap(ByVal newval As String)
'    If Trim(CanEvap_) <> Trim(newval) Then
'        CanEvap_ = Trim(newval)
'        isChanged_ = True
'    End If
'End Property

Public Property Get FactoryRegisNo() As String
    FactoryRegisNo = factoryRegisNo_
End Property
Public Property Let FactoryRegisNo(ByVal newval As String)
    If Trim(factoryRegisNo_) <> Trim(newval) Then
        factoryRegisNo_ = Trim(newval)
        isChanged_ = True
    End If
End Property
 Public Property Get LicenseStartDate() As String
    LicenseStartDate = licenseStartDate_
 End Property
 Public Property Let LicenseStartDate(ByVal newval As String)
     If Trim(licenseStartDate_) <> Trim(newval) Then
        licenseStartDate_ = Trim(newval)
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

Public Property Get IndustryTypeIDNew() As String
    IndustryTypeIDNew = industryTypeIDNew_
End Property
Public Property Let IndustryTypeIDNew(ByVal newval As String)
    If Trim(industryTypeIDNew_) <> Trim(newval) Then
        industryTypeIDNew_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get TypeDesc() As String
    TypeDesc = TypeDesc_
End Property
Public Property Let TypeDesc(ByVal newval As String)
    If Trim(TypeDesc_) <> Trim(newval) Then
        TypeDesc_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get QuotationNo() As String
    QuotationNo = quotationNo_
End Property
Public Property Let QuotationNo(ByVal newval As String)
    If Trim(quotationNo_) <> Trim(newval) Then
        quotationNo_ = Trim(newval)
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

Public Property Get TreatmentUnitID() As String
    TreatmentUnitID = treatmentUnitID_
End Property
Public Property Let TreatmentUnitID(ByVal newval As String)
    If Trim(treatmentUnitID_) <> Trim(newval) Then
        treatmentUnitID_ = Trim(newval)
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
    menifestID = menifestID_
End Property
Public Property Let menifestID(ByVal newval As String)
    menifestID_ = Trim(newval)
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

Public Property Get acceptedDate() As String
    acceptedDate = acceptedDate_
End Property
Public Property Let acceptedDate(ByVal newval As String)
    If Trim(acceptedDate_) <> Trim(newval) Then
        acceptedDate_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get AcceptStatusDate() As String
    AcceptStatusDate = AcceptStatusDate_
End Property
Public Property Let AcceptStatusDate(ByVal newval As String)
    If Trim(AcceptStatusDate_) <> Trim(newval) Then
        AcceptStatusDate_ = Trim(newval)
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

'************* for AP
Public Property Get HeatValue() As Double
    HeatValue = HeatValue_
End Property
Public Property Let HeatValue(ByVal newval As Double)
    If HeatValue_ <> newval Then
        HeatValue_ = newval
        isChanged_ = True
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
    Status_ = Trim(newval)
    isChanged_ = True
End Property

Public Property Get EnviAcceptedDateAuto() As String
    EnviAcceptedDateAuto = EnviAcceptedDateAuto_
End Property

Public Property Let EnviAcceptedDateAuto(ByVal newval As String)
    EnviAcceptedDateAuto_ = Trim(newval)
End Property

Public Property Get SaleSendDateAuto() As String
    SaleSendDateAuto = SaleSendDateAuto_
End Property

Public Property Let SaleSendDateAuto(ByVal newval As String)
    SaleSendDateAuto_ = Trim(newval)
End Property

Public Property Get QuoPrice() As String
    QuoPrice = QuoPrice_
End Property
Public Property Let QuoPrice(ByVal newval As String)
    QuoPrice_ = Trim(newval)
End Property
Public Property Get LicenseWeigh() As Double
    LicenseWeigh = LicenseWeigh_
End Property
Public Property Let LicenseWeigh(ByVal newval As Double)
    LicenseWeigh_ = newval
End Property
Public Property Get LicenseWeighBWG() As Double
    LicenseWeighBWG = LicenseWeighBWG_
End Property
Public Property Let LicenseWeighBWG(ByVal newval As Double)
    LicenseWeighBWG_ = newval
End Property
Public Property Get WasteCaseID() As String
 WasteCaseID = WasteCaseID_
End Property
Public Property Let WasteCaseID(ByVal newval As String)
WasteCaseID_ = newval
End Property
Public Property Get SDate() As String
 SDate = SDate_
End Property
Public Property Let SDate(ByVal newval As String)
SDate_ = newval
End Property
Public Property Get EDate() As String
 EDate = EDate_
End Property
Public Property Let EDate(ByVal newval As String)
EDate_ = newval
End Property
Public Property Get RefME() As String
 RefME = RefME_
End Property
Public Property Let RefME(ByVal newval As String)
RefME_ = newval
End Property
Public Sub SaveAccdate()
On Error GoTo ErrD
Dim tmpRec As ADODB.Recordset
Dim tmpRecwn As ADODB.Recordset
Dim rs As ADODB.Recordset
Dim tSql As String
Dim x%
    Set tmpRec = New ADODB.Recordset
    Set tmpRecwn = New ADODB.Recordset
'    If isNew_ = "N" Then
        tmpRec.Open "Select * from tbWasteDataCR Where WasteDataID='" & id_ & "'", DBManager.GetDBConnection, adOpenDynamic, adLockOptimistic
'    Else
'        tmpRec.Open "Select * from tbWasteData_New Where WasteDataID='" & id_ & "'", DBManager.GetDBConnection, adOpenDynamic, adLockOptimistic
'    End If
    If tmpRec.EOF Then
        Exit Sub
    End If
    If IsDate(AcceptStatusDate_) = True Then
        tmpRec("StatusAcceptDate") = Trim(AcceptStatusDate_)
    End If
    tmpRec.Update
    tmpRec.Requery
    tmpRec.Close
    Set tmpRec = Nothing
    isChanged_ = False
    Exit Sub
ErrD:
    MsgBox Err.Description & vbCrLf & "กรุณาลองใหม่หรือติดต่อผู้พัฒนาระบบ", vbCritical, "Error"
Exit Sub
End Sub

Public Sub Save()
On Error GoTo ErrD
Dim tmpRec As ADODB.Recordset
Dim tmpRecwn As ADODB.Recordset
Dim rs As ADODB.Recordset
Dim tSql As String
Dim x%
    Set tmpRec = New ADODB.Recordset
    Set tmpRecwn = New ADODB.Recordset
'    If isNew_ = "N" Then
        tmpRec.Open "Select * from tbWasteDataCR Where WasteDataID='" & id_ & "'", DBManager.GetDBConnection, adOpenDynamic, adLockOptimistic
'    Else
'        tmpRec.Open "Select * from tbWasteData_New Where WasteDataID='" & id_ & "'", DBManager.GetDBConnection, adOpenDynamic, adLockOptimistic
'    End If
    If tmpRec.EOF Then
        tmpRec.AddNew
        tmpRec("WasteDataID") = id_
'        If isNew_ = "N" Then
'            tSql = "Select * from tbWasteData_New Where WasteDataID='" & id_ & "'"
'            Set RS = New ADODB.Recordset
'            RS.Open tSql, DBManager.GetDBConnection, adOpenForwardOnly, adLockReadOnly
'            If RS.RecordCount > 0 Then
'                tmpRec("CreatedDate") = RS!CreatedDate
'            Else
'                tmpRec("CreatedDate") = TodayDate
'            End If
'            Set RS = Nothing
'        Else
'
'        End If
        tmpRec("CreatedDate") = TodayDate
        tmpRec("WasteYears") = Trim(WasteYears_)
    End If
    tmpRec("CustomerID") = Trim(customerID_)
    tmpRec("CompanyID") = Trim(companyID_)
    tmpRec("WasteNo") = Trim(wasteNo_)
    tmpRec("WasteName") = Trim(wasteName_)
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
    If Trim(RefME_) <> "" Then tmpRec("RefME") = Trim(RefME_)
    tmpRec("IsCancel") = isCancel_
    tmpRec("WasteCreateNo") = Trim(wasteCreateNo_)
    tmpRec("FactoryRegisNo") = Trim(factoryRegisNo_)
    tmpRec("14Digit") = Trim(Digit14_)
    tmpRec("CanEvap") = Trim(CanEvap_)
    tmpRec("EvapDesc") = Trim(EvapDesc_)
    tmpRec("Precipitation") = Trim(Precipitation_)
    tmpRec("PrecipitationDesc") = Trim(PrecipitationDesc_)
    tmpRec("Neutralization") = Trim(Neutralization_)
    tmpRec("NeutralizationDesc") = Trim(NeutralizationDesc_)
    tmpRec("Remark") = Trim(Remark_)
    
    tmpRec("CommentWasteDesc") = Trim(CommentWasteDesc_)
    tmpRec("RemarkDesc") = Trim(RemarkDesc_)

    
    tmpRec("LicenseStartDate") = Trim(licenseStartDate_)
    tmpRec("LicenseExpDate") = Trim(licenseExpDate_)
    tmpRec("isLabResultNeeded") = Trim(isLabResultNeeded_)
    tmpRec("IndustryTypeID") = Trim(industryTypeID_)
    If industryTypeIDNew_ <> "" Then
    tmpRec("IndustryTypeIDNew") = Trim(industryTypeIDNew_)
    End If
    tmpRec("isAcceptedByEnvi") = Trim(isAcceptedByEnvi_)
    tmpRec("isRequest") = Trim(isRequest_)
    If IsDate(AcceptStatusDate_) = True Then
        tmpRec("StatusAcceptDate") = Trim(AcceptStatusDate_)
    End If
    If IsDate(acceptedDate_) = True Then
        tmpRec("AcceptedDate") = Trim(acceptedDate_)
    End If
    tmpRec("SaleSendDateAuto") = Trim(SaleSendDateAuto_)
    tmpRec("TreatmentType") = Trim(treatmentType_)
    tmpRec!isNew = isNew_
    tmpRec!isAcceptedByEnvi = isAcceptedByEnvi_
    tmpRec!isRequest = isRequest_
    tmpRec!isSample = isSample_
    tmpRec!isQuotationReq = isQuotationReq_
    tmpRec("Status") = Trim(Status_)
    tmpRec("QuoPrice") = Trim(QuoPrice_)
    tmpRec("LicenseWeigh") = Trim(LicenseWeigh_)
    tmpRec("LicenseWeighBWG") = Trim(LicenseWeighBWG_)
    tmpRec("WasteCaseID") = Trim(WasteCaseID_)
    tmpRec!isTruckTen = isTruckTen_
    tmpRec!isPickup = isPickup_
    tmpRec!isLugger = isLugger_
    tmpRec!isRollOff = isRollOff_
    tmpRec!isHuerta = isHuerta_
    tmpRec!isTrailerW = isTrailerW_
    tmpRec!isTankcar = isTankcar_
    tmpRec!isFactoryCar = isFactoryCar_
    tmpRec!isOrtherCar = isOrtherCar_
    tmpRec("CarOrther") = Trim(CarOrther_)
    tmpRec("CommentWaste") = Trim(CommentWaste_)
    tmpRec("FactoryCar") = Trim(FactoryCar_)
    If IsDate(SDate_) = True Then
        tmpRec("StartDate") = Trim(SDate_)
    End If
    If IsDate(EDate_) = True Then
        tmpRec("EndDate") = Trim(EDate_)
    End If

     If Not contractObj_ Is Nothing Then tmpRec("ContractID") = contractObj_.ID
    tmpRec.Update
    tmpRec.Requery
    wasteNo_ = Trim("" & tmpRec("WasteNo"))
    tmpRec.Close
    Set tmpRec = Nothing
    If Not colKorOrNo_ Is Nothing Then
        With DBManager.GetDBConnection
            .Execute "Delete tbKorOrNumber Where WasteDataID='" & id_ & "'"
            For x = 1 To colKorOrNo_.Count
                If Trim(colKorOrNo_(x).KorOrNumber) <> "" Then
                    .Execute "Insert into tbKorOrNumber (KorOrID,WasteDataID,KorOrNumber,KorDate) Values ('" & colKorOrNo_(x).ID & "','" & colKorOrNo_(x).wastedataID & "','" & colKorOrNo_(x).KorOrNumber & "','" & colKorOrNo_(x).KorDate & "')"
                End If
            Next
        End With
    End If
    isChanged_ = False
    Exit Sub
ErrD:
    MsgBox Err.Description & vbCrLf & "กรุณาลองใหม่หรือติดต่อผู้พัฒนาระบบ", vbCritical, "Error"
Exit Sub
End Sub

Public Sub Add()
On Error GoTo ErrD
Dim x%
Dim tmpRec As ADODB.Recordset
Set tmpRec = New ADODB.Recordset
Dim Waste_ID As String
Waste_ID = GetGUID
Wastelab_ID = Waste_ID
tmpRec.Open "Select * from tbWasteDataCR Where WasteDataID = '" & Waste_ID & "' ", DBManager.GetDBConnection, adOpenDynamic, adLockOptimistic
If tmpRec.EOF Then
    tmpRec.AddNew
    tmpRec("WasteDataID") = Waste_ID
    tmpRec("CreatedDate") = TodayDate
    tmpRec("WasteYears") = "2556"
    tmpRec("AcceptedDate") = ""
    tmpRec("StatusAcceptDate") = ""
    tmpRec("isRequest") = "N"
    tmpRec("RequestDate") = TodayDate
    tmpRec("Status") = "R"
    tmpRec("WasteNo") = "= New ="
    tmpRec("isNew") = "Y"
    tmpRec("isAcceptedByEnvi") = "N"
    tmpRec("AcceptedDate") = ""
    tmpRec("SaleSendDateAuto") = TodayDate
    tmpRec("EnviAcceptedDateAuto") = ""
    tmpRec("WasteType") = ""
End If
    tmpRec("WasteReferrent") = WasteReferent
    tmpRec("CustomerID") = Trim(customerID_)
    tmpRec("CompanyID") = Trim(companyID_)
    tmpRec("WasteName") = Trim(wasteName_)
    tmpRec("EUCode1ID") = Trim(euCode1ID_)
    tmpRec("EUCode2ID") = Trim(euCode2ID_)
    tmpRec("EUCode3ID") = Trim(euCode3ID_)
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
    tmpRec("WasteCreateNo") = Trim(wasteCreateNo_)
    tmpRec("FactoryRegisNo") = Trim(factoryRegisNo_)
    tmpRec("14Digit") = Trim(Digit14_)
    tmpRec("CanEvap") = Trim(CanEvap_)
    tmpRec("EvapDesc") = Trim(EvapDesc_)
    tmpRec("Precipitation") = Trim(Precipitation_)
    tmpRec("PrecipitationDesc") = Trim(PrecipitationDesc_)
    tmpRec("Neutralization") = Trim(Neutralization_)
    tmpRec("NeutralizationDesc") = Trim(NeutralizationDesc_)
    tmpRec("Remark") = Trim(Remark_)
    
    tmpRec("CommentWasteDesc") = Trim(CommentWasteDesc_)
    tmpRec("RemarkDesc") = Trim(RemarkDesc_)
    
    tmpRec("isLabResultNeeded") = Trim(isLabResultNeeded_)
    tmpRec("IndustryTypeID") = Trim(industryTypeID_)
    If industryTypeIDNew_ <> "" Then
        tmpRec("IndustryTypeIDNew") = Trim(industryTypeIDNew_)
    End If
    tmpRec("TreatmentType") = Trim(treatmentType_)
    tmpRec("LicenseStartDate") = Trim(licenseStartDate_)
    tmpRec("LicenseExpDate") = Trim(licenseExpDate_)
    tmpRec("isSample") = isSample_
    tmpRec("QuoPrice") = QuoPrice_
    tmpRec("LicenseWeigh") = Trim(LicenseWeigh_)
    tmpRec("LicenseWeighBWG") = Trim(LicenseWeighBWG_)
    tmpRec("WasteCaseID") = Trim(WasteCaseID_)
    tmpRec("isQuotationReq") = isQuotationReq_
    tmpRec("isTruckTen") = isTruckTen_
    tmpRec("isPickup") = isPickup_
    tmpRec("isLugger") = isLugger_
    tmpRec("isRollOff") = isRollOff_
    tmpRec("isHuerta") = isHuerta_
    tmpRec("isTrailerW") = isTrailerW_
    tmpRec("isTankcar") = isTankcar_
    tmpRec("isFactoryCar") = isFactoryCar_
    tmpRec("isOrtherCar") = isOrtherCar_
    tmpRec("CarOrther") = Trim(CarOrther_)
    tmpRec("CommentWaste") = Trim(CommentWaste_)
    tmpRec("FactoryCar") = Trim(FactoryCar_)
    If IsDate(SDate_) = True Then
        tmpRec("StartDate") = Trim(SDate_)
    End If
    If IsDate(EDate_) = True Then
        tmpRec("EndDate") = Trim(EDate_)
    End If
    If Trim(RefME_) <> "" Then tmpRec("RefME") = Trim(RefME_)

     If Not contractObj_ Is Nothing Then tmpRec("ContractID") = contractObj_.ID
    tmpRec.Update
    tmpRec.Requery
    tmpRec.Close
    Set tmpRec = Nothing
    If Not colKorOrNo_ Is Nothing Then
        With DBManager.GetDBConnection
            .Execute "Delete tbKorOrNumber Where WasteDataID='" & id_ & "'"
            For x = 1 To colKorOrNo_.Count
                If Trim(colKorOrNo_(x).KorOrNumber) <> "" Then
                    .Execute "Insert into tbKorOrNumber (KorOrID,WasteDataID,KorOrNumber,KorDate) Values ('" & colKorOrNo_(x).ID & "','" & colKorOrNo_(x).wastedataID & "','" & colKorOrNo_(x).KorOrNumber & "','" & colKorOrNo_(x).KorDate & "')"
                End If
            Next
        End With
    End If
    isChanged_ = False
    Exit Sub
ErrD:
    MsgBox Err.Description & vbCrLf & "กรุณาลองใหม่หรือติดต่อผู้พัฒนาระบบ", vbCritical, "Error"
Exit Sub
End Sub

Public Function Delete() As String
On Error GoTo delErr:
    With DBManager.GetDBConnection
        .Execute "Delete tbOldTreatment Where WasteDataLabID in (Select WasteDataLabID from tbWasteDataLAB Where WasteDataID='" & id_ & "')"
        .Execute "Delete tbTreatment Where WasteDataLabID in (Select WasteDataLabID from tbWasteDataLAB Where WasteDataID='" & id_ & "')"
        .Execute "Delete tbWasteDataLAB Where WasteDataID='" & id_ & "'"
        .Execute "Delete tbWasteDataCR Where WasteDataID='" & id_ & "'"
        .Execute "Delete tbRunWasteNo Where WasteDataID='" & id_ & "'"
        .Execute "Update tbWasteData_New SET  isAcceptedByEnvi = 'N', AcceptedDate = '" & TodayDate & "', isNew = 'Y'"
    End With
    Delete = ""
    Exit Function
delErr:
    Delete = Err.Description
End Function

Public Function DeleteNew(ByVal WasteID As String) As String
On Error GoTo delErr:
    With DBManager.GetDBConnection
        .Execute "Delete tbOldTreatment_New Where WasteDataLabID in (Select WasteDataLabID from tbWasteDataLAB Where WasteDataID='" & WasteID & "')"
        .Execute "Delete tbTreatment_New Where WasteDataLabID in (Select WasteDataLabID from tbWasteDataLAB Where WasteDataID='" & WasteID & "')"
        .Execute "Delete tbWasteDataLAB_New Where WasteDataID='" & WasteID & "'"
        .Execute "Delete tbWasteData_New Where WasteDataID='" & WasteID & "'"
        .Execute "Delete tbRunWasteNo Where WasteDataID='" & WasteID & "'"
    End With
    DeleteNew = ""
    Exit Function
delErr:
    DeleteNew = Err.Description
End Function
Public Sub SetDataFromRecordset(selRec As ADODB.Recordset)
    id_ = Trim("" & selRec("WasteDataID"))
    WasteYears_ = Trim("" & selRec("WasteYears"))
    WasteReferrent_ = Trim("" & selRec("WasteReferrent"))
    customerID_ = Trim("" & selRec("CustomerID"))
    customerName_ = Trim("" & selRec("CompanyName"))
    companyID_ = Trim("" & selRec("CompanyID"))
    wasteNo_ = Trim("" & selRec("WasteNo"))
    wasteName_ = Trim("" & selRec("WasteName"))
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
    industryTypeIDNew_ = Trim("" & selRec("IndustryTypeIDNew"))
    industryTypeName_ = Trim("" & selRec("IndustryTypeName"))
'    maxQty_ = selRec("MaxQty")
'    maxUnit_ = Trim("" & selRec("MaxUnit"))
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
    wasteCreateNo_ = Trim("" & selRec("WasteCreateNo"))
    factoryRegisNo_ = Trim("" & selRec("FactoryRegisNo"))
    Digit14_ = Trim("" & selRec("14Digit"))
    
    CanEvap_ = Trim("" & selRec("CanEvap"))
    EvapDesc_ = Trim("" & selRec("EvapDesc"))
    Precipitation_ = Trim("" & selRec("Precipitation"))
    PrecipitationDesc_ = Trim("" & selRec("PrecipitationDesc"))
    Neutralization_ = Trim("" & selRec("Neutralization"))
    NeutralizationDesc_ = Trim("" & selRec("NeutralizationDesc"))
    Remark_ = Trim("" & selRec("Remark"))
    CommentWasteDesc_ = Trim("" & selRec("CommentWasteDesc"))
    RemarkDesc_ = Trim("" & selRec("RemarkDesc"))
        
        
    licenseStartDate_ = IIf(IsNull(selRec!LicenseStartDate), "", selRec!LicenseStartDate)
    licenseExpDate_ = Trim("" & selRec("LicenseExpDate"))
    isAcceptedByEnvi_ = Trim("" & selRec("isAcceptedByEnvi"))
    acceptedDate_ = Trim("" & selRec("AcceptedDate"))
    treatmentType_ = Trim("" & selRec("TreatmentType"))
    isRequest_ = IIf(IsNull(selRec!isRequest), "N", selRec!isRequest)
    RequestDate_ = IIf(IsNull(selRec!RequestDate), "", selRec!RequestDate)
    EnviAcceptedDateAuto_ = IIf(IsNull(selRec!EnviAcceptedDateAuto), "", selRec!EnviAcceptedDateAuto)
    SaleSendDateAuto_ = IIf(IsNull(selRec!SaleSendDateAuto), "", selRec!SaleSendDateAuto)
    isNew_ = IIf(IsNull(selRec!isNew), "N", selRec!isNew)
    isNew_ = Trim("" & selRec("isNew"))
    QuoPrice_ = Trim("" & selRec("QuoPrice"))
    LicenseWeigh_ = selRec("LicenseWeigh")
    LicenseWeighBWG_ = selRec("LicenseWeighBWG")
    WasteCaseID_ = Trim("" & selRec("WasteCaseID"))
    isTruckTen_ = IIf(IsNull(selRec!isTruckTen), "N", selRec!isTruckTen)
    isPickup_ = IIf(IsNull(selRec!isPickup), "N", selRec!isPickup)
    isLugger_ = IIf(IsNull(selRec!isLugger), "N", selRec!isLugger)
    isRollOff_ = IIf(IsNull(selRec!isRollOff), "N", selRec!isRollOff)
    isHuerta_ = IIf(IsNull(selRec!isHuerta), "N", selRec!isHuerta)
    isTrailerW_ = IIf(IsNull(selRec!isTrailerW), "N", selRec!isTrailerW)
    isTankcar_ = IIf(IsNull(selRec!isTankcar), "N", selRec!isTankcar)
    isFactoryCar_ = IIf(IsNull(selRec!isFactoryCar), "N", selRec!isFactoryCar)
    isOrtherCar_ = IIf(IsNull(selRec!isOrtherCar), "N", selRec!isOrtherCar)
    
    CarOrther_ = Trim("" & selRec("CarOrther"))
    CommentWaste_ = Trim("" & selRec("CommentWaste"))
    FactoryCar_ = Trim("" & selRec("FactoryCar"))
    AcceptStatusDate_ = ("" & selRec("StatusAcceptDate"))
    SDate_ = ("" & selRec("StartDate"))
    EDate_ = ("" & selRec("EndDate"))
    If Mid(IIf(IsNull(selRec!RejectDetail), "N", selRec!RejectDetail), 1, 1) = "Y" Then
        isReject_ = "Y"
    Else
        isReject_ = "N"
    End If
    If IsNull(selRec!RejectDetail) = False Then
        If Mid(selRec!RejectDetail, 1, 1) = "Y" Or Mid(selRec!RejectDetail, 1, 1) = "N" Then RejectDesc_ = Mid(selRec!RejectDetail, 2)
    Else
        RejectDesc_ = ""
    End If
    
    isSample_ = IIf(IsNull(selRec!isSample), "", selRec!isSample)
    
    If isExist(selRec, "StaffCountPerCar") Then
        If IsNull(selRec("StaffCountPerCar")) = False Then staffCount_ = selRec("StaffCountPerCar")
    End If
    If isExist(selRec, "ExtraLaborFee") Then
        If IsNull(selRec("ExtraLaborFee")) = False Then staffCharge_ = selRec("ExtraLaborFee")
    End If
    If isExist(selRec, "TruckTypeID") Then carTypeID_ = Trim("" & selRec("TruckTypeID"))
    If isExist(selRec, "TruckTypeDesc") Then carTypeName_ = Trim("" & selRec("TruckTypeDesc"))
    If isExist(selRec, "QuotationID") Then quotationID_ = Trim("" & selRec("QuotationID"))
    If isExist(selRec, "QuotationNo") Then quotationNo_ = Trim("" & selRec("QuotationNo"))
    If isExist(selRec, "isTrailer") Then isTrailer_ = Trim("" & selRec("isTrailer"))
    If isExist(selRec, "TransportFee") Then transFee_ = selRec("TransportFee")
    If isExist(selRec, "TrasnferUnitID") Then transUnitID_ = Trim("" & selRec("TrasnferUnitID"))
    If isExist(selRec, "TransUnitName") Then transUnitName_ = Trim("" & selRec("TransUnitName"))
    If isExist(selRec, "TreatmentRate") Then treatmentFee_ = selRec("TreatmentRate")
    If isExist(selRec, "MinWeightPerCar") Then minWeightPerCar_ = selRec("MinWeightPerCar")
    If isExist(selRec, "TreatmentUnitID") Then treatmentUnitID_ = Trim("" & selRec("TreatmentUnitID"))
    If isExist(selRec, "TreatmentUnitName") Then treatmentUnitName_ = Trim("" & selRec("TreatmentUnitName"))
    If isExist(selRec, "isPriceIncTransport") Then isPriceIncTrans_ = Trim("" & selRec("isPriceIncTransport"))

    If isExist(selRec, "NonMinRate") Then nonMinRate_ = selRec("NonMinRate")
    If isExist(selRec, "HazMinRate") Then hazMinRate_ = selRec("HazMinRate")
    If isExist(selRec, "NBLMinRate") Then nblMinRate_ = selRec("NBLMinRate")
    If isExist(selRec, "HBLMinRate") Then hblMinRate_ = selRec("HBLMinRate")
    
    If isExist(selRec, "StartDate") And isExist(selRec, "EndDate") And isExist(selRec, "DocNo") Then
        Set contractObj_ = New BWGContract
        With contractObj_
            .docNo = Trim("" & selRec("DocNo"))
            .EndDate = Trim("" & selRec("EndDate"))
            .ID = Trim("" & selRec("ContractID"))
            .RemDesc = Trim("" & selRec("RemDesc"))
            .isChanged = False
            .StartDate = Trim("" & selRec("StartDate"))
        End With
    End If
    
    isQuotationReq_ = IIf(IsNull(selRec!isQuotationReq), "N", selRec!isQuotationReq)
    Status_ = IIf(IsNull(selRec!Status), "D", selRec!Status)
    RefME_ = Trim("" & selRec("RefME"))
End Sub

Public Sub EnviAccepted()
On Error GoTo ErrD
Dim tSql As String
Dim rs As ADODB.Recordset
Dim WType As String
    DBManager.GetDBConnection.BeginTrans
    tSql = "SELECT * FROM tbWasteDataCR" & vbCrLf
    tSql = tSql & "WHERE WasteDataID = '" & id_ & "'"
    Set rs = New ADODB.Recordset
    Set rs = GetRS(tSql, adOpenForwardOnly, adLockReadOnly)
    If Not rs.EOF Then
            DBManager.GetDBConnection.Execute "Update tbWasteDataCR Set isNew='N', isRequest='Y', Status = 'P', RequestDate = '" & TodayDate & "',isAcceptedByEnvi='Y',AcceptedDate='" & TodayDate & "',EnviAcceptedDateAuto= '" & TodayDate & "'  Where WasteDataID='" & id_ & "'"
            isNew_ = "N"
            isRequest_ = "Y"
            Status_ = "P"
            RequestDate_ = TodayDate
            isAcceptedByEnvi_ = "Y"
            acceptedDate_ = TodayDate
            EnviAcceptedDateAuto_ = TodayDate
            WType = rs!WasteType
    Else
         tSql = "Update tbWasteDataCR Set isNew='N', isRequest='Y',RequestDate = '" & TodayDate & "', isAcceptedByEnvi='Y',AcceptedDate='" & TodayDate & "',EnviAcceptedDateAuto= '" & TodayDate & "',Status = 'P'  Where WasteDataID='" & id_ & "'"
        DBManager.GetDBConnection.Execute tSql
'         tSql = "Update tbWasteData_New Set isNew='N', isRequest='Y',RequestDate = '" & TodayDate & "', isAcceptedByEnvi='Y',AcceptedDate='" & TodayDate & "',EnviAcceptedDateAuto= '" & TodayDate & "',Status = 'P'  Where WasteDataID='" & id_ & "'"
'        DBManager.GetDBConnection.Execute tSql
            isNew_ = "N"
            isRequest_ = "Y"
            Status_ = "P"
            RequestDate_ = TodayDate
            isAcceptedByEnvi_ = "Y"
            acceptedDate_ = TodayDate
            EnviAcceptedDateAuto_ = TodayDate
    End If
    Set rs = Nothing
'    tSql = "Update tbWasteData_New Set isNew='N', Status = 'P',isRequest='Y',RequestDate = '" & TodayDate & "', isAcceptedByEnvi='Y',AcceptedDate='" & TodayDate & "',EnviAcceptedDateAuto='" & TodayDate & "'  Where WasteDataID='" & id_ & "'"
'    DBManager.GetDBConnection.Execute tSql
CreateWaste:
    Set rs = New ADODB.Recordset
    Set rs = DBManager.GetDBConnection.Execute("Exec sp_CreateWasteCode '" & id_ & "'")
    If rs.RecordCount = 0 Then
        GoTo CreateWaste
    Else
      wasteNo_ = rs(0).Value
    End If
    If Trim(WType) = "DOC" Then
        tSql = "Update tbWasteDataCR Set Status = 'A', isAcceptedByEnvi='Y',isRequest='Y',isNew='N',isTreatAble='Y',RequestDate = '" & TodayDate & "', AcceptedDate='" & TodayDate & "',EnviAcceptedDateAuto='" & TodayDate & "' ,WasteNo = '" & wasteNo_ & "' Where WasteDataID='" & id_ & "'"
        DBManager.GetDBConnection.Execute tSql
    Else
        tSql = "Update tbWasteDataCR Set Status = 'P', isRequest='Y',isNew='N',RequestDate = '" & TodayDate & "', AcceptedDate='" & TodayDate & "',EnviAcceptedDateAuto='" & TodayDate & "' ,WasteNo = '" & wasteNo_ & "' Where WasteDataID='" & id_ & "'"
        DBManager.GetDBConnection.Execute tSql
    End If
'    tSql = "SELECT WasteNo FROM vw_WasteDataCRSearch "
'    tSql = tSql & "WHERE WasteDataID = '" & id_ & "'"
'    Set RS = New ADODB.Recordset
'    Set RS.O = GetRS(tSql, adOpenForwardOnly, adLockReadOnly)
'    If Not RS.EOF Then wasteNo_ = RS!WasteNo
'    Dim tmpRec As ADODB.Recordset
'    Set tmpRec = New ADODB.Recordset
'    Dim sqlStr As String
'    sqlStr = "SELECT WasteNo FROM vw_WasteDataCRSearch  WHERE WasteDataID = '" & id_ & "' "
'    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
'    If tmpRec.EOF = False Then
'    wasteNo_ = tmpRec(0)
'    End If
'
    isNew_ = "N"
    isRequest_ = "Y"
    Status_ = "P"
    RequestDate_ = TodayDate
    isAcceptedByEnvi_ = "Y"
    acceptedDate_ = TodayDate
    EnviAcceptedDateAuto_ = TodayDate
    DBManager.GetDBConnection.CommitTrans
    Exit Sub
ErrD:    DBManager.GetDBConnection.RollbackTrans
    MsgBox "การตรวจรับยังไม่เสร็จสิ้น ข้อมูลมีปัญหา กรุณาปิดโปรแกรมแล้วตรวจสอบข้อมูลอีกครั้ง"
    Exit Sub
End Sub

Public Sub Request()
Dim tSql As String
Dim rs As ADODB.Recordset
    tSql = "SELECT * FROM tbWasteDataCR" & vbCrLf
    tSql = tSql & "WHERE WasteDataID = '" & id_ & "'"
    Set rs = New ADODB.Recordset
    Set rs = GetRS(tSql, adOpenForwardOnly, adLockReadOnly)
    If Not rs.EOF Then
        DBManager.GetDBConnection.Execute "Update tbWasteDataCR Set isNew = 'N', isRequest = 'N', RequestDate='" & TodayDate & "',SaleSendDateAuto= '" & TodayDate & "'   Where WasteDataID='" & id_ & "'"
        DBManager.GetDBConnection.Execute "Update tbWasteData_New Set isNew = 'N', isRequest = 'N', RequestDate='" & TodayDate & "',SaleSendDateAuto= '" & TodayDate & "'  Where WasteDataID='" & id_ & "'"
        If Status_ = "W" Then
           DBManager.GetDBConnection.Execute "Update tbWasteDataCR Set Status ='R', RequestDate='" & TodayDate & "' ,SaleSendDateAuto= '" & TodayDate & "'  Where WasteDataID='" & id_ & "'"
           DBManager.GetDBConnection.Execute "Update tbWasteData_New Set Status ='R', RequestDate='" & TodayDate & "',SaleSendDateAuto= '" & TodayDate & "'  Where WasteDataID='" & id_ & "'"
        Else
            DBManager.GetDBConnection.Execute "Update tbWasteDataCR Set Status ='R', RequestDate='" & TodayDate & "',SaleSendDateAuto= '" & TodayDate & "'  Where WasteDataID='" & id_ & "'"
            DBManager.GetDBConnection.Execute "Update tbWasteData_New Set Status ='R', RequestDate='" & TodayDate & "',SaleSendDateAuto= '" & TodayDate & "'  Where WasteDataID='" & id_ & "'"
        End If
      Else
        tSql = "SELECT * FROM tbWasteData_New" & vbCrLf
        tSql = tSql & "WHERE WasteDataID = '" & id_ & "'"
        Set rs = New ADODB.Recordset
        Set rs = GetRS(tSql, adOpenForwardOnly, adLockReadOnly)
        If Not rs.EOF Then
'            Dim dataDate As String
'            dataDate = SaleSendDateAuto_
'            If dataDate = "" Then
'                dataDate = TodayDate
'            Else
'                dataDate = SaleSendDateAuto_
'            End If
            DBManager.GetDBConnection.Execute "Update tbWasteData_New Set isNew = 'Y', isRequest = 'Y', RequestDate='" & TodayDate & "', isAcceptedByEnvi = '', AcceptedDate = '', RejectDetail = '',SaleSendDateAuto= '" & TodayDate & "'   Where WasteDataID='" & id_ & "'"
            DBManager.GetDBConnection.Execute "Update tbWasteData_New Set Status='R' Where WasteDataID='" & id_ & "'"
            isNew_ = "Y"
            isRequest_ = "Y"
            RequestDate_ = TodayDate
            isAcceptedByEnvi_ = ""
            acceptedDate_ = ""
            RejectDesc_ = ""
            SaleSendDateAuto_ = TodayDate
            Status_ = "R"
        End If
    End If
    isRequest_ = "Y"
    RequestDate_ = TodayDate
End Sub

Public Sub Reject(Optional ByVal RejectDetail As String = "ไม่ระบุเหตุผล")
Dim tSql As String
Dim rs As ADODB.Recordset
    tSql = "SELECT * FROM tbWasteDataCr" & vbCrLf
    tSql = tSql & "WHERE WasteDataID = '" & id_ & "'"
    Set rs = New ADODB.Recordset
    Set rs = GetRS(tSql, adOpenForwardOnly, adLockReadOnly)
    If Not rs.EOF Then
        DBManager.GetDBConnection.Execute "Update tbWasteDataCr Set isNew = 'Y', isRequest = 'N', RequestDate= '" & TodayDate & "', isAcceptedByEnvi = 'N', AcceptedDate = '" & TodayDate & "', RejectDetail = 'Y" & Trim(RejectDetail) & "'  Where WasteDataID='" & id_ & "'"
        DBManager.GetDBConnection.Execute "Update tbWasteDataCr Set Status = 'J', AcceptedDate = '', RejectDetail = 'Y" & Trim(RejectDetail) & "'  Where WasteDataID='" & id_ & "'"
    End If
    tSql = "SELECT * FROM tbWasteDataCR" & vbCrLf
    tSql = tSql & "WHERE WasteDataID = '" & id_ & "'"
       Set rs = New ADODB.Recordset
    Set rs = GetRS(tSql, adOpenForwardOnly, adLockReadOnly)
    If Not rs.EOF Then
        DBManager.GetDBConnection.Execute "Update tbWasteDataCR Set Status = 'J', AcceptedDate = ''  Where WasteDataID='" & id_ & "'"
    End If
    isReject_ = "Y"
    RejectDate_ = TodayDate
End Sub
