VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BWGJobDataDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim id_ As String
Dim CustomerID_ As String
Dim jobID_ As String
Dim jobNo_ As String
Dim wasteDataID_ As String
Dim WasteName_ As String
Dim wasteNo_ As String
Dim wasteType_ As String
Dim wasteEUCode_ As String
Dim wasteQuotationNo_ As String
Dim WasteGenNo_ As String
Dim estWasteQty_ As Double
Dim wasteQtyUnitID_ As String
Dim wasteQtyUnitName_ As String
Dim containerTypeDesc_ As String
Dim containerCount_ As Double
Dim detRemDesc_ As String
Dim destinationDesc_ As String
Dim rowID_ As Integer
Dim carNo_ As Integer
Dim tsdfID_ As String
Dim tsdfName_ As String
Dim tsdfRegisNo_ As String
Dim tsdfTumbolID_ As String

Dim QuotationID_ As String
Dim laborCharge_ As Double
Dim truckTypeID_ As String
Dim truckTypeName_ As String
Dim isTrailer_ As String
Dim transportFee_ As Double
Dim minWeightPerCar_ As Double
Dim transUnitID_ As String
Dim transUnitName_ As String
Dim JobTreatmentRate_ As Double
Dim JobTreatmentUnitID_ As String
Dim JobTreatmentUnitName_ As String
Dim laborCount_ As Integer
Dim isPriceIncTrans_ As String
Dim IsDiscountPriceTrans_  As String

'''''This part for menifest detail
Dim DisposerWeight_ As Double
Dim treatmentCharge_ As Double
Dim MenifestID_ As String
Dim menifestActQty_ As Double
Dim menifestDocNo_ As String
Dim menifestDocNoRef_ As String
Dim menifestWorkDate_ As String
Dim menifestCarRegisID_ As String
'''''''''''''''''''''''
Dim ChargeWType_ As String
Dim TimeTableNo_ As String

''''This part for job fail
Dim failType_ As String
Dim failRemark_ As String
Dim failCheck_ As Integer
'''''''''''''''''''''''
Dim nonMinRate_ As Double
Dim hazMinRate_ As Double
Dim nblMinRate_ As Double
Dim hblMinRate_ As Double
Dim MinRate_ As Double
Dim isChanged_ As Boolean

Dim QuoMaxWeightPerCar_ As Double
Dim QuoMaxRate_ As Double
Dim QuoMaxRateType_ As String

Dim QuoTreatmentRate_ As Double
Dim QuoTreatmentUnitID_ As String
Dim HAType_ As String
Dim IsNewWaste_ As String
Dim IsNewWasteCR_ As String
Dim InWeight_ As Double
Dim TransRate_ As String
Dim QuoCost_ As Double
Dim QuoCostUnit_ As String


Dim CanEvap_ As String
Dim EvapDesc_ As String
Dim Precipitation_ As String
Dim PrecipitationDesc_ As String
Dim Neutralization_ As String
Dim NeutralizationDesc_ As String
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



Public Property Get CommentWasteDesc() As String
    CommentWasteDesc = CommentWasteDesc_
End Property
Public Property Let CommentWasteDesc(ByVal newval As String)
    CommentWasteDesc_ = Trim(newval)
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

Public Property Get CustomerID() As String
    CustomerID = CustomerID_
End Property
Public Property Let CustomerID(ByVal newval As String)
    If Trim(CustomerID_) <> Trim(newval) Then
        CustomerID_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get JobID() As String
    JobID = jobID_
End Property
Public Property Let JobID(ByVal newval As String)
    If Trim(jobID_) <> Trim(newval) Then
        jobID_ = Trim(newval)
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

Public Property Get wastedataID() As String
    wastedataID = wasteDataID_
End Property
Public Property Let wastedataID(ByVal newval As String)
    If Trim(wasteDataID_) <> Trim(newval) Then
        wasteDataID_ = Trim(newval)
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

Public Property Get WasteNo() As String
    WasteNo = wasteNo_
End Property
Public Property Let WasteNo(ByVal newval As String)
    If Trim(wasteNo_) <> Trim(newval) Then
        wasteNo_ = Trim(newval)
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

Public Property Get WasteEUCode() As String
    WasteEUCode = wasteEUCode_
End Property
Public Property Let WasteEUCode(ByVal newval As String)
    If Trim(wasteEUCode_) <> Trim(newval) Then
        wasteEUCode_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get WasteQuotationNo() As String
    WasteQuotationNo = wasteQuotationNo_
End Property
Public Property Let WasteQuotationNo(ByVal newval As String)
    If Trim(wasteQuotationNo_) <> Trim(newval) Then
        wasteQuotationNo_ = Trim(newval)
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

Public Property Get FailType() As String
    FailType = failType_
End Property
Public Property Let FailType(ByVal newval As String)
    If Trim(failType_) <> Trim(newval) Then
        failType_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get FailRemark() As String
    FailRemark = failRemark_
End Property
Public Property Let FailRemark(ByVal newval As String)
    If Trim(failRemark_) <> Trim(newval) Then
        failRemark_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get FailCheck() As Integer
    FailCheck = failCheck_
End Property
Public Property Let FailCheck(ByVal newval As Integer)
    If Trim(failCheck_) <> Trim(newval) Then
        failCheck_ = Trim(newval)
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
Public Property Get InWeight() As Double
    InWeight = InWeight_
End Property
Public Property Let InWeight(ByVal newval As Double)
    If InWeight_ <> newval Then
        InWeight_ = newval
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
        isChanged_ = True
    End If
End Property

Public Property Get EstWasteQty() As Double
    EstWasteQty = estWasteQty_
End Property
Public Property Let EstWasteQty(ByVal newval As Double)
    If estWasteQty_ <> newval Then
        estWasteQty_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get WasteQtyUnitID() As String
    WasteQtyUnitID = wasteQtyUnitID_
End Property
Public Property Let WasteQtyUnitID(ByVal newval As String)
    If Trim(wasteQtyUnitID_) <> Trim(newval) Then
        wasteQtyUnitID_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get HAType() As String
    HAType = HAType_
End Property
Public Property Let HAType(ByVal newval As String)
    If Trim(HAType_) <> Trim(newval) Then
        HAType_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get IsNewWasteCR() As String
    IsNewWasteCR = IsNewWasteCR_
End Property
Public Property Let IsNewWasteCR(ByVal newval As String)
    If Trim(IsNewWasteCR_) <> Trim(newval) Then
        IsNewWasteCR_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get IsNewWaste() As String
    IsNewWaste = IsNewWaste_
End Property
Public Property Let IsNewWaste(ByVal newval As String)
    If Trim(IsNewWaste_) <> Trim(newval) Then
        IsNewWaste_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get WasteQtyUnitName() As String
    WasteQtyUnitName = wasteQtyUnitName_
End Property
Public Property Let WasteQtyUnitName(ByVal newval As String)
    If Trim(wasteQtyUnitName_) <> Trim(newval) Then
        wasteQtyUnitName_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get ContainerTypeDesc() As String
    ContainerTypeDesc = containerTypeDesc_
End Property
Public Property Let ContainerTypeDesc(ByVal newval As String)
    If Trim(containerTypeDesc_) <> Trim(newval) Then
        containerTypeDesc_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get ContainerCount() As Double
    ContainerCount = containerCount_
End Property
Public Property Let ContainerCount(ByVal newval As Double)
    If containerCount_ <> newval Then
        containerCount_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get Remdesc() As String
    Remdesc = detRemDesc_
End Property
Public Property Let Remdesc(ByVal newval As String)
    If Trim(detRemDesc_) <> Trim(newval) Then
        detRemDesc_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get DestinationDesc() As String
    DestinationDesc = destinationDesc_
End Property
Public Property Let DestinationDesc(ByVal newval As String)
    If Trim(destinationDesc_) <> Trim(newval) Then
        destinationDesc_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get RowID() As Integer
    RowID = rowID_
End Property
Public Property Let RowID(ByVal newval As Integer)
    If rowID_ <> newval Then
        rowID_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get carNo() As Integer
    carNo = carNo_
End Property
Public Property Let carNo(ByVal newval As Integer)
    If carNo_ <> newval Then
        carNo_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get TSDFID() As String
    TSDFID = Trim(tsdfID_)
End Property
Public Property Let TSDFID(ByVal newval As String)
    If Trim(tsdfID_) <> Trim(newval) Then
        tsdfID_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get TSDFName() As String
    TSDFName = Trim(tsdfName_)
End Property
Public Property Let TSDFName(ByVal newval As String)
    If Trim(tsdfName_) <> Trim(newval) Then
        tsdfName_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get TSDFRegisNo() As String
    TSDFRegisNo = Trim(tsdfRegisNo_)
End Property
Public Property Let TSDFRegisNo(ByVal newval As String)
    If Trim(tsdfRegisNo_) <> Trim(newval) Then
        tsdfRegisNo_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get TSDFTumbolID() As String
    TSDFTumbolID = Trim(tsdfTumbolID_)
End Property
Public Property Let TSDFTumbolID(ByVal newval As String)
    If Trim(tsdfTumbolID_) <> Trim(newval) Then
        tsdfTumbolID_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get DisposerWeight() As Double
    DisposerWeight = DisposerWeight_
End Property
Public Property Let DisposerWeight(ByVal newval As Double)
    If DisposerWeight_ <> newval Then
        DisposerWeight_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get TreatmentCharge() As Double
    TreatmentCharge = treatmentCharge_
End Property
Public Property Let TreatmentCharge(ByVal newval As Double)
    If treatmentCharge_ <> newval Then
        treatmentCharge_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get menifestID() As String
    menifestID = MenifestID_
End Property
Public Property Let menifestID(ByVal newval As String)
    If Trim(MenifestID_) <> Trim(newval) Then
        MenifestID_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get MenifestActQty() As Double
    MenifestActQty = menifestActQty_
End Property
Public Property Let MenifestActQty(ByVal newval As Double)
    If menifestActQty_ <> newval Then
        menifestActQty_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get MenifestDocNoRef() As String
    MenifestDocNoRef = menifestDocNoRef_
End Property
Public Property Let MenifestDocNoRef(ByVal newval As String)
    If Trim(menifestDocNoRef_) <> Trim(newval) Then
        menifestDocNoRef_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get MenifestDocNo() As String
    MenifestDocNo = menifestDocNo_
End Property
Public Property Let MenifestDocNo(ByVal newval As String)
    If Trim(menifestDocNo_) <> Trim(newval) Then
        menifestDocNo_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get MenifestWorkDate() As String
    MenifestWorkDate = menifestWorkDate_
End Property
Public Property Let MenifestWorkDate(ByVal newval As String)
    If Trim(menifestWorkDate_) <> Trim(newval) Then
        menifestWorkDate_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get MenifestCarRegisID() As String
    MenifestCarRegisID = menifestCarRegisID_
End Property
Public Property Let MenifestCarRegisID(ByVal newval As String)
    If Trim(menifestCarRegisID_) <> Trim(newval) Then
        menifestCarRegisID_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get ChargeWeightType() As String
    ChargeWeightType = ChargeWType_
End Property
Public Property Let ChargeWeightType(ByVal newval As String)
    If Trim(ChargeWType_) <> Trim(newval) Then
        ChargeWType_ = Trim(newval)
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

Public Property Get QuotationID() As String
    QuotationID = QuotationID_
End Property
Public Property Let QuotationID(ByVal newval As String)
    If Trim(QuotationID_) <> Trim(newval) Then
        QuotationID_ = Trim(newval)
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

Public Property Get JobTreatmentRate() As Double
    JobTreatmentRate = JobTreatmentRate_
End Property
Public Property Let JobTreatmentRate(ByVal newval As Double)
    If JobTreatmentRate_ <> newval Then
        JobTreatmentRate_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get JobTreatmentUnitID() As String
    JobTreatmentUnitID = JobTreatmentUnitID_
End Property
Public Property Let JobTreatmentUnitID(ByVal newval As String)
    If Trim(JobTreatmentUnitID_) <> Trim(newval) Then
        JobTreatmentUnitID_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get QuoTreatmentRate() As Double
    QuoTreatmentRate = QuoTreatmentRate_
End Property
Public Property Let QuoTreatmentRate(ByVal newval As Double)
    If QuoTreatmentRate_ <> newval Then
        QuoTreatmentRate_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get QuoTreatmentUnitID() As String
    QuoTreatmentUnitID = QuoTreatmentUnitID_
End Property
Public Property Let QuoTreatmentUnitID(ByVal newval As String)
    If Trim(QuoTreatmentUnitID_) <> Trim(newval) Then
        QuoTreatmentUnitID_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get QuoCost() As Double
    QuoCost = QuoCost_
End Property
Public Property Let QuoCost(ByVal newval As Double)
    If QuoCost_ <> newval Then
        QuoCost_ = newval
        isChanged_ = True
    End If
End Property
Public Property Get QuoCostUnit() As String
    QuoCostUnit = QuoCostUnit_
End Property
Public Property Let QuoCostUnit(ByVal newval As String)
    If Trim(QuoCostUnit_) <> Trim(newval) Then
        QuoCostUnit_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get QuoTreatmentUnitName() As String
    QuoTreatmentUnitName = JobTreatmentUnitName_
End Property
Public Property Let QuoTreatmentUnitName(ByVal newval As String)
    If Trim(JobTreatmentUnitName_) <> Trim(newval) Then
        JobTreatmentUnitName_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get QuoTruckTypeID() As String
    QuoTruckTypeID = truckTypeID_
End Property
Public Property Let QuoTruckTypeID(ByVal newval As String)
    If Trim(truckTypeID_) <> Trim(newval) Then
        truckTypeID_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get QuoTruckTypeName() As String
    QuoTruckTypeName = truckTypeName_
End Property
Public Property Let QuoTruckTypeName(ByVal newval As String)
    If Trim(truckTypeName_) <> Trim(newval) Then
        truckTypeName_ = Trim(newval)
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

Public Property Get JobTransportFee() As Double
    JobTransportFee = transportFee_
End Property
Public Property Let JobTransportFee(ByVal newval As Double)
    If transportFee_ <> newval Then
        transportFee_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get JobTransUnitID() As String
    JobTransUnitID = transUnitID_
End Property
Public Property Let JobTransUnitID(ByVal newval As String)
    If Trim(transUnitID_) <> Trim(newval) Then
        transUnitID_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get QuoMinWeightPerCar() As Double
    QuoMinWeightPerCar = minWeightPerCar_
End Property
Public Property Let QuoMinWeightPerCar(ByVal newval As Double)
    If minWeightPerCar_ <> newval Then
        minWeightPerCar_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get QuoMaxWeightPerCar() As Double
    QuoMaxWeightPerCar = QuoMaxWeightPerCar_
End Property
Public Property Let QuoMaxWeightPerCar(ByVal newval As Double)
    If QuoMaxWeightPerCar_ <> newval Then
        QuoMaxWeightPerCar_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get QuoMaxRate() As Double
    QuoMaxRate = QuoMaxRate_
End Property
Public Property Let QuoMaxRate(ByVal newval As Double)
    If QuoMaxRate_ <> newval Then
        QuoMaxRate_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get QuoMaxRateType() As String
    QuoMaxRateType = QuoMaxRateType_
End Property
Public Property Let QuoMaxRateType(ByVal newval As String)
    If Trim(QuoMaxRateType_) <> Trim(newval) Then
        QuoMaxRateType_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get QuoTransportFee() As Double
    QuoTransportFee = transportFee_
End Property
Public Property Let QuoTransportFee(ByVal newval As Double)
    If transportFee_ <> newval Then
        transportFee_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get QuoTransUnitID() As String
    QuoTransUnitID = transUnitID_
End Property
Public Property Let QuoTransUnitID(ByVal newval As String)
    If Trim(transUnitID_) <> Trim(newval) Then
        transUnitID_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get QuoTransUnitName() As String
    QuoTransUnitName = transUnitName_
End Property
Public Property Let QuoTransUnitName(ByVal newval As String)
    If Trim(transUnitName_) <> Trim(newval) Then
        transUnitName_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get TransRate() As String
    TransRate = TransRate_
End Property
Public Property Let TransRate(ByVal newval As String)
    If Trim(TransRate_) <> Trim(newval) Then
        TransRate_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get QuoLaborCount() As Integer
    QuoLaborCount = laborCount_
End Property
Public Property Let QuoLaborCount(ByVal newval As Integer)
    If laborCount_ <> newval Then
        laborCount_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get QuoLaborCharge() As Integer
    QuoLaborCharge = laborCharge_
End Property
Public Property Let QuoLaborCharge(ByVal newval As Integer)
    If laborCharge_ <> newval Then
        laborCharge_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get isChanged() As Boolean
    isChanged = isChanged_
End Property
Public Property Let isChanged(ByVal newval As Boolean)
    isChanged_ = newval
End Property

Public Sub Save()
On Error GoTo ErrD
Dim tmpRec As ADODB.Recordset
Dim tmpCon As ADODB.Connection
Dim sqlStr As String
    sqlStr = "SELECT * FROM tbJobDataDetail WHERE JobDetID='" & id_ & "'"
    Set tmpRec = New ADODB.Recordset
    GetDBConnection
    tmpRec.open sqlStr, DBConn, adOpenDynamic, adLockOptimistic
    If tmpRec.EOF Then
        tmpRec.AddNew
        tmpRec("JobDetID") = id_
    End If
    tmpRec("JobID") = Trim(jobID_)
    tmpRec("WasteDataID") = Trim(wasteDataID_)
    tmpRec("EstWasteQty") = Trim(estWasteQty_)
    tmpRec("WasteQtyUnitID") = Trim(wasteQtyUnitID_)
    tmpRec("ContainerTypeDesc") = Trim(containerTypeDesc_)
    tmpRec("ContainerCount") = Trim(containerCount_)
    tmpRec("DetRemDesc") = Trim(detRemDesc_)
    tmpRec("DestinationDesc") = Trim(destinationDesc_)
    tmpRec("TSDFID") = Trim(tsdfID_)
    tmpRec("MarkForDelete") = "N"
    tmpRec("RowID") = rowID_
    tmpRec("CarNo") = carNo_
    
    tmpRec("CanEvap") = Trim(CanEvap_)
    tmpRec("EvapDesc") = Trim(EvapDesc_)
    tmpRec("Precipitation") = Trim(Precipitation_)
    tmpRec("PrecipitationDesc") = Trim(PrecipitationDesc_)
    tmpRec("Neutralization") = Trim(Neutralization_)
    tmpRec("NeutralizationDesc") = Trim(NeutralizationDesc_)
    tmpRec("CommentWaste") = Trim(CommentWaste_)
    tmpRec("CommentWasteDesc") = Trim(CommentWasteDesc_)
    
    tmpRec("QuotationID") = Trim(QuotationID_)
    tmpRec("ExtraLaborFee") = laborCharge_
    tmpRec("TruckTypeID") = Trim(truckTypeID_)
    tmpRec("TransportFee") = transportFee_
    tmpRec("MinWeightPerCar") = minWeightPerCar_
    tmpRec("TransferUnitID") = Trim(transUnitID_)
    tmpRec("TreatmentRate") = JobTreatmentRate_
    tmpRec("TreatmentUnitID") = Trim(JobTreatmentUnitID_)
    tmpRec("LaborCount") = laborCount_
    tmpRec("isPriceIncTrans") = Trim(isPriceIncTrans_)
    tmpRec("IsDiscountPriceTrans") = Trim(IsDiscountPriceTrans_)
    tmpRec("NonMinRate") = nonMinRate_
    tmpRec("HazMinRate") = hazMinRate_
    tmpRec("NBLMinRate") = nblMinRate_
    tmpRec("HBLMinRate") = hblMinRate_
    tmpRec("IsNewWaste") = IsNewWaste_
    tmpRec("IsNewWasteCR") = IsNewWasteCR_
    
    tmpRec.Update
'    tmpRec.Close
    
    Set tmpCon = Nothing
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
    DBConnExc "Delete tbJobDataDetail WHERE JobDetID='" & id_ & "'"
    Delete = ""
    Exit Function
delErr:
    Delete = err.Description
End Function

Public Sub SetDataFromRecordset(selRec As ADODB.Recordset)
    id_ = Trim("" & selRec("JobDetID"))
    CustomerID_ = Trim("" & selRec!CustomerID)
    jobID_ = Trim("" & selRec("JobID"))
    jobNo_ = Trim("" & selRec("JobNo"))
    wasteDataID_ = Trim("" & selRec("WasteDataID"))
    WasteName_ = Trim("" & selRec("WasteName"))
    wasteEUCode_ = Trim("" & selRec("EUCode"))
    wasteNo_ = Trim("" & selRec("WasteNo"))
    WasteGenNo_ = Trim("" & selRec("WasteCreateNo"))
    If IsNull(selRec("EstWasteQty")) = False Then estWasteQty_ = selRec("EstWasteQty")
    wasteQtyUnitID_ = Trim("" & selRec("WasteQtyUnitID"))
    wasteQtyUnitName_ = Trim("" & selRec("UnitName"))
    containerTypeDesc_ = Trim("" & selRec("ContainerTypeDesc"))
    If IsNull(selRec("ContainerCount")) = False Then containerCount_ = selRec("ContainerCount")
    detRemDesc_ = Trim("" & selRec("DetRemDesc"))
    destinationDesc_ = Trim("" & selRec("DestinationDesc"))
    wasteQuotationNo_ = Trim("" & selRec("QuotationNo"))
    wasteType_ = Trim("" & selRec("WasteType"))
    
    tsdfID_ = Trim("" & selRec("TSDFID"))
    tsdfName_ = Trim("" & selRec("TSDFCompanyName"))
    tsdfRegisNo_ = Trim("" & selRec("TSDFRegisNo"))
    tsdfTumbolID_ = Trim("" & selRec("TSDFRegisNo"))
    
    If IsNull(selRec("RowID")) = False Then rowID_ = selRec("RowID")
    If IsNull(selRec("CarNo")) = False Then carNo_ = selRec("CarNo")
    
    If isExist(selRec, "MenifestID") Then MenifestID_ = Trim("" & selRec("MenifestID"))
    If isExist(selRec, "ActualQty") Then menifestActQty_ = Trim("" & selRec("ActualQty"))
    If isExist(selRec, "ChargeWType") Then ChargeWType_ = Trim("" & selRec("ChargeWType"))
    If isExist(selRec, "DocNo") Then menifestDocNo_ = Trim("" & selRec("DocNo"))
    If isExist(selRec, "DocNoRef") Then menifestDocNo_ = Trim("" & selRec("DocNoRef"))
    If isExist(selRec, "WorkDate") Then menifestWorkDate_ = Trim("" & selRec("WorkDate"))
    If isExist(selRec, "CarRegisID") Then menifestCarRegisID_ = Trim("" & selRec("CarRegisID"))
    
    If isExist(selRec, "DisposerWeight") Then
        If IsNull(selRec("DisposerWeight")) = False Then DisposerWeight_ = selRec("DisposerWeight")
    End If
    If isExist(selRec, "TreatmentCharge") Then
        If IsNull(selRec("TreatmentCharge")) = False Then treatmentCharge_ = selRec("TreatmentCharge")
    End If
    
    If IsNull(selRec("TransportFee")) = False Then transportFee_ = selRec("TransportFee")
    If IsNull(selRec("MinWeightPerCar")) = False Then minWeightPerCar_ = selRec("MinWeightPerCar")
    
    If IsNull(selRec("TreatmentRate")) = False Then JobTreatmentRate_ = selRec("TreatmentRate")
    JobTreatmentUnitID_ = Trim("" & selRec("TreatmentUnitID"))
    JobTreatmentUnitName_ = Trim("" & selRec("TreatmentUnitID"))   'Trim("" & selRec("TreatmentUnitName"))
    
    If IsNull(selRec("QuoTreatmentRate")) = False Then QuoTreatmentRate_ = selRec("QuoTreatmentRate")
    If IsNull(selRec("QuoTreatmentUnitID")) = False Then QuoTreatmentUnitID_ = selRec("QuoTreatmentUnitID")
    
    truckTypeID_ = Trim("" & selRec("TruckTypeID"))
    truckTypeName_ = Trim("" & selRec("TruckTypeDesc"))
    isTrailer_ = Trim("" & selRec("isTrailer"))
    transUnitID_ = Trim("" & selRec("TransferUnitID"))
    transUnitName_ = Trim("" & selRec("TransferUnitID"))    'Trim("" & selRec("TransUnitName"))
    If IsNull(selRec("TransportFee")) = False Then transportFee_ = selRec("TransportFee")
    If IsNull(selRec("MinWeightPerCar")) = False Then minWeightPerCar_ = selRec("MinWeightPerCar")
    isPriceIncTrans_ = Trim("" & selRec("isPriceIncTransport"))
    IsDiscountPriceTrans_ = IIf(IsNull(selRec!IsDiscountPriceTrans), "N", selRec!IsDiscountPriceTrans)
    QuotationID_ = Trim("" & selRec("QuotationID"))
    If IsNull(selRec("StaffCountPerCar")) = False Then laborCount_ = selRec("StaffCountPerCar")
    If IsNull(selRec("ExtraLaborFee")) = False Then laborCharge_ = selRec("ExtraLaborFee")
    
    If isExist(selRec, "FailType") Then failType_ = Trim("" & selRec("FailType"))
    If isExist(selRec, "FailRemark") Then failRemark_ = Trim("" & selRec("FailRemark"))
    If isExist(selRec, "FailCheck") Then failCheck_ = selRec("FailCheck")
    
    If isExist(selRec, "TimeTableNo") Then TimeTableNo_ = Trim("" & selRec("TimeTableNo"))
    nonMinRate_ = selRec("NonMinRate")
    hazMinRate_ = selRec("HazMinRate")
    nblMinRate_ = selRec("NBLMinRate")
    hblMinRate_ = selRec("HBLMinRate")
    '*** MaxWeight, MaxRate
    If isExist(selRec, "MaxWeight") Then QuoMaxWeightPerCar_ = Trim("0" & selRec("MaxWeight"))
    If isExist(selRec, "MaxRate") Then QuoMaxRate_ = Trim("0" & selRec("MaxRate"))
    If isExist(selRec, "MaxRateType") Then QuoMaxRateType_ = Trim("" & selRec("MaxRateType"))
    If IsNull(selRec("HaType")) = False Then HAType_ = selRec("HaType")
    If IsNull(selRec("IsNewWaste")) = False Then IsNewWaste_ = selRec("IsNewWaste")
    If IsNull(selRec("IsNewWasteCR")) = False Then IsNewWasteCR_ = selRec("IsNewWasteCR")
    
    If isExist(selRec, "InWeight") Then
        If IsNull(selRec("InWeight")) = False Then InWeight_ = selRec("InWeight")
    End If
    If isExist(selRec, "TransRate") Then
        If IsNull(selRec("TransRate")) = False Then TransRate_ = selRec("TransRate")
    End If
    If isExist(selRec, "Cost") Then
        If IsNull(selRec("Cost")) = False Then QuoCost_ = selRec("Cost")
    End If
    If isExist(selRec, "CostUnit") Then
        If IsNull(selRec("CostUnit")) = False Then QuoCostUnit_ = selRec("CostUnit")
    End If
    
    CanEvap_ = Trim("" & selRec("CanEvap"))
    EvapDesc_ = Trim("" & selRec("EvapDesc"))
    Precipitation_ = Trim("" & selRec("Precipitation"))
    PrecipitationDesc_ = Trim("" & selRec("PrecipitationDesc"))
    Neutralization_ = Trim("" & selRec("Neutralization"))
    NeutralizationDesc_ = Trim("" & selRec("NeutralizationDesc"))
    CommentWaste_ = Trim("" & selRec("CommentWaste"))
    CommentWasteDesc_ = Trim("" & selRec("CommentWasteDesc"))
End Sub
