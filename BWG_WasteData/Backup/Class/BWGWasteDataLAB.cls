VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BWGWasteDataLAB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim id_ As String
Dim wasteDataID_ As String
Dim generateDetail_ As String
Dim wasteChar_ As String
Dim powderOrLump_ As String
Dim CharDescOther_ As String
Dim isPumpAble_ As String
Dim isExplode_ As String
Dim isFlameAble_ As String
Dim flashPoint_ As String
Dim isCorrosive_ As String
Dim phValue_ As String
Dim isReact_ As String
Dim reactWith_ As String
Dim isCarcinogenic_ As String
Dim carcinogenicDesc_ As String
Dim isToxic_ As String
Dim toxicType_ As String
Dim toxicDesc_ As String
Dim isRadioActive_ As String
Dim radioActiveDesc_ As String
Dim isPesticide_ As String
Dim pesticideDesc_ As String
Dim otherDesc_ As String
Dim mixOrSeperate_ As String
Dim inOrOutDoor_ As String
Dim storageType_ As String
Dim storageDesc_ As String
Dim oldEliminator_ As String
Dim colOldTreatment_ As Collection
Dim colNewTreatment_ As Collection
Dim colGenTreatments_ As Collection

Dim isOdour_ As String
Dim Odour_ As String
Dim isColor_ As String
Dim Color_ As String
Dim isHospital_ As String
Dim isPlasticBag_ As String
Dim PlasticBag_ As String
Dim isPlasticDrum_ As String
Dim PlasticDrum_ As String
Dim isPlasticGallon_ As String
Dim PlasticGallon_ As String
Dim isPaperBox_ As String
Dim PaperBox_ As String
Dim isBottle_ As String
Dim Bottle_ As String
Dim AnalysisType_ As Integer
Dim AnalysisDesc_ As String
Dim isBlending_ As String
Dim CanEvap_ As String
Dim isChanged_ As Boolean
Dim isNew_ As Boolean

Dim analysisNo_ As String
Dim Physicalstate_ As String
Dim Viscosity_ As String
Dim Bulkdensity_ As String
Dim OxidizerScreen_ As String
Dim FreeChlorine_ As String
Dim Nitrite_ As String
Dim Cyanide_ As String
Dim HeatingValue_ As String
Dim MoistureContent_ As String
Dim WaterContent_ As String
Dim AshContent_ As String
Dim CL_ As String
Dim F_ As String
Dim S_ As String
Dim SludgeContent_  As String
Dim SolidContent_ As String
Dim HeavyAs_ As String
Dim HeavyCd_ As String
Dim HeavyCr_ As String
Dim HeavyHg_ As String
Dim HeavyPb_ As String
Dim HeavyFe_ As String
Dim reportNo_ As String
Dim samplingDate_ As String
Dim ReceivedDate_ As String
Dim AnalysisDate_ As String
Dim LBstatus_ As String
Dim CommantLab_ As String
Dim Cu_ As String
Dim Ni_ As String
Dim Mn_ As String
Dim Zn_ As String
Dim Cut_ As String
Dim ExtAs_ As String
Dim RemarkAs_ As String
Dim ExtCd_ As String
Dim RemarkCd_ As String
Dim ExtCr_ As String
Dim RemarkCr_ As String
Dim ExtCu_ As String
Dim RemarkCu_ As String
Dim ExtPb_ As String
Dim RemarkPb_ As String
Dim ExtNi_ As String
Dim RemarkNi_ As String
Dim ExtZn_ As String
Dim RemarkZn_ As String
Dim ExtHg_ As String
Dim RemarkHg_ As String
Dim LabCode_ As String
Dim Concentrate_ As String
Dim SamplingBy_ As String
Dim SamplingByNo_ As String
Dim NCVValue_ As String
Dim SS_ As String
Dim Acidity_ As String
Dim Alkalinity_ As String
Dim Silicon_ As String
Dim SiliconRemark_ As String
Dim Aluminum_ As String
Dim AluminumRemark_ As String
Dim Iron_ As String
Dim IronRemark_ As String
Dim Calcium_ As String
Dim CalciumRemark_ As String
Dim InFlammable_  As String

Public Property Get InFlammable() As Boolean
    If InFlammable_ = "Y" Then
        InFlammable = True
    Else
        InFlammable = False
    End If
End Property
Public Property Let InFlammable(ByVal newval As Boolean)
Dim tmpStr As String
    If newval Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(InFlammable_) <> Trim(tmpStr) Then
        InFlammable_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property




Public Property Get ID() As String
    ID = id_
End Property
Public Property Let ID(ByVal newval As String)
    id_ = Trim(newval)
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
Public Property Get CanEvap() As String
    CanEvap = CanEvap_
End Property
Public Property Let CanEvap(ByVal newval As String)
    If Trim(CanEvap_) <> Trim(newval) Then
        CanEvap_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get GenerateDetail() As String
    GenerateDetail = generateDetail_
End Property
Public Property Let GenerateDetail(ByVal newval As String)
    If Trim(generateDetail_) <> Trim(newval) Then
        generateDetail_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get WasteCharacter() As String
    WasteCharacter = wasteChar_
End Property
Public Property Let WasteCharacter(ByVal newval As String)
    If Trim(wasteChar_) <> Trim(newval) Then
        wasteChar_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get PowderOrLump() As String
    PowderOrLump = powderOrLump_
End Property
Public Property Let PowderOrLump(ByVal newval As String)
    If Trim(powderOrLump_) <> Trim(newval) Then
        powderOrLump_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get CharDescOther() As String
    CharDescOther = CharDescOther_
End Property
Public Property Let CharDescOther(ByVal newval As String)
    If Trim(CharDescOther_) <> Trim(newval) Then
        CharDescOther_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get isPumpAble() As Boolean
    If isPumpAble_ = "Y" Then
        isPumpAble = True
    Else
        isPumpAble = False
    End If
End Property
Public Property Let isPumpAble(ByVal newval As Boolean)
Dim tmpStr As String
    If newval Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(isPumpAble_) <> Trim(tmpStr) Then
        isPumpAble_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property

Public Property Get isExplode() As Boolean
    If isExplode_ = "Y" Then
        isExplode = True
    Else
        isExplode = False
    End If
End Property
Public Property Let isExplode(ByVal newval As Boolean)
Dim tmpStr As String
    If newval Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(isExplode_) <> Trim(tmpStr) Then
        isExplode_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property

Public Property Get isFlameAble() As Boolean
    If isFlameAble_ = "Y" Then
        isFlameAble = True
    Else
        isFlameAble = False
    End If
End Property
Public Property Let isFlameAble(ByVal newval As Boolean)
Dim tmpStr As String
    If newval Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(isFlameAble_) <> Trim(tmpStr) Then
        isFlameAble_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property

Public Property Get FlashPoint() As String
    FlashPoint = flashPoint_
End Property
Public Property Let FlashPoint(ByVal newval As String)
        flashPoint_ = newval
        isChanged_ = True
End Property

Public Property Get isCorrosive() As Boolean
    If isCorrosive_ = "Y" Then
        isCorrosive = True
    Else
        isCorrosive = False
    End If
End Property
Public Property Let isCorrosive(ByVal newval As Boolean)
Dim tmpStr As String
    If newval Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(isCorrosive_) <> Trim(tmpStr) Then
        isCorrosive_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property

Public Property Get pHValue() As String
    pHValue = phValue_
End Property
Public Property Let pHValue(ByVal newval As String)
    If Trim(phValue_) <> Trim(newval) Then
        phValue_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get isReact() As Boolean
    If isReact_ = "Y" Then
        isReact = True
    Else
        isReact = False
    End If
End Property
Public Property Let isReact(ByVal newval As Boolean)
Dim tmpStr As String
    If newval Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(isReact_) <> Trim(tmpStr) Then
        isReact_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property

Public Property Get ReActWith() As String
    ReActWith = reactWith_
End Property
Public Property Let ReActWith(ByVal newval As String)
    If Trim(reactWith_) <> Trim(newval) Then
        reactWith_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get isCarcinoGenic() As Boolean
    If isCarcinogenic_ = "Y" Then
        isCarcinoGenic = True
    Else
        isCarcinoGenic = False
    End If
End Property
Public Property Let isCarcinoGenic(ByVal newval As Boolean)
Dim tmpStr As String
    If newval Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(isCarcinogenic_) <> Trim(tmpStr) Then
        isCarcinogenic_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property

Public Property Get CarcinoGenicDesc() As String
    CarcinoGenicDesc = carcinogenicDesc_
End Property
Public Property Let CarcinoGenicDesc(ByVal newval As String)
    If Trim(carcinogenicDesc_) <> Trim(newval) Then
        carcinogenicDesc_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get isToxic() As Boolean
    If isToxic_ = "Y" Then
        isToxic = True
    Else
        isToxic = False
    End If
End Property
Public Property Let isToxic(ByVal newval As Boolean)
Dim tmpStr As String
    If newval Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(isToxic_) <> Trim(tmpStr) Then
        isToxic_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property

Public Property Get ToxicType() As String
    ToxicType = toxicType_
End Property
Public Property Let ToxicType(ByVal newval As String)
    If Trim(toxicType_) <> Trim(newval) Then
        toxicType_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get ToxicDesc() As String
    ToxicDesc = toxicDesc_
End Property
Public Property Let ToxicDesc(ByVal newval As String)
    If Trim(toxicDesc_) <> Trim(newval) Then
        toxicDesc_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get isRadioActive() As Boolean
    If isRadioActive_ = "Y" Then
        isRadioActive = True
    Else
        isRadioActive = False
    End If
End Property
Public Property Let isRadioActive(ByVal newval As Boolean)
Dim tmpStr As String
    If newval Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(isRadioActive_) <> Trim(tmpStr) Then
        isRadioActive_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property

Public Property Get RadioActiveDesc() As String
    RadioActiveDesc = radioActiveDesc_
End Property
Public Property Let RadioActiveDesc(ByVal newval As String)
    If Trim(radioActiveDesc_) <> Trim(newval) Then
        radioActiveDesc_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get isPesticide() As Boolean
    If isPesticide_ = "Y" Then
        isPesticide = True
    Else
        isPesticide = False
    End If
End Property
Public Property Let isPesticide(ByVal newval As Boolean)
Dim tmpStr As String
    If newval Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(isPesticide_) <> Trim(tmpStr) Then
        isPesticide_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property

Public Property Get PesticideDesc() As String
    PesticideDesc = pesticideDesc_
End Property
Public Property Let PesticideDesc(ByVal newval As String)
    If Trim(pesticideDesc_) <> Trim(newval) Then
        pesticideDesc_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get OtherDesc() As String
    OtherDesc = otherDesc_
End Property
Public Property Let OtherDesc(ByVal newval As String)
    If Trim(otherDesc_) <> Trim(newval) Then
        otherDesc_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get MixOrSeperate() As String
    MixOrSeperate = mixOrSeperate_
End Property
Public Property Let MixOrSeperate(ByVal newval As String)
    If Trim(mixOrSeperate_) <> Trim(newval) Then
        mixOrSeperate_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get InOrOutDoor() As String
    InOrOutDoor = inOrOutDoor_
End Property
Public Property Let InOrOutDoor(ByVal newval As String)
    If Trim(inOrOutDoor_) <> Trim(newval) Then
        inOrOutDoor_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get StorageType() As String
    StorageType = storageType_
End Property
Public Property Let StorageType(ByVal newval As String)
    If Trim(storageType_) <> Trim(newval) Then
        storageType_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get StorageDesc() As String
    StorageDesc = storageDesc_
End Property
Public Property Let StorageDesc(ByVal newval As String)
    If Trim(storageDesc_) <> Trim(newval) Then
        storageDesc_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get OldEliminator() As String
    OldEliminator = oldEliminator_
End Property
Public Property Let OldEliminator(ByVal newval As String)
    If Trim(oldEliminator_) <> Trim(newval) Then
        oldEliminator_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get OldTreatments() As Collection
Dim tmpS As New BWGSearchManager
    If isNew_ = False Then
        If colOldTreatment_ Is Nothing Then Set colOldTreatment_ = tmpS.TreatmentDataSearch("TreatmentID in (Select TreatmentID from tbOldTreatment Where WasteDataLabID='" & id_ & "')")
    Else
        If colOldTreatment_ Is Nothing Then Set colOldTreatment_ = tmpS.TreatmentDataSearch("TreatmentID in (Select TreatmentID from tbOldTreatment_New Where WasteDataLabID='" & id_ & "')")
    End If
    Set OldTreatments = colOldTreatment_
End Property
Public Property Set OldTreatments(ByVal newval As Collection)
    Set colOldTreatment_ = newval
    isChanged_ = True
End Property

Public Property Get NewTreatments() As Collection
Dim tmpS As New BWGSearchManager
    If colNewTreatment_ Is Nothing Then Set colNewTreatment_ = tmpS.TreatmentDataSearch("TreatmentID in (Select TreatmentID from tbTreatment Where WasteDataLabID='" & id_ & "')")
    Set NewTreatments = colNewTreatment_
End Property
Public Property Set NewTreatments(ByVal newval As Collection)
    Set colNewTreatment_ = newval
    isChanged_ = True
End Property

Public Sub RequeryOldTreatment()
    Set colOldTreatment_ = Nothing
End Sub

Public Sub RequeryNewTreatment()
    Set colNewTreatment_ = Nothing
End Sub

'*********** New Variable 13/06/2551 ******************
Public Property Get isOdour() As Boolean
    If isOdour_ = "Y" Then
        isOdour = True
    Else
        isOdour = False
    End If
End Property
Public Property Let isOdour(ByVal newval As Boolean)
    If newval Then
        isOdour_ = "Y"
    Else
        isOdour_ = "N"
    End If
    isChanged_ = True
End Property

Public Property Get Odour() As String
    Odour = Odour_
End Property
Public Property Let Odour(ByVal newval As String)
    If Trim(Odour_) <> Trim(newval) Then
        Odour_ = Trim(newval)
        isChanged_ = True
    End If
    isChanged_ = True
End Property

Public Property Get isColor() As Boolean
    If isColor_ = "Y" Then
        isColor = True
    Else
        isColor = False
    End If
End Property
Public Property Let isColor(ByVal newval As Boolean)
    If newval Then
        isColor_ = "Y"
    Else
        isColor_ = "N"
    End If
    isChanged_ = True
End Property

Public Property Get Color() As String
    Color = Color_
End Property
Public Property Let Color(ByVal newval As String)
    Color_ = Trim(newval)
    isChanged_ = True
End Property

Public Property Get isHospital() As Boolean
    If isHospital_ = "Y" Then
        isHospital = True
    Else
        isHospital = False
    End If
End Property
Public Property Let isHospital(ByVal newval As Boolean)
    If newval Then
        isHospital_ = "Y"
    Else
        isHospital_ = "N"
    End If
    isChanged_ = True
End Property

Public Property Get isPlasticBag() As Boolean
    If isPlasticBag_ = "Y" Then
        isPlasticBag = True
    Else
        isPlasticBag = False
    End If
End Property
Public Property Let isPlasticBag(ByVal newval As Boolean)
    If newval Then
        isPlasticBag_ = "Y"
    Else
        isPlasticBag_ = "N"
    End If
    isChanged_ = True
End Property

Public Property Get PlasticBag() As String
    PlasticBag = PlasticBag_
End Property
Public Property Let PlasticBag(ByVal newval As String)
    PlasticBag_ = Trim(newval)
    isChanged_ = True
End Property

Public Property Get isPlasticDrum() As Boolean
    If isPlasticDrum_ = "Y" Then
        isPlasticDrum = True
    Else
        isPlasticDrum = False
    End If
End Property
Public Property Let isPlasticDrum(ByVal newval As Boolean)
    If newval Then
        isPlasticDrum_ = "Y"
    Else
        isPlasticDrum_ = "N"
    End If
    isChanged_ = True
End Property

Public Property Get PlasticDrum() As String
    PlasticDrum = PlasticDrum_
End Property
Public Property Let PlasticDrum(ByVal newval As String)
    PlasticDrum_ = Trim(newval)
    isChanged_ = True
End Property

Public Property Get isPlasticGallon() As Boolean
    If isPlasticGallon_ = "Y" Then
        isPlasticGallon = True
    Else
        isPlasticGallon = False
    End If
End Property
Public Property Let isPlasticGallon(ByVal newval As Boolean)
    If newval Then
        isPlasticGallon_ = "Y"
    Else
        isPlasticGallon_ = "N"
    End If
    isChanged_ = True
End Property

Public Property Get PlasticGallon() As String
    PlasticGallon = PlasticGallon_
End Property
Public Property Let PlasticGallon(ByVal newval As String)
    PlasticGallon_ = Trim(newval)
    isChanged_ = True
End Property

Public Property Get isPaperBox() As Boolean
    If isPaperBox_ = "Y" Then
        isPaperBox = True
    Else
        isPaperBox = False
    End If
End Property
Public Property Let isPaperBox(ByVal newval As Boolean)
    If newval Then
        isPaperBox_ = "Y"
    Else
        isPaperBox_ = "N"
    End If
    isChanged_ = True
End Property

Public Property Get PaperBox() As String
    PaperBox = PaperBox_
End Property
Public Property Let PaperBox(ByVal newval As String)
    PaperBox_ = Trim(newval)
    isChanged_ = True
End Property

Public Property Get isBottle() As Boolean
    If isBottle_ = "Y" Then
        isBottle = True
    Else
        isBottle = False
    End If
End Property
Public Property Let isBottle(ByVal newval As Boolean)
    If newval Then
        isBottle_ = "Y"
    Else
        isBottle_ = "N"
    End If
    isChanged_ = True
End Property

Public Property Get Bottle() As String
    Bottle = Bottle_
End Property
Public Property Let Bottle(ByVal newval As String)
    Bottle_ = Trim(newval)
    isChanged_ = True
End Property

Public Property Get AnalysisType() As Integer
    AnalysisType = AnalysisType_
End Property
Public Property Let AnalysisType(ByVal newval As Integer)
    AnalysisType_ = newval
    isChanged_ = True
End Property


Public Property Get isBlending() As String
    isBlending = isBlending_
End Property
Public Property Let isBlending(ByVal newval As String)
    isBlending_ = newval
    isChanged_ = True
End Property
'*********** New Variable 13/06/2551 ******************

Public Property Get isChanged() As Boolean
    isChanged = isChanged_
End Property
Public Property Let isChanged(ByVal newval As Boolean)
    isChanged_ = newval
End Property

Public Property Get isNew() As Boolean
    isNew = isNew_
End Property
Public Property Let isNew(ByVal newval As Boolean)
    isNew_ = newval
End Property
Public Property Get AnalysisNo() As String
    AnalysisNo = analysisNo_
End Property
Public Property Let AnalysisNo(ByVal newval As String)
    analysisNo_ = newval
End Property
Public Property Get Physicalstate() As String
    Physicalstate = Physicalstate_
End Property
Public Property Let Physicalstate(ByVal newval As String)
    Physicalstate_ = newval
End Property
Public Property Get Viscosity() As String
    Viscosity = Viscosity_
End Property
Public Property Let Viscosity(ByVal newval As String)
    Viscosity_ = newval
End Property
Public Property Get Bulkdensity() As String
    Bulkdensity = Bulkdensity_
End Property
Public Property Let Bulkdensity(ByVal newval As String)
    Bulkdensity_ = newval
End Property
Public Property Get OxidizerScreen() As String
    OxidizerScreen = OxidizerScreen_
End Property
Public Property Let OxidizerScreen(ByVal newval As String)
    OxidizerScreen_ = newval
End Property
Public Property Get FreeChlorine() As String
    FreeChlorine = FreeChlorine_
End Property
Public Property Let FreeChlorine(ByVal newval As String)
    FreeChlorine_ = newval
End Property
Public Property Get Nitrite() As String
    Nitrite = Nitrite_
End Property
Public Property Let Nitrite(ByVal newval As String)
    Nitrite_ = newval
End Property
Public Property Get Cyanide() As String
    Cyanide = Cyanide_
End Property
Public Property Let Cyanide(ByVal newval As String)
    Cyanide_ = newval
End Property
Public Property Get HeatingValue() As String
    HeatingValue = HeatingValue_
End Property
Public Property Let HeatingValue(ByVal newval As String)
    HeatingValue_ = newval
End Property
Public Property Get NCVValue() As String
    NCVValue = NCVValue_
End Property
Public Property Let NCVValue(ByVal newval As String)
    NCVValue_ = newval
End Property

Public Property Get MoistureContent() As String
    MoistureContent = MoistureContent_
End Property
Public Property Let MoistureContent(ByVal newval As String)
    MoistureContent_ = newval
End Property
Public Property Get WaterContent() As String
    WaterContent = WaterContent_
End Property
Public Property Let WaterContent(ByVal newval As String)
    WaterContent_ = newval
End Property
Public Property Get AshContent() As String
    AshContent = AshContent_
End Property
Public Property Let AshContent(ByVal newval As String)
    AshContent_ = newval
End Property

Public Property Get CL() As String
    CL = CL_
End Property
Public Property Let CL(ByVal newval As String)
    CL_ = newval
End Property
Public Property Get f() As String
    f = F_
End Property
Public Property Let f(ByVal newval As String)
    F_ = newval
End Property
Public Property Get s() As String
    s = S_
End Property
Public Property Let s(ByVal newval As String)
    S_ = newval
End Property
Public Property Get SludgeContent() As String
    SludgeContent = SludgeContent_
End Property
Public Property Let SludgeContent(ByVal newval As String)
    SludgeContent_ = newval
End Property
Public Property Get SolidContent() As String
    SolidContent = SolidContent_
End Property
Public Property Let SolidContent(ByVal newval As String)
    SolidContent_ = newval
End Property
Public Property Get HeavyAs() As String
    HeavyAs = HeavyAs_
End Property
Public Property Let HeavyAs(ByVal newval As String)
    HeavyAs_ = newval
End Property
Public Property Get HeavyCd() As String
    HeavyCd = HeavyCd_
End Property
Public Property Let HeavyCd(ByVal newval As String)
    HeavyCd_ = newval
End Property
Public Property Get HeavyCr() As String
    HeavyCr = HeavyCr_
End Property
Public Property Let HeavyCr(ByVal newval As String)
    HeavyCr_ = newval
End Property
Public Property Get HeavyHg() As String
    HeavyHg = HeavyHg_
End Property
Public Property Let HeavyHg(ByVal newval As String)
    HeavyHg_ = newval
End Property
Public Property Get HeavyPb() As String
    HeavyPb = HeavyPb_
End Property
Public Property Let HeavyPb(ByVal newval As String)
    HeavyPb_ = newval
End Property
Public Property Get HeavyFe() As String
    HeavyFe = HeavyFe_
End Property
Public Property Let HeavyFe(ByVal newval As String)
    HeavyFe_ = newval
End Property
Public Property Get ReportNo() As String
    ReportNo = reportNo_
End Property
Public Property Let ReportNo(ByVal newval As String)
    reportNo_ = newval
End Property
Public Property Get SamplingDate() As String
    SamplingDate = samplingDate_
End Property
Public Property Let SamplingDate(ByVal newval As String)
    samplingDate_ = newval
End Property
Public Property Get ReceivedDate() As String
    ReceivedDate = ReceivedDate_
End Property
Public Property Let ReceivedDate(ByVal newval As String)
    ReceivedDate_ = newval
End Property
Public Property Get AnalysisDate() As String
    AnalysisDate = AnalysisDate_
End Property
Public Property Let AnalysisDate(ByVal newval As String)
    AnalysisDate_ = newval
End Property
Public Property Get LBstatus() As String
    LBstatus = LBstatus_
End Property
Public Property Let LBstatus(ByVal newval As String)
    LBstatus_ = newval
End Property
Public Property Get CommantLab() As String
    CommantLab = CommantLab_
End Property
Public Property Let CommantLab(ByVal newval As String)
    CommantLab_ = newval
End Property
Public Property Get Cu() As String
    Cu = Cu_
End Property
Public Property Let Cu(ByVal newval As String)
    Cu_ = newval
End Property
Public Property Get Ni() As String
    Ni = Ni_
End Property
Public Property Let Ni(ByVal newval As String)
    Ni_ = newval
End Property
Public Property Get Mn() As String
    Mn = Mn_
End Property
Public Property Let Mn(ByVal newval As String)
    Mn_ = newval
End Property
Public Property Get Concentrate() As String
    Concentrate = Concentrate_
End Property
Public Property Let Concentrate(ByVal newval As String)
    Concentrate_ = newval
End Property
Public Property Get Zn() As String
    Zn = Zn_
End Property
Public Property Let Zn(ByVal newval As String)
    Zn_ = newval
End Property
Public Property Get Cut() As String
    Cut = Cut_
End Property
Public Property Let Cut(ByVal newval As String)
    Cut_ = newval
End Property
Public Property Get ExtAs() As String
    ExtAs = ExtAs_
End Property
Public Property Let ExtAs(ByVal newval As String)
    ExtAs_ = newval
End Property
Public Property Get RemarkAs() As String
    RemarkAs = RemarkAs_
End Property
Public Property Let RemarkAs(ByVal newval As String)
    RemarkAs_ = newval
End Property
Public Property Get ExtCd() As String
    ExtCd = ExtCd_
End Property
Public Property Let ExtCd(ByVal newval As String)
    ExtCd_ = newval
End Property
Public Property Get RemarkCd() As String
    RemarkCd = RemarkCd_
End Property
Public Property Let RemarkCd(ByVal newval As String)
    RemarkCd_ = newval
End Property
Public Property Get ExtCr() As String
    ExtCr = ExtCr_
End Property
Public Property Let ExtCr(ByVal newval As String)
    ExtCr_ = newval
End Property
Public Property Get RemarkCr() As String
    RemarkCr = RemarkCr_
End Property
Public Property Let RemarkCr(ByVal newval As String)
    RemarkCr_ = newval
End Property
Public Property Get ExtCu() As String
    ExtCu = ExtCu_
End Property
Public Property Let ExtCu(ByVal newval As String)
    ExtCu_ = newval
End Property
Public Property Get RemarkCu() As String
    RemarkCu = RemarkCu_
End Property
Public Property Let RemarkCu(ByVal newval As String)
    RemarkCu_ = newval
End Property
Public Property Get ExtPb() As String
    ExtPb = ExtPb_
End Property
Public Property Let ExtPb(ByVal newval As String)
    ExtPb_ = newval
End Property
Public Property Get RemarkPb() As String
    RemarkPb = RemarkPb_
End Property
Public Property Let RemarkPb(ByVal newval As String)
    RemarkPb_ = newval
End Property
Public Property Get ExtNi() As String
    ExtNi = ExtNi_
End Property
Public Property Let ExtNi(ByVal newval As String)
    ExtNi_ = newval
End Property
Public Property Get RemarkNi() As String
    RemarkNi = RemarkNi_
End Property
Public Property Let RemarkNi(ByVal newval As String)
    RemarkNi_ = newval
End Property
Public Property Get ExtZn() As String
    ExtZn = ExtZn_
End Property
Public Property Let ExtZn(ByVal newval As String)
    ExtZn_ = newval
End Property
Public Property Get RemarkZn() As String
    RemarkZn = RemarkZn_
End Property
Public Property Let RemarkZn(ByVal newval As String)
    RemarkZn_ = newval
End Property
Public Property Get ExtHg() As String
    ExtHg = ExtHg_
End Property
Public Property Let ExtHg(ByVal newval As String)
    ExtHg_ = newval
End Property
Public Property Get RemarkHg() As String
    RemarkHg = RemarkHg_
End Property
Public Property Let RemarkHg(ByVal newval As String)
    RemarkHg_ = newval
End Property
Public Property Get LabCode() As String
    LabCode = LabCode_
End Property
Public Property Let LabCode(ByVal newval As String)
    LabCode_ = newval
End Property
Public Property Let SS(ByVal newval As String)
    SS_ = newval
End Property
Public Property Get SS() As String
    SS = SS_
End Property
Public Property Let Acidity(ByVal newval As String)
    Acidity_ = newval
End Property
Public Property Get Acidity() As String
    Acidity = Acidity_
End Property
Public Property Let Alkalinity(ByVal newval As String)
    Alkalinity_ = newval
End Property
Public Property Get Alkalinity() As String
    Alkalinity = Alkalinity_
End Property

Public Property Let AnalysisDesc(ByVal newval As String)
    AnalysisDesc_ = Trim(newval)
    isChanged_ = True
End Property
Public Property Get AnalysisDesc() As String
    AnalysisDesc = AnalysisDesc_
End Property
Public Property Get SamplingBy() As String
    SamplingBy = SamplingBy_
End Property
Public Property Let SamplingBy(ByVal newval As String)
    SamplingBy_ = Trim(newval)
    isChanged_ = True
End Property
Public Property Get SamplingByNo() As String
    SamplingByNo = SamplingByNo_
End Property
Public Property Let SamplingByNo(ByVal newval As String)
    SamplingByNo_ = Trim(newval)
    isChanged_ = True
End Property
Public Property Get Silicon() As String
    Silicon = Silicon_
End Property
Public Property Let Silicon(ByVal newval As String)
    Silicon_ = Trim(newval)
    isChanged_ = True
End Property
Public Property Get SiliconRemark() As String
    SiliconRemark = SiliconRemark_
End Property
Public Property Let SiliconRemark(ByVal newval As String)
    SiliconRemark_ = Trim(newval)
    isChanged_ = True
End Property
Public Property Get Aluminum() As String
    Aluminum = Aluminum_
End Property
Public Property Let Aluminum(ByVal newval As String)
    Aluminum_ = Trim(newval)
    isChanged_ = True
End Property
Public Property Get AluminumRemark() As String
    AluminumRemark = AluminumRemark_
End Property
Public Property Let AluminumRemark(ByVal newval As String)
    AluminumRemark_ = Trim(newval)
    isChanged_ = True
End Property
Public Property Get Iron() As String
    Iron = Iron_
End Property
Public Property Let Iron(ByVal newval As String)
    Iron_ = Trim(newval)
    isChanged_ = True
End Property
Public Property Get IronRemark() As String
    IronRemark = IronRemark_
End Property
Public Property Let IronRemark(ByVal newval As String)
    IronRemark_ = Trim(newval)
    isChanged_ = True
End Property
Public Property Get Calcium() As String
    Calcium = Calcium_
End Property
Public Property Let Calcium(ByVal newval As String)
    Calcium_ = Trim(newval)
    isChanged_ = True
End Property
Public Property Get CalciumRemark() As String
    CalciumRemark = CalciumRemark_
End Property
Public Property Let CalciumRemark(ByVal newval As String)
    CalciumRemark_ = Trim(newval)
    isChanged_ = True
End Property

Public Sub Save()
On Error GoTo ErrD
Dim tmpRec As ADODB.Recordset
Dim x%
    Set tmpRec = New ADODB.Recordset
'    If isNew_ = False Then
        DBManager.GetDBConnection.Execute "DELETE FROM tbOldTreatment WHERE WasteDataLabID = '" & id_ & "'"
        DBManager.GetDBConnection.Execute "DELETE FROM tbTreatment WHERE WasteDataLabID = '" & id_ & "'"
'        DBManager.GetDBConnection.Execute "DELETE FROM tbWasteDataLab WHERE WasteDataID = '" & wasteDataID_ & "'"
        tmpRec.Open "Select * from tbWasteDataLab Where WasteDataLabID='" & id_ & "'", DBManager.GetDBConnection, adOpenDynamic, adLockOptimistic
'    Else
'        tmpRec.Open "Select * from tbWasteDataLab_New Where WasteDataLabID='" & id_ & "'", DBManager.GetDBConnection, adOpenDynamic, adLockOptimistic
'    End If
    If tmpRec.EOF Then
        tmpRec.AddNew
        tmpRec("WasteDataLabID") = id_
    End If
    tmpRec("WasteDataID") = Trim(wasteDataID_)
    tmpRec("GenerateDetail") = Trim(generateDetail_)
    tmpRec("WasteCharacter") = Trim(wasteChar_)
    tmpRec("PowderOrLump") = Trim(powderOrLump_)
    tmpRec("CharDescOther") = Trim(CharDescOther_)
    tmpRec("isPumpable") = Trim(isPumpAble_)
    tmpRec("isExplode") = Trim(isExplode_)
    tmpRec("isFlamable") = Trim(isFlameAble_)
    tmpRec("FlashPoint") = Trim(flashPoint_)
    tmpRec("isCorrosive") = Trim(isCorrosive_)
    tmpRec("pHValue") = Trim(phValue_)
    tmpRec("isReact") = Trim(isReact_)
    tmpRec("ReactWith") = Trim(reactWith_)
    tmpRec("isCarcinogenic") = Trim(isCarcinogenic_)
    tmpRec("CarcinogenicDesc") = Trim(carcinogenicDesc_)
    tmpRec("isToxic") = Trim(isToxic_)
    tmpRec("ToxicType") = Trim(toxicType_)
    tmpRec("ToxicDesc") = Trim(toxicDesc_)
    tmpRec("isRadioActive") = Trim(isRadioActive_)
    tmpRec("RadioActiveDesc") = Trim(radioActiveDesc_)
    If isPesticide_ = "N" Then
        tmpRec("isPesticide") = ""
    Else
        tmpRec("isPesticide") = Trim(isPesticide_)
    End If
    tmpRec("PesticideDesc") = Trim(pesticideDesc_)
    tmpRec("OtherDesc") = Trim(otherDesc_)
    tmpRec("MixOrSeperate") = Trim(mixOrSeperate_)
    tmpRec("InOrOutDoor") = Trim(inOrOutDoor_)
    tmpRec("StorageType") = Trim(storageType_)
    tmpRec("StorageDesc") = Trim(storageDesc_)
    tmpRec("OldEliminator") = Trim(oldEliminator_)
    tmpRec("AnalysisNo") = Trim(analysisNo_)
    tmpRec("Physicalstate") = Trim(Physicalstate_)
    tmpRec("Viscosity") = Trim(Viscosity_)
    tmpRec("Bulkdensity") = Trim(Bulkdensity_)
    tmpRec("OxidizerScreen") = Trim(OxidizerScreen_)
    tmpRec("FreeChlorine") = Trim(FreeChlorine_)
    tmpRec("Nitrite") = Trim(Nitrite_)
    tmpRec("Cyanide") = Trim(Cyanide_)
    tmpRec("HeatingValue") = Trim(HeatingValue_)
    tmpRec("NCVValue") = Trim(NCVValue_)
    tmpRec("MoistureContent") = Trim(MoistureContent_)
    tmpRec("WaterContent") = Trim(WaterContent_)
    tmpRec("AshContent") = Trim(AshContent_)
    tmpRec("CL") = Trim(CL_)
    tmpRec("F") = Trim(F_)
    tmpRec("S") = Trim(S_)
    tmpRec("SludgeContent") = Trim(SludgeContent_)
    tmpRec("SolidContent") = Trim(SolidContent_)
    tmpRec("HeavyAs") = Trim(HeavyAs_)
    tmpRec("HeavyCd") = Trim(HeavyCd_)
    tmpRec("HeavyCr") = Trim(HeavyCr_)
    tmpRec("HeavyHg") = Trim(HeavyHg_)
    tmpRec("HeavyPb") = Trim(HeavyPb_)
    tmpRec("HeavyFe") = Trim(HeavyFe_)
    tmpRec("ReportNo") = Trim(reportNo_)
    tmpRec("SS") = Trim(SS_)
    tmpRec("Acidity") = Trim(Acidity_)
    tmpRec("Alkalinity") = Trim(Alkalinity)
    tmpRec("SamplingDate") = Trim(samplingDate_)
    tmpRec("ReceivedDate") = Trim(ReceivedDate_)
    tmpRec("AnalyticalDate") = Trim(AnalysisDate_)
    tmpRec("SamplingBy") = Trim(SamplingBy_)
    tmpRec("LBstatus") = Trim(LBstatus_)
    tmpRec("CommantLab") = Trim(CommantLab_)
    '**********************------*******************************************
    tmpRec!isHospital = Trim(isHospital_)
    tmpRec!isOdour = Trim(isOdour_)
    tmpRec!Odour = Trim(Odour_)
    tmpRec!isColor = Trim(isColor_)
    tmpRec!Color = Trim(Color_)
    tmpRec!isPlasticBag = Trim(isPlasticBag_)
    tmpRec!PlasticBag = Trim(PlasticBag_)
    tmpRec!isPlasticDrum = Trim(isPlasticDrum_)
    tmpRec!PlasticDrum = Trim(PlasticDrum_)
    tmpRec!isPlasticGallon = Trim(isPlasticGallon_)
    tmpRec!PlasticGallon = Trim(PlasticGallon_)
    tmpRec!isPaperBox = Trim(isPaperBox_)
    tmpRec!PaperBox = Trim(PaperBox_)
    tmpRec!isBottle = Trim(isBottle_)
    tmpRec!Bottle = Trim(Bottle_)
    tmpRec!AnalysisType = Trim(AnalysisType_)
    tmpRec!AnalysisDesc = Trim(AnalysisDesc_)
    tmpRec!isBlending = Trim(isBlending_)
    tmpRec!Cu = Trim(Cu_)
    tmpRec!Ni = Trim(Ni_)
    tmpRec!Mn = Trim(Mn_)
    tmpRec!Zn = Trim(Zn_)
    tmpRec!Cut = Trim(Cut_)
    tmpRec!ExtAs = Trim(ExtAs_)
    tmpRec!RemarkAs = Trim(RemarkAs_)
    tmpRec!ExtCd = Trim(ExtCd_)
    tmpRec!RemarkCd = Trim(RemarkCd_)
    tmpRec!ExtCr = Trim(ExtCr_)
    tmpRec!RemarkCr = Trim(RemarkCr_)
    tmpRec!ExtCu = Trim(ExtCu_)
    tmpRec!RemarkCu = Trim(RemarkCu_)
    tmpRec!ExtPb = Trim(ExtPb_)
    tmpRec!RemarkPb = Trim(RemarkPb_)
    tmpRec!ExtNi = Trim(ExtNi_)
    tmpRec!RemarkNi = Trim(RemarkNi_)
    tmpRec!ExtZn = Trim(ExtZn_)
    tmpRec!RemarkZn = Trim(RemarkZn_)
    tmpRec!ExtHg = Trim(ExtHg_)
    tmpRec!RemarkHg = Trim(RemarkHg_)
    tmpRec!LabCode = Trim(LabCode_)
    tmpRec!Concentrate = Trim(Concentrate_)
    tmpRec!SamplingBy = Trim(SamplingBy_)
    tmpRec!SamplingByNo = Trim(SamplingByNo_)
    tmpRec!Silicon = Trim(Silicon_)
    tmpRec!SiliconRemark = Trim(SiliconRemark_)
    tmpRec!Aluminum = Trim(Aluminum_)
    tmpRec!AluminumRemark = Trim(AluminumRemark_)
    tmpRec!Iron = Trim(Iron_)
    tmpRec!IronRemark = Trim(IronRemark_)
    tmpRec!Calcium = Trim(Calcium_)
    tmpRec!CalciumRemark = Trim(CalciumRemark_)
    tmpRec!InFlammable = Trim(InFlammable_)
    tmpRec.Update
    tmpRec.Close
    With DBManager.GetDBConnection
        If Not colNewTreatment_ Is Nothing Then
            .Execute "DELETE FROM  tbOldTreatment  WHERE WasteDataLabID = '" & id_ & "'"
            .Execute "Delete tbTreatment Where WasteDataLabID='" & id_ & "'"
'            If isNew_ = False Then
                .Execute "Delete tbTreatment Where WasteDataLabID='" & id_ & "'"
                For x = 1 To colNewTreatment_.Count
                    .Execute "Insert into tbTreatment (WasteDataLabID,TreatmentID) Values ('" & id_ & "','" & colNewTreatment_(x).ID & "')"
                Next
'            Else
'                .Execute "Delete tbTreatment_New Where WasteDataLabID='" & id_ & "'"
'                For x = 1 To colNewTreatment_.Count
'                    .Execute "Insert into tbTreatment_New (WasteDataLabID,TreatmentID) Values ('" & id_ & "','" & colNewTreatment_(x).ID & "')"
'                Next
'            End If
        End If
            .Execute "update tbwastedataCR set CanEvap='" & CanEvap_ & "' where wastedataid='" & wasteDataID_ & "' "
    End With
    Set tmpRec = Nothing
    isChanged_ = False
    Exit Sub
ErrD:
    MsgBox Err.Description & vbCrLf & "กรุณาลองใหม่หรือติดต่อผู้พัฒนาระบบ", vbCritical, "Error"
    Exit Sub
End Sub
Public Sub Add()
 On Error GoTo ErrD
Dim tmpRec As ADODB.Recordset
Dim x%
    Set tmpRec = New ADODB.Recordset
    Dim Lab_ID As String
    Lab_ID = GetGUID
    tmpRec.Open "Select * from tbWasteDataLab  Where WasteDataLabID='" & Lab_ID & "'", DBManager.GetDBConnection, adOpenDynamic, adLockOptimistic
    If tmpRec.EOF Then
        tmpRec.AddNew
        tmpRec("WasteDataLabID") = Lab_ID
    End If
    tmpRec("WasteDataID") = Wastelab_ID
    tmpRec("GenerateDetail") = Trim(generateDetail_)
    tmpRec("WasteCharacter") = Trim(wasteChar_)
    tmpRec("PowderOrLump") = Trim(powderOrLump_)
    tmpRec("CharDescOther") = Trim(CharDescOther_)
    tmpRec("isPumpable") = Trim(isPumpAble_)
    tmpRec("isExplode") = Trim(isExplode_)
    tmpRec("isFlamable") = Trim(isFlameAble_)
    tmpRec("FlashPoint") = Trim(flashPoint_)
    tmpRec("isCorrosive") = Trim(isCorrosive_)
    tmpRec("pHValue") = Trim(phValue_)
    tmpRec("isReact") = Trim(isReact_)
    tmpRec("ReactWith") = Trim(reactWith_)
    tmpRec("isCarcinogenic") = Trim(isCarcinogenic_)
    tmpRec("CarcinogenicDesc") = Trim(carcinogenicDesc_)
    tmpRec("isToxic") = Trim(isToxic_)
    tmpRec("ToxicType") = Trim(toxicType_)
    tmpRec("ToxicDesc") = Trim(toxicDesc_)
    tmpRec("isRadioActive") = Trim(isRadioActive_)
    tmpRec("RadioActiveDesc") = Trim(radioActiveDesc_)
    If isPesticide_ = "N" Then
        tmpRec("isPesticide") = ""
    Else
        tmpRec("isPesticide") = Trim(isPesticide_)
    End If
    tmpRec("PesticideDesc") = Trim(pesticideDesc_)
    tmpRec("OtherDesc") = Trim(otherDesc_)
    tmpRec("MixOrSeperate") = Trim(mixOrSeperate_)
    tmpRec("InOrOutDoor") = Trim(inOrOutDoor_)
    tmpRec("StorageType") = Trim(storageType_)
    tmpRec("StorageDesc") = Trim(storageDesc_)
    tmpRec("OldEliminator") = Trim(oldEliminator_)
    'tmpRec("AnalysisNo") = Trim(analysisNo_)
    tmpRec("Physicalstate") = Trim(Physicalstate_)
    tmpRec("Viscosity") = Trim(Viscosity_)
    tmpRec("Bulkdensity") = Trim(Bulkdensity_)
    tmpRec("OxidizerScreen") = Trim(OxidizerScreen_)
    tmpRec("FreeChlorine") = Trim(FreeChlorine_)
    tmpRec("Nitrite") = Trim(Nitrite_)
    tmpRec("Cyanide") = Trim(Cyanide_)
    tmpRec("HeatingValue") = Trim(HeatingValue_)
    tmpRec("MoistureContent") = Trim(MoistureContent_)
    tmpRec("WaterContent") = Trim(WaterContent_)
    tmpRec("AshContent") = Trim(AshContent_)
    tmpRec("CL") = Trim(CL_)
    tmpRec("F") = Trim(F_)
    tmpRec("S") = Trim(S_)
    tmpRec("SludgeContent") = Trim(SludgeContent_)
    tmpRec("SolidContent") = Trim(SolidContent_)
    tmpRec("HeavyAs") = Trim(HeavyAs_)
    tmpRec("HeavyCd") = Trim(HeavyCd_)
    tmpRec("HeavyCr") = Trim(HeavyCr_)
    tmpRec("HeavyHg") = Trim(HeavyHg_)
    tmpRec("HeavyPb") = Trim(HeavyPb_)
    tmpRec("HeavyFe") = Trim(HeavyFe_)
    tmpRec("ReportNo") = Trim(reportNo_)
    tmpRec("SS") = Trim(SS_)
    tmpRec("Acidity") = Trim(Acidity_)
    tmpRec("Alkalinity") = Trim(Alkalinity)
    tmpRec("SamplingDate") = Trim(samplingDate_)
    tmpRec("ReceivedDate") = Trim(ReceivedDate_)
    tmpRec("AnalyticalDate") = Trim(AnalysisDate_)
    tmpRec("LBstatus") = Trim(LBstatus_)
    tmpRec("CommantLab") = Trim(CommantLab_)
    '**********************------*******************************************
    tmpRec!isHospital = Trim(isHospital_)
    tmpRec!isOdour = Trim(isOdour_)
    tmpRec!Odour = Trim(Odour_)
    tmpRec!isColor = Trim(isColor_)
    tmpRec!Color = Trim(Color_)
    tmpRec!isPlasticBag = Trim(isPlasticBag_)
    tmpRec!PlasticBag = Trim(PlasticBag_)
    tmpRec!isPlasticDrum = Trim(isPlasticDrum_)
    tmpRec!PlasticDrum = Trim(PlasticDrum_)
    tmpRec!isPlasticGallon = Trim(isPlasticGallon_)
    tmpRec!PlasticGallon = Trim(PlasticGallon_)
    tmpRec!isPaperBox = Trim(isPaperBox_)
    tmpRec!PaperBox = Trim(PaperBox_)
    tmpRec!isBottle = Trim(isBottle_)
    tmpRec!Bottle = Trim(Bottle_)
    tmpRec!AnalysisType = Trim(AnalysisType_)
    tmpRec!AnalysisDesc = Trim(AnalysisDesc_)
    tmpRec!isBlending = Trim(isBlending_)
    tmpRec!Cu = Trim(Cu_)
    tmpRec!Ni = Trim(Ni_)
    tmpRec!Mn = Trim(Mn_)
    tmpRec!Zn = Trim(Zn_)
    tmpRec!Cut = Trim(Cut_)
    tmpRec!ExtAs = Trim(ExtAs_)
    tmpRec!RemarkAs = Trim(RemarkAs_)
    tmpRec!ExtCd = Trim(ExtCd_)
    tmpRec!RemarkCd = Trim(RemarkCd_)
    tmpRec!ExtCr = Trim(ExtCr_)
    tmpRec!RemarkCr = Trim(RemarkCr_)
    tmpRec!ExtCu = Trim(ExtCu_)
    tmpRec!RemarkCu = Trim(RemarkCu_)
    tmpRec!ExtPb = Trim(ExtPb_)
    tmpRec!RemarkPb = Trim(RemarkPb_)
    tmpRec!ExtNi = Trim(ExtNi_)
    tmpRec!RemarkNi = Trim(RemarkNi_)
    tmpRec!ExtZn = Trim(ExtZn_)
    tmpRec!RemarkZn = Trim(RemarkZn_)
    tmpRec!ExtHg = Trim(ExtHg_)
    tmpRec!RemarkHg = Trim(RemarkHg_)
    tmpRec!LabCode = Trim(LabCode_)
    tmpRec!Concentrate = Trim(Concentrate_)
    tmpRec!SamplingBy = Trim(SamplingBy_)
    tmpRec!SamplingByNo = Trim(SamplingByNo_)
    tmpRec!Silicon = Trim(Silicon_)
    tmpRec!SiliconRemark = Trim(SiliconRemark_)
    tmpRec!Aluminum = Trim(Aluminum_)
    tmpRec!AluminumRemark = Trim(AluminumRemark_)
    tmpRec!Iron = Trim(Iron_)
    tmpRec!IronRemark = Trim(IronRemark_)
    tmpRec!Calcium = Trim(Calcium_)
    tmpRec!CalciumRemark = Trim(CalciumRemark_)
    tmpRec!InFlammable = Trim(InFlammable_)
    
    tmpRec.Update
    tmpRec.Close
    With DBManager.GetDBConnection
        If Not colNewTreatment_ Is Nothing Then
            .Execute "DELETE FROM  tbOldTreatment  WHERE WasteDataLabID = '" & Lab_ID & "'"
            .Execute "Delete tbTreatment Where WasteDataLabID='" & Lab_ID & "'"
'            If isNew_ = False Then
                .Execute "Delete tbTreatment Where WasteDataLabID='" & Lab_ID & "'"
'                For x = 1 To colNewTreatment_.Count
'                    .Execute "Insert into tbTreatment (WasteDataLabID,TreatmentID) Values ('" & Lab_ID & "','" & colNewTreatment_(x).ID & "')"
'                Next
'            Else
'                .Execute "Delete tbTreatment_New Where WasteDataLabID='" & id_ & "'"
'                For x = 1 To colNewTreatment_.Count
'                    .Execute "Insert into tbTreatment_New (WasteDataLabID,TreatmentID) Values ('" & id_ & "','" & colNewTreatment_(x).ID & "')"
'                Next
'            End If
        End If
    End With
    Set tmpRec = Nothing
    isChanged_ = False
    Exit Sub
ErrD:
    MsgBox Err.Description & vbCrLf & "กรุณาลองใหม่หรือติดต่อผู้พัฒนาระบบ", vbCritical, "Error"
    Exit Sub
End Sub

Public Function Delete() As String
On Error GoTo delErr:
    With DBManager.GetDBConnection
        .Execute "Delete tbOldTreatment Where WasteDataLabID='" & id_ & "'"
        .Execute "Delete tbTreatment Where WasteDataLabID='" & id_ & "'"
        .Execute "Delete tbWasteDataLab Where WasteDataLabID='" & id_ & "'"
    End With
    Delete = ""
    Exit Function
delErr:
    Delete = Err.Description
End Function

Public Sub SetDataFromRecordset(selRec As ADODB.Recordset)
    id_ = selRec("WasteDataLabID")
    wasteDataID_ = Trim("" & selRec("WasteDataID"))
    generateDetail_ = Trim("" & selRec("GenerateDetail"))
    wasteChar_ = Trim("" & selRec("WasteCharacter"))
    powderOrLump_ = Trim("" & selRec("PowderOrLump"))
    isPumpAble_ = Trim("" & selRec("isPumpable"))
    isExplode_ = Trim("" & selRec("isExplode"))
    isFlameAble_ = Trim("" & selRec("isFlamable"))
    If IsNull(selRec("FlashPoint")) = False Then flashPoint_ = selRec("FlashPoint")
    isCorrosive_ = Trim("" & selRec("isCorrosive"))
    If IsNull(selRec("pHValue")) = False Then phValue_ = selRec("pHValue")
    isReact_ = Trim("" & selRec("isReact"))
    reactWith_ = Trim("" & selRec("ReactWith"))
    isCarcinogenic_ = Trim("" & selRec("isCarcinogenic"))
    carcinogenicDesc_ = Trim("" & selRec("CarcinogenicDesc"))
    isToxic_ = Trim("" & selRec("isToxic"))
    toxicType_ = Trim("" & selRec("ToxicType"))
    toxicDesc_ = Trim("" & selRec("ToxicDesc"))
    isRadioActive_ = Trim("" & selRec("isRadioActive"))
    radioActiveDesc_ = Trim("" & selRec("RadioActiveDesc"))
    isPesticide_ = Trim("" & selRec("isPesticide"))
    pesticideDesc_ = Trim("" & selRec("PesticideDesc"))
    otherDesc_ = Trim("" & selRec("OtherDesc"))
    mixOrSeperate_ = Trim("" & selRec("MixOrSeperate"))
    inOrOutDoor_ = Trim("" & selRec("InOrOutDoor"))
    storageType_ = Trim("" & selRec("StorageType"))
    storageDesc_ = Trim("" & selRec("StorageDesc"))
    oldEliminator_ = Trim("" & selRec("OldEliminator"))
    analysisNo_ = Trim("" & selRec("AnalysisNo"))
    Physicalstate_ = Trim("" & selRec("Physicalstate"))
    Viscosity_ = Trim("" & selRec("VisCosity"))
    Bulkdensity_ = Trim("" & selRec("Bulkdensity"))
    OxidizerScreen_ = Trim("" & selRec("OxidizerScreen"))
    FreeChlorine_ = Trim("" & selRec("FreeChlorine"))
    Nitrite_ = Trim("" & selRec("Nitrite"))
    Cyanide_ = Trim("" & selRec("Cyanide"))
    HeatingValue_ = Trim("" & selRec("Heatingvalue"))
    NCVValue_ = Trim("" & selRec("NCVValue"))
    MoistureContent_ = Trim("" & selRec("MoistureContent"))
    WaterContent_ = Trim("" & selRec("WaterContent"))
    AshContent_ = Trim("" & selRec("AshContent"))
    CL_ = Trim("" & selRec("CL"))
    F_ = Trim("" & selRec("F"))
    S_ = Trim("" & selRec("S"))
    SludgeContent_ = Trim("" & selRec("SludgeContent"))
    SolidContent_ = Trim("" & selRec("SolidContent"))
    HeavyAs_ = Trim("" & selRec("HeavyAs"))
    HeavyCd_ = Trim("" & selRec("HeavyCd"))
    HeavyCr_ = Trim("" & selRec("HeavyCr"))
    HeavyHg_ = Trim("" & selRec("HeavyHg"))
    HeavyPb_ = Trim("" & selRec("HeavyPb"))
    HeavyFe_ = Trim("" & selRec("HeavyFe"))
    reportNo_ = Trim("" & selRec("ReportNo"))
    samplingDate_ = Trim("" & selRec("SamplingDate"))
    ReceivedDate_ = Trim("" & selRec("ReceivedDate"))
    AnalysisDate_ = Trim("" & selRec("AnalytiCalDate"))
    SamplingBy_ = Trim("" & selRec("SamplingBy"))
    LBstatus_ = Trim("" & selRec("LBstatus"))
    CommantLab_ = Trim("" & selRec("CommantLab"))
        '******* Add for more waste storage method ************
    isHospital_ = IIf(IsNull(selRec!isHospital), "N", Trim(selRec!isHospital))
    isOdour_ = IIf(IsNull(selRec!isOdour), "N", Trim(selRec!isOdour))
    Odour_ = IIf(IsNull(selRec!Odour), "", Trim(selRec!Odour))
    isColor_ = IIf(IsNull(selRec!isColor), "N", Trim(selRec!isColor))
    Color_ = IIf(IsNull(selRec!Color), "", Trim(selRec!Color))
    isPlasticBag_ = IIf(IsNull(selRec!isPlasticBag), "N", Trim(selRec!isPlasticBag))
    PlasticBag_ = IIf(IsNull(selRec!PlasticBag), "", Trim(selRec!PlasticBag))
    isPlasticDrum_ = IIf(IsNull(selRec!isPlasticDrum), "N", Trim(selRec!isPlasticDrum))
    PlasticDrum_ = IIf(IsNull(selRec!PlasticDrum), "", Trim(selRec!PlasticDrum))
    isPlasticGallon_ = IIf(IsNull(selRec!isPlasticGallon), "N", Trim(selRec!isPlasticGallon))
    PlasticGallon_ = IIf(IsNull(selRec!PlasticGallon), "", Trim(selRec!PlasticGallon))
    isPaperBox_ = IIf(IsNull(selRec!isPaperBox), "N", Trim(selRec!isPaperBox))
    PaperBox_ = IIf(IsNull(selRec!PaperBox), "", Trim(selRec!PaperBox))
    isBottle_ = IIf(IsNull(selRec!isBottle), "N", Trim(selRec!isBottle))
    Bottle_ = IIf(IsNull(selRec!Bottle), "", Trim(selRec!Bottle))
    AnalysisType_ = IIf(IsNull(selRec!AnalysisType), 0, selRec!AnalysisType)
    AnalysisDesc_ = IIf(IsNull(selRec!AnalysisDesc), "", Trim(selRec!AnalysisDesc))
    isBlending_ = IIf(IsNull(selRec!isBlending), "", Trim(selRec!isBlending))
    CharDescOther_ = IIf(IsNull(selRec!CharDescOther), "", Trim(selRec!CharDescOther))
    InFlammable_ = IIf(IsNull(selRec!InFlammable), "N", selRec!InFlammable)
    Cu_ = Trim("" & selRec("Cu"))
    Ni_ = Trim("" & selRec("Ni"))
    Mn_ = Trim("" & selRec("Mn"))
    Zn_ = Trim("" & selRec("Zn"))
    Cut_ = Trim("" & selRec("Cut"))
    ExtAs_ = Trim("" & selRec("ExtAs"))
    RemarkAs_ = Trim("" & selRec("RemarkAs"))
    ExtCd_ = Trim("" & selRec("ExtCd"))
    RemarkCd_ = Trim("" & selRec("RemarkCd"))
    ExtCr_ = Trim("" & selRec("ExtCr"))
    RemarkCr_ = Trim("" & selRec("RemarkCr"))
    ExtCu_ = Trim("" & selRec("ExtCu"))
    RemarkCu_ = Trim("" & selRec("RemarkCu"))
    ExtPb_ = Trim("" & selRec("ExtPb"))
    RemarkPb_ = Trim("" & selRec("RemarkPb"))
    ExtNi_ = Trim("" & selRec("ExtNi"))
    RemarkNi_ = Trim("" & selRec("RemarkNi"))
    ExtZn_ = Trim("" & selRec("ExtZn"))
    RemarkZn_ = Trim("" & selRec("RemarkZn"))
    ExtHg_ = Trim("" & selRec("ExtHg"))
    RemarkHg_ = Trim("" & selRec("RemarkHg"))
    LabCode_ = Trim("" & selRec("LabCode"))
    Concentrate_ = Trim("" & selRec("Concentrate"))
    SamplingBy_ = Trim("" & selRec("SamplingBy"))
    SamplingByNo_ = Trim("" & selRec("SamplingByNo"))
    SS_ = Trim("" & selRec("ss"))
    Acidity_ = Trim("" & selRec("Acidity"))
    Alkalinity_ = Trim("" & selRec("Alkalinity"))
   Silicon_ = IIf(IsNull(selRec!Silicon), "", Trim(selRec!Silicon))
    SiliconRemark_ = IIf(IsNull(selRec!SiliconRemark), "", Trim(selRec!SiliconRemark))
    Aluminum_ = IIf(IsNull(selRec!Aluminum), "", Trim(selRec!Aluminum))
    AluminumRemark_ = IIf(IsNull(selRec!AluminumRemark), "", Trim(selRec!AluminumRemark))
    Iron_ = IIf(IsNull(selRec!Iron), "", Trim(selRec!Iron))
    IronRemark_ = IIf(IsNull(selRec!IronRemark), "", Trim(selRec!IronRemark))
   Calcium_ = IIf(IsNull(selRec!Calcium), "", Trim(selRec!Calcium))
    CalciumRemark_ = IIf(IsNull(selRec!CalciumRemark), "", Trim(selRec!CalciumRemark))
    

End Sub
