VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BWGMenifestWasteAdded"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Class นี้ใช้สำหรับ Waste ที่ add เข้าไปใหม่ใน Menifest นั้น
Dim id_ As String
Dim MenifestID_ As String
Dim wasteDataID_ As String
Dim WasteName_ As String
Dim wasteNo_ As String
Dim wasteType_ As String
Dim wasteEUCode_ As String
Dim WasteGenNo_ As String
Dim containerCount_ As Double
Dim containerDesc_ As String
Dim CustWeight_ As Double
Dim DisposerWeight_ As Double
Dim treatmentCharge_ As Double
Dim QuotationID_ As String
Dim QuotationNo_ As String
Dim transportFee_ As Double
Dim transferUnitID_ As String
Dim TreatmentRate_ As Double
Dim TreatmentUnitID_ As String
Dim isPriceIncTrans_ As String
Dim IsDiscountPriceTrans_ As String
Dim minWeightPerCar_ As Double
Dim rowID_ As Integer
Dim laborCount_ As Integer
Dim extraLaborFee_ As Double
Dim chargeWeightType_ As String
Dim nonMinRate_ As Double
Dim hazMinRate_ As Double
Dim nblMinRate_ As Double
Dim hblMinRate_ As Double
Dim isChanged_ As Boolean
Dim IsNewWaste_ As String
Dim InWeight_ As Double
Dim IsNewWasteCR_ As String
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

Public Property Get IsNewWasteCR() As String
    IsNewWasteCR = IsNewWasteCR_
End Property
Public Property Let IsNewWasteCR(ByVal newval As String)
    If Trim(IsNewWasteCR_) <> Trim(newval) Then
        IsNewWasteCR_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get id() As String
    id = id_
End Property
Public Property Let id(ByVal newval As String)
    id_ = Trim(newval)
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

Public Property Get WasteGenNo() As String
    WasteGenNo = WasteGenNo_
End Property
Public Property Let WasteGenNo(ByVal newval As String)
    If Trim(WasteGenNo_) <> Trim(newval) Then
        WasteGenNo_ = Trim(newval)
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

Public Property Get ContainerTypeDesc() As String
    ContainerTypeDesc = containerDesc_
End Property
Public Property Let ContainerTypeDesc(ByVal newval As String)
    If Trim(containerDesc_) <> Trim(newval) Then
        containerDesc_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get EstWasteQty() As Double
    EstWasteQty = CustWeight_
End Property
Public Property Let EstWasteQty(ByVal newval As Double)
    If CustWeight_ <> newval Then
        CustWeight_ = newval
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

Public Property Get QuotationID() As String
    QuotationID = QuotationID_
End Property
Public Property Let QuotationID(ByVal newval As String)
    If Trim(QuotationID_) <> Trim(newval) Then
        QuotationID_ = Trim(newval)
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

Public Property Get TransportFee() As Double
    TransportFee = transportFee_
End Property
Public Property Let TransportFee(ByVal newval As Double)
    If transportFee_ <> newval Then
        transportFee_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get TransferUnitID() As String
    TransferUnitID = transferUnitID_
End Property
Public Property Let TransferUnitID(ByVal newval As String)
    If Trim(transferUnitID_) <> Trim(newval) Then
        transferUnitID_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get QuoTreatmentRate() As Double
    QuoTreatmentRate = TreatmentRate_
End Property
Public Property Let QuoTreatmentRate(ByVal newval As Double)
    If TreatmentRate_ <> newval Then
        TreatmentRate_ = newval
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

Public Property Get ChargeWeightType() As String
    ChargeWeightType = chargeWeightType_
End Property
Public Property Let ChargeWeightType(ByVal newval As String)
    If Trim(chargeWeightType_) <> Trim(newval) Then
        chargeWeightType_ = Trim(newval)
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

Public Property Get QuoTreatmentUnitName() As String
    QuoTreatmentUnitName = TreatmentUnitID_
End Property
Public Property Let QuoTreatmentUnitName(ByVal newval As String)
    If Trim(TreatmentUnitID_) <> Trim(newval) Then
        TreatmentUnitID_ = Trim(newval)
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

Public Property Get MinWeightPerCar() As Double
    MinWeightPerCar = minWeightPerCar_
End Property
Public Property Let MinWeightPerCar(ByVal newval As Double)
    If minWeightPerCar_ <> newval Then
        minWeightPerCar_ = newval
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

Public Property Get RowID() As Integer
    RowID = rowID_
End Property
Public Property Let RowID(ByVal newval As Integer)
    If rowID_ <> newval Then
        rowID_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get LaborCount() As Integer
    LaborCount = laborCount_
End Property
Public Property Let LaborCount(ByVal newval As Integer)
    If laborCount_ <> newval Then
        laborCount_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get ExtraLaborFee() As Integer
    ExtraLaborFee = extraLaborFee_
End Property
Public Property Let ExtraLaborFee(ByVal newval As Integer)
    If extraLaborFee_ <> newval Then
        extraLaborFee_ = newval
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
Dim sqlStr As String
    sqlStr = "SELECT *, CONVERT(DATETIME, GETDATE(), 103) AS GETDATE FROM tbMenifestWasteAdded WHERE WasteAddID='" & id_ & "'"
    Set tmpRec = New ADODB.Recordset
    GetDBConnection
    tmpRec.open sqlStr, DBConn, adOpenDynamic, adLockOptimistic
    If tmpRec.EOF Then
        tmpRec.AddNew
        tmpRec("WasteAddID") = id_
    End If
    tmpRec("MenifestID") = Trim(MenifestID_)
    tmpRec("WasteDataID") = Trim(wasteDataID_)
    tmpRec("ContainerCount") = Trim(containerCount_)
    tmpRec("ContainerDesc") = Trim(containerDesc_)
    tmpRec("CustWeight") = Trim(CustWeight_)
    tmpRec("DisposerWeight") = Trim(DisposerWeight_)
    tmpRec("TreatmentCharge") = Trim(treatmentCharge_)
    tmpRec("QuotationID") = Trim(QuotationID_)
    tmpRec("TransportFee") = Trim(transportFee_)
    tmpRec("TransferUnitID") = Trim(transferUnitID_)
    tmpRec("TreatmentRate") = Trim(TreatmentRate_)
    tmpRec("TreatmentUnitID") = Trim(TreatmentUnitID_)
    tmpRec("isPriceIncTrans") = Trim(isPriceIncTrans_)
    tmpRec("IsDiscountPriceTrans") = Trim(IsDiscountPriceTrans_)
    tmpRec("MinWeightPerCar") = Trim(minWeightPerCar_)
    tmpRec("LaborCount") = Trim(laborCount_)
    tmpRec("ExtraLaborFee") = Trim(extraLaborFee_)
    tmpRec("RowID") = Trim(rowID_)
    tmpRec("NonMinRate") = nonMinRate_
    tmpRec("HazMinRate") = hazMinRate_
    tmpRec("NBLMinRate") = nblMinRate_
    tmpRec("HBLMinRate") = hblMinRate_
    tmpRec("IsNewWaste") = IsNewWaste_
    tmpRec("IsNewWasteCR") = IsNewWasteCR_
    tmpRec("InWeight") = Trim(InWeight_)
    tmpRec!LastUser = Trim(Trim(CurrentUser.StaffFName) & " " & Trim(CurrentUser.StaffLName))
    tmpRec!LastUpdate = GetSystemDateTime.SystemDateTime
    tmpRec("CanEvap") = Trim(CanEvap_)
    tmpRec("EvapDesc") = Trim(EvapDesc_)
    tmpRec("Precipitation") = Trim(Precipitation_)
    tmpRec("PrecipitationDesc") = Trim(PrecipitationDesc_)
    tmpRec("Neutralization") = Trim(Neutralization_)
    tmpRec("NeutralizationDesc") = Trim(NeutralizationDesc_)
    tmpRec("CommentWaste") = Trim(CommentWaste_)
    tmpRec("CommentWasteDesc") = Trim(CommentWasteDesc_)
    
    
    
    
    tmpRec.Update
'    tmpRec.Close
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
    DBConnExc "Delete tbMenifestWasteAdded WHERE WasteAddID='" & id_ & "'"
    Delete = ""
    Exit Function
delErr:
    Delete = err.Description
End Function

Public Sub SetDataFromRecordset(selRec As ADODB.Recordset)
     id_ = "" & selRec("WasteAddID")
    MenifestID_ = "" & selRec("MenifestID")
    wasteDataID_ = "" & selRec("WasteDataID")
    WasteName_ = "" & selRec("WasteName")
    wasteNo_ = "" & selRec("WasteNo")
    wasteType_ = "" & selRec("WasteType")
    WasteGenNo_ = "" & selRec("WasteCreateNo")
    wasteEUCode_ = Trim("" & selRec("EUCode"))
    containerCount_ = selRec("ContainerCount")
    containerDesc_ = "" & selRec("ContainerDesc")
    CustWeight_ = selRec("CustWeight")
    DisposerWeight_ = selRec("DisposerWeight")
    treatmentCharge_ = selRec("TreatmentCharge")
    QuotationID_ = "" & selRec("QuotationID")
    QuotationNo_ = "" & selRec("QuotationNo")
    transportFee_ = selRec("TransportFee")
    transferUnitID_ = "" & selRec("TransferUnitID")
    TreatmentRate_ = selRec("TreatmentRate")
    TreatmentUnitID_ = "" & selRec("TreatmentUnitID")
    isPriceIncTrans_ = "" & selRec("isPriceIncTrans")
    IsDiscountPriceTrans_ = "" & selRec("IsDiscountPriceTrans")
    chargeWeightType_ = "" & selRec("ChargeWType")
    minWeightPerCar_ = selRec("MinWeightPerCar")
    laborCount_ = selRec("LaborCount")
    extraLaborFee_ = selRec("ExtraLaborFee")
    rowID_ = selRec("RowID")
    nonMinRate_ = selRec("NonMinRate")
    hazMinRate_ = selRec("HazMinRate")
    nblMinRate_ = selRec("NBLMinRate")
    hblMinRate_ = selRec("HBLMinRate")
    If IsNull(selRec("IsNewWaste")) = False Then IsNewWaste_ = selRec("IsNewWaste")
    If IsNull(selRec("IsNewWasteCR")) = False Then IsNewWasteCR_ = selRec("IsNewWasteCR")
    If IsNull(selRec("InWeight")) = False Then InWeight_ = selRec("InWeight")
    
    CanEvap_ = Trim("" & selRec("CanEvap"))
    EvapDesc_ = Trim("" & selRec("EvapDesc"))
    Precipitation_ = Trim("" & selRec("Precipitation"))
    PrecipitationDesc_ = Trim("" & selRec("PrecipitationDesc"))
    Neutralization_ = Trim("" & selRec("Neutralization"))
    NeutralizationDesc_ = Trim("" & selRec("NeutralizationDesc"))
    CommentWaste_ = Trim("" & selRec("CommentWaste"))
    CommentWasteDesc_ = Trim("" & selRec("CommentWasteDesc"))

    
End Sub
