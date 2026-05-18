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
Dim jobID_ As String
Dim jobNo_ As String
Dim wasteDataID_ As String
Dim wasteName_ As String
Dim wasteNo_ As String
Dim wasteType_ As String
Dim wasteEUCode_ As String
Dim wasteQuotationNo_ As String
Dim wasteGenNo_ As String
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

Dim quotationID_ As String
Dim laborCharge_ As Double
Dim truckTypeID_ As String
Dim truckTypeName_ As String
Dim isTrailer_ As String
Dim transportFee_ As Double
Dim minWeightPerCar_ As Double
Dim transUnitID_ As String
Dim transUnitName_ As String
Dim treatmentRate_ As Double
Dim treatmentUnitID_ As String
Dim treatmentUnitName_ As String
Dim laborCount_ As Integer
Dim isPriceIncTrans_ As String

'''''This part for menifest detail
Dim disposerWeight_ As Double
Dim treatmentCharge_ As Double
Dim MenifestID_ As String
Dim menifestActQty_ As Double
Dim menifestDocNo_ As String
Dim menifestWorkDate_ As String
Dim menifestCarRegisID_ As String
'''''''''''''''''''''''
Dim chargeWType_ As String

''''This part for job fail
Dim failType_ As String
Dim failRemark_ As String
'''''''''''''''''''''''
Dim nonMinRate_ As Double
Dim hazMinRate_ As Double
Dim nblMinRate_ As Double
Dim hblMinRate_ As Double
Dim isChanged_ As Boolean

Public Property Get ID() As String
    ID = id_
End Property
Public Property Let ID(ByVal newVal As String)
    id_ = Trim(newVal)
End Property

Public Property Get JobID() As String
    JobID = jobID_
End Property
Public Property Let JobID(ByVal newVal As String)
    If Trim(jobID_) <> Trim(newVal) Then
        jobID_ = Trim(newVal)
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

Public Property Get WasteDataID() As String
    WasteDataID = wasteDataID_
End Property
Public Property Let WasteDataID(ByVal newVal As String)
    If Trim(wasteDataID_) <> Trim(newVal) Then
        wasteDataID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get WasteName() As String
    WasteName = wasteName_
End Property
Public Property Let WasteName(ByVal newVal As String)
    If Trim(wasteName_) <> Trim(newVal) Then
        wasteName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get WasteNo() As String
    WasteNo = wasteNo_
End Property
Public Property Let WasteNo(ByVal newVal As String)
    If Trim(wasteNo_) <> Trim(newVal) Then
        wasteNo_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get WasteType() As String
    WasteType = wasteType_
End Property
Public Property Let WasteType(ByVal newVal As String)
    If Trim(wasteType_) <> Trim(newVal) Then
        wasteType_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get WasteEUCode() As String
    WasteEUCode = wasteEUCode_
End Property
Public Property Let WasteEUCode(ByVal newVal As String)
    If Trim(wasteEUCode_) <> Trim(newVal) Then
        wasteEUCode_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get WasteQuotationNo() As String
    WasteQuotationNo = wasteQuotationNo_
End Property
Public Property Let WasteQuotationNo(ByVal newVal As String)
    If Trim(wasteQuotationNo_) <> Trim(newVal) Then
        wasteQuotationNo_ = Trim(newVal)
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

Public Property Get FailType() As String
    FailType = failType_
End Property
Public Property Let FailType(ByVal newVal As String)
    If Trim(failType_) <> Trim(newVal) Then
        failType_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get FailRemark() As String
    FailRemark = failRemark_
End Property
Public Property Let FailRemark(ByVal newVal As String)
    If Trim(failRemark_) <> Trim(newVal) Then
        failRemark_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get NonMinRate() As Double
    NonMinRate = nonMinRate_
End Property
Public Property Let NonMinRate(ByVal newVal As Double)
    If nonMinRate_ <> newVal Then
        nonMinRate_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get HazMinRate() As Double
    HazMinRate = hazMinRate_
End Property
Public Property Let HazMinRate(ByVal newVal As Double)
    If hazMinRate_ <> newVal Then
        hazMinRate_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get NBLMinRate() As Double
    NBLMinRate = nblMinRate_
End Property
Public Property Let NBLMinRate(ByVal newVal As Double)
    If nblMinRate_ <> newVal Then
        nblMinRate_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get HBLMinRate() As Double
    HBLMinRate = hblMinRate_
End Property
Public Property Let HBLMinRate(ByVal newVal As Double)
    If hblMinRate_ <> newVal Then
        hblMinRate_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get EstWasteQty() As Double
    EstWasteQty = estWasteQty_
End Property
Public Property Let EstWasteQty(ByVal newVal As Double)
    If estWasteQty_ <> newVal Then
        estWasteQty_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get WasteQtyUnitID() As String
    WasteQtyUnitID = wasteQtyUnitID_
End Property
Public Property Let WasteQtyUnitID(ByVal newVal As String)
    If Trim(wasteQtyUnitID_) <> Trim(newVal) Then
        wasteQtyUnitID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get WasteQtyUnitName() As String
    WasteQtyUnitName = wasteQtyUnitName_
End Property
Public Property Let WasteQtyUnitName(ByVal newVal As String)
    If Trim(wasteQtyUnitName_) <> Trim(newVal) Then
        wasteQtyUnitName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get ContainerTypeDesc() As String
    ContainerTypeDesc = containerTypeDesc_
End Property
Public Property Let ContainerTypeDesc(ByVal newVal As String)
    If Trim(containerTypeDesc_) <> Trim(newVal) Then
        containerTypeDesc_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get ContainerCount() As Double
    ContainerCount = containerCount_
End Property
Public Property Let ContainerCount(ByVal newVal As Double)
    If containerCount_ <> newVal Then
        containerCount_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get RemDesc() As String
    RemDesc = detRemDesc_
End Property
Public Property Let RemDesc(ByVal newVal As String)
    If Trim(detRemDesc_) <> Trim(newVal) Then
        detRemDesc_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get DestinationDesc() As String
    DestinationDesc = destinationDesc_
End Property
Public Property Let DestinationDesc(ByVal newVal As String)
    If Trim(destinationDesc_) <> Trim(newVal) Then
        destinationDesc_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get rowID() As Integer
    rowID = rowID_
End Property
Public Property Let rowID(ByVal newVal As Integer)
    If rowID_ <> newVal Then
        rowID_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get carNo() As Integer
    carNo = carNo_
End Property
Public Property Let carNo(ByVal newVal As Integer)
    If carNo_ <> newVal Then
        carNo_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get TSDFID() As String
    TSDFID = Trim(tsdfID_)
End Property
Public Property Let TSDFID(ByVal newVal As String)
    If Trim(tsdfID_) <> Trim(newVal) Then
        tsdfID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get TSDFName() As String
    TSDFName = Trim(tsdfName_)
End Property
Public Property Let TSDFName(ByVal newVal As String)
    If Trim(tsdfName_) <> Trim(newVal) Then
        tsdfName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get TSDFRegisNo() As String
    TSDFRegisNo = Trim(tsdfRegisNo_)
End Property
Public Property Let TSDFRegisNo(ByVal newVal As String)
    If Trim(tsdfRegisNo_) <> Trim(newVal) Then
        tsdfRegisNo_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get TSDFTumbolID() As String
    TSDFTumbolID = Trim(tsdfTumbolID_)
End Property
Public Property Let TSDFTumbolID(ByVal newVal As String)
    If Trim(tsdfTumbolID_) <> Trim(newVal) Then
        tsdfTumbolID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get DisposerWeight() As Double
    DisposerWeight = disposerWeight_
End Property
Public Property Let DisposerWeight(ByVal newVal As Double)
    If disposerWeight_ <> newVal Then
        disposerWeight_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get TreatmentCharge() As Double
    TreatmentCharge = treatmentCharge_
End Property
Public Property Let TreatmentCharge(ByVal newVal As Double)
    If treatmentCharge_ <> newVal Then
        treatmentCharge_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get MenifestID() As String
    MenifestID = MenifestID_
End Property
Public Property Let MenifestID(ByVal newVal As String)
    If Trim(MenifestID_) <> Trim(newVal) Then
        MenifestID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get MenifestActQty() As Double
    MenifestActQty = menifestActQty_
End Property
Public Property Let MenifestActQty(ByVal newVal As Double)
    If menifestActQty_ <> newVal Then
        menifestActQty_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get MenifestDocNo() As String
    MenifestDocNo = menifestDocNo_
End Property
Public Property Let MenifestDocNo(ByVal newVal As String)
    If Trim(menifestDocNo_) <> Trim(newVal) Then
        menifestDocNo_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get MenifestWorkDate() As String
    MenifestWorkDate = menifestWorkDate_
End Property
Public Property Let MenifestWorkDate(ByVal newVal As String)
    If Trim(menifestWorkDate_) <> Trim(newVal) Then
        menifestWorkDate_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get MenifestCarRegisID() As String
    MenifestCarRegisID = menifestCarRegisID_
End Property
Public Property Let MenifestCarRegisID(ByVal newVal As String)
    If Trim(menifestCarRegisID_) <> Trim(newVal) Then
        menifestCarRegisID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get ChargeWeightType() As String
    ChargeWeightType = chargeWType_
End Property
Public Property Let ChargeWeightType(ByVal newVal As String)
    If Trim(chargeWType_) <> Trim(newVal) Then
        chargeWType_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get QuotationID() As String
    QuotationID = quotationID_
End Property
Public Property Let QuotationID(ByVal newVal As String)
    If Trim(quotationID_) <> Trim(newVal) Then
        quotationID_ = Trim(newVal)
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
Public Property Let isPriceIncTrans(ByVal newVal As Boolean)
Dim tmpStr As String
    If newVal Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If isPriceIncTrans_ <> Trim(tmpStr) Then
        isPriceIncTrans_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property

Public Property Get QuoTreatmentRate() As Double
    QuoTreatmentRate = treatmentRate_
End Property
Public Property Let QuoTreatmentRate(ByVal newVal As Double)
    If treatmentRate_ <> newVal Then
        treatmentRate_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get QuoTreatmentUnitID() As String
    QuoTreatmentUnitID = treatmentUnitID_
End Property
Public Property Let QuoTreatmentUnitID(ByVal newVal As String)
    If Trim(treatmentUnitID_) <> Trim(newVal) Then
        treatmentUnitID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get QuoTreatmentUnitName() As String
    QuoTreatmentUnitName = treatmentUnitName_
End Property
Public Property Let QuoTreatmentUnitName(ByVal newVal As String)
    If Trim(treatmentUnitName_) <> Trim(newVal) Then
        treatmentUnitName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get QuoTruckTypeID() As String
    QuoTruckTypeID = truckTypeID_
End Property
Public Property Let QuoTruckTypeID(ByVal newVal As String)
    If Trim(truckTypeID_) <> Trim(newVal) Then
        truckTypeID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get QuoTruckTypeName() As String
    QuoTruckTypeName = truckTypeName_
End Property
Public Property Let QuoTruckTypeName(ByVal newVal As String)
    If Trim(truckTypeName_) <> Trim(newVal) Then
        truckTypeName_ = Trim(newVal)
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
    If Trim(isTrailer_) <> Trim(tmpStr) Then
        isTrailer_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property

Public Property Get QuoTransportFee() As Double
    QuoTransportFee = transportFee_
End Property
Public Property Let QuoTransportFee(ByVal newVal As Double)
    If transportFee_ <> newVal Then
        transportFee_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get QuoMinWeightPerCar() As Double
    QuoMinWeightPerCar = minWeightPerCar_
End Property
Public Property Let QuoMinWeightPerCar(ByVal newVal As Double)
    If minWeightPerCar_ <> newVal Then
        minWeightPerCar_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get QuoTransUnitID() As String
    QuoTransUnitID = transUnitID_
End Property
Public Property Let QuoTransUnitID(ByVal newVal As String)
    If Trim(transUnitID_) <> Trim(newVal) Then
        transUnitID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get QuoTransUnitName() As String
    QuoTransUnitName = transUnitName_
End Property
Public Property Let QuoTransUnitName(ByVal newVal As String)
    If Trim(transUnitName_) <> Trim(newVal) Then
        transUnitName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get QuoLaborCount() As Integer
    QuoLaborCount = laborCount_
End Property
Public Property Let QuoLaborCount(ByVal newVal As Integer)
    If laborCount_ <> newVal Then
        laborCount_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get QuoLaborCharge() As Integer
    QuoLaborCharge = laborCharge_
End Property
Public Property Let QuoLaborCharge(ByVal newVal As Integer)
    If laborCharge_ <> newVal Then
        laborCharge_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get isChanged() As Boolean
    isChanged = isChanged_
End Property
Public Property Let isChanged(ByVal newVal As Boolean)
    isChanged_ = newVal
End Property

Public Sub Save()
Dim tmpRec As ADODB.Recordset
Dim tmpCon As ADODB.Connection
    Set tmpRec = New ADODB.Recordset
    Set tmpCon = DBConn
    tmpRec.Open "Select * from tbJobDataDetail Where JobDetID='" & id_ & "'", tmpCon, adOpenDynamic, adLockOptimistic
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
    
    tmpRec("QuotationID") = Trim(quotationID_)
    tmpRec("ExtraLaborFee") = laborCharge_
    tmpRec("TruckTypeID") = Trim(truckTypeID_)
    tmpRec("TransportFee") = transportFee_
    tmpRec("MinWeightPerCar") = minWeightPerCar_
    tmpRec("TransferUnitID") = Trim(transUnitID_)
    tmpRec("TreatmentRate") = treatmentRate_
    tmpRec("TreatmentUnitID") = Trim(treatmentUnitID_)
    tmpRec("LaborCount") = laborCount_
    tmpRec("isPriceIncTrans") = Trim(isPriceIncTrans_)
    
    tmpRec("NonMinRate") = nonMinRate_
    tmpRec("HazMinRate") = hazMinRate_
    tmpRec("NBLMinRate") = nblMinRate_
    tmpRec("HBLMinRate") = hblMinRate_
    
    tmpRec.Update
    tmpRec.Close
    
    Set tmpCon = Nothing
    Set tmpRec = Nothing
    isChanged_ = False
End Sub

Public Function Delete() As String
On Error GoTo delErr:
    DBConnExc "Delete tbJobDataDetail Where JobDetID='" & id_ & "'"
    Delete = ""
    Exit Function
delErr:
    Delete = Err.Description
End Function

Public Sub SetDataFromRecordset(selRec As ADODB.Recordset)

    id_ = Trim("" & selRec("JobDetID"))
    jobID_ = Trim("" & selRec("JobID"))
    jobNo_ = Trim("" & selRec("JobNo"))
    wasteDataID_ = Trim("" & selRec("WasteDataID"))
    wasteName_ = Trim("" & selRec("WasteName"))
    wasteEUCode_ = Trim("" & selRec("EUCode"))
    wasteNo_ = Trim("" & selRec("WasteNo"))
    wasteGenNo_ = Trim("" & selRec("WasteCreateNo"))
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
    If isExist(selRec, "ChargeWType") Then chargeWType_ = Trim("" & selRec("ChargeWType"))
    If isExist(selRec, "DocNo") Then menifestDocNo_ = Trim("" & selRec("DocNo"))
    If isExist(selRec, "WorkDate") Then menifestWorkDate_ = Trim("" & selRec("WorkDate"))
    If isExist(selRec, "CarRegisID") Then menifestCarRegisID_ = Trim("" & selRec("CarRegisID"))
    
    If isExist(selRec, "DisposerWeight") Then
        If IsNull(selRec("DisposerWeight")) = False Then disposerWeight_ = selRec("DisposerWeight")
    End If
    If isExist(selRec, "TreatmentCharge") Then
        If IsNull(selRec("TreatmentCharge")) = False Then treatmentCharge_ = selRec("TreatmentCharge")
    End If
    If IsNull(selRec("TreatmentRate")) = False Then treatmentRate_ = selRec("TreatmentRate")
    If IsNull(selRec("TransportFee")) = False Then transportFee_ = selRec("TransportFee")
    If IsNull(selRec("MinWeightPerCar")) = False Then minWeightPerCar_ = selRec("MinWeightPerCar")
    
    treatmentUnitID_ = Trim("" & selRec("TreatmentUnitID"))
    treatmentUnitName_ = Trim("" & selRec("TreatmentUnitID"))   'Trim("" & selRec("TreatmentUnitName"))
    
    truckTypeID_ = Trim("" & selRec("TruckTypeID"))
    truckTypeName_ = Trim("" & selRec("TruckTypeDesc"))
    isTrailer_ = Trim("" & selRec("isTrailer"))
    transUnitID_ = Trim("" & selRec("TrasnferUnitID"))
    transUnitName_ = Trim("" & selRec("TrasnferUnitID"))    'Trim("" & selRec("TransUnitName"))
    If IsNull(selRec("TransportFee")) = False Then transportFee_ = selRec("TransportFee")
    If IsNull(selRec("MinWeightPerCar")) = False Then minWeightPerCar_ = selRec("MinWeightPerCar")
    isPriceIncTrans_ = Trim("" & selRec("isPriceIncTransport"))
    quotationID_ = Trim("" & selRec("QuotationID"))
    If IsNull(selRec("StaffCountPerCar")) = False Then laborCount_ = selRec("StaffCountPerCar")
    If IsNull(selRec("ExtraLaborFee")) = False Then laborCharge_ = selRec("ExtraLaborFee")
    If isExist(selRec, "FailType") Then failType_ = Trim("" & selRec("FailType"))
    If isExist(selRec, "FailRemark") Then failRemark_ = Trim("" & selRec("FailRemark"))

    nonMinRate_ = selRec("NonMinRate")
    hazMinRate_ = selRec("HazMinRate")
    nblMinRate_ = selRec("NBLMinRate")
    hblMinRate_ = selRec("HBLMinRate")
End Sub

