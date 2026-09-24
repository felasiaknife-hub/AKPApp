VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BWGJobDataCarUsed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim id_ As String
Dim jobID_ As String
Dim truckSubTypeID_ As String
Dim truckSubTypeName_ As String
Dim truckTypeID_ As String
Dim truckTypeName_ As String
Dim truckCount_ As Integer
Dim workerPerCar_ As Integer
Dim rowID_ As Integer
Dim CarRem_ As String
Dim destinationDesc_ As String
Dim goSameTime_ As String
Dim takePhoto_ As String
Dim carNo_ As Integer
Dim isTrailer_ As String
Dim canBeTrailer_ As String
Dim vendorTripCount_ As Integer
Dim colTimeTable_ As Collection
Dim menifestCount_ As Integer
Dim isChanged_ As Boolean
Dim TankType_ As String
Dim Tankcount_ As Integer
Dim TankTypeB_ As Integer
Dim TankTypeS_ As String
Dim TankTypeSCount_ As Integer
Dim TankTypeSB_ As Integer
Dim TankPrice_ As Integer
Dim BagType_ As String
Dim BagCount_ As Integer
Dim BagPrice_ As Integer
Dim TankPriceV_ As Integer
Dim TankPriceSV_ As Integer
Dim TankPriceS_ As Integer
Dim BagPriceV_ As Integer
Dim Container_ As Integer
Dim BagBack_ As Integer
'เพิ่มเติม method สำหรับการจัดการบ็อก
 Dim BoxType_ As String
 Dim BoxCount_ As Integer
 Dim PhotoS_ As String
 
 Dim Lack_ As Integer
 Dim PaperBox_ As Integer
 Dim BoutS_ As Integer

 Dim PhotoOnWeight_ As String
 Dim PhotoOnBelt_ As String
 Dim PhotoOnBoxfeed_ As String
 Dim CustPhoto_ As String
 Dim CustCon_ As String
 Dim CloseWeight_ As String
 Dim ConName_ As String
 Dim IsChk_ As String
 Dim EmpTrain_ As String
 Dim CarCheck_ As String
Dim AKPBoxCode_ As String
Dim AKPBoxCode2_ As String
 Dim NewCustJob_ As String
 Dim NewCustJobBI_ As String
 Dim LackNew_ As Integer
 Dim LackGas_ As Integer
 Dim Plate_ As Integer
 Dim Plastic_ As Integer
 Dim BoutCage_ As Integer
 Dim CovidBin_ As Integer
 Dim CovidBin200_ As Integer
 Dim Rack_ As Integer
 Dim TruckRateID_ As String

Public Property Get id() As String
    id = id_
End Property
Public Property Let id(ByVal newval As String)
    id_ = Trim(newval)
End Property

Public Property Get TruckRateID() As String
    TruckRateID = TruckRateID_
End Property
Public Property Let TruckRateID(ByVal newval As String)
    If Trim(jobID_) <> Trim(newval) Then
        TruckRateID_ = Trim(newval)
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
    TruckTypeID = truckTypeID_
End Property
Public Property Let TruckTypeID(ByVal newval As String)
    If Trim(truckTypeID_) <> Trim(newval) Then
        truckTypeID_ = Trim(newval)
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

Public Property Get TruckCount() As Integer
    TruckCount = truckCount_
End Property
Public Property Let TruckCount(ByVal newval As Integer)
    If truckCount_ <> newval Then
        truckCount_ = newval
        isChanged_ = True
    End If
End Property
Public Property Get PhotoS() As Integer
    PhotoS = PhotoS_
End Property
Public Property Let PhotoS(ByVal newval As Integer)
    If PhotoS_ <> newval Then
        PhotoS_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get WorkerPerCar() As Integer
    WorkerPerCar = workerPerCar_
End Property
Public Property Let WorkerPerCar(ByVal newval As Integer)
    If workerPerCar_ <> newval Then
        workerPerCar_ = newval
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

Public Property Get Remark() As String
    Remark = CarRem_
End Property
Public Property Let Remark(ByVal newval As String)
    If Trim(CarRem_) <> Trim(newval) Then
        CarRem_ = Trim(newval)
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

Public Property Get GoSameTime() As String
    GoSameTime = goSameTime_
End Property
Public Property Let GoSameTime(ByVal newval As String)
    If Trim(goSameTime_) <> Trim(newval) Then
        goSameTime_ = Trim(newval)
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

Public Property Get carNo() As Integer
    carNo = carNo_
End Property
Public Property Let carNo(ByVal newval As Integer)
    If Trim(carNo_) <> Trim(newval) Then
        carNo_ = Trim(newval)
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

Public Property Get CanBeTrailer() As Boolean
    If Trim(canBeTrailer_) = "Y" Then
        CanBeTrailer = True
    Else
        CanBeTrailer = False
    End If
End Property
Public Property Let CanBeTrailer(ByVal newval As Boolean)
Dim tmpStr As String
    If newval Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(canBeTrailer_) <> Trim(tmpStr) Then
        canBeTrailer_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property

Public Property Get VendorTripCount() As Integer
    VendorTripCount = vendorTripCount_
End Property
Public Property Let VendorTripCount(ByVal newval As Integer)
    If Trim(vendorTripCount_) <> Trim(newval) Then
        vendorTripCount_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get MenifestCount() As Integer
    MenifestCount = menifestCount_
End Property
Public Property Let MenifestCount(ByVal newval As Integer)
    If menifestCount_ <> newval Then
        menifestCount_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get TimeTables() As Collection
Dim tmpS As New BWGSearchManager
    If colTimeTable_ Is Nothing Then
        Set colTimeTable_ = tmpS.JobDataTimeTableSearch("JobDataCarID='" & id_ & "'")
    End If
    Set TimeTables = colTimeTable_
End Property
Public Property Set TimeTables(ByVal newval As Collection)
    Set colTimeTable_ = Trim(newval)
    isChanged_ = True
End Property

Public Sub RequeryTimeTable()
    Set colTimeTable_ = Nothing
End Sub

Public Property Get BoxType() As String
    BoxType = BoxType_
End Property
Public Property Let BoxType(ByVal newval As String)
    If BoxType_ <> newval Then
        BoxType_ = newval
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
Public Property Get TankType() As String
    TankType = TankType_
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
Public Property Get Container() As Integer
    Container = Container_
End Property
Public Property Let Container(ByVal newval As Integer)
    If Container_ <> newval Then
        Container_ = newval
        isChanged_ = True
    End If
End Property
Public Property Get BagBack() As Integer
    BagBack = BagBack_
End Property
Public Property Let BagBack(ByVal newval As Integer)
    If BagBack_ <> newval Then
        BagBack_ = newval
        isChanged_ = True
    End If
End Property
Public Property Get LackNew() As Integer
    LackNew = LackNew_
End Property
Public Property Let LackNew(ByVal newval As Integer)
    If LackNew_ <> newval Then
        LackNew_ = newval
        isChanged_ = True
    End If
End Property
Public Property Get LackGas() As Integer
    LackGas = LackGas_
End Property
Public Property Let LackGas(ByVal newval As Integer)
    If LackGas_ <> newval Then
        LackGas_ = newval
        isChanged_ = True
    End If
End Property
Public Property Get Plate() As Integer
    Plate = Plate_
End Property
Public Property Let Plate(ByVal newval As Integer)
    If Plate_ <> newval Then
        Plate_ = newval
        isChanged_ = True
    End If
End Property
Public Property Get Plastic() As Integer
    Plastic = Plastic_
End Property
Public Property Let Plastic(ByVal newval As Integer)
    If Plastic_ <> newval Then
        Plastic_ = newval
        isChanged_ = True
    End If
End Property
Public Property Get BoutCage() As Integer
    BoutCage = BoutCage_
End Property
Public Property Let BoutCage(ByVal newval As Integer)
    If BoutCage_ <> newval Then
        BoutCage_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get Lack() As Integer
    Lack = Lack_
End Property
Public Property Let Lack(ByVal newval As Integer)
    If Lack_ <> newval Then
        Lack_ = newval
        isChanged_ = True
    End If
End Property
Public Property Get PaperBox() As Integer
    PaperBox = PaperBox_
End Property
Public Property Let PaperBox(ByVal newval As Integer)
    If PaperBox_ <> newval Then
        PaperBox_ = newval
        isChanged_ = True
    End If
End Property
Public Property Get BoutS() As Integer
    BoutS = BoutS_
End Property
Public Property Let BoutS(ByVal newval As Integer)
    If BoutS_ <> newval Then
        BoutS_ = newval
        isChanged_ = True
    End If
End Property
Public Property Get CovidBin() As Integer
    CovidBin = CovidBin_
End Property
Public Property Let CovidBin(ByVal newval As Integer)
    If CovidBin_ <> newval Then
        CovidBin_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get CovidBin200() As Integer
    CovidBin200 = CovidBin200_
End Property
Public Property Let CovidBin200(ByVal newval As Integer)
    If CovidBin200_ <> newval Then
        CovidBin200_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get Rack() As Integer
    Rack = Rack_
End Property
Public Property Let Rack(ByVal newval As Integer)
    If Rack_ <> newval Then
        Rack_ = newval
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
Public Property Get ConName() As String
    ConName = ConName_
End Property
Public Property Let ConName(ByVal newval As String)
    If Trim(ConName_) <> Trim(newval) Then
        ConName_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get IsChk() As String
    IsChk = IsChk_
End Property
Public Property Let IsChk(ByVal newval As String)
    If Trim(IsChk_) <> Trim(newval) Then
        IsChk_ = Trim(newval)
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
Public Property Get PhotoOnBoxfeed() As String
    PhotoOnBoxfeed = PhotoOnBoxfeed_
End Property
Public Property Let PhotoOnBoxfeed(ByVal newval As String)
    If Trim(PhotoOnBoxfeed_) <> Trim(newval) Then
        PhotoOnBoxfeed_ = Trim(newval)
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


Public Property Get NewCustJobBI() As String
    NewCustJobBI = NewCustJobBI_
End Property
Public Property Let NewCustJobBI(ByVal newval As String)
    If Trim(NewCustJobBI_) <> Trim(newval) Then
        NewCustJobBI_ = Trim(newval)
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
Public Property Get AKPBoxCode() As String
    AKPBoxCode = AKPBoxCode_
End Property
Public Property Let AKPBoxCode(ByVal newval As String)
    If Trim(AKPBoxCode_) <> Trim(newval) Then
        AKPBoxCode_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get AKPBoxCode2() As String
    AKPBoxCode2 = AKPBoxCode2_
End Property
Public Property Let AKPBoxCode2(ByVal newval As String)
    If Trim(AKPBoxCode2_) <> Trim(newval) Then
        AKPBoxCode2_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get isChanged() As Boolean
    isChanged = isChanged_
End Property
Public Property Let isChanged(ByVal newval As Boolean)
    isChanged_ = newval
End Property

Public Sub Save(Optional WorkDate As String = "", Optional UpdateType As String = "NORMAL")
On Error GoTo ErrD
Dim tmpRec As ADODB.Recordset
Dim x%
Dim sqlStr As String
    sqlStr = "SELECT * FROM tbJobDataCarUsed WHERE JobDataCarID='" & id_ & "'"
    Set tmpRec = New ADODB.Recordset
    GetDBConnection
    tmpRec.open sqlStr, DBConn, adOpenDynamic, adLockOptimistic
    If tmpRec.EOF Then
        tmpRec.AddNew
        tmpRec("JobDataCarID") = id_
    End If
    tmpRec("JobID") = Trim(jobID_)
    tmpRec("TruckSubTypeID") = Trim(truckSubTypeID_)
    tmpRec("TruckCount") = Trim(truckCount_)
    tmpRec("WorkerPerCar") = Trim(workerPerCar_)
    tmpRec("RowID") = Trim(rowID_)
    tmpRec("CarRem") = Trim(CarRem_)
    tmpRec("DestinationDesc") = Trim(destinationDesc_)
    tmpRec("GoSameTime") = Trim(goSameTime_)
    tmpRec("TakePhoto") = Trim(takePhoto_)
    tmpRec("CarNo") = Trim(carNo_)
    tmpRec("MarkForDelete") = "N"
    tmpRec("isTrailer") = Trim(isTrailer_)
    tmpRec("VendorTripCount") = vendorTripCount_
    tmpRec("BoxType") = BoxType_
    tmpRec("BoxCount") = BoxCount_
    tmpRec("TankType") = TankType_
    tmpRec("Tankcount") = Tankcount_
    tmpRec("TankTypeB") = TankTypeB_
    tmpRec("TankTypeS") = TankTypeS_
    tmpRec("TankTypeSCount") = TankTypeSCount_
    tmpRec("TankTypeSB") = TankTypeSB_
    tmpRec("TankPrice") = TankPrice_
    tmpRec("BagType") = BagType_
    tmpRec("BagCount") = BagCount_
    tmpRec("BagPrice") = BagPrice_
    tmpRec("TankPriceV") = TankPriceV_
    tmpRec("TankPriceSV") = TankPriceSV_
     tmpRec("TankPriceS") = TankPriceS_
    tmpRec("BagPriceV") = BagPriceV_
    tmpRec("Container") = Container_
    tmpRec("BagBack") = BagBack_
    tmpRec("Lack") = Lack_
    tmpRec("PaperBox") = PaperBox_
    tmpRec("BoutS") = BoutS_
    tmpRec("CovidBin") = CovidBin_
    tmpRec("CovidBin200") = CovidBin200_
    tmpRec("Rack") = Rack_
    tmpRec("PhotoOnWeight") = PhotoOnWeight_
    tmpRec("PhotoOnBelt") = PhotoOnBelt_
    tmpRec("PhotoOnBoxfeed") = PhotoOnBoxfeed_
    tmpRec("NewCustJob") = NewCustJob_
    tmpRec("NewCustJobBI") = NewCustJobBI_
    tmpRec("CustPhoto") = CustPhoto_
    tmpRec("CustCon") = CustCon_
    tmpRec("Closeweight") = CloseWeight_
    tmpRec("ConName") = ConName_
    tmpRec("IsChk") = IsChk_
    tmpRec("EmpTrain") = EmpTrain_
    tmpRec("CarCheck") = CarCheck_
    tmpRec("LackNew") = LackNew_
    tmpRec("LackGas") = LackGas_
    tmpRec("Plate") = Plate_
    tmpRec("Plastic") = Plastic_
    tmpRec("BoutCage") = BoutCage_
    tmpRec("PhotoS") = PhotoS_
    tmpRec("TruckRateID") = TruckRateID_
    If UpdateType = "COPY" Then
        tmpRec("BoxType") = ""
        tmpRec("BoxCount") = "0"
    End If
    tmpRec.Update
    
    'Update BoxReserve
    If BoxType_ <> "" Then
        If BoxCount_ > 0 Then
            sqlStr = "SELECT * FROM BWG_BOX.dbo.Box_BoxReserve WHERE JobDataCarID = '" & id_ & "' AND TruckSubTypeID = '" & Trim(truckSubTypeID_) & "'"
            Set tmpRec = New ADODB.Recordset
            tmpRec.open sqlStr, DBConStock, adOpenDynamic, adLockOptimistic
            If tmpRec.EOF Then
                tmpRec.AddNew
                tmpRec("JobDataCarID") = id_
                tmpRec("TruckSubTypeID") = Trim(truckSubTypeID_)
            End If
            tmpRec("BoxType") = BoxType_
            tmpRec("BoxReservNumber") = BoxCount_
            tmpRec.Update
            Set tmpRec = Nothing
        End If
    End If
'    tmpRec.Close
'    With DBConn
        DBConnExc "Update tbJobDataTimeTable Set MarkForDelete='Y' WHERE JobDataCarID='" & id_ & "'"
        If Not colTimeTable_ Is Nothing Then
            For x = 1 To colTimeTable_.Count
                colTimeTable_(x).WorkerCount = workerPerCar_
                colTimeTable_(x).Save WorkDate
            Next
        End If
        DBConnExc "Delete tbJobDataTimeTable WHERE MarkForDelete='Y' And JobDataCarID='" & id_ & "'"
'    End With
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
    DBConnExc "DELETE FROM BWG_BOX.dbo.Box_BoxReserve WHERE JobDataCarID='" & id_ & "'"
    DBConnExc "Delete tbJobDataCarUsed WHERE JobDataCarID='" & id_ & "'"
    Delete = ""
    Exit Function
delErr:
    Delete = err.Description
End Function

Public Sub SetDataFromRecordset(selRec As ADODB.Recordset)
    id_ = Trim("" & selRec("JobDataCarID"))
    jobID_ = Trim("" & selRec("JobID"))
    truckSubTypeID_ = Trim("" & selRec("TruckSubTypeID"))
    truckSubTypeName_ = Trim("" & selRec("SubTypeDesc"))
    truckTypeID_ = Trim("" & selRec("TruckTypeID"))
    truckTypeName_ = Trim("" & selRec("TruckTypeDesc"))
    truckCount_ = selRec("TruckCount")
    workerPerCar_ = selRec("WorkerPerCar")
    rowID_ = selRec("RowID")
    CarRem_ = Trim("" & selRec("CarRem"))
    destinationDesc_ = Trim("" & selRec("DestinationDesc"))
    goSameTime_ = Trim("" & selRec("GoSameTime"))
    takePhoto_ = Trim("" & selRec("TakePhoto"))
    isTrailer_ = Trim("" & selRec("isTrailer"))
    canBeTrailer_ = Trim("" & selRec("CanBeTrailer"))
    vendorTripCount_ = selRec("VendorTripCount")
    If IsNull(selRec("CarNo")) = False Then carNo_ = selRec("CarNo")
    If IsNull(selRec("MenifestCount")) = False Then menifestCount_ = selRec("MenifestCount")
    If IsNull(selRec("BoxType")) = False Then BoxType_ = Trim("" & selRec("BoxType"))
    If IsNull(selRec("BoxCount")) = False Then BoxCount_ = Trim("" & selRec("BoxCount"))
    If IsNull(selRec("TankType")) = False Then TankType_ = Trim("" & selRec("TankType"))
    If IsNull(selRec("Tankcount")) = False Then Tankcount_ = Trim("" & selRec("Tankcount"))
    If IsNull(selRec("TankTypeB")) = False Then TankTypeB_ = Trim("" & selRec("TankTypeB"))
    If IsNull(selRec("TankTypeS")) = False Then TankTypeS_ = Trim("" & selRec("TankTypeS"))
    If IsNull(selRec("TankTypeSCount")) = False Then TankTypeSCount_ = Trim("" & selRec("TankTypeSCount"))
    If IsNull(selRec("TankTypeSB")) = False Then TankTypeSB_ = Trim("" & selRec("TankTypeSB"))
    If IsNull(selRec("TankPrice")) = False Then TankPrice_ = Trim("" & selRec("TankPrice"))
    If IsNull(selRec("BagType")) = False Then BagType_ = Trim("" & selRec("BagType"))
    If IsNull(selRec("BagCount")) = False Then BagCount_ = Trim("" & selRec("BagCount"))
    If IsNull(selRec("BagPrice")) = False Then BagPrice_ = Trim("" & selRec("BagPrice"))
    If IsNull(selRec("TankPriceV")) = False Then TankPriceV_ = Trim("" & selRec("TankPriceV"))
    If IsNull(selRec("TankPriceSV")) = False Then TankPriceSV_ = Trim("" & selRec("TankPriceSV"))
    If IsNull(selRec("TankPriceS")) = False Then TankPriceS_ = Trim("" & selRec("TankPriceS"))
    If IsNull(selRec("BagPriceV")) = False Then BagPriceV_ = Trim("" & selRec("BagPriceV"))
    If IsNull(selRec("BagBack")) = False Then BagBack_ = Trim("" & selRec("BagBack"))
    If IsNull(selRec("Container")) = False Then Container_ = Trim("" & selRec("Container"))
    If IsNull(selRec("Lack")) = False Then Lack_ = Trim("" & selRec("Lack"))
    If IsNull(selRec("PaperBox")) = False Then PaperBox_ = Trim("" & selRec("PaperBox"))
    If IsNull(selRec("BoutS")) = False Then BoutS_ = Trim("" & selRec("BoutS"))
    If IsNull(selRec("CovidBin")) = False Then CovidBin_ = Trim("" & selRec("CovidBin"))
    If IsNull(selRec("CovidBin200")) = False Then CovidBin200_ = Trim("" & selRec("CovidBin200"))
    If IsNull(selRec("Rack")) = False Then Rack_ = Trim("" & selRec("Rack"))
    PhotoOnWeight_ = Trim("" & selRec("PhotoOnWeight"))
    PhotoOnBelt_ = Trim("" & selRec("PhotoOnBelt"))
    PhotoOnBoxfeed_ = Trim("" & selRec("PhotoOnBoxfeed"))
    NewCustJob_ = Trim("" & selRec("NewCustJob"))
    NewCustJobBI_ = Trim("" & selRec("NewCustJobBI"))
    CustPhoto_ = Trim("" & selRec("CustPhoto"))
    CustCon_ = Trim("" & selRec("CustCon"))
    CloseWeight_ = Trim("" & selRec("Closeweight"))
    ConName_ = Trim("" & selRec("ConName"))
    IsChk_ = Trim("" & selRec("IsChk"))
    EmpTrain_ = Trim("" & selRec("EmpTrain"))
    CarCheck_ = Trim("" & selRec("CarCheck"))
    TruckRateID_ = Trim("" & selRec("TruckRateID"))
    
    If Trim("" & selRec("AKPBoxCode")) <> "" Then
        AKPBoxCode_ = Trim("" & selRec("AKPBoxCode"))
    End If
    If Trim("" & selRec("AKPBoxCode2")) <> "" Then
        AKPBoxCode2_ = Trim("" & selRec("AKPBoxCode2"))
    End If
     If IsNull(selRec("PhotoS")) = False Then PhotoS_ = Trim("" & selRec("PhotoS"))
    If IsNull(selRec("LackNew")) = False Then LackNew_ = Trim("" & selRec("LackNew"))
    If IsNull(selRec("LackGas")) = False Then LackGas_ = Trim("" & selRec("LackGas"))
    If IsNull(selRec("Plate")) = False Then Plate_ = Trim("" & selRec("Plate"))
    If IsNull(selRec("Plastic")) = False Then Plastic_ = Trim("" & selRec("Plastic"))
    If IsNull(selRec("BoutCage")) = False Then BoutCage_ = Trim("" & selRec("BoutCage"))
    
End Sub

Public Sub ClearTimeTable()
    Set colTimeTable_ = Nothing
    DBConnExc "Delete tbJobDataTimeTable WHERE JobDataCarID='" & id_ & "'"
End Sub
