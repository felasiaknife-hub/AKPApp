VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AKPFin09DIW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim isChanged_ As Boolean
Dim DocNo_ As String
Dim MenifestNo_ As String
Dim WasteGenNo_ As String
Dim TransRegisNo1_ As String
Dim TransportName2_ As String
Dim CarRegisID_ As String
Dim WorkDate_ As String
Dim IsCanceled_ As String
Dim isClosed_ As String
Dim DriverName_ As String
Dim JobDataCarID_ As String
Dim CompanyName_ As String
Dim CustAddr_ As String
Dim CustomerCode_ As String
Dim RegisterNo_ As String
Dim WasteManageID_ As String
Dim WasteName_ As String
Dim CustWeight_ As Double
Dim DisposerWeight_ As Double
Dim SubTypeDesc_ As String
Dim TruckTypeDesc_ As String
Dim FactoryRegisNo_ As String
Dim EUCode_ As String
Dim PriorityLevel_ As Integer
Dim TreatmentUnitID_ As String
Dim RefME_ As String
Dim HAType_ As String
Dim ChargeWType_ As String
Dim TreatmentRate_ As Double

Public Property Get DocNo() As String
    DocNo = DocNo_
End Property
Public Property Let DocNo(ByVal newval As String)
    If Trim(DocNo_) <> Trim(newval) Then
        DocNo_ = Trim(newval)
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
Public Property Get DisposerWeight() As Double
    DisposerWeight = DisposerWeight_
End Property
Public Property Let DisposerWeight(ByVal newval As Double)
    If DisposerWeight_ <> newval Then
        DisposerWeight_ = newval
        isChanged_ = True
    End If
End Property
Public Property Get TreatmentRate() As Double
    TreatmentRate = TreatmentRate_
End Property
Public Property Let TreatmentRate(ByVal newval As Double)
    If TreatmentRate_ <> newval Then
        TreatmentRate_ = newval
        isChanged_ = True
    End If
End Property
Public Property Get CustWeight() As Double
    CustWeight = CustWeight_
End Property
Public Property Let CustWeight(ByVal newval As Double)
    If CustWeight_ <> newval Then
        CustWeight_ = newval
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
Public Property Get CarRegisID() As String
    CarRegisID = CarRegisID_
End Property
Public Property Let CarRegisID(ByVal newval As String)
    If Trim(CarRegisID_) <> Trim(newval) Then
        CarRegisID_ = Trim(newval)
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
Public Property Get IsCanceled() As String
    IsCanceled = IsCanceled_
End Property
Public Property Let IsCanceled(ByVal newval As String)
    If Trim(IsCanceled_) <> Trim(newval) Then
        IsCanceled_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get isClosed() As String
    isClosed = isClosed_
End Property
Public Property Let isClosed(ByVal newval As String)
    If Trim(isClosed_) <> Trim(newval) Then
        isClosed_ = Trim(newval)
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
Public Property Get JobDataCarID() As String
    JobDataCarID = JobDataCarID_
End Property
Public Property Let JobDataCarID(ByVal newval As String)
    If Trim(JobDataCarID_) <> Trim(newval) Then
        JobDataCarID_ = Trim(newval)
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
Public Property Get CustAddr() As String
    CustAddr = CustAddr_
End Property
Public Property Let CustAddr(ByVal newval As String)
    If Trim(CustAddr_) <> Trim(newval) Then
        CustAddr_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get CustomerCode() As String
    CustomerCode = CustomerCode_
End Property
Public Property Let CustomerCode(ByVal newval As String)
    If Trim(CustomerCode_) <> Trim(newval) Then
        CustomerCode_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get RegisterNo() As String
    RegisterNo = RegisterNo_
End Property
Public Property Let RegisterNo(ByVal newval As String)
    If Trim(RegisterNo_) <> Trim(newval) Then
        RegisterNo_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get WasteManageID() As String
    WasteManageID = WasteManageID_
End Property
Public Property Let WasteManageID(ByVal newval As String)
    If Trim(WasteManageID_) <> Trim(newval) Then
        WasteManageID_ = Trim(newval)
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
Public Property Get SubTypeDesc() As String
    SubTypeDesc = SubTypeDesc_
End Property
Public Property Let SubTypeDesc(ByVal newval As String)
    If Trim(SubTypeDesc_) <> Trim(newval) Then
        SubTypeDesc_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get TruckTypeDesc() As String
    TruckTypeDesc = TruckTypeDesc_
End Property
Public Property Let TruckTypeDesc(ByVal newval As String)
    If Trim(TruckTypeDesc_) <> Trim(newval) Then
        TruckTypeDesc_ = Trim(newval)
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
Public Property Get EUCode() As String
    EUCode = EUCode_
End Property
Public Property Let EUCode(ByVal newval As String)
    If Trim(EUCode_) <> Trim(newval) Then
        EUCode_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get PriorityLevel() As String
    PriorityLevel = PriorityLevel_
End Property
Public Property Let PriorityLevel(ByVal newval As String)
    If Trim(PriorityLevel_) <> Trim(newval) Then
        PriorityLevel_ = Trim(newval)
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
Public Property Get RefME() As String
    RefME = RefME_
End Property
Public Property Let RefME(ByVal newval As String)
    If Trim(RefME_) <> Trim(newval) Then
        RefME_ = Trim(newval)
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
Public Property Get ChargeWType() As String
    ChargeWType = ChargeWType_
End Property
Public Property Let ChargeWType(ByVal newval As String)
    If Trim(ChargeWType_) <> Trim(newval) Then
        ChargeWType_ = Trim(newval)
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
    sqlStr = "SELECT * FROM TbReportFin09_2"
    Set tmpRec = New ADODB.Recordset
    GetDBConnection
    tmpRec.open sqlStr, DBConn, adOpenDynamic, adLockOptimistic
        tmpRec.AddNew
        tmpRec("DocNo") = DocNo_
        tmpRec("MenifestNo") = MenifestNo_
        tmpRec("WasteGenNo") = WasteGenNo_
        tmpRec("TransRegisNo1") = TransRegisNo1_
        tmpRec("TransportName2") = TransportName2_
        tmpRec("CarRegisID") = CarRegisID_
        tmpRec("WorkDate") = WorkDate_
        tmpRec("IsCanceled") = IsCanceled_
        tmpRec("isClosed") = isClosed_
        tmpRec("DriverName") = DriverName_
        tmpRec("JobDataCarID") = JobDataCarID_
        tmpRec("CompanyName") = CompanyName_
        tmpRec("CustAddr") = CustAddr_
        tmpRec("CustomerCode") = CustomerCode_
        tmpRec("WasteName") = WasteName_
        tmpRec("CustWeight") = CustWeight_
        tmpRec("DisposerWeight") = DisposerWeight_
        tmpRec("SubTypeDesc") = SubTypeDesc_
        tmpRec("TruckTypeDesc") = TruckTypeDesc_
        tmpRec("FactoryRegisNo") = FactoryRegisNo_
        tmpRec("EUCode") = EUCode_
        tmpRec("PriorityLevel") = PriorityLevel_
        tmpRec("TreatmentUnitID") = TreatmentUnitID_
        tmpRec("RefME") = RefME_
        tmpRec("HAType") = HAType_
        tmpRec("ChargeWType") = ChargeWType_
        tmpRec("TreatmentRate") = TreatmentRate_
        tmpRec.Update
        tmpRec.Requery

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
    DBConnExc "Delete TbReportFin09_2"
    Delete = ""
    Exit Function
delErr:
    Delete = err.Description
End Function
