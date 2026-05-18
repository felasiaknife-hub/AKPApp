VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BWGCustVisitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim id_ As String
Dim CustomerID_ As String
Dim CompanyName_ As String
Dim ContactName_ As String
Dim Position_ As String
Dim Tel_ As String
Dim Email_ As String
Dim TotalPeople_ As Double
Dim VisitDate_ As String
Dim VisitType_ As String
Dim VisitTime_ As String
Dim Point_ As Double
Dim FullPoint_ As Integer
Dim ServiceStat_ As String
Dim Remark_ As String
Dim isCancel_ As String
Dim isComplete_ As String
Dim LookOnline_ As String
Dim isUse_ As String
Dim CustomerCode_ As String
Dim CustTypeFlag_ As String
Dim isChanged_ As Boolean
Dim SaleStaff_ As String
Dim IsInform_ As String
Dim IsNewCust_ As String
Dim IsNoCome_ As String
Dim CheckList_ As String
Dim CCTV_ As String
Dim EmpPPE_ As String
Dim FieldClean_ As String
Dim EmpReady_ As String
Dim WorkProcess_ As String
Dim Welcome_ As String
Dim NameDisplay_  As String

Public Property Get isChanged() As Boolean
    isChanged = isChanged_
End Property
Public Property Let isChanged(ByVal newval As Boolean)
    isChanged_ = newval
End Property
Public Property Get ID() As String
    ID = id_
End Property
Public Property Let ID(ByVal newval As String)
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
Public Property Get SaleStaff() As String
    SaleStaff = SaleStaff_
End Property
Public Property Let SaleStaff(ByVal newval As String)
    If Trim(SaleStaff_) <> Trim(newval) Then
        SaleStaff_ = Trim(newval)
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

Public Property Get Position() As String
    Position = Position_
End Property
Public Property Let Position(ByVal newval As String)
    If Trim(Position_) <> Trim(newval) Then
        Position_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get ContactName() As String
    ContactName = ContactName_
End Property
Public Property Let ContactName(ByVal newval As String)
    If Trim(ContactName_) <> Trim(newval) Then
        ContactName_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get Tel() As String
    Tel = Tel_
End Property
Public Property Let Tel(ByVal newval As String)
    If Trim(Tel_) <> Trim(newval) Then
        Tel_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get Email() As String
    Email = Email_
End Property
Public Property Let Email(ByVal newval As String)
    If Trim(Email_) <> Trim(newval) Then
        Email_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get TotalPeople() As Double
    TotalPeople = TotalPeople_
End Property
Public Property Let TotalPeople(ByVal newval As Double)
    If Trim(TotalPeople_) <> Trim(newval) Then
        TotalPeople_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get VisitDate() As String
    VisitDate = VisitDate_
End Property
Public Property Let VisitDate(ByVal newval As String)
    If Trim(VisitDate_) <> Trim(newval) Then
        VisitDate_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get VisitType() As String
    VisitType = VisitType_
End Property
Public Property Let VisitType(ByVal newval As String)
    If Trim(VisitType_) <> Trim(newval) Then
        VisitType_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get VisitTime() As String
    VisitTime = VisitTime_
End Property
Public Property Let VisitTime(ByVal newval As String)
    If Trim(VisitTime_) <> Trim(newval) Then
        VisitTime_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get Point() As Double
    Point = Point_
End Property
Public Property Let Point(ByVal newval As Double)
    If Trim(Point_) <> Trim(newval) Then
        Point_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get FullPoint() As Integer
    FullPoint = FullPoint_
End Property
Public Property Let FullPoint(ByVal newval As Integer)
    If Trim(FullPoint_) <> Trim(newval) Then
        FullPoint_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get ServiceStat() As String
    ServiceStat = ServiceStat_
End Property
Public Property Let ServiceStat(ByVal newval As String)
    If Trim(ServiceStat_) <> Trim(newval) Then
        ServiceStat_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get Remark() As String
    Remark = Remark_
End Property
Public Property Let Remark(ByVal newval As String)
    If Trim(Remark_) <> Trim(newval) Then
        Remark_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get isCancel() As Boolean
    If Trim(isCancel_) = "Y" Then
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
Public Property Get isComplete() As Boolean
    If Trim(isComplete_) = "Y" Then
        isComplete = True
    Else
        isComplete = False
    End If
End Property
Public Property Let isComplete(ByVal newval As Boolean)
Dim tmpStr As String
    If newval Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(isComplete_) <> Trim(tmpStr) Then
        isComplete_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property
Public Property Get LookOnline() As Boolean
    If Trim(LookOnline_) = "Y" Then
        LookOnline = True
    Else
        LookOnline = False
    End If
End Property
Public Property Let LookOnline(ByVal newval As Boolean)
Dim tmpStr As String
    If newval Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(LookOnline_) <> Trim(tmpStr) Then
        LookOnline_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property
Public Property Get IsInform() As Boolean
    If Trim(IsInform_) = "Y" Then
        IsInform = True
    Else
        IsInform = False
    End If
End Property
Public Property Let IsInform(ByVal newval As Boolean)
Dim tmpStr As String
    If newval Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(IsInform_) <> Trim(tmpStr) Then
        IsInform_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property
Public Property Get IsNewCust() As Boolean
    If Trim(IsNewCust_) = "Y" Then
        IsNewCust = True
    Else
        IsNewCust = False
    End If
End Property
Public Property Let IsNewCust(ByVal newval As Boolean)
Dim tmpStr As String
    If newval Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(IsNewCust_) <> Trim(tmpStr) Then
        IsNewCust_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property
Public Property Get IsNoCome() As Boolean
    If Trim(IsNoCome_) = "Y" Then
        IsNoCome = True
    Else
        IsNoCome = False
    End If
End Property
Public Property Let IsNoCome(ByVal newval As Boolean)
Dim tmpStr As String
    If newval Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(IsNoCome_) <> Trim(tmpStr) Then
        IsNoCome_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property
Public Property Get CheckList() As Boolean
    If Trim(CheckList_) = "Y" Then
        CheckList = True
    Else
        CheckList = False
    End If
End Property
Public Property Let CheckList(ByVal newval As Boolean)
Dim tmpStr As String
    If newval Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(CheckList_) <> Trim(tmpStr) Then
        CheckList_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property
Public Property Get CCTV() As Boolean
    If Trim(CCTV_) = "Y" Then
        CCTV = True
    Else
        CCTV = False
    End If
End Property
Public Property Let CCTV(ByVal newval As Boolean)
Dim tmpStr As String
    If newval Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(CCTV_) <> Trim(tmpStr) Then
        CCTV_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property
Public Property Get EmpPPE() As Boolean
    If Trim(EmpPPE_) = "Y" Then
        EmpPPE = True
    Else
        EmpPPE = False
    End If
End Property
Public Property Let EmpPPE(ByVal newval As Boolean)
Dim tmpStr As String
    If newval Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(EmpPPE_) <> Trim(tmpStr) Then
        EmpPPE_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property
Public Property Get FieldClean() As Boolean
    If Trim(FieldClean_) = "Y" Then
        FieldClean = True
    Else
        FieldClean = False
    End If
End Property
Public Property Let FieldClean(ByVal newval As Boolean)
Dim tmpStr As String
    If newval Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(FieldClean_) <> Trim(tmpStr) Then
        FieldClean_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property
Public Property Get EmpReady() As Boolean
    If Trim(EmpReady_) = "Y" Then
        EmpReady = True
    Else
        EmpReady = False
    End If
End Property
Public Property Let EmpReady(ByVal newval As Boolean)
Dim tmpStr As String
    If newval Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(EmpReady_) <> Trim(tmpStr) Then
        EmpReady_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property
Public Property Get WorkProcess() As Boolean
    If Trim(WorkProcess_) = "Y" Then
        WorkProcess = True
    Else
        WorkProcess = False
    End If
End Property
Public Property Let WorkProcess(ByVal newval As Boolean)
Dim tmpStr As String
    If newval Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(WorkProcess_) <> Trim(tmpStr) Then
        WorkProcess_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property
Public Property Get Welcome() As Boolean
    If Trim(Welcome_) = "Y" Then
        Welcome = True
    Else
        Welcome = False
    End If
End Property
Public Property Let Welcome(ByVal newval As Boolean)
Dim tmpStr As String
    If newval Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(Welcome_) <> Trim(tmpStr) Then
        Welcome_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property
Public Property Get isUse() As Boolean
    If Trim(isUse_) = "Y" Then
        isUse = True
    Else
        isUse = False
    End If
End Property
Public Property Let isUse(ByVal newval As Boolean)
Dim tmpStr As String
    If newval Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(isUse_) <> Trim(tmpStr) Then
        isUse_ = Trim(tmpStr)
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


Public Property Get CustTypeFlag() As String
    CustTypeFlag = CustTypeFlag_
End Property
Public Property Let CustTypeFlag(ByVal newval As String)
    If Trim(CustTypeFlag_) <> Trim(newval) Then
        CustTypeFlag_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get NameDisplay() As String
    NameDisplay = NameDisplay_
End Property
Public Property Let NameDisplay(ByVal newval As String)
    If Trim(NameDisplay_) <> Trim(newval) Then
        NameDisplay_ = Trim(newval)
        isChanged_ = True
    End If
End Property


Public Sub Save()
On Error GoTo ErrD
Dim tmpRec As ADODB.Recordset, rs As ADODB.Recordset
Dim x%, isNew As Boolean
Dim sqlStr As String
Dim LastUpdate As String
    sqlStr = "SELECT *, CONVERT(DATETIME, GETDATE(), 103) AS GETDATE FROM tbCustomerVisitHeader WHERE VisitID='" & id_ & "'"
    Set tmpRec = New ADODB.Recordset
    GetDBConnection
    tmpRec.open sqlStr, DBConn, adOpenDynamic, adLockOptimistic
    isNew = False
    If tmpRec.EOF Then
        tmpRec.AddNew
        tmpRec("VisitID") = id_
        isNew = True
    End If
    tmpRec("CustomerID") = Trim(CustomerID_)
    tmpRec("CompanyName") = Trim(CompanyName_)
    tmpRec("NameDisplay") = Trim(NameDisplay_)
    tmpRec("SaleStaff") = Trim(SaleStaff_)
    tmpRec("ContactName") = Trim(ContactName_)
    tmpRec("Position") = Trim(Position_)
    tmpRec("Tel") = Trim(Tel_)
    tmpRec("Email") = Trim(Email_)
    tmpRec("TotalPeople") = Trim(TotalPeople_)
    tmpRec("VisitDate") = Trim(VisitDate_)
    tmpRec("VisitType") = Trim(VisitType_)
    tmpRec("VisitTime") = Trim(VisitTime_)
    If Point_ >= "0" Then
        tmpRec("Point") = Trim(Point_)
    End If
    If FullPoint_ >= "0" Then
        tmpRec("FullPoint") = Trim(FullPoint_)
    End If
    tmpRec("isCancel") = IIf(Trim(isCancel_) = "", "N", Trim(isCancel_))
    tmpRec("isComplete") = IIf(Trim(isComplete_) = "", "N", Trim(isComplete_))
    tmpRec("LookOnline") = IIf(Trim(LookOnline_) = "", "N", Trim(LookOnline_))
    tmpRec("IsInform") = IIf(Trim(IsInform_) = "", "N", Trim(IsInform_))
    tmpRec("IsNewCust") = IIf(Trim(IsNewCust_) = "", "N", Trim(IsNewCust_))
    tmpRec("isUse") = IIf(Trim(isUse_) = "", "N", Trim(isUse_))
    tmpRec("ServiceStat") = Trim(ServiceStat_)
    tmpRec("Remark") = Trim(Remark_)
    tmpRec("IsNoCome") = IIf(Trim(IsNoCome_) = "", "N", Trim(IsNoCome_))
    tmpRec("CheckList") = IIf(Trim(CheckList_) = "", "N", Trim(CheckList_))
    tmpRec("CCTV") = IIf(Trim(CCTV_) = "", "N", Trim(CCTV_))
    tmpRec("EmpPPE") = IIf(Trim(EmpPPE_) = "", "N", Trim(EmpPPE_))
    tmpRec("FieldClean") = IIf(Trim(FieldClean_) = "", "N", Trim(FieldClean_))
    tmpRec("EmpReady") = IIf(Trim(EmpReady_) = "", "N", Trim(EmpReady_))
    tmpRec("WorkProcess") = IIf(Trim(WorkProcess_) = "", "N", Trim(WorkProcess_))
    tmpRec("Welcome") = IIf(Trim(Welcome_) = "", "N", Trim(Welcome_))
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

Public Sub SetDataFromRecordset(selRec As ADODB.Recordset)
    id_ = Trim("" & selRec("VisitID"))
    CustomerID_ = Trim("" & selRec("CustomerID"))
    CompanyName_ = Trim("" & selRec("CompanyName"))
     If isExist(selRec, "SaleStaff") Then SaleStaff_ = Trim("" & selRec("SaleStaff"))
     If isExist(selRec, "SaleStaffName") Then SaleStaff_ = Trim("" & selRec("SaleStaffName"))
    ContactName_ = Trim("" & selRec("ContactName"))
    Position_ = Trim("" & selRec("Position"))
    If isExist(selRec, "Tel") Then Tel_ = Trim("" & selRec("Tel"))
    Email_ = Trim("" & selRec("Email"))
    If IsNull(selRec("TotalPeople")) = False Then TotalPeople_ = Trim("" & selRec("TotalPeople"))
    VisitDate_ = Trim("" & selRec("VisitDate"))
    VisitType_ = Trim("" & selRec("VisitType"))
    VisitTime_ = Trim("" & selRec("VisitTime"))
    If IsNull(selRec("Point")) = False Then Point_ = Trim("" & selRec("Point"))
    If IsNull(selRec("FullPoint")) = False Then FullPoint_ = Trim("" & selRec("FullPoint"))
    isCancel_ = Trim("" & selRec("isCancel"))
    isComplete_ = Trim("" & selRec("isComplete"))
    LookOnline_ = Trim("" & selRec("LookOnline"))
    IsInform_ = Trim("" & selRec("IsInform"))
    IsNewCust_ = Trim("" & selRec("IsNewCust"))
    isUse_ = Trim("" & selRec("isUse"))
    ServiceStat_ = Trim("" & selRec("ServiceStat"))
    Remark_ = Trim("" & selRec("Remark"))
    IsNoCome_ = Trim("" & selRec("IsNoCome"))
    CheckList_ = Trim("" & selRec("CheckList"))
    CCTV_ = Trim("" & selRec("CCTV"))
    EmpPPE_ = Trim("" & selRec("EmpPPE"))
    FieldClean_ = Trim("" & selRec("FieldClean"))
     EmpReady_ = Trim("" & selRec("EmpReady"))
    WorkProcess_ = Trim("" & selRec("WorkProcess"))
    Welcome_ = Trim("" & selRec("Welcome"))
    NameDisplay_ = Trim("" & selRec("NameDisplay"))
End Sub
