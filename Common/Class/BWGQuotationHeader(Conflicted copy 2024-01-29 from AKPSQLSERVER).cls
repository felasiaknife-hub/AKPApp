VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BWGQuotationHeader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim id_ As String
Dim QuotationNo_ As String
Dim BMEQuotation_ As String
Dim IssuedDate_ As String
Dim AttentionTo_ As String
Dim CustomerID_ As String
Dim CustomerCode_ As String
Dim CustomerName_ As String
Dim issuedStaffID_ As String
Dim issuedStaffName_ As String
Dim isCancel_ As String
Dim cancelDate_ As String
Dim isApproved_ As String
Dim Approveddate_ As String
Dim ApprovedStaffID_ As String
Dim ApprovedStaffName_ As String
Dim discountAmt_ As Double
Dim discountType_ As String
Dim cancelStaffID_ As String
Dim cancelStaffName_ As String
Dim isPriceIncTransport_ As String
Dim isPriceIncVAT_ As String
Dim StaffCountPerCar_ As Integer
Dim minWeightPerCar_ As Double
Dim colQuotationDet_ As Collection
Dim colRemarks_ As Collection
Dim colSelTruck_ As Collection
Dim companyID_ As String
Dim extraLaborFee_ As Double
Dim truckTypeID_ As String
Dim truckTypeName_ As String
Dim transferRate_ As Double
Dim transferUnitID_ As String
Dim transferUnitName_ As String
Dim signName_ As String
Dim signPosition_ As String
Dim isTrailer_ As String
Dim nonMinRate_ As Double
Dim hazMinRate_ As Double
Dim nblMinRate_ As Double
Dim hblMinRate_ As Double
Dim isClosed_ As String
Dim closedDate_ As String
Dim closeStaffID_ As String
Dim closeStaffName_ As String
Dim Remdesc_ As String
Dim isChanged_ As Boolean
Dim isBmeNew_ As String
Dim QuotationType_ As Integer
Dim isQuotationReq_ As String
Dim isWasteApproveAll_ As Boolean
Dim Status_ As String
Dim isJob_ As String
Dim isCase_ As String
Dim SaleAP_ As String
Dim LastUpdateUser_ As String
Dim Promies_ As String
Dim StartPromiesDate_ As String
Dim EndPromiesDate_ As String
Dim RefQuotationID_ As String
Dim TRuck1ID_ As String
Dim Truck1Trailer_ As String
Dim Truck1Rate_ As String
Dim Truck1Unit_ As String
Dim TRuck2ID_ As String
Dim Truck2Trailer_ As String
Dim Truck2Rate_ As String
Dim Truck2Unit_ As String
Dim TRuck3ID_ As String
Dim Truck3Trailer_ As String
Dim Truck3Rate_ As String
Dim Truck3Unit_ As String
Dim TRuck4ID_ As String
Dim Truck4Trailer_ As String
Dim Truck4Rate_ As String
Dim Truck4Unit_ As String
Dim TRuck5ID_ As String
Dim Truck5Trailer_ As String
Dim Truck5Rate_ As String
Dim Truck5Unit_ As String
Dim TRuck6ID_ As String
Dim Truck6Trailer_ As String
Dim Truck6Rate_ As String
Dim Truck6Unit_ As String
Dim TRuck7ID_ As String
Dim Truck7Trailer_ As String
Dim Truck7Rate_ As String
Dim Truck7Unit_ As String
Dim TRuck8ID_ As String
Dim Truck8Trailer_ As String
Dim Truck8Rate_ As String
Dim Truck8Unit_ As String
Dim TRuck9ID_ As String
Dim Truck9Trailer_ As String
Dim Truck9Rate_ As String
Dim Truck9Unit_ As String
Dim TransRate_ As String

Dim MaxWeight_ As Double
Dim MaxRate_ As Double
Dim MaxRateType_ As String
Dim ProjectID_ As String
Public Property Get ProjectID() As String
    ProjectID = ProjectID_
End Property
Public Property Let ProjectID(ByVal newval As String)
    If Trim(ProjectID_) <> Trim(newval) Then
        ProjectID_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get isBmeNew() As String
    isBmeNew = isBmeNew_
End Property
Public Property Let isBmeNew(ByVal newval As String)
    If Trim(isBmeNew_) <> Trim(newval) Then
        isBmeNew_ = Trim(newval)
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

Public Property Get TRuck1ID() As String
    TRuck1ID = TRuck1ID_
End Property
Public Property Let TRuck1ID(ByVal newval As String)
    If Trim(TRuck1ID_) <> Trim(newval) Then
        TRuck1ID_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get Truck1Trailer() As String
    Truck1Trailer = Truck1Trailer_
End Property
Public Property Let Truck1Trailer(ByVal newval As String)
    If Trim(Truck1Trailer_) <> Trim(newval) Then
        Truck1Trailer_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get Truck1Rate() As String
    Truck1Rate = Truck1Rate_
End Property
Public Property Let Truck1Rate(ByVal newval As String)
    If Trim(Truck1Rate_) <> Trim(newval) Then
        Truck1Rate_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get Truck1Unit() As String
    Truck1Unit = Truck1Unit_
End Property
Public Property Let Truck1Unit(ByVal newval As String)
    If Trim(Truck1Unit_) <> Trim(newval) Then
        Truck1Unit_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get TRuck2ID() As String
    TRuck2ID = TRuck2ID_
End Property
Public Property Let TRuck2ID(ByVal newval As String)
    If Trim(TRuck2ID_) <> Trim(newval) Then
        TRuck2ID_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get Truck2Trailer() As String
    Truck2Trailer = Truck2Trailer_
End Property
Public Property Let Truck2Trailer(ByVal newval As String)
    If Trim(Truck2Trailer_) <> Trim(newval) Then
        Truck2Trailer_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get Truck2Rate() As String
    Truck2Rate = Truck2Rate_
End Property
Public Property Let Truck2Rate(ByVal newval As String)
    If Trim(Truck2Rate_) <> Trim(newval) Then
        Truck2Rate_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get Truck2Unit() As String
    Truck2Unit = Truck2Unit_
End Property
Public Property Let Truck2Unit(ByVal newval As String)
    If Trim(Truck2Unit_) <> Trim(newval) Then
        Truck2Unit_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get TRuck3ID() As String
    TRuck3ID = TRuck3ID_
End Property
Public Property Let TRuck3ID(ByVal newval As String)
    If Trim(TRuck3ID_) <> Trim(newval) Then
        TRuck3ID_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get TRuck3Trailer() As String
    TRuck3Trailer = Truck3Trailer_
End Property
Public Property Let TRuck3Trailer(ByVal newval As String)
    If Trim(Truck3Trailer_) <> Trim(newval) Then
        Truck3Trailer_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get TRuck3Rate() As String
    TRuck3Rate = Truck3Rate_
End Property
Public Property Let TRuck3Rate(ByVal newval As String)
    If Trim(Truck3Rate_) <> Trim(newval) Then
        Truck3Rate_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get TRuck3Unit() As String
    TRuck3Unit = Truck3Unit_
End Property
Public Property Let TRuck3Unit(ByVal newval As String)
    If Trim(Truck3Unit_) <> Trim(newval) Then
        Truck3Unit_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get TRuck4ID() As String
    TRuck4ID = TRuck4ID_
End Property
Public Property Let TRuck4ID(ByVal newval As String)
    If Trim(TRuck4ID_) <> Trim(newval) Then
        TRuck4ID_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get TRuck4Trailer() As String
    TRuck4Trailer = Truck4Trailer_
End Property
Public Property Let TRuck4Trailer(ByVal newval As String)
    If Trim(Truck4Trailer_) <> Trim(newval) Then
        Truck4Trailer_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get TRuck4Rate() As String
    TRuck4Rate = Truck4Rate_
End Property
Public Property Let TRuck4Rate(ByVal newval As String)
    If Trim(Truck4Rate_) <> Trim(newval) Then
        Truck4Rate_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get TRuck4Unit() As String
    TRuck4Unit = Truck4Unit_
End Property
Public Property Let TRuck4Unit(ByVal newval As String)
    If Trim(Truck4Unit_) <> Trim(newval) Then
        Truck4Unit_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get TRuck5ID() As String
    TRuck5ID = TRuck5ID_
End Property
Public Property Let TRuck5ID(ByVal newval As String)
    If Trim(TRuck5ID_) <> Trim(newval) Then
        TRuck5ID_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get TRuck5Trailer() As String
    TRuck5Trailer = Truck5Trailer_
End Property
Public Property Let TRuck5Trailer(ByVal newval As String)
    If Trim(Truck5Trailer_) <> Trim(newval) Then
        Truck5Trailer_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get TRuck5Rate() As String
    TRuck5Rate = Truck5Rate_
End Property
Public Property Let TRuck5Rate(ByVal newval As String)
    If Trim(Truck5Rate_) <> Trim(newval) Then
        Truck5Rate_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get TRuck5Unit() As String
    TRuck5Unit = Truck5Unit_
End Property
Public Property Let TRuck5Unit(ByVal newval As String)
    If Trim(Truck5Unit_) <> Trim(newval) Then
        Truck5Unit_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get TRuck6ID() As String
    TRuck6ID = TRuck6ID_
End Property
Public Property Let TRuck6ID(ByVal newval As String)
    If Trim(TRuck6ID_) <> Trim(newval) Then
        TRuck6ID_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get TRuck6Trailer() As String
    TRuck6Trailer = Truck6Trailer_
End Property
Public Property Let TRuck6Trailer(ByVal newval As String)
    If Trim(Truck6Trailer_) <> Trim(newval) Then
        Truck6Trailer_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get TRuck6Rate() As String
    TRuck6Rate = Truck6Rate_
End Property
Public Property Let TRuck6Rate(ByVal newval As String)
    If Trim(Truck6Rate_) <> Trim(newval) Then
        Truck6Rate_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get TRuck6Unit() As String
    TRuck6Unit = Truck6Unit_
End Property
Public Property Let TRuck6Unit(ByVal newval As String)
    If Trim(Truck6Unit_) <> Trim(newval) Then
        Truck6Unit_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get TRuck7ID() As String
    TRuck7ID = TRuck7ID_
End Property
Public Property Let TRuck7ID(ByVal newval As String)
    If Trim(TRuck7ID_) <> Trim(newval) Then
        TRuck7ID_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get TRuck7Trailer() As String
    TRuck7Trailer = Truck7Trailer_
End Property
Public Property Let TRuck7Trailer(ByVal newval As String)
    If Trim(Truck7Trailer_) <> Trim(newval) Then
        Truck7Trailer_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get TRuck7Rate() As String
    TRuck7Rate = Truck7Rate_
End Property
Public Property Let TRuck7Rate(ByVal newval As String)
    If Trim(Truck7Rate_) <> Trim(newval) Then
        Truck7Rate_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get TRuck7Unit() As String
    TRuck7Unit = Truck7Unit_
End Property
Public Property Let TRuck7Unit(ByVal newval As String)
    If Trim(Truck7Unit_) <> Trim(newval) Then
        Truck7Unit_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get TRuck8ID() As String
    TRuck8ID = TRuck8ID_
End Property
Public Property Let TRuck8ID(ByVal newval As String)
    If Trim(TRuck8ID_) <> Trim(newval) Then
        TRuck8ID_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get TRuck8Trailer() As String
    TRuck8Trailer = Truck8Trailer_
End Property
Public Property Let TRuck8Trailer(ByVal newval As String)
    If Trim(Truck8Trailer_) <> Trim(newval) Then
        Truck8Trailer_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get TRuck8Rate() As String
    TRuck8Rate = Truck8Rate_
End Property
Public Property Let TRuck8Rate(ByVal newval As String)
    If Trim(Truck8Rate_) <> Trim(newval) Then
        Truck8Rate_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get TRuck8Unit() As String
    TRuck8Unit = Truck8Unit_
End Property
Public Property Let TRuck8Unit(ByVal newval As String)
    If Trim(Truck8Unit_) <> Trim(newval) Then
        Truck8Unit_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get TRuck9ID() As String
    TRuck9ID = TRuck9ID_
End Property
Public Property Let TRuck9ID(ByVal newval As String)
    If Trim(TRuck9ID_) <> Trim(newval) Then
        TRuck9ID_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get TRuck9Trailer() As String
    TRuck9Trailer = Truck9Trailer_
End Property
Public Property Let TRuck9Trailer(ByVal newval As String)
    If Trim(Truck9Trailer_) <> Trim(newval) Then
        Truck9Trailer_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get TRuck9Rate() As String
    TRuck9Rate = Truck9Rate_
End Property
Public Property Let TRuck9Rate(ByVal newval As String)
    If Trim(Truck9Rate_) <> Trim(newval) Then
        Truck9Rate_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get TRuck9Unit() As String
    TRuck9Unit = Truck9Unit_
End Property
Public Property Let TRuck9Unit(ByVal newval As String)
    If Trim(Truck9Unit_) <> Trim(newval) Then
        Truck9Unit_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get ID() As String
    ID = id_
End Property
Public Property Let ID(ByVal newval As String)
    id_ = Trim(newval)
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

Public Property Get CompanyID() As String
    CompanyID = companyID_
End Property
Public Property Let CompanyID(ByVal newval As String)
    If Trim(companyID_) <> Trim(newval) Then
        companyID_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get ExtraLaborFee() As Double
    ExtraLaborFee = extraLaborFee_
End Property
Public Property Let ExtraLaborFee(ByVal newval As Double)
    If extraLaborFee_ <> newval Then
        extraLaborFee_ = newval
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

Public Property Get TransferUnitID() As String
    TransferUnitID = transferUnitID_
End Property
Public Property Let TransferUnitID(ByVal newval As String)
    If Trim(transferUnitID_) <> Trim(newval) Then
        transferUnitID_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get TransferRate() As Double
    TransferRate = transferRate_
End Property
Public Property Let TransferRate(ByVal newval As Double)
    If transferRate_ <> newval Then
        transferRate_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get TransferUnitName() As String
    TransferUnitName = transferUnitName_
End Property
Public Property Let TransferUnitName(ByVal newval As String)
    If Trim(transferUnitName_) <> Trim(newval) Then
        transferUnitName_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get SignName() As String
    SignName = signName_
End Property
Public Property Let SignName(ByVal newval As String)
    If Trim(signName_) <> Trim(newval) Then
        signName_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get SignPosition() As String
    SignPosition = signPosition_
End Property
Public Property Let SignPosition(ByVal newval As String)
    If Trim(signPosition_) <> Trim(newval) Then
        signPosition_ = Trim(newval)
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

Public Property Get isClosed() As Boolean
    If isClosed_ = "Y" Then
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
        isClosed_ = tmpStr
        isChanged_ = True
    End If
End Property

Public Property Get ClosedDate() As String
    ClosedDate = closedDate_
End Property
Public Property Let ClosedDate(ByVal newval As String)
    If Trim(closedDate_) <> Trim(newval) Then
        closedDate_ = Trim(newval)
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

Public Property Get Remdesc() As String
    Remdesc = Remdesc_
End Property
Public Property Let Remdesc(ByVal newval As String)
    If Trim(Remdesc_) <> Trim(newval) Then
        Remdesc_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get IssuedDate() As String
    IssuedDate = IssuedDate_
End Property
Public Property Let IssuedDate(ByVal newval As String)
    If Trim(IssuedDate_) <> Trim(newval) Then
        IssuedDate_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get AttentionTo() As String
    AttentionTo = AttentionTo_
End Property
Public Property Let AttentionTo(ByVal newval As String)
    If Trim(AttentionTo_) <> Trim(newval) Then
        AttentionTo_ = Trim(newval)
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
    CustomerCode = CustomerCode_
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

Public Property Get IssuedStaffID() As String
    IssuedStaffID = issuedStaffID_
End Property
Public Property Let IssuedStaffID(ByVal newval As String)
    If Trim(issuedStaffID_) <> Trim(newval) Then
        issuedStaffID_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get IssuedStaffName() As String
    IssuedStaffName = issuedStaffName_
End Property
Public Property Let IssuedStaffName(ByVal newval As String)
    If Trim(issuedStaffName_) <> Trim(newval) Then
        issuedStaffName_ = Trim(newval)
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

Public Property Get CancelDate() As String
    CancelDate = cancelDate_
End Property
Public Property Let CancelDate(ByVal newval As String)
    If Trim(cancelDate_) <> Trim(newval) Then
        cancelDate_ = Trim(newval)
        isChanged_ = True
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

Public Property Get ApprovedStaffID() As String
    ApprovedStaffID = ApprovedStaffID_
End Property
Public Property Let ApprovedStaffID(ByVal newval As String)
    If Trim(ApprovedStaffID_) <> Trim(newval) Then
        ApprovedStaffID_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get ApprovedStaffName() As String
    ApprovedStaffName = ApprovedStaffName_
End Property
Public Property Let ApprovedStaffName(ByVal newval As String)
    If Trim(ApprovedStaffName_) <> Trim(newval) Then
        ApprovedStaffName_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get DiscountAmt() As Double
    DiscountAmt = discountAmt_
End Property
Public Property Let DiscountAmt(ByVal newval As Double)
    If discountAmt_ <> newval Then
        discountAmt_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get DiscountType() As String
    DiscountType = discountType_
End Property
Public Property Let DiscountType(ByVal newval As String)
    If Trim(discountType_) <> Trim(newval) Then
        discountType_ = Trim(newval)
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

Public Property Get IsPriceIncTransport() As Boolean
    If Trim(isPriceIncTransport_) = "Y" Then
        IsPriceIncTransport = True
    Else
        IsPriceIncTransport = False
    End If
End Property
Public Property Let IsPriceIncTransport(ByVal newval As Boolean)
Dim tmpStr As String
    If newval Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(isPriceIncTransport_) <> Trim(tmpStr) Then
        isPriceIncTransport_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property

Public Property Get IsPriceIncVAT() As Boolean
    If Trim(isPriceIncVAT_) = "Y" Then
        IsPriceIncVAT = True
    Else
        IsPriceIncVAT = False
    End If
End Property
Public Property Let IsPriceIncVAT(ByVal newval As Boolean)
Dim tmpStr As String
    If newval Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(isPriceIncVAT_) <> Trim(tmpStr) Then
        isPriceIncVAT_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property

Public Property Get StaffCountPerCar() As Double
    StaffCountPerCar = StaffCountPerCar_
End Property
Public Property Let StaffCountPerCar(ByVal newval As Double)
    If StaffCountPerCar_ <> newval Then
        StaffCountPerCar_ = newval
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

Public Property Get isWasteApproveAll() As Boolean
        isWasteApproveAll = isWasteApproveAll_
End Property
Public Property Let isWasteApproveAll(ByVal newval As Boolean)
    If isWasteApproveAll_ <> newval Then
        isWasteApproveAll_ = newval
    End If
End Property

Public Property Get QuotationDetails() As Collection

Dim tmpS As New BWGSearchManager
Dim i As Integer
    If isQuotationReq_ <> "Y" Then
        If colQuotationDet_ Is Nothing Then Set colQuotationDet_ = tmpS.QuotationDetailSearch("QuotationID='" & id_ & "'")
    Else
        If colQuotationDet_ Is Nothing Then Set colQuotationDet_ = tmpS.QuotationDetailSearch("QuotationID='" & id_ & "'", "vw_QuotationDetailSearch_New")
    End If
    Set QuotationDetails = colQuotationDet_
    
    isWasteApproveAll_ = True
    For i = 1 To colQuotationDet_.Count
        Select Case colQuotationDet_(i).WasteStatus
            Case "A", "IA"
                isWasteApproveAll_ = True
            Case Else
                isWasteApproveAll_ = False
                Exit For
        End Select
    Next i
End Property
Public Property Set QuotationDetails(ByVal newval As Collection)
    Set colQuotationDet_ = newval
    isChanged_ = True
End Property

Public Sub RequeryQuotationDetail()
    Set colQuotationDet_ = Nothing
End Sub
Public Property Get selTruck() As Collection
Dim tmpS As New BWGSearchManager
    If colSelTruck_ Is Nothing Then Set colSelTruck_ = tmpS.TruckTypeSearch("TruckTypeid='" & id_ & "'")
    Set selTruck = colSelTruck_
End Property
Public Property Set selTruck(ByVal newval As Collection)
    Set colSelTruck_ = newval
    isChanged_ = True
End Property

Public Property Get RemarkDetails() As Collection
Dim tmpS As New BWGSearchManager
    If colRemarks_ Is Nothing Then Set colRemarks_ = tmpS.QuotationRemark("QuotationID='" & id_ & "'")
    Set RemarkDetails = colRemarks_
End Property
Public Property Set RemarkDetails(ByVal newval As Collection)
    Set colRemarks_ = newval
    isChanged_ = True
End Property
Public Sub RequeryRemark()
    Set colRemarks_ = Nothing
End Sub

Public Property Get QuotationType() As Integer
    QuotationType = QuotationType_
End Property
Public Property Let QuotationType(ByVal newval As Integer)
    If QuotationType_ <> newval Then
        QuotationType_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get isChanged() As Boolean
    isChanged = isChanged_
End Property
Public Property Let isChanged(ByVal newval As Boolean)
    isChanged_ = newval
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

Public Property Get Status() As Integer
    Status = Status_
End Property
Public Property Let Status(ByVal newval As Integer)
    If Status_ <> newval Then
        Status_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get isJob() As String
    isJob = isJob_
End Property
Public Property Let isJob(ByVal newval As String)
    If isJob_ <> newval Then
        isJob_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get isCase() As String
    isCase = isCase_
End Property
Public Property Let isCase(ByVal newval As String)
    If isCase_ <> newval Then
        isCase_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get LastUpdateUser() As String
    LastUpdateUser = LastUpdateUser_
End Property
Public Property Let LastUpdateUser(ByVal newval As String)
    LastUpdateUser_ = newval
End Property

Public Property Get SaleAP() As String
    SaleAP = SaleAP_
End Property
Public Property Let SaleAP(ByVal newval As String)
    If Trim(SaleAP_) <> Trim(newval) Then
        SaleAP_ = Trim(newval)
    End If
End Property
Public Property Get Promies() As String
    Promies = Promies_
End Property
Public Property Let Promies(ByVal newval As String)
    If Trim(Promies_) <> Trim(newval) Then
        Promies_ = Trim(newval)
    End If
End Property
Public Property Get StartPromiesDate() As String
    StartPromiesDate = StartPromiesDate_
End Property
Public Property Let StartPromiesDate(ByVal newval As String)
    If Trim(StartPromiesDate_) <> Trim(newval) Then
     StartPromiesDate_ = Trim(newval)
    End If
End Property
Public Property Get EndPromiesDate() As String
    EndPromiesDate = EndPromiesDate_
End Property
Public Property Let EndPromiesDate(ByVal newval As String)
    If Trim(EndPromiesDate_) <> Trim(newval) Then
        EndPromiesDate_ = Trim(newval)
    End If
End Property

Public Property Get MaxWeight() As Double
    MaxWeight = MaxWeight_
End Property
Public Property Let MaxWeight(ByVal newval As Double)
    If MaxWeight_ <> newval Then
        MaxWeight_ = newval
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
Public Property Get RefQuotationID() As String
    RefQuotationID = RefQuotationID_
End Property
Public Property Let RefQuotationID(ByVal newval As String)
    If RefQuotationID_ <> newval Then
        RefQuotationID_ = newval
        isChanged_ = True
    End If
End Property
Public Property Get MaxRateType() As String
    MaxRateType = MaxRateType_
End Property
Public Property Let MaxRateType(ByVal newval As String)
    If Trim(MaxRateType_) <> Trim(newval) Then
        MaxRateType_ = Trim(newval)
    End If
End Property

Public Sub Approve()
Dim tmpRec As ADODB.Recordset
Dim x%
Dim sqlStr As String
    sqlStr = "SELECT *, CONVERT(DATETIME, GETDATE(), 103) AS GETDATE FROM tbQuotationHeader WHERE QuotationID='" & id_ & "'"
    Set tmpRec = New ADODB.Recordset
    GetDBConnection
    tmpRec.open sqlStr, DBConn, adOpenDynamic, adLockOptimistic
    If Not tmpRec.EOF Then
        tmpRec("isApproved") = Trim(isApproved_)
        tmpRec("ApprovedDate") = Trim(Approveddate_)
        tmpRec("ApproveStaffID") = Trim(ApprovedStaffID_)
        
        tmpRec!LastUser = Trim(Trim(CurrentUser.StaffFName) & " " & Trim(CurrentUser.StaffLName))
        tmpRec!LastUpdate = GetSystemDateTime.SystemDateTime
        
        tmpRec.Update
    End If
    Set tmpRec = Nothing
    isChanged_ = False
    CloseDBConnection
End Sub

Public Sub Save()
'On Error GoTo ErrD
Dim tmpRec As ADODB.Recordset
Dim x%
Dim sqlStr As String
    sqlStr = "SELECT *, CONVERT(DATETIME, GETDATE(), 103) AS GETDATE FROM tbQuotationHeader WHERE QuotationID='" & id_ & "'"
    Set tmpRec = New ADODB.Recordset
    GetDBConnection
    tmpRec.open sqlStr, DBConn, adOpenDynamic, adLockOptimistic
    If tmpRec.EOF Then
        tmpRec.AddNew
        tmpRec("QuotationID") = id_
    End If
    tmpRec("IssueDate") = Trim(IssuedDate_)
    tmpRec("AttentionTo") = Trim(AttentionTo_)
    tmpRec("CustomerID") = Trim(CustomerID_)
    tmpRec("IssuedStaffID") = Trim(issuedStaffID_)
    tmpRec("isCancel") = Trim(isCancel_)
    tmpRec("isBmeNew") = Trim(isBmeNew_)
    tmpRec("CancelDate") = Trim(cancelDate_)
    tmpRec("isApproved") = Trim(isApproved_)
    tmpRec("ApprovedDate") = Trim(Approveddate_)
    tmpRec("ApproveStaffID") = Trim(ApprovedStaffID_)
    tmpRec("DiscountAmt") = Trim(discountAmt_)
    tmpRec("DisCountType") = Trim(discountType_)
    tmpRec("CancelStaffID") = Trim(cancelStaffID_)
    tmpRec("isPriceIncTransport") = Trim(isPriceIncTransport_)
    tmpRec("StaffCountPerCar") = StaffCountPerCar_
    tmpRec("MinWeightPerCar") = minWeightPerCar_
    tmpRec("CompanyID") = companyID_
    tmpRec("ExtraLaborFee") = extraLaborFee_
    tmpRec("isPriceIncVAT") = isPriceIncVAT_
    tmpRec("TruckTypeID") = truckTypeID_
    tmpRec("TransportFee") = transferRate_
    tmpRec("TransferUnitID") = transferUnitID_
    tmpRec("SignName") = Trim(signName_)
    tmpRec("SignPosition") = Trim(signPosition_)
    If Trim(isTrailer_) = "" Then isTrailer_ = "N"
    tmpRec("isTrailer") = Trim(isTrailer_)
    tmpRec("NonMinRate") = nonMinRate_
    tmpRec("HazMinRate") = hazMinRate_
    tmpRec("NBLMinRate") = nblMinRate_
    tmpRec("HBLMinRate") = hblMinRate_
    If Trim(isClosed_) = "" Then isClosed_ = "N"
    tmpRec("isClosed") = Trim(isClosed_)
    tmpRec("CloseDate") = Trim(closedDate_)
    tmpRec("CloseStaffID") = Trim(closeStaffID_)
    tmpRec("QuoRem") = Trim(Remdesc_)
    tmpRec("Promies") = Trim(Promies_)
    tmpRec("StartPromiesDate") = Trim(StartPromiesDate_)
    tmpRec("EndPromiesDate") = Trim(EndPromiesDate_)
    tmpRec("ProjectID") = Trim(ProjectID_)
    If LastUpdateUser_ <> CurrentUser.StaffFName + " " + CurrentUser.StaffLName Then LastUpdateUser_ = CurrentUser.StaffFName + " " + CurrentUser.StaffLName
    tmpRec!LastUpdateUser = LastUpdateUser_
    tmpRec!QuotationType = QuotationType_
    tmpRec!isQuotationReq = isQuotationReq_
    tmpRec!BMEQuotation = BMEQuotation_
    tmpRec!TransRate = TransRate_
    If Trim(isJob_) = "" Then
        tmpRec!isJobs = "N"
    Else
        tmpRec!isJobs = isJob_
    End If
    tmpRec!WasteCaseID = isCase_
    tmpRec!MaxWeight = MaxWeight_
    tmpRec!MaxRate = MaxRate_
    tmpRec!MaxRateType = MaxRateType_
    tmpRec!RefQuotationID = RefQuotationID_
    
    tmpRec!LastUser = Trim(Trim(CurrentUser.StaffFName) & " " & Trim(CurrentUser.StaffLName))
    tmpRec!LastUpdate = GetSystemDateTime.SystemDateTime
    
    tmpRec.Update
    tmpRec.Requery
    QuotationNo_ = Trim("" & tmpRec("QuotationNo"))
'    tmpRec.Close
        DBConnExc "Update tbQuotationDetail Set MarkAsDelete='Y' WHERE QuotationID='" & id_ & "'"
        DBConnExc "Update tbQuotationRemark Set MarkAsDelete='Y' WHERE QuotationID='" & id_ & "'"
        DBConnExc "Update tbJobProject Set ProjectStatus='Q' WHERE ProjectID='" & ProjectID_ & "'"
        If Not colQuotationDet_ Is Nothing Then
            For x = 1 To colQuotationDet_.Count
                If QuotationType_ = 1 Then
                    colQuotationDet_(x).BlendingRate = 0
                    colQuotationDet_(x).ParameterRate = 0
                End If
                colQuotationDet_(x).Save
            Next
        End If
        If Not colRemarks_ Is Nothing Then
            For x = 1 To colRemarks_.Count
                colRemarks_(x).Save
            Next
        End If
        DBConnExc "Delete tbQuotationDetail WHERE MarkAsDelete='Y' And QuotationID='" & id_ & "'"
        DBConnExc "Delete tbQuotationRemark WHERE MarkAsDelete='Y' And QuotationID='" & id_ & "'"
    Set tmpRec = Nothing
    isChanged_ = False
    CloseDBConnection
    Exit Sub
ErrD:
    GetMsg "", vbCritical
    Exit Sub
End Sub
Public Sub SaveQuoTruck()
'On Error GoTo ErrD
Dim tmpRec As ADODB.Recordset
Dim x%
Dim sqlStr As String
    sqlStr = "SELECT *, CONVERT(DATETIME, GETDATE(), 103) AS GETDATE FROM tbQuotationTruck WHERE QuotationID='" & id_ & "'"
    Set tmpRec = New ADODB.Recordset
    GetDBConnection
    tmpRec.open sqlStr, DBConn, adOpenDynamic, adLockOptimistic
    If tmpRec.EOF Then
        tmpRec.AddNew
        tmpRec("QuotationID") = id_
    End If
    tmpRec("TRuck1ID") = Trim(TRuck1ID_)
    tmpRec("Truck1Trailer") = Trim(Truck1Trailer_)
    If Truck1Rate_ <> "" Then
        tmpRec("Truck1Rate") = Trim(Truck1Rate_)
    End If
    tmpRec("Truck1Unit") = Trim(Truck1Unit_)
    tmpRec("TRuck2ID") = Trim(TRuck2ID_)
    tmpRec("TRuck2Trailer") = Trim(Truck2Trailer_)
    If Truck2Rate_ <> "" Then
        tmpRec("TRuck2Rate") = Trim(Truck2Rate_)
    End If
    tmpRec("TRuck2Unit") = Trim(Truck2Unit_)
    tmpRec("TRuck3ID") = Trim(TRuck3ID_)
    tmpRec("TRuck3Trailer") = Trim(Truck3Trailer_)
    If Truck3Rate_ <> "" Then
        tmpRec("TRuck3Rate") = Trim(Truck3Rate_)
    End If
    tmpRec("TRuck3Unit") = Trim(Truck3Unit_)
    tmpRec("TRuck4ID") = Trim(TRuck4ID_)
    tmpRec("TRuck4Trailer") = Trim(Truck4Trailer_)
    If Truck4Rate_ <> "" Then
        tmpRec("TRuck4Rate") = Trim(Truck4Rate_)
    End If
    tmpRec("TRuck4Unit") = Trim(Truck4Unit_)
    tmpRec("TRuck5ID") = Trim(TRuck5ID_)
    tmpRec("TRuck5Trailer") = Trim(Truck5Trailer_)
    If Truck5Rate_ <> "" Then
        tmpRec("TRuck5Rate") = Trim(Truck5Rate_)
    End If
    tmpRec("TRuck5Unit") = Trim(Truck5Unit_)
    tmpRec("TRuck6ID") = Trim(TRuck6ID_)
    tmpRec("TRuck6Trailer") = Trim(Truck6Trailer_)
    If Truck6Rate_ <> "" Then
        tmpRec("TRuck6Rate") = Trim(Truck6Rate_)
    End If
    tmpRec("TRuck6Unit") = Trim(Truck6Unit_)
    tmpRec("TRuck7ID") = Trim(TRuck7ID_)
    tmpRec("TRuck7Trailer") = Trim(Truck7Trailer_)
    If Truck7Rate_ <> "" Then
        tmpRec("TRuck7Rate") = Trim(Truck7Rate_)
    End If
    tmpRec("TRuck7Unit") = Trim(Truck7Unit_)
    tmpRec("TRuck8ID") = Trim(TRuck8ID_)
    tmpRec("TRuck8Trailer") = Trim(Truck8Trailer_)
    If Truck8Rate_ <> "" Then
        tmpRec("TRuck8Rate") = Trim(Truck8Rate_)
    End If
    tmpRec("TRuck8Unit") = Trim(Truck8Unit_)
    tmpRec("TRuck9ID") = Trim(TRuck9ID_)
    tmpRec("TRuck9Trailer") = Trim(Truck9Trailer_)
    If Truck9Rate_ <> "" Then
        tmpRec("TRuck9Rate") = Trim(Truck9Rate_)
    End If
    tmpRec("TRuck9Unit") = Trim(Truck9Unit_)

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
    DBConnExc "Delete tbQuotationHeader WHERE QuotationID='" & id_ & "'"
    Delete = ""
    Exit Function
delErr:
    Delete = err.Description
End Function

Public Sub SetDataFromRecordset(selRec As ADODB.Recordset)
    id_ = Trim("" & selRec("QuotationID"))
    QuotationNo_ = Trim("" & selRec("QuotationNo"))
    BMEQuotation_ = Trim("" & selRec("BMEQuotation"))
    IssuedDate_ = Trim("" & selRec("IssueDate"))
    AttentionTo_ = Trim("" & selRec("AttentionTo"))
    CustomerID_ = Trim("" & selRec("CustomerID"))
    CustomerCode_ = Trim("" & selRec("CustomerCode"))
    CustomerName_ = Trim("" & selRec("CompanyName"))
    issuedStaffID_ = Trim("" & selRec("IssuedStaffID"))
    issuedStaffName_ = Trim("" & selRec("IssuedStaffName"))
    isCancel_ = Trim("" & selRec("isCancel"))
    isBmeNew_ = Trim("" & selRec("isBmeNew"))
    cancelDate_ = Trim("" & selRec("CancelDate"))
    isApproved_ = Trim("" & selRec("isApproved"))
    Approveddate_ = Trim("" & selRec("ApprovedDate"))
    ApprovedStaffID_ = Trim("" & selRec("ApproveStaffID"))
    ApprovedStaffName_ = Trim("" & selRec("ApprovedStaffName"))
    If IsNull(selRec("DiscountAmt")) = False Then discountAmt_ = selRec("DiscountAmt")
    discountType_ = Trim("" & selRec("DisCountType"))
    cancelStaffID_ = Trim("" & selRec("CancelStaffID"))
    cancelStaffName_ = Trim("" & selRec("CancelStaffName"))
    isPriceIncTransport_ = Trim("" & selRec("isPriceIncTransport"))
    companyID_ = Trim("" & selRec("CompanyID"))
    
    truckTypeID_ = Trim("" & selRec("TruckTypeID"))
    truckTypeName_ = Trim("" & selRec("TruckTypeDesc"))
    If IsNull(selRec("TransportFee")) = False Then transferRate_ = selRec("TransportFee")
    transferUnitID_ = Trim("" & selRec("TransferUnitID"))
    transferUnitName_ = Trim("" & selRec("TransferUnitID"))
    
    If IsNull(selRec("ExtraLaborFee")) = False Then extraLaborFee_ = selRec("ExtraLaborFee")
    isPriceIncVAT_ = Trim("" & selRec("isPriceIncVAT"))
    If IsNull(selRec("StaffCountPerCar")) = False Then StaffCountPerCar_ = selRec("StaffCountPerCar")
    If IsNull(selRec("MinWeightPerCar")) = False Then minWeightPerCar_ = selRec("MinWeightPerCar")
    signName_ = Trim("" & selRec("SignName"))
    signPosition_ = Trim("" & selRec("SignPosition"))
    isTrailer_ = Trim("" & selRec("isTrailer"))
    
    isClosed_ = Trim("" & selRec("isClosed"))
    closedDate_ = Trim("" & selRec("CloseDate"))
    closeStaffID_ = Trim("" & selRec("CloseStaffID"))
    closeStaffName_ = Trim("" & selRec("CloseStaffFName")) & " " & Trim("" & selRec("CloseStaffLName"))
    Remdesc_ = Trim("" & selRec("QuoRem"))
    nonMinRate_ = selRec("NonMinRate")
    hazMinRate_ = selRec("HazMinRate")
    nblMinRate_ = selRec("NBLMinRate")
    hblMinRate_ = selRec("HBLMinRate")
    QuotationType_ = CDbl(selRec!QuotationType)
    isQuotationReq_ = IIf(IsNull(selRec!isQuotationReq), "N", selRec!isQuotationReq)
    Promies_ = Trim("" & selRec("Promies"))
    StartPromiesDate_ = Trim("" & selRec("StartPromiesDate"))
    EndPromiesDate_ = Trim("" & selRec("EndPromiesDate"))
   RefQuotationID_ = IIf(IsNull(selRec!RefQuotationID), "", selRec!RefQuotationID)
   
   TRuck1ID_ = IIf(IsNull(selRec!TRuck1ID), "", selRec!TRuck1ID)
   Truck1Trailer_ = IIf(IsNull(selRec!Truck1Trailer), "", selRec!Truck1Trailer)
   Truck1Rate_ = IIf(IsNull(selRec!Truck1Rate), "", selRec!Truck1Rate)
   Truck1Unit_ = IIf(IsNull(selRec!Truck1Unit), "", selRec!Truck1Unit)
   
   TRuck2ID_ = IIf(IsNull(selRec!TRuck2ID), "", selRec!TRuck2ID)
   Truck2Trailer_ = IIf(IsNull(selRec!Truck2Trailer), "", selRec!Truck2Trailer)
   Truck2Rate_ = IIf(IsNull(selRec!Truck2Rate), "", selRec!Truck2Rate)
   Truck2Unit_ = IIf(IsNull(selRec!Truck2Unit), "", selRec!Truck2Unit)

   TRuck3ID_ = IIf(IsNull(selRec!TRuck3ID), "", selRec!TRuck3ID)
   Truck3Trailer_ = IIf(IsNull(selRec!TRuck3Trailer), "", selRec!TRuck3Trailer)
   Truck3Rate_ = IIf(IsNull(selRec!TRuck3Rate), "", selRec!TRuck3Rate)
   Truck3Unit_ = IIf(IsNull(selRec!TRuck3Unit), "", selRec!TRuck3Unit)

   TRuck4ID_ = IIf(IsNull(selRec!TRuck4ID), "", selRec!TRuck4ID)
   Truck4Trailer_ = IIf(IsNull(selRec!TRuck4Trailer), "", selRec!TRuck4Trailer)
   Truck4Rate_ = IIf(IsNull(selRec!TRuck4Rate), "", selRec!TRuck4Rate)
   Truck4Unit_ = IIf(IsNull(selRec!TRuck4Unit), "", selRec!TRuck4Unit)

   TRuck5ID_ = IIf(IsNull(selRec!TRuck5ID), "", selRec!TRuck5ID)
   Truck5Trailer_ = IIf(IsNull(selRec!TRuck5Trailer), "", selRec!TRuck5Trailer)
   Truck5Rate_ = IIf(IsNull(selRec!TRuck5Rate), "", selRec!TRuck5Rate)
   Truck5Unit_ = IIf(IsNull(selRec!TRuck5Unit), "", selRec!TRuck5Unit)

   TRuck6ID_ = IIf(IsNull(selRec!TRuck6ID), "", selRec!TRuck6ID)
   Truck6Trailer_ = IIf(IsNull(selRec!TRuck6Trailer), "", selRec!TRuck6Trailer)
   Truck6Rate_ = IIf(IsNull(selRec!TRuck6Rate), "", selRec!TRuck6Rate)
   Truck6Unit_ = IIf(IsNull(selRec!TRuck6Unit), "", selRec!TRuck6Unit)

   TRuck7ID_ = IIf(IsNull(selRec!TRuck7ID), "", selRec!TRuck7ID)
   Truck7Trailer_ = IIf(IsNull(selRec!TRuck7Trailer), "", selRec!TRuck7Trailer)
   Truck7Rate_ = IIf(IsNull(selRec!TRuck7Rate), "", selRec!TRuck7Rate)
   Truck7Unit_ = IIf(IsNull(selRec!TRuck7Unit), "", selRec!TRuck7Unit)

   TRuck8ID_ = IIf(IsNull(selRec!TRuck8ID), "", selRec!TRuck8ID)
   Truck8Trailer_ = IIf(IsNull(selRec!TRuck8Trailer), "", selRec!TRuck8Trailer)
   Truck8Rate_ = IIf(IsNull(selRec!TRuck8Rate), "", selRec!TRuck8Rate)
   Truck8Unit_ = IIf(IsNull(selRec!TRuck8Unit), "", selRec!TRuck8Unit)

   TRuck9ID_ = IIf(IsNull(selRec!TRuck9ID), "", selRec!TRuck9ID)
   Truck9Trailer_ = IIf(IsNull(selRec!TRuck9Trailer), "", selRec!TRuck9Trailer)
   Truck9Rate_ = IIf(IsNull(selRec!TRuck9Rate), "", selRec!TRuck9Rate)
   Truck9Unit_ = IIf(IsNull(selRec!TRuck9Unit), "", selRec!TRuck9Unit)

    If IsNull(selRec("isJobs")) = False Then isJob_ = Trim("" & selRec("isJobs"))
    If IsNull(selRec("WasteCaseID")) = False Then isCase_ = Trim("" & selRec("WasteCaseID"))
    
    If IsNull(selRec("LastUpdateUser")) = False Then LastUpdateUser_ = selRec("LastUpdateUser")
    If Trim(LastUpdateUser_) = "" Then LastUpdateUser_ = issuedStaffName_
    If isExist(selRec, "APSale") Then SaleAP_ = IIf(IsNull(selRec!APSale), "", selRec!APSale)
    
    If isExist(selRec, "MAXWeight") Then MaxWeight_ = IIf(IsNull(selRec!MaxWeight), "0", selRec!MaxWeight)
    If isExist(selRec, "MAXRate") Then MaxRate_ = IIf(IsNull(selRec!MaxRate), "0", selRec!MaxRate)
    If isExist(selRec, "MAXRateType") Then MaxRateType_ = IIf(IsNull(selRec!MaxRateType), "", selRec!MaxRateType)
    
    If isExist(selRec, "TransRate") Then TransRate_ = IIf(IsNull(selRec!TransRate), "", selRec!TransRate)
    If isExist(selRec, "ProjectID") Then ProjectID_ = IIf(IsNull(selRec!ProjectID), "", selRec!ProjectID)
    
End Sub

Public Sub CancelQuotation()
    isCancel_ = "Y"
    cancelDate_ = TodayDate
    cancelStaffID_ = CurrentUser.ID
    cancelStaffName_ = CurrentUser.StaffFName & " " & CurrentUser.StaffLName
    DBConnExc "Update tbQuotationHeader Set isCancel='Y',CancelDate='" & TodayDate & "',CancelStaffID='" & cancelStaffID_ & "',QuoRem='" & Replace(Remdesc_, "'", "") & "' WHERE QuotationID='" & id_ & "'"
    isChanged_ = False
End Sub

Public Sub CloseQuotation()
    isClosed_ = "Y"
    closedDate_ = TodayDate
    closeStaffID_ = CurrentUser.ID
    closeStaffName_ = CurrentUser.StaffFName & " " & CurrentUser.StaffLName
    DBConnExc "Update tbQuotationHeader Set isClosed='Y',CloseDate='" & TodayDate & "',CloseStaffID='" & closeStaffID_ & "',QuoRem='" & Replace(Remdesc_, "'", "") & "' WHERE QuotationID='" & id_ & "'"
    isChanged_ = False
End Sub

Public Function WantApprove() As String
On Error GoTo ErrD
Dim i As Integer
Dim tSql As String
    WantApprove = ""
    For i = 1 To colQuotationDet_.Count
        tSql = "UPDATE tbWasteDataCR SET Status = 'W' WHERE WasteDataID = '" & colQuotationDet_(i).wastedataID & "' AND (Status <> 'P') AND (Status <> 'R')"
'        DBConnExc tSql
    Next
    Exit Function
ErrD:
    WantApprove = "เกิดข้อผิดพลาด ระหว่างส่งข้อมูลให้สิ่งแวดล้อม ลองคลิกที่ปุ่ม (อนุมัติใบเสนอราคา) อีกครั้ง"
    Exit Function
End Function
