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
Dim quotationNo_ As String
Dim issuedDate_ As String
Dim attentionTo_ As String
Dim customerID_ As String
Dim customerName_ As String
Dim issuedStaffID_ As String
Dim issuedStaffName_ As String
Dim isCancel_ As String
Dim cancelDate_ As String
Dim isApproved_ As String
Dim approvedDate_ As String
Dim approvedStaffID_ As String
Dim approvedStaffName_ As String
Dim discountAmt_ As Double
Dim discountType_ As String
Dim cancelStaffID_ As String
Dim cancelStaffName_ As String
Dim isPriceIncTransport_ As String
Dim isPriceIncVAT_ As String
Dim staffCountPerCar_ As Integer
Dim minWeightPerCar_ As Double
Dim colQuotationDet_ As Collection
Dim colRemarks_ As Collection
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
Dim remDesc_ As String
Dim isChanged_ As Boolean

Public Property Get ID() As String
    ID = id_
End Property
Public Property Let ID(ByVal newval As String)
    id_ = Trim(newval)
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

Public Property Get companyID() As String
    companyID = companyID_
End Property
Public Property Let companyID(ByVal newval As String)
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

Public Property Get RemDesc() As String
    RemDesc = remDesc_
End Property
Public Property Let RemDesc(ByVal newval As String)
    If Trim(remDesc_) <> Trim(newval) Then
        remDesc_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get IssuedDate() As String
    IssuedDate = issuedDate_
End Property
Public Property Let IssuedDate(ByVal newval As String)
    If Trim(issuedDate_) <> Trim(newval) Then
        issuedDate_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get AttentionTo() As String
    AttentionTo = attentionTo_
End Property
Public Property Let AttentionTo(ByVal newval As String)
    If Trim(attentionTo_) <> Trim(newval) Then
        attentionTo_ = Trim(newval)
        isChanged_ = True
    End If
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

Public Property Get IsApproved() As Boolean
    If Trim(isApproved_) = "Y" Then
        IsApproved = True
    Else
        IsApproved = False
    End If
End Property
Public Property Let IsApproved(ByVal newval As Boolean)
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

Public Property Get ApprovedDate() As String
    ApprovedDate = approvedDate_
End Property
Public Property Let ApprovedDate(ByVal newval As String)
    If Trim(approvedDate_) <> Trim(newval) Then
        approvedDate_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get ApprovedStaffID() As String
    ApprovedStaffID = approvedStaffID_
End Property
Public Property Let ApprovedStaffID(ByVal newval As String)
    If Trim(approvedStaffID_) <> Trim(newval) Then
        approvedStaffID_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get ApprovedStaffName() As String
    ApprovedStaffName = approvedStaffName_
End Property
Public Property Let ApprovedStaffName(ByVal newval As String)
    If Trim(approvedStaffName_) <> Trim(newval) Then
        approvedStaffName_ = Trim(newval)
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
    StaffCountPerCar = staffCountPerCar_
End Property
Public Property Let StaffCountPerCar(ByVal newval As Double)
    If staffCountPerCar_ <> newval Then
        staffCountPerCar_ = newval
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

Public Property Get QuotationDetails() As Collection
Dim tmpS As New BWGSearchManager
    If colQuotationDet_ Is Nothing Then Set colQuotationDet_ = tmpS.QuotationDetailSearch("QuotationID='" & id_ & "'")
    Set QuotationDetails = colQuotationDet_
End Property
Public Property Set QuotationDetails(ByVal newval As Collection)
    Set colQuotationDet_ = newval
    isChanged_ = True
End Property
Public Sub RequeryQuotationDetail()
    Set colQuotationDet_ = Nothing
End Sub





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

Public Property Get isChanged() As Boolean
    isChanged = isChanged_
End Property
Public Property Let isChanged(ByVal newval As Boolean)
    isChanged_ = newval
End Property

Public Sub Save()
On Error GoTo ErrD
Dim tmpRec As ADODB.Recordset
Dim x%
    Set tmpRec = New ADODB.Recordset
    tmpRec.Open "Select * from tbQuotationHeader Where QuotationID='" & id_ & "'", DBManager.GetDBConnection, adOpenDynamic, adLockOptimistic
    If tmpRec.EOF Then
        tmpRec.AddNew
        tmpRec("QuotationID") = id_
    End If
    tmpRec("IssueDate") = Trim(issuedDate_)
    tmpRec("AttentionTo") = Trim(attentionTo_)
    tmpRec("CustomerID") = Trim(customerID_)
    tmpRec("IssuedStaffID") = Trim(issuedStaffID_)
    tmpRec("isCancel") = Trim(isCancel_)
    tmpRec("CancelDate") = Trim(cancelDate_)
    tmpRec("isApproved") = Trim(isApproved_)
    tmpRec("ApprovedDate") = Trim(approvedDate_)
    tmpRec("ApproveStaffID") = Trim(approvedStaffID_)
    tmpRec("DiscountAmt") = Trim(discountAmt_)
    tmpRec("DisCountType") = Trim(discountType_)
    tmpRec("CancelStaffID") = Trim(cancelStaffID_)
    tmpRec("isPriceIncTransport") = Trim(isPriceIncTransport_)
    tmpRec("StaffCountPerCar") = staffCountPerCar_
    tmpRec("MinWeightPerCar") = minWeightPerCar_
    tmpRec("CompanyID") = companyID_
    tmpRec("ExtraLaborFee") = extraLaborFee_
    tmpRec("isPriceIncVAT") = isPriceIncVAT_
    tmpRec("TruckTypeID") = truckTypeID_
    tmpRec("TransportFee") = transferRate_
    tmpRec("TrasnferUnitID") = transferUnitID_
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
    tmpRec("QuoRem") = Trim(remDesc_)
    
    tmpRec.Update
    tmpRec.Requery
    quotationNo_ = Trim("" & tmpRec("QuotationNo"))
    tmpRec.Close
    With DBManager.GetDBConnection
        .Execute "Update tbQuotationDetail Set MarkAsDelete='Y' Where QuotationID='" & id_ & "'"
        .Execute "Update tbQuotationRemark Set MarkAsDelete='Y' Where QuotationID='" & id_ & "'"
        If Not colQuotationDet_ Is Nothing Then
            For x = 1 To colQuotationDet_.Count
                colQuotationDet_(x).Save
            Next
        End If
        If Not colRemarks_ Is Nothing Then
            For x = 1 To colRemarks_.Count
                colRemarks_(x).Save
            Next
        End If
        .Execute "Delete tbQuotationDetail Where MarkAsDelete='Y' And QuotationID='" & id_ & "'"
        .Execute "Delete tbQuotationRemark Where MarkAsDelete='Y' And QuotationID='" & id_ & "'"
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
    DBManager.GetDBConnection.Execute "Delete tbQuotationHeader Where QuotationID='" & id_ & "'"
    Delete = ""
    Exit Function
delErr:
    Delete = Err.Description
End Function

Public Sub SetDataFromRecordset(selRec As ADODB.Recordset)
    id_ = Trim("" & selRec("QuotationID"))
    quotationNo_ = Trim("" & selRec("QuotationNo"))
    issuedDate_ = Trim("" & selRec("IssueDate"))
    attentionTo_ = Trim("" & selRec("AttentionTo"))
    customerID_ = Trim("" & selRec("CustomerID"))
    customerName_ = Trim("" & selRec("CompanyName"))
    issuedStaffID_ = Trim("" & selRec("IssuedStaffID"))
    issuedStaffName_ = Trim("" & selRec("IssuedStaffName"))
    isCancel_ = Trim("" & selRec("isCancel"))
    cancelDate_ = Trim("" & selRec("CancelDate"))
    isApproved_ = Trim("" & selRec("isApproved"))
    approvedDate_ = Trim("" & selRec("ApprovedDate"))
    approvedStaffID_ = Trim("" & selRec("ApproveStaffID"))
    approvedStaffName_ = Trim("" & selRec("ApprovedStaffName"))
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
    If IsNull(selRec("StaffCountPerCar")) = False Then staffCountPerCar_ = selRec("StaffCountPerCar")
    If IsNull(selRec("MinWeightPerCar")) = False Then minWeightPerCar_ = selRec("MinWeightPerCar")
    signName_ = Trim("" & selRec("SignName"))
    signPosition_ = Trim("" & selRec("SignPosition"))
    isTrailer_ = Trim("" & selRec("isTrailer"))
    
    isClosed_ = Trim("" & selRec("isClosed"))
    closedDate_ = Trim("" & selRec("CloseDate"))
    closeStaffID_ = Trim("" & selRec("CloseStaffID"))
    closeStaffName_ = Trim("" & selRec("CloseStaffFName")) & " " & Trim("" & selRec("CloseStaffLName"))
    remDesc_ = Trim("" & selRec("QuoRem"))
    
    nonMinRate_ = selRec("NonMinRate")
    hazMinRate_ = selRec("HazMinRate")
    nblMinRate_ = selRec("NBLMinRate")
    hblMinRate_ = selRec("HBLMinRate")
End Sub

Public Sub CancelQuotation()
    isCancel_ = "Y"
    cancelDate_ = TodayDate
    cancelStaffID_ = CurrentUser.ID
    cancelStaffName_ = CurrentUser.StaffFName & " " & CurrentUser.StaffLName
    DBManager.GetDBConnection.Execute "Update tbQuotationHeader Set isCancel='Y',CancelDate='" & TodayDate & "',CancelStaffID='" & cancelStaffID_ & "',QuoRem='" & Replace(remDesc_, "'", "") & "' Where QuotationID='" & id_ & "'"
    isChanged_ = False
End Sub

Public Sub CloseQuotation()
    isClosed_ = "Y"
    closedDate_ = TodayDate
    closeStaffID_ = CurrentUser.ID
    closeStaffName_ = CurrentUser.StaffFName & " " & CurrentUser.StaffLName
    DBManager.GetDBConnection.Execute "Update tbQuotationHeader Set isClosed='Y',CloseDate='" & TodayDate & "',CloseStaffID='" & closeStaffID_ & "',QuoRem='" & Replace(remDesc_, "'", "") & "' Where QuotationID='" & id_ & "'"
    isChanged_ = False
End Sub

