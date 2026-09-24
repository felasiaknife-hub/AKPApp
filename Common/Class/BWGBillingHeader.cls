VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BWGBillingHeader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim id_ As String
Dim IssuedDate_ As String
Dim CustomerID_ As String
Dim CustomerName_ As String
Dim companyID_ As String
Dim companyName_ As String
Dim BillingNo_ As String
Dim invoiceNo_ As String
Dim receiptNo_ As String
Dim payDueDate_ As String
Dim discountAmt_ As Double
Dim AmtB4VAT_ As Double
Dim VATAmt_ As Double
Dim TotalBillAmt_ As Double
Dim billRemark_ As String
Dim isCancel_ As String
Dim cancelDate_ As String
Dim CancelStaff_ As String
Dim CancelDetail_ As String
Dim discountType_ As String
Dim wTaxRate_ As Double
Dim paidAmt_ As Double
Dim isCustAccepted_ As String
Dim billAddrID_ As String
Dim billAddr_ As String
Dim billRecStaff_ As String
Dim colBillDetail_ As Collection
Dim colTripControl_ As Collection
Dim isChanged_ As Boolean
Dim isSave_ As Boolean
Dim IncVAT_ As Integer
Dim LastUser_ As String
Dim LastUpdate_ As String
Dim LastUserRE_ As String
Dim LastUpdateRE_ As String
Dim DatePast_ As String
Dim isFiled_ As String
Dim isServiceCustomer_ As String
Dim IsInv_ As String
Dim invDate_ As String
Dim PreInvNo_ As String
Dim PreDate_ As String
Dim PreNote_ As String
Dim PreStaff_ As String

Public Property Get ID() As String
    ID = id_
End Property
Public Property Let ID(ByVal newval As String)
    id_ = Trim(newval)
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

Public Property Get CustomerID() As String
    CustomerID = CustomerID_
End Property
Public Property Let CustomerID(ByVal newval As String)
    If Trim(CustomerID_) <> Trim(newval) Then
        CustomerID_ = Trim(newval)
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

Public Property Get CompanyID() As String
    CompanyID = companyID_
End Property
Public Property Let CompanyID(ByVal newval As String)
    If Trim(companyID_) <> Trim(newval) Then
        companyID_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get CompanyName() As String
    CompanyName = companyName_
End Property
Public Property Let CompanyName(ByVal newval As String)
    If Trim(companyName_) <> Trim(newval) Then
        companyName_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get BillingNo() As String
    BillingNo = BillingNo_
End Property
Public Property Let BillingNo(ByVal newval As String)
    If Trim(BillingNo_) <> Trim(newval) Then
        BillingNo_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get InvoiceNo() As String
    InvoiceNo = invoiceNo_
End Property
Public Property Let InvoiceNo(ByVal newval As String)
    If Trim(invoiceNo_) <> Trim(newval) Then
        invoiceNo_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get ReceiptNo() As String
    ReceiptNo = receiptNo_
End Property
Public Property Let ReceiptNo(ByVal newval As String)
    If Trim(receiptNo_) <> Trim(newval) Then
        receiptNo_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get PayDueDate() As String
    PayDueDate = payDueDate_
End Property
Public Property Let PayDueDate(ByVal newval As String)
    If Trim(payDueDate_) <> Trim(newval) Then
        payDueDate_ = Trim(newval)
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

Public Property Get AmtB4VAT() As Double
    AmtB4VAT = AmtB4VAT_
End Property
Public Property Let AmtB4VAT(ByVal newval As Double)
    If AmtB4VAT_ <> newval Then
        AmtB4VAT_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get VATAmt() As Double
    VATAmt = VATAmt_
End Property
Public Property Let VATAmt(ByVal newval As Double)
    If VATAmt_ <> newval Then
        VATAmt_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get TotalBillAmt() As Double
    TotalBillAmt = TotalBillAmt_
End Property
Public Property Let TotalBillAmt(ByVal newval As Double)
    If TotalBillAmt_ <> newval Then
        TotalBillAmt_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get BillRemark() As String
    BillRemark = billRemark_
End Property
Public Property Let BillRemark(ByVal newval As String)
    If Trim(billRemark_) <> Trim(newval) Then
        billRemark_ = Trim(newval)
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

Public Property Get CancelStaff() As String
    CancelStaff = CancelStaff_
End Property
Public Property Let CancelStaff(ByVal newval As String)
    If Trim(CancelStaff_) <> Trim(newval) Then
        CancelStaff_ = Trim(newval)
    End If
End Property

Public Property Get CancelDetail() As String
    CancelDetail = CancelDetail_
End Property
Public Property Let CancelDetail(ByVal newval As String)
    If Trim(CancelDetail_) <> Trim(newval) Then
        CancelDetail_ = Trim(newval)
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

Public Property Get WTaxRate() As Double
    WTaxRate = wTaxRate_
End Property
Public Property Let WTaxRate(ByVal newval As Double)
    If wTaxRate_ <> newval Then
        wTaxRate_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get PaidAmt() As Double
    PaidAmt = paidAmt_
End Property
Public Property Let PaidAmt(ByVal newval As Double)
    If paidAmt_ <> newval Then
        paidAmt_ = newval
        isChanged_ = True
    End If
End Property
Public Property Get isServiceCustomer() As Boolean
    If Trim(isServiceCustomer_) = "Y" Then
        isServiceCustomer = True
    Else
        isServiceCustomer = False
    End If
End Property
Public Property Let isServiceCustomer(ByVal newval As Boolean)
Dim tmpStr As String
    If newval Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(isServiceCustomer_) <> Trim(tmpStr) Then
        isServiceCustomer_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property

Public Property Get isCustAccepted() As Boolean
    If Trim(isCustAccepted_) = "Y" Then
        isCustAccepted = True
    Else
        isCustAccepted = False
    End If
End Property
Public Property Let isCustAccepted(ByVal newval As Boolean)
Dim tmpStr As String
    If newval Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(isCustAccepted_) <> Trim(tmpStr) Then
        isCustAccepted_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property

Public Property Get isFiled() As Boolean
 If Trim(isFiled_) = "Y" Then
        isFiled = True
    Else
        isFiled = False
    End If
End Property
Public Property Let isFiled(ByVal newval As Boolean)
    Dim tmpStr As String
    If newval Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(isFiled_) <> Trim(tmpStr) Then
        isFiled_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property

Public Property Get BillAddrID() As String
    BillAddrID = billAddrID_
End Property
Public Property Let BillAddrID(ByVal newval As String)
    If Trim(billAddrID_) <> Trim(newval) Then
        billAddrID_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get BillAddr() As String
    BillAddr = billAddr_
End Property
Public Property Let BillAddr(ByVal newval As String)
    If Trim(billAddr_) <> Trim(newval) Then
        billAddr_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get BillRecStaff() As String
    BillRecStaff = billRecStaff_
End Property
Public Property Let BillRecStaff(ByVal newval As String)
    If Trim(billRecStaff_) <> Trim(newval) Then
        billRecStaff_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get BillingDetails() As Collection
Dim tmpS As New BWGSearchManager
    If colBillDetail_ Is Nothing Then Set colBillDetail_ = tmpS.BillingDetailSearch("BillingHeaderID='" & id_ & "'")
    Set tmpS = Nothing
    Set BillingDetails = colBillDetail_
End Property
Public Property Set BillingDetails(ByVal newval As Collection)
    Set colBillDetail_ = newval
    isChanged_ = True
End Property
Public Sub RequeryBillingDetail()
    Set colBillDetail_ = Nothing
End Sub

Public Property Get TripControls() As Collection
Dim tmpS As New BWGSearchManager
    If colTripControl_ Is Nothing Then Set colTripControl_ = tmpS.TripControlInBillSearch("BillingHeaderID='" & id_ & "'")
    Set tmpS = Nothing
    Set TripControls = colTripControl_
End Property
Public Property Set TripControls(ByVal newval As Collection)
    Set colTripControl_ = newval
    isChanged_ = True
End Property
Public Sub RequeryTripControl()
    Set colTripControl_ = Nothing
End Sub

Public Property Get isChanged() As Boolean
    isChanged = isChanged_
End Property
Public Property Let isChanged(ByVal newval As Boolean)
    isChanged_ = newval
End Property

Public Property Get isSave() As Boolean
    isSave = isSave_
End Property
Public Property Let isSave(ByVal newval As Boolean)
    isSave_ = newval
End Property

Public Property Get IncVAT() As Integer
    IncVAT = IncVAT_
End Property
Public Property Let IncVAT(ByVal newval As Integer)
    IncVAT_ = newval
    isChanged_ = True
End Property

Public Property Get LastUser() As String
    LastUser = LastUser_
End Property
Public Property Let LastUser(ByVal newval As String)
    If Trim(LastUser_) <> Trim(newval) Then
        LastUser_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get LastUpdate() As String
    LastUpdate = LastUpdate_
End Property
Public Property Let LastUpdate(ByVal newval As String)
    If Trim(LastUpdate_) <> Trim(newval) Then
        LastUpdate_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get LastUserRE() As String
    LastUserRE = LastUserRE_
End Property
Public Property Let LastUserRE(ByVal newval As String)
    If Trim(LastUserRE_) <> Trim(newval) Then
        LastUserRE_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get LastUpdateRE() As String
    LastUpdateRE = LastUpdateRE_
End Property
Public Property Let LastUpdateRE(ByVal newval As String)
    If Trim(LastUpdateRE_) <> Trim(newval) Then
        LastUpdateRE_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get DatePast() As String
    DatePast = DatePast_
End Property
Public Property Let DatePast(ByVal newval As String)
    If Trim(DatePast_) <> Trim(newval) Then
        DatePast_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get IsInv() As String
    IsInv = IsInv_
End Property
Public Property Let IsInv(ByVal newval As String)
    If Trim(IsInv_) <> Trim(newval) Then
        IsInv_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get InvDate() As String
    InvDate = invDate_
End Property
Public Property Let InvDate(ByVal newval As String)
    If Trim(invDate_) <> Trim(newval) Then
        invDate_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get PreInvNo() As String
    PreInvNo = PreInvNo_
End Property
Public Property Let PreInvNo(ByVal newval As String)
    If Trim(PreInvNo_) <> Trim(newval) Then
        PreInvNo_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get PreDate() As String
    PreDate = PreDate_
End Property
Public Property Let PreDate(ByVal newval As String)
    If Trim(PreDate_) <> Trim(newval) Then
        PreDate_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get PreNote() As String
    PreNote = PreNote_
End Property
Public Property Let PreNote(ByVal newval As String)
    If Trim(PreNote_) <> Trim(newval) Then
        PreNote_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get PreStaff() As String
    PreStaff = PreStaff_
End Property
Public Property Let PreStaff(ByVal newval As String)
    If Trim(PreStaff_) <> Trim(newval) Then
        PreStaff_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Sub Save()
On Error GoTo ErrD
Dim tmpRec As ADODB.Recordset
Dim x%
Dim SaveErr As Boolean
Dim ErrState As String, sqlStr As String
Dim CompleteDetailWithError As Boolean
    SaveErr = False
    isSave_ = True
'    DBConn.BeginTrans
    GetDBConnection
    Set tmpRec = New ADODB.Recordset
    
    sqlStr = "SELECT *, CONVERT(DATETIME, GETDATE(), 103) AS GETDATE FROM tbBillingHeader WHERE BillingHeaderID='" & id_ & "'"
    tmpRec.open sqlStr, DBConn, adOpenKeyset, adLockOptimistic
    If tmpRec.EOF Then
        tmpRec.AddNew
        tmpRec("BillingHeaderID") = id_
        tmpRec("isFiled") = "N"
    End If
    tmpRec("IssuedDate") = Trim(IssuedDate_)
    tmpRec("CustomerID") = Trim(CustomerID_)
    tmpRec("CompanyID") = Trim(companyID_)
    'tmpRec("BillingNo") = Trim(billingNo_)
    tmpRec("InvoiceNo") = Trim(invoiceNo_)
    tmpRec("ReceiptNo") = Trim(receiptNo_)
    tmpRec("PayDueDate") = Trim(payDueDate_)
    tmpRec("DiscountAmt") = Trim(discountAmt_)
    tmpRec("AmtB4VAT") = Trim(AmtB4VAT_)
    tmpRec("VATAmt") = Trim(VATAmt_)
    tmpRec("TotalBillAmt") = Trim(TotalBillAmt_)
    tmpRec("BillRemark") = Trim(billRemark_)
    If Trim(isCancel_) = "" Then isCancel_ = "N"
    tmpRec("isCancel") = Trim(isCancel_)
    'tmpRec("CancelDate") = Trim(cancelDate_)
    tmpRec("DiscountType") = Trim(discountType_)
    tmpRec("WTax") = Trim(wTaxRate_)
    tmpRec("isCustAccepted") = Trim(isCustAccepted_)
    tmpRec("isServiceCustomer") = Trim(isServiceCustomer_)
    tmpRec("isSendCustomer") = Trim(isServiceCustomer_)
    tmpRec("isFiled") = Trim(isFiled_)
    tmpRec("IsInv") = Trim(IsInv_)
    tmpRec("InvDate") = Trim(invDate_)
'    tmpRec("PreInvNo") = Trim(PreInvNo_)
    tmpRec("PreDate") = Trim(PreDate_)
    tmpRec("PreNote") = Trim(PreNote_)
    tmpRec("PreStaff") = Trim(PreStaff_)
    tmpRec("BillAddrID") = Trim(billAddrID_)
    tmpRec("BillRecStaff") = Trim(billRecStaff_)
    tmpRec!IncVAT = IIf(IncVAT_ = 0, "Y", "N")
    tmpRec("DatePast") = Trim(DatePast_)
    tmpRec!LastUser = Trim(Trim(CurrentUser.StaffFName) & " " & Trim(CurrentUser.StaffLName))
    tmpRec!LastUpdate = GetSystemDateTime.SystemDateTime
    tmpRec.Update
    tmpRec.Requery
    BillingNo_ = Trim("" & tmpRec("BillingNo"))
    PreInvNo_ = Trim("" & tmpRec("PreInvNo"))
'    tmpRec.Close

    ErrState = "กำลังบันทึกเลขที่ Invoice"
    Set tmpRec = Nothing
    CloseDBConnection
    If SaveErr = True Then Exit Sub
    '*** Initializing Status in Menifest IsINV = "N"
    ErrState = UpdateManifest(id_, "N", CurrentUser.ID, "")
    If ErrState <> "" Then
        MsgBox ErrState, vbCritical, "Error"
    End If
    DBConnExc "Update tbBillingDetail Set MarkForDelete='Y' WHERE BillingHeaderID='" & id_ & "'"
RetrySaveDetail:
    For x = 1 To colBillDetail_.Count
        colBillDetail_(x).Save x
        ErrState = "กำลังบันทึกรายละเอียดใน Invoice ลำดับที่ " & x
        If CompleteDetailWithError = False Then CompleteDetailWithError = colBillDetail_(x).CompleteWithError
        
        '***** Try to saving data deail.
        Dim i As Integer
        While CompleteDetailWithError = True
            If CompleteDetailWithError = True Then colBillDetail_(x).Save x
            CompleteDetailWithError = colBillDetail_(x).CompleteWithError
        Wend
        '***** End *****
        DoEvents
        frmMain.Stb1.Panels(1).Text = ErrState
        frmMain.Stb1.Panels(2).Text = "เมนิเฟสเลขที่ " & colBillDetail_(x).MenifestDocNo
    
        If CompleteDetailWithError = True Then
            MsgBox "เกิดข้อผิดพลาดขณะทำการบันทึกรายการ (" & colBillDetail_(x).MenifestDocNo & ") ใน Invoice" & vbCrLf & "ปิดหน้าต่างนี้ แล้วเปิด Invoice ขึ้นมาตรวจสอบความถูกต้องอีกครั้ง", vbInformation
            CompleteDetailWithError = False
'            GoTo RetrySaveDetail
        End If
    Next
    
'    If CompleteDetailWithError = True Then
'        MsgBox "เกิดข้อผิดพลาดขณะทำการบันทึกรายการใน Invoice ระบบจะทำการบันทึกซ้ำอีกครั้ง" & vbCrLf & "ก่อนกดปุ่ม OK ให้ตรวจสอบว่ามีการเปิดใบเมนิเฟสที่กำลังบันทึกอยู่หรือไม่", vbInformation
'        GoTo RetrySaveDetail
'    End If
    
    DBConnExc "Delete tbBillingDetail WHERE MarkForDelete='Y' AND BillingHeaderID = '" & id_ & "'"
    
    If Not colTripControl_ Is Nothing Then
    For x = 1 To colTripControl_.Count
        colTripControl_(x).UpdateCustBillNo
    Next
    End If
    
    isChanged_ = True
    '*** Update Status in Menifest IsINV = "Y"
    ErrState = UpdateManifest(id_, "Y", CurrentUser.ID, id_)
    If ErrState <> "" Then
        MsgBox ErrState, vbCritical, "Error"
    End If
    
    Set tmpRec = Nothing
    CompleteDetailWithError = False
    CloseDBConnection
    Exit Sub
ErrD:
    SaveErr = True
    isSave_ = False
'    DBConn.RollbackTrans
    frmMain.Stb1.Panels(1).Text = "เกิดข้อผิดพลาดขณะทำการบันทึก Inv Header, กรุณาปิดใบ Invoice และทำการบันทึกใหม่อีกครั้ง"
    frmMain.Stb1.Panels(2).Text = ""
    MsgBox "เกิดข้อผิดพลาดขณะทำการบันทึก Inv Header, กรุณาปิดใบ Invoice และทำการบันทึกใหม่อีกครั้ง", vbCritical, "Error"
    Exit Sub
End Sub

Private Function UpdateManifest(ByVal invID As String, ByVal Value_ As String, ByVal UserInv_ As String, ByVal SetInvID As String) As String
On Error GoTo ErrD
Dim tSql As String
Dim RS As ADODB.Recordset
    UpdateManifest = ""
    tSql = "SELECT     BillingHeaderID, SellID, RowID" & vbCrLf
    tSql = tSql & "From dbo.tbBillingDetail" & vbCrLf
    tSql = tSql & "WHERE     (BillingHeaderID = '" & invID & "')" & vbCrLf
    tSql = tSql & "ORDER BY RowID"
    Set RS = GetRS(tSql, adOpenForwardOnly, adLockOptimistic)
    Do While Not RS.EOF
        tSql = "UPDATE tbMenifestHeader SET isInv = '" & Value_ & "', UserInv = '" & UserInv_ & "', InvID = '" & SetInvID & "'"
        tSql = tSql & "WHERE MenifestID = '" & RS!SellID & "'"
        DBConnExc tSql
        RS.MoveNext
    Loop
'    RS.Close
Exit Function
ErrD:
    UpdateManifest = "เกิดข้อผิดพลาดขณะทำการอัเดทสถานะในใบเมนิเฟสเป็น (" & Value_ & ") การกระทำทั้งหมดถูกยกเลิก"
    Exit Function
End Function

Public Function UpdateReceive() As String
On Error GoTo ErrD
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
    sqlStr = "SELECT * , CONVERT(DATETIME, GETDATE(), 103) AS GETDATE FROM tbBillingHeader WHERE BillingHeaderID = '" & id_ & "'"
    Set tmpRec = New ADODB.Recordset
    GetDBConnection
    tmpRec.open sqlStr, DBConn, adOpenDynamic, adLockOptimistic
    If tmpRec.EOF = False Then
        If Trim(isCustAccepted_) = "" Then isCustAccepted_ = "N"
        If Trim(payDueDate_) = "" Then payDueDate_ = ""
        
        tmpRec!isCustAccepted = Trim(isCustAccepted_)
        tmpRec!PayDueDate = Trim(payDueDate_)
        tmpRec!LastUserRE = Trim(Trim(CurrentUser.StaffFName) & " " & Trim(CurrentUser.StaffLName))
        tmpRec!LastUpdateRE = tmpRec!GetDate
        tmpRec.Update
        UpdateReceive = "Update Completed"
    End If
    Set tmpRec = Nothing
    CloseDBConnection
    Exit Function
ErrD:
    UpdateReceive = Trim(err.Description)
    Exit Function
End Function

Public Function Delete() As String
On Error GoTo delErr:
    DBConnExc "Delete tbBillingHeader WHERE BillingHeaderID='" & id_ & "'"
    Delete = ""
    Exit Function
delErr:
    Delete = err.Description
End Function

Public Sub SetDataFromRecordset(selRec As ADODB.Recordset)
    id_ = Trim("" & selRec("BillingHeaderID"))
    IssuedDate_ = Trim("" & selRec("IssuedDate"))
    CustomerID_ = Trim("" & selRec("CustomerID"))
    CustomerName_ = Trim("" & selRec("CompanyName"))
    companyID_ = Trim("" & selRec("CompanyID"))
    companyName_ = Trim("" & selRec("CompanyNameTH"))
    invoiceNo_ = Trim("" & selRec("InvoiceNo"))
    receiptNo_ = Trim("" & selRec("ReceiptNo"))
    payDueDate_ = Trim("" & selRec("PayDueDate"))
    discountAmt_ = selRec("DiscountAmt")
    AmtB4VAT_ = selRec("AmtB4VAT")
    VATAmt_ = selRec("VATAmt")
    TotalBillAmt_ = selRec("TotalBillAmt")
    billRemark_ = Trim("" & selRec("BillRemark"))
    isCancel_ = Trim("" & selRec("isCancel"))
    cancelDate_ = Trim("" & selRec("CancelDate"))
    CancelStaff_ = Trim("" & selRec("CancelStaff"))
    CancelDetail_ = Trim("" & selRec("CancelDetail"))
    discountType_ = Trim("" & selRec("DiscountType"))
    wTaxRate_ = selRec("WTax")
    If IsNull(selRec("PaidAmt")) = False Then paidAmt_ = selRec("PaidAmt")
    isCustAccepted_ = Trim("" & selRec("isCustAccepted"))
    isServiceCustomer_ = Trim("" & selRec("IsServiceCustomer"))
    isFiled_ = Trim("" & selRec("isFiled"))
    billAddrID_ = Trim("" & selRec("BillAddrID"))
    billAddr_ = Trim("" & selRec("Addr1")) & " " & Trim("" & selRec("TumbolName")) & " " & Trim("" & selRec("AumphurName")) & " " & Trim("" & selRec("ProvinceName")) & " " & Trim("" & selRec("ZipCode"))
    billRecStaff_ = Trim("" & selRec("BillRecStaff"))
    IncVAT_ = IIf(selRec!IncVAT = "Y", 0, 1)
    LastUser_ = Trim("" & selRec!LastUser)
    DatePast_ = Trim("" & selRec!DatePast)
    LastUpdate_ = Trim("" & selRec!LastUpdate)
    LastUserRE_ = Trim("" & selRec!LastUserRE)
    LastUpdateRE_ = Trim("" & selRec!LastUpdateRE)
    IsInv_ = Trim("" & selRec!IsInv)
    invDate_ = Trim("" & selRec!InvDate)
    BillingNo_ = Trim("" & selRec("BillingNo"))
    PreInvNo_ = Trim("" & selRec!PreInvNo)
    PreDate_ = Trim("" & selRec!PreDate)
    PreNote_ = Trim("" & selRec!PreNote)
    PreStaff_ = Trim("" & selRec!PreStaff)
End Sub

Public Function CancelBill(ByVal CancelStaff As String, ByVal CancelDetail As String) As String
On Error GoTo ErrD
Dim tmpRec As New ADODB.Recordset
Dim sqlStr As String
    sqlStr = "SELECT * FROM vw_CheckBillInv WHERE isCancel='N' And BillingHeaderID='" & id_ & "'"
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec.EOF Then
        '*** Update Manifest No is INV = 'N' in detail
        'UpdateManifest id_, "N", CurrentUser.ID    '*** ทดสอบปิดการยกเลิก Inv  2554-02-01 12:13
        sqlStr = "UPDATE tbMenifestHeader SET isInv = 'N', InvID = '' WHERE MenifestID IN (SELECT SellID FROM tbBillingDetail WHERE BillingHeaderID='" & id_ & "')"
        DBConnExc sqlStr
        
        DBConnExc "Update tbBillingHeader Set isCancel = 'Y', CancelDate = '" & TodayDate & "', CancelStaff = '" & CancelStaff & "', CancelDetail = '" & CancelDetail & "' WHERE BillingHeaderID = '" & id_ & "'"
        DBConnExc "Update tbBillingDetail Set isCancel='Y' WHERE BillingHeaderID='" & id_ & "'"
        DBConnExc "Update tbJobDataTimeTable Set CustBillID='' WHERE CustBillID='" & id_ & "'"
        
        isCancel_ = "Y"
        cancelDate_ = TodayDate
        CancelBill = ""
    Else
        CancelBill = "ใบวางบิลนี้ถูกออกใบกำกับภาษีแล้ว ไม่สามารถยกเลิกได้"
    End If
    tmpRec.Close
    Set tmpRec = Nothing
    Exit Function
ErrD:
    GetMsg "", vbCritical
    CancelBill = err.Description
    Exit Function
End Function
