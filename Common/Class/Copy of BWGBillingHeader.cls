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
Dim issuedDate_ As String
Dim customerID_ As String
Dim customerName_ As String
Dim companyID_ As String
Dim companyName_ As String
Dim billingNo_ As String
Dim invoiceNo_ As String
Dim receiptNo_ As String
Dim payDueDate_ As String
Dim discountAmt_ As Double
Dim amtB4Vat_ As Double
Dim vatAmt_ As Double
Dim totalBillAmt_ As Double
Dim billRemark_ As String
Dim isCancel_ As String
Dim cancelDate_ As String
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


Public Property Get ID() As String
    ID = id_
End Property
Public Property Let ID(ByVal newVal As String)
    id_ = Trim(newVal)
End Property

Public Property Get IssuedDate() As String
    IssuedDate = issuedDate_
End Property
Public Property Let IssuedDate(ByVal newVal As String)
    If Trim(issuedDate_) <> Trim(newVal) Then
        issuedDate_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get CustomerID() As String
    CustomerID = customerID_
End Property
Public Property Let CustomerID(ByVal newVal As String)
    If Trim(customerID_) <> Trim(newVal) Then
        customerID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get CustomerName() As String
    CustomerName = customerName_
End Property
Public Property Let CustomerName(ByVal newVal As String)
    If Trim(customerName_) <> Trim(newVal) Then
        customerName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get companyID() As String
    companyID = companyID_
End Property
Public Property Let companyID(ByVal newVal As String)
    If Trim(companyID_) <> Trim(newVal) Then
        companyID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get CompanyName() As String
    CompanyName = companyName_
End Property
Public Property Let CompanyName(ByVal newVal As String)
    If Trim(companyName_) <> Trim(newVal) Then
        companyName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get BillingNo() As String
    BillingNo = billingNo_
End Property
Public Property Let BillingNo(ByVal newVal As String)
    If Trim(billingNo_) <> Trim(newVal) Then
        billingNo_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get InvoiceNo() As String
    InvoiceNo = invoiceNo_
End Property
Public Property Let InvoiceNo(ByVal newVal As String)
    If Trim(invoiceNo_) <> Trim(newVal) Then
        invoiceNo_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get ReceiptNo() As String
    ReceiptNo = receiptNo_
End Property
Public Property Let ReceiptNo(ByVal newVal As String)
    If Trim(receiptNo_) <> Trim(newVal) Then
        receiptNo_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get PayDueDate() As String
    PayDueDate = payDueDate_
End Property
Public Property Let PayDueDate(ByVal newVal As String)
    If Trim(payDueDate_) <> Trim(newVal) Then
        payDueDate_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get DiscountAmt() As Double
    DiscountAmt = discountAmt_
End Property
Public Property Let DiscountAmt(ByVal newVal As Double)
    If discountAmt_ <> newVal Then
        discountAmt_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get AmtB4VAT() As Double
    AmtB4VAT = amtB4Vat_
End Property
Public Property Let AmtB4VAT(ByVal newVal As Double)
    If amtB4Vat_ <> newVal Then
        amtB4Vat_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get VATAmt() As Double
    VATAmt = vatAmt_
End Property
Public Property Let VATAmt(ByVal newVal As Double)
    If vatAmt_ <> newVal Then
        vatAmt_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get TotalBillAmt() As Double
    TotalBillAmt = totalBillAmt_
End Property
Public Property Let TotalBillAmt(ByVal newVal As Double)
    If totalBillAmt_ <> newVal Then
        totalBillAmt_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get BillRemark() As String
    BillRemark = billRemark_
End Property
Public Property Let BillRemark(ByVal newVal As String)
    If Trim(billRemark_) <> Trim(newVal) Then
        billRemark_ = Trim(newVal)
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
Public Property Let isCancel(ByVal newVal As Boolean)
Dim tmpStr As String
    If newVal Then
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
Public Property Let CancelDate(ByVal newVal As String)
    If Trim(cancelDate_) <> Trim(newVal) Then
        cancelDate_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get DiscountType() As String
    DiscountType = discountType_
End Property
Public Property Let DiscountType(ByVal newVal As String)
    If Trim(discountType_) <> Trim(newVal) Then
        discountType_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get WTaxRate() As Double
    WTaxRate = wTaxRate_
End Property
Public Property Let WTaxRate(ByVal newVal As Double)
    If wTaxRate_ <> newVal Then
        wTaxRate_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get PaidAmt() As Double
    PaidAmt = paidAmt_
End Property
Public Property Let PaidAmt(ByVal newVal As Double)
    If paidAmt_ <> newVal Then
        paidAmt_ = newVal
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
Public Property Let isCustAccepted(ByVal newVal As Boolean)
Dim tmpStr As String
    If newVal Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(isCustAccepted_) <> Trim(tmpStr) Then
        isCustAccepted_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property

Public Property Get BillAddrID() As String
    BillAddrID = billAddrID_
End Property
Public Property Let BillAddrID(ByVal newVal As String)
    If Trim(billAddrID_) <> Trim(newVal) Then
        billAddrID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get BillAddr() As String
    BillAddr = billAddr_
End Property
Public Property Let BillAddr(ByVal newVal As String)
    If Trim(billAddr_) <> Trim(newVal) Then
        billAddr_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get BillRecStaff() As String
    BillRecStaff = billRecStaff_
End Property
Public Property Let BillRecStaff(ByVal newVal As String)
    If Trim(billRecStaff_) <> Trim(newVal) Then
        billRecStaff_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get BillingDetails() As Collection
Dim tmpS As New BWGSearchManager
    If colBillDetail_ Is Nothing Then Set colBillDetail_ = tmpS.BillingDetailSearch("BillingHeaderID='" & id_ & "'")
    Set tmpS = Nothing
    Set BillingDetails = colBillDetail_
End Property
Public Property Set BillingDetails(ByVal newVal As Collection)
    Set colBillDetail_ = newVal
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
Public Property Set TripControls(ByVal newVal As Collection)
    Set colTripControl_ = newVal
    isChanged_ = True
End Property
Public Sub RequeryTripControl()
    Set colTripControl_ = Nothing
End Sub

Public Property Get isChanged() As Boolean
    isChanged = isChanged_
End Property
Public Property Let isChanged(ByVal newVal As Boolean)
    isChanged_ = newVal
End Property

Public Property Get isSave() As Boolean
    isSave = isSave_
End Property
Public Property Let isSave(ByVal newVal As Boolean)
    isSave_ = newVal
End Property

Public Property Get IncVAT() As Integer
    IncVAT = IncVAT_
End Property
Public Property Let IncVAT(ByVal newVal As Integer)
    IncVAT_ = newVal
    isChanged_ = True
End Property

Public Sub Save()
On Error GoTo ErrD
Dim tmpRec As ADODB.Recordset
Dim x%
Dim SaveErr As Boolean
Dim ErrState As String, sqlStr As String
    SaveErr = False
    isSave_ = True
    DBManager.GetDBConnection.BeginTrans
    Set tmpRec = New ADODB.Recordset
    sqlStr = "Select * from tbBillingHeader Where BillingHeaderID='" & id_ & "'"
'    Set tmpRec = GetRS(sqlStr, adOpenKeyset, adLockOptimistic)
    tmpRec.Open sqlStr, DBManager.GetDBConnection, adOpenDynamic, adLockOptimistic

    If tmpRec.EOF Then
        tmpRec.AddNew
        tmpRec("BillingHeaderID") = id_
    Else
        If UpdateManifest(id_, "N", CurrentUser.ID) = False Then Exit Sub
    End If
    tmpRec("IssuedDate") = Trim(issuedDate_)
    tmpRec("CustomerID") = Trim(customerID_)
    tmpRec("CompanyID") = Trim(companyID_)
    'tmpRec("BillingNo") = Trim(billingNo_)
    tmpRec("InvoiceNo") = Trim(invoiceNo_)
    tmpRec("ReceiptNo") = Trim(receiptNo_)
    tmpRec("PayDueDate") = Trim(payDueDate_)
    tmpRec("DiscountAmt") = Trim(discountAmt_)
    tmpRec("AmtB4VAT") = Trim(amtB4Vat_)
    tmpRec("VATAmt") = Trim(vatAmt_)
    tmpRec("TotalBillAmt") = Trim(totalBillAmt_)
    tmpRec("BillRemark") = Trim(billRemark_)
    If Trim(isCancel_) = "" Then isCancel_ = "N"
    tmpRec("isCancel") = Trim(isCancel_)
    'tmpRec("CancelDate") = Trim(cancelDate_)
    tmpRec("DiscountType") = Trim(discountType_)
    tmpRec("WTax") = Trim(wTaxRate_)
    tmpRec("isCustAccepted") = Trim(isCustAccepted_)
    tmpRec("BillAddrID") = Trim(billAddrID_)
    tmpRec("BillRecStaff") = Trim(billRecStaff_)
    tmpRec!IncVAT = IIf(IncVAT_ = 0, "Y", "N")
    tmpRec.Update
    tmpRec.Requery
    billingNo_ = Trim("" & tmpRec("BillingNo"))
    tmpRec.Close
    
    DBManager.GetDBConnection.CommitTrans
    ErrState = "กำลังบันทึกเลขที่ Invoice"
    Set tmpRec = Nothing
    If SaveErr = True Then Exit Sub
    DBManager.GetDBConnection.BeginTrans
    With DBManager.GetDBConnection
        If UpdateManifest(id_, "N", CurrentUser.ID) = False Then
            Exit Sub
        End If
        .Execute "Update tbBillingDetail Set MarkForDelete='Y' Where BillingHeaderID='" & id_ & "'"
        For x = 1 To colBillDetail_.Count
                
            colBillDetail_(x).Save
            ErrState = "กำลังบันทึกรายละเอียดใน Invoice ลำดับที่ " & x
        Next
        
        .Execute "Delete tbBillingDetail Where MarkForDelete='Y'"
        
        If Not colTripControl_ Is Nothing Then
        For x = 1 To colTripControl_.Count
            colTripControl_(x).UpdateCustBillNo
        Next
        End If
    End With
    isChanged_ = True
    
    DBManager.GetDBConnection.CommitTrans
    
    If UpdateManifest(id_, "Y", CurrentUser.ID) = False Then
        Exit Sub
    End If
    
    Exit Sub
    
ErrD:
    SaveErr = True
    isSave_ = False
    DBManager.GetDBConnection.RollbackTrans
    MsgBox Err.Description & vbCrLf & "เกิดข้อผิดพลาด การกระทำทั้งหมดถูกยกเลิก, ถ้าไม่แน่ใจให้ติดต่อเจ้าหน้าที่ IT", vbCritical, "Error"
    Exit Sub
End Sub

Private Function UpdateManifest(ByVal invID As String, ByVal Value_ As String, ByVal UserInv_ As String) As Boolean
On Error GoTo ErrD
Dim tSql As String
Dim RS As ADODB.Recordset
    UpdateManifest = True
    tSql = "SELECT     TOP (100) PERCENT BillingHeaderID, SellID, RowID" & vbCrLf
    tSql = tSql & "From dbo.tbBillingDetail" & vbCrLf
    tSql = tSql & "WHERE     (BillingHeaderID = '" & invID & "')" & vbCrLf
    tSql = tSql & "ORDER BY RowID"
    Set RS = GetRS(tSql, adOpenForwardOnly, adLockOptimistic)
    Do While Not RS.EOF
        tSql = "UPDATE tbMenifestHeader SET isInv = '" & Value_ & "', UserInv = '" & UserInv_ & "'"
        tSql = tSql & "WHERE MenifestID = '" & RS!SellID & "'"
        DBManager.GetDBConnection.Execute tSql
        RS.MoveNext
    Loop
Exit Function
ErrD:
    UpdateManifest = False
    MsgBox Err.Description & vbCrLf & "เกิดข้อผิดพลาดขณะทำการอัเดทสถานะในใบเมนิเฟส (" & Value_ & ") การกระทำทั้งหมดถูกยกเลิก", vbCritical, "Error"
End Function

Public Function Delete() As String
On Error GoTo delErr:
    DBManager.GetDBConnection.Execute "Delete tbBillingHeader Where BillingHeaderID='" & id_ & "'"
    Delete = ""
    Exit Function
delErr:
    Delete = Err.Description
End Function

Public Sub SetDataFromRecordset(selRec As ADODB.Recordset)
    id_ = Trim("" & selRec("BillingHeaderID"))
    issuedDate_ = Trim("" & selRec("IssuedDate"))
    customerID_ = Trim("" & selRec("CustomerID"))
    customerName_ = Trim("" & selRec("CompanyName"))
    companyID_ = Trim("" & selRec("CompanyID"))
    companyName_ = Trim("" & selRec("CompanyNameTH"))
    billingNo_ = Trim("" & selRec("BillingNo"))
    invoiceNo_ = Trim("" & selRec("InvoiceNo"))
    receiptNo_ = Trim("" & selRec("ReceiptNo"))
    payDueDate_ = Trim("" & selRec("PayDueDate"))
    discountAmt_ = selRec("DiscountAmt")
    amtB4Vat_ = selRec("AmtB4VAT")
    vatAmt_ = selRec("VATAmt")
    totalBillAmt_ = selRec("TotalBillAmt")
    billRemark_ = Trim("" & selRec("BillRemark"))
    isCancel_ = Trim("" & selRec("isCancel"))
    cancelDate_ = Trim("" & selRec("CancelDate"))
    discountType_ = Trim("" & selRec("DiscountType"))
    wTaxRate_ = selRec("WTax")
    If IsNull(selRec("PaidAmt")) = False Then paidAmt_ = selRec("PaidAmt")
    isCustAccepted_ = Trim("" & selRec("isCustAccepted"))
    billAddrID_ = Trim("" & selRec("BillAddrID"))
    billAddr_ = Trim("" & selRec("Addr1")) & " " & Trim("" & selRec("TumbolName")) & " " & Trim("" & selRec("AumphurName")) & " " & Trim("" & selRec("ProvinceName")) & " " & Trim("" & selRec("ZipCode"))
    billRecStaff_ = Trim("" & selRec("BillRecStaff"))
    IncVAT_ = IIf(selRec!IncVAT = "Y", 0, 1)
End Sub

Public Function CancelBill() As String
On Error GoTo ErrD
Dim tmpRec As New ADODB.Recordset
Dim sqlStr As String
    sqlStr = "Select * from vw_CheckBillInv Where isCancel='N' And BillingHeaderID='" & id_ & "'"
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
'    tmpRec.Open sqlStr, DBManager.GetDBConnection, adOpenKeyset, adLockOptimistic
    If tmpRec.EOF Then
        With DBManager.GetDBConnection
            .Execute "Update tbBillingHeader Set isCancel='Y',CancelDate='" & TodayDate & "' Where BillingHeaderID='" & id_ & "'"
            .Execute "Update tbBillingDetail Set isCancel='Y' Where BillingHeaderID='" & id_ & "'"
            .Execute "Update tbJobDataTimeTable Set CustBillID='' Where CustBillID='" & id_ & "'"
        End With
        
        If UpdateManifest(id_, "N", CurrentUser.ID) = False Then
            Exit Function
        End If
        
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
    MsgBox Err.Description & vbCrLf & "กรุณาลองใหม่หรือติดต่อผู้พัฒนาระบบ", vbCritical, "Error"
    Exit Function
End Function
