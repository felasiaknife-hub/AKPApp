VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BWGBillPayHeader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim id_ As String
Dim billPayDate_ As String
Dim supplierID_ As String
Dim supplierName_ As String
Dim companyID_ As String
Dim companyName_ As String
Dim billPayNo_ As String
Dim BillingNo_ As String
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
Dim wTax_ As Long
Dim vatRate_ As Double

Dim sasomAmt_ As Double
Dim wTaxAmt_ As Double
Dim deductB4VatAmt_ As Double
Dim fundAmt_ As Double
Dim deductAfterVatAmt_ As Double

Dim colPayDetail_ As Collection
Dim isChanged_ As Boolean

Public Property Get ID() As String
    ID = id_
End Property
Public Property Let ID(ByVal newVal As String)
    id_ = Trim(newVal)
End Property

Public Property Get BillPayDate() As String
    BillPayDate = billPayDate_
End Property
Public Property Let BillPayDate(ByVal newVal As String)
    If Trim(billPayDate_) <> Trim(newVal) Then
        billPayDate_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get SupplierID() As String
    SupplierID = supplierID_
End Property
Public Property Let SupplierID(ByVal newVal As String)
    If Trim(supplierID_) <> Trim(newVal) Then
        supplierID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get SupplierName() As String
    SupplierName = supplierName_
End Property
Public Property Let SupplierName(ByVal newVal As String)
    If Trim(supplierName_) <> Trim(newVal) Then
        supplierName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get CompanyID() As String
    CompanyID = companyID_
End Property
Public Property Let CompanyID(ByVal newVal As String)
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

Public Property Get BillPayNo() As String
    BillPayNo = billPayNo_
End Property
Public Property Let BillPayNo(ByVal newVal As String)
    If Trim(billPayNo_) <> Trim(newVal) Then
        billPayNo_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get BillingNo() As String
    BillingNo = BillingNo_
End Property
Public Property Let BillingNo(ByVal newVal As String)
    If Trim(BillingNo_) <> Trim(newVal) Then
        BillingNo_ = Trim(newVal)
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
    BillRemark = Trim(billRemark_)
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
    CancelDate = Trim(cancelDate_)
End Property
Public Property Let CancelDate(ByVal newVal As String)
    If Trim(cancelDate_) <> Trim(newVal) Then
        cancelDate_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get DiscountType() As String
    DiscountType = Trim(discountType_)
End Property
Public Property Let DiscountType(ByVal newVal As String)
    If Trim(discountType_) <> Trim(newVal) Then
        discountType_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get WTax() As Long
    WTax = wTax_
End Property
Public Property Let WTax(ByVal newVal As Long)
    If wTax_ <> newVal Then
        wTax_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get VATRate() As Double
    VATRate = vatRate_
End Property
Public Property Let VATRate(ByVal newVal As Double)
    If vatRate_ <> newVal Then
        vatRate_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get SasomAmt() As Double
    SasomAmt = sasomAmt_
End Property
Public Property Let SasomAmt(ByVal newVal As Double)
    If sasomAmt_ <> newVal Then
        sasomAmt_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get WTaxAmt() As Double
    WTaxAmt = wTaxAmt_
End Property
Public Property Let WTaxAmt(ByVal newVal As Double)
    If wTaxAmt_ <> newVal Then
        wTaxAmt_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get DeductB4VATAmt() As Double
    DeductB4VATAmt = deductB4VatAmt_
End Property
Public Property Let DeductB4VATAmt(ByVal newVal As Double)
    If deductB4VatAmt_ <> newVal Then
        deductB4VatAmt_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get FundAmt() As Double
    FundAmt = fundAmt_
End Property
Public Property Let FundAmt(ByVal newVal As Double)
    If fundAmt_ <> newVal Then
        fundAmt_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get DeductAfterVATAmt() As Double
    DeductAfterVATAmt = deductAfterVatAmt_
End Property
Public Property Let DeductAfterVATAmt(ByVal newVal As Double)
    If deductAfterVatAmt_ <> newVal Then
        deductAfterVatAmt_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get BillPayDetails() As Collection
Dim tmpS As New BWGSearchManager
    If colPayDetail_ Is Nothing Then Set colPayDetail_ = tmpS.BillPayDetailSearch("BillPayHeaderID='" & id_ & "'")
    Set BillPayDetails = colPayDetail_
    Set tmpS = Nothing
End Property
Public Property Set BillPayDetails(ByVal newVal As Collection)
    Set colPayDetail_ = newVal
    isChanged_ = True
End Property
Public Sub RequeryBillPayDetail()
    Set colPayDetail_ = Nothing
End Sub

Public Property Get isChanged() As Boolean
    isChanged = isChanged_
End Property
Public Property Let isChanged(ByVal newVal As Boolean)
    isChanged_ = newVal
End Property

Public Sub Save()
On Error GoTo ErrD
Dim tmpRec As ADODB.Recordset
Dim x%, sqlStr As String
    sqlStr = "SELECT *, CONVERT(DATETIME, GETDATE(), 103) AS GETDATE FROM tbBillPayHeader WHERE BillPayHeaderID='" & id_ & "'"
    Set tmpRec = New ADODB.Recordset
    GetDBConnection
    tmpRec.open sqlStr, DBConn, adOpenDynamic, adLockOptimistic
    If tmpRec.EOF Then
        tmpRec.AddNew
        tmpRec("BillPayHeaderID") = id_
    End If
    tmpRec("BillPayDate") = billPayDate_
    tmpRec("SupplierID") = supplierID_
    tmpRec("CompanyID") = companyID_
    tmpRec("BillingNo") = BillingNo_
    tmpRec("InvoiceNo") = invoiceNo_
    tmpRec("ReceiptNo") = receiptNo_
    tmpRec("PayDueDate") = payDueDate_
    tmpRec("DiscountAmt") = discountAmt_
    tmpRec("AmtB4VAT") = amtB4Vat_
    tmpRec("VATAmt") = vatAmt_
    tmpRec("TotalBillAmt") = totalBillAmt_
    tmpRec("BillRemark") = billRemark_
    If Trim(isCancel_) = "" Then isCancel_ = "N"
    tmpRec("isCancel") = isCancel_
    tmpRec("CancelDate") = cancelDate_
    tmpRec("DiscountType") = discountType_
    tmpRec("WTaxRate") = wTax_
    tmpRec("VATRate") = vatRate_
    
    tmpRec("SasomAmt") = sasomAmt_
    tmpRec("WTaxAmt") = wTaxAmt_
    tmpRec("DeductB4Vat") = deductB4VatAmt_
    tmpRec("FundAmt") = fundAmt_
    tmpRec("DeductAfterVAT") = deductAfterVatAmt_
    
    tmpRec!LastUser = Trim(Trim(CurrentUser.StaffFName) & " " & Trim(CurrentUser.StaffLName))
    tmpRec!LastUpdate = GetSystemDateTime.SystemDateTime
    
    tmpRec.Update
    tmpRec.Requery
    billPayNo_ = Trim("" & tmpRec("BillPayNo"))
'    tmpRec.Close
    Set tmpRec = Nothing
    
    DBConnExc "Update tbBillPayDetail Set MarkForDelete='Y' WHERE BillPayHeaderID='" & id_ & "'"
    For x = 1 To colPayDetail_.Count
        colPayDetail_(x).Save
    Next
    DBConnExc "Delete tbBillPayDetail WHERE MarkForDelete='Y'"
    isChanged_ = False
    CloseDBConnection
    Exit Sub
ErrD:
    GetMsg "", vbCritical
    Exit Sub
End Sub

Public Function Delete() As String
On Error GoTo delErr:
    DBConnExc "Delete tbBillPayHeader WHERE BillPayHeaderID='" & id_ & "'"
    Delete = ""
    Exit Function
delErr:
    Delete = err.Description
End Function

Public Sub SetDataFromRecordset(selRec As ADODB.Recordset)
    id_ = Trim("" & selRec("BillPayHeaderID"))
    billPayDate_ = Trim("" & selRec("BillPayDate"))
    supplierID_ = Trim("" & selRec("SupplierID"))
    supplierName_ = Trim("" & selRec("VendorName"))
    companyID_ = Trim("" & selRec("CompanyID"))
    companyName_ = Trim("" & selRec("CompanyNameTH"))
    billPayNo_ = Trim("" & selRec("BillPayNo"))
    BillingNo_ = Trim("" & selRec("BillingNo"))
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
    wTax_ = selRec("WTaxRate")
    vatRate_ = selRec("VATRate")
    
    sasomAmt_ = selRec("SasomAmt")
    wTaxAmt_ = selRec("WTaxAmt")
    deductB4VatAmt_ = selRec("DeductB4Vat")
    fundAmt_ = selRec("FundAmt")
    deductAfterVatAmt_ = selRec("DeductAfterVAT")
    
End Sub

Public Sub CancelBill()
    DBConnExc "Update tbBillPayHeader Set isCancel='Y',CancelDate='" & TodayDate & "' WHERE BillPayHeaderID='" & id_ & "'"
    DBConnExc "Update tbBillPayDetail Set isCancel='Y' WHERE BillPayHeaderID='" & id_ & "'"
    isCancel_ = "Y"
    cancelDate_ = TodayDate
End Sub
