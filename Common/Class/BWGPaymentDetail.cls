VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BWGPaymentDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim id_ As String
Dim PaymentHeaderID_ As String
Dim rowID_ As Integer
Dim billDesc_ As String
Dim amtB4Vat_ As Double
Dim discountAmt_ As Double
Dim vatAmt_ As Double
Dim totalAmtToPay_ As Double
Dim isCancel_ As String
Dim expenseID_ As String
Dim expenseVendorID_ As String
Dim expenseVendorName_ As String
Dim expenseBillDesc_ As String
Dim expenseDate_ As String
Dim expenseBuyQty_ As Double
Dim expenseNetPrice_ As Double
Dim expenseisPaid_ As String
Dim beforeOrAfterVAT_ As String
Dim venderDeductRate_ As Double
Dim venderFundAmt_ As Double

Dim JobDataCarID_ As String
Dim TripNo_ As Integer
Dim TimeTableNo_ As String
Dim fromAumphurID_ As String
Dim fromAumphurName_ As String
Dim fromProvinceID_ As String
Dim fromProvinceName_ As String
Dim toAumphurID_ As String
Dim toAumphurName_ As String
Dim toProvinceID_ As String
Dim toProvinceName_ As String
Dim workDate_ As String

Dim truckTypeID_ As String
Dim truckTypeDesc_ As String
Dim truckSubTypeID_ As String
Dim truckSubTypeDesc_ As String
Dim carRegisNo_ As String
Dim TrailerTruckNo_ As String
Dim DriverName_ As String
Dim CustomerName_ As String
Dim CompanyCode_ As String
Dim isChanged_ As Boolean

Public Property Get ID() As String
    ID = id_
End Property
Public Property Let ID(ByVal newVal As String)
    id_ = Trim(newVal)
End Property

Public Property Get PaymentHeaderID() As String
    PaymentHeaderID = Trim(PaymentHeaderID_)
End Property
Public Property Let PaymentHeaderID(ByVal newVal As String)
    If Trim(PaymentHeaderID_) <> Trim(newVal) Then
        PaymentHeaderID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get RowID() As Integer
    RowID = rowID_
End Property
Public Property Let RowID(ByVal newVal As Integer)
    If rowID_ <> newVal Then
        rowID_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get BillDesc() As String
    BillDesc = Trim(billDesc_)
End Property
Public Property Let BillDesc(ByVal newVal As String)
    If Trim(billDesc_) <> Trim(newVal) Then
        billDesc_ = Trim(newVal)
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

Public Property Get DiscountAmt() As Double
    DiscountAmt = discountAmt_
End Property
Public Property Let DiscountAmt(ByVal newVal As Double)
    If discountAmt_ <> newVal Then
        discountAmt_ = newVal
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

Public Property Get TotalAmtToPay() As Double
    TotalAmtToPay = totalAmtToPay_
End Property
Public Property Let TotalAmtToPay(ByVal newVal As Double)
    If totalAmtToPay_ <> newVal Then
        totalAmtToPay_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get isCancel() As Boolean
    If isCancel_ = "Y" Then
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

Public Property Get ExpenseID() As String
    ExpenseID = Trim(expenseID_)
End Property
Public Property Let ExpenseID(ByVal newVal As String)
    If Trim(expenseID_) <> Trim(newVal) Then
        expenseID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get ExpenseVendorID() As String
    ExpenseVendorID = Trim(expenseVendorID_)
End Property
Public Property Let ExpenseVendorID(ByVal newVal As String)
    If Trim(expenseVendorID_) <> Trim(newVal) Then
        expenseVendorID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get ExpenseVendorName() As String
    ExpenseVendorName = Trim(expenseVendorName_)
End Property
Public Property Let ExpenseVendorName(ByVal newVal As String)
    If Trim(expenseVendorName_) <> Trim(newVal) Then
        expenseVendorName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get ExpenseBillDesc() As String
    ExpenseBillDesc = Trim(expenseBillDesc_)
End Property
Public Property Let ExpenseBillDesc(ByVal newVal As String)
    If Trim(expenseBillDesc_) <> Trim(newVal) Then
        expenseBillDesc_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get ExpenseDate() As String
    ExpenseDate = Trim(expenseDate_)
End Property
Public Property Let ExpenseDate(ByVal newVal As String)
    If Trim(expenseDate_) <> Trim(newVal) Then
        expenseDate_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get ExpenseBuyQty() As Double
    ExpenseBuyQty = expenseBuyQty_
End Property
Public Property Let ExpenseBuyQty(ByVal newVal As Double)
    If expenseBuyQty_ <> newVal Then
        expenseBuyQty_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get ExpenseNetPrice() As Double
    ExpenseNetPrice = expenseNetPrice_
End Property
Public Property Let ExpenseNetPrice(ByVal newVal As Double)
    If expenseNetPrice_ <> newVal Then
        expenseNetPrice_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get ExpenseIsPaid() As Boolean
    If Trim(expenseisPaid_) = "Y" Then
        ExpenseIsPaid = True
    Else
        ExpenseIsPaid = False
    End If
End Property
Public Property Let ExpenseIsPaid(ByVal newVal As Boolean)
Dim tmpStr As String
    If newVal Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(expenseisPaid_) <> Trim(tmpStr) Then
        expenseisPaid_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property

Public Property Get BeforeOrAfterVAT() As String
    BeforeOrAfterVAT = Trim(beforeOrAfterVAT_)
End Property
Public Property Let BeforeOrAfterVAT(ByVal newVal As String)
    If Trim(beforeOrAfterVAT_) <> Trim(newVal) Then
        beforeOrAfterVAT_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get VenderDeductRate() As Double
    VenderDeductRate = venderDeductRate_
End Property
Public Property Let VenderDeductRate(ByVal newVal As Double)
    If venderDeductRate_ <> newVal Then
        venderDeductRate_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get VenderFundAmt() As Double
    VenderFundAmt = venderFundAmt_
End Property
Public Property Let VenderFundAmt(ByVal newVal As Double)
    If venderFundAmt_ <> newVal Then
        venderFundAmt_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get JobDataCarID() As String
    JobDataCarID = Trim(JobDataCarID_)
End Property
Public Property Let JobDataCarID(ByVal newVal As String)
    If Trim(JobDataCarID_) <> Trim(newVal) Then
        JobDataCarID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get TripNo() As String
    TripNo = TripNo_
End Property
Public Property Let TripNo(ByVal newVal As String)
    If TripNo_ <> newVal Then
       TripNo_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get TimeTableNo() As String
    TimeTableNo = Trim(TimeTableNo_)
End Property
Public Property Let TimeTableNo(ByVal newVal As String)
    If Trim(TimeTableNo_) <> Trim(newVal) Then
        TimeTableNo_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get FromAumphurID() As String
    FromAumphurID = Trim(fromAumphurID_)
End Property
Public Property Let FromAumphurID(ByVal newVal As String)
    If Trim(fromAumphurID_) <> Trim(newVal) Then
        fromAumphurID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get FromAumphurName() As String
    FromAumphurName = Trim(fromAumphurName_)
End Property
Public Property Let FromAumphurName(ByVal newVal As String)
    If Trim(fromAumphurName_) <> Trim(newVal) Then
        fromAumphurName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get FromProvinceID() As String
    FromProvinceID = Trim(fromProvinceID_)
End Property
Public Property Let FromProvinceID(ByVal newVal As String)
    If Trim(fromProvinceID_) <> Trim(newVal) Then
        fromProvinceID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get FromProvinceName() As String
    FromProvinceName = Trim(fromProvinceName_)
End Property
Public Property Let FromProvinceName(ByVal newVal As String)
    If Trim(fromProvinceName_) <> Trim(newVal) Then
        fromProvinceName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get ToAumphurID() As String
    ToAumphurID = Trim(toAumphurID_)
End Property
Public Property Let ToAumphurID(ByVal newVal As String)
    If Trim(toAumphurID_) <> Trim(newVal) Then
        toAumphurID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get ToAumphurName() As String
    ToAumphurName = Trim(toAumphurName_)
End Property
Public Property Let ToAumphurName(ByVal newVal As String)
    If Trim(toAumphurName_) <> Trim(newVal) Then
        toAumphurName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get ToProvinceID() As String
    ToProvinceID = Trim(toProvinceID_)
End Property
Public Property Let ToProvinceID(ByVal newVal As String)
    If Trim(toProvinceID_) <> Trim(newVal) Then
        toProvinceID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get ToProvinceName() As String
    ToProvinceName = Trim(toProvinceName_)
End Property
Public Property Let ToProvinceName(ByVal newVal As String)
    If Trim(toProvinceName_) <> Trim(newVal) Then
        toProvinceName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get WorkDate() As String
    WorkDate = Trim(workDate_)
End Property
Public Property Let WorkDate(ByVal newVal As String)
    If Trim(workDate_) <> Trim(newVal) Then
        workDate_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get TruckTypeID() As String
    TruckTypeID = Trim(truckTypeID_)
End Property
Public Property Let TruckTypeID(ByVal newVal As String)
    If Trim(truckTypeID_) <> Trim(newVal) Then
        truckTypeID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get TruckTypeDesc() As String
    TruckTypeDesc = Trim(truckTypeDesc_)
End Property
Public Property Let TruckTypeDesc(ByVal newVal As String)
    If Trim(truckTypeDesc_) <> Trim(newVal) Then
        truckTypeDesc_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get TruckSubTypeID() As String
    TruckSubTypeID = Trim(truckSubTypeID_)
End Property
Public Property Let TruckSubTypeID(ByVal newVal As String)
    If Trim(truckSubTypeID_) <> Trim(newVal) Then
        truckSubTypeID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get TruckSubTypeDesc() As String
    TruckSubTypeDesc = Trim(truckSubTypeDesc_)
End Property
Public Property Let TruckSubTypeDesc(ByVal newVal As String)
    If Trim(truckSubTypeDesc_) <> Trim(newVal) Then
        truckSubTypeDesc_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get CarRegisNo() As String
    CarRegisNo = Trim(carRegisNo_)
End Property
Public Property Let CarRegisNo(ByVal newVal As String)
    If Trim(carRegisNo_) <> Trim(newVal) Then
        carRegisNo_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get TrailerTruckNo() As String
    TrailerTruckNo = Trim(TrailerTruckNo_)
End Property
Public Property Let TrailerTruckNo(ByVal newVal As String)
    If Trim(TrailerTruckNo_) <> Trim(newVal) Then
        TrailerTruckNo_ = Trim(newVal)
    End If
End Property

Public Property Get DriverName() As String
    DriverName = Trim(DriverName_)
End Property
Public Property Let DriverName(ByVal newVal As String)
    If Trim(DriverName_) <> Trim(newVal) Then
        DriverName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get CustomerName() As String
    CustomerName = Trim(CustomerName_)
End Property
Public Property Let CustomerName(ByVal newVal As String)
    If Trim(CustomerName_) <> Trim(newVal) Then
        CustomerName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get CompanyCode() As String
    CompanyCode = Trim(CompanyCode_)
End Property
Public Property Let CompanyCode(ByVal newVal As String)
    If Trim(CompanyCode_) <> Trim(newVal) Then
        CompanyCode_ = Trim(newVal)
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
On Error GoTo ErrD
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
    sqlStr = "SELECT * FROM tbPaymentDetail WHERE PaymentDetID = '" & id_ & "'"
    Set tmpRec = New ADODB.Recordset
    GetDBConnection
    tmpRec.open sqlStr, DBConn, adOpenDynamic, adLockOptimistic
    If tmpRec.EOF Then
        tmpRec.AddNew
        tmpRec("PaymentDetID") = id_
    End If
    tmpRec("PaymentID") = PaymentHeaderID_
    tmpRec("RowID") = rowID_
    tmpRec("BillDesc") = billDesc_
    tmpRec("AmtB4VAT") = amtB4Vat_
    tmpRec("DiscountAmt") = discountAmt_
    tmpRec("VATAmt") = vatAmt_
    tmpRec("TotalAmtToPay") = totalAmtToPay_
    If Trim(isCancel_) = "" Then isCancel_ = "N"
    tmpRec("isCancel") = isCancel_
    tmpRec("ExpenseID") = expenseID_
    tmpRec("BeforeOrAfterVAT") = beforeOrAfterVAT_
    tmpRec("DeductRate") = venderDeductRate_
    tmpRec("FundAmt") = venderFundAmt_
    tmpRec("MarkForDelete") = "N"
    tmpRec("TimeTableNo") = TimeTableNo_
    tmpRec("TripNo") = TripNo_
    tmpRec!CompanyCodeTrip = CompanyCode_
    tmpRec.Update
'     DBConnExc "Update tbJobdata WHERE PaymentDetID = '" & id_ & "'"
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
    DBConnExc "Delete tbPaymentDetail WHERE PaymentDetID = '" & id_ & "'"
    Delete = ""
    Exit Function
delErr:
    Delete = err.Description
End Function

Public Sub SetDataFromRecordset(selRec As ADODB.Recordset)
    id_ = Trim("" & selRec("PaymentDetID"))
    PaymentHeaderID_ = Trim("" & selRec("PaymentID"))
    rowID_ = selRec("RowID")
    billDesc_ = Trim("" & selRec("BillDesc"))
    amtB4Vat_ = selRec("AmtB4VAT")
    discountAmt_ = selRec("DiscountAmt")
    vatAmt_ = selRec("VATAmt")
    totalAmtToPay_ = selRec("TotalAmtToPay")
    isCancel_ = Trim("" & selRec("isCancel"))
    expenseID_ = Trim("" & selRec("ExpenseID"))
    expenseVendorID_ = Trim("" & selRec("VendorID"))
    expenseVendorName_ = Trim("" & selRec("VendorName"))
    expenseBillDesc_ = Trim("" & selRec("ExpenseBillDesc"))
    expenseDate_ = Trim("" & selRec("ExpenseDate"))
    If IsNull(selRec("BuyQty")) = False Then expenseBuyQty_ = selRec("BuyQty")
    If IsNull(selRec("NetPrice")) = False Then expenseNetPrice_ = selRec("NetPrice")
    expenseisPaid_ = Trim("" & selRec("isPaid"))
    beforeOrAfterVAT_ = Trim("" & selRec("BeforeOrAfterVAT"))
    If IsNull(selRec("DeductRate")) = False Then venderDeductRate_ = selRec("DeductRate")
    If IsNull(selRec("FundAmt")) = False Then venderFundAmt_ = selRec("FundAmt")
    
    JobDataCarID_ = Trim("" & selRec("JobDataCarid"))
    TripNo_ = Trim("" & selRec("TripNo"))
    TimeTableNo_ = Trim("" & selRec("TimeTableNo"))
    
    If isExist(selRec, "CompanyCodeTrip") Then CompanyCode_ = Trim("" & selRec!CompanyCodeTrip)
End Sub

