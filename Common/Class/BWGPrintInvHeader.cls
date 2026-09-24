VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BWGPrintInvHeader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim id_ As String
Dim invCompanyID_ As String
Dim invNo_ As String
Dim invDate_ As String
Dim isCancel_ As String
Dim cancelDate_ As String
Dim isPrintAfterChanged_ As String
Dim CustomerID_ As String
Dim CustomerName_ As String
Dim billAddrID_ As String
Dim billAddr_ As String
Dim billTumbol_ As String
Dim billAumphur_ As String
Dim billProvince_ As String
Dim billZipCode_ As String
Dim creditDay_ As Integer
Dim dueDate_ As String
Dim invRemark_ As String
Dim paidCashAmt_ As Double
Dim colInvDetail_ As Collection
Dim colInvDetail2_ As Collection
Dim colCheqPayment_ As Collection
Dim isChanged_ As Boolean
Dim IncVAT_ As Integer
Dim AdjTax_ As String
Dim NetValue_ As Double
Dim Rest_ As String
Dim ReStaffID_ As String
Dim BankName_ As String
Dim BankNo_ As String

Public Property Get ID() As String
    ID = id_
End Property
Public Property Let ID(ByVal newVal As String)
    id_ = Trim(newVal)
End Property

Public Property Get InvCompanyID() As String
    InvCompanyID = invCompanyID_
End Property
Public Property Let InvCompanyID(ByVal newVal As String)
    If Trim(invCompanyID_) <> Trim(newVal) Then
        invCompanyID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property
Public Property Get BankName() As String
    BankName = BankName_
End Property
Public Property Let BankName(ByVal newVal As String)
    If Trim(BankName_) <> Trim(newVal) Then
        BankName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property
Public Property Get BankNo() As String
    BankNo = BankNo_
End Property
Public Property Let BankNo(ByVal newVal As String)
    If Trim(BankNo_) <> Trim(newVal) Then
        BankNo_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get InvNo() As String
    InvNo = invNo_
End Property
Public Property Let InvNo(ByVal newVal As String)
    If Trim(invNo_) <> Trim(newVal) Then
        invNo_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get InvDate() As String
    InvDate = invDate_
End Property
Public Property Let InvDate(ByVal newVal As String)
    If Trim(invDate_) <> Trim(newVal) Then
        invDate_ = Trim(newVal)
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

Public Property Get isPrintAfterChanged() As Boolean
    If Trim(isPrintAfterChanged_) = "Y" Then
        isPrintAfterChanged = True
    Else
        isPrintAfterChanged = False
    End If
End Property
Public Property Let isPrintAfterChanged(ByVal newVal As Boolean)
Dim tmpStr As String
    If newVal Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(isPrintAfterChanged_) <> Trim(tmpStr) Then
        isPrintAfterChanged_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property

Public Property Get CustomerID() As String
    CustomerID = CustomerID_
End Property
Public Property Let CustomerID(ByVal newVal As String)
    If Trim(CustomerID_) <> Trim(newVal) Then
        CustomerID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get CustomerName() As String
    CustomerName = CustomerName_
End Property
Public Property Let CustomerName(ByVal newVal As String)
    If Trim(CustomerName_) <> Trim(newVal) Then
        CustomerName_ = Trim(newVal)
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

Public Property Get BillTumbol() As String
    BillTumbol = billTumbol_
End Property
Public Property Let BillTumbol(ByVal newVal As String)
    If Trim(billTumbol_) <> Trim(newVal) Then
        billTumbol_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get BillAumphur() As String
    BillAumphur = billAumphur_
End Property
Public Property Let BillAumphur(ByVal newVal As String)
    If Trim(billAumphur_) <> Trim(newVal) Then
        billAumphur_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get BillProvince() As String
    BillProvince = billProvince_
End Property
Public Property Let BillProvince(ByVal newVal As String)
    If Trim(billProvince_) <> Trim(newVal) Then
        billProvince_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get BillZipCode() As String
    BillZipCode = billZipCode_
End Property
Public Property Let BillZipCode(ByVal newVal As String)
    If Trim(billZipCode_) <> Trim(newVal) Then
        billZipCode_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get CreditDay() As Integer
    CreditDay = creditDay_
End Property
Public Property Let CreditDay(ByVal newVal As Integer)
    If creditDay_ <> newVal Then
        creditDay_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get DueDate() As String
    DueDate = dueDate_
End Property
Public Property Let DueDate(ByVal newVal As String)
    If Trim(dueDate_) <> Trim(newVal) Then
        dueDate_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get InvRemark() As String
    InvRemark = invRemark_
End Property
Public Property Let InvRemark(ByVal newVal As String)
    If Trim(invRemark_) <> Trim(newVal) Then
        invRemark_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get PaidCashAmt() As Double
    PaidCashAmt = paidCashAmt_
End Property
Public Property Let PaidCashAmt(ByVal newVal As Double)
    If paidCashAmt_ <> newVal Then
        paidCashAmt_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get InvDetails() As Collection
Dim tmpS As New BWGSearchManager
    If colInvDetail_ Is Nothing Then Set colInvDetail_ = tmpS.PrintInvDetailSearch("PrintInvID='" & id_ & "'")
    Set InvDetails = colInvDetail_
    Set tmpS = Nothing
End Property
Public Property Set InvDetails(ByVal newVal As Collection)
    Set colInvDetail_ = newVal
    isChanged_ = True
End Property

Public Property Get InvDetails2() As Collection
Dim tmpS As New BWGSearchManager
    If colInvDetail2_ Is Nothing Then Set colInvDetail2_ = tmpS.PrintInvDetail_AdjustSearch("PrintInvID='" & id_ & "'")
    Set InvDetails2 = colInvDetail2_
    Set tmpS = Nothing
End Property
Public Property Set InvDetails2(ByVal newVal As Collection)
    Set colInvDetail2_ = newVal
    isChanged_ = True
End Property

Public Sub RequeryInvDetails()
    Set colInvDetail_ = Nothing
End Sub

Public Property Get CheqPayments() As Collection
Dim tmpS As New BWGSearchManager
    If colCheqPayment_ Is Nothing Then Set colCheqPayment_ = tmpS.InvCheqPaymentSearch("PrintInvID='" & id_ & "'")
    Set CheqPayments = colCheqPayment_
    Set tmpS = Nothing
End Property
Public Property Set CheqPayments(ByVal newVal As Collection)
    Set colCheqPayment_ = newVal
    isChanged_ = True
End Property
Public Sub RequeryCheqPayment()
    Set colCheqPayment_ = Nothing
End Sub

Public Property Get isChanged() As Boolean
    isChanged = isChanged_
End Property
Public Property Let isChanged(ByVal newVal As Boolean)
    isChanged_ = newVal
End Property

Public Property Get IncVAT() As Integer
    IncVAT = IncVAT_
End Property
Public Property Let IncVAT(ByVal newVal As Integer)
    IncVAT_ = newVal
    isChanged_ = True
End Property

Public Property Get AdjTax() As String
    AdjTax = AdjTax_
End Property
Public Property Let AdjTax(ByVal newVal As String)
    AdjTax_ = newVal
End Property

Public Property Get NetValue() As Double
    NetValue = NetValue_
End Property
Public Property Let NetValue(ByVal newVal As Double)
    NetValue_ = newVal
    isChanged_ = True
End Property
Public Property Get ReSt() As String
    ReSt = Rest_
End Property
Public Property Let ReSt(ByVal newVal As String)
    Rest_ = ReSt
End Property
Public Property Get ReStaffID() As String
    ReStaffID = ReStaffID_
End Property
Public Property Let ReStaffID(ByVal newVal As String)
    ReStaffID_ = ReStaffID
End Property

Public Sub Save()
On Error GoTo ErrD
Dim x%
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
    sqlStr = "SELECT *, CONVERT(DATETIME, GETDATE(), 103) AS GETDATE FROM tbPrintInvHeader WHERE PrintInvID='" & id_ & "'"
    Set tmpRec = New ADODB.Recordset
    GetDBConnection
    tmpRec.open sqlStr, DBConn, adOpenDynamic, adLockOptimistic
    If tmpRec.EOF Then
        tmpRec.AddNew
        tmpRec("PrintInvID") = id_
    End If
    tmpRec("InvCompanyID") = Trim(invCompanyID_)
    tmpRec("InvDate") = Trim(invDate_)
    If Trim(isCancel_) = "" Then isCancel_ = "N"
    tmpRec("isCancel") = Trim(isCancel_)
    If Trim(isPrintAfterChanged_) = "" Then isPrintAfterChanged_ = "N"
    tmpRec("isPrintAfterChanged") = Trim(isPrintAfterChanged_)
    tmpRec("CustomerID") = Trim(CustomerID_)
    tmpRec("BillAddrID") = Trim(billAddrID_)
    tmpRec("CreditDay") = Trim(creditDay_)
    tmpRec("DueDate") = Trim(dueDate_)
    tmpRec("InvRemark") = Trim(invRemark_)
    tmpRec("PaidCashAmt") = paidCashAmt_
    tmpRec!IncVAT = IncVAT_
    tmpRec!AdjTax = AdjTax_
    tmpRec!NetValue = NetValue_
    
    tmpRec!LastUser = Trim(Trim(CurrentUser.StaffFName) & " " & Trim(CurrentUser.StaffLName))
    tmpRec!LastUpdate = GetSystemDateTime.SystemDateTime
    
    tmpRec!ReState = Rest_
    tmpRec!ReStaffinfoID = ReStaffID_
    tmpRec!BankName = BankName_
    tmpRec!BankNo = BankNo_
    tmpRec.Update
    tmpRec.Requery
    invNo_ = Trim("" & tmpRec("InvNo"))
'    tmpRec.Close
    Set tmpRec = Nothing
    
'    With DBConn
        DBConnExc "Update tbInvCheqPayment Set MarkForDelete='Y' WHERE PrintInvID='" & id_ & "'"
        DBConnExc "Update tbPrintInvDetail Set MarkForDelete='Y' WHERE PrintInvID='" & id_ & "'"
        For x = 1 To colInvDetail_.Count
            colInvDetail_(x).Save
        Next
        For x = 1 To colCheqPayment_.Count
            colCheqPayment_(x).Save
        Next
        DBConnExc "Delete tbInvCheqPayment WHERE MarkForDelete='Y' And PrintInvID='" & id_ & "'"
        DBConnExc "Delete tbPrintInvDetail WHERE MarkForDelete='Y' And PrintInvID='" & id_ & "'"
'    End With
    
        DBConnExc "Update tbPrintRVDetail_Adjust Set MarkForDelete='Y' WHERE PrintInvID='" & id_ & "'"
        DBConnExc "UPDATE tbAdjustDeptHeader SET Status = 'N' WHERE AdjustID IN (SELECT AdjustID FROM tbPrintRVDetail_Adjust WHERE PrintInvID='" & id_ & "')"
        For x = 1 To colInvDetail2_.Count
            colInvDetail2_(x).Save
        Next
        DBConnExc "Delete tbPrintRVDetail_Adjust WHERE MarkForDelete='Y' And PrintInvID='" & id_ & "'"
    
    isChanged_ = False
    CloseDBConnection
    Exit Sub
ErrD:
    GetMsg "", vbCritical
    Exit Sub
End Sub

Public Function Delete() As String
On Error GoTo delErr:
    DBConnExc "Delete tbPrintInvHeader WHERE PrintInvID='" & id_ & "'"
    Delete = ""
    Exit Function
delErr:
    Delete = err.Description
End Function

Public Sub SetDataFromRecordset(selRec As ADODB.Recordset)
Dim x%
    id_ = Trim("" & selRec("PrintInvID"))
    invCompanyID_ = Trim("" & selRec("InvCompanyID"))
    invNo_ = Trim("" & selRec("InvNo"))
    invDate_ = Trim("" & selRec("InvDate"))
    isCancel_ = Trim("" & selRec("isCancel"))
    cancelDate_ = Trim("" & selRec("CancelDate"))
    isPrintAfterChanged_ = Trim("" & selRec("isPrintAfterChanged"))
    CustomerID_ = Trim("" & selRec("CustomerID"))
    CustomerName_ = Trim("" & selRec("CustomerName"))
    billAddrID_ = Trim("" & selRec("BillAddrID"))
    billAddr_ = Trim("" & selRec("Addr1"))
    billTumbol_ = Trim("" & selRec("TumbolName"))
    billAumphur_ = Trim("" & selRec("AumphurName"))
    billProvince_ = Trim("" & selRec("ProvinceName"))
    billZipCode_ = Trim("" & selRec("ZipCode"))
    creditDay_ = selRec("CreditDay")
    dueDate_ = Trim("" & selRec("DueDate"))
    invRemark_ = Trim("" & selRec("InvRemark"))
    paidCashAmt_ = selRec("PaidCashAmt")
    IncVAT_ = IIf(IsNull(selRec!IncVAT), 0, selRec!IncVAT)
    AdjTax_ = Trim("" & selRec("AdjTax"))
    NetValue_ = IIf(IsNull(selRec!NetValue), 0, selRec!NetValue)
    Rest_ = Trim("" & selRec("ReState"))
    ReStaffID_ = Trim("" & selRec("ReStaffinfoID"))
    BankName_ = Trim("" & selRec("BankName"))
    BankNo_ = Trim("" & selRec("BankNo"))
End Sub

Public Function CancelBill() As String
On Error GoTo ErrD
Dim tmpRec As New ADODB.Recordset
Dim sqlStr As String
    sqlStr = "SELECT * FROM vw_CheckInvRE WHERE isCancel='N' And isDeleted='N' And PrintInvID='" & id_ & "'"
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec.EOF Then
        DBConnExc "Update tbPrintInvHeader Set isCancel='Y',CancelDate='" & TodayDate & "' WHERE PrintInvID='" & id_ & "'"
         isCancel_ = "Y"
         cancelDate_ = TodayDate
        CancelBill = ""
    Else
        CancelBill = "ไม่สามารถยกเลิกใบกำกับนี้ได้ เพราะมีการทำการรับชำระเงินแล้ว"
    End If
    tmpRec.Close
    Set tmpRec = Nothing
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function
