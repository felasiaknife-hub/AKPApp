VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BWGPayDeptHeader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim id_ As String
Dim payDeptBillNo_ As String
Dim payDate_ As String
Dim companyID_ As String
Dim companyName_ As String
Dim supplierID_ As String
Dim supplierName_ As String
Dim payDueDate_ As String
Dim supBillingNo_ As String
Dim supInvNo_ As String
Dim supReceiptNo_ As String
Dim billRemark_ As String
Dim totalServiceFee_ As Double
Dim vatAmt_ As Double
Dim grandTotalFee_ As Double
Dim isChanged_ As Boolean

Public Property Get ID() As String
    ID = id_
End Property
Public Property Let ID(ByVal newVal As String)
    id_ = Trim(newVal)
End Property

Public Property Get PayDeptBillNo() As String
    PayDeptBillNo = payDeptBillNo_
End Property
Public Property Let PayDeptBillNo(ByVal newVal As String)
    If Trim(payDeptBillNo_) <> Trim(newVal) Then
        payDeptBillNo_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get PayDate() As String
    PayDate = payDate_
End Property
Public Property Let PayDate(ByVal newVal As String)
    If Trim(payDate_) <> Trim(newVal) Then
        payDate_ = Trim(newVal)
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

Public Property Get PayDueDate() As String
    PayDueDate = payDueDate_
End Property
Public Property Let PayDueDate(ByVal newVal As String)
    If Trim(payDueDate_) <> Trim(newVal) Then
        payDueDate_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get SupBillingNo() As String
    SupBillingNo = supBillingNo_
End Property
Public Property Let SupBillingNo(ByVal newVal As String)
    If Trim(supBillingNo_) <> Trim(newVal) Then
        supBillingNo_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get SupInvNo() As String
    SupInvNo = supInvNo_
End Property
Public Property Let SupInvNo(ByVal newVal As String)
    If Trim(supInvNo_) <> Trim(newVal) Then
        supInvNo_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get SupReceiptNo() As String
    SupReceiptNo = supReceiptNo_
End Property
Public Property Let SupReceiptNo(ByVal newVal As String)
    If Trim(supReceiptNo_) <> Trim(newVal) Then
        supReceiptNo_ = Trim(newVal)
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

Public Property Get TotalServiceFee() As Double
    TotalServiceFee = totalServiceFee_
End Property
Public Property Let TotalServiceFee(ByVal newVal As Double)
    If totalServiceFee_ <> newVal Then
        totalServiceFee_ = newVal
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

Public Property Get GrandTotalFee() As Double
    GrandTotalFee = grandTotalFee_
End Property
Public Property Let GrandTotalFee(ByVal newVal As Double)
    If grandTotalFee_ <> newVal Then
        grandTotalFee_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get isChanged() As Boolean
    isChanged = isChanged_
End Property
Public Property Let isChanged(ByVal newVal As Boolean)
    isChanged_ = newVal
End Property

'Public Sub Save()
''ทำตรงนี้ต่อ
'Dim tmpRec As ADODB.Recordset
'    Set tmpRec = New ADODB.Recordset
'    tmpRec.Open "SELECT * FROM tbLABUnit WHERE LABUnitID='" & id_ & "'", DBConn, adOpenDynamic, adLockOptimistic
'    If tmpRec.EOF Then
'        tmpRec.AddNew
'        tmpRec("LABUnitID") = id_
'    End If
'    tmpRec("UnitName") = Trim(unitName_)
'    tmpRec.Update
'    tmpRec.Close
'    Set tmpRec = Nothing
'    isChanged_ = False
'End Sub
'
'Public Function Delete() As String
'On Error GoTo delErr:
'    DBConnExc "Delete tbLABUnit WHERE LABUnitID='" & id_ & "'"
'    Delete = ""
'    Exit Function
'delErr:
'    Delete = Err.Description
'End Function
'
'Public Sub SetDataFromRecordset(selRec As ADODB.Recordset)
'    id_ = selRec("LABUnitID")
'    unitName_ = Trim("" & selRec("UnitName"))
'End Sub
'
