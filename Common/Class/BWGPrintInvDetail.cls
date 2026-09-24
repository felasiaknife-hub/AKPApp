VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BWGPrintInvDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim id_ As String
Dim printInvID_ As String
Dim rowID_ As Integer
Dim billingHeaderID_ As String
Dim BillingNo_ As String
Dim billingDate_  As String
Dim billingTotalAmt_ As Double
Dim billingPaidAmt_ As Double
Dim wTaxRate_ As Double
Dim invBillDesc_ As String
Dim realRecAmt_ As Double
Dim isChanged_ As Boolean

Public Property Get ID() As String
    ID = id_
End Property
Public Property Let ID(ByVal newVal As String)
    id_ = Trim(newVal)
End Property

Public Property Get PrintInvID() As String
    PrintInvID = printInvID_
End Property
Public Property Let PrintInvID(ByVal newVal As String)
    If Trim(printInvID_) <> Trim(newVal) Then
        printInvID_ = Trim(newVal)
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

Public Property Get BillingHeaderID() As String
    BillingHeaderID = billingHeaderID_
End Property
Public Property Let BillingHeaderID(ByVal newVal As String)
    If Trim(billingHeaderID_) <> Trim(newVal) Then
        billingHeaderID_ = Trim(newVal)
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

Public Property Get BillingDate() As String
    BillingDate = billingDate_
End Property
Public Property Let BillingDate(ByVal newVal As String)
    If Trim(billingDate_) <> Trim(newVal) Then
        billingDate_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get BillingTotalAmt() As Double
    BillingTotalAmt = billingTotalAmt_
End Property
Public Property Let BillingTotalAmt(ByVal newVal As Double)
    If billingTotalAmt_ <> newVal Then
        billingTotalAmt_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get BillingPaidAmt() As Double
    BillingPaidAmt = billingPaidAmt_
End Property
Public Property Let BillingPaidAmt(ByVal newVal As Double)
    If billingPaidAmt_ <> newVal Then
        billingPaidAmt_ = newVal
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

Public Property Get InvBillDesc() As String
    InvBillDesc = invBillDesc_
End Property
Public Property Let InvBillDesc(ByVal newVal As String)
    If Trim(invBillDesc_) <> Trim(newVal) Then
        invBillDesc_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get RealRecAmt() As Double
    RealRecAmt = realRecAmt_
End Property
Public Property Let RealRecAmt(ByVal newVal As Double)
    If realRecAmt_ <> newVal Then
        realRecAmt_ = newVal
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
    sqlStr = "SELECT * FROM tbPrintInvDetail WHERE PrintInvDetID='" & id_ & "'"
    Set tmpRec = New ADODB.Recordset
    GetDBConnection
    tmpRec.open sqlStr, DBConn, adOpenDynamic, adLockOptimistic
    If tmpRec.EOF Then
        tmpRec.AddNew
        tmpRec("PrintInvDetID") = id_
    End If
    tmpRec("PrintInvID") = Trim(printInvID_)
    tmpRec("RowID") = Trim(rowID_)
    tmpRec("BillingHeaderID") = Trim(billingHeaderID_)
    tmpRec("InvBillDesc") = Trim(invBillDesc_)
    tmpRec("RealRecAmt") = Trim(realRecAmt_)
    tmpRec("MarkForDelete") = "N"
    tmpRec.Update
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
    DBConnExc "Delete tbPrintInvDetail WHERE PrintInvDetID='" & id_ & "'"
    Delete = ""
    Exit Function
delErr:
    Delete = err.Description
End Function

Public Sub SetDataFromRecordset(selRec As ADODB.Recordset)
    id_ = Trim("" & selRec("PrintInvDetID"))
    printInvID_ = Trim("" & selRec("PrintInvID"))
    rowID_ = Trim("" & selRec("RowID"))
    billingHeaderID_ = Trim("" & selRec("BillingHeaderID"))
    BillingNo_ = Trim("" & selRec("BillingNo"))
    billingDate_ = Trim("" & selRec("IssuedDate"))
    invBillDesc_ = Trim("" & selRec("InvBillDesc"))
    If IsNull(selRec("RealRecAmt")) = False Then realRecAmt_ = Trim("" & selRec("RealRecAmt"))
    If IsNull(selRec("TotalBillAmt")) = False Then billingTotalAmt_ = Trim("" & selRec("TotalBillAmt"))
    If IsNull(selRec("PaidAmt")) = False Then billingPaidAmt_ = Trim("" & selRec("PaidAmt"))
    
    If isExist(selRec, "WTax") Then
        If IsNull(selRec("WTax")) = False Then wTaxRate_ = Trim("" & selRec("WTax"))
    End If
End Sub

