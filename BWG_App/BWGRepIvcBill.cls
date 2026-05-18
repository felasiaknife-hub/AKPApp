VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BWGRepIvcBill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim BillingNo_ As String
Dim IssuedDate_ As String
Dim AmtB4VAT_ As String
Dim VATAmt_ As String
Dim TotalBillAmt_ As String
Dim UserN_ As String
Dim CustomerName_ As String
Public Property Get BillingNo() As String
    BillingNo = BillingNo_
End Property
Public Property Let BillingNo(ByVal newVal As String)
    BillingNo_ = Trim(newVal)
End Property
Public Property Get IssuedDate() As String
    IssuedDate = IssuedDate_
End Property
Public Property Let IssuedDate(ByVal newVal As String)
    If Trim(IssuedDate_) <> Trim(newVal) Then
        IssuedDate_ = Trim(newVal)
    End If
End Property
Public Property Get AmtB4VAT() As Double
    AmtB4VAT = AmtB4VAT_
End Property
Public Property Let AmtB4VAT(ByVal newVal As Double)
    If Trim(AmtB4VAT_) <> Trim(newVal) Then
        AmtB4VAT_ = Trim(newVal)
    End If
End Property
Public Property Get VATAmt() As Double
    VATAmt = VATAmt_
End Property
Public Property Let VATAmt(ByVal newVal As Double)
    If Trim(VATAmt_) <> Trim(newVal) Then
        VATAmt_ = Trim(newVal)
    End If
End Property
Public Property Get TotalBillAmt() As Double
    TotalBillAmt = TotalBillAmt_
End Property
Public Property Let TotalBillAmt(ByVal newVal As Double)
    If Trim(TotalBillAmt_) <> Trim(newVal) Then
        TotalBillAmt_ = Trim(newVal)
    End If
End Property
Public Property Get UserN() As String
    UserN = UserN_
End Property
Public Property Let UserN(ByVal newVal As String)
    If Trim(UserN_) <> Trim(newVal) Then
        UserN_ = Trim(newVal)
    End If
End Property
Public Property Get CustomerName() As String
    CustomerName = CustomerName_
End Property
Public Property Let CustomerName(ByVal newVal As String)
    If CustomerName_ <> newVal Then
        CustomerName_ = newVal
    End If
End Property
Public Sub Save()
On Error GoTo ErrD
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpRec2 As ADODB.Recordset
Dim sqlStr2 As String
Dim i As Integer
Dim TSFee As Integer

    sqlStr = "SELECT *, CONVERT(DATETIME, GETDATE(), 103) AS GETDATE FROM TbTmpRepIvcBill "
    Set tmpRec = New ADODB.Recordset
    GetDBConnection
    tmpRec.open sqlStr, DBConn, adOpenDynamic, adLockOptimistic
    tmpRec.AddNew
    tmpRec("BillingNo") = BillingNo_
    tmpRec("IssuedDate") = IssuedDate_
    tmpRec("AmtB4VAT") = AmtB4VAT_
    tmpRec("VATAmt") = VATAmt_
    tmpRec("TotalBillAmt") = TotalBillAmt_
    tmpRec("UserN") = UserN_
    tmpRec("CustomerName") = CustomerName_
    tmpRec.Update
'    tmpRec.Close
    Set tmpRec = Nothing
    CloseDBConnection
    Exit Sub
ErrD:
    GetMsg "", vbCritical
    Exit Sub
End Sub
Public Function Delete() As String
On Error GoTo delErr:
    DBConnExc "Delete TbTmpRepIvcBill WHERE BillingNo='" & BillingNo_ & "' and UserN='" & UserN_ & "' "
    Delete = ""
    Exit Function
delErr:
    Delete = err.Description
End Function
Public Function DeleteAll() As String
    DBConnExc "Delete TbTmpRepIvcBill WHERE UserN='" & UserN_ & "' "
    Exit Function
End Function

