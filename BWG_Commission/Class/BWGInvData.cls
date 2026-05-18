VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BWGInvData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'***This part for Invoice
Dim InvID_ As String
Dim MenifestID_ As String
Dim BillingNo_ As String
Dim DocNo_ As String
Dim InvDate_ As String
Dim PayDueDate_ As String
Dim DiscountAmt_ As Double
Dim AmtB4VAT_ As Double
Dim VATAmt_ As Double
Dim TotalBillAmt_ As Double
Dim VATRate_ As Double
Dim ManiAmtB4VAT_ As Double
Dim ManiDiscount_ As Double
Dim MeniVATAmt_ As Double
Dim TotalAmtToPay_ As Double
Dim isCancel_ As String
Dim isChanged_ As Boolean

Public Property Get InvID() As String
    InvID = InvID_
End Property

Public Property Let InvID(ByVal newVal As String)
    InvID_ = Trim(newVal)
End Property

Public Property Get MenifestID() As String
    MenifestID = MenifestID_
End Property

Public Property Let MenifestID(ByVal newVal As String)
    If Trim(MenifestID_) <> Trim(newVal) Then
        MenifestID_ = Trim(newVal)
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

Public Property Get DocNo() As String
    DocNo = DocNo_
End Property

Public Property Let DocNo(ByVal newVal As String)
    If Trim(DocNo_) <> Trim(newVal) Then
        DocNo_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get InvDate() As String
    InvDate = InvDate_
End Property

Public Property Let InvDate(ByVal newVal As String)
    If Trim(InvDate_) <> Trim(newVal) Then
        InvDate_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get PayDueDate() As String
    PayDueDate = PayDueDate_
End Property

Public Property Let PayDueDate(ByVal newVal As String)
    If Trim(PayDueDate_) <> Trim(newVal) Then
        PayDueDate_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get DiscountAmt() As Double
    DiscountAmt = DiscountAmt_
End Property

Public Property Let DiscountAmt(ByVal newVal As Double)
    If Trim(DiscountAmt_) <> Trim(newVal) Then
        DiscountAmt_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get AmtB4VAT() As Double
    AmtB4VAT = AmtB4VAT_
End Property

Public Property Let AmtB4VAT(ByVal newVal As Double)
    If Trim(AmtB4VAT_) <> Trim(newVal) Then
        AmtB4VAT_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get VATAmt() As Double
    VATAmt = VATAmt_
End Property

Public Property Let VATAmt(ByVal newVal As Double)
    If Trim(VATAmt_) <> Trim(newVal) Then
        VATAmt_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get TotalBillAmt() As Double
    TotalBillAmt = TotalBillAmt_
End Property

Public Property Let TotalBillAmt(ByVal newVal As Double)
    If Trim(TotalBillAmt_) <> Trim(newVal) Then
        TotalBillAmt_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get VATRate() As Double
    VATRate = VATRate_
End Property

Public Property Let VATRate(ByVal newVal As Double)
    If Trim(VATRate_) <> Trim(newVal) Then
        VATRate_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get ManiAmtB4VAT() As Double
    ManiAmtB4VAT = ManiAmtB4VAT_
End Property

Public Property Let ManiAmtB4VAT(ByVal newVal As Double)
    If Trim(ManiAmtB4VAT_) <> Trim(newVal) Then
        ManiAmtB4VAT_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get ManiDiscount() As Double
    ManiDiscount = ManiDiscount_
End Property

Public Property Let ManiDiscount(ByVal newVal As Double)
    If Trim(ManiDiscount_) <> Trim(newVal) Then
        ManiDiscount_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get MeniVATAmt() As Double
    MeniVATAmt = MeniVATAmt_
End Property

Public Property Let MeniVATAmt(ByVal newVal As Double)
    If Trim(MeniVATAmt_) <> Trim(newVal) Then
        MeniVATAmt_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get TotalAmtToPay() As Double
    TotalAmtToPay = TotalAmtToPay_
End Property

Public Property Let TotalAmtToPay(ByVal newVal As Double)
    If Trim(TotalAmtToPay_) <> Trim(newVal) Then
        TotalAmtToPay_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get isCancel() As String
    isCancel = isCancel_
End Property

Public Property Let isCancel(ByVal newVal As String)
    If Trim(isCancel_) <> Trim(newVal) Then
    End If
End Property

Public Property Get isChanged() As Boolean
    isChanged = isChanged_
End Property

Public Property Let isChanged(ByVal newVal As Boolean)
    isChanged_ = newVal
End Property

Public Sub SetDataFromRecordset(selRec As ADODB.Recordset)
    InvID_ = Trim(IIf(IsNull(selRec!BillingDetailID), "", selRec!BillingDetailID))
    MenifestID_ = Trim(IIf(IsNull(selRec!MenifestID), "", selRec!MenifestID))
    BillingNo_ = Trim(IIf(IsNull(selRec!BillingNo), "", selRec!BillingNo))
    DocNo_ = Trim(IIf(IsNull(selRec!DocNo), "", selRec!DocNo))
    InvDate_ = Trim(IIf(IsNull(selRec!InvDate), "", selRec!InvDate))
    PayDueDate_ = Trim(IIf(IsNull(selRec!PayDueDate), "", selRec!PayDueDate))
    DiscountAmt_ = IIf(IsNull(selRec!DiscountAmt), 0, selRec!DiscountAmt)
    AmtB4VAT_ = selRec!AmtB4VAT
    VATAmt_ = selRec!VATAmt
    TotalBillAmt_ = selRec!TotalBillAmt
    VATRate_ = selRec!VATRate
    ManiAmtB4VAT_ = selRec!ManiAmtB4VAT
    ManiDiscount_ = selRec!ManiDiscount
    MeniVATAmt_ = selRec!MeniVATAmt
    TotalAmtToPay_ = selRec!TotalAmtToPay
    isCancel_ = Trim(IIf(IsNull(selRec!isCancel), "N", selRec!isCancel))
End Sub
