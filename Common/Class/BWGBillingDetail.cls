VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BWGBillingDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim id_ As String
Dim billHeaderID_ As String
Dim rowID_ As Integer
Dim MenifestID_ As String
Dim menifestDocNo_ As String
Dim menifestNo_ As String
Dim menifestDate_ As String
Dim menifestAddr_ As String
Dim menifestAmt_ As Double
Dim menifestOtherDesc_ As String
Dim menifestOtherAmt_ As Double
Dim wasteType_  As String
Dim billDesc_ As String
Dim amtB4Vat_ As Double
Dim discountAmt_ As Double
Dim vatAmt_ As Double
Dim totalAmtToPay_ As Double
Dim isCancel_ As String
Dim discountType_ As String
Dim vatRate_ As Double
Dim colMenifestDet_ As Collection
Dim isChanged_ As Boolean
Dim CompleteWithError_ As Boolean

Public Property Get ID() As String
    ID = id_
End Property
Public Property Let ID(ByVal newVal As String)
    id_ = Trim(newVal)
End Property

Public Property Get BillHeaderID() As String
    BillHeaderID = Trim(billHeaderID_)
End Property
Public Property Let BillHeaderID(ByVal newVal As String)
    If Trim(billHeaderID_) <> Trim(newVal) Then
        billHeaderID_ = Trim(newVal)
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

Public Property Get menifestID() As String
    menifestID = Trim(MenifestID_)
End Property
Public Property Let menifestID(ByVal newVal As String)
    If Trim(MenifestID_) <> Trim(newVal) Then
        MenifestID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get MenifestDocNo() As String
    MenifestDocNo = Trim(menifestDocNo_)
End Property
Public Property Let MenifestDocNo(ByVal newVal As String)
    If Trim(menifestDocNo_) <> Trim(newVal) Then
        menifestDocNo_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get MenifestNo() As String
    MenifestNo = Trim(menifestNo_)
End Property
Public Property Let MenifestNo(ByVal newVal As String)
    If Trim(menifestNo_) <> Trim(newVal) Then
        menifestNo_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get MenifestDate() As String
    MenifestDate = Trim(menifestDate_)
End Property
Public Property Let MenifestDate(ByVal newVal As String)
    If Trim(menifestDate_) <> Trim(newVal) Then
        menifestDate_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get MenifestAddr() As String
    MenifestAddr = Trim(menifestAddr_)
End Property
Public Property Let MenifestAddr(ByVal newVal As String)
    If Trim(menifestAddr_) <> Trim(newVal) Then
        menifestAddr_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get MenifestAmt() As Double
    MenifestAmt = menifestAmt_
End Property
Public Property Let MenifestAmt(ByVal newVal As Double)
    If menifestAmt_ <> newVal Then
        menifestAmt_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get MenifestOtherDesc() As String
    MenifestOtherDesc = Trim(menifestOtherDesc_)
End Property
Public Property Let MenifestOtherDesc(ByVal newVal As String)
    If Trim(menifestOtherDesc_) <> Trim(newVal) Then
        menifestOtherDesc_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get MenifestOtherAmt() As Double
    MenifestOtherAmt = menifestOtherAmt_
End Property
Public Property Let MenifestOtherAmt(ByVal newVal As Double)
    If menifestOtherAmt_ <> newVal Then
        menifestOtherAmt_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get WasteType() As String
    WasteType = Trim(wasteType_)
End Property
Public Property Let WasteType(ByVal newVal As String)
    If Trim(wasteType_) <> Trim(newVal) Then
        wasteType_ = Trim(newVal)
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

Public Property Get DiscountType() As String
    DiscountType = Trim(discountType_)
End Property
Public Property Let DiscountType(ByVal newVal As String)
    If Trim(discountType_) <> Trim(newVal) Then
        discountType_ = Trim(newVal)
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

Public Property Get isChanged() As Boolean
    isChanged = isChanged_
End Property
Public Property Let isChanged(ByVal newVal As Boolean)
    isChanged_ = newVal
End Property

Public Property Get CompleteWithError() As Boolean
    CompleteWithError = CompleteWithError_
End Property
Public Property Let CompleteWithError(ByVal newVal As Boolean)
    CompleteWithError_ = newVal
End Property

Public Sub Save(Optional ByVal x As Integer = 0)
On Error GoTo ErrD
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
    GetDBConnection
    sqlStr = "SELECT * FROM tbBillingDetail WHERE BillingDetailID='" & id_ & "'"
    Set tmpRec = New ADODB.Recordset
    tmpRec.open sqlStr, DBConn, adOpenKeyset, adLockOptimistic
    If tmpRec.EOF Then
        tmpRec.AddNew
        tmpRec("BillingDetailID") = id_
    End If
    tmpRec("BillingHeaderID") = billHeaderID_
    If rowID_ <> x Then
        tmpRec("RowID") = x
    Else
        tmpRec("RowID") = rowID_
    End If
    tmpRec("SellID") = MenifestID_
    tmpRec("BillDesc") = billDesc_
    tmpRec("AmtB4VAT") = amtB4Vat_
    tmpRec("DiscountAmt") = discountAmt_
    tmpRec("VATAmt") = vatAmt_
    tmpRec("TotalAmtToPay") = totalAmtToPay_
    If Trim(isCancel_) = "" Then isCancel_ = "N"
    tmpRec("isCancel") = isCancel_
    tmpRec("DiscountType") = discountType_
    tmpRec("VATRate") = vatRate_
    tmpRec!isCancel = "N"
    tmpRec("MarkForDelete") = "N"
    tmpRec.Update
'    tmpRec.Close
    Set tmpRec = Nothing
    isChanged_ = False
    If Not CompleteWithError_ = False Then CompleteWithError_ = False
    CloseDBConnection
    Exit Sub
ErrD:
    frmMain.Stb1.Panels(2).Text = "System Error: " & err.Number
    'GetMsg "", vbCritical
    CompleteWithError_ = True
    Exit Sub
End Sub

Public Function Delete() As String
On Error GoTo delErr:
    DBConnExc "Delete tbBillingDetail WHERE BillingDetailID='" & id_ & "'"
    Delete = ""
    Exit Function
delErr:
    Delete = err.Description
End Function

Public Sub SetDataFromRecordset(selRec As ADODB.Recordset)
    id_ = Trim("" & selRec("BillingDetailID"))
    billHeaderID_ = Trim("" & selRec("BillingHeaderID"))
    rowID_ = selRec("RowID")
    MenifestID_ = Trim("" & selRec("SellID"))
    menifestDocNo_ = Trim("" & selRec("DocNo"))
    menifestNo_ = Trim("" & selRec("MenifestNo"))
    menifestDate_ = Trim("" & selRec("WorkDate"))
    menifestAddr_ = Trim("" & selRec("SiteName")) & " " & Trim("" & selRec("Addr1")) & " " & Trim("" & selRec("TumbolName")) & " " & Trim("" & selRec("AumphurName")) & " " & Trim("" & selRec("ProvinceName")) & " " & Trim("" & selRec("ZipCode"))
    menifestAmt_ = selRec("InCollectFee") + selRec("InTransFee") + selRec("InWorkerFee") + selRec("InEquipFee") + selRec("InOtherFee")
    menifestOtherDesc_ = Trim("" & selRec("InOtherDesc"))
    menifestOtherAmt_ = Trim("" & selRec("InOtherFee"))
    billDesc_ = Trim("" & selRec("BillDesc"))
    amtB4Vat_ = selRec("AmtB4VAT")
    discountAmt_ = selRec("DiscountAmt")
    vatAmt_ = selRec("VATAmt")
    totalAmtToPay_ = selRec("TotalAmtToPay")
    isCancel_ = Trim("" & selRec("isCancel"))
    wasteType_ = Trim("" & selRec("NonOrHz"))
    discountType_ = Trim("" & selRec("DiscountType"))
    vatRate_ = selRec("VATRate")
End Sub

Public Function GetJobDetails() As Collection
Dim tmpS As New BWGSearchManager
    Set GetJobDetails = tmpS.JobDetailInTripSearch("MenifestID='" & MenifestID_ & "'", "WasteType,TreatmentUnitID")
    Set tmpS = Nothing
End Function
