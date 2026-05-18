VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BWGCustomerBillInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim id_ As String
Dim customerID_ As String
Dim customerName_ As String
Dim contactName_ As String
Dim telNo_ As String
Dim faxNo_ As String
Dim billingDetail_ As String
Dim cheqDetail_ As String
Dim needPO_ As String
Dim needInvoice_ As String
Dim needBillingNote_ As String
Dim needContract_ As String
Dim needQuotation_ As String
Dim otherDocDesc_ As String
Dim remDesc_ As String
Dim billRecType_ As String
Dim billDay1_ As Integer
Dim billDay2_ As Integer
Dim billDay3_ As Integer
Dim billDay4_ As Integer
Dim billDay5_ As Integer
Dim billWeekCount_ As Integer
Dim cheqRecType_ As String
Dim cheqDay1_ As Integer
Dim cheqDay2_ As Integer
Dim cheqDay3_ As Integer
Dim cheqDay4_ As Integer
Dim cheqDay5_ As Integer
Dim cheqWeekCount_ As Integer
Dim isChanged_ As Boolean


Public Property Get ID() As String
    ID = id_
End Property
Public Property Let ID(ByVal newVal As String)
    id_ = Trim(newVal)
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

Public Property Get ContactName() As String
    ContactName = contactName_
End Property
Public Property Let ContactName(ByVal newVal As String)
    If Trim(contactName_) <> Trim(newVal) Then
        contactName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get TelNo() As String
    TelNo = telNo_
End Property
Public Property Let TelNo(ByVal newVal As String)
    If Trim(telNo_) <> Trim(newVal) Then
        telNo_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get FaxNo() As String
    FaxNo = faxNo_
End Property
Public Property Let FaxNo(ByVal newVal As String)
    If Trim(faxNo_) <> Trim(newVal) Then
        faxNo_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get BillingDetail() As String
    BillingDetail = billingDetail_
End Property
Public Property Let BillingDetail(ByVal newVal As String)
    If Trim(billingDetail_) <> Trim(newVal) Then
        billingDetail_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get CheqDetail() As String
    CheqDetail = cheqDetail_
End Property
Public Property Let CheqDetail(ByVal newVal As String)
    If Trim(cheqDetail_) <> Trim(newVal) Then
        cheqDetail_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get NeedPO() As Boolean
    If needPO_ = "Y" Then
        NeedPO = True
    Else
        NeedPO = False
    End If
End Property
Public Property Let NeedPO(ByVal newVal As Boolean)
Dim tmpStr As String
    If newVal Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(needPO_) <> Trim(tmpStr) Then
        needPO_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property

Public Property Get NeedInvoice() As Boolean
    If needInvoice_ = "Y" Then
        NeedInvoice = True
    Else
        NeedInvoice = False
    End If
End Property
Public Property Let NeedInvoice(ByVal newVal As Boolean)
Dim tmpStr As String
    If newVal Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(needInvoice_) <> Trim(tmpStr) Then
        needInvoice_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property

Public Property Get NeedBillingNote() As Boolean
    If needBillingNote_ = "Y" Then
        NeedBillingNote = True
    Else
        NeedBillingNote = False
    End If
End Property
Public Property Let NeedBillingNote(ByVal newVal As Boolean)
Dim tmpStr As String
    If newVal Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(needBillingNote_) <> Trim(tmpStr) Then
        needBillingNote_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property

Public Property Get NeedContract() As Boolean
    If needContract_ = "Y" Then
        NeedContract = True
    Else
        NeedContract = False
    End If
End Property
Public Property Let NeedContract(ByVal newVal As Boolean)
Dim tmpStr As String
    If newVal Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(needContract_) <> Trim(tmpStr) Then
        needContract_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property

Public Property Get NeedQuotation() As Boolean
    If needQuotation_ = "Y" Then
        NeedQuotation = True
    Else
        NeedQuotation = False
    End If
End Property
Public Property Let NeedQuotation(ByVal newVal As Boolean)
Dim tmpStr As String
    If newVal Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(needQuotation_) <> Trim(tmpStr) Then
        needQuotation_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property

Public Property Get OtherDocDesc() As String
    OtherDocDesc = otherDocDesc_
End Property
Public Property Let OtherDocDesc(ByVal newVal As String)
    If Trim(otherDocDesc_) <> Trim(newVal) Then
        otherDocDesc_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get RemDesc() As String
    RemDesc = remDesc_
End Property
Public Property Let RemDesc(ByVal newVal As String)
    If Trim(remDesc_) <> Trim(newVal) Then
        remDesc_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get BillRecType() As String
    BillRecType = billRecType_
End Property
Public Property Let BillRecType(ByVal newVal As String)
    If Trim(billRecType_) <> Trim(newVal) Then
        billRecType_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get BillDay1() As Integer
    BillDay1 = billDay1_
End Property
Public Property Let BillDay1(ByVal newVal As Integer)
    If billDay1_ <> newVal Then
        billDay1_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get BillDay2() As Integer
    BillDay2 = billDay2_
End Property
Public Property Let BillDay2(ByVal newVal As Integer)
    If billDay2_ <> newVal Then
        billDay2_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get BillDay3() As Integer
    BillDay3 = billDay3_
End Property
Public Property Let BillDay3(ByVal newVal As Integer)
    If billDay3_ <> newVal Then
        billDay3_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get BillDay4() As Integer
    BillDay4 = billDay4_
End Property
Public Property Let BillDay4(ByVal newVal As Integer)
    If billDay4_ <> newVal Then
        billDay4_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get BillDay5() As Integer
    BillDay5 = billDay5_
End Property
Public Property Let BillDay5(ByVal newVal As Integer)
    If billDay5_ <> newVal Then
        billDay5_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get BillWeekCount() As Integer
    BillWeekCount = billWeekCount_
End Property
Public Property Let BillWeekCount(ByVal newVal As Integer)
    If billWeekCount_ <> newVal Then
        billWeekCount_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get CheqRecType() As String
    CheqRecType = Trim(cheqRecType_)
End Property
Public Property Let CheqRecType(ByVal newVal As String)
    If Trim(cheqRecType_) <> Trim(newVal) Then
        cheqRecType_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get CheqDay1() As Integer
    CheqDay1 = cheqDay1_
End Property
Public Property Let CheqDay1(ByVal newVal As Integer)
    If cheqDay1_ <> newVal Then
        cheqDay1_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get CheqDay2() As Integer
    CheqDay2 = cheqDay2_
End Property
Public Property Let CheqDay2(ByVal newVal As Integer)
    If cheqDay2_ <> newVal Then
        cheqDay2_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get CheqDay3() As Integer
    CheqDay3 = cheqDay3_
End Property
Public Property Let CheqDay3(ByVal newVal As Integer)
    If cheqDay3_ <> newVal Then
        cheqDay3_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get CheqDay4() As Integer
    CheqDay4 = cheqDay4_
End Property
Public Property Let CheqDay4(ByVal newVal As Integer)
    If cheqDay4_ <> newVal Then
        cheqDay4_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get CheqDay5() As Integer
    CheqDay5 = cheqDay5_
End Property
Public Property Let CheqDay5(ByVal newVal As Integer)
    If cheqDay5_ <> newVal Then
        cheqDay5_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get CheqWeekCount() As Integer
    CheqWeekCount = cheqWeekCount_
End Property
Public Property Let CheqWeekCount(ByVal newVal As Integer)
    If cheqWeekCount_ <> newVal Then
        cheqWeekCount_ = newVal
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
Dim tmpRec As ADODB.Recordset
    Set tmpRec = New ADODB.Recordset
    tmpRec.Open "Select * from tbCustomerBillingInfo Where BillingInfoID='" & id_ & "'", DBConn, adOpenDynamic, adLockOptimistic
    If tmpRec.EOF Then
        tmpRec.AddNew
        tmpRec("BillingInfoID") = id_
    End If
    tmpRec("CustomerID") = Trim(customerID_)
    tmpRec("ContactName") = Trim(contactName_)
    tmpRec("TelNo") = Trim(telNo_)
    tmpRec("FaxNo") = Trim(faxNo_)
    tmpRec("BillingDetail") = Trim(billingDetail_)
    tmpRec("RecCheqDetail") = Trim(cheqDetail_)
    tmpRec("NeedPO") = Trim(needPO_)
    tmpRec("NeedInvoice") = Trim(needInvoice_)
    tmpRec("NeedBillingNote") = Trim(needBillingNote_)
    tmpRec("NeedContract") = Trim(needContract_)
    tmpRec("NeedQuotation") = Trim(needQuotation_)
    tmpRec("RemDesc") = Trim(remDesc_)
    tmpRec("OtherDocDesc") = Trim(otherDocDesc_)
    
    tmpRec("CheqRecType") = Trim(billRecType_)
    tmpRec("BillDay1") = billDay1_
    tmpRec("BillDay2") = billDay2_
    tmpRec("BillDay3") = billDay3_
    tmpRec("BillDay4") = billDay4_
    tmpRec("BillDay5") = billDay5_
    tmpRec("BillWeekCount") = billWeekCount_
    tmpRec("OtherDocDesc") = Trim(cheqRecType_)
    tmpRec("CheqDay1") = cheqDay1_
    tmpRec("CheqDay2") = cheqDay2_
    tmpRec("CheqDay3") = cheqDay3_
    tmpRec("CheqDay4") = cheqDay4_
    tmpRec("CheqDay5") = cheqDay5_
    tmpRec("CheqWeekCount") = cheqWeekCount_
    tmpRec.Update
    tmpRec.Close
    Set tmpRec = Nothing
    isChanged_ = False
End Sub

Public Function Delete() As String
On Error GoTo delErr:
    DBConnExc "Delete tbCustomerBillingInfo Where BillingInfoID='" & id_ & "'"
    Delete = ""
    Exit Function
delErr:
    Delete = Err.Description
End Function

Public Sub SetDataFromRecordset(selRec As ADODB.Recordset)
    id_ = Trim("" & selRec("BillingInfoID"))
    customerID_ = Trim("" & selRec("CustomerID"))
    customerName_ = Trim("" & selRec("CompanyName"))
    contactName_ = Trim("" & selRec("ContactName"))
    telNo_ = Trim("" & selRec("TelNo"))
    faxNo_ = Trim("" & selRec("FaxNo"))
    billingDetail_ = Trim("" & selRec("BillingDetail"))
    needPO_ = Trim("" & selRec("NeedPO"))
    needInvoice_ = Trim("" & selRec("NeedInvoice"))
    needBillingNote_ = Trim("" & selRec("NeedBillingNote"))
    needContract_ = Trim("" & selRec("NeedContract"))
    needQuotation_ = Trim("" & selRec("NeedQuotation"))
    remDesc_ = Trim("" & selRec("RemDesc"))
    cheqDetail_ = Trim("" & selRec("RecCheqDetail"))
    otherDocDesc_ = Trim("" & selRec("OtherDocDesc"))
    
    billRecType_ = Trim("" & selRec("CheqRecType"))
    If IsNull(selRec("BillDay1")) = False Then billDay1_ = selRec("BillDay1")
    If IsNull(selRec("BillDay2")) = False Then billDay2_ = selRec("BillDay2")
    If IsNull(selRec("BillDay3")) = False Then billDay3_ = selRec("BillDay3")
    If IsNull(selRec("BillDay4")) = False Then billDay4_ = selRec("BillDay4")
    If IsNull(selRec("BillDay5")) = False Then billDay5_ = selRec("BillDay5")
    If IsNull(selRec("BillWeekCount")) = False Then billWeekCount_ = selRec("BillWeekCount")
    cheqRecType_ = Trim("" & selRec("OtherDocDesc"))
    If IsNull(selRec("CheqDay1")) = False Then cheqDay1_ = selRec("CheqDay1")
    If IsNull(selRec("CheqDay2")) = False Then cheqDay2_ = selRec("CheqDay2")
    If IsNull(selRec("CheqDay3")) = False Then cheqDay3_ = selRec("CheqDay3")
    If IsNull(selRec("CheqDay4")) = False Then cheqDay4_ = selRec("CheqDay4")
    If IsNull(selRec("CheqDay5")) = False Then cheqDay5_ = selRec("CheqDay5")
    If IsNull(selRec("CheqWeekCount")) = False Then cheqWeekCount_ = selRec("CheqWeekCount")
End Sub

Public Function getBillingDesc() As String
Dim tmpStr As String
    If Trim(billRecType_) = "1" Then
        tmpStr = "วางบิลได้ทุกวัน"
    ElseIf Trim(billRecType_) = "2" Then
        tmpStr = "วางบิลได้เฉพาะวันที่ "
        If billDay1_ > 0 Then
            tmpStr = tmpStr & ", " & billDay1_
        End If
        If billDay2_ > 0 Then
            tmpStr = tmpStr & ", " & billDay2_
        End If
        If billDay3_ > 0 Then
            tmpStr = tmpStr & ", " & billDay3_
        End If
        If billDay4_ > 0 Then
            tmpStr = tmpStr & ", " & billDay4_
        End If
        If billDay5_ > 0 Then
            tmpStr = tmpStr & ", " & billDay5_
        End If
        If Right(Trim(tmpStr), 1) = "," Then tmpStr = Left(Trim(tmpStr), Len(Trim(tmpStr)) - 1)
    ElseIf Trim(billRecType_) = "3" Then
        tmpStr = "วางบิลได้ทุกวัน "
        If billDay1_ > 0 Then
            tmpStr = tmpStr & ", จันทร์"
        End If
        If billDay2_ > 0 Then
            tmpStr = tmpStr & ", อังคาร"
        End If
        If billDay3_ > 0 Then
            tmpStr = tmpStr & ", พุธ"
        End If
        If billDay4_ > 0 Then
            tmpStr = tmpStr & ", พฤหัสบดี"
        End If
        If billDay5_ > 0 Then
            tmpStr = tmpStr & ", ศุกร์"
        End If
        If Right(Trim(tmpStr), 1) = "," Then tmpStr = Left(Trim(tmpStr), Len(Trim(tmpStr)) - 1)
        If billWeekCount_ > 0 Then
            tmpStr = tmpStr & " สัปดาห์ที่ " & billWeekCount_ & " ของเดือน"
        End If
    ElseIf Trim(billRecType_) = "4" Then
        tmpStr = billingDetail_
    End If
    getBillingDesc = tmpStr
End Function

Public Function getCheqDesc() As String
Dim tmpStr As String
    If Trim(cheqRecType_) = "1" Then
        tmpStr = "รับเช็คได้ทุกวัน"
    ElseIf Trim(cheqRecType_) = "2" Then
        tmpStr = "รับเช็คได้เฉพาะวันที่ "
        If cheqDay1_ > 0 Then
            tmpStr = tmpStr & ", " & cheqDay1_
        End If
        If cheqDay2_ > 0 Then
            tmpStr = tmpStr & ", " & cheqDay2_
        End If
        If cheqDay3_ > 0 Then
            tmpStr = tmpStr & ", " & cheqDay3_
        End If
        If cheqDay4_ > 0 Then
            tmpStr = tmpStr & ", " & cheqDay4_
        End If
        If cheqDay5_ > 0 Then
            tmpStr = tmpStr & ", " & cheqDay5_
        End If
        If Right(Trim(tmpStr), 1) = "," Then tmpStr = Left(Trim(tmpStr), Len(Trim(tmpStr)) - 1)
    ElseIf Trim(cheqRecType_) = "3" Then
        tmpStr = "รับเช็คได้ทุกวัน "
        If cheqDay1_ > 0 Then
            tmpStr = tmpStr & ", จันทร์"
        End If
        If cheqDay2_ > 0 Then
            tmpStr = tmpStr & ", อังคาร"
        End If
        If cheqDay3_ > 0 Then
            tmpStr = tmpStr & ", พุธ"
        End If
        If cheqDay4_ > 0 Then
            tmpStr = tmpStr & ", พฤหัสบดี"
        End If
        If cheqDay5_ > 0 Then
            tmpStr = tmpStr & ", ศุกร์"
        End If
        If Right(Trim(tmpStr), 1) = "," Then tmpStr = Left(Trim(tmpStr), Len(Trim(tmpStr)) - 1)
        If cheqWeekCount_ > 0 Then
            tmpStr = tmpStr & " สัปดาห์ที่ " & cheqWeekCount_ & " ของเดือน"
        End If
    ElseIf Trim(cheqRecType_) = "4" Then
        tmpStr = cheqDetail_
    End If
    getCheqDesc = tmpStr
End Function
