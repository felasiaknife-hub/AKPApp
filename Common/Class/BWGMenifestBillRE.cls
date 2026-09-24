VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BWGMenifestBillRE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim id_ As String
Dim TimeTableNo_ As String
Dim DocNo_ As String
Dim menifestNo_ As String
Dim workDate_ As String
Dim iscanceled_ As String
Dim CustomerName_ As String
Dim companyNameTh_ As String
Dim actualQty_ As String
Dim inCollectFee_ As String
Dim salefname_ As String
Dim salelname_ As String
Dim BillingNo_ As String
Dim issuedDate_  As String
Dim printInvID_ As String
Dim reNo_ As String
Dim reDate_ As String
Dim isCancel_ As String
Dim Status_ As String
Dim MStatus_ As String
Dim isChanged_ As Boolean

Public Property Get ID() As String
    ID = id_
End Property
Public Property Let ID(ByVal newVal As String)
    id_ = Trim(newVal)
End Property

Public Property Get TimeTableNo() As String
    TimeTableNo = TimeTableNo_
End Property

Public Property Get docNo() As String
    docNo = DocNo_
End Property

Public Property Let JobID(ByVal newVal As String)
    If Trim(DocNo_) <> Trim(newVal) Then
        DocNo_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get MenifestNo() As String
    MenifestNo = menifestNo_
End Property
Public Property Let MenifestNo(ByVal newVal As String)
    If Trim(menifestNo_) <> Trim(newVal) Then
        menifestNo_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get WorkDate() As String
    WorkDate = workDate_
End Property
Public Property Let WorkDate(ByVal newVal As String)
    If Trim(workDate_) <> Trim(newVal) Then
        workDate_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get IsCanceled() As String
    IsCanceled = iscanceled_
End Property
Public Property Let IsCanceled(ByVal newVal As String)
    If Trim(iscanceled_) <> Trim(newVal) Then
        iscanceled_ = Trim(newVal)
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
Public Property Get ActualQty() As String
    ActualQty = actualQty_
End Property
Public Property Let ActualQty(ByVal newVal As String)
    If Trim(actualQty_) <> Trim(newVal) Then
        actualQty_ = Trim(newVal)
        isChanged_ = True
    End If
End Property
Public Property Get InCollectFee() As String
 InCollectFee = inCollectFee_
End Property
Public Property Let InCollectFee(ByVal newVal As String)
    If Trim(inCollectFee_) <> Trim(newVal) Then
        inCollectFee_ = Trim(newVal)
        isChanged_ = True
    End If
End Property
Public Property Get CompanyNameTH() As String
    CompanyNameTH = companyNameTh_
End Property
Public Property Let CompanyNameTH(ByVal newVal As String)
     If Trim(companyNameTh_) <> Trim(newVal) Then
        companyNameTh_ = Trim(newVal)
        isChanged_ = True
     End If
End Property
Public Property Get SaleFName() As String
    SaleFName = salefname_
End Property
Public Property Let SaleFName(ByVal newVal As String)
    If Trim(salefname_) <> Trim(newVal) Then
        salefname_ = Trim(newVal)
        isChanged_ = True
    End If
End Property
Public Property Get SaleLName() As String
    SaleLName = salelname_
End Property
Public Property Let SaleLName(ByVal newVal As String)
    If Trim(salelname_) <> Trim(newVal) Then
        salelname_ = Trim(newVal)
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
Public Property Get IssuedDate() As String
    IssuedDate = issuedDate_
End Property

Public Property Let IssuedDate(ByVal newVal As String)
    If Trim(issuedDate_) <> Trim(newVal) Then
        issuedDate_ = Trim(newVal)
        isChanged_ = True
    End If
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
Public Property Get RENo() As String
    RENo = reNo_
End Property

Public Property Let RENo(ByVal newVal As String)
    If Trim(reNo_) <> Trim(newVal) Then
        reNo_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get REDate() As String
    REDate = reDate_
End Property
Public Property Let REDate(ByVal newVal As String)
    If Trim(reDate_) <> Trim(newVal) Then
        reDate_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get isCancel() As String
    isCancel = isCancel_
End Property
Public Property Let isCancel(ByVal newVal As String)
    If Trim(isCancel_) <> Trim(newVal) Then
        isCancel_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get Status() As String
    Status = Status_
End Property
Public Property Let Status(ByVal newVal As String)
    If Trim(Status_) <> Trim(newVal) Then
        Status_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get mStatus() As String
    mStatus = MStatus_
End Property
Public Property Let mStatus(ByVal newVal As String)
    If Trim(MStatus_) <> Trim(newVal) Then
        MStatus_ = Trim(newVal)
    End If
End Property

Public Sub SetDataFromRecordset(selRec As ADODB.Recordset)
    id_ = Trim("" & selRec("MenifestID"))
    TimeTableNo_ = Trim("" & selRec("TimeTableNo"))
    DocNo_ = Trim("" & selRec("DocNo"))
    menifestNo_ = Trim("" & selRec("MenifestNo"))
    workDate_ = Trim("" & selRec("WorkDate"))
    iscanceled_ = Trim("" & selRec("isCanceled"))
    CustomerName_ = Trim("" & selRec("CustomerName"))
    companyNameTh_ = Trim("" & selRec("CompanyNameTH"))
    actualQty_ = Trim("" & selRec("ActualQty"))
    inCollectFee_ = Trim("" & selRec("InCollectFee"))
    salefname_ = Trim("" & selRec("SaleFName"))
    salelname_ = Trim("" & selRec("SaleLName"))
    If IsNull(selRec("BillingNo")) = False Then BillingNo_ = selRec("BillingNo")
    If IsNull(selRec("IssuedDate")) = False Then issuedDate_ = selRec("IssuedDate")
    
    If isExist(selRec, "PrintInvID") Then
        If IsNull(selRec("PrintInvID")) = False Then printInvID_ = selRec("PrintInvID")
    End If

    If IsNull(selRec("RENo")) = False Then reNo_ = selRec("RENo")
    If IsNull(selRec("REDate")) = False Then reDate_ = selRec("REDate")
    If IsNull(selRec("isCancel")) = False Then isCancel_ = selRec("isCancel")
    If IsNull(selRec("Status")) = False Then Status_ = selRec("Status")
    If IsNull(selRec("MStatus")) = False Then MStatus_ = selRec("MStatus")
End Sub


