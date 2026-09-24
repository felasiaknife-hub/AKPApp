VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BWGSendSampleHeader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim id_ As String
Dim sendBillNo_ As String
Dim CustomerID_ As String
Dim CustomerName_ As String
Dim customerAddr_ As String
Dim companyID_ As String
Dim sendDate_ As String
Dim issuedStaffID_ As String
Dim issuedStaffName_ As String
Dim sendStaffID_ As String
Dim sendStaffName_ As String
Dim sampleRecStaffID_ As String
Dim sampleRecStaffName_ As String
Dim isTestExtByVender_ As String
Dim labVenderID_ As String
Dim labVenderName_ As String
Dim Remdesc_ As String

Dim reportNo_ As String
Dim sampleSetName_ As String
Dim intSampleTypeID_ As String
Dim intSampleTypeName_ As String
Dim headerType_ As String
Dim templateID_ As String
Dim templateID2_ As String
Dim CostLab_ As String
Dim colSendSampleDet_ As Collection
Dim colLABVender_ As Collection
Dim isApproveQT_ As String
Dim isChanged_ As Boolean

Public Property Get ID() As String
    ID = id_
End Property
Public Property Let ID(ByVal newVal As String)
    id_ = Trim(newVal)
End Property

Public Property Get SendBillNo() As String
    SendBillNo = sendBillNo_
End Property
Public Property Let SendBillNo(ByVal newVal As String)
    If Trim(sendBillNo_) <> Trim(newVal) Then
        sendBillNo_ = Trim(newVal)
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

Public Property Get CustomerAddr() As String
    CustomerAddr = customerAddr_
End Property
Public Property Let CustomerAddr(ByVal newVal As String)
    If Trim(customerAddr_) <> Trim(newVal) Then
        customerAddr_ = Trim(newVal)
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

Public Property Get SendDate() As String
    SendDate = sendDate_
End Property
Public Property Let SendDate(ByVal newVal As String)
    If Trim(sendDate_) <> Trim(newVal) Then
        sendDate_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get IssuedStaffID() As String
    IssuedStaffID = issuedStaffID_
End Property
Public Property Let IssuedStaffID(ByVal newVal As String)
    If Trim(issuedStaffID_) <> Trim(newVal) Then
        issuedStaffID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get IssuedStaffName() As String
    IssuedStaffName = issuedStaffName_
End Property
Public Property Let IssuedStaffName(ByVal newVal As String)
    If Trim(issuedStaffName_) <> Trim(newVal) Then
        issuedStaffName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get SendStaffID() As String
    SendStaffID = sendStaffID_
End Property
Public Property Let SendStaffID(ByVal newVal As String)
    If Trim(sendStaffID_) <> Trim(newVal) Then
        sendStaffID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get SendStaffName() As String
    SendStaffName = sendStaffName_
End Property
Public Property Let SendStaffName(ByVal newVal As String)
    If Trim(sendStaffName_) <> Trim(newVal) Then
        sendStaffName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get SampleRecStaffID() As String
    SampleRecStaffID = sampleRecStaffID_
End Property
Public Property Let SampleRecStaffID(ByVal newVal As String)
    If Trim(sampleRecStaffID_) <> Trim(newVal) Then
        sampleRecStaffID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get SampleRecStaffName() As String
    SampleRecStaffName = sampleRecStaffName_
End Property
Public Property Let SampleRecStaffName(ByVal newVal As String)
    If Trim(sampleRecStaffName_) <> Trim(newVal) Then
        sampleRecStaffName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get IsTestExtByVender() As Boolean
    If Trim(isTestExtByVender_) = "Y" Then
        IsTestExtByVender = True
    Else
        IsTestExtByVender = False
    End If
End Property
Public Property Let IsTestExtByVender(ByVal newVal As Boolean)
Dim tmpStr As String
    If newVal Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(isTestExtByVender_) <> Trim(tmpStr) Then
        isTestExtByVender_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property

Public Property Get LABVenderID() As String
    LABVenderID = labVenderID_
End Property
Public Property Let LABVenderID(ByVal newVal As String)
    If Trim(labVenderID_) <> Trim(newVal) Then
        labVenderID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get LABVenderName() As String
    LABVenderName = labVenderName_
End Property
Public Property Let LABVenderName(ByVal newVal As String)
    If Trim(labVenderName_) <> Trim(newVal) Then
        labVenderName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get Remdesc() As String
    Remdesc = Remdesc_
End Property
Public Property Let Remdesc(ByVal newVal As String)
    If Trim(Remdesc_) <> Trim(newVal) Then
        Remdesc_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get ReportNo() As String
    ReportNo = reportNo_
End Property
Public Property Let ReportNo(ByVal newVal As String)
    If Trim(reportNo_) <> Trim(newVal) Then
        reportNo_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get SampleSetName() As String
    SampleSetName = sampleSetName_
End Property
Public Property Let SampleSetName(ByVal newVal As String)
    If Trim(sampleSetName_) <> Trim(newVal) Then
        sampleSetName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get IntSampleTypeID() As String
    IntSampleTypeID = intSampleTypeID_
End Property
Public Property Let IntSampleTypeID(ByVal newVal As String)
    If Trim(intSampleTypeID_) <> Trim(newVal) Then
        intSampleTypeID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get IntSampleTypeName() As String
    IntSampleTypeName = intSampleTypeName_
End Property
Public Property Let IntSampleTypeName(ByVal newVal As String)
    If Trim(intSampleTypeName_) <> Trim(newVal) Then
        intSampleTypeName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get HeaderType() As String
    HeaderType = headerType_
End Property
Public Property Let HeaderType(ByVal newVal As String)
    If Trim(headerType_) <> Trim(newVal) Then
        headerType_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get TemplateID() As String
    TemplateID = templateID_
End Property
Public Property Let TemplateID(ByVal newVal As String)
    If Trim(templateID_) <> Trim(newVal) Then
        templateID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property
Public Property Get TemplateID2() As String
    TemplateID2 = templateID2_
End Property
Public Property Let TemplateID2(ByVal newVal As String)
    If Trim(templateID2_) <> Trim(newVal) Then
        templateID2_ = Trim(newVal)
        isChanged_ = True
    End If
End Property
Public Property Get CostLab() As String
    CostLab = CostLab_
End Property
Public Property Let CostLab(ByVal newVal As String)
    If Trim(CostLab_) <> Trim(newVal) Then
        CostLab_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get isApproveQT() As String
    isApproveQT = isApproveQT_
End Property
Public Property Let isApproveQT(ByVal newVal As String)
If Trim(isApproveQT_) <> Trim(newVal) Then
        isApproveQT_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get isChanged() As Boolean
    isChanged = isChanged_
End Property
Public Property Let isChanged(ByVal newVal As Boolean)
    isChanged_ = newVal
End Property


Public Property Get SendSampleDetList() As Collection
Dim tmpS As New BWGSearchManager
    If colSendSampleDet_ Is Nothing Then Set colSendSampleDet_ = tmpS.SendSampleDetailSearch("SendSampleID='" & id_ & "'")
    Set SendSampleDetList = colSendSampleDet_
End Property
Public Property Set SendSampleDetList(newVal As Collection)
    Set colSendSampleDet_ = newVal
    isChanged_ = True
End Property

Public Sub RequerySendSampleDetList()
    Set colSendSampleDet_ = Nothing
End Sub

Public Property Get LABVender() As Collection
Dim tmpS As New BWGSearchManager
    If colLABVender_ Is Nothing Then Set colLABVender_ = tmpS.TransportVenderSearch("SendSampleID='" & id_ & "'")
    Set LABVender = colLABVender_
End Property
Public Property Set LABVender(newVal As Collection)
    Set colLABVender_ = newVal
    isChanged_ = True
End Property

Public Sub RequeryLABVender()
    Set colLABVender_ = Nothing
End Sub

Public Sub Save()
On Error GoTo ErrD
Dim tmpRec As ADODB.Recordset
Dim x%
Dim sqlStr As String
    sqlStr = "SELECT * FROM tbSendSampleHeader WHERE SendSampleID='" & id_ & "'"
    Set tmpRec = New ADODB.Recordset
    GetDBConnection
    tmpRec.open sqlStr, DBConn, adOpenDynamic, adLockOptimistic
    If tmpRec.EOF Then
        tmpRec.AddNew
        tmpRec("SendSampleID") = id_
    End If
    tmpRec("CustomerID") = Trim(CustomerID_)
    tmpRec("CompanyID") = Trim(companyID_)
    tmpRec("SendDate") = Trim(sendDate_)
    tmpRec("IssuedStaffID") = Trim(issuedStaffID_)
    tmpRec("SendStaffID") = Trim(sendStaffID_)
    tmpRec("SampleRecStaffID") = Trim(sampleRecStaffID_)
    tmpRec("LabVendorID") = Trim(labVenderID_)
    tmpRec("isTestByExtVendor") = Trim(isTestExtByVender_)
    tmpRec("RemDesc") = Trim(Remdesc_)
    tmpRec("HeaderType") = Trim(headerType_)
    tmpRec("ReportNo") = Trim(reportNo_)
    tmpRec("SampleSetName") = Trim(sampleSetName_)
    tmpRec("IntSampleTypeID") = Trim(intSampleTypeID_)
    tmpRec("TemplateID") = Trim(templateID_)
    tmpRec("TemplateID2") = Trim(templateID2_)
    
    If IsNull(tmpRec("CostLab")) = False Then
        tmpRec("CostLab") = Trim(CostLab_)
    End If
    
    tmpRec("isApproveQT") = Trim(isApproveQT_)
    tmpRec.Update
    tmpRec.Requery
    sendBillNo_ = Trim(tmpRec("SendBillNo"))
    reportNo_ = Trim(tmpRec("ReportNo"))
'    tmpRec.Close
'    With DBConn
        '.Execute "Delete tbSendSampleDetail WHERE SendSampleID='" & id_ & "'"
        DBConnExc "Update tbSendSampleDetail Set MarkAsDelete='Y' WHERE SendSampleID='" & id_ & "'"
        If Not colSendSampleDet_ Is Nothing Then
            For x = 1 To colSendSampleDet_.Count
                If Trim(colSendSampleDet_(x).wastedataID) <> "" Or Trim(colSendSampleDet_(x).SamplePointName) <> "" Then colSendSampleDet_(x).Save
            Next
        End If
        DBConnExc "Delete tbSampleParamReq WHERE SendSampleDetID in (Select SendSampleDetID From tbSendSampleDetail WHERE SendSampleID='" & id_ & "' And MarkAsDelete='Y')"
        DBConnExc "Delete tbSendSampleDetail WHERE SendSampleID='" & id_ & "' And MarkAsDelete='Y'"
'    End With
    
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
    DBConnExc "Delete tbSampleParamReq WHERE  SendSampleDetID in (Select SendSampleDetID from tbSendSampleDetail WHERE SendSampleID='" & id_ & "')"
    DBConnExc "Delete tbSendSampleDetail WHERE  SendSampleID='" & id_ & "'"
    DBConnExc "Delete tbSendSampleHeader WHERE SendSampleID='" & id_ & "'"
    Delete = ""
    Exit Function
delErr:
    Delete = err.Description
End Function

Public Sub SetDataFromRecordset(selRec As ADODB.Recordset)
    id_ = selRec("SendSampleID")
    sendBillNo_ = Trim("" & selRec("SendBillNo"))
    CustomerID_ = Trim("" & selRec("CustomerID"))
    CustomerName_ = Trim("" & selRec("CompanyName"))
    companyID_ = Trim("" & selRec("CompanyID"))
    sendDate_ = Trim("" & selRec("SendDate"))
    issuedStaffID_ = Trim("" & selRec("IssuedStaffID"))
    issuedStaffName_ = Trim("" & selRec("IssuedStaffName"))
    sendStaffID_ = Trim("" & selRec("SendStaffID"))
    'sendStaffName_ = Trim("" & selRec("SendStaffName"))
    sampleRecStaffID_ = Trim("" & selRec("SampleRecStaffID"))
    'sampleRecStaffName_ = Trim("" & selRec("SampleRecStaffName"))
    isTestExtByVender_ = Trim("" & selRec("isTestByExtVendor"))
    labVenderID_ = Trim("" & selRec("LabVendorID"))
    labVenderName_ = Trim("" & selRec("NameTH"))
    Remdesc_ = Trim("" & selRec("RemDesc"))
    reportNo_ = Trim("" & selRec("ReportNo"))
    sampleSetName_ = Trim("" & selRec("SampleSetName"))
    intSampleTypeID_ = Trim("" & selRec("IntSampleTypeID"))
    intSampleTypeName_ = Trim("" & selRec("SampleTypeName"))
    headerType_ = Trim("" & selRec("HeaderType"))
    templateID_ = Trim("" & selRec("TemplateID"))
    templateID2_ = Trim("" & selRec("TemplateID2"))
    CostLab_ = Trim("" & selRec("CostLab"))
    isApproveQT_ = Trim("" & selRec("isApproveQT"))
'   If isExist(selRec, "isApproveQT") Then isApproveQT_ = IIf(IsNull(selRec!isApproveQT), "N", selRec!isApproveQT)
End Sub


