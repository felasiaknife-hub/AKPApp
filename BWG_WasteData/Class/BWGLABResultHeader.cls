VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BWGLABResultHeader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim id_ As String
Dim reportNo_ As String
Dim reportDate_ As String
Dim samplingDate_ As String
Dim receiveDate_ As String
Dim analyticalDateStart_ As String
Dim analyticalDateEnd_ As String
Dim analysisNo_ As String
Dim customerID_ As String
Dim customerName_ As String
Dim customerAddr_ As String
Dim sampleDesc_ As String
Dim wastDataID_ As String
Dim wasteName_ As String
Dim wasteNo_ As String
Dim templateID_ As String
Dim templateNo_ As String
Dim sendSampleDetID_ As String
Dim companyID_ As String
Dim isCancel_ As String
Dim isApprove4Print_ As String
Dim isFinalApproved_ As String
Dim colLabResultDet_ As Collection
Dim isChanged_ As Boolean

Public Property Get ID() As String
    ID = id_
End Property
Public Property Let ID(ByVal newVal As String)
    id_ = Trim(newVal)
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

Public Property Get ReportDate() As String
    ReportDate = reportDate_
End Property
Public Property Let ReportDate(ByVal newVal As String)
    If Trim(reportDate_) <> Trim(newVal) Then
        reportDate_ = Trim(newVal)
        isChanged_ = True
    End If
End Property


Public Property Get SamplingDate() As String
    SamplingDate = samplingDate_
End Property
Public Property Let SamplingDate(ByVal newVal As String)
    If Trim(samplingDate_) <> Trim(newVal) Then
        samplingDate_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get ReceivedDate() As String
    ReceivedDate = receiveDate_
End Property
Public Property Let ReceivedDate(ByVal newVal As String)
    If Trim(receiveDate_) <> Trim(newVal) Then
        receiveDate_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get AnalyticalDateStart() As String
    AnalyticalDateStart = analyticalDateStart_
End Property
Public Property Let AnalyticalDateStart(ByVal newVal As String)
    If Trim(analyticalDateStart_) <> Trim(newVal) Then
        analyticalDateStart_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get AnalyticalDateEnd() As String
    AnalyticalDateEnd = analyticalDateEnd_
End Property
Public Property Let AnalyticalDateEnd(ByVal newVal As String)
    If Trim(analyticalDateEnd_) <> Trim(newVal) Then
        analyticalDateEnd_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get AnalysisNo() As String
    AnalysisNo = analysisNo_
End Property
Public Property Let AnalysisNo(ByVal newVal As String)
    If Trim(analysisNo_) <> Trim(newVal) Then
        analysisNo_ = Trim(newVal)
        isChanged_ = True
    End If
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

Public Property Get CustomerAddr() As String
    CustomerAddr = customerAddr_
End Property
Public Property Let CustomerAddr(ByVal newVal As String)
    If Trim(customerAddr_) <> Trim(newVal) Then
        customerAddr_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get SampleDesc() As String
    SampleDesc = sampleDesc_
End Property
Public Property Let SampleDesc(ByVal newVal As String)
    If Trim(sampleDesc_) <> Trim(newVal) Then
        sampleDesc_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get WasteDataID() As String
    WasteDataID = wastDataID_
End Property
Public Property Let WasteDataID(ByVal newVal As String)
    If Trim(wastDataID_) <> Trim(newVal) Then
        wastDataID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get WasteName() As String
    WasteName = wasteName_
End Property
Public Property Let WasteName(ByVal newVal As String)
    If Trim(wasteName_) <> Trim(newVal) Then
        wasteName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get WasteNo() As String
    WasteNo = wasteNo_
End Property
Public Property Let WasteNo(ByVal newVal As String)
    If Trim(wasteNo_) <> Trim(newVal) Then
        wasteNo_ = Trim(newVal)
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

Public Property Get TemplateNo() As String
    TemplateNo = templateNo_
End Property
Public Property Let TemplateNo(ByVal newVal As String)
    If Trim(templateNo_) <> Trim(newVal) Then
        templateNo_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get SendSampleDetID() As String
    SendSampleDetID = sendSampleDetID_
End Property
Public Property Let SendSampleDetID(ByVal newVal As String)
    If Trim(sendSampleDetID_) <> Trim(newVal) Then
        sendSampleDetID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get companyID() As String
    companyID = companyID_
End Property
Public Property Let companyID(ByVal newVal As String)
    If Trim(companyID_) <> Trim(newVal) Then
        companyID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get isCancel() As Boolean
    isCancel = isCancel_
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

Public Property Get isApprove4Print() As Boolean
    If Trim(isApprove4Print_) = "Y" Then
        isApprove4Print = True
    Else
        isApprove4Print = False
    End If
End Property
Public Property Let isApprove4Print(ByVal newVal As Boolean)
Dim tmpStr As String
    If newVal Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(isApprove4Print_) <> Trim(tmpStr) Then
        isApprove4Print_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property

Public Property Get isFinalApproved() As Boolean
    If Trim(isFinalApproved_) = "Y" Then
        isFinalApproved = True
    Else
        isFinalApproved = False
    End If
End Property
Public Property Let isFinalApproved(ByVal newVal As Boolean)
Dim tmpStr As String
    If newVal Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(isFinalApproved_) <> Trim(tmpStr) Then
        isFinalApproved_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property

Public Property Get LabResultDetails() As Collection
Dim tmpS As New BWGSearchManager
    If colLabResultDet_ Is Nothing Then Set colLabResultDet_ = tmpS.LabResultDetailSearch("LabResultID='" & id_ & "'")
    Set LabResultDetails = colLabResultDet_
End Property
Public Property Set LabResultDetails(ByVal newVal As Collection)
    Set colLabResultDet_ = newVal
    isChanged_ = True
End Property
Public Sub RequeryLabResultDetail()
    Set colLabResultDet_ = Nothing
End Sub

Public Property Get isChanged() As Boolean
    isChanged = isChanged_
End Property
Public Property Let isChanged(ByVal newVal As Boolean)
    isChanged_ = newVal
End Property

Public Sub Save()
On Error GoTo ErrD
Dim tmpRec As ADODB.Recordset
Dim x%
    Set tmpRec = New ADODB.Recordset
    tmpRec.Open "Select * from tbLABResultHeader Where LABResultID='" & id_ & "'", DBManager.GetDBConnection, adOpenDynamic, adLockOptimistic
    If tmpRec.EOF Then
        tmpRec.AddNew
        tmpRec("LABResultID") = id_
    End If
    tmpRec("ReportDate") = Trim(reportDate_)
    tmpRec("SamplingDate") = Trim(samplingDate_)
    tmpRec("ReceivedDate") = Trim(receiveDate_)
    tmpRec("AnalyticalDateStart") = Trim(analyticalDateStart_)
    tmpRec("AnalyticalDateEnd") = Trim(analyticalDateEnd_)
    tmpRec("CustomerID") = Trim(customerID_)
    tmpRec("SampleDesc") = Trim(sampleDesc_)
    tmpRec("WasteDataID") = Trim(wastDataID_)
    tmpRec("TemplateID") = Trim(templateID_)
    tmpRec("SendSampleDetID") = Trim(sendSampleDetID_)
    tmpRec("CompanyID") = Trim(companyID_)
    tmpRec("AnalysisNo") = Trim(analysisNo_)
    tmpRec("ReportNo") = Trim(reportNo_)
    If Trim(isCancel_) = "" Then isCancel_ = "N"
    tmpRec("isCancel") = Trim(isCancel_)
    If Trim(isApprove4Print_) = "" Then isApprove4Print_ = "N"
    tmpRec("isApprove4Print") = Trim(isApprove4Print_)
    If Trim(isFinalApproved_) = "" Then isFinalApproved_ = "N"
    tmpRec("isFinalApproved") = Trim(isFinalApproved_)
    tmpRec.Update
    tmpRec.Requery
    analysisNo_ = Trim("" & tmpRec("AnalysisNo"))
    reportNo_ = Trim("" & tmpRec("ReportNo"))
    tmpRec.Close
    Set tmpRec = Nothing
    DBManager.GetDBConnection.Execute "Delete tbLABResultDetail Where LABResultID='" & id_ & "'"
    For x = 1 To colLabResultDet_.Count
        colLabResultDet_(x).Save
    Next
    isChanged_ = False
    Exit Sub
ErrD:
    MsgBox Err.Description & vbCrLf & "กรุณาลองใหม่หรือติดต่อผู้พัฒนาระบบ", vbCritical, "Error"
    Exit Sub
End Sub

Public Function Delete() As String
On Error GoTo delErr:
    With DBManager.GetDBConnection
        .Execute "Delete tbLABResultDetail Where LABResultID='" & id_ & "'"
        .Execute "Delete tbLABResultHeader Where LABResultID='" & id_ & "'"
    End With
    Delete = ""
    Exit Function
delErr:
    Delete = Err.Description
End Function

Public Sub SetDataFromRecordset(selRec As ADODB.Recordset)
    id_ = selRec("LABResultID")
    reportNo_ = Trim("" & selRec("ReportNo"))
    reportDate_ = Trim("" & selRec("ReportDate"))
    samplingDate_ = Trim("" & selRec("SamplingDate"))
    receiveDate_ = Trim("" & selRec("ReceivedDate"))
    analyticalDateStart_ = Trim("" & selRec("AnalyticalDateStart"))
    analyticalDateEnd_ = Trim("" & selRec("AnalyticalDateEnd"))
    analysisNo_ = Trim("" & selRec("AnalysisNo"))
    customerID_ = Trim("" & selRec("CustomerID"))
    customerName_ = Trim("" & selRec("CompanyName"))
    sampleDesc_ = Trim("" & selRec("SampleDesc"))
    wastDataID_ = Trim("" & selRec("WasteDataID"))
    wasteName_ = Trim("" & selRec("WasteName"))
    wasteNo_ = Trim("" & selRec("WasteNo"))
    templateID_ = Trim("" & selRec("TemplateID"))
    templateNo_ = Trim("" & selRec("TemplateNo"))
    sendSampleDetID_ = Trim("" & selRec("SendSampleDetID"))
    companyID_ = Trim("" & selRec("CompanyID"))
    isCancel_ = Trim("" & selRec("isCancel"))
    isApprove4Print_ = Trim("" & selRec("isApprove4Print"))
    isFinalApproved_ = Trim("" & selRec("isFinalApproved"))
    customerAddr_ = Trim("" & selRec("Addr1") & " " & selRec("TumbolName") & " " & selRec("AumphurName") & " " & selRec("ProvinceName"))
End Sub

Public Sub GetLabTemplete()
Dim tmpS As New BWGSearchManager
Dim tmpCol As Collection, y%, itemCount
Dim x%, tmpPlate As BWGLABTemplateHeader, tmpDet As BWGLABResultDetail, tmpPlateDet As BWGLABTemplateDetail
    Set tmpCol = tmpS.LABTemplateHeaderSearch
    Set colLabResultDet_ = Nothing
    Set colLabResultDet_ = New Collection
    itemCount = 1
    For x = 1 To tmpCol.Count
        Set tmpPlate = tmpCol(x)
        For y = 1 To tmpPlate.LABTemplateDet.Count
            Set tmpPlateDet = tmpPlate.LABTemplateDet(y)
            Set tmpDet = New BWGLABResultDetail
            tmpDet.ID = GetGUID
            With tmpDet
                .CommentDesc = tmpPlateDet.CommentDesc
                .CompareType = tmpPlateDet.CompareType
                .hasStandard = tmpPlateDet.hasStandard
                .ItemNo = itemCount
                itemCount = itemCount + 1
                .LABResultID = id_
                .LabUnitID = tmpPlateDet.LabUnitID
                .LabUnitName = tmpPlateDet.LabUnitName
                .MethodID = tmpPlateDet.MethodID
                .MethodName = tmpPlateDet.MethodName
                .ParameterID = tmpPlateDet.ParameterID
                .ParameterName = tmpPlateDet.ParameterName
                .ResultType = tmpPlateDet.ResultType
                .StandardValue = tmpPlateDet.StandardValue
                .TestTypeID = tmpPlate.TestTypeID
                .TestTypeName = tmpPlate.TestTypeName
            End With
            colLabResultDet_.Add tmpDet, "'" & tmpDet.ID & "'"
            Set tmpDet = Nothing
            Set tmpPlateDet = Nothing
        Next
    Next
End Sub

Public Sub Approve4Print()
    DBManager.GetDBConnection.Execute "Update tbLABResultHeader Set isApprove4Print='Y' Where LABResultID='" & id_ & "'"
    isApprove4Print_ = "Y"
End Sub

Public Sub FinalApprove()
    DBManager.GetDBConnection.Execute "Update tbLABResultHeader Set isFinalApproved='Y' Where LABResultID='" & id_ & "'"
    isFinalApproved_ = "Y"
End Sub
