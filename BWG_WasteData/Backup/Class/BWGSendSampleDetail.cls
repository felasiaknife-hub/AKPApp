VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BWGSendSampleDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim id_ As String
Dim sendSampleID_ As String
Dim rowPos_ As Long
Dim wasteDataID_ As String
Dim wasteNo_ As String
Dim sendBillNo_ As String
Dim sendDate_ As String
Dim wasteDataName_ As String
Dim SampleType_ As String
Dim paramName_ As String
Dim SampleQty_ As Double
Dim PackageType_ As String
Dim samplePointName_ As String
Dim colSampleParamReq_ As Collection
Dim colWasteDataCR_ As Collection
Dim ColorA_  As String
Dim ColorB_ As String
Dim ColorC_ As String
Dim ColorD_ As String
Dim BetNo_ As String
Dim isLab_ As String
Dim Remark_ As String
Dim IsLabEnv_ As String
Dim TRemark_ As String
Dim isPretreat_ As String
Dim PtRemark_ As String
Dim isChanged_ As Boolean



Public Property Get ID() As String
    ID = id_
End Property
Public Property Let ID(ByVal newval As String)
    id_ = Trim(newval)
End Property

Public Property Get SendSampleID() As String
    SendSampleID = sendSampleID_
End Property
Public Property Let SendSampleID(ByVal newval As String)
    If Trim(sendSampleID_) <> Trim(newval) Then
        sendSampleID_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get RowPos() As Long
    RowPos = rowPos_
End Property
Public Property Let RowPos(ByVal newval As Long)
    If rowPos_ <> newval Then
        rowPos_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get wastedataID() As String
    wastedataID = wasteDataID_
End Property

Public Property Let wastedataID(ByVal newval As String)
    If Trim(wasteDataID_) <> Trim(newval) Then
        wasteDataID_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get WasteNo() As String
    WasteNo = wasteNo_
End Property

Public Property Let WasteNo(ByVal newval As String)
    If Trim(wasteNo_) <> Trim(newval) Then
        wasteNo_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get SendBillNo() As String
    SendBillNo = sendBillNo_
End Property

Public Property Let SendBillNo(ByVal newval As String)
    If Trim(sendBillNo_) <> Trim(newval) Then
        sendBillNo_ = Trim(newval)
    End If
End Property

Public Property Get SendDate() As String
    SendDate = sendDate_
End Property

Public Property Let SendDate(ByVal newval As String)
    If Trim(sendDate_) <> Trim(newval) Then
        sendDate_ = Trim(newval)
    End If
End Property

Public Property Get WasteDataName() As String
    WasteDataName = wasteDataName_
End Property

Public Property Let WasteDataName(ByVal newval As String)
    If Trim(wasteDataName_) <> Trim(newval) Then
        wasteDataName_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get SampleType() As String
    SampleType = SampleType_
End Property

Public Property Let SampleType(ByVal newval As String)
    If Trim(SampleType_) <> Trim(newval) Then
        SampleType_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get ParamName() As String
    ParamName = paramName_
End Property

Public Property Let ParamName(ByVal newval As String)
    If Trim(paramName_) <> Trim(newval) Then
        paramName_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get SampleQty() As Double
    SampleQty = SampleQty_
End Property
Public Property Let SampleQty(ByVal newval As Double)
    If SampleQty_ <> newval Then
        SampleQty_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get PackageType() As String
    PackageType = PackageType_
End Property

Public Property Let PackageType(ByVal newval As String)
    If Trim(PackageType_) <> Trim(newval) Then
        PackageType_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get SamplePointName() As String
    SamplePointName = samplePointName_
End Property

Public Property Let SamplePointName(ByVal newval As String)
    If Trim(samplePointName_) <> Trim(newval) Then
        samplePointName_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get SampleParamReq() As Collection
Dim tmpS As New BWGSearchManager
    If colSampleParamReq_ Is Nothing Then Set colSampleParamReq_ = tmpS.ParameterLABSearch("ParameterID in (Select ParameterID from tbSampleParamReq Where SendSampleDetID='" & id_ & "')")
    Set SampleParamReq = colSampleParamReq_
End Property
Public Property Set SampleParamReq(ByVal newval As Collection)
    Set colSampleParamReq_ = newval
    isChanged_ = True
End Property
Public Sub RequerySampleParamReq()
    Set colSampleParamReq_ = Nothing
End Sub

Public Property Get WasteDataCR() As Collection
Dim tmpS As New BWGSearchManager
    If colWasteDataCR_ Is Nothing Then Set colWasteDataCR_ = tmpS.WasteDataCRSearch("WasteDataID ='" & id_ & "'")
    Set WasteDataCR = colWasteDataCR_
End Property
Public Property Set WasteDataCR(ByVal newval As Collection)
    Set colWasteDataCR_ = newval
    isChanged_ = True
End Property
Public Sub RequeryWasteDataCR()
    Set colWasteDataCR_ = Nothing
End Sub


Public Property Get ColorA() As String
    ColorA = ColorA_
End Property
Public Property Let ColorA(ByVal newval As String)
    ColorA_ = Trim(newval)
End Property

Public Property Get ColorB() As String
    ColorB = ColorB_
End Property
Public Property Let ColorB(ByVal newval As String)
    ColorB_ = Trim(newval)
End Property

Public Property Get ColorC() As String
    ColorC = ColorC_
End Property
Public Property Let ColorC(ByVal newval As String)
    ColorC_ = Trim(newval)
End Property

Public Property Get ColorD() As String
    ColorD = ColorD_
End Property
Public Property Let ColorD(ByVal newval As String)
    ColorD_ = Trim(newval)
End Property

Public Property Get BetNo() As String
    BetNo = BetNo_
End Property
Public Property Let BetNo(ByVal newval As String)
    BetNo_ = Trim(newval)
End Property

Public Property Get isLab() As String
    isLab = isLab_
End Property
Public Property Let isLab(ByVal newval As String)
    isLab_ = Trim(newval)
End Property

Public Property Get Remark() As String
    Remark = Remark_
End Property
Public Property Let Remark(ByVal newval As String)
    Remark_ = Trim(newval)
End Property

Public Property Get IsLabEnv() As String
    IsLabEnv = IsLabEnv_
End Property
Public Property Let IsLabEnv(ByVal newval As String)
    IsLabEnv_ = Trim(newval)
End Property

Public Property Get TRemark() As String
    TRemark = TRemark_
End Property
Public Property Let TRemark(ByVal newval As String)
    TRemark_ = Trim(newval)
End Property

Public Property Get isPretreat() As String
    isPretreat = isPretreat_
End Property
Public Property Let isPretreat(ByVal newval As String)
    isPretreat_ = Trim(newval)
End Property

Public Property Get PtRemark() As String
    PtRemark = PtRemark_
End Property
Public Property Let PtRemark(ByVal newval As String)
    PtRemark_ = Trim(newval)
End Property

Public Property Get isChanged() As Boolean
    isChanged = isChanged_
End Property
Public Property Let isChanged(ByVal newval As Boolean)
    isChanged_ = newval
End Property

Public Sub Save()
On Error GoTo ErrD
Dim tmpRec As ADODB.Recordset
Dim x%
    Set tmpRec = New ADODB.Recordset
    tmpRec.Open "Select * from tbSendSampleDetail Where SendSampleDetID='" & id_ & "'", DBManager.GetDBConnection, adOpenDynamic, adLockOptimistic
    If tmpRec.EOF Then
        tmpRec.AddNew
        tmpRec("SendSampleDetID") = id_
    End If
    tmpRec("SendSampleID") = Trim(sendSampleID_)
    tmpRec("RowPos") = Trim(rowPos_)
    tmpRec("WasteDataID") = Trim(wasteDataID_)
    tmpRec("SampleType") = Trim(SampleType_)
    tmpRec("SampleQty") = Trim(SampleQty_)
    tmpRec("PackageType") = Trim(PackageType_)
    tmpRec("SamplePointName") = Trim(samplePointName_)
    tmpRec("MarkAsDelete") = "N"
    tmpRec.Update
    tmpRec.Close
    
    With DBManager.GetDBConnection
        .Execute "Delete tbSampleParamReq Where SendSampleDetID='" & id_ & "'"
        If Not colSampleParamReq_ Is Nothing Then
            For x = 1 To colSampleParamReq_.Count
                .Execute "Insert into tbSampleParamReq (ParameterID,SendSampleDetID) Values ('" & colSampleParamReq_(x).ID & "','" & id_ & "')"
            Next
        End If
    End With

    Set tmpRec = Nothing
    isChanged_ = False
    Exit Sub
ErrD:
    MsgBox Err.Description & vbCrLf & "กรุณาลองใหม่หรือติดต่อผู้พัฒนาระบบ", vbCritical, "Error"
    Exit Sub
End Sub

Public Function Delete() As String
On Error GoTo delErr:
    With DBManager.GetDBConnection
        .Execute "Delete tbSampleParamReq Where SendSampleDetID='" & id_ & "'"
        .Execute "Delete tbSendSampleDetail Where SendSampleDetID='" & id_ & "'"
    End With
    Delete = ""
    Exit Function
delErr:
    Delete = Err.Description
End Function

Public Sub SetDataFromRecordset(selRec As ADODB.Recordset)
    id_ = selRec("SendSampleDetID")
    sendSampleID_ = Trim("" & selRec("SendSampleID"))
    If IsNull(selRec("RowPos")) = False Then rowPos_ = selRec("RowPos")
    wasteDataID_ = Trim("" & selRec("WasteDataID"))
    wasteNo_ = Trim("" & selRec("WasteNo"))
    wasteDataName_ = Trim("" & selRec("WasteName"))
    sendBillNo_ = Trim("" & selRec("SendBillNo"))
    sendDate_ = Trim("" & selRec("SendDate"))
    SampleType_ = Trim("" & selRec("SampleType"))
    If IsNull(selRec("SampleQty")) = False Then SampleQty_ = selRec("SampleQty")
    PackageType_ = Trim("" & selRec("PackageType"))
    samplePointName_ = Trim("" & selRec("SamplePointName"))
    ColorA_ = Trim("" & selRec("ColorA"))
    ColorB_ = Trim("" & selRec("ColorB"))
    ColorC_ = Trim("" & selRec("ColorC"))
    ColorD_ = Trim("" & selRec("ColorD"))
    BetNo_ = Trim("" & selRec("BetNo"))
    isLab_ = Trim("" & selRec("isLab"))
    Remark_ = Trim("" & selRec("Remark"))
    IsLabEnv_ = Trim("" & selRec("IsLabEnv"))
    TRemark_ = Trim("" & selRec("TRemark"))
    isPretreat_ = Trim("" & selRec("isPretreat"))
    PtRemark_ = Trim("" & selRec("PtRemark"))
End Sub

