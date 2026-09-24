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
Dim wasteDataName_ As String
Dim sampleType_ As String
Dim paramName_ As String
Dim sampleQty_ As Double
Dim packageType_ As String
Dim samplePointName_ As String
Dim colSampleParamReq_ As Collection
Dim colWasteDataCR_ As Collection
Dim isChanged_ As Boolean

Public Property Get ID() As String
    ID = id_
End Property
Public Property Let ID(ByVal newVal As String)
    id_ = Trim(newVal)
End Property

Public Property Get SendSampleID() As String
    SendSampleID = sendSampleID_
End Property
Public Property Let SendSampleID(ByVal newVal As String)
    If Trim(sendSampleID_) <> Trim(newVal) Then
        sendSampleID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get RowPos() As Long
    RowPos = rowPos_
End Property
Public Property Let RowPos(ByVal newVal As Long)
    If rowPos_ <> newVal Then
        rowPos_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get wastedataID() As String
    wastedataID = wasteDataID_
End Property

Public Property Let wastedataID(ByVal newVal As String)
    If Trim(wasteDataID_) <> Trim(newVal) Then
        wasteDataID_ = Trim(newVal)
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


Public Property Get WasteDataName() As String
    WasteDataName = wasteDataName_
End Property

Public Property Let WasteDataName(ByVal newVal As String)
    If Trim(wasteDataName_) <> Trim(newVal) Then
        wasteDataName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get SampleType() As String
    SampleType = sampleType_
End Property

Public Property Let SampleType(ByVal newVal As String)
    If Trim(sampleType_) <> Trim(newVal) Then
        sampleType_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get ParamName() As String
    ParamName = paramName_
End Property

Public Property Let ParamName(ByVal newVal As String)
    If Trim(paramName_) <> Trim(newVal) Then
        paramName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get SampleQty() As Double
    SampleQty = sampleQty_
End Property
Public Property Let SampleQty(ByVal newVal As Double)
    If sampleQty_ <> newVal Then
        sampleQty_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get PackageType() As String
    PackageType = packageType_
End Property

Public Property Let PackageType(ByVal newVal As String)
    If Trim(packageType_) <> Trim(newVal) Then
        packageType_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get SamplePointName() As String
    SamplePointName = samplePointName_
End Property

Public Property Let SamplePointName(ByVal newVal As String)
    If Trim(samplePointName_) <> Trim(newVal) Then
        samplePointName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get SampleParamReq() As Collection
Dim tmpS As New BWGSearchManager
    If colSampleParamReq_ Is Nothing Then Set colSampleParamReq_ = tmpS.ParameterLABSearch("ParameterID in (Select ParameterID from tbSampleParamReq WHERE SendSampleDetID='" & id_ & "')")
    Set SampleParamReq = colSampleParamReq_
End Property
Public Property Set SampleParamReq(ByVal newVal As Collection)
    Set colSampleParamReq_ = newVal
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
Public Property Set WasteDataCR(ByVal newVal As Collection)
    Set colWasteDataCR_ = newVal
    isChanged_ = True
End Property
Public Sub RequeryWasteDataCR()
    Set colWasteDataCR_ = Nothing
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
Dim sqlStr As String
    sqlStr = "SELECT * FROM tbSendSampleDetail WHERE SendSampleDetID='" & id_ & "'"
    Set tmpRec = New ADODB.Recordset
    GetDBConnection
    tmpRec.open sqlStr, DBConn, adOpenDynamic, adLockOptimistic
    If tmpRec.EOF Then
        tmpRec.AddNew
        tmpRec("SendSampleDetID") = id_
    End If
    tmpRec("SendSampleID") = Trim(sendSampleID_)
    tmpRec("RowPos") = Trim(rowPos_)
    tmpRec("WasteDataID") = Trim(wasteDataID_)
    tmpRec("SampleType") = Trim(sampleType_)
    tmpRec("SampleQty") = Trim(sampleQty_)
    tmpRec("PackageType") = Trim(packageType_)
    tmpRec("SamplePointName") = Trim(samplePointName_)
    tmpRec("MarkAsDelete") = "N"
    tmpRec.Update
'    tmpRec.Close
    
'    With DBConn
        DBConnExc "Delete tbSampleParamReq WHERE SendSampleDetID='" & id_ & "'"
        If Not colSampleParamReq_ Is Nothing Then
            For x = 1 To colSampleParamReq_.Count
                DBConnExc "Insert into tbSampleParamReq (ParameterID,SendSampleDetID) Values ('" & colSampleParamReq_(x).ID & "','" & id_ & "')"
            Next
        End If
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
    DBConnExc "Delete tbSampleParamReq WHERE SendSampleDetID='" & id_ & "'"
    DBConnExc "Delete tbSendSampleDetail WHERE SendSampleDetID='" & id_ & "'"
    Delete = ""
    Exit Function
delErr:
    Delete = err.Description
End Function

Public Sub SetDataFromRecordset(selRec As ADODB.Recordset)
    id_ = selRec("SendSampleDetID")
    sendSampleID_ = Trim("" & selRec("SendSampleID"))
    If IsNull(selRec("RowPos")) = False Then rowPos_ = selRec("RowPos")
    wasteDataID_ = Trim("" & selRec("WasteDataID"))
    wasteNo_ = Trim("" & selRec("WasteNo"))
    wasteDataName_ = Trim("" & selRec("WasteName"))
    sampleType_ = Trim("" & selRec("SampleType"))
    If IsNull(selRec("SampleQty")) = False Then sampleQty_ = selRec("SampleQty")
    packageType_ = Trim("" & selRec("PackageType"))
    samplePointName_ = Trim("" & selRec("SamplePointName"))
End Sub

