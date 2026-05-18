VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BWGJobDataCarUsed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim id_ As String
Dim jobID_ As String
Dim truckSubTypeID_ As String
Dim truckSubTypeName_ As String
Dim truckTypeID_ As String
Dim truckTypeName_ As String
Dim truckCount_ As Integer
Dim workerPerCar_ As Integer
Dim rowID_ As Integer
Dim carRem_ As String
Dim destinationDesc_ As String
Dim goSameTime_ As String
Dim takePhoto_ As String
Dim carNo_ As Integer
Dim isTrailer_ As String
Dim canBeTrailer_ As String
Dim vendorTripCount_ As Integer
Dim colTimeTable_ As Collection
Dim menifestCount_ As Integer
Dim isChanged_ As Boolean

Public Property Get ID() As String
    ID = id_
End Property
Public Property Let ID(ByVal newVal As String)
    id_ = Trim(newVal)
End Property

Public Property Get JobID() As String
    JobID = jobID_
End Property
Public Property Let JobID(ByVal newVal As String)
    If Trim(jobID_) <> Trim(newVal) Then
        jobID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get TruckSubTypeID() As String
    TruckSubTypeID = truckSubTypeID_
End Property
Public Property Let TruckSubTypeID(ByVal newVal As String)
    If Trim(truckSubTypeID_) <> Trim(newVal) Then
        truckSubTypeID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get TruckSubTypeName() As String
    TruckSubTypeName = truckSubTypeName_
End Property
Public Property Let TruckSubTypeName(ByVal newVal As String)
    If Trim(truckSubTypeName_) <> Trim(newVal) Then
        truckSubTypeName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get TruckTypeID() As String
    TruckTypeID = truckTypeID_
End Property
Public Property Let TruckTypeID(ByVal newVal As String)
    If Trim(truckTypeID_) <> Trim(newVal) Then
        truckTypeID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get TruckTypeName() As String
    TruckTypeName = truckTypeName_
End Property
Public Property Let TruckTypeName(ByVal newVal As String)
    If Trim(truckTypeName_) <> Trim(newVal) Then
        truckTypeName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get TruckCount() As Integer
    TruckCount = truckCount_
End Property
Public Property Let TruckCount(ByVal newVal As Integer)
    If truckCount_ <> newVal Then
        truckCount_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get WorkerPerCar() As Integer
    WorkerPerCar = workerPerCar_
End Property
Public Property Let WorkerPerCar(ByVal newVal As Integer)
    If workerPerCar_ <> newVal Then
        workerPerCar_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get rowID() As Integer
    rowID = rowID_
End Property
Public Property Let rowID(ByVal newVal As Integer)
    If rowID_ <> newVal Then
        rowID_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get Remark() As String
    Remark = carRem_
End Property
Public Property Let Remark(ByVal newVal As String)
    If Trim(carRem_) <> Trim(newVal) Then
        carRem_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get DestinationDesc() As String
    DestinationDesc = destinationDesc_
End Property
Public Property Let DestinationDesc(ByVal newVal As String)
    If Trim(destinationDesc_) <> Trim(newVal) Then
        destinationDesc_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get GoSameTime() As String
    GoSameTime = goSameTime_
End Property
Public Property Let GoSameTime(ByVal newVal As String)
    If Trim(goSameTime_) <> Trim(newVal) Then
        goSameTime_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get TakePhoto() As String
    TakePhoto = takePhoto_
End Property
Public Property Let TakePhoto(ByVal newVal As String)
    If Trim(takePhoto_) <> Trim(newVal) Then
        takePhoto_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get carNo() As Integer
    carNo = carNo_
End Property
Public Property Let carNo(ByVal newVal As Integer)
    If Trim(carNo_) <> Trim(newVal) Then
        carNo_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get isTrailer() As Boolean
    If Trim(isTrailer_) = "Y" Then
        isTrailer = True
    Else
        isTrailer = False
    End If
End Property
Public Property Let isTrailer(ByVal newVal As Boolean)
Dim tmpStr As String
    If newVal Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(isTrailer_) <> Trim(tmpStr) Then
        isTrailer_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property

Public Property Get CanBeTrailer() As Boolean
    If Trim(canBeTrailer_) = "Y" Then
        CanBeTrailer = True
    Else
        CanBeTrailer = False
    End If
End Property
Public Property Let CanBeTrailer(ByVal newVal As Boolean)
Dim tmpStr As String
    If newVal Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(canBeTrailer_) <> Trim(tmpStr) Then
        canBeTrailer_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property

Public Property Get VendorTripCount() As Integer
    VendorTripCount = vendorTripCount_
End Property
Public Property Let VendorTripCount(ByVal newVal As Integer)
    If Trim(vendorTripCount_) <> Trim(newVal) Then
        vendorTripCount_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get MenifestCount() As Integer
    MenifestCount = menifestCount_
End Property
Public Property Let MenifestCount(ByVal newVal As Integer)
    If menifestCount_ <> newVal Then
        menifestCount_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get TimeTables() As Collection
Dim tmpS As New BWGSearchManager
    If colTimeTable_ Is Nothing Then
        Set colTimeTable_ = tmpS.JobDataTimeTableSearch("JobDataCarID='" & id_ & "'")
    End If
    Set TimeTables = colTimeTable_
End Property
Public Property Set TimeTables(ByVal newVal As Collection)
    Set colTimeTable_ = Trim(newVal)
    isChanged_ = True
End Property

Public Sub RequeryTimeTable()
    Set colTimeTable_ = Nothing
End Sub

Public Property Get isChanged() As Boolean
    isChanged = isChanged_
End Property
Public Property Let isChanged(ByVal newVal As Boolean)
    isChanged_ = newVal
End Property

Public Sub Save(Optional WorkDate As String = "")
On Error GoTo ErrD
Dim tmpRec As ADODB.Recordset
Dim x%
    Set tmpRec = New ADODB.Recordset
    tmpRec.Open "Select * from tbJobDataCarUsed Where JobDataCarID='" & id_ & "'", DBManager.GetDBConnection, adOpenDynamic, adLockOptimistic
    If tmpRec.EOF Then
        tmpRec.AddNew
        tmpRec("JobDataCarID") = id_
    End If
    tmpRec("JobID") = Trim(jobID_)
    tmpRec("TruckSubTypeID") = Trim(truckSubTypeID_)
    tmpRec("TruckCount") = Trim(truckCount_)
    tmpRec("WorkerPerCar") = Trim(workerPerCar_)
    tmpRec("RowID") = Trim(rowID_)
    tmpRec("CarRem") = Trim(carRem_)
    tmpRec("DestinationDesc") = Trim(destinationDesc_)
    tmpRec("GoSameTime") = Trim(goSameTime_)
    tmpRec("TakePhoto") = Trim(takePhoto_)
    tmpRec("CarNo") = Trim(carNo_)
    tmpRec("MarkForDelete") = "N"
    tmpRec("isTrailer") = Trim(isTrailer_)
    tmpRec("VendorTripCount") = vendorTripCount_
    tmpRec.Update
    tmpRec.Close
    With DBManager.GetDBConnection
        .Execute "Update tbJobDataTimeTable Set MarkForDelete='Y' Where JobDataCarID='" & id_ & "'"
        If Not colTimeTable_ Is Nothing Then
            For x = 1 To colTimeTable_.Count
                colTimeTable_(x).WorkerCount = workerPerCar_
                colTimeTable_(x).Save WorkDate
            Next
        End If
        .Execute "Delete tbJobDataTimeTable Where MarkForDelete='Y' And JobDataCarID='" & id_ & "'"
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
    DBManager.GetDBConnection.Execute "Delete tbJobDataCarUsed Where JobDataCarID='" & id_ & "'"
    Delete = ""
    Exit Function
delErr:
    Delete = Err.Description
End Function

Public Sub SetDataFromRecordset(selRec As ADODB.Recordset)
    id_ = Trim("" & selRec("JobDataCarID"))
    jobID_ = Trim("" & selRec("JobID"))
    truckSubTypeID_ = Trim("" & selRec("TruckSubTypeID"))
    truckSubTypeName_ = Trim("" & selRec("SubTypeDesc"))
    truckTypeID_ = Trim("" & selRec("TruckTypeID"))
    truckTypeName_ = Trim("" & selRec("TruckTypeDesc"))
    truckCount_ = selRec("TruckCount")
    workerPerCar_ = selRec("WorkerPerCar")
    rowID_ = selRec("RowID")
    carRem_ = Trim("" & selRec("CarRem"))
    destinationDesc_ = Trim("" & selRec("DestinationDesc"))
    goSameTime_ = Trim("" & selRec("GoSameTime"))
    takePhoto_ = Trim("" & selRec("TakePhoto"))
    isTrailer_ = Trim("" & selRec("isTrailer"))
    canBeTrailer_ = Trim("" & selRec("CanBeTrailer"))
    vendorTripCount_ = selRec("VendorTripCount")
    If IsNull(selRec("CarNo")) = False Then carNo_ = selRec("CarNo")
    If IsNull(selRec("MenifestCount")) = False Then menifestCount_ = selRec("MenifestCount")
End Sub

Public Sub ClearTimeTable()
    Set colTimeTable_ = Nothing
    DBManager.GetDBConnection.Execute "Delete tbJobDataTimeTable Where JobDataCarID='" & id_ & "'"
End Sub
