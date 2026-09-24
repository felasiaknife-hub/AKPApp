VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BWGBoxTransactionD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim id_ As String
Dim JobDataCarID_ As String
Dim TripNo_ As Integer
Dim CarRegisID_ As String
Dim VendorTruckID_ As String
Dim DriverName_ As String
Dim TimeTableNo_ As String
Dim BoxOwnersCode_ As String
Dim BoxType_ As String
Dim BoxNo_ As String
Dim BoxTransactionStatus_ As String
Dim BoxID_ As String
Dim isChanged_ As Boolean

Public Property Get ID() As String
    ID = id_
End Property
Public Property Let ID(ByVal newVal As String)
    id_ = Trim(newVal)
End Property

Public Property Get JobDataCarID() As String
    JobDataCarID = JobDataCarID_
End Property
Public Property Let JobDataCarID(ByVal newVal As String)
    JobDataCarID_ = Trim(newVal)
End Property

Public Property Get VendorTruckID() As String
    VendorTruckID = VendorTruckID_
End Property
Public Property Let VendorTruckID(ByVal newVal As String)
    VendorTruckID_ = Trim(newVal)
End Property

Public Property Get TripNo() As Integer
    TripNo = TripNo_
End Property
Public Property Let TripNo(ByVal newVal As Integer)
    TripNo_ = Trim(newVal)
End Property

Public Property Get CarRegisID() As String
    CarRegisID = CarRegisID_
End Property
Public Property Let CarRegisID(ByVal newVal As String)
    CarRegisID_ = Trim(newVal)
End Property

Public Property Get DriverName() As String
    DriverName = DriverName_
End Property
Public Property Let DriverName(ByVal newVal As String)
    DriverName_ = Trim(newVal)
End Property

Public Property Get TimeTableNo() As String
    TimeTableNo = TimeTableNo_
End Property
Public Property Let TimeTableNo(ByVal newVal As String)
    TimeTableNo_ = Trim(newVal)
End Property

Public Property Get BoxOwnersCode() As String
    BoxOwnersCode = BoxOwnersCode_
End Property
Public Property Let BoxOwnersCode(ByVal newVal As String)
    BoxOwnersCode_ = Trim(newVal)
End Property

Public Property Get BoxType() As String
    BoxType = BoxType_
End Property
Public Property Let BoxType(ByVal newVal As String)
    BoxType_ = Trim(newVal)
End Property

Public Property Get BoxNo() As String
    BoxNo = BoxNo_
End Property
Public Property Let BoxNo(ByVal newVal As String)
    BoxNo_ = Trim(newVal)
End Property

Public Property Get TransactionStatus() As String
    TransactionStatus = BoxTransactionStatus_
End Property
Public Property Let TransactionStatus(ByVal newVal As String)
    BoxTransactionStatus_ = Trim(newVal)
End Property

Public Property Get BoxID() As String
    BoxID = BoxID_
End Property
Public Property Let BoxID(ByVal newVal As String)
    BoxID_ = Trim(newVal)
End Property

Public Property Get isChanged() As Boolean
    isChanged = isChanged_
End Property
Public Property Let isChanged(ByVal newVal As Boolean)
    isChanged_ = Trim(newVal)
End Property

Public Sub Save()
On Error GoTo ErrD
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String

    If isChanged_ = False Then Exit Sub
    Set tmpRec = New ADODB.Recordset
    If BoxTransactionStatus_ = "TFO" Then
        sqlStr = "SELECT * FROM Box_TransactionOutD WHERE TransactionID = '" & id_ & "' AND BoxOwnersCode = '" & BoxOwnersCode_ & "' AND BoxType = '" & BoxType_ & "' AND BoxCode = '" & BoxNo_ & "'"
    ElseIf BoxTransactionStatus_ = "TFI" Then
        sqlStr = "SELECT * FROM Box_TransactionInD WHERE TransactionID = '" & id_ & "' AND BoxOwnersCode = '" & BoxOwnersCode_ & "' AND BoxType = '" & BoxType_ & "' AND BoxCode = '" & BoxNo_ & "'"
    Else
        Exit Sub
    End If
    tmpRec.open sqlStr, DBConn, adOpenDynamic, adLockOptimistic
    If tmpRec.EOF Then
        tmpRec.AddNew
        tmpRec!TransactionID = id_
        tmpRec!JobDataCarID = JobDataCarID_
        tmpRec!VendorTruckID = VendorTruckID_
        tmpRec!TripNo = TripNo_
        tmpRec!CarRegistID = CarRegisID_
        tmpRec!DriverName = DriverName_
        tmpRec!TimeTableNo = TimeTableNo_
        tmpRec!BoxOwnersCode = BoxOwnersCode_
        tmpRec!BoxType = BoxType_
        tmpRec!BoxCode = BoxNo_
        tmpRec!BoxID = BoxOwnersCode_ & BoxType_ & BoxNo_
    End If
    tmpRec.Update
    tmpRec.Requery
    Call UpdateStatus(BoxOwnersCode_ & BoxType_ & BoxNo_, 0)
    Call UpdateReserve(JobDataCarID_, BoxType_, -1)
        
    Set tmpRec = Nothing
    DBConn.Execute sqlStr
    isChanged_ = False
    Exit Sub
ErrD:
    GetMsg "เกิดข้อผิดพลาดขณะบันทึกข้อมูล", vbCritical
    Exit Sub
End Sub

Public Function Delete() As String
On Error GoTo delErr:
    GetDBConnection DBConStock
    If BoxTransactionStatus_ = "TFO" Then
        DBConn.Execute "DELETE Box_TransactionOutD WHERE TransactionID = '" & id_ & "'"
    ElseIf BoxTransactionStatus_ = "TFI" Then
        DBConn.Execute "DELETE Box_TransactionInD WHERE TransactionID = '" & id_ & "'"
    End If

    Delete = ""
    Exit Function
delErr:
    Delete = err.Description
End Function

Public Sub UpdateStatus(ByVal BoxID As String, ByVal BoxStatus As Integer)
    Dim sqlStr As String
    If BoxTransactionStatus_ = "TFO" Then
        sqlStr = "UPDATE Box_BoxNo SET BoxStatus = " & BoxStatus
        sqlStr = sqlStr & " WHERE BoxID = '" & BoxID & "'"
    ElseIf BoxTransactionStatus_ = "TFI" Then
        sqlStr = "UPDATE Box_BoxNo SET BoxStatus = " & BoxStatus
        sqlStr = sqlStr & " WHERE BoxID = '" & BoxID & "'"
    End If
    DBConn.Execute sqlStr
End Sub

Public Sub UpdateReserve(ByVal JobDataCarID As String, ByVal BoxType As String, ByVal iCount As Integer)
    Dim sqlStr As String
    sqlStr = "UPDATE Box_BoxReserve SET BoxReservNumber = BoxReservNumber + " & iCount & vbCrLf
    sqlStr = sqlStr & "WHERE JobDataCarID = '" & JobDataCarID & "' AND BoxType = '" & BoxType_ & "'"
    DBConn.Execute sqlStr
End Sub

Public Sub SetDataFromRecordset(selRec As ADODB.Recordset)
    id_ = selRec!TransactionID
    JobDataCarID_ = selRec!JobDataCarID
    VendorTruckID_ = selRec!VendorTruckID
    TripNo_ = selRec!TripNo
    CarRegisID_ = selRec!CarRegistID
    DriverName_ = selRec!DriverName
    TimeTableNo_ = selRec!TimeTableNo
    BoxOwnersCode_ = selRec!BoxOwnersCode
    BoxType_ = selRec!BoxType
    BoxNo_ = selRec!BoxCode
    BoxTransactionStatus_ = selRec!BoxTransactionStatus
    BoxID_ = BoxOwnersCode_ & BoxType_ & BoxNo_
End Sub

Public Function TransactionDetailSearch(Optional strCond As String = "", Optional sType As String = "TFO") As Collection
On Error Resume Next
Dim tmpItem As BWGBoxTransactionD
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If sType = "TFO" Then
        sqlStr = "SELECT *" & vbCrLf
        sqlStr = sqlStr & "FROM         dbo.Box_TransactionOutD" & vbCrLf
    ElseIf sType = "TFI" Then
        sqlStr = "SELECT *" & vbCrLf
        sqlStr = sqlStr & "FROM         dbo.Box_TransactionInD" & vbCrLf
    End If
    
    sqlStr = sqlStr & "WHERE (" & strCond & ") ORDER BY BoxOwnersCode, BoxType"
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly, DBConStock)
    Do Until tmpRec.EOF
        Set tmpItem = New BWGBoxTransactionD
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.ID & "_" & tmpItem.BoxID & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set TransactionDetailSearch = tmpCol
    Set tmpCol = Nothing
End Function
