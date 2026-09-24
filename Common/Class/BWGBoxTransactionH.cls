VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BWGBoxTransactionH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim id_ As String
Dim JobDataCarID_ As String
Dim TimeTableNo_ As String
Dim VendorTruckID_ As String
Dim TransactionDate_ As String
Dim TransactionStatus_ As String
Dim TransactionStatusDate_ As String
Dim TransactionDesc_ As String
Dim TransactionType_ As String  '1=Transfer, 2=Receive
Dim isChanged_ As Boolean

'for detail
Dim colTransDet_ As Collection

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

Public Property Get TimeTableNo() As String
    TimeTableNo = TimeTableNo_
End Property
Public Property Let TimeTableNo(ByVal newVal As String)
    TimeTableNo_ = Trim(newVal)
End Property

Public Property Get VendorTruckID() As String
    VendorTruckID = VendorTruckID_
End Property
Public Property Let VendorTruckID(ByVal newVal As String)
    VendorTruckID_ = Trim(newVal)
End Property

Public Property Get TransactionDate() As String
    TransactionDate = TransactionDate_
End Property
Public Property Let TransactionDate(ByVal newVal As String)
    If Trim(TransactionDate_) <> Trim(newVal) Then
        TransactionDate_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get TransactionStatus() As String
    TransactionStatus = TransactionStatus_
End Property
Public Property Let TransactionStatus(ByVal newVal As String)
    TransactionStatus_ = Trim(newVal)
End Property

Public Property Get TransactionStatusDate() As String
    TransactionStatusDate = TransactionStatusDate_
End Property
Public Property Let TransactionStatusDate(ByVal newVal As String)
    TransactionStatusDate_ = Trim(newVal)
End Property

Public Property Get TransactionDesc() As String
    TransactionDesc = TransactionDesc_
End Property
Public Property Let TransactionDesc(ByVal newVal As String)
    If Trim(TransactionDesc_) <> Trim(newVal) Then
        TransactionDesc_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get TransactionType() As String
    TransactionType = TransactionType_
End Property
Public Property Let TransactionType(ByVal newVal As String)
    TransactionType_ = Trim(newVal)
End Property

Public Property Get isChanged() As Boolean
    isChanged = isChanged_
End Property
Public Property Let isChanged(ByVal newVal As Boolean)
    isChanged_ = newVal
End Property

'for detail
Public Property Get TransactionDetails() As Collection
Dim tmpS As New BWGBoxTransactionD
    If colTransDet_ Is Nothing Then Set colTransDet_ = tmpS.TransactionDetailSearch("TransactionID='" & id_ & "'", TransactionType_)
    Set tmpS = Nothing
    Set TransactionDetails = colTransDet_
End Property
Public Property Set TransactionDetails(ByVal newVal As Collection)
    Set colTransDet_ = newVal
End Property

Public Sub Save()
    On Error GoTo ErrD
    Dim tmpRec As ADODB.Recordset
    Dim sqlStr As String
    Dim X  As Integer
    GetDBConnection DBConStock
    Set tmpRec = New ADODB.Recordset
    
    If TransactionType_ = "TFO" Then
        sqlStr = "SELECT * FROM Box_TransactionOutH WHERE TransactionID='" & id_ & "'"
    ElseIf TransactionType_ = "TFI" Then
        sqlStr = "SELECT * FROM Box_TransactionInH WHERE TransactionID='" & id_ & "'"
    End If
    
    tmpRec.open sqlStr, DBConn, adOpenDynamic, adLockOptimistic
    If tmpRec.EOF Then
        tmpRec.AddNew
        tmpRec!TransactionID = id_
        tmpRec!JobDataCarID = JobDataCarID_
        tmpRec!TimeTableNo = TimeTableNo_
        tmpRec!VendorTruckID = VendorTruckID_
    End If
    tmpRec!TransactionDate = TransactionDate_
    tmpRec!TransactionType = TransactionType_
    tmpRec!TransactionStatus = TransactionStatus_
    tmpRec!TransactionStatusDate = TransactionStatusDate_
    tmpRec!TransactionDesc = TransactionDesc_
    tmpRec.Update
    tmpRec.Requery

    If Not colTransDet_ Is Nothing Then
'        DBConn.Execute "DELETE tbTransactionDetail WHERE TransactionID = '" & id_ & "'"
        For X = 1 To colTransDet_.Count
            colTransDet_(X).TransactionStatus = "TFO"
            colTransDet_(X).Save
        Next
    End If
        
    isChanged_ = False
    CloseDBConnection
    Exit Sub
ErrD:
    GetMsg "เกิดข้อผิดพลาดขณะบันทึกข้อมูล", vbCritical
    Exit Sub
End Sub

Public Function Delete() As String
On Error GoTo delErr:
    Dim X As Integer
    GetDBConnection DBConStock
    DBConn.BeginTrans
    
    If Not colTransDet_ Is Nothing Then
        For X = 1 To colTransDet_.Count
'            colTransDet_(X).TransactionStatus = TransactionType_
'            colTransDet_(X).Delete
'            colTransDet_(X).UpdateStatus colTransDet_(X).BoxID, 1
'            colTransDet_(X).UpdateReserve colTransDet_(X).JobDataCarID, colTransDet_(X).BoxType, 1
            DBConn.Execute " UPDATE  Box_BoxNo  SET  BoxStatus = '1'  WHERE BoxID='" & colTransDet_(X).BoxID & "'"

        Next
    End If

'    If TransactionType_ = "TFO" Then
'        DBConn.Execute "Delete Box_TransactionOutD WHERE TransactionID='" & id_ & "'"
'        DBConn.Execute "Delete Box_TransactionOutH WHERE TransactionID='" & id_ & "'"
'    ElseIf TransactionType_ = "TFI" Then
'        DBConn.Execute "Delete Box_TransactionInD WHERE TransactionID='" & id_ & "'"
'        DBConn.Execute "Delete Box_TransactionInH WHERE TransactionID='" & id_ & "'"
'    End If
    
    DBConn.CommitTrans
    Delete = ""
    Exit Function
delErr:
    Delete = err.Description
    DBConn.RollbackTrans
End Function

Public Function Cancel() As String
On Error GoTo ErrD:
    Dim ToDate As Date
    Dim X As Integer
    GetDBConnection DBConStock
    DBConn.BeginTrans
    If TransactionType_ = "TFO" Then
        DBConn.Execute "UPDATE Box_TransactionOutH  SET TransactionStatus = 'CANCEL', TransactionStatusDate = '" & Format(Date, "yyyy/MM/dd") & "', TransactionDesc = '" & Trim(CurrentUser.StaffFName & " " & CurrentUser.StaffLName) & "' WHERE TransactionID='" & id_ & "'"
    ElseIf TransactionType_ = "TFI" Then
        DBConn.Execute "UPDATE Box_TransactionInH  SET TransactionStatus = 'CANCEL', TransactionStatusDate = '" & Format(Date, "yyyy/MM/dd") & "', TransactionDesc = '" & Trim(CurrentUser.StaffFName & " " & CurrentUser.StaffLName) & "' WHERE TransactionID='" & id_ & "'"
    End If
    
    If Not colTransDet_ Is Nothing Then
        For X = 1 To colTransDet_.Count
            colTransDet_(X).TransactionStatus = TransactionType_
            colTransDet_(X).Delete
            colTransDet_(X).UpdateStatus colTransDet_(X).BoxID, 1
        Next
    End If
    
    DBConn.CommitTrans
    Cancel = ""
    Exit Function
ErrD:
    Cancel = err.Description
    DBConn.RollbackTrans
End Function

Public Sub SetDataFromRecordset(selRec As ADODB.Recordset)
    id_ = selRec!TransactionID
    JobDataCarID_ = selRec!JobDataCarID
    TimeTableNo_ = selRec!TimeTableNo
    VendorTruckID_ = selRec!VendorTruckID
    TransactionDate_ = selRec!TransactionDate
    TransactionType_ = selRec!TransactionType
    TransactionDesc_ = selRec!TransactionDesc
    TransactionStatus_ = selRec!TransactionStatus
    TransactionStatusDate_ = selRec!TransactionStatusDate
End Sub

Public Function TransactionHeaderSearch(Optional strCond As String = "", Optional sType As String = "TFO") As Collection
On Error Resume Next
Dim tmpItem As BWGBoxTransactionH
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If sType = "TFO" Then
        sqlStr = "SELECT DISTINCT  *  FROM Box_TransactionOutH WHERE (" & strCond & ") ORDER BY TransactionDate"
    ElseIf sType = "TFI" Then
        sqlStr = "SELECT DISTINCT  *  FROM Box_TransactionInH WHERE (" & strCond & ") ORDER BY TransactionDate"
    End If
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly, DBConStock)
    Do Until tmpRec.EOF
        Set tmpItem = New BWGBoxTransactionH
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.ID & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set TransactionHeaderSearch = tmpCol
    Set tmpCol = Nothing
End Function
