VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BWGAdjDeptDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim id_ As String
Dim AdjustID_ As String
Dim rowPos_ As Integer
Dim billingHeaderID_ As String
Dim manifestID_ As String
Dim jobDetID_ As String
Dim BillingNo_ As String
Dim adjBillDesc_ As String
Dim adjBillQty_ As Double
Dim adjUnitPrice_ As Double
Dim adjUnitName_ As String
Dim menifestDocNo_ As String
Dim menifestNo_ As String
Dim isChanged_ As Boolean

Public Property Get ID() As String
    ID = id_
End Property
Public Property Let ID(ByVal newVal As String)
    id_ = Trim(newVal)
End Property

Public Property Get AdjustID() As String
    AdjustID = Trim(AdjustID_)
End Property
Public Property Let AdjustID(ByVal newVal As String)
    If Trim(AdjustID_) <> Trim(newVal) Then
        AdjustID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get RowPos() As Integer
    RowPos = rowPos_
End Property
Public Property Let RowPos(ByVal newVal As Integer)
    If rowPos_ <> newVal Then
        rowPos_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get BillingHeaderID() As String
    BillingHeaderID = Trim(billingHeaderID_)
End Property
Public Property Let BillingHeaderID(ByVal newVal As String)
    If Trim(billingHeaderID_) <> Trim(newVal) Then
        billingHeaderID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get ManifestID() As String
    ManifestID = Trim(manifestID_)
End Property
Public Property Let ManifestID(ByVal newVal As String)
    If Trim(manifestID_) <> Trim(newVal) Then
        manifestID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get JobDetID() As String
    JobDetID = Trim(jobDetID_)
End Property
Public Property Let JobDetID(ByVal newVal As String)
    If Trim(jobDetID_) <> Trim(newVal) Then
        jobDetID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get BillingNo() As String
    BillingNo = Trim(BillingNo_)
End Property
Public Property Let BillingNo(ByVal newVal As String)
    If Trim(BillingNo_) <> Trim(newVal) Then
        BillingNo_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get AdjBillDesc() As String
    AdjBillDesc = Trim(adjBillDesc_)
End Property
Public Property Let AdjBillDesc(ByVal newVal As String)
    If Trim(adjBillDesc_) <> Trim(newVal) Then
        adjBillDesc_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get AdjBillQty() As Double
    AdjBillQty = adjBillQty_
End Property
Public Property Let AdjBillQty(ByVal newVal As Double)
    If adjBillQty_ <> newVal Then
        adjBillQty_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get AdjUnitPrice() As Double
    AdjUnitPrice = adjUnitPrice_
End Property
Public Property Let AdjUnitPrice(ByVal newVal As Double)
    If adjUnitPrice_ <> newVal Then
        adjUnitPrice_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get AdjUnitName() As String
    AdjUnitName = Trim(adjUnitName_)
End Property
Public Property Let AdjUnitName(ByVal newVal As String)
    If Trim(adjUnitName_) <> Trim(newVal) Then
        adjUnitName_ = Trim(newVal)
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

Public Property Get isChanged() As Boolean
    isChanged = isChanged_
End Property
Public Property Let isChanged(ByVal newVal As Boolean)
    isChanged_ = newVal
End Property

Public Sub Save()
On Error GoTo ErrD
Dim tmpRec As ADODB.Recordset
Dim x%, sqlStr As String
    sqlStr = "SELECT * FROM tbAdjustDeptDetail WHERE AdjustDetID='" & id_ & "'"
    Set tmpRec = New ADODB.Recordset
    GetDBConnection
    tmpRec.open sqlStr, DBConn, adOpenDynamic, adLockOptimistic
    If tmpRec.EOF Then
        tmpRec.AddNew
        tmpRec("AdjustDetID") = id_
    End If
    tmpRec("AdjustID") = Trim(AdjustID_)
    tmpRec("RowPos") = Trim(rowPos_)
    tmpRec("BillingHeaderID") = Trim(billingHeaderID_)
    tmpRec("AdjBillDesc") = Trim(adjBillDesc_)
    tmpRec("AdjBillQty") = Trim(adjBillQty_)
    tmpRec("AdjBillUnitPrice") = Trim(adjUnitPrice_)
    tmpRec("AdjBillUnitName") = Trim(adjUnitName_)
    tmpRec("MenifestID") = Trim(manifestID_)
    tmpRec("JobDetID") = Trim(jobDetID_)
    tmpRec.Update
'    tmpRec.Close
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
    DBConnExc "Delete tbAdjustDeptDetail WHERE ApproveLogID='" & id_ & "'"
    Delete = ""
    Exit Function
delErr:
    Delete = err.Description
End Function

Public Sub SetDataFromRecordset(selRec As ADODB.Recordset)
    id_ = Trim("" & selRec("AdjustDetID"))
    AdjustID_ = Trim("" & selRec("AdjustID"))
    rowPos_ = Trim("" & selRec("RowPos"))
    billingHeaderID_ = Trim("" & selRec("BillingHeaderID"))
    manifestID_ = Trim("" & selRec("MenifestID"))
    jobDetID_ = Trim("" & selRec("JobDetID"))
    BillingNo_ = Trim("" & selRec("BillingNo"))
    adjBillDesc_ = Trim("" & selRec("AdjBillDesc"))
    adjBillQty_ = Trim("" & selRec("AdjBillQty"))
    adjUnitPrice_ = Trim("" & selRec("AdjBillUnitPrice"))
    adjUnitName_ = Trim("" & selRec("AdjBillUnitName"))
    menifestDocNo_ = Trim("" & selRec("DocNo"))
    menifestNo_ = Trim("" & selRec("MenifestNo"))
End Sub
