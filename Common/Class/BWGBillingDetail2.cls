VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BWGBillingDetail2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim id_ As String
Dim billHeaderID_ As String
Dim rowID_ As Integer
Dim MenifestID_ As String
Dim menifestDocNo_ As String
Dim menifestNo_ As String
Dim menifestDate_ As String
Dim jobDetID_ As String
Dim wasteDataID_ As String
Dim wasteName_ As String
Dim wasteType_  As String
Dim billDesc_ As String
Dim showWeight_ As Double
Dim weightUnit_ As String
Dim ratePrice_ As Double
Dim rateUnit_ As String
Dim isDescription_ As String
Dim isChanged_ As Boolean

Public Property Get ID() As String
    ID = id_
End Property
Public Property Let ID(ByVal newVal As String)
    id_ = Trim(newVal)
End Property

Public Property Get BillHeaderID() As String
    BillHeaderID = Trim(billHeaderID_)
End Property
Public Property Let BillHeaderID(ByVal newVal As String)
    If Trim(billHeaderID_) <> Trim(newVal) Then
        billHeaderID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get RowID() As Integer
    RowID = rowID_
End Property
Public Property Let RowID(ByVal newVal As Integer)
    If rowID_ <> newVal Then
        rowID_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get menifestID() As String
    menifestID = Trim(MenifestID_)
End Property
Public Property Let menifestID(ByVal newVal As String)
    If Trim(MenifestID_) <> Trim(newVal) Then
        MenifestID_ = Trim(newVal)
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

Public Property Get MenifestDate() As String
    MenifestDate = Trim(menifestDate_)
End Property
Public Property Let MenifestDate(ByVal newVal As String)
    If Trim(menifestDate_) <> Trim(newVal) Then
        menifestDate_ = Trim(newVal)
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

Public Property Get wastedataID() As String
    wastedataID = Trim(wasteDataID_)
End Property
Public Property Let wastedataID(ByVal newVal As String)
    If Trim(wasteDataID_) <> Trim(newVal) Then
        wasteDataID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get WasteName() As String
    WasteName = Trim(wasteName_)
End Property
Public Property Let WasteName(ByVal newVal As String)
    If Trim(wasteName_) <> Trim(newVal) Then
        wasteName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get WasteType() As String
    WasteType = Trim(wasteType_)
End Property
Public Property Let WasteType(ByVal newVal As String)
    If Trim(wasteType_) <> Trim(newVal) Then
        wasteType_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get BillDesc() As String
    BillDesc = Trim(billDesc_)
End Property
Public Property Let BillDesc(ByVal newVal As String)
    If Trim(billDesc_) <> Trim(newVal) Then
        billDesc_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get ShowWeight() As Double
    ShowWeight = showWeight_
End Property
Public Property Let ShowWeight(ByVal newVal As Double)
    If showWeight_ <> newVal Then
        showWeight_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get WeightUnit() As String
    WeightUnit = Trim(weightUnit_)
End Property
Public Property Let WeightUnit(ByVal newVal As String)
    If Trim(weightUnit_) <> Trim(newVal) Then
        weightUnit_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get RatePrice() As Double
    RatePrice = ratePrice_
End Property
Public Property Let RatePrice(ByVal newVal As Double)
    If ratePrice_ <> newVal Then
        ratePrice_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get RateUnit() As String
    RateUnit = Trim(rateUnit_)
End Property
Public Property Let RateUnit(ByVal newVal As String)
    If Trim(rateUnit_) <> Trim(newVal) Then
        rateUnit_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get isDescription() As Boolean
    If Trim(isDescription_) = "Y" Then
        isDescription = True
    Else
        isDescription = False
    End If
End Property
Public Property Let isDescription(ByVal newVal As Boolean)
Dim tmpStr As String
    If newVal Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(isDescription_) <> Trim(tmpStr) Then
        isDescription_ = Trim(tmpStr)
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
Dim sqlStr As String
    sqlStr = "SELECT * FROM tbBillingDetail2 WHERE BillingDetailID='" & id_ & "'"
    Set tmpRec = New ADODB.Recordset
    GetDBConnection
    tmpRec.open sqlStr, DBConn, adOpenDynamic, adLockOptimistic
    If tmpRec.EOF Then
        tmpRec.AddNew
        tmpRec("BillingDetailID") = id_
    End If
    tmpRec("BillingHeaderID") = billHeaderID_
    tmpRec("RowID") = rowID_
    tmpRec("MenifestID") = MenifestID_
    tmpRec("JobDetID") = jobDetID_
    tmpRec("BillDesc2") = billDesc_
    tmpRec("ShowWeight") = showWeight_
    tmpRec("WeightUnit") = weightUnit_
    tmpRec("RatePrice") = ratePrice_
    tmpRec("RateUnit") = rateUnit_
    If Trim(isDescription_) = "" Then isDescription_ = "N"
    tmpRec("isDescription") = isDescription_
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
    DBConnExc "Delete tbBillingDetail2 WHERE BillingDetailID='" & id_ & "'"
    Delete = ""
    Exit Function
delErr:
    Delete = err.Description
End Function

Public Sub SetDataFromRecordset(selRec As ADODB.Recordset)
    id_ = Trim("" & selRec("BillingDetailID"))
    billHeaderID_ = Trim("" & selRec("BillingHeaderID"))
    rowID_ = selRec("RowID")
    MenifestID_ = Trim("" & selRec("MenifestID"))
    menifestDocNo_ = Trim("" & selRec("DocNo"))
    menifestNo_ = Trim("" & selRec("MenifestNo"))
    menifestDate_ = Trim("" & selRec("WorkDate"))
    jobDetID_ = Trim("" & selRec("JobDetID"))
    wasteDataID_ = Trim("" & selRec("WasteDataID"))
    wasteName_ = Trim("" & selRec("WasteName"))
    wasteType_ = Trim("" & selRec("WasteType"))
    billDesc_ = Trim("" & selRec("BillDesc2"))
    showWeight_ = Trim("" & selRec("ShowWeight"))
    weightUnit_ = Trim("" & selRec("WeightUnit"))
    ratePrice_ = selRec("RatePrice")
    rateUnit_ = selRec("RateUnit")
    isDescription_ = selRec("isDescription")
End Sub
