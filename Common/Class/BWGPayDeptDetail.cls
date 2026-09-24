VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BWGPayDeptDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim id_ As String
Dim MenifestID_ As String
Dim menifestNo_ As String
Dim menifestDocNo_ As String
Dim menifestDate_ As String
Dim menifestCustName_ As String
Dim expenseID_ As String
Dim expenseVendorName_ As String
Dim expenseDate_ As String
Dim billDesc_ As String
Dim amtToPay_ As Double
Dim isChanged_ As Boolean

Public Property Get ID() As String
    ID = id_
End Property
Public Property Let ID(ByVal newVal As String)
    id_ = Trim(newVal)
End Property

Public Property Get menifestID() As String
    menifestID = MenifestID_
End Property
Public Property Let menifestID(ByVal newVal As String)
    If Trim(MenifestID_) <> Trim(newVal) Then
        MenifestID_ = Trim(newVal)
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

Public Property Get MenifestDocNo() As String
    MenifestDocNo = menifestDocNo_
End Property
Public Property Let MenifestDocNo(ByVal newVal As String)
    If Trim(menifestDocNo_) <> Trim(newVal) Then
        menifestDocNo_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get MenifestDate() As String
    MenifestDate = menifestDate_
End Property
Public Property Let MenifestDate(ByVal newVal As String)
    If Trim(menifestDate_) <> Trim(newVal) Then
        menifestDate_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get MenifestCustName() As String
    MenifestCustName = menifestCustName_
End Property
Public Property Let MenifestCustName(ByVal newVal As String)
    If Trim(menifestCustName_) <> Trim(newVal) Then
        menifestCustName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get ExpenseID() As String
    ExpenseID = expenseID_
End Property
Public Property Let ExpenseID(ByVal newVal As String)
    If Trim(expenseID_) <> Trim(newVal) Then
        expenseID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get ExpenseVendorName() As String
    ExpenseVendorName = expenseVendorName_
End Property
Public Property Let ExpenseVendorName(ByVal newVal As String)
    If Trim(expenseVendorName_) <> Trim(newVal) Then
        expenseVendorName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get ExpenseDate() As String
    ExpenseDate = expenseDate_
End Property
Public Property Let ExpenseDate(ByVal newVal As String)
    If Trim(expenseDate_) <> Trim(newVal) Then
        expenseDate_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get BillDesc() As String
    BillDesc = billDesc_
End Property
Public Property Let BillDesc(ByVal newVal As String)
    If Trim(billDesc_) <> Trim(newVal) Then
        billDesc_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get AmtToPay() As Double
    AmtToPay = amtToPay_
End Property
Public Property Let AmtToPay(ByVal newVal As Double)
    If amtToPay_ <> newVal Then
        amtToPay_ = newVal
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
    sqlStr = "SELECT * FROM tbPayDeptDetail WHERE PayDeptID='" & id_ & "'"
    Set tmpRec = New ADODB.Recordset
    GetDBConnection
    tmpRec.open sqlStr, DBConn, adOpenDynamic, adLockOptimistic
    If tmpRec.EOF Then
        tmpRec.AddNew
        tmpRec("PayDeptID") = id_
    End If
    tmpRec("MenifestID") = Trim(MenifestID_)
    tmpRec("ExpenseID") = Trim(expenseID_)
    tmpRec("BillDesc") = Trim(billDesc_)
    tmpRec("AmtToPay") = Trim(amtToPay_)
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
    DBConnExc "Delete tbPayDeptDetail WHERE PayDeptID='" & id_ & "'"
    Delete = ""
    Exit Function
delErr:
    Delete = err.Description
End Function

Public Sub SetDataFromRecordset(selRec As ADODB.Recordset)
    id_ = Trim("" & selRec("PayDeptID"))
    MenifestID_ = Trim("" & selRec("MenifestID"))
    menifestNo_ = Trim("" & selRec("MenifestNo"))
    menifestDocNo_ = Trim("" & selRec("DocNo"))
    menifestDate_ = Trim("" & selRec("WorkDate"))
    menifestCustName_ = Trim("" & selRec("CustomerName"))
    expenseID_ = Trim("" & selRec("ExpenseID"))
    expenseVendorName_ = Trim("" & selRec("VendorName"))
    expenseDate_ = Trim("" & selRec("ExpenseDate"))
    billDesc_ = Trim("" & selRec("BillDesc"))
    amtToPay_ = selRec("AmtToPay")
End Sub
