VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BWGVendorOtherExpense"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim id_ As String
Dim vendorID_ As String
Dim vendorName_ As String
Dim expenseDate_ As String
Dim billDesc_ As String
Dim buyQty_ As Double
Dim netPrice_ As Double
Dim isPaid_ As String
Dim remDesc_ As String
Dim isChanged_ As Boolean

Public Property Get ID() As String
    ID = id_
End Property
Public Property Let ID(ByVal newVal As String)
    id_ = Trim(newVal)
End Property

Public Property Get VendorID() As String
    VendorID = vendorID_
End Property
Public Property Let VendorID(ByVal newVal As String)
    If Trim(vendorID_) <> Trim(newVal) Then
        vendorID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get VendorName() As String
    VendorName = vendorName_
End Property
Public Property Let VendorName(ByVal newVal As String)
    If Trim(vendorName_) <> Trim(newVal) Then
        vendorName_ = Trim(newVal)
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

Public Property Get BuyQty() As Double
    BuyQty = buyQty_
End Property
Public Property Let BuyQty(ByVal newVal As Double)
    If buyQty_ <> newVal Then
        buyQty_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get NetPrice() As Double
    NetPrice = netPrice_
End Property
Public Property Let NetPrice(ByVal newVal As Double)
    If netPrice_ <> newVal Then
        netPrice_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get isPaid() As Boolean
    If Trim(isPaid_) = "Y" Then
        isPaid = True
    Else
        isPaid = False
    End If
End Property
Public Property Let isPaid(ByVal newVal As Boolean)
Dim tmpStr As String
    If newVal Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If isPaid_ <> Trim(tmpStr) Then
        isPaid_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property

Public Property Get RemDesc() As String
    RemDesc = remDesc_
End Property
Public Property Let RemDesc(ByVal newVal As String)
    If Trim(remDesc_) <> Trim(newVal) Then
        remDesc_ = Trim(newVal)
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
    sqlStr = "SELECT *, CONVERT(DATETIME, GETDATE(), 103) AS GETDATE FROM tbVendorOtherExpense WHERE ExpenseID='" & id_ & "'"
    Set tmpRec = New ADODB.Recordset
    GetDBConnection
    tmpRec.open sqlStr, DBConn, adOpenDynamic, adLockOptimistic
    If tmpRec.EOF Then
        tmpRec.AddNew
        tmpRec("ExpenseID") = id_
    End If
    tmpRec("VendorID") = Trim(vendorID_)
    tmpRec("ExpenseDate") = Trim(expenseDate_)
    tmpRec("BillDesc") = Trim(billDesc_)
    tmpRec("BuyQty") = buyQty_
    tmpRec("NetPrice") = netPrice_
    tmpRec("isPaid") = Trim(isPaid_)
    tmpRec("RemDesc") = Trim(remDesc_)
    
    tmpRec!LastUser = Trim(Trim(CurrentUser.StaffFName) & " " & Trim(CurrentUser.StaffLName))
    tmpRec!LastUpdate = GetSystemDateTime.SystemDateTime
    
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
    DBConnExc "Delete tbVendorOtherExpense WHERE ExpenseID='" & id_ & "'"
    Delete = ""
    Exit Function
delErr:
    Delete = err.Description
End Function

Public Sub SetDataFromRecordset(selRec As ADODB.Recordset)
    id_ = selRec("ExpenseID")
    vendorID_ = Trim("" & selRec("VendorID"))
    vendorName_ = Trim("" & selRec("NameTH"))
    expenseDate_ = Trim("" & selRec("ExpenseDate"))
    billDesc_ = Trim("" & selRec("BillDesc"))
    If IsNull(selRec("BuyQty")) = False Then buyQty_ = selRec("BuyQty")
    If IsNull(selRec("NetPrice")) = False Then netPrice_ = selRec("NetPrice")
    isPaid_ = Trim("" & selRec("isPaid"))
    remDesc_ = Trim("" & selRec("RemDesc"))
End Sub
