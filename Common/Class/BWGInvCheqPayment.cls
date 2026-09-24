VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BWGInvCheqPayment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim id_ As String
Dim printInvID_ As String
Dim cheqNo_ As String
Dim bankName_ As String
Dim cheqAmt_ As Double
Dim cheqDate_ As String
Dim isChanged_ As Boolean

Public Property Get ID() As String
    ID = id_
End Property
Public Property Let ID(ByVal newVal As String)
    id_ = Trim(newVal)
End Property

Public Property Get PrintInvID() As String
    PrintInvID = printInvID_
End Property
Public Property Let PrintInvID(ByVal newVal As String)
    If Trim(printInvID_) <> Trim(newVal) Then
        printInvID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get CheqNo() As String
    CheqNo = cheqNo_
End Property
Public Property Let CheqNo(ByVal newVal As String)
    If Trim(cheqNo_) <> Trim(newVal) Then
        cheqNo_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get BankName() As String
    BankName = bankName_
End Property
Public Property Let BankName(ByVal newVal As String)
    If Trim(bankName_) <> Trim(newVal) Then
        bankName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get CheqAmt() As Double
    CheqAmt = cheqAmt_
End Property
Public Property Let CheqAmt(ByVal newVal As Double)
    If cheqAmt_ <> newVal Then
        cheqAmt_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get cheqDate() As String
    cheqDate = cheqDate_
End Property
Public Property Let cheqDate(ByVal newVal As String)
    If Trim(cheqDate_) <> Trim(newVal) Then
        cheqDate_ = Trim(newVal)
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
    sqlStr = "SELECT * FROM tbInvCheqPayment WHERE CheqPaymentID='" & id_ & "'"
    Set tmpRec = New ADODB.Recordset
    GetDBConnection
    tmpRec.open sqlStr, DBConn, adOpenDynamic, adLockOptimistic
    If tmpRec.EOF Then
        tmpRec.AddNew
        tmpRec("CheqPaymentID") = id_
    End If
    tmpRec("PrintInvID") = Trim(printInvID_)
    tmpRec("CheqNo") = Trim(cheqNo_)
    tmpRec("BankName") = Trim(bankName_)
    tmpRec("CheqAmt") = cheqAmt_
    tmpRec("CheqDate") = Trim(cheqDate_)
    tmpRec("MarkForDelete") = "N"
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
    DBConnExc "Delete tbInvCheqPayment WHERE CheqPaymentID='" & id_ & "'"
    Delete = ""
    Exit Function
delErr:
    Delete = err.Description
End Function

Public Sub SetDataFromRecordset(selRec As ADODB.Recordset)
    id_ = selRec("CheqPaymentID")
    printInvID_ = Trim("" & selRec("PrintInvID"))
    cheqNo_ = Trim("" & selRec("CheqNo"))
    bankName_ = Trim("" & selRec("BankName"))
    If IsNull(selRec("CheqAmt")) = False Then cheqAmt_ = Trim("" & selRec("CheqAmt"))
    cheqDate_ = Trim("" & selRec("CheqDate"))
End Sub
