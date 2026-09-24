VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BWGQTDiscountDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim id_ As String
Dim QuoDiscountHeaderID_  As String
Dim QuotationID_ As String
Dim QuotationNo_ As String
Dim QuotationDate_ As String
Dim QuotationStatus_ As String

Public Property Get ID() As String
    ID = id_
End Property
Public Property Let ID(ByVal newVal As String)
    id_ = Trim(newVal)
End Property

Public Property Get QuoDisCountHeaderID() As String
    QuoDisCountHeaderID = QuoDiscountHeaderID_
End Property
Public Property Let QuoDisCountHeaderID(ByVal newVal As String)
    QuoDiscountHeaderID_ = Trim(newVal)
End Property

Public Property Get QuotationID() As String
    QuotationID = QuotationID_
End Property
Public Property Let QuotationID(ByVal newVal As String)
    QuotationID_ = Trim(newVal)
End Property

Public Property Get QuotationNo() As String
    QuotationNo = QuotationNo_
End Property
Public Property Let QuotationNo(ByVal newVal As String)
    QuotationNo_ = Trim(newVal)
End Property

Public Property Get QuotationDate() As String
    QuotationDate = QuotationDate_
End Property
Public Property Let QuotationDate(ByVal newVal As String)
    QuotationDate_ = Trim(newVal)
End Property
 
 Public Property Get QuotationStatus() As String
    QuotationStatus = QuotationStatus_
 End Property
 Public Property Let QuotationStatus(ByVal newVal As String)
    QuotationStatus_ = Trim(newVal)
 End Property
 
 Public Sub Save()
On Error GoTo ErrD
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
    sqlStr = "SELECT * FROM tbQuotationDiscountDetail WHERE QuoDiscountDetailID='" & id_ & "'"
    Set tmpRec = New ADODB.Recordset
    GetDBConnection
    tmpRec.open sqlStr, DBConn, adOpenDynamic, adLockOptimistic
    If tmpRec.EOF Then
        tmpRec.AddNew
        tmpRec("QuoDiscountDetailID") = id_
    End If
        tmpRec("QuoDisCountHeaderID") = Trim(QuoDiscountHeaderID_)
        tmpRec("QuotationID") = Trim(QuotationID_)
        tmpRec("QuotationNO") = Trim(QuotationNo_)
        tmpRec("QuotationDate") = Trim(QuotationDate_)
        tmpRec("QuotationStatus") = Trim(QuotationStatus_)
        tmpRec.Update
        Set tmpRec = Nothing
        CloseDBConnection
        Exit Sub
ErrD:
    GetMsg "", vbCritical
    Exit Sub
End Sub

Public Sub SetDataFromRecordset(selRec As ADODB.Recordset)
    id_ = Trim("" & selRec("QuoDiscountDetailID"))
    QuoDiscountHeaderID_ = Trim("" & selRec("QuoDisCountHeaderID"))
    QuotationID_ = Trim("" & selRec("QuotationID"))
    QuotationNo_ = Trim("" & selRec("QuotationNo"))
    QuotationDate_ = Trim("" & selRec("QuotationDate"))
    QuotationStatus_ = Trim("" & selRec("QuotationStatus"))
End Sub
