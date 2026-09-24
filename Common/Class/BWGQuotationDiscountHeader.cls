VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BWGQTDiscountHeader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim id_ As String
Dim CustomerID_ As String
Dim customerName_ As String
Dim QuoDisNo_ As String
Dim CreateDate_ As String
Dim discount_ As Integer
Dim isCancel_ As String
Dim detail_ As String
Dim StartDate_ As String
Dim EndDate_ As String
Dim isQuotationDiscountAll_ As Boolean
Dim colQuotationDet_  As Collection

 Public Property Get ID() As String
    ID = id_
End Property
Public Property Let ID(ByVal newVal As String)
    id_ = Trim(newVal)
End Property

Public Property Get QuoDisNo() As String
    QuoDisNo = QuoDisNo_
End Property
Public Property Let QuoDisNo(ByVal newVal As String)
    QuoDisNo_ = Trim(newVal)
End Property

Public Property Get CustomerID() As String
    CustomerID = CustomerID_
End Property
Public Property Let CustomerID(ByVal newVal As String)
    CustomerID_ = Trim(newVal)
End Property
Public Property Get CustomerName() As String
    CustomerName = customerName_
End Property
Public Property Let CustomerName(ByVal newVal As String)
        customerName_ = Trim(newVal)
End Property
Public Property Get CreateDate() As String
    CreateDate = CreateDate_
End Property
Public Property Let CreateDate(ByVal newVal As String)
    CreateDate_ = Trim(newVal)
End Property

Public Property Get Discount() As Integer
    Discount = discount_
End Property
Public Property Let Discount(ByVal newVal As Integer)
    discount_ = Trim(newVal)
End Property

Public Property Get isCancel() As String
    isCancel = isCancel_
End Property
Public Property Let isCancel(ByVal newVal As String)
    isCancel_ = Trim(newVal)
End Property

Public Property Get Detail() As String
    Detail = detail_
End Property
Public Property Let Detail(ByVal newVal As String)
    detail_ = Trim(newVal)
End Property

Public Property Get StartDate() As String
    StartDate = StartDate_
End Property
Public Property Let StartDate(ByVal newVal As String)
    StartDate_ = Trim(newVal)
End Property

Public Property Get EndDate() As String
    EndDate = EndDate_
End Property
Public Property Let EndDate(ByVal newVal As String)
    EndDate_ = Trim(newVal)
End Property

Public Property Get isQuotationDiscountAll() As Boolean
        isQuotationDiscountAll = isQuotationDiscountAll_
End Property
Public Property Let isQuotationDiscountAll(ByVal newVal As Boolean)
    If isQuotationDiscountAll <> newVal Then
        isQuotationDiscountAll = newVal
    End If
End Property

Public Property Get QuotationDiscountDetails() As Collection
Dim tmpS As New BWGSearchManager
Dim i As Integer
    Set colQuotationDet_ = tmpS.QuotationDetailDiscountSearch("QuoDiscountHeaderID='" & id_ & "'")
    Set QuotationDiscountDetails = colQuotationDet_
End Property

Public Sub Save()
    Dim tmpRec As ADODB.Recordset
    Dim x%
    Dim sqlStr As String
        sqlStr = "SELECT * FROM tbQuotationDiscountHeader WHERE QuoDiscountHeaderID ='" & id_ & "'"
        Set tmpRec = New ADODB.Recordset
        GetDBConnection
        tmpRec.open sqlStr, DBConn, adOpenDynamic, adLockOptimistic
        If tmpRec.EOF Then
            tmpRec.AddNew
            tmpRec("QuoDiscountHeaderID") = id_
            tmpRec("CreateDate") = TodayDate
            tmpRec("QuoDisNo") = Trim(QuoDisNo_)
        End If
        tmpRec("CustomerID") = Trim(CustomerID_)
        tmpRec("Discount") = Trim(discount_)
        tmpRec("isCancel") = Trim(isCancel_)
        tmpRec("Detail") = Trim(detail_)
        tmpRec("StartDate") = Trim(StartDate_)
        tmpRec("EndDate") = Trim(EndDate_)
        tmpRec.Update
End Sub

Public Sub SetDataFromRecordset(selRec As ADODB.Recordset)
    id_ = Trim("" & selRec("QuoDiscountHeaderID"))
    QuoDisNo_ = Trim("" & selRec("QuoDisNo"))
    CustomerID_ = Trim("" & selRec("CustomerID"))
    customerName_ = Trim("" & selRec("CustomerName"))
    CreateDate_ = Trim("" & selRec("CreateDate"))
    discount_ = selRec("Discount")
    isCancel_ = Trim("" & selRec("isCancel"))
    detail_ = Trim("" & selRec("Detail"))
    StartDate_ = Trim("" & selRec("StartDate"))
    EndDate_ = Trim("" & selRec("EndDate"))
End Sub

Public Sub CancelQuotationDiscount()
On Error GoTo ErrD
    isCancel_ = "Y"
    DBConnExc "Update tbQuotationDiscountHeader Set isCancel='Y' WHERE QuoDiscountHeaderID='" & id_ & "'"
    DBConnExc "Update tbQuotationHeader Set isQuoDisount='N' WHERE QuotationID in (Select QuotationID From tbQuotationDiscountDetail WHERE QuoDiscountHeaderID='" & id_ & "')"
ErrD:
    GetMsg "", vbCritical
    Exit Sub
End Sub
