VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BWGAdjDeptHeader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim id_ As String
Dim issuedDate_ As String
Dim AdjustBillNo_ As String
Dim companyID_ As String
Dim companyName_ As String
Dim CustomerID_ As String
Dim customerName_ As String
Dim adjReason_ As String
Dim issuedStaffID_ As String
Dim issuedStaffName_ As String
Dim recBillPerson_ As String
Dim AdjType_ As String
Dim isCancel_ As String
Dim cancelStaffID_ As String
Dim billAddrID_ As String
Dim billAddr_ As String
Dim oldBillAmt_ As Double
Dim correctBillAmt_ As Double
Dim vatRate_ As Double
Dim colAdjDet_ As Collection
Dim isChanged_ As Boolean

Public Property Get ID() As String
    ID = id_
End Property
Public Property Let ID(ByVal newVal As String)
    id_ = Trim(newVal)
End Property

Public Property Get IssuedDate() As String
    IssuedDate = Trim(issuedDate_)
End Property
Public Property Let IssuedDate(ByVal newVal As String)
    If Trim(issuedDate_) <> Trim(newVal) Then
        issuedDate_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get AdjustBillNo() As String
    AdjustBillNo = Trim(AdjustBillNo_)
End Property
Public Property Let AdjustBillNo(ByVal newVal As String)
    If Trim(AdjustBillNo_) <> Trim(newVal) Then
        AdjustBillNo_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get CompanyID() As String
    CompanyID = Trim(companyID_)
End Property
Public Property Let CompanyID(ByVal newVal As String)
    If Trim(companyID_) <> Trim(newVal) Then
        companyID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get CompanyName() As String
    CompanyName = Trim(companyName_)
End Property
Public Property Let CompanyName(ByVal newVal As String)
    If Trim(companyName_) <> Trim(newVal) Then
        companyName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get CustomerID() As String
    CustomerID = Trim(CustomerID_)
End Property
Public Property Let CustomerID(ByVal newVal As String)
    If Trim(CustomerID_) <> Trim(newVal) Then
        CustomerID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get BillAddrID() As String
    BillAddrID = Trim(billAddrID_)
End Property
Public Property Let BillAddrID(ByVal newVal As String)
    If Trim(billAddrID_) <> Trim(newVal) Then
        billAddrID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get BillAddr() As String
    BillAddr = Trim(billAddr_)
End Property
Public Property Let BillAddr(ByVal newVal As String)
    If Trim(billAddr_) <> Trim(newVal) Then
        billAddr_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get CustomerName() As String
    CustomerName = Trim(customerName_)
End Property
Public Property Let CustomerName(ByVal newVal As String)
    If Trim(customerName_) <> Trim(newVal) Then
        customerName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get AdjReason() As String
    AdjReason = Trim(adjReason_)
End Property
Public Property Let AdjReason(ByVal newVal As String)
    If Trim(adjReason_) <> Trim(newVal) Then
        adjReason_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get IssuedStaffID() As String
    IssuedStaffID = Trim(issuedStaffID_)
End Property
Public Property Let IssuedStaffID(ByVal newVal As String)
    If Trim(issuedStaffID_) <> Trim(newVal) Then
        issuedStaffID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get IssuedStaffName() As String
    IssuedStaffName = Trim(issuedStaffName_)
End Property
Public Property Let IssuedStaffName(ByVal newVal As String)
    If Trim(issuedStaffName_) <> Trim(newVal) Then
        issuedStaffName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get RecBillPerson() As String
    RecBillPerson = Trim(recBillPerson_)
End Property
Public Property Let RecBillPerson(ByVal newVal As String)
    If Trim(recBillPerson_) <> Trim(newVal) Then
        recBillPerson_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get AdjustType() As String
    'A = ลดหนี้ ไม่ใช่ใบกำกับ
    'B = ลดหนี้ ่ใบกำกับ
    'C = เพิ่มหนี้  ไม่ใช่ใบกำกับ
    'D = เพิ่มหนี้ ่ใบกำกับ
    AdjustType = Trim(AdjType_)
End Property
Public Property Let AdjustType(ByVal newVal As String)
    If Trim(AdjType_) <> Trim(newVal) Then
        AdjType_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get isCancel() As Boolean
    If isCancel_ = "Y" Then
        isCancel = True
    Else
        isCancel = False
    End If
End Property
Public Property Let isCancel(ByVal newVal As Boolean)
Dim tmpStr As String
    If newVal Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(isCancel_) <> Trim(tmpStr) Then
        isCancel_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property

Public Property Get CancelStaffID() As String
    CancelStaffID = Trim(cancelStaffID_)
End Property
Public Property Let CancelStaffID(ByVal newVal As String)
    If Trim(cancelStaffID_) <> Trim(newVal) Then
        cancelStaffID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get OldBillAmt() As Double
    OldBillAmt = oldBillAmt_
End Property
Public Property Let OldBillAmt(ByVal newVal As Double)
    If oldBillAmt_ <> newVal Then
        oldBillAmt_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get CorrectBillAmt() As Double
    CorrectBillAmt = correctBillAmt_
End Property
Public Property Let CorrectBillAmt(ByVal newVal As Double)
    If correctBillAmt_ <> newVal Then
        correctBillAmt_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get VATRate() As Double
    VATRate = vatRate_
End Property
Public Property Let VATRate(ByVal newVal As Double)
    If vatRate_ <> newVal Then
        vatRate_ = newVal
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
    sqlStr = "SELECT * FROM tbAdjustDeptHeader WHERE AdjustID='" & id_ & "'"
    Set tmpRec = New ADODB.Recordset
    GetDBConnection
    tmpRec.open sqlStr, DBConn, adOpenDynamic, adLockOptimistic
    If tmpRec.EOF Then
        tmpRec.AddNew
        tmpRec("AdjustID") = id_
    End If
    tmpRec("IssuedDate") = Trim(issuedDate_)
    tmpRec("CompanyID") = Trim(companyID_)
    tmpRec("CustomerID") = Trim(CustomerID_)
    tmpRec("AdjReason") = Trim(adjReason_)
    tmpRec("IssuedStaffID") = Trim(issuedStaffID_)
    tmpRec("RecBillPerson") = Trim(recBillPerson_)
    tmpRec("AdjType") = Trim(AdjType_)
    tmpRec("OldBillAmt") = Trim(oldBillAmt_)
    tmpRec("CustBillAddrID") = Trim(billAddrID_)
    tmpRec("CorrectBillAmt") = Trim(correctBillAmt_)
    tmpRec("AdjVATRate") = Trim(vatRate_)
    If isCancel_ = "" Then isCancel_ = "N"
    tmpRec("isCancel") = Trim(isCancel_)
    tmpRec("CancelStaffID") = Trim(cancelStaffID_)
    tmpRec.Update
    tmpRec.Requery
    AdjustBillNo_ = Trim("" & tmpRec("AdjustBillNo"))
'    tmpRec.Close
    
    DBConnExc "Delete tbAdjustDeptDetail WHERE AdjustID='" & id_ & "'"
    For x = 1 To colAdjDet_.Count
        colAdjDet_(x).Save
    Next
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
    DBConnExc "Delete AdjustDetID WHERE AdjustID='" & id_ & "'"
    DBConnExc "Delete tbAdjustDeptHeader WHERE AdjustID='" & id_ & "'"
    Delete = ""
    Exit Function
delErr:
    Delete = err.Description
End Function

Public Sub SetDataFromRecordset(selRec As ADODB.Recordset)
    id_ = Trim("" & selRec("AdjustID"))
    issuedDate_ = Trim("" & selRec("IssuedDate"))
    AdjustBillNo_ = Trim("" & selRec("AdjustBillNo"))
    companyID_ = Trim("" & selRec("CompanyID"))
    companyName_ = Trim("" & selRec("CompanyNameTH"))
    CustomerID_ = Trim("" & selRec("CustomerID"))
    customerName_ = Trim("" & selRec("CustomerName"))
    adjReason_ = Trim("" & selRec("AdjReason"))
    issuedStaffID_ = Trim("" & selRec("IssuedStaffID"))
    issuedStaffName_ = Trim("" & selRec("StaffFName")) & " " & Trim("" & selRec("StaffLName"))
    recBillPerson_ = Trim("" & selRec("RecBillPerson"))
    AdjType_ = Trim("" & selRec("AdjType"))
    oldBillAmt_ = selRec("OldBillAmt")
    correctBillAmt_ = selRec("CorrectBillAmt")
    vatRate_ = selRec("AdjVATRate")
    isCancel_ = selRec("isCancel")
    cancelStaffID_ = selRec("CancelStaffID")
    billAddrID_ = Trim("" & selRec("CustBillAddrID"))
    billAddr_ = Trim("" & selRec("Addr1")) & " " & Trim("" & selRec("TumbolName")) & " " & Trim("" & selRec("AumphurName")) & " " & Trim("" & selRec("ProvinceName")) & " " & Trim("" & selRec("ZipCode"))
End Sub

Public Property Get AdjDetails() As Collection
Dim tmpS As New BWGSearchManager
    If colAdjDet_ Is Nothing Then Set colAdjDet_ = tmpS.AdjustDeptDetailSearch("AdjustID='" & id_ & "'")
    Set tmpS = Nothing
    Set AdjDetails = colAdjDet_
End Property
Public Property Set AdjDetails(ByVal newVal As Collection)
    Set colAdjDet_ = newVal
End Property
Public Sub RequeryAdjDetail()
    Set colAdjDet_ = Nothing
End Sub

Public Sub CancelBill()
    DBConnExc "Update tbAdjustDeptHeader Set isCancel='Y' WHERE AdjustID='" & id_ & "'"
    isCancel_ = "Y"
    cancelStaffID_ = CurrentUser.ID
    isChanged_ = False
End Sub

