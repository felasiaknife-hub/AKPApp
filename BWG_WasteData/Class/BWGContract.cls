VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BWGContract"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim id_ As String
Dim startDate_  As String
Dim endDate_ As String
Dim remDesc_ As String
Dim docNo_ As String
Dim customerID_ As String
Dim customerName_ As String
Dim companyID_ As String
Dim companyName_ As String
Dim isChanged_ As Boolean

Public Property Get ID() As String
    ID = id_
End Property
Public Property Let ID(ByVal newVal As String)
    id_ = Trim(newVal)
End Property

Public Property Get StartDate() As String
    StartDate = startDate_
End Property
Public Property Let StartDate(ByVal newVal As String)
    If Trim(startDate_) <> Trim(newVal) Then
        startDate_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get CustomerID() As String
    CustomerID = customerID_
End Property
Public Property Let CustomerID(ByVal newVal As String)
    If Trim(customerID_) <> Trim(newVal) Then
        customerID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get CustomerName() As String
    CustomerName = customerName_
End Property
Public Property Let CustomerName(ByVal newVal As String)
    If Trim(customerName_) <> Trim(newVal) Then
        customerName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get companyID() As String
    companyID = companyID_
End Property
Public Property Let companyID(ByVal newVal As String)
    If Trim(companyID_) <> Trim(newVal) Then
        companyID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get CompanyName() As String
    CompanyName = companyName_
End Property
Public Property Let CompanyName(ByVal newVal As String)
    If Trim(companyName_) <> Trim(newVal) Then
        companyName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get EndDate() As String
    EndDate = endDate_
End Property
Public Property Let EndDate(ByVal newVal As String)
    If Trim(endDate_) <> Trim(newVal) Then
        endDate_ = Trim(newVal)
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

Public Property Get docNo() As String
    docNo = docNo_
End Property
Public Property Let docNo(ByVal newVal As String)
    If Trim(docNo_) <> Trim(newVal) Then
        docNo_ = Trim(newVal)
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
    Set tmpRec = New ADODB.Recordset
    tmpRec.Open "Select * from tbCustomerContract Where ContractID='" & id_ & "'", DBManager.GetDBConnection, adOpenDynamic, adLockOptimistic
    If tmpRec.EOF Then
        tmpRec.AddNew
        tmpRec("ContractID") = id_
    End If
    tmpRec("StartDate") = Trim(startDate_)
    tmpRec("EndDate") = Trim(endDate_)
    tmpRec("RemDesc") = Trim(remDesc_)
    tmpRec("DocNo") = Trim(docNo_)
    tmpRec("CustomerID") = Trim(customerID_)
    tmpRec("CompanyID") = Trim(companyID_)
    tmpRec.Update
    tmpRec.Close
    Set tmpRec = Nothing
    isChanged_ = False
    Exit Sub
ErrD:
    MsgBox Err.Description & vbCrLf & "กรุณาลองใหม่หรือติดต่อผู้พัฒนาระบบ", vbCritical, "Error"
    Exit Sub
End Sub

Public Function Delete() As String
On Error GoTo delErr:
    DBManager.GetDBConnection.Execute "Delete tbCustomerContract Where ContractID='" & id_ & "'"
    Delete = ""
    Exit Function
delErr:
    Delete = Err.Description
End Function

Public Sub SetDataFromRecordset(selRec As ADODB.Recordset)
    id_ = selRec("ContractID")
    startDate_ = Trim("" & selRec("StartDate"))
    endDate_ = Trim("" & selRec("EndDate"))
    remDesc_ = Trim("" & selRec("RemDesc"))
    docNo_ = Trim("" & selRec("DocNo"))
    customerID_ = Trim("" & selRec("CustomerID"))
    customerName_ = Trim("" & selRec("CustomerName"))
    companyID_ = Trim("" & selRec("CompanyID"))
    companyName_ = Trim("" & selRec("CompanyNameTH"))
End Sub
