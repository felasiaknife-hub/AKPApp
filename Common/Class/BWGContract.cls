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
Dim StartDate_  As String
Dim EndDate_ As String
Dim remDesc_ As String
Dim DocNo_ As String
Dim CustomerID_ As String
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
    StartDate = StartDate_
End Property
Public Property Let StartDate(ByVal newVal As String)
    If Trim(StartDate_) <> Trim(newVal) Then
        StartDate_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get CustomerID() As String
    CustomerID = CustomerID_
End Property
Public Property Let CustomerID(ByVal newVal As String)
    If Trim(CustomerID_) <> Trim(newVal) Then
        CustomerID_ = Trim(newVal)
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

Public Property Get CompanyID() As String
    CompanyID = companyID_
End Property
Public Property Let CompanyID(ByVal newVal As String)
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
    EndDate = EndDate_
End Property
Public Property Let EndDate(ByVal newVal As String)
    If Trim(EndDate_) <> Trim(newVal) Then
        EndDate_ = Trim(newVal)
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
    docNo = DocNo_
End Property
Public Property Let docNo(ByVal newVal As String)
    If Trim(DocNo_) <> Trim(newVal) Then
        DocNo_ = Trim(newVal)
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
    sqlStr = "SELECT *, CONVERT(DATETIME, GETDATE(), 103) AS GETDATE FROM tbCustomerContract WHERE ContractID='" & id_ & "'"
    Set tmpRec = New ADODB.Recordset
    GetDBConnection
    tmpRec.open sqlStr, DBConn, adOpenDynamic, adLockOptimistic
    If tmpRec.EOF Then
        tmpRec.AddNew
        tmpRec("ContractID") = id_
    End If
    tmpRec("StartDate") = Trim(StartDate_)
    tmpRec("EndDate") = Trim(EndDate_)
    tmpRec("RemDesc") = Trim(remDesc_)
    tmpRec("DocNo") = Trim(DocNo_)
    tmpRec("CustomerID") = Trim(CustomerID_)
    tmpRec("CompanyID") = Trim(companyID_)
    
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
    DBConnExc "Delete tbCustomerContract WHERE ContractID='" & id_ & "'"
    Delete = ""
    Exit Function
delErr:
    Delete = err.Description
End Function

Public Sub SetDataFromRecordset(selRec As ADODB.Recordset)
    id_ = selRec("ContractID")
    StartDate_ = Trim("" & selRec("StartDate"))
    EndDate_ = Trim("" & selRec("EndDate"))
    remDesc_ = Trim("" & selRec("RemDesc"))
    DocNo_ = Trim("" & selRec("DocNo"))
    CustomerID_ = Trim("" & selRec("CustomerID"))
    customerName_ = Trim("" & selRec("CustomerName"))
    companyID_ = Trim("" & selRec("CompanyID"))
    companyName_ = Trim("" & selRec("CompanyNameTH"))
End Sub
