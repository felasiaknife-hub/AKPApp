VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BWGCustomerAuction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim id_ As String
Dim StartDate_  As String
Dim EndDate_ As String
Dim Remdesc_ As String
Dim DocNo_ As String
Dim CustomerID_ As String
Dim CustomerName_ As String
Dim companyID_ As String
Dim companyName_ As String
Dim isChanged_ As Boolean
Dim IsOK_ As String
Dim Price_ As String
Dim NextDate_ As String

Public Property Get ID() As String
    ID = id_
End Property
Public Property Let ID(ByVal newval As String)
    id_ = Trim(newval)
End Property

Public Property Get StartDate() As String
    StartDate = StartDate_
End Property
Public Property Let StartDate(ByVal newval As String)
    If Trim(StartDate_) <> Trim(newval) Then
        StartDate_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get CustomerID() As String
    CustomerID = CustomerID_
End Property
Public Property Let CustomerID(ByVal newval As String)
    If Trim(CustomerID_) <> Trim(newval) Then
        CustomerID_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get CustomerName() As String
    CustomerName = CustomerName_
End Property
Public Property Let CustomerName(ByVal newval As String)
    If Trim(CustomerName_) <> Trim(newval) Then
        CustomerName_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get CompanyID() As String
    CompanyID = companyID_
End Property
Public Property Let CompanyID(ByVal newval As String)
    If Trim(companyID_) <> Trim(newval) Then
        companyID_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get CompanyName() As String
    CompanyName = companyName_
End Property
Public Property Let CompanyName(ByVal newval As String)
    If Trim(companyName_) <> Trim(newval) Then
        companyName_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get EndDate() As String
    EndDate = EndDate_
End Property
Public Property Let EndDate(ByVal newval As String)
    If Trim(EndDate_) <> Trim(newval) Then
        EndDate_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get NextDate() As String
    NextDate = NextDate_
End Property
Public Property Let NextDate(ByVal newval As String)
    If Trim(NextDate_) <> Trim(newval) Then
        NextDate_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get Remdesc() As String
    Remdesc = Remdesc_
End Property
Public Property Let Remdesc(ByVal newval As String)
    If Trim(Remdesc_) <> Trim(newval) Then
        Remdesc_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get docNo() As String
    docNo = DocNo_
End Property
Public Property Let docNo(ByVal newval As String)
    If Trim(DocNo_) <> Trim(newval) Then
        DocNo_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get Price() As String
    Price = Price_
End Property
Public Property Let Price(ByVal newval As String)
    If Trim(Price_) <> Trim(newval) Then
        Price_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get IsOK() As String
    IsOK = IsOK_
End Property
Public Property Let IsOK(ByVal newval As String)
    If Trim(IsOK_) <> Trim(newval) Then
        IsOK_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get isChanged() As Boolean
    isChanged = isChanged_
End Property
Public Property Let isChanged(ByVal newval As Boolean)
    isChanged_ = newval
End Property

Public Sub Save(ByVal AuctionID As String)
On Error GoTo ErrD
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
If AuctionID <> "" Then
    sqlStr = "SELECT *, CONVERT(DATETIME, GETDATE(), 103) AS GETDATE FROM tbCustomerAuction WHERE AuctionID=" & AuctionID & " "
    Else
    sqlStr = "SELECT *, CONVERT(DATETIME, GETDATE(), 103) AS GETDATE FROM tbCustomerAuction WHERE AuctionID=" & id_ & " "
End If
    Set tmpRec = New ADODB.Recordset
    GetDBConnection
    tmpRec.open sqlStr, DBConn, adOpenDynamic, adLockOptimistic
    If tmpRec.EOF Then
        tmpRec.AddNew
        tmpRec("AuctionID") = id_
    End If
    tmpRec("StartDate") = Trim(StartDate_)
    If IsOK_ = "Y" Then
    tmpRec("EndDate") = Trim(EndDate_)
    End If
    tmpRec("RemDesc") = Trim(Remdesc_)
    tmpRec("DocNo") = Trim(DocNo_)
    tmpRec("CustomerID") = Trim(CustomerID_)
    tmpRec("CompanyID") = Trim(companyID_)
    tmpRec("Price") = Trim(Price_)
    tmpRec("IsOK") = Trim(IsOK_)
    tmpRec("NextDate") = Trim(NextDate_)
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

Public Function Delete(ByVal AuctionID As String) As String
On Error GoTo delErr:
    DBConnExc "Delete tbCustomerAuction WHERE AuctionID=" & AuctionID & " "
    Delete = ""
    Exit Function
delErr:
    Delete = err.Description
End Function

Public Sub SetDataFromRecordset(selRec As ADODB.Recordset)
    id_ = selRec("AuctionID")
    StartDate_ = Trim("" & selRec("StartDate"))
    EndDate_ = Trim("" & selRec("EndDate"))
    Remdesc_ = Trim("" & selRec("RemDesc"))
    DocNo_ = Trim("" & selRec("DocNo"))
    CustomerID_ = Trim("" & selRec("CustomerID"))
    CustomerName_ = Trim("" & selRec("CustomerName"))
    companyID_ = Trim("" & selRec("CompanyID"))
    companyName_ = Trim("" & selRec("CompanyNameTH"))
    Price_ = Trim("" & selRec("Price"))
    IsOK_ = Trim("" & selRec("IsOK"))
    NextDate_ = Trim("" & selRec("NextDate"))
End Sub

