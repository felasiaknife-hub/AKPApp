VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BWGCustomerStaff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim id_ As String
Dim CustomerID_ As String
Dim fName_ As String
Dim lName_ As String
Dim positionName_ As String
Dim deptName_ As String
Dim TelNo_ As String
Dim faxNo_ As String
Dim Email_ As String
Dim Subjact_ As String
Dim isChanged_ As Boolean
Dim Address_ As String


Public Property Get ID() As String
    ID = id_
End Property
Public Property Let ID(ByVal newval As String)
    id_ = Trim(newval)
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

Public Property Get FirstName() As String
    FirstName = fName_
End Property
Public Property Let FirstName(ByVal newval As String)
    If Trim(fName_) <> Trim(newval) Then
        fName_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get Email() As String
    Email = Email_
End Property
Public Property Let Email(ByVal newval As String)
    If Trim(Email_) <> Trim(newval) Then
        Email_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get Subjact() As String
    Subjact = Subjact_
End Property
Public Property Let Subjact(ByVal newval As String)
    If Trim(Subjact_) <> Trim(newval) Then
        Subjact_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get Address() As String
    Address = Address_
End Property
Public Property Let Address(ByVal newval As String)
    If Trim(Address_) <> Trim(newval) Then
        Address_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get LastName() As String
    LastName = lName_
End Property
Public Property Let LastName(ByVal newval As String)
    If Trim(lName_) <> Trim(newval) Then
        lName_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get PositionName() As String
    PositionName = positionName_
End Property
Public Property Let PositionName(ByVal newval As String)
    If Trim(positionName_) <> Trim(newval) Then
        positionName_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get DeptName() As String
    DeptName = deptName_
End Property
Public Property Let DeptName(ByVal newval As String)
    If Trim(deptName_) <> Trim(newval) Then
        deptName_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get TelNo() As String
    TelNo = TelNo_
End Property
Public Property Let TelNo(ByVal newval As String)
    If Trim(TelNo_) <> Trim(newval) Then
        TelNo_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get FaxNo() As String
    FaxNo = faxNo_
End Property
Public Property Let FaxNo(ByVal newval As String)
    If Trim(faxNo_) <> Trim(newval) Then
        faxNo_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get isChanged() As Boolean
    isChanged = isChanged_
End Property
Public Property Let isChanged(ByVal newval As Boolean)
    isChanged_ = newval
End Property

Public Sub Save()
On Error GoTo ErrD
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
    sqlStr = "SELECT *, CONVERT(DATETIME, GETDATE(), 103) AS GETDATE FROM tbCustomerStaff WHERE CustStaffID='" & id_ & "'"
    Set tmpRec = New ADODB.Recordset
    GetDBConnection
    tmpRec.open sqlStr, DBConn, adOpenDynamic, adLockOptimistic
    If tmpRec.EOF Then
        tmpRec.AddNew
        tmpRec("CustStaffID") = id_
    End If
    tmpRec("CustomerID") = Trim(CustomerID_)
    tmpRec("FName") = Trim(fName_)
    tmpRec("LName") = Trim(lName_)
    tmpRec("PositionName") = Trim(positionName_)
    tmpRec("DeptName") = Trim(deptName_)
    tmpRec("TelNo") = Trim(TelNo_)
    tmpRec("FaxNo") = Trim(faxNo_)
    tmpRec("Email") = Trim(Email_)
    tmpRec("Subjact") = Trim(Subjact_)
    tmpRec("Address") = Trim(Address_)
    
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
    DBConnExc "Delete tbCustomerStaff WHERE CustStaffID='" & id_ & "'"
    Delete = ""
    Exit Function
delErr:
    Delete = err.Description
End Function

Public Sub SetDataFromRecordset(selRec As ADODB.Recordset)
    id_ = Trim("" & selRec("CustStaffID"))
    CustomerID_ = Trim("" & selRec("CustomerID"))
    fName_ = Trim("" & selRec("FName"))
    lName_ = Trim("" & selRec("LName"))
    positionName_ = Trim("" & selRec("PositionName"))
    deptName_ = Trim("" & selRec("DeptName"))
    TelNo_ = Trim("" & selRec("TelNo"))
    faxNo_ = Trim("" & selRec("FaxNo"))
    Email_ = Trim("" & selRec("Email"))
    Subjact_ = Trim("" & selRec("Subjact"))
    Address_ = Trim("" & selRec("Address"))
End Sub
