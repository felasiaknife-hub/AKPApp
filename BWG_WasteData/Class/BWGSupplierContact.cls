VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BWGSupplierContact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim id_ As String
Dim venderID_ As String
Dim contactName_ As String
Dim positionName_ As String
Dim telNo_ As String
Dim emailAddr_ As String
Dim carType_ As String
Dim isDriver_ As String
Dim isChanged_ As Boolean

Public Property Get ID() As String
    ID = id_
End Property
Public Property Let ID(ByVal newVal As String)
    id_ = Trim(newVal)
End Property

Public Property Get VendorID() As String
    VendorID = venderID_
End Property
Public Property Let VendorID(ByVal newVal As String)
    If Trim(venderID_) <> Trim(newVal) Then
        venderID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get ContactName() As String
    ContactName = contactName_
End Property
Public Property Let ContactName(ByVal newVal As String)
    If Trim(contactName_) <> Trim(newVal) Then
        contactName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get PositionName() As String
    PositionName = positionName_
End Property
Public Property Let PositionName(ByVal newVal As String)
    If Trim(positionName_) <> Trim(newVal) Then
        positionName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get TelNo() As String
    TelNo = telNo_
End Property
Public Property Let TelNo(ByVal newVal As String)
    If Trim(telNo_) <> Trim(newVal) Then
        telNo_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get EmailAddr() As String
    EmailAddr = emailAddr_
End Property
Public Property Let EmailAddr(ByVal newVal As String)
    If Trim(emailAddr_) <> Trim(newVal) Then
        emailAddr_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get CarType() As String
    CarType = carType_
End Property
Public Property Let CarType(ByVal newVal As String)
    If Trim(carType_) <> Trim(newVal) Then
        carType_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get isDriver() As Boolean
    If Trim(isDriver_) = "" Then Trim(isDriver_) = "Y"
    If Trim(isDriver_) = "Y" Then
        isDriver = True
    Else
        isDriver = False
    End If
End Property
Public Property Let isDriver(ByVal newVal As Boolean)
Dim tmpStr As String
    If newVal Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(isDriver_) <> Trim(tmpStr) Then
        isDriver_ = Trim(tmpStr)
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
    tmpRec.Open "Select * from tbSupplierContact Where ContactID='" & id_ & "'", DBManager.GetDBConnection, adOpenDynamic, adLockOptimistic
    If tmpRec.EOF Then
        tmpRec.AddNew
        tmpRec("ContactID") = id_
    End If
    tmpRec("VendorID") = Trim(venderID_)
    tmpRec("ContactName") = Trim(contactName_)
    tmpRec("PositionName") = Trim(positionName_)
    tmpRec("TelNo") = Trim(telNo_)
    tmpRec("EmailAddr") = Trim(emailAddr_)
    tmpRec("CarType") = Trim(carType_)
    tmpRec("isDriver") = Trim(isDriver_)
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
    DBManager.GetDBConnection.Execute "Delete tbSupplierContact Where ContactID='" & id_ & "'"
    Delete = ""
    Exit Function
delErr:
    Delete = Err.Description
End Function

Public Sub SetDataFromRecordset(selRec As ADODB.Recordset)
    id_ = selRec("ContactID")
    venderID_ = Trim("" & selRec("VendorID"))
    contactName_ = Trim("" & selRec("ContactName"))
    positionName_ = Trim("" & selRec("PositionName"))
    telNo_ = Trim("" & selRec("TelNo"))
    emailAddr_ = Trim("" & selRec("EmailAddr"))
    carType_ = Trim("" & selRec("CarType"))
    isDriver_ = Trim("" & selRec("isDriver"))
End Sub
',

