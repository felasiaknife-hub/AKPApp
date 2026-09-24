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
Dim ContactName_ As String
Dim positionName_ As String
Dim TelNo_ As String
Dim emailAddr_ As String
Dim carType_ As String
Dim isDriver_ As String
Dim isChanged_ As Boolean
Dim DVNumber_ As String
Dim DVType_ As String
Dim DVDateExp_ As String
Dim DVPNumber_ As String
Dim Stat_ As String

Public Property Get DVPNumber() As String
    DVPNumber = DVPNumber_
End Property
Public Property Let DVPNumber(ByVal newVal As String)
    If Trim(DVPNumber_) <> Trim(newVal) Then
        DVPNumber_ = Trim(newVal)
        isChanged_ = True
    End If
End Property
Public Property Get DVDateExp() As String
    DVDateExp = DVDateExp_
End Property
Public Property Let DVDateExp(ByVal newVal As String)
    If Trim(DVDateExp_) <> Trim(newVal) Then
        DVDateExp_ = Trim(newVal)
        isChanged_ = True
    End If
End Property
Public Property Get DVType() As String
    DVType = DVType_
End Property
Public Property Let DVType(ByVal newVal As String)
    If Trim(DVType_) <> Trim(newVal) Then
        DVType_ = Trim(newVal)
        isChanged_ = True
    End If
End Property
Public Property Get DVNumber() As String
    DVNumber = DVNumber_
End Property
Public Property Let DVNumber(ByVal newVal As String)
    If Trim(DVNumber_) <> Trim(newVal) Then
        DVNumber_ = Trim(newVal)
        isChanged_ = True
    End If
End Property
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
    ContactName = ContactName_
End Property
Public Property Let ContactName(ByVal newVal As String)
    If Trim(ContactName_) <> Trim(newVal) Then
        ContactName_ = Trim(newVal)
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
    TelNo = TelNo_
End Property
Public Property Let TelNo(ByVal newVal As String)
    If Trim(TelNo_) <> Trim(newVal) Then
        TelNo_ = Trim(newVal)
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
Public Property Get Stat() As String
    Stat = Stat_
End Property
Public Property Let Stat(ByVal newVal As String)
    If Trim(Stat_) <> Trim(newVal) Then
        Stat_ = Trim(newVal)
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

Public Sub Save(Optional ByVal MarkForDelete As String = "N")
On Error GoTo ErrD
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
    sqlStr = "SELECT *, CONVERT(DATETIME, GETDATE(), 103) AS GETDATE FROM tbSupplierContact WHERE ContactID='" & id_ & "'"
    Set tmpRec = New ADODB.Recordset
    GetDBConnection
    tmpRec.open sqlStr, DBConn, adOpenDynamic, adLockOptimistic
    If tmpRec.EOF Then
        tmpRec.AddNew
        tmpRec("ContactID") = id_
    End If
    tmpRec("VendorID") = Trim(venderID_)
    tmpRec("ContactName") = Trim(ContactName_)
    tmpRec("PositionName") = Trim(positionName_)
    tmpRec("TelNo") = Trim(TelNo_)
    tmpRec("EmailAddr") = Trim(emailAddr_)
    tmpRec("CarType") = Trim(carType_)
    tmpRec("isDriver") = Trim(isDriver_)
    tmpRec("DVNumber") = Trim(DVNumber_)
    tmpRec("DVType") = Trim(DVType_)
    tmpRec("DVDateExp") = Trim(DVDateExp_)
    tmpRec("DVPNumber") = Trim(DVPNumber_)
    tmpRec!LastUser = Trim(Trim(CurrentUser.StaffFName) & " " & Trim(CurrentUser.StaffLName))
    tmpRec!LastUpdate = GetSystemDateTime.SystemDateTime
    tmpRec!MarkForDelete = MarkForDelete
    tmpRec!Stat = Stat
    
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
    DBConnExc "Delete tbSupplierContact WHERE ContactID='" & id_ & "'"
    Delete = ""
    Exit Function
delErr:
    Delete = err.Description
End Function

Public Sub SetDataFromRecordset(selRec As ADODB.Recordset)
    id_ = selRec("ContactID")
    venderID_ = Trim("" & selRec("VendorID"))
    ContactName_ = Trim("" & selRec("ContactName"))
    positionName_ = Trim("" & selRec("PositionName"))
    TelNo_ = Trim("" & selRec("TelNo"))
    emailAddr_ = Trim("" & selRec("EmailAddr"))
    carType_ = Trim("" & selRec("CarType"))
    isDriver_ = Trim("" & selRec("isDriver"))
    DVNumber_ = Trim("" & selRec("DVNumber"))
    DVType_ = Trim("" & selRec("DVType"))
    DVDateExp_ = Trim("" & selRec("DVDateExp"))
    DVPNumber_ = Trim("" & selRec("DVPNumber"))
    Stat_ = Trim("" & selRec("Stat"))
End Sub
',

