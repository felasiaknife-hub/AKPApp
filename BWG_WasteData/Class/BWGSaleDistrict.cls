VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BWGSaleDistrict"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim id_ As String
Dim zoneID_ As String
Dim zoneName_ As String
Dim districtName_ As String
Dim saleStaffID_ As String
Dim saleStaffName_ As String
Dim isChanged_ As Boolean

Public Property Get ID() As String
    ID = id_
End Property
Public Property Let ID(ByVal newVal As String)
    id_ = Trim(newVal)
End Property

Public Property Get ZoneID() As String
    ZoneID = zoneID_
End Property
Public Property Let ZoneID(ByVal newVal As String)
    If Trim(zoneID_) <> Trim(newVal) Then
        zoneID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get ZoneName() As String
    ZoneName = zoneName_
End Property
Public Property Let ZoneName(ByVal newVal As String)
    If Trim(zoneName_) <> Trim(newVal) Then
        zoneName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get DistrictName() As String
    DistrictName = districtName_
End Property
Public Property Let DistrictName(ByVal newVal As String)
    If Trim(districtName_) <> Trim(newVal) Then
        districtName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get SaleStaffID() As String
    SaleStaffID = saleStaffID_
End Property
Public Property Let SaleStaffID(ByVal newVal As String)
    If Trim(saleStaffID_) <> Trim(newVal) Then
        saleStaffID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get SaleStaffName() As String
    SaleStaffName = saleStaffName_
End Property
Public Property Let SaleStaffName(ByVal newVal As String)
    If Trim(saleStaffName_) <> Trim(newVal) Then
        saleStaffName_ = Trim(newVal)
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
    tmpRec.Open "Select * from tbSaleDistrict Where SaleDistrictID='" & id_ & "'", DBManager.GetDBConnection, adOpenDynamic, adLockOptimistic
    If tmpRec.EOF Then
        tmpRec.AddNew
        tmpRec("SaleDistrictID") = id_
    End If
    tmpRec("SaleZoneID") = Trim(zoneID_)
    tmpRec("SDName") = Trim(districtName_)
    tmpRec("SaleStaffID") = Trim(saleStaffID_)
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
    DBManager.GetDBConnection.Execute "Delete tbSaleDistrict Where SaleDistrictID='" & id_ & "'"
    Delete = ""
    Exit Function
delErr:
    Delete = Err.Description
End Function

Public Sub SetDataFromRecordset(selRec As ADODB.Recordset)
    id_ = selRec("SaleDistrictID")
    zoneID_ = Trim("" & selRec("SaleZoneID"))
    zoneName_ = selRec("ZoneName")
    districtName_ = Trim("" & selRec("SDName"))
    saleStaffID_ = selRec("SaleStaffID")
    saleStaffName_ = Trim("" & selRec("StaffFName") & " " & selRec("StaffLName"))
End Sub
