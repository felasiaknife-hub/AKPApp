VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BWGTSDF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim id_ As String
Dim companyName_ As String
Dim tsdfRegisNo_ As String
Dim addr1_ As String
Dim tumbolID_ As String
Dim tumbolName_ As String
Dim aumphurID_ As String
Dim aumphurName_ As String
Dim provinceID_ As String
Dim provinceName_ As String
Dim zipCode_ As String
Dim telNo_ As String
Dim faxNo_ As String
Dim contactName_ As String
Dim canBlending_ As String
Dim canLandFill_ As String
Dim isChanged_ As Boolean


Public Property Get ID() As String
    ID = id_
End Property
Public Property Let ID(ByVal newVal As String)
    id_ = Trim(newVal)
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

Public Property Get TSDFRegisNo() As String
    TSDFRegisNo = tsdfRegisNo_
End Property
Public Property Let TSDFRegisNo(ByVal newVal As String)
    If Trim(tsdfRegisNo_) <> Trim(newVal) Then
        tsdfRegisNo_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get Addr1() As String
    Addr1 = addr1_
End Property
Public Property Let Addr1(ByVal newVal As String)
    If Trim(addr1_) <> Trim(newVal) Then
        addr1_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get TumbolID() As String
    TumbolID = tumbolID_
End Property
Public Property Let TumbolID(ByVal newVal As String)
    If Trim(tumbolID_) <> Trim(newVal) Then
        tumbolID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get TumbolName() As String
    TumbolName = tumbolName_
End Property
Public Property Let TumbolName(ByVal newVal As String)
    If Trim(tumbolName_) <> Trim(newVal) Then
        tumbolName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get AumphurID() As String
    AumphurID = aumphurID_
End Property
Public Property Let AumphurID(ByVal newVal As String)
    If Trim(aumphurID_) <> Trim(newVal) Then
        aumphurID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get AumphurName() As String
    AumphurName = aumphurName_
End Property
Public Property Let AumphurName(ByVal newVal As String)
    If Trim(aumphurName_) <> Trim(newVal) Then
        aumphurName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get ProvinceID() As String
    ProvinceID = provinceID_
End Property
Public Property Let ProvinceID(ByVal newVal As String)
    If Trim(provinceID_) <> Trim(newVal) Then
        provinceID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get ProvinceName() As String
    ProvinceName = provinceName_
End Property
Public Property Let ProvinceName(ByVal newVal As String)
    If Trim(provinceName_) <> Trim(newVal) Then
        provinceName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get ZipCode() As String
    ZipCode = zipCode_
End Property
Public Property Let ZipCode(ByVal newVal As String)
    If Trim(zipCode_) <> Trim(newVal) Then
        zipCode_ = Trim(newVal)
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

Public Property Get FaxNo() As String
    FaxNo = faxNo_
End Property
Public Property Let FaxNo(ByVal newVal As String)
    If Trim(faxNo_) <> Trim(newVal) Then
        faxNo_ = Trim(newVal)
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

Public Property Get CanBlending() As Boolean
    If Trim(canBlending_) = "Y" Then
        CanBlending = True
    Else
        CanBlending = False
    End If
End Property
Public Property Let CanBlending(ByVal newVal As Boolean)
Dim tmpStr As String
    If newVal Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(canBlending_) <> Trim(tmpStr) Then
        canBlending_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property

Public Property Get CanLandFill() As Boolean
    If Trim(canLandFill_) = "Y" Then
        CanLandFill = True
    Else
        CanLandFill = False
    End If
End Property
Public Property Let CanLandFill(ByVal newVal As Boolean)
Dim tmpStr As String
    If newVal Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(canLandFill_) <> Trim(tmpStr) Then
        canLandFill_ = Trim(tmpStr)
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
    tmpRec.Open "Select * from tbTSDFCompany Where TSDFID='" & id_ & "'", DBManager.GetDBConnection, adOpenDynamic, adLockOptimistic
    If tmpRec.EOF Then
        tmpRec.AddNew
        tmpRec("TSDFID") = id_
    End If
    tmpRec("CompanyName") = Trim(companyName_)
    tmpRec("TSDFRegisNo") = Trim(tsdfRegisNo_)
    tmpRec("Addr1") = Trim(addr1_)
    tmpRec("TumbolID") = Trim(tumbolID_)
    tmpRec("ZipCode") = Trim(zipCode_)
    tmpRec("TelNo") = Trim(telNo_)
    tmpRec("FaxNo") = Trim(faxNo_)
    tmpRec("ContactName") = Trim(contactName_)
    tmpRec("CanBlending") = Trim(canBlending_)
    tmpRec("CanLandFill") = Trim(canLandFill_)
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
    DBManager.GetDBConnection.Execute "Delete tbTSDFCompany Where TSDFID='" & id_ & "'"
    Delete = ""
    Exit Function
delErr:
    Delete = Err.Description
End Function

Public Sub SetDataFromRecordset(selRec As ADODB.Recordset)
    id_ = Trim("" & selRec("TSDFID"))
    companyName_ = Trim("" & selRec("CompanyName"))
    tsdfRegisNo_ = Trim("" & selRec("TSDFRegisNo"))
    addr1_ = Trim("" & selRec("Addr1"))
    tumbolID_ = Trim("" & selRec("TumbolID"))
    tumbolName_ = Trim("" & selRec("TumbolName"))
    aumphurID_ = Trim("" & selRec("AumphurID"))
    aumphurName_ = Trim("" & selRec("AumphurName"))
    provinceID_ = Trim("" & selRec("ProvinceID"))
    provinceName_ = Trim("" & selRec("ProvinceName"))
    zipCode_ = Trim("" & selRec("ZipCode"))
    telNo_ = Trim("" & selRec("TelNo"))
    faxNo_ = Trim("" & selRec("FaxNo"))
    contactName_ = Trim("" & selRec("ContactName"))
    canBlending_ = Trim("" & selRec("CanBlending"))
    canLandFill_ = Trim("" & selRec("CanLandFill"))
End Sub
