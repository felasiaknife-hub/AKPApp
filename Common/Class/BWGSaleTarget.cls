VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BWGSaleTarget"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim id_ As String
Dim saleDistrictID_ As String
Dim saleDistrictName_  As String
Dim saleZoneName_ As String
Dim saleStaffID_ As String
Dim SaleStaffName_ As String
Dim forMonth_ As Integer
Dim forYear_ As Integer
Dim nonTarget_ As Double
Dim hzTarget_ As Double
Dim nblTarget_ As Double
Dim hblTarget_ As Double
Dim HICTarget_ As Double
Dim HBITarget_ As Double
Dim BBHTarget_ As Double
Dim BBLTarget_ As Double
Dim KPITarget_ As Double
Dim companyID_ As String
Dim companyName_ As String
Dim LastUpdate_ As String
Dim InPro_ As Integer
Dim isOverAll_ As Integer
Dim isChanged_ As Boolean

Public Property Get ID() As String
    ID = id_
End Property
Public Property Let ID(ByVal newval As String)
    id_ = Trim(newval)
End Property

Public Property Get SaleDistrictID() As String
    SaleDistrictID = saleDistrictID_
End Property
Public Property Let SaleDistrictID(ByVal newval As String)
    If Trim(saleDistrictID_) <> Trim(newval) Then
        saleDistrictID_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get SaleDistrictName() As String
    SaleDistrictName = saleDistrictName_
End Property
Public Property Let SaleDistrictName(ByVal newval As String)
    If Trim(saleDistrictName_) <> Trim(newval) Then
        saleDistrictName_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get SaleZoneName() As String
    SaleZoneName = saleZoneName_
End Property
Public Property Let SaleZoneName(ByVal newval As String)
    If Trim(saleZoneName_) <> Trim(newval) Then
        saleZoneName_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get SaleStaffID() As String
    SaleStaffID = saleStaffID_
End Property
Public Property Let SaleStaffID(ByVal newval As String)
    If Trim(saleStaffID_) <> Trim(newval) Then
        saleStaffID_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get SaleStaffName() As String
    SaleStaffName = SaleStaffName_
End Property
Public Property Let SaleStaffName(ByVal newval As String)
    If Trim(SaleStaffName_) <> Trim(newval) Then
        SaleStaffName_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get ForMonth() As Integer
    ForMonth = forMonth_
End Property
Public Property Let ForMonth(ByVal newval As Integer)
    If forMonth_ <> newval Then
        forMonth_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get ForYear() As Integer
    ForYear = forYear_
End Property
Public Property Let ForYear(ByVal newval As Integer)
    If forYear_ <> newval Then
        forYear_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get NonTarget() As Double
    NonTarget = nonTarget_
End Property
Public Property Let NonTarget(ByVal newval As Double)
    If nonTarget_ <> newval Then
        nonTarget_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get HzTarget() As Double
    HzTarget = hzTarget_
End Property
Public Property Let HzTarget(ByVal newval As Double)
    If hzTarget_ <> newval Then
        hzTarget_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get NBLTarget() As Double
    NBLTarget = nblTarget_
End Property
Public Property Let NBLTarget(ByVal newval As Double)
    If nblTarget_ <> newval Then
        nblTarget_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get HBLTarget() As Double
    HBLTarget = hblTarget_
End Property
Public Property Let HBLTarget(ByVal newval As Double)
    If hblTarget_ <> newval Then
        hblTarget_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get HICTarget() As Double
    HICTarget = HICTarget_
End Property
Public Property Let HICTarget(ByVal newval As Double)
    If HICTarget_ <> newval Then
        HICTarget_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get KPITarget() As Double
    KPITarget = KPITarget_
End Property
Public Property Let KPITarget(ByVal newval As Double)
    If KPITarget_ <> newval Then
        KPITarget_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get HBITarget() As Double
    HBITarget = HBITarget_
End Property
Public Property Let HBITarget(ByVal newval As Double)
    If HBITarget_ <> newval Then
        HBITarget_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get BBHTarget() As Double
    BBHTarget = BBHTarget_
End Property
Public Property Let BBHTarget(ByVal newval As Double)
    If BBHTarget_ <> newval Then
        BBHTarget_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get BBLTarget() As Double
    BBLTarget = BBLTarget_
End Property
Public Property Let BBLTarget(ByVal newval As Double)
    If BBLTarget_ <> newval Then
        BBLTarget_ = newval
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

Public Property Get LastUpdate() As String
    LastUpdate = LastUpdate_
End Property
Public Property Let LastUpdate(ByVal newval As String)
    If Trim(LastUpdate_) <> Trim(newval) Then
        LastUpdate_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get InPro() As Integer
    InPro = InPro_
End Property
Public Property Let InPro(ByVal newval As Integer)
    If InPro_ <> newval Then
        InPro_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get IsOverall() As Integer
    IsOverall = isOverAll_
End Property
Public Property Let IsOverall(ByVal newval As Integer)
    If isOverAll_ <> newval Then
        isOverAll_ = newval
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
    sqlStr = "SELECT *, CONVERT(DATETIME, GETDATE(), 103) AS GETDATE FROM tbSaleTargetDistrict WHERE TargetID='" & id_ & "'"
    Set tmpRec = New ADODB.Recordset
    GetDBConnection
    tmpRec.open sqlStr, DBConn, adOpenDynamic, adLockOptimistic
    If tmpRec.EOF Then
        tmpRec.AddNew
        tmpRec("TargetID") = id_
    End If
    tmpRec("SaleDistrictID") = Trim(saleDistrictID_)
    tmpRec("SaleStaffID") = Trim(saleStaffID_)
    tmpRec("ForMonth") = Trim(forMonth_)
    tmpRec("ForYear") = Trim(forYear_)
    tmpRec("NonTarget") = Trim(nonTarget_)
    tmpRec("HazTarget") = Trim(hzTarget_)
    tmpRec("NBLTarget") = Trim(nblTarget_)
    tmpRec("HBLTarget") = Trim(hblTarget_)
    tmpRec("HICTarget") = Trim(HICTarget_)
    tmpRec("HBITarget") = Trim(HBITarget_)
    tmpRec("KPITarget") = Trim(KPITarget_)
    tmpRec("BBHTarget") = Trim(BBHTarget_)
    tmpRec("BBLTarget") = Trim(BBLTarget_)
    tmpRec("CompanyID") = Trim(companyID_)
    tmpRec("LastUpdateDate") = Trim(LastUpdate_)
    tmpRec("InPro") = Trim(InPro_)
    tmpRec("isOverAll") = Trim(isOverAll_)
    
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
    DBConnExc "Delete tbSaleTargetDistrict WHERE TargetID='" & id_ & "'"
    Delete = ""
    Exit Function
delErr:
    Delete = err.Description
End Function

Public Sub SetDataFromRecordset(selRec As ADODB.Recordset)
    id_ = Trim("" & selRec("TargetID"))
    saleDistrictID_ = Trim("" & selRec("SaleDistrictID"))
    saleStaffID_ = Trim("" & selRec("SaleStaffID"))
    saleDistrictName_ = Trim("" & selRec("SDName"))
    saleZoneName_ = Trim("" & selRec("ZoneName"))
    SaleStaffName_ = Trim("" & selRec("StaffFName") & " " & selRec("StaffLName"))
    If IsNull(selRec("ForMonth")) = False Then forMonth_ = selRec("ForMonth")
    If IsNull(selRec("ForYear")) = False Then forYear_ = selRec("ForYear")
    If IsNull(selRec("NonTarget")) = False Then nonTarget_ = selRec("NonTarget")
    If IsNull(selRec("HazTarget")) = False Then hzTarget_ = selRec("HazTarget")
    If IsNull(selRec("NBLTarget")) = False Then nblTarget_ = selRec("NBLTarget")
    If IsNull(selRec("HBLTarget")) = False Then hblTarget_ = selRec("HBLTarget")
    If IsNull(selRec("HICTarget")) = False Then HICTarget_ = selRec("HICTarget")
    If IsNull(selRec("HBITarget")) = False Then HBITarget_ = selRec("HBITarget")
    If IsNull(selRec("KPITarget")) = False Then KPITarget_ = selRec("KPITarget")
    If IsNull(selRec("BBHTarget")) = False Then BBHTarget_ = selRec("BBHTarget")
    If IsNull(selRec("BBLTarget")) = False Then BBLTarget_ = selRec("BBLTarget")
    companyID_ = Trim("" & selRec("CompanyID"))
    companyName_ = Trim("" & selRec("CompanyNameTH"))
    LastUpdate_ = Trim("" & selRec("LastUpdateDate"))
    InPro_ = Trim("" & selRec("InPro"))
    isOverAll_ = Trim("" & selRec("isOverAll"))
End Sub
