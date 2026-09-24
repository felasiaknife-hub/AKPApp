Attribute VB_Name = "modShared"
Option Explicit
Public colUnit_ As Collection
Public colReception_ As Collection
Public colCancelRemark_ As Collection
Public coltruckType_ As Collection

Public WasteTypes As Collection
Public WasteCases As Collection
Public ColFormList As Collection
Public ColUserGroup As Collection

Public AutoCalcMF As Boolean
Public cboIndex As Integer

Private colZone_ As Collection
Private colStaffInfo_ As Collection
Private colBankAccInfo_ As Collection
Private colVenderContact_ As Collection
Private colRem_ As Collection

Public Type TruckAvailableSet
    TruckTotal  As Integer
    TruckAvailable As Integer
    TruckUsed As Integer
End Type
Public TruckAvailable As TruckAvailableSet

Public Function TruckAvailableCount(ByVal WorkDate As String, ByVal SubTypeID1 As String, Optional ByVal SubTypeID2 As String)
    On Error GoTo ErrD
    Dim rs As New ADODB.Recordset
    Dim tSql As String, i As Integer, SubTypeID As String
    TruckAvailableCount = 0
    
    SubTypeID = "'" & Replace(SubTypeID1, "''", "") & "'"
    If Replace(SubTypeID2, "''", "") <> "" Then
        SubTypeID = "'" & Replace(SubTypeID1, "''", "") & "', '" & Replace(SubTypeID2, "''", "") & "'"
    End If
    
    tSql = "SELECT     TOP (100) PERCENT CONVERT(DATETIME, CAST(YEAR(WorkDate) - 543 AS VARCHAR(4)) + '/' + CAST(MONTH(WorkDate) AS VARCHAR(2))" & vbCrLf
    tSql = tSql & "                      + '/' + CAST(DAY(WorkDate) AS VARCHAR(2)), 111) AS WorkDate, ISNULL" & vbCrLf
    tSql = tSql & "                          ((SELECT     TOP (1) SUM(ISNULL(TruckTotal, 0)) AS Expr1" & vbCrLf
    tSql = tSql & "                              FROM         dbo.tbTruckTotalAvailable AS tbTruckTotalAvailable_2" & vbCrLf
    tSql = tSql & "                              WHERE     (SubTypeID IN (" & SubTypeID & ")) AND" & vbCrLf
    tSql = tSql & "                                                    (dbo.vw_JobData_Truck_Count.WorkDate >= TruckAvailableDate)" & vbCrLf
    tSql = tSql & "                              GROUP BY TruckAvailableDate" & vbCrLf
    tSql = tSql & "                              ORDER BY TruckAvailableDate DESC), 0) AS TruckTotal, ISNULL" & vbCrLf
    tSql = tSql & "                          ((SELECT     TOP (1) SUM(ISNULL(TruckTotal, 0)) AS Expr1" & vbCrLf
    tSql = tSql & "                              FROM         dbo.tbTruckTotalAvailable AS tbTruckTotalAvailable_2" & vbCrLf
    tSql = tSql & "                              WHERE     (SubTypeID IN (" & SubTypeID & ")) AND" & vbCrLf
    tSql = tSql & "                                                    (dbo.vw_JobData_Truck_Count.WorkDate >= TruckAvailableDate)" & vbCrLf
    tSql = tSql & "                              GROUP BY TruckAvailableDate" & vbCrLf
    tSql = tSql & "                              ORDER BY TruckAvailableDate DESC), 0) - ISNULL" & vbCrLf
    tSql = tSql & "                          ((SELECT     TOP (1) SUM(ISNULL(TruckNotAvailable, 0)) AS Expr1" & vbCrLf
    tSql = tSql & "                              From dbo.tbTruckManage" & vbCrLf
    tSql = tSql & "                              WHERE     (SubTypeID IN (" & SubTypeID & ")) AND" & vbCrLf
    tSql = tSql & "                                                    (dbo.vw_JobData_Truck_Count.WorkDate BETWEEN TruckNotAvailableDateStart AND TruckNotAvailableDateEnd)), 0) AS TruckAvailable," & vbCrLf
    tSql = tSql & "                      IsNull" & vbCrLf
    tSql = tSql & "                          ((SELECT     TOP (1) SUM(ISNULL(TruckNotAvailable, 0)) AS Expr1" & vbCrLf
    tSql = tSql & "                              FROM         dbo.tbTruckManage AS tbTruckManage_1" & vbCrLf
    tSql = tSql & "                              WHERE     (SubTypeID IN (" & SubTypeID & ")) AND" & vbCrLf
    tSql = tSql & "                                                    (dbo.vw_JobData_Truck_Count.WorkDate BETWEEN TruckNotAvailableDateStart AND TruckNotAvailableDateEnd)), 0)" & vbCrLf
    tSql = tSql & "                      AS TruckNotAvailable, ISNULL" & vbCrLf
    tSql = tSql & "                          ((SELECT     SUM(ISNULL(TruckCount, 0)) AS Expr1" & vbCrLf
    tSql = tSql & "                              FROM         dbo.vw_JobData_Truck_Count AS V1" & vbCrLf
    tSql = tSql & "                              WHERE     (SubTypeID IN (" & SubTypeID & ")) AND" & vbCrLf
    tSql = tSql & "                                                    (WorkDate = dbo.vw_JobData_Truck_Count.WorkDate) AND (ISNULL(isCancel, 'N') <> 'Y')), 0) AS TruckUsed, TruckTypeDesc," & vbCrLf
    tSql = tSql & "                      SubTypeDesc" & vbCrLf
    tSql = tSql & "From dbo.vw_JobData_Truck_Count" & vbCrLf
    tSql = tSql & "WHERE     (WorkDate = '" & WorkDate & "') AND (ISNULL(isCancel, 'N') <> 'Y') AND (SubTypeID = '" & SubTypeID1 & "')" & vbCrLf
    tSql = tSql & "GROUP BY WorkDate, TruckTypeDesc, SubTypeDesc, TruckTypeID, SubTypeID" & vbCrLf
    tSql = tSql & "ORDER BY TruckTypeDesc, SubTypeDesc"

'    Screen.MousePointer = 11 ' *** Set Sceeen Mouse for wait
    Set rs = GetRS(tSql, adOpenForwardOnly, adLockReadOnly)
    If rs.RecordCount = 0 Then
        TruckAvailableCount = 0
        Exit Function
    End If
    
    TruckAvailableCount = IIf(IsNull(rs!TruckAvailable), 0, rs!TruckAvailable) - IIf(IsNull(rs!TruckUsed), 0, rs!TruckUsed)
    
    TruckAvailable.TruckTotal = IIf(IsNull(rs!TruckTotal), 0, rs!TruckTotal)
    TruckAvailable.TruckUsed = IIf(IsNull(rs!TruckTotal), 0, rs!TruckUsed)
    TruckAvailable.TruckAvailable = TruckAvailableCount
'    Screen.MousePointer = 0 ' *** Set Sceeen Mouse for normal
    Exit Function
ErrD:
    MsgBox "เกิดข้อผิดพลาด ไม่สามารถค้นหาจำนวนรถที่เหลือได้"
    TruckAvailableCount = 1000
'    Screen.MousePointer = 0 ' *** Set Sceeen Mouse for normal
    Exit Function
End Function

Public Sub PopulateCarType(ByRef cbo As Object)
Dim x%, tmpS As New BWGSearchManager
    cbo.Clear
    Set coltruckType_ = tmpS.TruckTypeSearch
    cbo.AddItem "==เลือกทุกประเภทรถ=="
    For x = 1 To coltruckType_.Count
        cbo.AddItem coltruckType_(x).TypeDesc
    Next
    If cbo.ListCount > 0 Then
        cbo.ListIndex = 0
    End If
    Set tmpS = Nothing
End Sub
Public Sub PopulateUnit()
On Error Resume Next
Dim tmpS As New BWGSearchManager
    Set colUnit_ = tmpS.UnitGeneralSearch
    Set tmpS = Nothing
End Sub
Public Sub PopulateCompany(ByRef cbo As Object, ByRef cboIndex As Integer)
On Error Resume Next
Dim x%
    cbo.Clear
    If CurrentUser.getUserCompany().Count > 1 Then
        cbo.AddItem "==เลือกทุกบริษัท=="
        cboIndex = 0
    Else
        cboIndex = 1
    End If
    For x = 1 To CurrentUser.WorkCompanies.Count
        cbo.AddItem CurrentUser.WorkCompanies(x).CompanyNameTH
    Next
    If cbo.ListCount > 0 Then
        cbo.ListIndex = 1
    End If
End Sub
Public Sub PopulateReception() ' ปิดสำหรับ WasteData
On Error Resume Next
Dim tmpS As New BWGSearchManager
    Set colReception_ = tmpS.ReceptionSearch
    Set tmpS = Nothing
End Sub
Public Sub PopulateCancelRemark() 'ปิดสำหรับโปรแกรม Waste Data
On Error Resume Next
Dim tmpS As New BWGSearchManager
    Set colCancelRemark_ = tmpS.CancelRemarkSearch
    Set tmpS = Nothing
End Sub
Public Sub PopulateRem(ByRef cbo As Object)  'ปิดสำหรับโปรแกรม Waste Data
Dim x%, tmpS As New BWGSearchManager
    cbo.Clear
    cbo.AddItem "-เลือก-"
     Set colRem_ = tmpS.RemSearch
    For x = 1 To colRem_.Count
        cbo.AddItem colRem_(x).RemName
    Next
    If cbo.ListCount > 0 Then
        cbo.ListIndex = 0
    End If
End Sub
Public Sub PopulateStaffInfo(ByRef cbo As Object)  'ปิดสำหรับโปรแกรม Waste Data
Dim x%, tmpS As New BWGSearchManager
    cbo.Clear
    cbo.AddItem ""
    Set colStaffInfo_ = tmpS.BilInfoStaffSearch
    For x = 1 To colStaffInfo_.Count
        cbo.AddItem colStaffInfo_(x).InFoStaffName
    Next
    If cbo.ListCount > 0 Then
        cbo.ListIndex = -1
    End If
End Sub
Public Sub PopulateBankAccInfo(ByRef cbo As Object, Optional strCon As String)  'ปิดสำหรับโปรแกรม Waste Data
Dim x%, tmpS As New BWGSearchManager
    cbo.Clear
    cbo.AddItem ""
    Set colBankAccInfo_ = tmpS.BilInfoBankAccountSearch(strCon)
    For x = 1 To colBankAccInfo_.Count
        cbo.AddItem colBankAccInfo_(x).BankInfo
    Next
    If cbo.ListCount > 0 Then
        cbo.ListIndex = -1
    End If
End Sub

Public Sub PopulateVenderContact(ByRef cbo As Object, Optional strCon As String)
Dim x%, tmpS As New BWGSearchManager
    cbo.Clear
    cbo.AddItem "==ทั้งหมด=="
    Set colVenderContact_ = tmpS.SupplierContactSearch(strCon)
    For x = 1 To colVenderContact_.Count
        cbo.AddItem colVenderContact_(x).ContactName
    Next
    If cbo.ListCount > 0 Then
        cbo.ListIndex = 0
    End If
End Sub

Public Sub PopulateZone(ByRef cbo As Object)
Dim x%, tmpS As New BWGSearchManager
    cbo.Clear
    cbo.AddItem "==ทุกโซน=="
    Set colZone_ = tmpS.SaleZoneSearch
    For x = 1 To colZone_.Count
        cbo.AddItem colZone_(x).ZoneName
    Next
    If cbo.ListCount > 0 Then
        cbo.ListIndex = 0
    End If
End Sub

Public Sub PopulateFormList()
On Error GoTo ErrD
Dim tmpRec As ADODB.Recordset
Dim tmpCol As Collection
    If ColFormList Is Nothing Then
        Set tmpRec = New ADODB.Recordset
        Set tmpRec = GetRS("SELECT * FROM tbFormList ORDER BY FormDesc", adOpenForwardOnly, adLockReadOnly)
        If tmpRec Is Nothing Then Exit Sub
        Set tmpCol = New Collection
        Do Until tmpRec.EOF
            tmpCol.Add Trim("" & tmpRec("FormName")) & "_" & Trim("" & tmpRec("FormDesc")), "'" & Trim("" & tmpRec("FormName")) & "'"
            tmpRec.MoveNext
        Loop
        'tmpRec.Close
        Set tmpRec = Nothing
        Set ColFormList = tmpCol
        Set tmpCol = Nothing
    End If
    Exit Sub
ErrD:
    GetMsg "", vbCritical
    Exit Sub
End Sub

Public Sub PopulateUserGroup()
On Error GoTo ErrD
Dim tmpRec As ADODB.Recordset
Dim tmpCol As Collection
Dim tmpGroup As AWSUserGroup
    If ColUserGroup Is Nothing Then
        Set tmpRec = New ADODB.Recordset
        Set tmpRec = GetRS("SELECT * FROM tbUserGroup WHERE IsDeleted='N' ORDER BY GroupName", adOpenForwardOnly, adLockReadOnly)
        If tmpRec Is Nothing Then Exit Sub
        Set tmpCol = New Collection
        Do Until tmpRec.EOF
            Set tmpGroup = New AWSUserGroup
            tmpGroup.SetDataFromRecordset tmpRec
            tmpCol.Add tmpGroup, "'" & tmpGroup.ID & "'"
            Set tmpGroup = Nothing
            tmpRec.MoveNext
        Loop
        'tmpRec.Close
        Set tmpRec = Nothing
        Set ColUserGroup = tmpCol
        Set tmpCol = Nothing
    End If
    Exit Sub
ErrD:
    GetMsg "", vbCritical
    Exit Sub
End Sub

Public Sub WasteType()
    Dim tmpRec As New ADODB.Recordset
    Dim tmpType As BWGWasteType
    
    Set tmpRec = GetRS("SELECT * FROM tbWasteType where isactive='Y' ORDER BY PriorityLevel", adOpenForwardOnly, adLockReadOnly)
    Set WasteTypes = New Collection
    Do Until tmpRec.EOF
        Set tmpType = New BWGWasteType
        tmpType.SetDataFromRecordset tmpRec
        WasteTypes.Add tmpType, "'" & tmpType.WasteType & "'"
        Set tmpType = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
End Sub

Public Sub WasteCase()
    Dim tmpRec As New ADODB.Recordset
    Dim tmpType As BWGWasteCase
    
    Set tmpRec = GetRS("SELECT * FROM tbWasteDataCase ORDER BY CaseName", adOpenForwardOnly, adLockReadOnly)
    Set WasteCases = New Collection
    Do Until tmpRec.EOF
        Set tmpType = New BWGWasteCase
        tmpType.SetDataFromRecordset tmpRec
        WasteCases.Add tmpType, "'" & tmpType.CaseName & "'"
        Set tmpType = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
End Sub
    
Public Sub PopulateProvince(selCbo As ComboBox)
    selCbo.Clear
    With selCbo
        .AddItem "เชียงใหม่"
        .AddItem "เชียงราย"
        .AddItem "เพชรบุรี"
        .AddItem "เพชรบูรณ์"
        .AddItem "เลย"
        .AddItem "แพร่"
        .AddItem "แม่ฮ่องสอน"
        .AddItem "กระบี่"
        .AddItem "กรุงเทพมหานคร"
        .AddItem "กาญจนบุรี"
        .AddItem "กาฬสินธุ์"
        .AddItem "กำแพงเพชร"
        .AddItem "ขอนแก่น"
        .AddItem "จันทบุรี"
        .AddItem "ฉะเชิงเทรา"
        .AddItem "ชลบุรี"
        .AddItem "ชัยนาท"
        .AddItem "ชัยภูมิ"
        .AddItem "ชุมพร"
        .AddItem "ตรัง"
        .AddItem "ตราด"
        .AddItem "ตาก"
        .AddItem "นครนายก"
        .AddItem "นครปฐม"
        .AddItem "นครพนม"
        .AddItem "นครราชสีมา"
        .AddItem "นครศรีธรรมราช"
        .AddItem "นครสวรรค์"
        .AddItem "นนทบุรี"
        .AddItem "นราธิวาส"
        .AddItem "น่าน"
        .AddItem "บุรีรัมย์"
        .AddItem "ปทุมธานี"
        .AddItem "ประจวบคีรีขันธ์"
        .AddItem "ปราจีนบุรี"
        .AddItem "ปัตตานี"
        .AddItem "พระนครศรีอยุธยา"
        .AddItem "พะเยา"
        .AddItem "พังงา"
        .AddItem "พัทลุง"
        .AddItem "พิจิตร"
        .AddItem "พิษณุโลก"
        .AddItem "ภูเก็ต"
        .AddItem "มหาสารคาม"
        .AddItem "มุกดาหาร"
        .AddItem "ยโสธร"
        .AddItem "ยะลา"
        .AddItem "ร้อยเอ็ด"
        .AddItem "ระนอง"
        .AddItem "ระยอง"
        .AddItem "ราชบุรี"
        .AddItem "ลพบุรี"
        .AddItem "ลำปาง"
        .AddItem "ลำพูน"
        .AddItem "ศรีสะเกษ"
        .AddItem "สกลนคร"
        .AddItem "สงขลา"
        .AddItem "สตูล"
        .AddItem "สมุทรปราการ"
        .AddItem "สมุทรสงคราม"
        .AddItem "สมุทรสาคร"
        .AddItem "สระแก้ว"
        .AddItem "สระบุรี"
        .AddItem "สิงห์บุรี"
        .AddItem "สุโขทัย"
        .AddItem "สุพรรณบุรี"
        .AddItem "สุราษฎร์ธานี"
        .AddItem "สุรินทร์"
        .AddItem "หนองคาย"
        .AddItem "หนองบัวลำภู"
        .AddItem "อ่างทอง"
        .AddItem "อำนาจเจริญ"
        .AddItem "อุดรธานี"
        .AddItem "อุตรดิตถ์"
        .AddItem "อุทัยธานี"
        .AddItem "อุบลราชธานี"
    End With
End Sub
