VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmPrintORD02 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "F093"
   ClientHeight    =   3315
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6030
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3315
   ScaleWidth      =   6030
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport rpt1 
      Left            =   120
      Top             =   2640
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "Preview"
      Height          =   795
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2460
      Width           =   915
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   90
      TabIndex        =   8
      Top             =   810
      Width           =   5865
      Begin VB.ComboBox cboType 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1080
         Width           =   4395
      End
      Begin VB.ComboBox cboCompany 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   690
         Width           =   4395
      End
      Begin WasteManagment.ctlDate dtFrom 
         Height          =   315
         Left            =   1080
         TabIndex        =   0
         Top             =   300
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         BackColor       =   -2147483633
      End
      Begin WasteManagment.ctlDate dtTo 
         Height          =   315
         Left            =   3720
         TabIndex        =   1
         Top             =   300
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         BackColor       =   -2147483633
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "บริษัท"
         Height          =   195
         Index           =   0
         Left            =   540
         TabIndex        =   11
         Top             =   750
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ช่วงวันที่"
         Height          =   195
         Index           =   3
         Left            =   360
         TabIndex        =   10
         Top             =   360
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ถึงวันที่"
         Height          =   195
         Index           =   2
         Left            =   3120
         TabIndex        =   9
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "ปิด"
      Height          =   795
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2460
      Width           =   915
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "พิมพ์"
      Height          =   795
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2460
      Width           =   915
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   4
      Index           =   0
      X1              =   120
      X2              =   5760
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      Index           =   1
      X1              =   210
      X2              =   5910
      Y1              =   570
      Y2              =   570
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "พิมพ์รายงานยอดขายเทียบเป้าหมาย (ORD02)"
      BeginProperty Font 
         Name            =   "Angsana New"
         Size            =   15.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   1
      Left            =   150
      TabIndex        =   7
      Top             =   60
      Width           =   3915
   End
   Begin VB.Image Image2 
      Height          =   735
      Index           =   1
      Left            =   0
      Picture         =   "frPrintORD02.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11280
   End
End
Attribute VB_Name = "frmPrintORD02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sessionID_ As String

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdPreview_Click()
    PrintMain 0
End Sub

Private Sub cmdPrint_Click()
    PrintMain 1
End Sub

Private Sub Form_Load()
    CanAccess = False: CanUpdate = False
    SetPermission Me, CanAccess, CanUpdate
    If CanAccess = False Then
        MsgBox "ท่านไม่มีสิทธิใช้ระบบนี้", vbExclamation, "Warning"
        Unload Me
    End If
    
    PopulateCompany cboCompany, cboIndex
     With frmPicture
        cmdPreview.Picture = .picPreview.Picture
        cmdPrint.Picture = .picPrint.Picture
        cmdClose.Picture = .picExit.Picture
    End With
    With cboType
        .AddItem "สรุปจำนวนเที่ยว"
        .ListIndex = 0
    End With
    sessionID_ = GetGUID
End Sub

Private Sub PrintMain(printType As Integer)
'On Error GoTo ErrD
Dim tmpSql As String, rptName$
Dim tmpRec As New ADODB.Recordset, tmpStr As String
Dim startMonth$, startYear$, endMonth$, endYear$, tmpCompanyID As String
Dim sqlStr As String
    startMonth = dtFrom.MonthValue
    startYear = Left(dtFrom.ValueYMD, 4)
    endMonth = dtTo.MonthValue
    endYear = Left(dtTo.ValueYMD, 4)
    tmpCompanyID = ""
    If cboCompany.ListCount > 1 Then
        If cboCompany.ListIndex > 0 Then
            tmpCompanyID = Trim(CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).ID)
        End If
    Else
        If cboCompany.ListIndex >= 0 Then
            tmpCompanyID = Trim(CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).ID)
        End If
    End If
    
    Screen.MousePointer = 11

    If cboType.ListIndex = 0 Then
        If tmpCompanyID <> "FBAE386732E34AF1ABE2F69BB3F28AFF" Then
            DBConnExc "Update vw_UpdateJobWasteType set WasteType=WasteTypeMin WHERE ISNULL(WasteType, '') = '' "
'            'without SaleDistrictID
'            sqlStr = "SELECT SaleZoneID,ZoneName,SaleDistrictID,SDName,StaffID,StaffFName," & vbCrLf
'            sqlStr = sqlStr & "SumNoN=isnull((Select  SUM(TruckCount) From vw_TripCountByWasteType WHERE CompanyID='" & tmpCompanyID & "' And WorkDate>='" & dtFrom.ValueYMD & "' And WorkDate<='" & dtTo.ValueYMD & "' And UnderSaleStaffID=v.StaffID And TripWasteType='Non'),0)," & vbCrLf
'            sqlStr = sqlStr & "SumHaz=isnull((Select  SUM(TruckCount) From vw_TripCountByWasteType WHERE CompanyID='" & tmpCompanyID & "' And WorkDate>='" & dtFrom.ValueYMD & "' And WorkDate<='" & dtTo.ValueYMD & "' And UnderSaleStaffID=v.StaffID And TripWasteType='Haz'),0), " & vbCrLf
'            sqlStr = sqlStr & "SumHBL=isnull((Select  SUM(TruckCount) From vw_TripCountByWasteType WHERE CompanyID='" & tmpCompanyID & "' And WorkDate>='" & dtFrom.ValueYMD & "' And WorkDate<='" & dtTo.ValueYMD & "' And UnderSaleStaffID=v.StaffID And TripWasteType='HBL'),0), " & vbCrLf
'            sqlStr = sqlStr & "SumNBL=isnull((Select  SUM(TruckCount) From vw_TripCountByWasteType WHERE CompanyID='" & tmpCompanyID & "' And WorkDate>='" & dtFrom.ValueYMD & "' And WorkDate<='" & dtTo.ValueYMD & "' And UnderSaleStaffID=v.StaffID And TripWasteType='NBL'),0), " & vbCrLf
'            sqlStr = sqlStr & "SumHIC=isnull((Select  SUM(TruckCount) From vw_TripCountByWasteType WHERE CompanyID='" & tmpCompanyID & "' And WorkDate>='" & dtFrom.ValueYMD & "' And WorkDate<='" & dtTo.ValueYMD & "' And UnderSaleStaffID=v.StaffID And TripWasteType='HIC'),0), " & vbCrLf
'            sqlStr = sqlStr & "SumHBI=isnull((Select  SUM(TruckCount) From vw_TripCountByWasteType WHERE CompanyID='" & tmpCompanyID & "' And WorkDate>='" & dtFrom.ValueYMD & "' And WorkDate<='" & dtTo.ValueYMD & "' And UnderSaleStaffID=v.StaffID And TripWasteType='HBI'),0), " & vbCrLf
'            sqlStr = sqlStr & "SumOth=isnull((Select  SUM(TruckCount) From vw_TripCountByWasteType WHERE CompanyID='" & tmpCompanyID & "' And WorkDate>='" & dtFrom.ValueYMD & "' And WorkDate<='" & dtTo.ValueYMD & "' And UnderSaleStaffID=v.StaffID And TripWasteType NOT IN ('Non', 'Haz', 'NBL', 'HBL', 'HIC', 'HBI')),0) " & vbCrLf
'            sqlStr = sqlStr & "FROM vw_RptOrd02_2 v Group By SaleZoneID,ZoneName,SaleDistrictID,SDName,StaffID,StaffFName"
'
            'with SaledistrictID
            sqlStr = "SELECT SaleZoneID,ZoneName,SaleDistrictID,SDName,StaffID,StaffFName," & vbCrLf
            sqlStr = sqlStr & "SumNoN=isnull((Select  SUM(TruckCount) From vw_TripCountByWasteType WHERE CompanyID='" & tmpCompanyID & "' And WorkDate>='" & dtFrom.ValueYMD & "' And WorkDate<='" & dtTo.ValueYMD & "' And UnderSaleStaffID=v.StaffID And TripWasteType='Non' And SaleDistrictID=v.SaleDistrictID),0)," & vbCrLf
            sqlStr = sqlStr & "SumHaz=isnull((Select  SUM(TruckCount) From vw_TripCountByWasteType WHERE CompanyID='" & tmpCompanyID & "' And WorkDate>='" & dtFrom.ValueYMD & "' And WorkDate<='" & dtTo.ValueYMD & "' And UnderSaleStaffID=v.StaffID And TripWasteType='Haz' And SaleDistrictID=v.SaleDistrictID),0), " & vbCrLf
            sqlStr = sqlStr & "SumHBL=isnull((Select  SUM(TruckCount) From vw_TripCountByWasteType WHERE CompanyID='" & tmpCompanyID & "' And WorkDate>='" & dtFrom.ValueYMD & "' And WorkDate<='" & dtTo.ValueYMD & "' And UnderSaleStaffID=v.StaffID And TripWasteType='HBL' And SaleDistrictID=v.SaleDistrictID),0), " & vbCrLf
            sqlStr = sqlStr & "SumNBL=isnull((Select  SUM(TruckCount) From vw_TripCountByWasteType WHERE CompanyID='" & tmpCompanyID & "' And WorkDate>='" & dtFrom.ValueYMD & "' And WorkDate<='" & dtTo.ValueYMD & "' And UnderSaleStaffID=v.StaffID And TripWasteType='NBL' And SaleDistrictID=v.SaleDistrictID),0), " & vbCrLf
            sqlStr = sqlStr & "SumHIC=isnull((Select  SUM(TruckCount) From vw_TripCountByWasteType WHERE CompanyID='" & tmpCompanyID & "' And WorkDate>='" & dtFrom.ValueYMD & "' And WorkDate<='" & dtTo.ValueYMD & "' And UnderSaleStaffID=v.StaffID And TripWasteType='HIC' And SaleDistrictID=v.SaleDistrictID),0), " & vbCrLf
            sqlStr = sqlStr & "SumHBI=isnull((Select  SUM(TruckCount) From vw_TripCountByWasteType WHERE CompanyID='" & tmpCompanyID & "' And WorkDate>='" & dtFrom.ValueYMD & "' And WorkDate<='" & dtTo.ValueYMD & "' And UnderSaleStaffID=v.StaffID And TripWasteType='HBI' And SaleDistrictID=v.SaleDistrictID),0), " & vbCrLf
            sqlStr = sqlStr & "SumNIC=isnull((Select  SUM(TruckCount) From vw_TripCountByWasteType WHERE CompanyID='" & tmpCompanyID & "' And WorkDate>='" & dtFrom.ValueYMD & "' And WorkDate<='" & dtTo.ValueYMD & "' And UnderSaleStaffID=v.StaffID And TripWasteType='NIC' And SaleDistrictID=v.SaleDistrictID),0), " & vbCrLf
            sqlStr = sqlStr & "SumNBI=isnull((Select  SUM(TruckCount) From vw_TripCountByWasteType WHERE CompanyID='" & tmpCompanyID & "' And WorkDate>='" & dtFrom.ValueYMD & "' And WorkDate<='" & dtTo.ValueYMD & "' And UnderSaleStaffID=v.StaffID And TripWasteType='NBI' And SaleDistrictID=v.SaleDistrictID),0), " & vbCrLf
            sqlStr = sqlStr & "SumBBH=isnull((Select  SUM(TruckCount) From vw_TripCountByWasteType WHERE CompanyID='" & tmpCompanyID & "' And WorkDate>='" & dtFrom.ValueYMD & "' And WorkDate<='" & dtTo.ValueYMD & "' And UnderSaleStaffID=v.StaffID And TripWasteType='BBH' And SaleDistrictID=v.SaleDistrictID),0), " & vbCrLf
            sqlStr = sqlStr & "SumBBL=isnull((Select  SUM(TruckCount) From vw_TripCountByWasteType WHERE CompanyID='" & tmpCompanyID & "' And WorkDate>='" & dtFrom.ValueYMD & "' And WorkDate<='" & dtTo.ValueYMD & "' And UnderSaleStaffID=v.StaffID And TripWasteType='BBL' And SaleDistrictID=v.SaleDistrictID),0), " & vbCrLf
            sqlStr = sqlStr & "SumOth=isnull((Select  SUM(TruckCount) From vw_TripCountByWasteType WHERE CompanyID='" & tmpCompanyID & "' And WorkDate>='" & dtFrom.ValueYMD & "' And WorkDate<='" & dtTo.ValueYMD & "' And UnderSaleStaffID=v.StaffID And TripWasteType NOT IN ('Non', 'Haz', 'NBL', 'HBL', 'HIC', 'HBI', 'NIC', 'NBI', 'BBH', 'BBL') And SaleDistrictID=v.SaleDistrictID),0) " & vbCrLf
            sqlStr = sqlStr & "FROM vw_RptOrd02_2 v Group By SaleZoneID,ZoneName,SaleDistrictID,SDName,StaffID,StaffFName"
    
            Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
            If tmpRec Is Nothing Then Exit Sub
            
                DBConnExc "Delete tbRptOrd02 WHERE SessionID='" & sessionID_ & "'"
                Do Until tmpRec.EOF
                    DBConnExc "Insert into tbRptOrd02 (StaffID,ZoneID,SaleDistrictID,ZoneName,SDName,StaffFullName,NonAmt,HazAmt,NBLAmt,HBLAmt, HICAmt, HBIAmt, NICAmt, NBIAmt, BBHAmt, BBLAmt, OthAmt, SessionID,StaffType) Values (" & _
                                    "'" & Trim("" & tmpRec("StaffID")) & "','" & Trim("" & tmpRec("SaleZoneID")) & "','" & Trim("" & tmpRec("SaleDistrictID")) & "','" & Replace(Trim("" & tmpRec("ZoneName")), "โซน", "") & "','" & Trim("" & tmpRec("SDName")) & "','" & Trim("" & tmpRec("StaffFName")) & "'," & _
                                    tmpRec("SumNon") & "," & tmpRec("SumHaz") & "," & tmpRec("SumNBL") & "," & tmpRec("SumHBL") & "," & tmpRec("SumHIC") & "," & tmpRec("SumHBI") & "," & tmpRec("SumNIC") & "," & tmpRec("SumNBI") & "," & tmpRec("SumBBH") & "," & tmpRec("SumBBL") & "," & tmpRec("SumOth") & ",'" & sessionID_ & "','C')"
                    tmpRec.MoveNext
                Loop

            tmpRec.Close
            sqlStr = "SELECT SaleZoneID,ZoneName,SaleDistrictID,SDName,StaffID,StaffFName," & vbCrLf
            sqlStr = sqlStr & "SumNoN=isnull((Select  SUM(TruckCount) From vw_TripCountByWasteType WHERE CompanyID='" & tmpCompanyID & "' And WorkDate>='" & dtFrom.ValueYMD & "' And WorkDate<='" & dtTo.ValueYMD & "' And UnderSaleStaffID=v.StaffID And TripWasteType='Non' And SaleDistrictID=v.SaleDistrictID),0)," & vbCrLf
            sqlStr = sqlStr & "SumHaz=isnull((Select  SUM(TruckCount) From vw_TripCountByWasteType WHERE CompanyID='" & tmpCompanyID & "' And WorkDate>='" & dtFrom.ValueYMD & "' And WorkDate<='" & dtTo.ValueYMD & "' And UnderSaleStaffID=v.StaffID And TripWasteType='Haz' And SaleDistrictID=v.SaleDistrictID),0), " & vbCrLf
            sqlStr = sqlStr & "SumHBL=isnull((Select  SUM(TruckCount) From vw_TripCountByWasteType WHERE CompanyID='" & tmpCompanyID & "' And WorkDate>='" & dtFrom.ValueYMD & "' And WorkDate<='" & dtTo.ValueYMD & "' And UnderSaleStaffID=v.StaffID And TripWasteType='HBL' And SaleDistrictID=v.SaleDistrictID),0), " & vbCrLf
            sqlStr = sqlStr & "SumNBL=isnull((Select  SUM(TruckCount) From vw_TripCountByWasteType WHERE CompanyID='" & tmpCompanyID & "' And WorkDate>='" & dtFrom.ValueYMD & "' And WorkDate<='" & dtTo.ValueYMD & "' And UnderSaleStaffID=v.StaffID And TripWasteType='NBL' And SaleDistrictID=v.SaleDistrictID),0), " & vbCrLf
            sqlStr = sqlStr & "SumHIC=isnull((Select  SUM(TruckCount) From vw_TripCountByWasteType WHERE CompanyID='" & tmpCompanyID & "' And WorkDate>='" & dtFrom.ValueYMD & "' And WorkDate<='" & dtTo.ValueYMD & "' And UnderSaleStaffID=v.StaffID And TripWasteType='HIC' And SaleDistrictID=v.SaleDistrictID),0), " & vbCrLf
            sqlStr = sqlStr & "SumHBI=isnull((Select  SUM(TruckCount) From vw_TripCountByWasteType WHERE CompanyID='" & tmpCompanyID & "' And WorkDate>='" & dtFrom.ValueYMD & "' And WorkDate<='" & dtTo.ValueYMD & "' And UnderSaleStaffID=v.StaffID And TripWasteType='HBI' And SaleDistrictID=v.SaleDistrictID),0), " & vbCrLf
            sqlStr = sqlStr & "SumNIC=isnull((Select  SUM(TruckCount) From vw_TripCountByWasteType WHERE CompanyID='" & tmpCompanyID & "' And WorkDate>='" & dtFrom.ValueYMD & "' And WorkDate<='" & dtTo.ValueYMD & "' And UnderSaleStaffID=v.StaffID And TripWasteType='NIC' And SaleDistrictID=v.SaleDistrictID),0), " & vbCrLf
            sqlStr = sqlStr & "SumNBI=isnull((Select  SUM(TruckCount) From vw_TripCountByWasteType WHERE CompanyID='" & tmpCompanyID & "' And WorkDate>='" & dtFrom.ValueYMD & "' And WorkDate<='" & dtTo.ValueYMD & "' And UnderSaleStaffID=v.StaffID And TripWasteType='NBI' And SaleDistrictID=v.SaleDistrictID),0), " & vbCrLf
            sqlStr = sqlStr & "SumBBH=isnull((Select  SUM(TruckCount) From vw_TripCountByWasteType WHERE CompanyID='" & tmpCompanyID & "' And WorkDate>='" & dtFrom.ValueYMD & "' And WorkDate<='" & dtTo.ValueYMD & "' And UnderSaleStaffID=v.StaffID And TripWasteType='BBH' And SaleDistrictID=v.SaleDistrictID),0), " & vbCrLf
            sqlStr = sqlStr & "SumBBL=isnull((Select  SUM(TruckCount) From vw_TripCountByWasteType WHERE CompanyID='" & tmpCompanyID & "' And WorkDate>='" & dtFrom.ValueYMD & "' And WorkDate<='" & dtTo.ValueYMD & "' And UnderSaleStaffID=v.StaffID And TripWasteType='BBL' And SaleDistrictID=v.SaleDistrictID),0), " & vbCrLf
            sqlStr = sqlStr & "SumOth=isnull((Select  SUM(TruckCount) From vw_TripCountByWasteType WHERE CompanyID='" & tmpCompanyID & "' And WorkDate>='" & dtFrom.ValueYMD & "' And WorkDate<='" & dtTo.ValueYMD & "' And UnderSaleStaffID=v.StaffID And TripWasteType NOT IN ('Non', 'Haz', 'NBL', 'HBL', 'HIC', 'HBI', 'NIC', 'NBI', 'BBH', 'BBL') And SaleDistrictID=v.SaleDistrictID),0) " & vbCrLf
            sqlStr = sqlStr & "FROM vw_RptOrd02_2Job v WHERE (CompanyID='" & tmpCompanyID & "' or CompanyID is null) And StaffID<>SaleStaffID Group By SaleZoneID,ZoneName,SaleDistrictID,SDName,StaffID,StaffFName"
            Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    '        DBConnExc sqlStr, tmpRec
            If tmpRec Is Nothing Then Exit Sub
    '        If tmpRec.RecordCount > 0 Then tmpRec.MoveFirst
    '            .Execute "Delete tbRptOrd02 WHERE SessionID='" & sessionID_ & "'"
                Do While Not tmpRec.EOF
                    If tmpRec("SumNoN") + tmpRec("SumHaz") + tmpRec("SumHBL") + tmpRec("SumNBL") + tmpRec("SumHIC") + tmpRec("SumHBI") + tmpRec("SumNIC") + tmpRec("SumNBI") + tmpRec("SumBBH") + tmpRec("SumBBL") + tmpRec("SumOth") > 0 Then
                        DBConnExc "Insert into tbRptOrd02 (StaffID,ZoneID,SaleDistrictID,ZoneName,SDName,StaffFullName,NonAmt,HazAmt,NBLAmt,HBLAmt, HICAmt, HBIAmt, NICAmt, NBIAmt, BBHAmt, BBLAmt, OthAmt, SessionID,StaffType) Values (" & _
                                        "'" & Trim("" & tmpRec("StaffID")) & "','" & Trim("" & tmpRec("SaleZoneID")) & "','" & Trim("" & tmpRec("SaleDistrictID")) & "','" & Replace(Trim("" & tmpRec("ZoneName")), "โซน", "") & "','" & Trim("" & tmpRec("SDName")) & "','" & Trim("" & tmpRec("StaffFName")) & "'," & _
                                        tmpRec("SumNon") & "," & tmpRec("SumHaz") & "," & tmpRec("SumNBL") & "," & tmpRec("SumHBL") & "," & tmpRec("SumHIC") & "," & tmpRec("SumHBI") & "," & tmpRec("SumNIC") & "," & tmpRec("SumNBI") & "," & tmpRec("SumBBH") & "," & tmpRec("SumBBL") & "," & tmpRec("SumOth") & ",'" & sessionID_ & "','J')"
                    End If
                    tmpRec.MoveNext
                Loop
        Else
            DBConnExc "Update vw_UpdateJobWasteType set WasteType=TripWasteType"
'            'without SaleDistrictID
'            sqlStr = "SELECT SaleZoneID,ZoneName,SaleDistrictID,SDName,StaffID,StaffFName," & vbCrLf
'            sqlStr = sqlStr & "SumNoN=isnull((Select  SUM(TruckCount) From vw_TripCountByWasteType WHERE APManifestNo LIKE '%AP%' And WorkDate>='" & dtFrom.ValueYMD & "' And WorkDate<='" & dtTo.ValueYMD & "' And UnderSaleStaffID=v.StaffID And TripWasteType='Non'),0)," & vbCrLf
'            sqlStr = sqlStr & "SumHaz=isnull((Select  SUM(TruckCount) From vw_TripCountByWasteType WHERE APManifestNo LIKE '%AP%' And WorkDate>='" & dtFrom.ValueYMD & "' And WorkDate<='" & dtTo.ValueYMD & "' And UnderSaleStaffID=v.StaffID And TripWasteType='Haz'),0), " & vbCrLf
'            sqlStr = sqlStr & "SumHBL=isnull((Select  SUM(TruckCount) From vw_TripCountByWasteType WHERE APManifestNo LIKE '%AP%' And WorkDate>='" & dtFrom.ValueYMD & "' And WorkDate<='" & dtTo.ValueYMD & "' And UnderSaleStaffID=v.StaffID And TripWasteType='HBL'),0), " & vbCrLf
'            sqlStr = sqlStr & "SumNBL=isnull((Select  SUM(TruckCount) From vw_TripCountByWasteType WHERE APManifestNo LIKE '%AP%' And WorkDate>='" & dtFrom.ValueYMD & "' And WorkDate<='" & dtTo.ValueYMD & "' And UnderSaleStaffID=v.StaffID And TripWasteType='NBL'),0), " & vbCrLf
'            sqlStr = sqlStr & "SumHIC=isnull((Select  SUM(TruckCount) From vw_TripCountByWasteType WHERE APManifestNo LIKE '%AP%' And WorkDate>='" & dtFrom.ValueYMD & "' And WorkDate<='" & dtTo.ValueYMD & "' And UnderSaleStaffID=v.StaffID And TripWasteType='HIC'),0), " & vbCrLf
'            sqlStr = sqlStr & "SumHBI=isnull((Select  SUM(TruckCount) From vw_TripCountByWasteType WHERE APManifestNo LIKE '%AP%' And WorkDate>='" & dtFrom.ValueYMD & "' And WorkDate<='" & dtTo.ValueYMD & "' And UnderSaleStaffID=v.StaffID And TripWasteType='HBI'),0), " & vbCrLf
'            sqlStr = sqlStr & "SumOth=isnull((Select  SUM(TruckCount) From vw_TripCountByWasteType WHERE APManifestNo LIKE '%AP%' And WorkDate>='" & dtFrom.ValueYMD & "' And WorkDate<='" & dtTo.ValueYMD & "' And UnderSaleStaffID=v.StaffID And TripWasteType NOT IN ('Non', 'Haz', 'NBL', 'HBL', 'HIC', 'HBI')),0) " & vbCrLf
'            sqlStr = sqlStr & "FROM vw_RptOrd02_2 v Group By SaleZoneID,ZoneName,SaleDistrictID,SDName,StaffID,StaffFName"
            
'            'with SaleDistrictID
            sqlStr = "SELECT SaleZoneID,ZoneName,SaleDistrictID,SDName,StaffID,StaffFName," & vbCrLf
            sqlStr = sqlStr & "SumNoN=isnull((Select  SUM(TruckCount) From vw_TripCountByWasteType WHERE APManifestNo LIKE '%AP%' And WorkDate>='" & dtFrom.ValueYMD & "' And WorkDate<='" & dtTo.ValueYMD & "' And UnderSaleStaffID=v.StaffID And TripWasteType='Non' And SaleDistrictID=v.SaleDistrictID),0)," & vbCrLf
            sqlStr = sqlStr & "SumHaz=isnull((Select  SUM(TruckCount) From vw_TripCountByWasteType WHERE APManifestNo LIKE '%AP%' And WorkDate>='" & dtFrom.ValueYMD & "' And WorkDate<='" & dtTo.ValueYMD & "' And UnderSaleStaffID=v.StaffID And TripWasteType='Haz' And SaleDistrictID=v.SaleDistrictID),0), " & vbCrLf
            sqlStr = sqlStr & "SumHBL=isnull((Select  SUM(TruckCount) From vw_TripCountByWasteType WHERE APManifestNo LIKE '%AP%' And WorkDate>='" & dtFrom.ValueYMD & "' And WorkDate<='" & dtTo.ValueYMD & "' And UnderSaleStaffID=v.StaffID And TripWasteType='HBL' And SaleDistrictID=v.SaleDistrictID),0), " & vbCrLf
            sqlStr = sqlStr & "SumNBL=isnull((Select  SUM(TruckCount) From vw_TripCountByWasteType WHERE APManifestNo LIKE '%AP%' And WorkDate>='" & dtFrom.ValueYMD & "' And WorkDate<='" & dtTo.ValueYMD & "' And UnderSaleStaffID=v.StaffID And TripWasteType='NBL' And SaleDistrictID=v.SaleDistrictID),0), " & vbCrLf
            sqlStr = sqlStr & "SumHIC=isnull((Select  SUM(TruckCount) From vw_TripCountByWasteType WHERE APManifestNo LIKE '%AP%' And WorkDate>='" & dtFrom.ValueYMD & "' And WorkDate<='" & dtTo.ValueYMD & "' And UnderSaleStaffID=v.StaffID And TripWasteType='HIC' And SaleDistrictID=v.SaleDistrictID),0), " & vbCrLf
            sqlStr = sqlStr & "SumHBI=isnull((Select  SUM(TruckCount) From vw_TripCountByWasteType WHERE APManifestNo LIKE '%AP%' And WorkDate>='" & dtFrom.ValueYMD & "' And WorkDate<='" & dtTo.ValueYMD & "' And UnderSaleStaffID=v.StaffID And TripWasteType='HBI' And SaleDistrictID=v.SaleDistrictID),0), " & vbCrLf
            sqlStr = sqlStr & "SumNIC=isnull((Select  SUM(TruckCount) From vw_TripCountByWasteType WHERE APManifestNo LIKE '%AP%' And WorkDate>='" & dtFrom.ValueYMD & "' And WorkDate<='" & dtTo.ValueYMD & "' And UnderSaleStaffID=v.StaffID And TripWasteType='NIC' And SaleDistrictID=v.SaleDistrictID),0), " & vbCrLf
            sqlStr = sqlStr & "SumNBI=isnull((Select  SUM(TruckCount) From vw_TripCountByWasteType WHERE APManifestNo LIKE '%AP%' And WorkDate>='" & dtFrom.ValueYMD & "' And WorkDate<='" & dtTo.ValueYMD & "' And UnderSaleStaffID=v.StaffID And TripWasteType='NBI' And SaleDistrictID=v.SaleDistrictID),0), " & vbCrLf
            sqlStr = sqlStr & "SumBBH=isnull((Select  SUM(TruckCount) From vw_TripCountByWasteType WHERE APManifestNo LIKE '%AP%' And WorkDate>='" & dtFrom.ValueYMD & "' And WorkDate<='" & dtTo.ValueYMD & "' And UnderSaleStaffID=v.StaffID And TripWasteType='BBH' And SaleDistrictID=v.SaleDistrictID),0), " & vbCrLf
            sqlStr = sqlStr & "SumBBL=isnull((Select  SUM(TruckCount) From vw_TripCountByWasteType WHERE APManifestNo LIKE '%AP%' And WorkDate>='" & dtFrom.ValueYMD & "' And WorkDate<='" & dtTo.ValueYMD & "' And UnderSaleStaffID=v.StaffID And TripWasteType='BBL' And SaleDistrictID=v.SaleDistrictID),0), " & vbCrLf
            sqlStr = sqlStr & "SumOth=isnull((Select  SUM(TruckCount) From vw_TripCountByWasteType WHERE APManifestNo LIKE '%AP%' And WorkDate>='" & dtFrom.ValueYMD & "' And WorkDate<='" & dtTo.ValueYMD & "' And UnderSaleStaffID=v.StaffID And TripWasteType NOT IN ('Non', 'Haz', 'NBL', 'HBL', 'HIC', 'HBI', 'NIC', 'NBI', 'BBH', 'BBL') And SaleDistrictID=v.SaleDistrictID),0) " & vbCrLf
            sqlStr = sqlStr & "FROM vw_RptOrd02_2 v Group By SaleZoneID,ZoneName,SaleDistrictID,SDName,StaffID,StaffFName"
            
            Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
            If tmpRec Is Nothing Then Exit Sub

                DBConnExc "Delete tbRptOrd02 WHERE SessionID='" & sessionID_ & "'"
                Do Until tmpRec.EOF
                    DBConnExc "Insert into tbRptOrd02 (StaffID,ZoneID,SaleDistrictID,ZoneName,SDName,StaffFullName,NonAmt,HazAmt,NBLAmt,HBLAmt, HICAmt, HBIAmt, NICAmt, NBIAmt, BBHAmt, BBLAmt, OthAmt, SessionID,StaffType) Values (" & _
                                    "'" & Trim("" & tmpRec("StaffID")) & "','" & Trim("" & tmpRec("SaleZoneID")) & "','" & Trim("" & tmpRec("SaleDistrictID")) & "','" & Replace(Trim("" & tmpRec("ZoneName")), "โซน", "") & "','" & Trim("" & tmpRec("SDName")) & "','" & Trim("" & tmpRec("StaffFName")) & "'," & _
                                    tmpRec("SumNon") & "," & tmpRec("SumHaz") & "," & tmpRec("SumNBL") & "," & tmpRec("SumHBL") & "," & tmpRec("SumHIC") & "," & tmpRec("SumHBI") & "," & tmpRec("SumNIC") & "," & tmpRec("SumNBI") & "," & tmpRec("SumBBH") & "," & tmpRec("SumBBL") & "," & tmpRec("SumOth") & ",'" & sessionID_ & "','C')"
                    tmpRec.MoveNext
                Loop

            tmpRec.Close
            sqlStr = "SELECT SaleZoneID,ZoneName,SaleDistrictID,SDName,StaffID,StaffFName," & vbCrLf
            sqlStr = sqlStr & "SumNoN=isnull((Select  SUM(TruckCount) From vw_TripCountByWasteType WHERE APManifestNo LIKE '%AP%' And WorkDate>='" & dtFrom.ValueYMD & "' And WorkDate<='" & dtTo.ValueYMD & "' And UnderSaleStaffID=v.StaffID And TripWasteType='Non' And SaleDistrictID=v.SaleDistrictID),0)," & vbCrLf
            sqlStr = sqlStr & "SumHaz=isnull((Select  SUM(TruckCount) From vw_TripCountByWasteType WHERE APManifestNo LIKE '%AP%' And WorkDate>='" & dtFrom.ValueYMD & "' And WorkDate<='" & dtTo.ValueYMD & "' And UnderSaleStaffID=v.StaffID And TripWasteType='Haz' And SaleDistrictID=v.SaleDistrictID),0), " & vbCrLf
            sqlStr = sqlStr & "SumHBL=isnull((Select  SUM(TruckCount) From vw_TripCountByWasteType WHERE APManifestNo LIKE '%AP%' And WorkDate>='" & dtFrom.ValueYMD & "' And WorkDate<='" & dtTo.ValueYMD & "' And UnderSaleStaffID=v.StaffID And TripWasteType='HBL' And SaleDistrictID=v.SaleDistrictID),0), " & vbCrLf
            sqlStr = sqlStr & "SumNBL=isnull((Select  SUM(TruckCount) From vw_TripCountByWasteType WHERE APManifestNo LIKE '%AP%' And WorkDate>='" & dtFrom.ValueYMD & "' And WorkDate<='" & dtTo.ValueYMD & "' And UnderSaleStaffID=v.StaffID And TripWasteType='NBL' And SaleDistrictID=v.SaleDistrictID),0), " & vbCrLf
            sqlStr = sqlStr & "SumHIC=isnull((Select  SUM(TruckCount) From vw_TripCountByWasteType WHERE APManifestNo LIKE '%AP%' And WorkDate>='" & dtFrom.ValueYMD & "' And WorkDate<='" & dtTo.ValueYMD & "' And UnderSaleStaffID=v.StaffID And TripWasteType='HIC' And SaleDistrictID=v.SaleDistrictID),0), " & vbCrLf
            sqlStr = sqlStr & "SumHBI=isnull((Select  SUM(TruckCount) From vw_TripCountByWasteType WHERE APManifestNo LIKE '%AP%' And WorkDate>='" & dtFrom.ValueYMD & "' And WorkDate<='" & dtTo.ValueYMD & "' And UnderSaleStaffID=v.StaffID And TripWasteType='HBI' And SaleDistrictID=v.SaleDistrictID),0), " & vbCrLf
            sqlStr = sqlStr & "SumNIC=isnull((Select  SUM(TruckCount) From vw_TripCountByWasteType WHERE APManifestNo LIKE '%AP%' And WorkDate>='" & dtFrom.ValueYMD & "' And WorkDate<='" & dtTo.ValueYMD & "' And UnderSaleStaffID=v.StaffID And TripWasteType='NIC' And SaleDistrictID=v.SaleDistrictID),0), " & vbCrLf
            sqlStr = sqlStr & "SumNBI=isnull((Select  SUM(TruckCount) From vw_TripCountByWasteType WHERE APManifestNo LIKE '%AP%' And WorkDate>='" & dtFrom.ValueYMD & "' And WorkDate<='" & dtTo.ValueYMD & "' And UnderSaleStaffID=v.StaffID And TripWasteType='NBI' And SaleDistrictID=v.SaleDistrictID),0), " & vbCrLf
            sqlStr = sqlStr & "SumBBH=isnull((Select  SUM(TruckCount) From vw_TripCountByWasteType WHERE APManifestNo LIKE '%AP%' And WorkDate>='" & dtFrom.ValueYMD & "' And WorkDate<='" & dtTo.ValueYMD & "' And UnderSaleStaffID=v.StaffID And TripWasteType='BBH' And SaleDistrictID=v.SaleDistrictID),0), " & vbCrLf
            sqlStr = sqlStr & "SumBBL=isnull((Select  SUM(TruckCount) From vw_TripCountByWasteType WHERE APManifestNo LIKE '%AP%' And WorkDate>='" & dtFrom.ValueYMD & "' And WorkDate<='" & dtTo.ValueYMD & "' And UnderSaleStaffID=v.StaffID And TripWasteType='BBL' And SaleDistrictID=v.SaleDistrictID),0), " & vbCrLf
            sqlStr = sqlStr & "SumOth=isnull((Select  SUM(TruckCount) From vw_TripCountByWasteType WHERE APManifestNo LIKE '%AP%' And WorkDate>='" & dtFrom.ValueYMD & "' And WorkDate<='" & dtTo.ValueYMD & "' And UnderSaleStaffID=v.StaffID And TripWasteType NOT IN ('Non', 'Haz', 'NBL', 'HBL', 'HIC', 'HBI', 'NIC', 'NBI', 'BBH', 'BBL') And SaleDistrictID=v.SaleDistrictID),0) " & vbCrLf
            sqlStr = sqlStr & "FROM vw_RptOrd02_2Job v WHERE StaffID<>SaleStaffID Group By SaleZoneID,ZoneName,SaleDistrictID,SDName,StaffID,StaffFName"
            
            Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
            If tmpRec Is Nothing Then Exit Sub

                Do While Not tmpRec.EOF
                    If tmpRec("SumNoN") + tmpRec("SumHaz") + tmpRec("SumHBL") + tmpRec("SumNBL") + tmpRec("SumHIC") + tmpRec("SumHBI") + tmpRec("SumNIC") + tmpRec("SumNBI") + tmpRec("SumBBH") + tmpRec("SumBBL") + tmpRec("SumOth") > 0 Then
                        DBConnExc "Insert into tbRptOrd02 (StaffID,ZoneID,SaleDistrictID,ZoneName,SDName,StaffFullName,NonAmt,HazAmt,NBLAmt,HBLAmt, HICAmt, HBIAmt, NICAmt, NBIAmt, BBHAmt, BBLAmt, OthAmt, SessionID,StaffType) Values (" & _
                                        "'" & Trim("" & tmpRec("StaffID")) & "','" & Trim("" & tmpRec("SaleZoneID")) & "','" & Trim("" & tmpRec("SaleDistrictID")) & "','" & Replace(Trim("" & tmpRec("ZoneName")), "โซน", "") & "','" & Trim("" & tmpRec("SDName")) & "','" & Trim("" & tmpRec("StaffFName")) & "'," & _
                                        tmpRec("SumNon") & "," & tmpRec("SumHaz") & "," & tmpRec("SumNBL") & "," & tmpRec("SumHBL") & "," & tmpRec("SumHIC") & "," & tmpRec("SumHBI") & "," & tmpRec("SumNIC") & "," & tmpRec("SumNBI") & "," & tmpRec("SumBBH") & "," & tmpRec("SumBBL") & "," & tmpRec("SumOth") & ",'" & sessionID_ & "','J')"
                    End If
                    tmpRec.MoveNext
                Loop

        End If
                
    End If
    
    rptName = "ORD02_2.rpt"
    Screen.MousePointer = 0
    With rpt1
        .LogonInfo(0) = "DSN = " & DSN_NAME & ";UID = " & strUid & ";PWD = " & strPwd
        .ReportFileName = App.Path & "\reports\" & rptName
        tmpSql = ""
        tmpSql = "SELECT * FROM tbRptOrd02 " & vbCrLf
        tmpSql = tmpSql & "WHERE SessionID='" & sessionID_ & "' "
        tmpSql = tmpSql & vbCrLf & "ORDER BY ZoneName,SDName,StaffType,StaffFullName"
        .Formulas(0) = "DateRange=" & Chr(34) & "ในช่วงวันที่ขนตั้งแต่ " & dtFrom.ValueDMY & " ถึงวันที่ " & dtTo.ValueDMY & Chr(34)
        .Formulas(2) = "ShowCompany=" & Chr(34) & cboCompany.Text & Chr(34)
        .SQLQuery = tmpSql
        .Destination = printType
        .PrintReport
        If Trim(.LastErrorString) <> "" Then MsgBox .LastErrorString
    End With
    Exit Sub
ErrD:
    Screen.MousePointer = 0
    GetMsg "", vbCritical
    Exit Sub
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    DBConnExc "Delete tbRptOrd02 WHERE SessionID='" & sessionID_ & "'"
End Sub
