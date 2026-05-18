VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmPrintQUO03 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "F102"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6060
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   6060
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3015
      Left            =   90
      TabIndex        =   10
      Top             =   780
      Width           =   5925
      Begin VB.OptionButton optincrease 
         BackColor       =   &H80000005&
         Caption         =   "รายได้สะสม"
         Height          =   375
         Left            =   3000
         TabIndex        =   18
         Top             =   2400
         Width           =   2175
      End
      Begin VB.OptionButton optnormal 
         BackColor       =   &H80000005&
         Caption         =   "ปกติ"
         Height          =   375
         Left            =   1080
         TabIndex        =   17
         Top             =   2400
         Value           =   -1  'True
         Width           =   1815
      End
      Begin VB.ComboBox cboTruckType 
         Height          =   315
         Left            =   1140
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1140
         Width           =   4395
      End
      Begin VB.ComboBox cboZone 
         Height          =   315
         Left            =   1140
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1530
         Width           =   4395
      End
      Begin VB.ComboBox cboCompany 
         Height          =   315
         Left            =   1140
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1920
         Width           =   4395
      End
      Begin VB.TextBox txtCustomer 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1140
         TabIndex        =   2
         Top             =   750
         Width           =   4365
      End
      Begin WasteManagment.ctlDate dtFrom 
         Height          =   315
         Left            =   1140
         TabIndex        =   0
         Top             =   360
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         BackColor       =   -2147483633
      End
      Begin WasteManagment.ctlDate dtTo 
         Height          =   315
         Left            =   3780
         TabIndex        =   1
         Top             =   360
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         BackColor       =   -2147483633
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ประเภทรถ"
         Height          =   195
         Index           =   6
         Left            =   330
         TabIndex        =   16
         Top             =   1200
         Width           =   750
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "โซน"
         Height          =   195
         Index           =   5
         Left            =   750
         TabIndex        =   15
         Top             =   1560
         Width           =   300
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "บริษัทที่ออก"
         Height          =   195
         Index           =   4
         Left            =   270
         TabIndex        =   14
         Top             =   1950
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "จนท. Mkt."
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   13
         Top             =   810
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ช่วงวันที่ขน"
         Height          =   195
         Index           =   3
         Left            =   270
         TabIndex        =   12
         Top             =   420
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ถึงวันที่"
         Height          =   195
         Index           =   2
         Left            =   3180
         TabIndex        =   11
         Top             =   420
         Width           =   495
      End
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "พิมพ์"
      Height          =   795
      Left            =   4110
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3900
      Width           =   915
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "ปิด"
      Height          =   795
      Left            =   5070
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3900
      Width           =   915
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "Preview"
      Height          =   795
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3900
      Width           =   915
   End
   Begin Crystal.CrystalReport rpt1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowPrintSetupBtn=   -1  'True
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
      Caption         =   "รายชื่อลูกค้าที่เปิดใหม่จำแนกตามลูกค้า (QUO03)"
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
      TabIndex        =   9
      Top             =   60
      Width           =   4170
   End
   Begin VB.Image Image2 
      Height          =   735
      Index           =   1
      Left            =   0
      Picture         =   "frPrintQUO03.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11280
   End
End
Attribute VB_Name = "frmPrintQUO03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim curStaff_ As BWGCompanyStaff
Dim coltruckType_ As Collection
Dim colZone_ As Collection

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdPreview_Click()
    PrintMain 0
End Sub

Private Sub cmdPrint_Click()
    PrintMain 1
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{Tab}"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    CanAccess = False: CanUpdate = False
    SetPermission Me, CanAccess, CanUpdate
    If CanAccess = False Then
        MsgBox "ท่านไม่มีสิทธิใช้ระบบนี้", vbExclamation, "Warning"
        Unload Me
    End If
    
    PopulateCompany cboCompany, cboIndex
    PopulateZone
    PopulateTruckType
     With frmPicture
        cmdPreview.Picture = .picPreview.Picture
        cmdPrint.Picture = .picPrint.Picture
        cmdClose.Picture = .picExit.Picture
    End With
End Sub

Private Sub PrintMain(printType As Integer)
Dim tmpSql As String
 If optnormal.Value = True Then
    With rpt1
        .LogonInfo(0) = "DSN = " & DSN_NAME & ";UID = " & strUid & ";PWD = " & strPwd
        .ReportFileName = App.Path & "\reports\QUO03.rpt"
        tmpSql = ""
        tmpSql = "SELECT * FROM vw_RptQUO03 " & vbCrLf
        tmpSql = tmpSql & "WHERE FirstWorkDate>='" & dtFrom.ValueYMD & "' And FirstWorkDate<='" & dtTo.ValueYMD & "' "
        If Not curStaff_ Is Nothing Then
            If Trim(curStaff_.id) <> "" Then
                tmpSql = tmpSql & " And UnderSaleStaffID='" & curStaff_.id & "' "
            End If
        End If
        If cboCompany.ListCount > 1 Then
            If cboCompany.ListIndex > 0 Then
                tmpSql = tmpSql & " And CompanyID='" & CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).id & "' "
            End If
        Else
            If cboCompany.ListIndex >= 0 Then
                tmpSql = tmpSql & " And CompanyID='" & CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).id & "' "
            End If
        End If
        If cboZone.ListIndex > 0 Then
            tmpSql = tmpSql & " And SaleZoneID='" & colZone_(cboZone.ListIndex).id & "' "
        End If
        If cboTruckType.ListIndex > 0 Then
            tmpSql = tmpSql & " And TruckTypeID='" & coltruckType_(cboTruckType.ListIndex).id & "' "
        End If
        tmpSql = tmpSql & vbCrLf & " ORDER BY FirstWorkDate,CustomerName"
        .Formulas(0) = "DateRange=" & Chr(34) & "ช่วงวันที่ขนครั้งแรก " & dtFrom.ValueDMY & " ถึงวันที่ " & dtTo.ValueDMY & Chr(34)
        .Formulas(1) = "ShowPrintDate=" & Chr(34) & FormatYMD_to_DMY(TodayDate, "/", "/") & Chr(34)
        .Formulas(2) = "ShowWorkCompany=" & Chr(34) & cboCompany.Text & Chr(34)
        .SQLQuery = tmpSql
        .Destination = printType
        .PrintReport
        If Trim(.LastErrorString) <> "" Then MsgBox .LastErrorString
    End With
ElseIf optincrease.Value = True Then
    With rpt1
           .LogonInfo(0) = "DSN = " & DSN_NAME & ";UID = " & strUid & ";PWD = " & strPwd
           .ReportFileName = App.Path & "\reports\QUO03_A.rpt"
           tmpSql = ""
           tmpSql = "SELECT * FROM vw_RptQUO03_A" & vbCrLf
           tmpSql = tmpSql & "WHERE  workyear>='" & Left(dtFrom.ValueYMD, 4) & "' and workyear<='" & Left(dtTo.ValueYMD, 4) & "' and workdate<='" & dtTo.ValueYMD & "' "
           If Not curStaff_ Is Nothing Then
               If Trim(curStaff_.id) <> "" Then
                   tmpSql = tmpSql & " And UnderSaleStaffID='" & curStaff_.id & "' "
               End If
           End If
           If cboCompany.ListCount > 1 Then
               If cboCompany.ListIndex > 0 Then
                   tmpSql = tmpSql & " And CompanyID='" & CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).id & "' "
               End If
           Else
               If cboCompany.ListIndex >= 0 Then
                   tmpSql = tmpSql & " And CompanyID='" & CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).id & "' "
               End If
           End If
           If cboZone.ListIndex > 0 Then
               tmpSql = tmpSql & " And SaleZoneID='" & colZone_(cboZone.ListIndex).id & "' "
           End If
           If cboTruckType.ListIndex > 0 Then
               tmpSql = tmpSql & " And TruckTypeID='" & coltruckType_(cboTruckType.ListIndex).id & "' "
           End If
           tmpSql = tmpSql & vbCrLf & " ORDER BY firstdate,CompanyName"
           .Formulas(0) = "DateRange=" & Chr(34) & "ช่วงวันที่ขนครั้งแรก 01/01" & Left(dtFrom.ValueYMD, 4) & " ถึงวันที่ " & dtTo.ValueDMY & Chr(34)
           .Formulas(1) = "ShowPrintDate=" & Chr(34) & FormatYMD_to_DMY(TodayDate, "/", "/") & Chr(34)
           .Formulas(2) = "ShowWorkCompany=" & Chr(34) & cboCompany.Text & Chr(34)
           .SQLQuery = tmpSql
           .Destination = printType
           .PrintReport
           If Trim(.LastErrorString) <> "" Then MsgBox .LastErrorString
    End With
Else
End If



End Sub

Private Sub txtCustomer_Validate(Cancel As Boolean)
    If curStaff_ Is Nothing Then Set curStaff_ = New BWGCompanyStaff
    If Trim(txtCustomer.Text) = "" And Trim(curStaff_.StaffFName & " " & curStaff_.StaffLName) <> Trim(txtCustomer.Text) Then
        Set curStaff_ = Nothing
        Set curStaff_ = New BWGCompanyStaff
    Else
        If Trim(curStaff_.StaffFName & " " & curStaff_.StaffLName) <> Trim(txtCustomer.Text) Then
            Set curStaff_ = Nothing
            Set curStaff_ = New BWGCompanyStaff
            Load frmStaffSearch
            With frmStaffSearch
                .getCriteria (txtCustomer.Text)
                .SetCallerForm Me
                .Show 1
            End With
            If curStaff_.id = "" Then txtCustomer.Text = ""
        End If
    End If
End Sub

Public Sub SetDataStaff(selStaff As BWGCompanyStaff, Optional isAssigned As Boolean = True)
Dim tmpRec As New ADODB.Recordset
Dim x%
Dim tmpWaste As BWGWasteDataCR
    Set curStaff_ = Nothing
    Set curStaff_ = selStaff
    If curStaff_ Is Nothing Then Exit Sub
    With curStaff_
        txtCustomer.Text = .StaffFName & " " & .StaffLName
    End With
End Sub



Private Sub PopulateZone()
Dim tmpS As New BWGSearchManager
Dim x%
    Set colZone_ = tmpS.SaleZoneSearch
    cboZone.Clear
    cboZone.AddItem "ทุก Zone"
    For x = 1 To colZone_.Count
        cboZone.AddItem colZone_(x).ZoneName
    Next
    cboZone.ListIndex = 0
End Sub

Private Sub PopulateTruckType()
Dim tmpS As New BWGSearchManager
Dim x%
    Set coltruckType_ = tmpS.TruckTypeSearch
    cboTruckType.Clear
    cboTruckType.AddItem "ทุกประเภทรถ"
    For x = 1 To coltruckType_.Count
        cboTruckType.AddItem coltruckType_(x).TypeDesc
    Next
    cboTruckType.ListIndex = 0
End Sub

