VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frPrintORD01 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " "
   ClientHeight    =   3930
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6060
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3930
   ScaleWidth      =   6060
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2145
      Left            =   60
      TabIndex        =   9
      Top             =   810
      Width           =   5925
      Begin VB.ComboBox cboCompany 
         Height          =   315
         Left            =   1140
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   750
         Width           =   4455
      End
      Begin VB.ComboBox cboYear 
         Height          =   315
         Left            =   3780
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   360
         Width           =   1815
      End
      Begin VB.ComboBox cboMonth 
         Height          =   315
         Left            =   1140
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   360
         Width           =   1815
      End
      Begin VB.TextBox txtSale 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1140
         TabIndex        =   4
         Top             =   1530
         Width           =   4425
      End
      Begin VB.ComboBox cboZone 
         Height          =   315
         Left            =   1140
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1140
         Width           =   4455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "บริษัท"
         Height          =   195
         Index           =   2
         Left            =   600
         TabIndex        =   14
         Top             =   810
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ปี"
         Height          =   195
         Index           =   0
         Left            =   3570
         TabIndex        =   13
         Top             =   420
         Width           =   105
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "โซน"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   8
         Left            =   750
         TabIndex        =   12
         Top             =   1200
         Width           =   300
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ผู้แทนขาย"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   330
         TabIndex        =   11
         Top             =   1590
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "เดือน"
         Height          =   195
         Index           =   3
         Left            =   660
         TabIndex        =   10
         Top             =   420
         Width           =   390
      End
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "พิมพ์"
      Height          =   795
      Left            =   4110
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3060
      Width           =   915
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "ปิด"
      Height          =   795
      Left            =   5070
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3060
      Width           =   915
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "Preview"
      Height          =   795
      Left            =   3150
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3060
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
      Caption         =   "จำนวนเที่ยวที่เปิด Order การขนส่งรายเดือน (ORD01)"
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
      TabIndex        =   8
      Top             =   60
      Width           =   4665
   End
   Begin VB.Image Image2 
      Height          =   735
      Index           =   1
      Left            =   0
      Picture         =   "frPrintORD01.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11280
   End
End
Attribute VB_Name = "frPrintORD01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim curStaff_ As BWGCompanyStaff
Dim colZone_ As Collection
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

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{Tab}"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
Dim x%
    sessionID_ = GetGUID
    PopulateCompany
    PopulateZone
     With frmPicture
        cmdPreview.Picture = .picPreview.Picture
        cmdPrint.Picture = .picPrint.Picture
        cmdClose.Picture = .picExit.Picture
    End With
    With cboMonth
        .AddItem "มกราคม"
        .AddItem "กุมภาพันธ์"
        .AddItem "มีนาคม"
        .AddItem "เมษายน"
        .AddItem "พฤษภาคม"
        .AddItem "มิถุนายน"
        .AddItem "กรกฎาคม"
        .AddItem "สิงหาคม"
        .AddItem "กันยายน"
        .AddItem "ตุลาคม"
        .AddItem "พฤศจิกายน"
        .AddItem "ธันวาคม"
        .ListIndex = CInt(Mid(TodayDate, 6, 2)) - 1
    End With
    With cboYear
        For x = 1 To 10
            .AddItem (CInt(Left(TodayDate, 4)) - 10 + x)
        Next
        SetTextToCombo Left(TodayDate, 4), cboYear
    End With
End Sub

Private Sub PrintMain(printType As Integer)
Dim tmpSql As String, iCount As Long
Dim tmpCon As ADODB.Connection, tmpRec As New ADODB.Recordset
    Set tmpCon = DBManager.GetDBConnection
    tmpSql = "Select * from vw_PrintOrd01TripCount Where WorkDate>='" & cboYear.Text & "/" & Format(cboMonth.ListIndex + 1, "00") & "/01' And WorkDate<='" & cboYear.Text & "/" & Format(cboMonth.ListIndex + 1, "00") & "/31'"
    tmpSql = tmpSql & " And CompanyID='" & CurrentUser.WorkCompanies(cboCompany.ListIndex + 1).ID & "'"
    If cboZone.ListIndex > 0 Then
        tmpSql = tmpSql & " And SaleZoneID='" & colZone_(cboZone.ListIndex).ID & "'"
    End If
    If Trim(txtSale.Text) <> "" Then
        tmpSql = tmpSql & " And UnderSaleStaffID='" & curStaff_.ID & "'"
    End If
    tmpRec.Open tmpSql, tmpCon, adOpenStatic, adLockReadOnly
    tmpCon.Execute "Delete tbPrintOrd01 Where SessionID='" & sessionID_ & "'"
    iCount = 1
    Do Until tmpRec.EOF
        tmpCon.Execute "Insert into tbPrintOrd01 (AutoKey,CompanyID,Day" & Trim(Format(Right(Trim(tmpRec("WorkDate")), 2), "##")) & ",SessionID,SaleStaffID,SaleDistrictID) Values (" & _
         iCount & ",'" & tmpRec("CustomerID") & "'," & tmpRec("TruckCount") & ",'" & sessionID_ & "','" & tmpRec("UnderSaleStaffID") & "','" & tmpRec("SaleDistrictID") & "')"
        tmpRec.MoveNext
        iCount = iCount + 1
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set tmpCon = Nothing
    With rpt1
        .LogonInfo(0) = "DSN = " & DSN_NAME & ";UID = " & uID & ";PWD = " & uPwd
        .ReportFileName = App.Path & "\reports\ORD01.rpt"
        tmpSql = ""
        tmpSql = "Select * from vw_PrintOrd01 " & vbCrLf
        tmpSql = tmpSql & "Where SessionID='" & sessionID_ & "' "
        tmpSql = tmpSql & vbCrLf & " Order By StaffFName,StaffLName,ZoneName,CompanyName"
        .Formulas(0) = "DateRange=" & Chr(34) & cboMonth.Text & "  พ.ศ. " & cboYear.Text & Chr(34)
        .Formulas(1) = "ShowPrintDate=" & Chr(34) & FormatYMD_to_DMY(TodayDate, "/", "/") & Chr(34)
        .Formulas(2) = "ShowCompany=" & Chr(34) & cboCompany.Text & Chr(34)
        .SQLQuery = tmpSql
        .Destination = printType
        .PrintReport
        If Trim(.LastErrorString) <> "" Then MsgBox .LastErrorString
    End With
End Sub

Public Sub SetDataStaff(selStaff As BWGCompanyStaff, Optional isAssigned As Boolean = True)
Dim tmpRec As New ADODB.Recordset
Dim x%
    Set curStaff_ = Nothing
    Set curStaff_ = selStaff
    If curStaff_ Is Nothing Then Exit Sub
    With curStaff_
        txtSale.Text = .StaffFName & " " & .StaffLName
    End With
End Sub

Private Sub PopulateZone()
Dim x%, tmpS As New BWGSearchManager
    cboZone.Clear
    cboZone.AddItem "ทุกโซน"
    Set colZone_ = tmpS.SaleZoneSearch
    For x = 1 To colZone_.Count
        cboZone.AddItem colZone_(x).ZoneName
    Next
    If cboZone.ListCount > 0 Then
        cboZone.ListIndex = 0
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    DBManager.GetDBConnection.Execute "Delete tbPrintOrd01 Where SessionID='" & sessionID_ & "'"
End Sub

Private Sub txtSale_Validate(Cancel As Boolean)
    If curStaff_ Is Nothing Then Set curStaff_ = New BWGCompanyStaff
    If Trim(txtSale.Text) = "" And Trim(curStaff_.StaffFName & " " & curStaff_.StaffLName) <> Trim(txtSale.Text) Then
        Set curStaff_ = Nothing
        Set curStaff_ = New BWGCompanyStaff
    Else
        If Trim(curStaff_.StaffFName & " " & curStaff_.StaffLName) <> Trim(txtSale.Text) Then
            Set curStaff_ = Nothing
            Set curStaff_ = New BWGCompanyStaff
            Load frStaffSearch
            With frStaffSearch
                .getCriteria (txtSale.Text)
                .setCallerForm Me
                .Show 1
            End With
            If curStaff_.ID = "" Then txtSale.Text = ""
        End If
    End If
End Sub

Private Sub PopulateCompany()
Dim x%
    cboCompany.Clear
    For x = 1 To CurrentUser.WorkCompanies.Count
        cboCompany.AddItem CurrentUser.WorkCompanies(x).CompanyNameTH
    Next
    If cboCompany.ListCount > 0 Then
        cboCompany.ListIndex = 0
    End If
End Sub

