VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmPrintCR01 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "F066"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6135
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   6135
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3800
      Left            =   90
      TabIndex        =   12
      Top             =   780
      Width           =   5925
      Begin VB.OptionButton Option3 
         BackColor       =   &H80000005&
         Caption         =   "เลือกทั้งหมด"
         Height          =   255
         Left            =   1080
         TabIndex        =   5
         Top             =   2400
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H80000005&
         Caption         =   "เลือก ลูกค้าที่ยังไม่มีการขน"
         Height          =   255
         Left            =   1080
         TabIndex        =   7
         Top             =   3120
         Width           =   2235
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H80000005&
         Caption         =   "เลือก ลูกค้าที่มีการขนแล้ว"
         Height          =   255
         Left            =   1080
         TabIndex        =   6
         Top             =   2760
         Width           =   2175
      End
      Begin VB.ComboBox cboZone 
         Height          =   315
         Left            =   1140
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1380
         Width           =   4395
      End
      Begin VB.ComboBox cboCompany 
         Height          =   315
         Left            =   1140
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1890
         Width           =   4395
      End
      Begin VB.TextBox txtCustomer 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1140
         TabIndex        =   2
         Top             =   870
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
         Caption         =   "โซน"
         Height          =   195
         Index           =   5
         Left            =   750
         TabIndex        =   17
         Top             =   1410
         Width           =   300
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "บริษัทที่ออก"
         Height          =   195
         Index           =   4
         Left            =   270
         TabIndex        =   16
         Top             =   1920
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "จนท. Mkt."
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   15
         Top             =   930
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ช่วงวันที่"
         Height          =   195
         Index           =   3
         Left            =   450
         TabIndex        =   14
         Top             =   420
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ถึงวันที่"
         Height          =   195
         Index           =   2
         Left            =   3180
         TabIndex        =   13
         Top             =   420
         Width           =   495
      End
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "พิมพ์"
      Height          =   795
      Left            =   4110
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4620
      Width           =   915
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "ปิด"
      Height          =   795
      Left            =   5070
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4620
      Width           =   915
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "Preview"
      Height          =   795
      Left            =   3150
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4620
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
      Caption         =   "รายงานจำนวนใบเสนอราคาและจำนวนเที่ยว (CR01)"
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
      TabIndex        =   11
      Top             =   60
      Width           =   4500
   End
   Begin VB.Image Image2 
      Height          =   735
      Index           =   1
      Left            =   0
      Picture         =   "frPrintCR01.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11280
   End
End
Attribute VB_Name = "frmPrintCR01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim curStaff_ As BWGCompanyStaff
Dim colZone_ As Collection
Dim sessionID_  As String

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
    sessionID_ = GetGUID
    PopulateCompany cboCompany, cboIndex
    PopulateZone
     With frmPicture
        cmdPreview.Picture = .picPreview.Picture
        cmdPrint.Picture = .picPrint.Picture
        cmdClose.Picture = .picExit.Picture
    End With
End Sub

Private Sub PrintMain(printType As Integer)
On Error GoTo ErrD
Dim tmpSql As String, tmpRec As New ADODB.Recordset
Dim tmpQuo As New ADODB.Recordset
Dim tmpCon As ADODB.Connection, iCount%, x%, tmpCol As New Collection
Dim tmpCust As BWGCustomer
    With rpt1
                    tmpSql = "SELECT * FROM vw_RptQUO01 WHERE IssueDate>='" & dtFrom.ValueYMD & "' And IssueDate<='" & dtTo.ValueYMD & "' "
                    If Not curStaff_ Is Nothing Then
                        If Trim(curStaff_.ID) <> "" Then
                            tmpSql = tmpSql & " And StaffID='" & curStaff_.ID & "' "
                        End If
                    End If
                    If cboCompany.ListCount > 1 Then
                        If cboCompany.ListIndex > 0 Then
                            tmpSql = tmpSql & " And CompanyID='" & CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).ID & "' "
                        End If
                    Else
                        If cboCompany.ListIndex >= 0 Then
                            tmpSql = tmpSql & " And CompanyID='" & CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).ID & "' "
                        End If
                    End If
                    If cboZone.ListIndex > 0 Then
                        tmpSql = tmpSql & " And SaleZoneID='" & colZone_(cboZone.ListIndex).ID & "' "
                    End If
                    If Option1.Value = True Then
                        tmpSql = tmpSql & " And CustomerID  IN (Select CustomerID From tbMenifestHeader)"
                    ElseIf Option2.Value = True Then
                         tmpSql = tmpSql & " And CustomerID NOT IN (Select CustomerID From tbMenifestHeader) And JobCount = '0' "
                    End If
                    Set tmpRec = GetRS(tmpSql, adOpenForwardOnly, adLockReadOnly)
                    If tmpRec Is Nothing Then Exit Sub
                    DBConnExc "Delete tbPrintCR01 WHERE SessionID='" & sessionID_ & "'"
                    GetDBConnection
                    Do Until tmpRec.EOF
                        iCount = 0
                        If isExist(tmpCol, "'" & Trim("" & tmpRec("CustomerID")) & "_" & Trim("" & tmpRec("SaleDistrictID")) & "_" & Trim("" & tmpRec("StaffID")) & "'") Then
                            x = 0
                        Else
                            x = 1
                            Set tmpCust = New BWGCustomer
                            tmpCust.ID = Trim("" & tmpRec("CustomerID"))
                            tmpCol.Add tmpCust, "'" & Trim("" & tmpRec("CustomerID")) & "_" & Trim("" & tmpRec("SaleDistrictID")) & "_" & Trim("" & tmpRec("StaffID")) & "'"
                            Set tmpCust = Nothing
                        End If
                        DBConn.Execute "Update tbPrintCR01 Set QuoCount=QuoCount+1,CustCount=CustCount+" & x & " WHERE SaleStaffID='" & tmpRec("StaffID") & "' And SaleDisctrictID='" & tmpRec("SaleDistrictID") & "' And SessionID='" & sessionID_ & "'", iCount
                        If iCount = 0 Then
                            DBConn.Execute "Insert into tbPrintCR01 (SaleStaffID,SaleDisctrictID,SessionID,SaleZoneID,SaleFullName,SDName,ZoneName,QuoCount,CustCount,NewCustCount,NewTripCount) Values (" & _
                                                    "'" & tmpRec("StaffID") & "','" & tmpRec("SaleDistrictID") & "','" & sessionID_ & "','" & tmpRec("SaleZoneID") & "','" & tmpRec("StaffFName") & " " & tmpRec("StaffLName") & "','" & tmpRec("SDName") & "','" & tmpRec("ZoneName") & "',1," & x & ",0,0)"
                        End If
                        tmpRec.MoveNext
                    Loop
                    tmpRec.Close
                    
                    tmpSql = "SELECT * FROM vw_RptQUO03 WHERE FirstWorkDate>='" & dtFrom.ValueYMD & "' And FirstWorkDate<='" & dtTo.ValueYMD & "' "
                    If Not curStaff_ Is Nothing Then
                        If Trim(curStaff_.ID) <> "" Then
                            tmpSql = tmpSql & " And UnderSaleStaffID='" & curStaff_.ID & "' "
                        End If
                    End If
                    If cboCompany.ListCount > 1 Then
                        If cboCompany.ListIndex > 0 Then
                            tmpSql = tmpSql & " And CompanyID='" & CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).ID & "' "
                        End If
                    Else
                        If cboCompany.ListIndex >= 0 Then
                            tmpSql = tmpSql & " And CompanyID='" & CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).ID & "' "
                        End If
                    End If
                    If cboZone.ListIndex > 0 Then
                        tmpSql = tmpSql & " And SaleZoneID='" & colZone_(cboZone.ListIndex).ID & "' "
                    End If
                    If Option1.Value = True Then
                        tmpSql = tmpSql & " And CustomerID  IN (Select CustomerID From tbMenifestHeader)"
                    ElseIf Option2.Value = True Then
                         tmpSql = tmpSql & " And CustomerID NOT IN (Select CustomerID From tbMenifestHeader)"
                    End If
                    Set tmpRec = GetRS(tmpSql, adOpenForwardOnly, adLockReadOnly)
                    If tmpRec Is Nothing Then Exit Sub
                    Set tmpCol = Nothing
                    Set tmpCol = New Collection
                    CloseDBConnection
                    GetDBConnection
                    Do Until tmpRec.EOF
                        iCount = 0
                        If isExist(tmpCol, "'" & Trim("" & tmpRec("CustomerID")) & "_" & Trim("" & tmpRec("SaleDistrictID")) & "_" & Trim("" & tmpRec("UnderSaleStaffID")) & "'") Then
                            x = 0
                        Else
                            x = 1
                            Set tmpCust = New BWGCustomer
                            tmpCust.ID = Trim("" & tmpRec("CustomerID"))
                            tmpCol.Add tmpCust, "'" & Trim("" & tmpRec("CustomerID")) & "_" & Trim("" & tmpRec("SaleDistrictID")) & "_" & Trim("" & tmpRec("UnderSaleStaffID")) & "'"
                            Set tmpCust = Nothing
                        End If
                        DBConn.Execute "Update tbPrintCR01 Set NewCustCount=NewCustCount+" & x & ",NewTripCount=NewTripCount+" & tmpRec("TruckCount") & " WHERE SaleStaffID='" & tmpRec("UnderSaleStaffID") & "' And SaleDisctrictID='" & tmpRec("SaleDistrictID") & "' And SessionID='" & sessionID_ & "'", iCount
                        If iCount = 0 Then
                            DBConn.Execute "Insert into tbPrintCR01 (SaleStaffID,SaleDisctrictID,SessionID,SaleZoneID,SaleFullName,SDName,ZoneName,QuoCount,CustCount,NewCustCount,NewTripCount) Values (" & _
                                                    "'" & tmpRec("UnderSaleStaffID") & "','" & tmpRec("SaleDistrictID") & "','" & sessionID_ & "','" & tmpRec("SaleZoneID") & "','" & tmpRec("StaffFName") & " " & tmpRec("StaffLName") & "','" & tmpRec("SDName") & "','" & tmpRec("ZoneName") & "',0,0," & x & "," & tmpRec("TruckCount") & ")"
                        End If
                        tmpRec.MoveNext
                    Loop
                     tmpRec.Close
        Set tmpRec = Nothing
        Set tmpCol = Nothing
        Set tmpCust = Nothing
        Set tmpCol = Nothing
        
        .LogonInfo(0) = "DSN = " & DSN_NAME & ";UID = " & strUid & ";PWD = " & strPwd
        .ReportFileName = App.Path & "\reports\CR01.rpt"
        tmpSql = ""
        tmpSql = "SELECT * FROM tbPrintCR01 " & vbCrLf
        tmpSql = tmpSql & "WHERE SessionID='" & sessionID_ & "' "
        tmpSql = tmpSql & vbCrLf & " ORDER BY ZoneName,SaleFullName"
        .Formulas(0) = "DateRange=" & Chr(34) & "ช่วงวันที่ " & dtFrom.ValueDMY & " ถึงวันที่ " & dtTo.ValueDMY & Chr(34)
        .Formulas(1) = "ShowPrintDate=" & Chr(34) & FormatYMD_to_DMY(TodayDate, "/", "/") & Chr(34)
        .Formulas(2) = "ShowWorkCompany=" & Chr(34) & cboCompany.Text & Chr(34)
        .SQLQuery = tmpSql
        .Destination = printType
        .PrintReport
        If Trim(.LastErrorString) <> "" Then MsgBox .LastErrorString
    End With
    CloseDBConnection
    Exit Sub
ErrD:
    GetMsg "", vbCritical
    Exit Sub
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    DBConnExc "Delete tbPrintCR01 WHERE SessionID='" & sessionID_ & "'"
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
            If curStaff_.ID = "" Then txtCustomer.Text = ""
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

