VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form FrmPrintCR06 
   Caption         =   "Form1"
   ClientHeight    =   5505
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6285
   LinkTopic       =   "Form1"
   ScaleHeight     =   5505
   ScaleWidth      =   6285
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPreview 
      Caption         =   "Preview"
      Height          =   795
      Left            =   3150
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4300
      Width           =   915
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "ปิด"
      Height          =   795
      Left            =   5070
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4300
      Width           =   915
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "พิมพ์"
      Height          =   795
      Left            =   4110
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4300
      Width           =   915
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3500
      Left            =   60
      TabIndex        =   0
      Top             =   780
      Width           =   5925
      Begin VB.ComboBox cboRPT 
         Height          =   315
         Left            =   1140
         TabIndex        =   17
         Text            =   "cboRPT"
         Top             =   1920
         Width           =   4395
      End
      Begin VB.ComboBox cboCompany 
         Height          =   315
         Left            =   1140
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   720
         Width           =   4395
      End
      Begin VB.ComboBox cboType 
         Height          =   315
         Left            =   1140
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1080
         Width           =   4395
      End
      Begin VB.TextBox txtCustomer 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1140
         TabIndex        =   1
         Top             =   1500
         Width           =   4395
      End
      Begin WasteManagment.ctlDate dtFrom 
         Height          =   315
         Left            =   1140
         TabIndex        =   14
         Top             =   360
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         BackColor       =   -2147483633
      End
      Begin WasteManagment.ctlDate dtTo 
         Height          =   315
         Left            =   3840
         TabIndex        =   15
         Top             =   360
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         BackColor       =   -2147483633
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "รายงาน"
         Height          =   255
         Left            =   480
         TabIndex        =   16
         Top             =   1920
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ช่วงวันที่"
         Height          =   195
         Index           =   3
         Left            =   420
         TabIndex        =   8
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
         TabIndex        =   7
         Top             =   420
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ประเภทกาก"
         Height          =   195
         Index           =   4
         Left            =   210
         TabIndex        =   6
         Top             =   1170
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "บริษัท"
         Height          =   195
         Index           =   0
         Left            =   600
         TabIndex        =   5
         Top             =   780
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ลูกค้า"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   660
         TabIndex        =   4
         Top             =   1560
         Width           =   375
      End
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "รายงาน ใบเสนอราคา และ Order ปลายทางอื่น"
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
      Index           =   8
      Left            =   480
      TabIndex        =   13
      Top             =   0
      Width           =   3915
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      Index           =   3
      X1              =   90
      X2              =   5790
      Y1              =   510
      Y2              =   510
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   4
      Index           =   2
      X1              =   0
      X2              =   5640
      Y1              =   420
      Y2              =   420
   End
   Begin VB.Image Image2 
      Height          =   735
      Index           =   1
      Left            =   0
      Picture         =   "FrmPrintCR06.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11280
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "รายงาน Order ที่มีการยืนยันการขนกาก (TRN05)"
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
      Left            =   120
      TabIndex        =   12
      Top             =   60
      Width           =   4200
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
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   4
      Index           =   0
      X1              =   120
      X2              =   5760
      Y1              =   480
      Y2              =   480
   End
End
Attribute VB_Name = "FrmPrintCR06"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim curCustomer_ As BWGCustomer

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
Dim x As Integer
    PopulateCompany
     With frmPicture
        cmdPreview.Picture = .picPreview.Picture
        cmdPrint.Picture = .picPrint.Picture
        cmdClose.Picture = .picExit.Picture
    End With
    With cboType
        .AddItem "ทุกประเภท"
        For x = 1 To WasteTypes.Count
            .AddItem WasteTypes(x).WasteType
        Next
        .ListIndex = 0
    End With
     With cboRPT
        .AddItem "ใบเสนอราคาปลายทางอื่น Detail"
        .AddItem "ใบเสนอราคาปลายทางอื่น สรุป"
       .AddItem "รายการขน Detail"
       .AddItem "รายการขน สรุป"
        .ListIndex = 0
    End With
  
End Sub
Private Sub PopulateCompany()
Dim x%
    cboCompany.Clear
    cboCompany.AddItem "== เลือกบริษัท =="
    For x = 1 To CurrentUser.WasteCompany.Count
        If CurrentUser.WasteCompany(x).isDeleted <> "Y" Then
        cboCompany.AddItem CurrentUser.WasteCompany(x).CompanyNameTH
        End If
        'If CurrentUser.WasteCompany(x).CompanyCode = "AKP" Then cboCompany.AddItem CurrentUser.WasteCompany(x).CompanyNameTH
        'If CurrentUser.WasteCompany(x).CompanyCode = "BWG" Then cboCompany.AddItem CurrentUser.WasteCompany(x).CompanyNameTH
        'If CurrentUser.WasteCompany(x).CompanyCode = "BWC" Then cboCompany.AddItem CurrentUser.WasteCompany(x).CompanyNameTH
    Next
        cboCompany.ListIndex = 0
End Sub


Private Sub PrintMain(printType As Integer)
Dim tmpSql As String
    With rpt1
        '.LogonInfo(0) = "DSN = " & DSN_NAME & ";UID = " & strUid & ";PWD = " & strPwd
        If cboRPT.ListIndex = 0 Then
            .ReportFileName = App.Path & "\reports\CR061.rpt"
             .SelectionFormula = "{tbQuotationHeader.IssueDate} in " & Chr(34) & dtFrom.ValueYMD & Chr(34) & " to " & Chr(34) & dtTo.ValueYMD & Chr(34)
        ElseIf cboRPT.ListIndex = 1 Then
            .ReportFileName = App.Path & "\reports\CR062.rpt"
             .SelectionFormula = "{tbQuotationHeader.IssueDate} in " & Chr(34) & dtFrom.ValueYMD & Chr(34) & " to " & Chr(34) & dtTo.ValueYMD & Chr(34)
        ElseIf cboRPT.ListIndex = 2 Then
            .ReportFileName = App.Path & "\reports\CR063.rpt"
             .SelectionFormula = "{vw_MenifestDetail.WorkDate}  in " & Chr(34) & dtFrom.ValueYMD & Chr(34) & " to " & Chr(34) & dtTo.ValueYMD & Chr(34)
        Else
            .ReportFileName = App.Path & "\reports\CR064.rpt"
             .SelectionFormula = " {vw_MenifestDetail.WorkDate}  in " & Chr(34) & dtFrom.ValueYMD & Chr(34) & " to " & Chr(34) & dtTo.ValueYMD & Chr(34)
        End If
        .Connect = "DSN = " & DSN_NAME & ";UID = " & strUid & ";PWD = " & strPwd
     
        'tmpSql = tmpSql & vbCrLf & " ORDER BY WorkDate,CompanyName"
        '.Formulas(0) = "WasteType=" & Chr(34) & cboType.Text & Chr(34)
        .Formulas(0) = "DateRange=" & Chr(34) & dtFrom.ValueDMY & " ถึงวันที่ " & dtTo.ValueDMY & Chr(34)
        .Formulas(1) = "CompanyName=" & Chr(34) & cboCompany.Text & Chr(34)
        '.SQLQuery = tmpSql
       '.SelectionFormula = "{tbQuotationHeader.IssueDate} in '" & dtFrom.ValueYMD & "' to '" & dtTo.ValueYMD & "' "
        If cboCompany.ListIndex > 0 Then
            .SelectionFormula = .SelectionFormula & " and {tbWasteDataCR.CompanyID}='" & CurrentUser.WasteCompany(cboCompany.ListIndex).ID & "'"
        End If
        If cboType.ListIndex > 0 Then
             .SelectionFormula = .SelectionFormula & " and {tbWasteDataCR.WasteType} ='" & cboType.Text & "'"
        End If
            If Not curCustomer_ Is Nothing Then
            If Trim(curCustomer_.ID) <> "" Then
            .SelectionFormula = .SelectionFormula & " and {tbCustomer.CustomerID} ='" & curCustomer_.ID & "'""' "
            End If
        End If
        .Destination = printType
        .PrintReport
        If Trim(.LastErrorString) <> "" Then MsgBox .LastErrorString
    End With
End Sub

Public Sub SetDataCustomer(selCustomer As BWGCustomer, Optional isAssigned As Boolean = True)
Dim tmpRec As New ADODB.Recordset
Dim x%
    Set curCustomer_ = Nothing
    Set curCustomer_ = selCustomer
    If curCustomer_ Is Nothing Then Exit Sub
    With curCustomer_
        txtCustomer.Text = .CustomerName
    End With
End Sub

Private Sub txtCustomer_Validate(Cancel As Boolean)
    If curCustomer_ Is Nothing Then Set curCustomer_ = New BWGCustomer
    If Trim(txtCustomer.Text) = "" And Trim(curCustomer_.CustomerName) <> Trim(txtCustomer.Text) Then
        Set curCustomer_ = Nothing
        Set curCustomer_ = New BWGCustomer
    Else
        If Trim(curCustomer_.CustomerName) <> Trim(txtCustomer.Text) Then
            Set curCustomer_ = Nothing
            Set curCustomer_ = New BWGCustomer
            Load frmCustomerSearch
            With frmCustomerSearch
                .getCriteria (txtCustomer.Text)
                .SetCallerForm Me
                .Show 1
            End With
            If Trim(curCustomer_.ID) = "" Then txtCustomer.Text = ""
        End If
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{Tab}"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0
End Sub


