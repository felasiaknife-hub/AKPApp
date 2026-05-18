VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmPrintTRN05 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " "
   ClientHeight    =   5235
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6060
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5235
   ScaleWidth      =   6060
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3500
      Left            =   60
      TabIndex        =   9
      Top             =   780
      Width           =   5925
      Begin VB.ComboBox cboWasteCR 
         Height          =   315
         Left            =   1110
         TabIndex        =   18
         Text            =   "Combo2"
         Top             =   2670
         Width           =   1815
      End
      Begin VB.ComboBox cboWaste 
         Height          =   315
         Left            =   1110
         TabIndex        =   17
         Text            =   "Combo1"
         Top             =   2280
         Width           =   1815
      End
      Begin VB.ComboBox cboCustType 
         Height          =   315
         Left            =   1110
         TabIndex        =   15
         Top             =   1890
         Width           =   1815
      End
      Begin VB.TextBox txtCustomer 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1110
         TabIndex        =   4
         Top             =   1500
         Width           =   4395
      End
      Begin VB.ComboBox cboType 
         Height          =   315
         Left            =   1140
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1110
         Width           =   4395
      End
      Begin VB.ComboBox cboCompany 
         Height          =   315
         Left            =   1140
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   720
         Width           =   4395
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
         Caption         =   "WasteCR"
         Height          =   195
         Index           =   7
         Left            =   350
         TabIndex        =   20
         Top             =   2715
         Width           =   690
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Waste"
         Height          =   195
         Index           =   6
         Left            =   550
         TabIndex        =   19
         Top             =   2325
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ประเภทลูกค้า"
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   16
         Top             =   1940
         Width           =   930
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ลูกค้า"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   660
         TabIndex        =   14
         Top             =   1560
         Width           =   375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "บริษัท"
         Height          =   195
         Index           =   0
         Left            =   600
         TabIndex        =   13
         Top             =   780
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ประเภทกาก"
         Height          =   195
         Index           =   4
         Left            =   210
         TabIndex        =   12
         Top             =   1170
         Width           =   840
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ช่วงวันที่"
         Height          =   195
         Index           =   3
         Left            =   420
         TabIndex        =   10
         Top             =   420
         Width           =   585
      End
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "พิมพ์"
      Height          =   795
      Left            =   4110
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4300
      Width           =   915
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "ปิด"
      Height          =   795
      Left            =   5070
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4300
      Width           =   915
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "Preview"
      Height          =   795
      Left            =   3150
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4300
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
      TabIndex        =   8
      Top             =   60
      Width           =   4200
   End
   Begin VB.Image Image2 
      Height          =   735
      Index           =   1
      Left            =   0
      Picture         =   "frPrintCR06.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11280
   End
End
Attribute VB_Name = "frmPrintTRN05"
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
    PopulateCompany cboCompany, cboIndex
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
    With cboCustType
        .AddItem "== ประเภทลูกค้า =="
        .AddItem "(C) ,(s)"
        .AddItem "(F) RDF Waste"
        .AddItem "(L) (p)"
        .AddItem "(T) (k)"
        .AddItem "(G) (e)"
        .AddItem "(R)"
        .ListIndex = 0
    End With
     With cboWaste
        .AddItem "== ประเภท Waste =="
        .AddItem "New"
        .AddItem "ปกติ"
        .ListIndex = 0
    End With
    With cboWasteCR
        .AddItem "== ประเภท Waste CR =="
        .AddItem "New"
        .AddItem "ปกติ"
        .ListIndex = 0
    End With
End Sub


Private Sub PrintMain(printType As Integer)
Dim tmpSql As String
    With rpt1
        .LogonInfo(0) = "DSN = " & DSN_NAME & ";UID = " & strUid & ";PWD = " & strPwd
        .ReportFileName = App.Path & "\reports\TRN05.rpt"
        tmpSql = ""
        tmpSql = "SELECT * FROM vw_RptTRN05 " & vbCrLf
        tmpSql = tmpSql & "WHERE WorkDate>='" & dtFrom.ValueYMD & "' And WorkDate<='" & dtTo.ValueYMD & "' "
        
        If cboCompany.ListCount > 1 Then
            If cboCompany.ListIndex > 0 Then
                If CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).CompanyCode = "AKP" Then
                    tmpSql = tmpSql & " AND CompanyID='" & CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).ID & "' "
                    tmpSql = tmpSql & " AND DocNo NOT LIKE ('%PST%')"
                ElseIf CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).CompanyCode = "PST" Then
                    tmpSql = tmpSql & "AND CompanyID='" & CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).ID & "' "
                Else
                    tmpSql = tmpSql & "AND CompanyID='" & CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).ID & "' "
                    tmpSql = tmpSql & " AND DocNo NOT LIKE ('%PST%')"
                End If
            End If
        Else
            If cboCompany.ListIndex >= 0 Then
                If CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).CompanyCode = "AKP" Then
                    tmpSql = tmpSql & "And CompanyID = '" & CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).ID & "' "
                    tmpSql = tmpSql & " AND DocNo NOT LIKE ('%PST%')"
                ElseIf CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).CompanyCode = "PST" Then
                    tmpSql = tmpSql & "And CompanyID = '" & CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).ID & "' "
                Else
                    tmpSql = tmpSql & "And CompanyID = '" & CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).ID & "' "
                    tmpSql = tmpSql & " AND DocNo NOT LIKE ('%PST%')"
                End If
            End If
        End If
        
        If cboType.ListIndex > 0 Then
            tmpSql = tmpSql & " AND WasteType='" & cboType.Text & "' "
        End If
        If cboCustType.ListIndex = 1 Then
        tmpSql = tmpSql & " AND CustTypeFlag in ('C','S') "
    ElseIf cboCustType.ListIndex = 2 Then
        tmpSql = tmpSql & " AND CustTypeFlag = 'F' "
   ElseIf cboCustType.ListIndex = 3 Then
        tmpSql = tmpSql & " AND CustTypeFlag in ('L','P') "
    ElseIf cboCustType.ListIndex = 4 Then
        tmpSql = tmpSql & " AND CustTypeFlag in ('T','K') "
    ElseIf cboCustType.ListIndex = 5 Then
        tmpSql = tmpSql & " AND CustTypeFlag in ('G','E') "
    ElseIf cboCustType.ListIndex = 6 Then
        tmpSql = tmpSql & " AND CustTypeFlag in ('R') "
    End If
        
        If Not curCustomer_ Is Nothing Then
            If Trim(curCustomer_.ID) <> "" Then
                tmpSql = tmpSql & " And CustomerID='" & curCustomer_.ID & "' "
            End If
        End If
        
        If cboWaste.ListIndex > 0 Then
            If cboWaste.ListIndex = 1 Then
               tmpSql = tmpSql & " And IsNewWaste='New' "
            Else
                 tmpSql = tmpSql & " And IsNewWaste='' "
            End If
        End If
         If cboWasteCR.ListIndex > 0 Then
            If cboWasteCR.ListIndex = 1 Then
               tmpSql = tmpSql & " And IsNewWasteCR='New' "
            Else
                 tmpSql = tmpSql & " And IsNewWasteCR='' "
            End If
        End If
        tmpSql = tmpSql & vbCrLf & " ORDER BY WorkDate,CompanyName"
        .Formulas(0) = "WasteType=" & Chr(34) & cboType.Text & Chr(34)
        .Formulas(1) = "DateRange=" & Chr(34) & dtFrom.ValueDMY & " ถึงวันที่ " & dtTo.ValueDMY & Chr(34)
        .Formulas(2) = "CompanyName=" & Chr(34) & cboCompany.Text & Chr(34)
        .SQLQuery = tmpSql
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

