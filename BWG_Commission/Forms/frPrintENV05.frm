VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frPrintENV05 
   BackColor       =   &H00FFFFFF&
   Caption         =   " "
   ClientHeight    =   3225
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6855
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3225
   ScaleWidth      =   6855
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport Rpt1 
      Left            =   900
      Top             =   2700
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "Preview"
      Height          =   375
      Left            =   3000
      TabIndex        =   5
      Top             =   2760
      Width           =   1155
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "พิมพ์"
      Height          =   375
      Left            =   4320
      TabIndex        =   6
      Top             =   2760
      Width           =   1155
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "ปิด"
      Height          =   375
      Left            =   5640
      TabIndex        =   7
      Top             =   2760
      Width           =   1155
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1815
      Left            =   60
      TabIndex        =   1
      Top             =   840
      Width           =   6735
      Begin VB.ComboBox cboCompany 
         Height          =   315
         Left            =   1140
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   780
         Width           =   4275
      End
      Begin Project1.ctlDate dtFrom 
         Height          =   345
         Left            =   1140
         TabIndex        =   2
         Top             =   360
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   609
         BackColor       =   -2147483633
      End
      Begin Project1.ctlDate dtTo 
         Height          =   345
         Left            =   3690
         TabIndex        =   3
         Top             =   360
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   609
         BackColor       =   -2147483633
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "บริษัท"
         Height          =   195
         Index           =   2
         Left            =   600
         TabIndex        =   10
         Top             =   840
         Width           =   420
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ช่วงวันที่ขน"
         Height          =   195
         Index           =   0
         Left            =   270
         TabIndex        =   9
         Top             =   420
         Width           =   795
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ถึงวันที่"
         Height          =   195
         Index           =   1
         Left            =   3090
         TabIndex        =   8
         Top             =   420
         Width           =   495
      End
   End
   Begin VB.Label AumphurSearch 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "แบบแสดงรหัสของชนิดและประเภทสิ่งปฏิกูลหรือวัสดุไม่ใช้แล้ว"
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
      Index           =   7
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   5355
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      Index           =   1
      X1              =   0
      X2              =   5805
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   4
      Index           =   0
      X1              =   0
      X2              =   5715
      Y1              =   540
      Y2              =   540
   End
   Begin VB.Image Image2 
      Height          =   795
      Left            =   0
      Picture         =   "frPrintENV05.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9195
   End
End
Attribute VB_Name = "frPrintENV05"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type TSDFCom
    CompanyName As String
    TSDFRegisNo As String
    Index As Integer
End Type

Private TSDF() As TSDFCom

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdPreview_Click()
    PrintMain 0
End Sub

Private Sub cmdPrint_Click()
    PrintMain 1
End Sub

Private Sub PrintMain(printType As Integer)
Dim tmpSql As String
    With Rpt1
        .LogonInfo(0) = "DSN = " & DSN_NAME & ";UID = " & uID & ";PWD = " & uPwd
        .ReportFileName = App.Path & "\reports\ENV05.rpt"
        tmpSql = ""
'        tmpSql = "Select * from Com_vwEUCode_Waste_Qty " & vbCrLf
'        tmpSql = tmpSql & "Where WorkDate >= '" & dtFrom.ValueYMD & "' And WorkDate <= '" & dtTo.ValueYMD & "' " & vbCrLf
'        tmpSql = tmpSql & "And CompanyID='" & CurrentUser.WorkCompanies(cboCompany.ListIndex + 1).ID & "' " & vbCrLf
'        tmpSql = tmpSql & "Order By EUCode, WasteName "
        tmpSql = "{Com_vwEUCode_Waste_Qty.WorkDate} >= '" & dtFrom.ValueYMD & "' AND {Com_vwEUCode_Waste_Qty.WorkDate} <= '" & dtTo.ValueYMD & "' " & vbCrLf
        tmpSql = tmpSql & " AND {Com_vwEUCode_Waste_Qty.TSDFRegisNo1} = '" & GetRegisNo & "' "
        
        .Formulas(0) = "DateRange=" & Chr(34) & "ช่วงวันที่   " & dtFrom.ValueDMY & "      ถึงวันที่   " & dtTo.ValueDMY & Chr(34)
        .Formulas(1) = "Destination=" & Chr(34) & "ปลายทางที่ส่งกำจัด  " & cboCompany.Text & Chr(34)
'        .SQLQuery = tmpSql
        .SelectionFormula = ""
        .SelectionFormula = tmpSql
        .Destination = printType
        .WindowState = crptMaximized
        .WindowShowPrintBtn = True
        .WindowShowPrintSetupBtn = True
        .WindowShowExportBtn = True
        .WindowTitle = "แบบแสดงรหัสของชนิดและประเภทสิ่งปฏิกูลเหรือวัสดุที่ไม่ใช้แล้ว"
        .PrintReport
        If Trim(.LastErrorString) <> "" Then MsgBox .LastErrorString
    End With
End Sub

Private Sub Form_Load()
    PopulateCompany
End Sub

Private Sub PopulateCompany()
Dim x%
Dim sqlStr As String
    cboCompany.Clear
    sqlStr = ""
    Call TSDFCompany(cboCompany, sqlStr)
    If cboCompany.ListCount > 0 Then
        cboCompany.ListIndex = 0
    End If
End Sub

Private Sub TSDFCompany(cbo As ComboBox, Optional strCond As String = "")
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim i As Integer: i = 0
    If Trim(strCond) = "" Then
        sqlStr = "Select CompanyName + ' (' + TSDFRegisNo +')' AS TSDFCom, CompanyName, TSDFRegisNo From tbTSDFCompany Order By CompanyName"
    Else
        sqlStr = "Select CompanyName + ' (' + TSDFRegisNo +')' AS TSDFCom, CompanyName, TSDFRegisNo From tbTSDFCompany Where (" & strCond & ") Order by CompanyName"
    End If
    
    Set tmpRec = New ADODB.Recordset
    tmpRec.Open sqlStr, DBManager.GetDBConnection, adOpenStatic, adLockReadOnly
    If tmpRec.EOF = True Then Exit Sub
    ReDim TSDF(tmpRec.RecordCount)
    Do Until tmpRec.EOF
        cbo.AddItem Trim(tmpRec!TSDFCom), i
        TSDF(i).CompanyName = tmpRec!CompanyName
        TSDF(i).TSDFRegisNo = tmpRec!TSDFRegisNo
        TSDF(i).Index = tmpRec.AbsolutePosition - 1
        i = i + 1
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
End Sub

Private Function GetRegisNo() As String
On Error GoTo Err
    GetRegisNo = TSDF(cboCompany.ListIndex).TSDFRegisNo
    Exit Function
Err:
    GetRegisNo = ""
End Function
