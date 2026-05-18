VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmBurnCost 
   Caption         =   "ต้นทุนการเผา"
   ClientHeight    =   2775
   ClientLeft      =   5505
   ClientTop       =   4290
   ClientWidth     =   6855
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2775
   ScaleWidth      =   6855
   Begin VB.CommandButton Command2 
      Caption         =   "ปิด"
      Height          =   420
      Left            =   5400
      TabIndex        =   31
      Top             =   2250
      Width           =   1320
   End
   Begin VB.CommandButton btSave 
      Caption         =   "บันทึก"
      Height          =   420
      Left            =   4050
      TabIndex        =   30
      Top             =   2250
      Width           =   1320
   End
   Begin VB.TextBox txtChem 
      Height          =   285
      Left            =   4995
      TabIndex        =   28
      Text            =   "0"
      Top             =   1305
      Width           =   1365
   End
   Begin VB.TextBox txtElec 
      Height          =   285
      Left            =   4995
      TabIndex        =   25
      Text            =   "0"
      Top             =   945
      Width           =   1365
   End
   Begin VB.TextBox txtGas 
      Height          =   285
      Left            =   4995
      TabIndex        =   22
      Text            =   "0"
      Top             =   585
      Width           =   1365
   End
   Begin MSComCtl2.DTPicker DTPDate 
      Height          =   330
      Left            =   1125
      TabIndex        =   3
      Top             =   45
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   582
      _Version        =   393216
      Format          =   84017153
      CurrentDate     =   44537
   End
   Begin VB.Frame Frame1 
      Caption         =   "ยอดเผา"
      Height          =   2220
      Left            =   45
      TabIndex        =   0
      Top             =   450
      Width           =   3615
      Begin VB.TextBox txtSludge 
         Height          =   285
         Left            =   1485
         TabIndex        =   19
         Text            =   "0"
         Top             =   1800
         Width           =   1365
      End
      Begin VB.TextBox txtEvap 
         Height          =   285
         Left            =   1485
         TabIndex        =   16
         Text            =   "0"
         Top             =   1485
         Width           =   1365
      End
      Begin VB.TextBox txtSW 
         Height          =   285
         Left            =   1485
         TabIndex        =   13
         Text            =   "0"
         Top             =   1170
         Width           =   1365
      End
      Begin VB.TextBox txtSolid 
         Height          =   285
         Left            =   1485
         TabIndex        =   10
         Text            =   "0"
         Top             =   855
         Width           =   1365
      End
      Begin VB.TextBox txtBoxFeed 
         Height          =   285
         Left            =   1485
         TabIndex        =   7
         Text            =   "0"
         Top             =   540
         Width           =   1365
      End
      Begin VB.TextBox txtWasteWater 
         Height          =   285
         Left            =   1485
         TabIndex        =   4
         Text            =   "0"
         Top             =   225
         Width           =   1365
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "ตัน"
         Height          =   195
         Left            =   2925
         TabIndex        =   20
         Top             =   1845
         Width           =   225
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Sludge(Pump) :"
         Height          =   195
         Left            =   315
         TabIndex        =   18
         Top             =   1800
         Width           =   1080
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "ลบ.ม"
         Height          =   195
         Left            =   2925
         TabIndex        =   17
         Top             =   1530
         Width           =   345
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Evap(Customer) :"
         Height          =   195
         Left            =   180
         TabIndex        =   15
         Top             =   1485
         Width           =   1215
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "ตัน"
         Height          =   195
         Left            =   2925
         TabIndex        =   14
         Top             =   1215
         Width           =   225
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Special Waste :"
         Height          =   195
         Left            =   270
         TabIndex        =   12
         Top             =   1170
         Width           =   1125
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "ตัน"
         Height          =   195
         Left            =   2925
         TabIndex        =   11
         Top             =   900
         Width           =   225
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Solid :"
         Height          =   195
         Left            =   945
         TabIndex        =   9
         Top             =   855
         Width           =   435
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "ตัน"
         Height          =   195
         Left            =   2925
         TabIndex        =   8
         Top             =   585
         Width           =   225
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Box feed :"
         Height          =   195
         Left            =   675
         TabIndex        =   6
         Top             =   540
         Width           =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "ลบ.ม"
         Height          =   195
         Left            =   2925
         TabIndex        =   5
         Top             =   270
         Width           =   345
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "ยอดน้ำเสีย :"
         Height          =   195
         Left            =   540
         TabIndex        =   1
         Top             =   225
         Width           =   855
      End
   End
   Begin VB.Label lblID 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   225
      TabIndex        =   32
      Top             =   2970
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      Caption         =   "บาท"
      Height          =   195
      Left            =   6435
      TabIndex        =   29
      Top             =   1350
      Width           =   285
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   "ค่าสารเคมี :"
      Height          =   195
      Left            =   4140
      TabIndex        =   27
      Top             =   1350
      Width           =   810
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "บาท"
      Height          =   195
      Left            =   6435
      TabIndex        =   26
      Top             =   990
      Width           =   285
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "ค่าไฟฟ้า :"
      Height          =   195
      Left            =   4275
      TabIndex        =   24
      Top             =   990
      Width           =   675
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "บาท"
      Height          =   195
      Left            =   6435
      TabIndex        =   23
      Top             =   630
      Width           =   285
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "ค่าแก้ส :"
      Height          =   195
      Left            =   4365
      TabIndex        =   21
      Top             =   630
      Width           =   585
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "วันที่"
      Height          =   195
      Left            =   720
      TabIndex        =   2
      Top             =   90
      Width           =   315
   End
End
Attribute VB_Name = "FrmBurnCost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub btSave_Click()
    Dim Sql, tSql As String
    Dim rs As New ADODB.Recordset
    Dim BID, BDate As String
    If Trim(txtWasteWater.Text) = "" Then txtWasteWater.Text = "0"
    If Trim(txtBoxFeed.Text) = "" Then txtBoxFeed.Text = "0"
    If Trim(txtSolid.Text) = "" Then txtSolid.Text = "0"
    If Trim(txtSW.Text) = "" Then txtSW.Text = "0"
    If Trim(txtEvap.Text) = "" Then txtEvap.Text = "0"
    If Trim(txtSludge.Text) = "" Then txtSludge.Text = "0"
    If Trim(txtGas.Text) = "" Then txtGas.Text = "0"
    If Trim(txtElec.Text) = "" Then txtElec.Text = "0"
    If Trim(txtChem.Text) = "" Then txtChem.Text = "0"
    
    Sql = "SELECT * FROM tbBurnCost " & vbCrLf
    Sql = Sql & "where id='" & lblID.Caption & "'" & vbCrLf
    Set rs = GetRS(Sql, adOpenForwardOnly, adLockOptimistic)
        If rs.RecordCount > 0 Then
            DBConnExc "Update tbBurnCost Set WasteWater=" & CDbl(txtWasteWater.Text) & ",BoxFeed=" & CDbl(txtBoxFeed.Text) & ",Solid=" & CDbl(txtSolid.Text) & " " & _
            ",SpecialWaste=" & CDbl(txtSW.Text) & ",Evap=" & CDbl(txtEvap.Text) & ",Sludge=" & CDbl(txtSludge.Text) & ",Gas=" & CDbl(txtGas.Text) & " " & _
            ",Elec=" & CDbl(txtElec.Text) & ",Chem=" & CDbl(txtChem.Text) & " where ID='" & lblID.Caption & "' "
        Else
        BID = GetGUID
        BDate = Format(DTPDate.Value, "yyyy/MM/dd")
            tSql = "INSERT INTO tbBurnCost(ID,burnDate,WasteWater,BoxFeed,Solid,SpecialWaste,Evap,Sludge,Gas,Elec,Chem) "
            tSql = tSql + " Values ('" & BID & "','" & BDate & "', " & CDbl(txtWasteWater.Text) & ", " & CDbl(txtBoxFeed.Text) & "," & CDbl(txtSolid.Text) & " " & _
            "," & CDbl(txtSW.Text) & "," & CDbl(txtEvap.Text) & "," & CDbl(txtSludge.Text) & "," & CDbl(txtGas.Text) & "," & CDbl(txtElec.Text) & "," & CDbl(txtChem.Text) & ")"
            DBConnExc tSql
        End If
        MsgBox "บันทึกข้อมูลเรียบร้อย", vbInformation + vbOKOnly, "บันทึกข้อมูล"
End Sub
Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub DTPDate_Change()
    Dim Sql, tSql As String
    Dim rs As New ADODB.Recordset
    Dim i As Integer
    Sql = "SELECT * FROM tbBurnCost " & vbCrLf
    Sql = Sql & "where burnDate='" & Format(DTPDate.Value, "yyyy/MM/dd") & "'" & vbCrLf
    Set rs = GetRS(Sql, adOpenForwardOnly, adLockOptimistic)
    If rs.RecordCount > 0 Then
        lblID.Caption = rs!ID
        txtWasteWater.Text = Format(rs!WasteWater, "#,##0.00")
        txtBoxFeed.Text = Format(rs!BoxFeed, "#,##0.00")
        txtSolid.Text = Format(rs!Solid, "#,##0.00")
        txtSW.Text = Format(rs!SpecialWaste, "#,##0.00")
        txtEvap.Text = Format(rs!Evap, "#,##0.00")
        txtSludge.Text = Format(rs!Sludge, "#,##0.00")
        txtGas.Text = Format(rs!Gas, "#,##0.00")
        txtElec.Text = Format(rs!Elec, "#,##0.00")
        txtChem.Text = Format(rs!Chem, "#,##0.00")
    Else
        lblID.Caption = ""
        txtWasteWater.Text = "0"
        txtBoxFeed.Text = "0"
        txtSolid.Text = "0"
        txtSW.Text = "0"
        txtEvap.Text = "0"
        txtSludge.Text = "0"
        txtGas.Text = "0"
        txtElec.Text = "0"
        txtChem.Text = "0"
    End If
        With txtWasteWater
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
        End With
End Sub
Private Sub DTPDate_Validate(Cancel As Boolean)
    Dim Sql, tSql As String
    Dim rs As New ADODB.Recordset
    Dim i As Integer
    Sql = "SELECT * FROM tbBurnCost " & vbCrLf
    Sql = Sql & "where burnDate='" & Format(DTPDate.Value, "yyyy/MM/dd") & "'" & vbCrLf
    Set rs = GetRS(Sql, adOpenForwardOnly, adLockOptimistic)
    If rs.RecordCount > 0 Then
        lblID.Caption = rs!ID
        txtWasteWater.Text = rs!WasteWater
        txtBoxFeed.Text = rs!BoxFeed
        txtSolid.Text = rs!Solid
        txtSW.Text = rs!SpecialWaste
        txtEvap.Text = rs!Evap
        txtSludge.Text = rs!Sludge
        txtGas.Text = rs!Gas
        txtElec.Text = rs!Elec
        txtChem.Text = rs!Chem
    Else
        lblID.Caption = ""
        txtWasteWater.Text = "0"
        txtBoxFeed.Text = "0"
        txtSolid.Text = "0"
        txtSW.Text = "0"
        txtEvap.Text = "0"
        txtSludge.Text = "0"
        txtGas.Text = "0"
        txtElec.Text = "0"
        txtChem.Text = "0"
    End If
End Sub

Private Sub Form_Load()
    DTPDate.Value = Date
End Sub

Private Sub txtBoxFeed_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        With txtSolid
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
        End With
            txtBoxFeed.Text = Format(txtBoxFeed.Text, "#,##0.00")
    End If
End Sub

Private Sub txtBoxFeed_Validate(Cancel As Boolean)
    If IsNumeric(txtBoxFeed.Text) = False Then
        MsgBox "ให้ใส่ตัวเลขเท่านั้น", vbInformation + vbOKOnly, "ผิดพลาด"
        txtBoxFeed.Text = "0"
    End If
            txtBoxFeed.Text = Format(txtBoxFeed.Text, "#,##0.00")
End Sub

Private Sub txtChem_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        With btSave
            .SetFocus
        End With
            txtChem.Text = Format(txtChem.Text, "#,##0.00")
    End If
End Sub
Private Sub txtChem_Validate(Cancel As Boolean)
    If IsNumeric(txtChem.Text) = False Then
        MsgBox "ให้ใส่ตัวเลขเท่านั้น", vbInformation + vbOKOnly, "ผิดพลาด"
        txtChem.Text = "0"
    End If
            txtChem.Text = Format(txtChem.Text, "#,##0.00")
End Sub
Private Sub txtElec_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        With txtChem
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
        End With
            txtElec.Text = Format(txtElec.Text, "#,##0.00")
    End If
End Sub
Private Sub txtElec_Validate(Cancel As Boolean)
    If IsNumeric(txtElec.Text) = False Then
        MsgBox "ให้ใส่ตัวเลขเท่านั้น", vbInformation + vbOKOnly, "ผิดพลาด"
        txtElec.Text = "0"
    End If
            txtElec.Text = Format(txtElec.Text, "#,##0.00")
End Sub
Private Sub txtEvap_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        With txtSludge
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
        End With
            txtEvap.Text = Format(txtEvap.Text, "#,##0.00")
    End If
End Sub

Private Sub txtEvap_Validate(Cancel As Boolean)
    If IsNumeric(txtEvap.Text) = False Then
        MsgBox "ให้ใส่ตัวเลขเท่านั้น", vbInformation + vbOKOnly, "ผิดพลาด"
        txtEvap.Text = "0"
    End If
            txtEvap.Text = Format(txtEvap.Text, "#,##0.00")
End Sub

Private Sub txtGas_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        With txtElec
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
        End With
            txtGas.Text = Format(txtGas.Text, "#,##0.00")
    End If
End Sub

Private Sub txtGas_Validate(Cancel As Boolean)
    If IsNumeric(txtGas.Text) = False Then
        MsgBox "ให้ใส่ตัวเลขเท่านั้น", vbInformation + vbOKOnly, "ผิดพลาด"
        txtGas.Text = "0"
    End If
            txtGas.Text = Format(txtGas.Text, "#,##0.00")
End Sub

Private Sub txtSludge_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        With txtGas
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
        End With
            txtSludge.Text = Format(txtSludge.Text, "#,##0.00")
    End If
End Sub

Private Sub txtSludge_Validate(Cancel As Boolean)
    If IsNumeric(txtSludge.Text) = False Then
        MsgBox "ให้ใส่ตัวเลขเท่านั้น", vbInformation + vbOKOnly, "ผิดพลาด"
        txtSludge.Text = "0"
    End If
            txtSludge.Text = Format(txtSludge.Text, "#,##0.00")
End Sub

Private Sub txtSolid_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        With txtSW
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
        End With
            txtSolid.Text = Format(txtSolid.Text, "#,##0.00")
    End If
End Sub

Private Sub txtSolid_Validate(Cancel As Boolean)
    If IsNumeric(txtSolid.Text) = False Then
        MsgBox "ให้ใส่ตัวเลขเท่านั้น", vbInformation + vbOKOnly, "ผิดพลาด"
        txtSolid.Text = "0"
    End If
            txtSolid.Text = Format(txtSolid.Text, "#,##0.00")
End Sub

Private Sub txtSW_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        With txtEvap
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
        End With
            txtSW.Text = Format(txtSW.Text, "#,##0.00")
    End If
End Sub

Private Sub txtSW_Validate(Cancel As Boolean)
    If IsNumeric(txtSW.Text) = False Then
        MsgBox "ให้ใส่ตัวเลขเท่านั้น", vbInformation + vbOKOnly, "ผิดพลาด"
        txtSW.Text = "0"
    End If
            txtSW.Text = Format(txtSW.Text, "#,##0.00")
End Sub

Private Sub txtWasteWater_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        With txtBoxFeed
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
        End With
            txtWasteWater.Text = Format(txtWasteWater.Text, "#,##0.00")
    End If
End Sub

Private Sub txtWasteWater_Validate(Cancel As Boolean)
    If IsNumeric(txtWasteWater.Text) = False Then
        MsgBox "ให้ใส่ตัวเลขเท่านั้น", vbInformation + vbOKOnly, "ผิดพลาด"
        txtWasteWater.Text = "0"
    End If
            txtWasteWater.Text = Format(txtWasteWater.Text, "#,##0.00")
End Sub
