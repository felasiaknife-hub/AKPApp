VERSION 5.00
Begin VB.Form frmCustActivity 
   Caption         =   "บันทึกการติดต่อ"
   ClientHeight    =   4380
   ClientLeft      =   7245
   ClientTop       =   4290
   ClientWidth     =   6630
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4380
   ScaleWidth      =   6630
   Begin VB.CheckBox Check7 
      Caption         =   "อื่นๆ"
      Height          =   195
      Left            =   5490
      TabIndex        =   19
      Top             =   990
      Width           =   690
   End
   Begin VB.CheckBox Check6 
      Caption         =   "วางบิลเก็บเงิน"
      Height          =   195
      Left            =   3915
      TabIndex        =   18
      Top             =   990
      Width           =   1455
   End
   Begin VB.CheckBox Check5 
      Caption         =   "ขนส่ง"
      Height          =   195
      Left            =   3015
      TabIndex        =   17
      Top             =   990
      Width           =   690
   End
   Begin VB.CheckBox Check4 
      Caption         =   "การติดตาม Order"
      Height          =   195
      Left            =   1170
      TabIndex        =   16
      Top             =   990
      Width           =   1635
   End
   Begin VB.CheckBox Check3 
      Caption         =   "สั่งงาน Order"
      Height          =   195
      Left            =   4860
      TabIndex        =   15
      Top             =   675
      Width           =   1365
   End
   Begin VB.CheckBox Check2 
      Caption         =   "ใบเสนอราคา"
      Height          =   195
      Left            =   2790
      TabIndex        =   14
      Top             =   675
      Width           =   1365
   End
   Begin VB.CheckBox Check1 
      Caption         =   "กากฯ ของเสีย"
      Height          =   195
      Left            =   1170
      TabIndex        =   13
      Top             =   675
      Width           =   1365
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ปิด"
      Height          =   600
      Left            =   5400
      TabIndex        =   10
      Top             =   3690
      Width           =   1050
   End
   Begin VB.ComboBox cbResult 
      Height          =   315
      Left            =   1125
      TabIndex        =   8
      Text            =   "Combo1"
      Top             =   3285
      Width           =   5370
   End
   Begin VB.ComboBox cbChannel 
      Height          =   315
      Left            =   1125
      TabIndex        =   6
      Text            =   "Combo1"
      Top             =   1350
      Width           =   5370
   End
   Begin VB.TextBox TxtDet 
      Height          =   1500
      Left            =   1125
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   1710
      Width           =   5370
   End
   Begin VB.CommandButton BtSave 
      Caption         =   "บันทึก"
      Height          =   600
      Left            =   4275
      TabIndex        =   0
      Top             =   3690
      Width           =   1050
   End
   Begin WasteManagment.ctlDate dtdate 
      Height          =   315
      Left            =   1125
      TabIndex        =   2
      Top             =   180
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   556
      BackColor       =   -2147483633
   End
   Begin VB.Label lblID 
      Height          =   240
      Left            =   2025
      TabIndex        =   12
      Top             =   3915
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Label lblCID 
      Height          =   330
      Left            =   4545
      TabIndex        =   11
      Top             =   180
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "ผลลัพท์ :"
      Height          =   195
      Left            =   450
      TabIndex        =   9
      Top             =   3330
      Width           =   600
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "ช่องทาง :"
      Height          =   195
      Left            =   450
      TabIndex        =   7
      Top             =   1395
      Width           =   615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "วันที่ : "
      Height          =   195
      Left            =   630
      TabIndex        =   5
      Top             =   225
      Width           =   450
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "เรื่อง :"
      Height          =   195
      Left            =   630
      TabIndex        =   4
      Top             =   720
      Width           =   420
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "ข้อความ :"
      Height          =   195
      Left            =   405
      TabIndex        =   3
      Top             =   1755
      Width           =   645
   End
End
Attribute VB_Name = "frmCustActivity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub BtSave_Click()
    On Error Resume Next
    Dim RS As ADODB.Recordset
    Dim RSC As ADODB.Recordset
    Dim tSql As String
    Dim tmpList
    Dim i As Integer
    If Trim(TxtDet.Text) = "" Then
        MsgBox "ระบุรายละเอียดก่อน", vbInformation + vbOKOnly, "ผิดพลาด"
        Exit Sub
    End If
    If cbChannel.ListIndex = 0 Then
        MsgBox "ระบุช่องทางก่อน", vbInformation + vbOKOnly, "ผิดพลาด"
        Exit Sub
    End If
    If cbResult.ListIndex = 0 Then
        MsgBox "ระบุผลลัพธ์ก่อน", vbInformation + vbOKOnly, "ผิดพลาด"
        Exit Sub
    End If
    If Trim(lblCID.Caption) = "" Then
        MsgBox "ไม่พบข้อมูลลูกค้าที่ระบุ กรุณาตรวจสอบ", vbInformation + vbOKOnly, "ผิดพลาด"
        Exit Sub
    End If
    Dim SubWaste, SubQT, SubOrder, SubTrackOrder, SubTransport, SubBill, SubOther As String
    If Check1.Value = 1 Then SubWaste = "Y" Else SubWaste = "N"
    If Check2.Value = 1 Then SubQT = "Y" Else SubQT = "N"
    If Check3.Value = 1 Then SubOrder = "Y" Else SubOrder = "N"
    If Check4.Value = 1 Then SubTrackOrder = "Y" Else SubTrackOrder = "N"
    If Check5.Value = 1 Then SubTransport = "Y" Else SubTransport = "N"
    If Check6.Value = 1 Then SubBill = "Y" Else SubBill = "N"
    If Check7.Value = 1 Then SubOther = "Y" Else SubOther = "N"
        If lblID.Caption = "" Then
            tSql = "INSERT INTO tbCustomerActivity(ActivityID ,CustomerID,ActivityDate,ConnectChanel,Description,"
            tSql = tSql + "ActivityResult,SubWaste,SubQT,SubOrder ,SubTrackOrder ,SubTransport,SubBill,SubOther,userid) "
            tSql = tSql + " Values ('" & GetGUID & "', '" & Trim(lblCID.Caption) & "', '" & Trim(dtdate.ValueYMD) & "','" & Trim(cbChannel.Text) & "',"
            tSql = tSql + "'" & TxtDet.Text & "','" & Trim(cbResult.Text) & "','" & SubWaste & "','" & SubQT & "','" & SubOrder & "','" & SubTrackOrder & "',"
            tSql = tSql + "'" & SubTransport & "','" & SubBill & "','" & SubOther & "','" & CurrentUser.ID & "' )"
            DBConnExc tSql
        Else
            Dim Sql As String
            Sql = "update tbCustomerActivity SET ActivityDate =  '" & Trim(dtdate.ValueYMD) & "',ConnectChanel='" & Trim(cbChannel.Text) & "',"
            Sql = Sql + "Description='" & TxtDet.Text & "',ActivityResult =  '" & Trim(cbResult.Text) & "',SubWaste='" & SubWaste & "', " & vbCrLf
            Sql = Sql + "SubQT='" & SubQT & "',SubOrder =  '" & SubOrder & "',SubTrackOrder='" & SubTrackOrder & "', SubTransport='" & SubTransport & "'," & vbCrLf
             Sql = Sql + "SubBill='" & SubBill & "',SubOther =  '" & SubOther & "',userid =  '" & CurrentUser.ID & "' " & vbCrLf
           Sql = Sql & " WHERE ActivityID = " & Trim(lblID.Caption) & ""
            DBConnExc Sql
        End If
        frmCustomer.SetActivity
        Unload Me
End Sub
Sub SetData(ByVal ID As String)
Dim i As Integer
Dim tmpList As ListItem
Dim RS As New ADODB.Recordset
Dim x%, tSql As String
    tSql = "SELECT * FROM tbCustomerActivity WHERE ActivityID=" & ID & " "
    Set RS = New ADODB.Recordset
    Set RS = GetRS(tSql, adOpenStatic, adLockOptimistic)
    If RS.RecordCount > 0 Then
        lblID.Caption = RS!ActivityID
        lblCID.Caption = RS!CustomerID
        dtdate.ValueDMY = Format(RS!ActivityDate, "dd/MM/yyyy")
        If RS!SubWaste = "Y" Then
            Check1.Value = 1
        End If
        If RS!SubQT = "Y" Then
            Check2.Value = 1
        End If
        If RS!SubOrder = "Y" Then
            Check3.Value = 1
        End If
        If RS!SubTrackOrder = "Y" Then
            Check4.Value = 1
        End If
        If RS!SubTransport = "Y" Then
            Check5.Value = 1
        End If
        If RS!SubBill = "Y" Then
            Check6.Value = 1
        End If
        If RS!SubOther = "Y" Then
            Check7.Value = 1
        End If
        cbChannel.Text = RS!ConnectChanel
        TxtDet.Text = RS!Description
        cbResult.Text = RS!ActivityResult
    End If
End Sub
Private Sub Command1_Click()
    Unload Me
End Sub
Private Sub Form_Load()
    With cbChannel
        .AddItem "===เลือก==="
        .AddItem "โทรศัพท์"
        .AddItem "Email"
        .AddItem "เข้าพบ"
        .AddItem "Application Line"
        .AddItem "อื่นๆ"
        .ListIndex = 0
    End With
    With cbResult
        .AddItem "===เลือก==="
        .AddItem "ดำเนินการเรื่องกากของเสียได้"
        .AddItem "ลูกค้าอนุมัติใบเสนอราคา"
        .AddItem "ลูกค้าพิจารณาใบเสนอราคาอยู่"
        .AddItem "ลูกค้าสั่ง Order"
        .AddItem "ลูกค้าไม่สั่ง Order"
        .AddItem "อยู่ระหว่างดำเนินการ"
        .AddItem "ดำเนินการต่อ"
        .ListIndex = 0
    End With
End Sub
