VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPaytoDeptor 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "F064"
   ClientHeight    =   9165
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11775
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9165
   ScaleWidth      =   11775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAddTimeTable 
      Caption         =   "เพิ่ม/แก้ไขใบคุม"
      Height          =   795
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   8340
      Width           =   1725
   End
   Begin VB.ComboBox cboVAT 
      Height          =   315
      Left            =   9900
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   1680
      Width           =   1605
   End
   Begin VB.ComboBox cboWTax 
      Height          =   315
      ItemData        =   "frPaytoDeptor.frx":0000
      Left            =   6900
      List            =   "frPaytoDeptor.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   1680
      Width           =   765
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "รายการเงินค้างที่จะนำมาหักในงวดนี้"
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   2
      Left            =   10260
      TabIndex        =   14
      Top             =   6840
      Visible         =   0   'False
      Width           =   975
      Begin VB.CommandButton cmdDeSelectAll 
         Caption         =   "-"
         Height          =   225
         Left            =   3030
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   0
         Width           =   375
      End
      Begin VB.CommandButton cmdSelectAll 
         Caption         =   "+"
         Height          =   225
         Left            =   2670
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   0
         Width           =   375
      End
      Begin MSComctlLib.ListView lvOther 
         Height          =   2025
         Left            =   60
         TabIndex        =   15
         Top             =   240
         Width           =   5715
         _ExtentX        =   10081
         _ExtentY        =   3572
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "วันที่"
            Object.Width           =   2249
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "รายการ"
            Object.Width           =   5371
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "ยอดเงิน"
            Object.Width           =   1826
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "รายการจ่ายเพิ่มเติมอื่นๆ"
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   1
      Left            =   10260
      TabIndex        =   46
      Top             =   6480
      Visible         =   0   'False
      Width           =   975
      Begin WasteManagment.ctlGrid grItem 
         Height          =   2055
         Left            =   60
         TabIndex        =   16
         Top             =   210
         Width           =   5715
         _ExtentX        =   10081
         _ExtentY        =   3625
      End
   End
   Begin VB.TextBox txtVendor 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1170
      TabIndex        =   3
      Top             =   1290
      Width           =   4275
   End
   Begin WasteManagment.ctlDate dtIssueDate 
      Height          =   345
      Left            =   3660
      TabIndex        =   1
      Top             =   870
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   609
      BackColor       =   -2147483633
   End
   Begin VB.TextBox txtRem 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1170
      MaxLength       =   80
      TabIndex        =   10
      Top             =   2070
      Width           =   10335
   End
   Begin VB.TextBox txtRecNo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   4020
      TabIndex        =   7
      Top             =   1680
      Width           =   1425
   End
   Begin VB.TextBox txtInvNo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1170
      TabIndex        =   6
      Top             =   1680
      Width           =   1725
   End
   Begin VB.TextBox txtBillNo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   9900
      TabIndex        =   5
      Top             =   1290
      Width           =   1635
   End
   Begin Crystal.CrystalReport rpt1 
      Left            =   3360
      Top             =   30
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin VB.ComboBox cboPrinter 
      Height          =   315
      Left            =   8190
      Style           =   2  'Dropdown List
      TabIndex        =   37
      Top             =   30
      Visible         =   0   'False
      Width           =   3525
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "เก็บข้อมูล"
      Height          =   795
      Left            =   5340
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   8340
      Width           =   1005
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "เปิดข้อมูลเก่า"
      Height          =   795
      Left            =   6390
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   8340
      Width           =   1065
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "สร้างใหม่"
      Height          =   795
      Left            =   7500
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   8340
      Width           =   1005
   End
   Begin VB.CommandButton cmdPayment 
      Caption         =   "บันทึกการจ่ายเงิน"
      Height          =   795
      Left            =   1830
      Picture         =   "frPaytoDeptor.frx":0004
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   8340
      Visible         =   0   'False
      Width           =   1725
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "ปิด"
      Height          =   795
      Left            =   10680
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   8340
      Width           =   1035
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "พิมพ์"
      Height          =   795
      Left            =   8550
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   8340
      Width           =   1005
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ยกเลิก"
      Height          =   795
      Left            =   9600
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   8340
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "รายการใบคุมที่ทำจ่ายงวดนี้"
      ForeColor       =   &H80000008&
      Height          =   4905
      Index           =   0
      Left            =   30
      TabIndex        =   29
      Top             =   2400
      Width           =   11715
      Begin WasteManagment.ctlGrid grBill 
         Height          =   4695
         Left            =   60
         TabIndex        =   11
         Top             =   180
         Width           =   11595
         _ExtentX        =   20452
         _ExtentY        =   8281
      End
   End
   Begin VB.ComboBox cboCompany 
      Height          =   315
      Left            =   6900
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   870
      Width           =   4635
   End
   Begin VB.TextBox txtSaleNo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1170
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Text            =   "== AUTO =="
      Top             =   900
      Width           =   1875
   End
   Begin WasteManagment.ctlDate dtDueDate 
      Height          =   345
      Left            =   6900
      TabIndex        =   4
      Top             =   1290
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   609
      BackColor       =   -2147483633
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   5220
      TabIndex        =   61
      Top             =   7500
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "อื่นๆ ก่อน VAT"
      Height          =   195
      Index           =   5
      Left            =   4020
      TabIndex        =   60
      Top             =   7500
      Width           =   1095
   End
   Begin VB.Label lbTotalMenifest 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   1890
      TabIndex        =   59
      Top             =   7500
      Width           =   1005
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "รวมยอด Menifest"
      Height          =   195
      Index           =   23
      Left            =   180
      TabIndex        =   58
      Top             =   7500
      Width           =   1275
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "หักค่าเที่ยว"
      Height          =   195
      Index           =   26
      Left            =   1020
      TabIndex        =   57
      Top             =   7860
      Width           =   765
   End
   Begin VB.Label lbFund 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   1860
      TabIndex        =   56
      Top             =   7860
      Width           =   1035
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "หักเงินสะสม"
      Height          =   195
      Index           =   24
      Left            =   4200
      TabIndex        =   55
      Top             =   7860
      Width           =   840
   End
   Begin VB.Label lbDeduct 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   5190
      TabIndex        =   54
      Top             =   7860
      Width           =   1005
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VAT"
      Height          =   195
      Index           =   22
      Left            =   9480
      TabIndex        =   53
      Top             =   1740
      Width           =   345
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ยอดที่ต้องสั่งจ่าย"
      Height          =   195
      Index           =   14
      Left            =   9060
      TabIndex        =   52
      Top             =   7860
      Width           =   1155
   End
   Begin VB.Label lbNetAmt 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   10410
      TabIndex        =   51
      Top             =   7860
      Width           =   1245
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TAX"
      Height          =   195
      Index           =   12
      Left            =   7020
      TabIndex        =   50
      Top             =   7860
      Width           =   315
   End
   Begin VB.Label lbWTaxAmt 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   7860
      TabIndex        =   49
      Top             =   7860
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      Height          =   195
      Index           =   10
      Left            =   7770
      TabIndex        =   48
      Top             =   1740
      Width           =   120
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ภาษีหัก ณ. ที่จ่าย"
      Height          =   195
      Index           =   9
      Left            =   5670
      TabIndex        =   47
      Top             =   1740
      Width           =   1155
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "เจ้าหนี้"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   8
      Left            =   630
      TabIndex        =   45
      Top             =   1350
      Width           =   465
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "หมายเหตุ"
      Height          =   195
      Index           =   7
      Left            =   420
      TabIndex        =   44
      Top             =   2130
      Width           =   675
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "กำหนดชำระ"
      Height          =   195
      Index           =   18
      Left            =   5970
      TabIndex        =   43
      Top             =   1350
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "เลขที่ใบเสร็จ"
      Height          =   195
      Index           =   17
      Left            =   3060
      TabIndex        =   42
      Top             =   1740
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ใบกำกับภาษี"
      Height          =   195
      Index           =   16
      Left            =   240
      TabIndex        =   41
      Top             =   1740
      Width           =   885
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "เลขใบวางบิล"
      Height          =   195
      Index           =   15
      Left            =   8970
      TabIndex        =   40
      Top             =   1350
      Width           =   885
   End
   Begin VB.Label lbCancel 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "ยกเลิก"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   645
      Left            =   4950
      TabIndex        =   39
      Top             =   0
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "เลือกเครื่องพิมพ์"
      Height          =   195
      Index           =   8
      Left            =   6960
      TabIndex        =   38
      Top             =   90
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   11880
      Y1              =   8280
      Y2              =   8280
   End
   Begin VB.Label lbTotalPrice 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   2340
      TabIndex        =   36
      Top             =   8580
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label lbVat 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   10320
      TabIndex        =   35
      Top             =   7500
      Width           =   1305
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "บาท"
      Height          =   195
      Index           =   20
      Left            =   3060
      TabIndex        =   34
      Top             =   9090
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ค่าบริการ"
      Height          =   195
      Index           =   19
      Left            =   360
      TabIndex        =   33
      Top             =   9090
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VAT"
      Height          =   195
      Index           =   6
      Left            =   9720
      TabIndex        =   32
      Top             =   7500
      Width           =   375
   End
   Begin VB.Label lbDAVat 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   7710
      TabIndex        =   31
      Top             =   7500
      Width           =   1245
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "อื่นๆ หลัง VAT"
      Height          =   195
      Index           =   4
      Left            =   6480
      TabIndex        =   30
      Top             =   7500
      Width           =   1035
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ทำรายการชำระเงินเจ้าหนี้"
      BeginProperty Font 
         Name            =   "Angsana New"
         Size            =   18
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   180
      TabIndex        =   28
      Top             =   90
      Width           =   2460
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      Index           =   1
      X1              =   240
      X2              =   7050
      Y1              =   630
      Y2              =   630
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   4
      Index           =   0
      X1              =   150
      X2              =   6960
      Y1              =   540
      Y2              =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "จ่ายในนาม"
      Height          =   195
      Index           =   1
      Left            =   6060
      TabIndex        =   27
      Top             =   930
      Width           =   780
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "วันที่"
      Height          =   195
      Index           =   2
      Left            =   3240
      TabIndex        =   26
      Top             =   960
      Width           =   315
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "เลขที่ใบจ่าย"
      Height          =   195
      Index           =   0
      Left            =   270
      TabIndex        =   25
      Top             =   960
      Width           =   825
   End
   Begin VB.Image Image2 
      Height          =   795
      Left            =   0
      Picture         =   "frPaytoDeptor.frx":046F
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12030
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFFF80&
      Height          =   315
      Index           =   3
      Left            =   7590
      Top             =   7440
      Width           =   1425
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFFF80&
      Height          =   315
      Index           =   2
      Left            =   10260
      Top             =   7440
      Width           =   1455
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFFF80&
      Height          =   315
      Index           =   1
      Left            =   1110
      Top             =   9030
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFFF80&
      Height          =   315
      Index           =   0
      Left            =   1800
      Top             =   7800
      Width           =   1155
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFFF80&
      Height          =   315
      Index           =   4
      Left            =   10260
      Top             =   7800
      Width           =   1455
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFFF80&
      Height          =   315
      Index           =   5
      Left            =   5160
      Top             =   7800
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFFF80&
      Height          =   315
      Index           =   6
      Left            =   7560
      Top             =   7800
      Width           =   1455
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFFF80&
      Height          =   315
      Index           =   7
      Left            =   1770
      Top             =   7440
      Width           =   1185
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFFF80&
      Height          =   315
      Index           =   8
      Left            =   5160
      Top             =   7440
      Width           =   1095
   End
End
Attribute VB_Name = "frmPaytoDeptor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim curVendor_ As BWGTransportVendor
Dim myBillPay_ As BWGBillPayHeader
Dim colMenifest_ As Collection
Dim colExpense_ As Collection
Dim isDetailChange_ As Boolean

Private Sub cboCompany_Validate(Cancel As Boolean)
    If cboCompany.ListIndex >= 0 Then
        If cboCompany.ListCount > 1 Then
            If cboCompany.ListIndex = 0 Then Exit Sub
            myBillPay_.CompanyID = CurrentUser.WorkCompanies(cboCompany.ListIndex).ID
        Else
            myBillPay_.CompanyID = CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).ID
        End If
        myBillPay_.CompanyName = cboCompany.Text
    Else
        myBillPay_.CompanyID = ""
        myBillPay_.CompanyName = ""
    End If
End Sub

Private Sub cboVAT_Click()
    If Me.Visible Then
        If Me.ActiveControl.Name = "cboVAT" Then
            If cboVAT.ListIndex = 0 Then
                myBillPay_.VATRate = 0
            Else
                myBillPay_.VATRate = SystemConfig.VATRate
            End If
            SumMoneyAmt
        End If
    End If
End Sub

Private Sub cboWTax_Click()
    If Me.Visible Then
        If Me.ActiveControl.Name = "cboWTax" Then
            myBillPay_.WTax = cboWTax.Text
            SumMoneyAmt
        End If
    End If
End Sub

Private Sub cboWTax_Validate(Cancel As Boolean)
    If IsNumeric(cboWTax.Text) Then
        myBillPay_.WTax = cboWTax.Text
    Else
        myBillPay_.WTax = 0
    End If
End Sub

Private Sub cmdCancel_Click()
    If MsgBox("ท่านต้องการยกเลิกการทำใบจ่ายเงินนี้หรือไม่?", vbQuestion + vbYesNo) = vbYes Then
        myBillPay_.CancelBill
        SetCancelScreen
    End If
End Sub

Private Sub cmdClear_Click()
Dim selResult As Integer
    If myBillPay_.isChanged And myBillPay_.isCancel = False Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนสร้างใหม่หรือไม่?", vbQuestion + vbYesNoCancel)
        If selResult = vbYes Then
            cmdSave_Click
            If myBillPay_.isChanged Then Exit Sub
        ElseIf selResult = vbCancel Then
            Exit Sub
        End If
    End If
    ClearScreen
    CreateNewBillPay
    txtVendor.SetFocus
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Public Sub SetDataTransportVender(selVendor As BWGTransportVendor, Optional isAssigned As Boolean = True)
    Set curVendor_ = selVendor
    txtVendor.Text = selVendor.NameTH
    SetTextToCombo selVendor.WTaxRate, cboWTax
    If isAssigned Then
        myBillPay_.SupplierID = curVendor_.ID
        myBillPay_.SupplierName = curVendor_.NameTH
    End If
    ShowOtherExp
End Sub

Private Sub ShowPayDetail()
Dim X%, i%, tmpDet As BWGBillPayDetail
Dim tmpCol As Collection
    Set tmpCol = myBillPay_.BillPayDetails
    grBill.ClearAllData
    With grBill.getGridObj
        .Rows = 20 + tmpCol.Count
        grItem.getGridObj.Rows = 1
        Dim XX%
        For X = 1 To tmpCol.Count
            Set tmpDet = tmpCol(X)
'            .TextMatrix(x, 0) = tmpDet.ID
'            .TextMatrix(x, 1) = FormatYMD_to_DMY(tmpDet.WorkDate)
'            .TextMatrix(x, 2) = tmpDet.TimeTableNo
'            .TextMatrix(x, 3) = tmpDet.FromAumphurName & ", " & tmpDet.FromProvinceName
'            .TextMatrix(x, 4) = tmpDet.ToAumphurName & ", " & tmpDet.ToProvinceName
'            .TextMatrix(x, 5) = Format(tmpDet.AmtB4VAT, "#,##0.00")
            Dim isOther1 As Boolean, isOther2 As Boolean
            For i = 1 To lvOther.ListItems.Count
                If tmpDet.ExpenseID = Replace(lvOther.ListItems(i).Key, "'", "") Then
                    lvOther.ListItems(i).Checked = True
                    isOther1 = True
                    Exit For
                End If
            Next i
            If Trim("" & tmpDet.BeforeOrAfterVAT) = "" Then
                isOther2 = False
            Else
                isOther2 = True
            End If
            If isOther1 = False Then
                If isOther2 = False Then
                    XX = XX + 1
                    .TextMatrix(XX, 0) = tmpDet.ID
                    .TextMatrix(XX, 1) = tmpDet.TimeTableNo
                    .TextMatrix(XX, 2) = tmpDet.BillDesc
                    .TextMatrix(XX, 3) = Format(tmpDet.AmtB4VAT, "#,##0.00")
                    Set tmpDet = Nothing
                Else
                    With grItem.getGridObj
                        .AddItem tmpDet.ID & vbTab & tmpDet.BillDesc & vbTab & Format(tmpDet.AmtB4VAT, "#,000.00") & vbTab & IIf(tmpDet.BeforeOrAfterVAT = "B", "ก่อน VAT", "หลัง VAT")
                    End With
                End If
            End If
            isOther1 = False
            isOther2 = False
        Next
    End With
    grItem.getGridObj.Rows = grItem.getGridObj.Rows + 10
    SumMoneyAmt
End Sub

Private Sub ShowOtherExp()
Dim X%, tmpItem As BWGVendorOtherExpense, tmpList As ListItem
Dim tmpCol As Collection, tmpS As New BWGSearchManager
    Set tmpCol = tmpS.VendorExpenseSearch("VendorID='" & myBillPay_.SupplierID & "' And (BillPayHeaderID is null Or BillPayHeaderID='" & myBillPay_.ID & "')")
    lvOther.ListItems.Clear
    For X = 1 To tmpCol.Count
        Set tmpItem = tmpCol(X)
        Set tmpList = lvOther.ListItems.Add(, "'" & tmpItem.ID & "'", FormatYMD_to_DMY(tmpItem.ExpenseDate, "/", "/"))
        tmpList.SubItems(1) = Trim(tmpItem.BillDesc)
        tmpList.SubItems(2) = Format(tmpItem.NetPrice, "#,##0.00")
        Set tmpItem = Nothing
        Set tmpList = Nothing
    Next
    Set colExpense_ = tmpCol
    Set tmpCol = Nothing
    Set tmpS = Nothing
End Sub

Private Sub cmdOpen_Click()
Dim selResult As Integer
    If myBillPay_.isChanged And myBillPay_.isCancel = False Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนทำงานต่อหรือไม่?", vbQuestion + vbYesNoCancel)
        If selResult = vbYes Then
            cmdSave_Click
            If myBillPay_.isChanged Then Exit Sub
        ElseIf selResult = vbCancel Then
            Exit Sub
        End If
    End If
    Load frmPaytoDeptorSearch
    With frmPaytoDeptorSearch
        .SetCallerForm Me
        .Show 1
    End With
End Sub

Private Sub cmdPayment_Click()
Dim selResult As Integer
    If Trim(txtVendor.Text) = "" Then
        MsgBox "กรุณาระบุชื่อเจ้าหนี้", vbExclamation
        txtVendor.SetFocus
        Exit Sub
    End If
    If myBillPay_.isChanged And myBillPay_.isCancel = False Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนทำงานต่อหรือไม่?", vbQuestion + vbYesNo)
        If selResult = vbYes Then
            cmdSave_Click
            If myBillPay_.isChanged Then Exit Sub
        Else 'If selResult = vbCancel Then
            Exit Sub
        End If
    End If
    Load frmCheqPayment
    With frmCheqPayment
        .SetCallerForm Me
        .TotalPayments lbNetAmt
        .Show 1
    End With
End Sub

Private Sub cmdPrint_Click()
Dim tmpSql As String
Dim selResult
Dim tmpStr As String
Dim AcceptState As Boolean, i As Integer
    tmpStr = CurrentUser.VerifyAccessRole(Me.Name, PrintRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If

    If Trim(txtVendor.Text) = "" Then
        MsgBox "กรุณาใส่เจ้าหนี้", vbExclamation
        Exit Sub
    End If
    If Trim(cboCompany.Text) = "" Then
        MsgBox "กรุณาระบุ บริษัท", vbExclamation
        cboCompany.SetFocus
        Exit Sub
    End If

    If myBillPay_.isChanged = True And cmdSave.Enabled Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนสร้างใหม่หรือไม่?", vbQuestion + vbYesNo)
        If selResult = vbYes Then
            cmdSave_Click
            If myBillPay_.isChanged Then Exit Sub
        Else 'If selResult = vbCancel Then
            Exit Sub
        End If
    End If

    With rpt1
        .WindowShowCloseBtn = True
        .WindowShowExportBtn = True
        .WindowShowPrintSetupBtn = True
        .LogonInfo(0) = "DSN = " & DSN_NAME & ";UID = " & strUid & ";PWD = " & strPwd
        .ReportFileName = App.Path & "\reports\Billpayvender.rpt"
        .Formulas(0) = "Company=" & Chr(34) & txtVendor.Text & Chr(34)
        .Formulas(1) = "CompanyAddr=" & Chr(34) & curVendor_.Addr & Chr(34)
        .Formulas(2) = "CompanyAddr1=" & Chr(34) & curVendor_.TumbolName & " " & curVendor_.AumphurName & " " & curVendor_.ProvinceName & " " & curVendor_.ZipCode & Chr(34)
        .Formulas(3) = "CompanyAddr2=" & Chr(34) & "โทร" & curVendor_.TelNo & Chr(34)
        .Formulas(4) = "BillpayNo=" & Chr(34) & txtSaleNo.Text & Chr(34)
        .Formulas(5) = "BillPayDate=" & Chr(34) & dtIssueDate.ValueDMY & Chr(34)
        .Formulas(6) = "CompanyBill=" & Chr(34) & cboCompany.Text & Chr(34)
        .Formulas(7) = "CompanyBillAddr=" & Chr(34) & CurrentUser.WorkCompanies(cboCompany.ListIndex).AddrTH1 & Chr(34)
        .Formulas(8) = "CompanyBillAddr1=" & Chr(34) & CurrentUser.WorkCompanies(cboCompany.ListIndex).AddrTH2 & CurrentUser.WorkCompanies(cboCompany.ListIndex).AddrTH3 & Chr(34)
        .Formulas(9) = "AmtB4Vat=" & Chr(34) & lbTotalMenifest.Caption & Chr(34)
        .Formulas(10) = "VATAmt=" & Chr(34) & lbVat.Caption & Chr(34)
        .Formulas(11) = "TotalAmtToPay=" & Chr(34) & lbNetAmt.Caption & Chr(34)
        
        tmpSql = " select * from vw_BillPayDetailSearch " & vbCrLf
        tmpSql = tmpSql & "WHERE BillPayHeaderID='" & myBillPay_.ID & "' " & vbCrLf
        tmpSql = tmpSql & "ORDER BY RowID"
        .SQLQuery = tmpSql
        .Destination = crptToWindow
        .PrintReport
        If Trim(.LastErrorString) <> "" Then MsgBox .LastErrorString
    End With
End Sub

Private Sub cmdSave_Click()
    If Trim(txtVendor.Text) = "" Then
        MsgBox "กรุณาเลือกเจ้าหนี้ที่ต้องการวางบิล", vbExclamation
        txtVendor.SetFocus
        Exit Sub
    End If
    If cboCompany.ListIndex < 0 Then
        MsgBox "กรุณาระบุบริษัทที่ออกใบวางบิล", vbExclamation
        cboCompany.SetFocus
        Exit Sub
    End If
    If myBillPay_.BillPayDetails.Count = 0 Then
        MsgBox "กรุณาระบุรายละเอียดที่จะทำการจ่ายเงิน", vbExclamation
        Exit Sub
    End If
    If myBillPay_.isChanged And myBillPay_.isCancel = False Then
        With myBillPay_
            .AmtB4VAT = lbTotalMenifest
            .VATAmt = lbVat
            .SasomAmt = lbDeduct
            .WTaxAmt = lbWTaxAmt
            .DeductB4VATAmt = Label3
            .FundAmt = lbFund
            .DeductAfterVATAmt = lbDAVat
            .TotalBillAmt = lbNetAmt
            .Save
        End With
        txtSaleNo.Text = myBillPay_.BillPayNo
    End If
    If myBillPay_.isChanged = False And myBillPay_.isCancel = False Then
        MsgBox "จัดเก็บข้อมูลเรียบร้อยแล้ว", vbInformation
    End If
End Sub

Private Sub cmdAddTimeTable_Click()
    If Trim(txtVendor.Text) = "" Then
        MsgBox "กรุณาระบุชื่อเจ้าหนี้", vbExclamation
        txtVendor.SetFocus
        Exit Sub
    End If
    isDetailChange_ = False
    Load frmTripControlAdd
    With frmTripControlAdd
        .SetCallerForm Me, curVendor_, myBillPay_
        .Show 1
    End With
    If isDetailChange_ Then
        ShowPayDetail
    End If
End Sub

Private Sub cmdSelectAll_Click()
Dim X%
    For X = 1 To lvOther.ListItems.Count
        If lvOther.ListItems(X).Checked = False Then
            lvOther.ListItems(X).Checked = True
            lvOther_ItemCheck lvOther.ListItems(X)
        End If
    Next
End Sub

Private Sub cmdDeSelectAll_Click()
Dim X%
    For X = 1 To lvOther.ListItems.Count
        If lvOther.ListItems(X).Checked = True Then
            lvOther.ListItems(X).Checked = False
            lvOther_ItemCheck lvOther.ListItems(X)
        End If
    Next
End Sub

Private Sub dtDueDate_Validate(Cancel As Boolean)
    myBillPay_.PayDueDate = dtDueDate.ValueYMD
End Sub

Private Sub dtIssueDate_Validate(Cancel As Boolean)
    myBillPay_.BillPayDate = dtIssueDate.ValueYMD
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 And Me.ActiveControl.Name <> "grItem" Then SendKeys "{Tab}"
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
    Label1(24).Caption = "หักเงินสะสม " & SystemConfig.VendorDeductRate & "%"
    Label1(26).Caption = "หักเข้ากองทุน (" & SystemConfig.VendorFundAmt & " บ./เที่ยว)"
    With grBill.getGridObj
        .Cols = 4
'        .TextMatrix(0, 1) = "วันที่"
'        .TextMatrix(0, 2) = "เลขที่ใบคุม"
'        .TextMatrix(0, 3) = "ต้นทาง"
'        .TextMatrix(0, 4) = "ปลายทาง"
'        .TextMatrix(0, 5) = "ยอดเงิน"
'        .ColWidth(0) = 0
'        .ColWidth(1) = 1500
'        .ColWidth(2) = 2265
'        .ColWidth(3) = 3060
'        .ColWidth(4) = 3060
'        .ColWidth(5) = 1455
'        .Rows = 50
'        .FixedAlignment(1) = 4
'        .FixedAlignment(2) = 4
'        .FixedAlignment(3) = 4
'        .FixedAlignment(4) = 4
'        .FixedAlignment(5) = 4
'        .ColAlignment(1) = 4
'        .ColAlignment(2) = 4
'        .ColAlignment(3) = 1
'        .ColAlignment(4) = 1
'        .ColAlignment(5) = 7
        
        .TextMatrix(0, 1) = "เลขที่ใบคุม"
        .TextMatrix(0, 2) = "รายการ"
        .TextMatrix(0, 3) = "จำนวนเงิน"
        .ColWidth(0) = 0
        .ColWidth(1) = 1500
        .ColWidth(3) = 2000
        .ColWidth(2) = .Width - (.ColWidth(1) + .ColWidth(3)) - 300
        .Rows = 50
        .FixedAlignment(1) = 4
        .FixedAlignment(2) = 4
        .FixedAlignment(3) = 4
        .ColAlignment(1) = 4
        .ColAlignment(2) = 1
        .ColAlignment(3) = 7
    End With
    With grBill
        .ColEnabled(1) = False
        .ColEnabled(2) = True
        .ColEnabled(3) = False
    End With
    With frmPicture
        cmdSave.Picture = .picSave.Picture
        cmdOpen.Picture = .picSearch.Picture
        cmdClear.Picture = .picNew.Picture
        cmdCancel.Picture = .picDelete.Picture
        cmdClose.Picture = .picExit.Picture
        cmdPrint.Picture = .picPrint
        cmdAddTimeTable.Picture = .picAdd
    End With
    With cboVAT
        .AddItem "ไม่คิด VAT"
        .AddItem "คิด VAT " & SystemConfig.VATRate & "%"
        .ListIndex = 0
    End With
    With grItem.getGridObj
        grItem.ColInputType(3) = ComboBoxEnum
        .Cols = 4
        .Rows = 50
        .TextMatrix(0, 1) = "รายการ"
        .TextMatrix(0, 2) = "จำนวนเงิน"
        .TextMatrix(0, 3) = "นำไปรวม"
        .ColWidth(0) = 0
        .ColWidth(1) = 3045
        .ColWidth(2) = 1200
        .ColWidth(3) = 1200
    End With
    With cboWTax
        .AddItem "0"
        .AddItem "1"
        .AddItem "3"
        .AddItem "5"
        .ListIndex = 0
    End With
    CreateNewBillPay
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload frmPaytoDeptorSearch
End Sub

Private Sub grBill_DblClick(Row As Integer, Col As Integer)
Dim tmpDet As BWGBillPayDetail
Dim tmpCol As Collection, tmpS As New BWGSearchManager
    With grBill.getGridObj
        If Trim(.TextMatrix(Row, 0)) <> "" Then
            Set tmpDet = myBillPay_.BillPayDetails("'" & .TextMatrix(Row, 0) & "'")
            'Set tmpCol = tmpS.JobDataTimeTableSearch("JobDataCarID='" & tmpDet.JobDataCarID & "' And TripNo=" & tmpDet.tripNo)
            Set tmpCol = tmpS.JobDataTimeTableSearch("TimeTableNo='" & tmpDet.TimeTableNo & "' And isMain='Y' And isCancel ='N'")
            If tmpCol.Count > 0 Then
                Load frmTripControl
                With frmTripControl
                    .SetTripControl tmpCol(1), tmpDet.CustomerName, tmpDet.TruckSubTypeDesc & ", " & tmpDet.TruckTypeDesc, tmpDet.CarRegisNo, tmpDet.FromAumphurID, tmpDet.ToAumphurID, tmpDet.TruckTypeID
                    .Show 1
                End With
            End If
            Set tmpCol = Nothing
            Set tmpS = Nothing
            Set tmpDet = Nothing
        End If
    End With
End Sub

Private Sub grBill_RolColChanged(Row As Integer, Col As Integer)
    If Col <> 2 Then grBill.ColEnabled(Col) = False
End Sub

Private Sub grItem_Click(Row As Integer, Col As Integer)
'    MsgBox grItem.getGridObj.ColWidth(col)
End Sub

Private Sub grItem_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim Row As Integer
    Dim tmpID As String, tmpVal As String
    Dim tmpItem As BWGBillPayDetail
    Dim tmpCol As Collection
    If KeyCode = vbKeyDelete Then
        Row = grItem.getGridObj.RowSel
        tmpID = Trim(grItem.getGridObj.TextMatrix(Row, 0))
        If myBillPay_.BillPayDetails Is Nothing Then Exit Sub
        If isExist(myBillPay_.BillPayDetails, "'" & tmpID & "'") Then
            RemoveTripControl tmpID
            Set tmpCol = myBillPay_.BillPayDetails
            tmpCol.Remove "'" & tmpID & "'"
            myBillPay_.isChanged = True
            grItem.getGridObj.RemoveItem Row
        End If
    End If
End Sub

Private Sub grItem_RolColChanged(Row As Integer, Col As Integer)
Dim tmpCbo As ComboBox
    With grItem.getGridObj
        If (Col = 1 And Trim(.TextMatrix(Row - 1, 1)) = "") Or (Col > 1 And Trim(.TextMatrix(Row, 1)) = "") Then
            grItem.ColEnabled(Col) = False
        Else
            grItem.ColEnabled(Col) = True
            If Col = 3 Then
                Set tmpCbo = grItem.getColObject(3)
                tmpCbo.Clear
                tmpCbo.AddItem "ก่อน VAT"
                tmpCbo.AddItem "หลัง VAT"
                Set tmpCbo = Nothing
            End If
        End If
    End With
End Sub

Private Sub grItem_UpdateValue(Row As Integer, Col As Integer)
Dim tmpItem As BWGBillPayDetail
Dim tmpID As String, tmpVal As String
    tmpID = Trim(grItem.getGridObj.TextMatrix(Row, 0))
    tmpVal = Trim(grItem.getGridObj.TextMatrix(Row, Col))
    If isExist(myBillPay_.BillPayDetails, "'" & tmpID & "'") Then
        Set tmpItem = myBillPay_.BillPayDetails("'" & tmpID & "'")
    Else
        Set tmpItem = New BWGBillPayDetail
        With tmpItem
            .ID = GetGUID
            .BillPayHeaderID = myBillPay_.ID
            .isCancel = False
            .RowID = getMaxRowID(myBillPay_.BillPayDetails) + 1
            myBillPay_.BillPayDetails.Add tmpItem, "'" & .ID & "'"
            grItem.getGridObj.TextMatrix(Row, 0) = Trim(.ID)
        End With
    End If
    Select Case Col
        Case 1
            tmpItem.BillDesc = Trim(tmpVal)
        Case 2
            If IsNumeric(tmpVal) Then
                tmpItem.AmtB4VAT = CDbl(tmpVal)
            Else
                tmpItem.AmtB4VAT = 0
            End If
            If Trim(tmpItem.BeforeOrAfterVAT) = "" Then
                tmpItem.BeforeOrAfterVAT = "B"
                grItem.getGridObj.TextMatrix(Row, 3) = "ก่อน VAT"
                If IsNumeric(grItem.getGridObj.TextMatrix(Row, Col)) = False Then
                    grItem.getGridObj.TextMatrix(Row, Col) = "0.00"
                End If
                grItem.getGridObj.TextMatrix(Row, Col) = Format(grItem.getGridObj.TextMatrix(Row, Col), "#,000.00")
            End If
            SumMoneyAmt
        Case 3
            If grItem.getColObject(Col).ListIndex = 0 Then
                tmpItem.BeforeOrAfterVAT = "B" 'Before
            ElseIf grItem.getColObject(Col).ListIndex = 1 Then
                tmpItem.BeforeOrAfterVAT = "A"  'After
            End If
            SumMoneyAmt
    End Select
    myBillPay_.isChanged = True
End Sub





Private Sub lvOther_DblClick()
On Error Resume Next
Dim tmpS As New BWGSearchManager
Dim tmpCol As Collection, tmpItem As BWGBillPayDetail
    If lvOther.SelectedItem Is Nothing Then Exit Sub
    Set tmpItem = colExpense_(lvOther.SelectedItem.Key)
    Set tmpCol = tmpS.VendorExpenseSearch("ExpenseID=" & lvOther.SelectedItem.Key & "")
    If tmpCol.Count > 0 Then
        Load frmVendorExpense
        With frmVendorExpense
            .SetDataVendorExpense tmpCol(1)
            .Show 1
        End With
    End If
    Set tmpItem = Nothing
    Set tmpS = Nothing
    Set tmpCol = Nothing
End Sub

Private Sub lvOther_ItemCheck(ByVal item As MSComctlLib.ListItem)
Dim X%, tmpCol As Collection
Dim tmpItem As BWGBillPayDetail, tmpID As String
Dim tmpExpense As BWGVendorOtherExpense
    Set tmpCol = myBillPay_.BillPayDetails
    tmpID = Replace(item.Key, "'", "")
    If item.Checked Then
        Set tmpExpense = colExpense_(item.Key)
        Set tmpItem = New BWGBillPayDetail
        With tmpItem
            .AmtB4VAT = tmpExpense.NetPrice
            .BillDesc = tmpExpense.BillDesc
            .BillPayHeaderID = myBillPay_.ID
            .ID = GetGUID
            .isCancel = False
            .ExpenseBillDesc = tmpExpense.BillDesc
            .ExpenseBuyQty = tmpExpense.BuyQty
            .ExpenseDate = tmpExpense.ExpenseDate
            .ExpenseID = tmpExpense.ID
            .ExpenseIsPaid = tmpExpense.isPaid
            .ExpenseNetPrice = tmpExpense.NetPrice
            .ExpenseVendorID = tmpExpense.VendorID
            .ExpenseVendorName = tmpExpense.VendorName
            If cboVAT.ListIndex = 1 Then
                .VATAmt = Format(.AmtB4VAT * (SystemConfig.VATRate) / 100, "0.00")
            ElseIf cboVAT.ListIndex = 0 Then
                .VATAmt = 0
            End If
            .TotalAmtToPay = .AmtB4VAT + .VATAmt
            tmpCol.Add tmpItem, "'" & tmpItem.ID & "'"
            .RowID = getMaxRowID(tmpCol) + 1
        End With
    Else
        For X = 1 To tmpCol.Count
            If Trim(tmpID) = tmpCol(X).ExpenseID Then
                tmpCol.Remove X
                Exit For
            End If
        Next
    End If
    myBillPay_.isChanged = True
    SumMoneyAmt
End Sub

Private Sub txtBillNo_Validate(Cancel As Boolean)
    myBillPay_.BillingNo = txtBillNo.Text
End Sub

Private Sub txtInvNo_Validate(Cancel As Boolean)
    myBillPay_.InvoiceNo = txtInvNo.Text
End Sub

Private Sub txtRecNo_Validate(Cancel As Boolean)
    myBillPay_.ReceiptNo = txtRecNo.Text
End Sub

Private Sub txtRem_Validate(Cancel As Boolean)
    myBillPay_.BillRemark = txtRem.Text
End Sub

Private Sub txtVendor_Validate(Cancel As Boolean)
Dim selStr As String
    If curVendor_ Is Nothing Then Set curVendor_ = New BWGTransportVendor
    selStr = Trim(curVendor_.NameTH)
    If Trim(txtVendor.Text) = "" And Trim(selStr) <> Trim(txtVendor.Text) Then
        Set curVendor_ = Nothing
        Set curVendor_ = New BWGTransportVendor
        With myBillPay_
            .SupplierID = ""
            .SupplierName = ""
        End With
    Else
        If Trim(selStr) <> Trim(txtVendor.Text) Then
            Set curVendor_ = Nothing
            Set curVendor_ = New BWGTransportVendor
            With myBillPay_
                .SupplierID = ""
                .SupplierName = ""
            End With
            Load frmTransportVenderSearch
            With frmTransportVenderSearch
                .Caption = "ค้นหาเจ้าหนี้การค้า Vendor"
                .getCriteria (txtVendor.Text)
                .SetCallerForm Me
                .Show 1
            End With
            If myBillPay_.SupplierID = "" Then txtVendor.Text = ""
        End If
    End If
End Sub

Private Sub SumMoneyAmt()
Dim tmpTotal As Double
Dim tmpTax As Double
Dim tmpTaxAmt As Double, afterVATAmt As Double
Dim X%, tripCount%
    With grBill.getGridObj
        For X = 1 To .Rows - 1
            If IsNumeric(.TextMatrix(X, .Cols - 1)) Then
                tmpTotal = tmpTotal + CDbl(.TextMatrix(X, .Cols - 1))
                tripCount = tripCount + 1
            End If
        Next
    End With
    lbTotalMenifest.Caption = Format(tmpTotal, "#,##0.00")
    lbDeduct.Caption = Format(CDbl(lbTotalMenifest.Caption) * SystemConfig.VendorDeductRate / 100, "#,##0.00")
    If IsNumeric(cboWTax.Text) Then
        tmpTax = cboWTax.Text
    Else
        tmpTax = 0
    End If
    lbWTaxAmt.Caption = Format(tmpTotal * tmpTax / 100, "#,##0.00")
    tmpTotal = 0
    With grItem.getGridObj
        For X = 1 To .Rows - 1
            If IsNumeric(.TextMatrix(X, 2)) Then
                If Trim(.TextMatrix(X, 3)) = "ก่อน VAT" Then
                    tmpTotal = tmpTotal + CDbl(.TextMatrix(X, 2))
                Else
                    afterVATAmt = afterVATAmt + CDbl(.TextMatrix(X, 2))
                End If
            End If
        Next
    End With
    Label3.Caption = Format(tmpTotal, "#,##0.00")
    If cboVAT.ListIndex = 1 Then
        lbVat.Caption = Format((CDbl(lbTotalMenifest.Caption) + CDbl(Label3.Caption)) * 0.07, "#,##0.00")
    ElseIf cboVAT.ListIndex = 0 Then
        lbVat.Caption = "0.00"
    End If
    lbFund.Caption = Format(tripCount * SystemConfig.VendorFundAmt, "#,##0.00")
    tmpTotal = 0
    For X = 1 To lvOther.ListItems.Count
        If lvOther.ListItems(X).Checked And IsNumeric(lvOther.ListItems(X).SubItems(2)) Then
            tmpTotal = tmpTotal + CDbl(lvOther.ListItems(X).SubItems(2))
        End If
    Next
    lbDAVat.Caption = Format(afterVATAmt + tmpTotal, "#,##0.00")
    tmpTotal = lbDAVat.Caption
    
    If IsNumeric(lbWTaxAmt.Caption) Then
        tmpTaxAmt = lbWTaxAmt.Caption
    End If
    
    lbNetAmt.Caption = Format(lbTotalMenifest - lbDeduct - lbWTaxAmt + Label3 + lbVat - lbFund + lbDAVat, "#,##0.00")
    
    'lbTotalPrice.Caption = Format(CDbl(lbNetPrice.Caption) / 1.07, "#,##0.00")
End Sub

Private Sub CreateNewBillPay()
    Set myBillPay_ = Nothing
    Set myBillPay_ = New BWGBillPayHeader
    With myBillPay_
        .ID = GetGUID
        .BillPayDate = dtIssueDate.ValueYMD
        .PayDueDate = dtDueDate.ValueYMD
        .isChanged = False
        If cboCompany.ListIndex >= 0 Then
            If cboCompany.ListCount > 1 Then
                If cboCompany.ListIndex = 0 Then Exit Sub
                myBillPay_.CompanyID = CurrentUser.WorkCompanies(cboCompany.ListIndex).ID
            Else
                myBillPay_.CompanyID = CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).ID
            End If
            myBillPay_.CompanyName = cboCompany.Text
        Else
            myBillPay_.CompanyID = ""
            myBillPay_.CompanyName = ""
        End If

        .isChanged = False
    End With
    SetCancelScreen
    myBillPay_.isChanged = False
End Sub

Private Sub ClearScreen()
    txtSaleNo.Text = "== AUTO =="
    dtIssueDate.ValueYMD = TodayDate
    grBill.ClearAllData
    grBill.getGridObj.Rows = 20
    txtVendor.Text = ""
    Set curVendor_ = Nothing
    Set curVendor_ = New BWGTransportVendor
    dtDueDate.ValueYMD = TodayDate
    txtBillNo.Text = ""
    txtInvNo.Text = ""
    txtRecNo.Text = ""
    cboWTax.ListIndex = 0
    cboVAT.ListIndex = 0
    txtRem.Text = ""
    lvOther.ListItems.Clear
    grItem.ClearAllData
    grItem.getGridObj.Rows = 50
    lbTotalMenifest.Caption = "0.00"
    lbDeduct.Caption = "0.00"
    lbFund.Caption = "0.00"
    lbVat.Caption = "0.00"
    lbDAVat.Caption = "0.00"
    lbWTaxAmt.Caption = "0.00"
    lbTotalPrice.Caption = "0.00"
    lbNetAmt.Caption = "0.00"
End Sub



Private Function getMaxRowID(selCol As Collection) As Integer
Dim X%, RowID%
    For X = 1 To selCol.Count
        If RowID < selCol(X).RowID Then
            RowID = selCol(X).RowID
        End If
    Next
    getMaxRowID = RowID
End Function

Public Sub SetDataPaytoDeptor(selBillPay As BWGBillPayHeader)
Dim tmpCol As Collection, tmpS As New BWGSearchManager
    Set myBillPay_ = Nothing
    Set myBillPay_ = selBillPay
    With myBillPay_
        txtSaleNo.Text = .BillPayNo
        dtIssueDate.ValueYMD = .BillPayDate
        SetTextToCombo .CompanyName, cboCompany
        Set tmpCol = tmpS.TransportVenderSearch("VendorID='" & .SupplierID & "'")
        If tmpCol.Count > 0 Then
            SetDataTransportVender tmpCol(1), False
        End If
        dtDueDate.ValueYMD = .PayDueDate
        txtBillNo.Text = .BillingNo
        txtInvNo.Text = .InvoiceNo
        txtRecNo.Text = .ReceiptNo
        SetTextToCombo .WTax, cboWTax
        If .VATRate > 0 Then
            cboVAT.ListIndex = 1
        Else
            cboVAT.ListIndex = 0
        End If
        txtRem.Text = .BillRemark
        Set tmpCol = Nothing
        Set tmpS = Nothing
        ShowPayDetail
        .isChanged = False
    End With
    SetCancelScreen
    SumMoneyAmt
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim selResult As Integer
    If myBillPay_ Is Nothing Then Exit Sub
    If myBillPay_.isChanged And myBillPay_.isCancel = False Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนปิดหน้าจอหรือไม่?", vbQuestion + vbYesNoCancel)
        If selResult = vbYes Then
            cmdSave_Click
            If myBillPay_.isChanged Then Cancel = 1
        ElseIf selResult = vbCancel Then
            Cancel = 1
        End If
    End If
End Sub

Private Sub SetCancelScreen()
Dim cVal As Boolean
    cVal = myBillPay_.isCancel
    lbCancel.Visible = cVal
    cmdSave.Enabled = Not cVal
    cmdCancel.Enabled = Not cVal
End Sub

Public Function GetTripControls() As Collection
Dim X%, tmpS As New BWGSearchManager
Dim tmpCol As New Collection, tmpTrip As BWGJobDataTimeTable
Dim tmpDet As BWGBillPayDetail
Dim selCol As Collection, tmpStr As String
    Set selCol = myBillPay_.BillPayDetails
    With selCol
        For X = 1 To .Count
            Set tmpDet = .item(X)
            Set tmpTrip = New BWGJobDataTimeTable
            If Trim(tmpDet.TimeTableNo) <> "" Then
                'tmpTrip.JobDataCarID = tmpDet.JobDataCarID
                tmpTrip.CompanyCode = tmpDet.CompanyCode
                tmpTrip.TripNo = tmpDet.TripNo
                tmpTrip.TimeTableNo = tmpDet.TimeTableNo
                tmpTrip.WorkDate = tmpDet.WorkDate
                tmpTrip.TruckSubTypeDesc = tmpDet.TruckSubTypeDesc
                tmpTrip.TruckSubTypeID = tmpDet.TruckSubTypeID
                tmpTrip.TruckTypeDesc = tmpDet.TruckTypeDesc
                tmpTrip.TruckTypeID = tmpDet.TruckTypeID
                tmpTrip.CarRegisNo = tmpDet.CarRegisNo
                tmpTrip.TrailerTruckNo = tmpDet.TrailerTruckNo
                tmpTrip.DriverName = tmpDet.DriverName
                tmpTrip.FromAumphurID = tmpDet.FromAumphurID
                tmpTrip.FromAumphurName = tmpDet.FromAumphurName
                tmpTrip.FromProvinceID = tmpDet.FromProvinceID
                tmpTrip.FromProvinceName = tmpDet.FromProvinceName
                tmpTrip.ToAumphurID = tmpDet.ToAumphurID
                tmpTrip.ToAumphurName = tmpDet.ToAumphurName
                tmpTrip.ToProvinceID = tmpDet.ToProvinceID
                tmpTrip.ToProvinceName = tmpDet.ToProvinceName
                tmpTrip.TripTranFee = tmpDet.AmtB4VAT
                tmpCol.Add tmpTrip, "'" & tmpTrip.TimeTableNo & "_" & tmpTrip.TripNo & "'"
            End If
            Set tmpTrip = Nothing
            Set tmpDet = Nothing
        Next
    End With

    Set GetTripControls = tmpCol
    Set tmpCol = Nothing
End Function

Public Sub AddTripControl(selItem As BWGJobDataTimeTable)
Dim tmpCol As Collection
Dim X%, isNew As Boolean
Dim tmpID As String, tmpCarID As String, tmpTripNo As Integer
Dim tmpDetail As New BWGBillPayDetail
    Set tmpCol = myBillPay_.BillPayDetails
    
    If tmpCol Is Nothing Then Set tmpCol = New Collection
    isNew = True
    tmpCarID = selItem.TimeTableNo
    tmpTripNo = selItem.TripNo
    For X = 1 To tmpCol.Count
        If Trim(tmpCol(X).TimeTableNo) = Trim(tmpCarID) And tmpCol(X).TripNo = tmpTripNo Then
            Set tmpDetail = tmpCol(X)
            isNew = False
            Exit For
        End If
    Next

    If isNew Then
        tmpDetail.ID = GetGUID
        tmpCol.Add tmpDetail, "'" & tmpDetail.ID & "'"
    End If
    tmpDetail.BillPayHeaderID = myBillPay_.ID
'    tmpDetail.JobDataCarID = selItem.JobDataCarID
    tmpDetail.CompanyCode = selItem.CompanyCode
    tmpDetail.TripNo = selItem.TripNo
    tmpDetail.TimeTableNo = selItem.TimeTableNo
    tmpDetail.WorkDate = selItem.WorkDate
    tmpDetail.TruckSubTypeDesc = selItem.TruckSubTypeDesc
    tmpDetail.TruckSubTypeID = selItem.TruckSubTypeID
    tmpDetail.TruckTypeDesc = selItem.TruckTypeDesc
    tmpDetail.TruckTypeID = selItem.TruckTypeID
    tmpDetail.CarRegisNo = selItem.CarRegisNo
    tmpDetail.TrailerTruckNo = selItem.TrailerTruckNo
    tmpDetail.DriverName = selItem.DriverName
    tmpDetail.FromAumphurID = selItem.FromAumphurID
    tmpDetail.FromAumphurName = selItem.FromAumphurName
    tmpDetail.FromProvinceID = selItem.FromProvinceID
    tmpDetail.FromProvinceName = selItem.FromProvinceName
    tmpDetail.ToAumphurID = selItem.ToAumphurID
    tmpDetail.ToAumphurName = selItem.ToAumphurName
    tmpDetail.ToProvinceID = selItem.ToProvinceID
    tmpDetail.ToProvinceName = selItem.ToProvinceName
    tmpDetail.RowID = tmpCol.Count
    'tmpDetail.AmtB4VAT = selItem.TripEquipFee + selItem.TripMoreWorkerFee + selItem.TripOtherFee + selItem.TripTranFee
    tmpDetail.AmtB4VAT = selItem.VDEquipFee + selItem.VDWorkerFee + selItem.VDOtherFee + selItem.VDTranFee
    tmpDetail.BillDesc = "ค่าขนส่งวันที่" & selItem.WorkDate & "เลขที่ " & selItem.TimeTableNo '& " ทะเบียน " & selItem.CarRegisNo & IIf(selItem.TrailerTruckNo = "", "", " - " & selItem.TrailerTruckNo & " (พ่วง)")
    tmpDetail.TotalAmtToPay = selItem.VDEquipFee + selItem.VDWorkerFee + selItem.VDOtherFee + selItem.VDTranFee
    isDetailChange_ = True
    myBillPay_.isChanged = True
    Set tmpDetail = Nothing
    
    If myBillPay_.BillPayDetails Is Nothing Then
        Set myBillPay_.BillPayDetails = tmpCol
        isDetailChange_ = True
        myBillPay_.isChanged = True
    End If
    Set tmpCol = Nothing
End Sub

Public Sub RemoveTripControl(selID As String)
Dim tmpCol As Collection
Dim X%, tmpKey As String
    Set tmpCol = myBillPay_.BillPayDetails
    If tmpCol Is Nothing Then Exit Sub
    For X = 1 To tmpCol.Count
        If "'" & tmpCol(X).TimeTableNo & "_" & tmpCol(X).TripNo & "'" = Trim(selID) Then
            tmpKey = "'" & tmpCol(X).ID & "'"
            Exit For
        End If
    Next
    If isExist(tmpCol, tmpKey) Then
        tmpCol.Remove tmpKey
        isDetailChange_ = True
        myBillPay_.isChanged = True
    End If
End Sub

Public Property Get RefID() As String
    RefID = myBillPay_.ID
End Property

