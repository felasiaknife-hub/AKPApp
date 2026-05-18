VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmExtCom2 
   Caption         =   "ค่าคอมฯ นอก"
   ClientHeight    =   10290
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   11280
   Icon            =   "frmExtCom2.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10290
   ScaleWidth      =   11280
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   10455
      Left            =   2640
      TabIndex        =   27
      Top             =   -120
      Width           =   9255
      Begin TabDlg.SSTab SSTab1 
         Height          =   8475
         Left            =   0
         TabIndex        =   17
         Top             =   120
         Width           =   8715
         _ExtentX        =   15372
         _ExtentY        =   14949
         _Version        =   393216
         Tabs            =   1
         TabHeight       =   520
         TabCaption(0)   =   "รายการค่าคอมฯนอก"
         TabPicture(0)   =   "frmExtCom2.frx":0ECA
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Result"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         Begin MSComctlLib.ListView Result 
            Height          =   2355
            Left            =   600
            TabIndex        =   28
            Top             =   660
            Width           =   2865
            _ExtentX        =   5054
            _ExtentY        =   4154
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   16777215
            BorderStyle     =   1
            Appearance      =   0
            NumItems        =   40
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "เขต"
               Object.Width           =   1411
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "SaleStaffID"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "ผู้แทนขาย"
               Object.Width           =   2469
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "CustomerID"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "ชื่อลูกค้า"
               Object.Width           =   5292
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   5
               Text            =   "วันที่ขน"
               Object.Width           =   1940
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "MenifestID"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Text            =   "เลขที่ Manifest"
               Object.Width           =   2999
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   8
               Text            =   "WasteDataID"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   9
               Text            =   "ชื่อกาก"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   10
               Text            =   "เลขที่ RE"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   11
               Text            =   "วันที่รับชำระ"
               Object.Width           =   1940
            EndProperty
            BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   12
               Text            =   "น้ำหนัก"
               Object.Width           =   1411
            EndProperty
            BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   13
               Text            =   "อัตราค่าบริการ"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   14
               Text            =   "หน่วย"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   15
               Text            =   "ค่าบริการ"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   16
               Text            =   "ค่าขนส่ง"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   17
               Text            =   "อัตราค่าคอม"
               Object.Width           =   1940
            EndProperty
            BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   18
               Text            =   "หน่วย"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   19
               Text            =   "ค่าคอมฯ"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   20
               Text            =   "อัตราภาษี"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   21
               Text            =   "มูลค่าภาษี"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   22
               Text            =   "ค่าคอมฯสุทธิ"
               Object.Width           =   2469
            EndProperty
            BeginProperty ColumnHeader(24) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   23
               Text            =   "WasteType"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(25) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   24
               Text            =   "SaleZoneID"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(26) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   25
               Text            =   "CompanyID"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(27) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   26
               Text            =   "StaffLName"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(28) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   27
               Text            =   "TimeTableNo"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(29) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   28
               Text            =   "TreatMentCharge"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(30) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   29
               Text            =   "WasteType"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(31) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   30
               Text            =   "ZoneName"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(32) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   31
               Text            =   "IsCancel"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(33) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   32
               Text            =   "SaleDistrictID"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(34) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   33
               Text            =   "ComRecName"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(35) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   34
               Text            =   "ComRatePC"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(36) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   35
               Text            =   "RateTypePC"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(37) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   36
               Text            =   "DocCustWeight"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(38) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   37
               Text            =   "Status"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(39) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   38
               Text            =   "BillDate"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(40) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   39
               Text            =   "BillNo"
               Object.Width           =   0
            EndProperty
         End
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   10455
      Left            =   0
      TabIndex        =   19
      Top             =   -120
      Width           =   2655
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "รายการที่ตั้งค่าคอมฯแล้ว"
         Height          =   255
         Left            =   300
         TabIndex        =   8
         Top             =   4620
         Width           =   2055
      End
      Begin VB.CommandButton cmdDoComExt_Del 
         Caption         =   "ยกเลิกยอดตั้งค่าคอมฯ"
         Height          =   435
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   6480
         Width           =   2115
      End
      Begin VB.ComboBox cboDateType 
         Height          =   315
         Left            =   1140
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   900
         Width           =   1275
      End
      Begin VB.CommandButton cmdPrintExt1 
         Caption         =   "พิมพ์ตั้งยอดค่าคอมฯ"
         Height          =   435
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   7560
         Width           =   2115
      End
      Begin VB.CommandButton cmdDoComExt_Cancel 
         Caption         =   "ยกเลิกการจ่ายค่าคอมฯ"
         Height          =   435
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   6960
         Width           =   2115
      End
      Begin VB.CommandButton cmdDoComExt_Base 
         Caption         =   "ตั้งยอดค่าคอมฯ"
         Height          =   435
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   5400
         Width           =   2115
      End
      Begin VB.CommandButton cmdPrintExt2 
         Caption         =   "พิมพ์ค่าคอมฯ ที่ตัดจ่ายแล้ว"
         Height          =   435
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   8040
         Width           =   2115
      End
      Begin VB.CommandButton cmdDoComExt_Pay 
         Caption         =   "ตัดจ่ายค่าคอมฯ"
         Height          =   435
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   5880
         Width           =   2115
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "ปิด"
         Height          =   435
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   8640
         Width           =   2115
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "ค้นหา"
         Height          =   375
         Left            =   300
         TabIndex        =   9
         Top             =   4920
         Width           =   2055
      End
      Begin VB.ComboBox cboZone 
         Height          =   315
         Left            =   180
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   2280
         Width           =   2235
      End
      Begin VB.TextBox txtSale 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   180
         TabIndex        =   5
         Text            =   "*"
         Top             =   2940
         Width           =   2235
      End
      Begin VB.ComboBox cboCompany 
         Height          =   315
         Left            =   180
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   540
         Width           =   2235
      End
      Begin VB.ComboBox cboWaste 
         Height          =   315
         Left            =   180
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   3540
         Width           =   2235
      End
      Begin VB.TextBox txtCustomer 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   180
         TabIndex        =   7
         Text            =   "*"
         Top             =   4200
         Width           =   2235
      End
      Begin Commission.ctlDate dtFrom 
         Height          =   375
         Left            =   600
         TabIndex        =   2
         Top             =   1260
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         BackColor       =   -2147483633
      End
      Begin Commission.ctlDate dtTo 
         Height          =   375
         Left            =   600
         TabIndex        =   3
         Top             =   1680
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         BackColor       =   -2147483633
      End
      Begin Crystal.CrystalReport rpt1 
         Left            =   1920
         Top             =   9240
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowState     =   2
         PrintFileLinesPerPage=   60
         WindowShowPrintSetupBtn=   -1  'True
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H00800080&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   4
         Left            =   240
         Top             =   9180
         Width           =   195
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ไม่อนุมัติจ่าย"
         ForeColor       =   &H00C000C0&
         Height          =   195
         Index           =   1
         Left            =   540
         TabIndex        =   34
         Top             =   9180
         Width           =   885
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ยกเลิก"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   0
         Left            =   540
         TabIndex        =   33
         Top             =   9420
         Width           =   480
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   3
         Left            =   240
         Top             =   9420
         Width           =   195
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "รูปแบบวันที่"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   32
         Top             =   960
         Width           =   915
      End
      Begin VB.Shape Shape1 
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   2
         Left            =   240
         Top             =   9660
         Width           =   195
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "รายการใหม่"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   540
         TabIndex        =   31
         Top             =   9660
         Width           =   840
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H00008000&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   1
         Left            =   240
         Top             =   10140
         Width           =   195
      End
      Begin VB.Label lbApproved 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ตัดจ่ายค่าคอมฯ แล้ว"
         ForeColor       =   &H00008000&
         Height          =   195
         Left            =   540
         TabIndex        =   30
         Top             =   10140
         Width           =   1395
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   0
         Left            =   240
         Top             =   9900
         Width           =   195
      End
      Begin VB.Label lbOpen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ตั้งยอดค่าคอมฯ แล้ว"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   540
         TabIndex        =   29
         Top             =   9900
         Width           =   1395
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ผู้แทนขาย"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   26
         Top             =   2700
         Width           =   765
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "โซน"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   8
         Left            =   180
         TabIndex        =   25
         Top             =   2040
         Width           =   300
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "บริษัท"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   24
         Top             =   300
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ชนิด waste"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   240
         TabIndex        =   23
         Top             =   3300
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "วันที่"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   180
         TabIndex        =   22
         Top             =   1320
         Width           =   345
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ถึง"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   6
         Left            =   225
         TabIndex        =   21
         Top             =   1740
         Width           =   210
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ชื่อลูกค้า"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   4
         Left            =   180
         TabIndex        =   20
         Top             =   3960
         Width           =   615
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   18
      Top             =   9915
      Width           =   11280
      _ExtentX        =   19897
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmExtCom2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim curStaff_ As BWGCompanyStaff
Dim colZone_ As Collection
Dim sessionID_ As String
Dim curCustomer_ As BWGCustomer
Dim Counts As Long
Dim lvIndex As Long
Dim tList As ListItem
Dim GetBaseDate As String
Dim GetBase As Boolean
Dim tmp_CommRE As String

Private Sub CheckRE()
On Error GoTo ErrD
Dim RS As ADODB.Recordset
Dim i As Integer, ErrCn As Integer
Dim tSql As String
    If Result.ListItems.Count > 0 Then
        For i = 1 To Result.ListItems.Count
            tSql = "SELECT  *  FROM Com_vwComEx2_Union" & vbCrLf
            tSql = tSql & SqlQuery(cboDateType.ListIndex)
            tSql = tSql & "AND MenifestID = '" & Result.ListItems(i).SubItems(6) & "' AND WasteDataID = '" & Result.ListItems(i).SubItems(8) & "'" & vbCrLf
            tSql = tSql & "ORDER BY TimeTableNo, WorkDate, DocNo, WasteType"
            Set RS = New ADODB.Recordset
            Set RS = GetRS(tSql, adOpenForwardOnly, adLockReadOnly)
            Do While Not RS.EOF
                Result.ListItems(i).SubItems(10) = RS!RENo
                Result.ListItems(i).SubItems(11) = Format(RS!REDate, "dd/mm/yyyy")
                Result.ListItems(i).SubItems(38) = Format(RS!IssuedDate, "dd/mm/yyyy")
                Result.ListItems(i).SubItems(39) = RS!BillingNo
                Result.Refresh
                RS.MoveNext
            Loop
            RS.Close
            Set RS = Nothing
        Next
    End If
    Exit Sub
ErrD:
    ErrCn = ErrCn + 1
    MsgBox Err.Description
    If ErrCn > 5 Then Exit Sub
    Resume Next
End Sub

Public Sub SetBaseDate(ByVal BaseDate As String, ByVal GetBase_ As Boolean)
    GetBaseDate = BaseDate
    GetBase = GetBase_
End Sub

Private Sub cboDateType_Validate(Cancel As Boolean)
    If cboDateType.ListIndex > 0 Then
        cmdDoComExt_Base.Enabled = False
    Else
        cmdDoComExt_Base.Enabled = True
    End If
End Sub

Private Sub cmdDoComExt_Base_Click()
On Error GoTo ErrD
Dim RS  As ADODB.Recordset
Dim Status As String, tSql As String
Dim ManifestID As String, wastedataID As String
Dim i, j, k As Integer, BaseCount As Integer
Dim tList As ListItem
    If Result.ListItems.Count = 0 Then
        MsgBox "No Item, please search Item again.", vbInformation, "Warning"
        Exit Sub
    End If
    Load frmBaseDate
    frmBaseDate.Show 1
    If GetBase = False Then Exit Sub
    ManifestID = ""
    wastedataID = ""
    BaseCount = 0
        
    For i = 1 To Result.ListItems.Count
        If Result.ListItems(i).Checked = True Then
            If Result.ListItems(i).SubItems(37) = "1" Then
                k = i
                If ManifestID = "" Then
                    ManifestID = "'" & Result.ListItems(i).SubItems(6) & "'"
                    wastedataID = "'" & Result.ListItems(i).SubItems(8) & "'"
                Else
                    ManifestID = ManifestID & ", '" & Result.ListItems(i).SubItems(6) & "'"
                    wastedataID = wastedataID & ", '" & Result.ListItems(i).SubItems(8) & "'"
                End If
            End If
        End If
    Next
    
    For i = 1 To Result.ListItems.Count
        If Result.ListItems(i).Checked = True Then
            tSql = "SELECT * FROM tbCom_Commission_External" & vbCrLf
            tSql = tSql & "WHERE MenifestID IN ('" & Result.ListItems(i).SubItems(6) & "')" & vbCrLf
            tSql = tSql & "AND WasteDataID IN ('" & Result.ListItems(i).SubItems(8) & "')" & vbCrLf
            Set RS = New ADODB.Recordset
            Set RS = GetRS(tSql, adOpenForwardOnly, adLockReadOnly)
            If RS.RecordCount > 0 Then
                If Result.ListItems(i).SubItems(37) = "1" Then
                    tSql = "UPDATE tbCom_Commission_External SET Status = 2, BaseDate = '" & GetBaseDate & "'," & vbCrLf
                    tSql = tSql & "RENo = '" & Result.ListItems(i).SubItems(10) & "', REDate = '" & Format(Result.ListItems(i).SubItems(11), "yyyy/mm/dd") & "'" & vbCrLf
                    tSql = tSql & "WHERE MenifestID IN ('" & Result.ListItems(i).SubItems(6) & "')" & vbCrLf
                    tSql = tSql & "AND WasteDataID IN ('" & Result.ListItems(i).SubItems(8) & "')" & vbCrLf
                    tSql = tSql & "AND isCanceled <> 'Y'"
                    DBConnExc tSql
                End If
            Else
                tSql = "INSERT INTO tbCom_Commission_External(SDName, StaffFName, StaffLName, CompanyName, WorkDate, TimeTableNo, MenifestID, DocNo, Weight, InCollectFee, TreatmentCharge, " & _
                      "TransportFee, NonOrHz, WasteType, ZoneName, CompanyID, SaleZoneID, UnderSaleStaffID, CustomerID, CustomerCode, isCanceled, SaleDistrictID, WasteDataID," & _
                      "WasteName , TreatmentRate, TreatmentUnitID, ComRateAmt, RateType, ComRecName, ComRatePC, RateTypePC, DocCustWeight, RENo, REDate, BaseDate, PaidDate, Status, isPriceIncTrans, TripTransFee_Inc)" & vbCrLf
                tSql = tSql & "SELECT     SDName, StaffFName, StaffLName, CompanyName, WorkDate, TimeTableNo, MenifestID, DocNo, Weight, InCollectFee, TreatmentCharge, " & vbCrLf
                tSql = tSql & "                      TransportFee, NonOrHz, WasteType, ZoneName, CompanyID, SaleZoneID, UnderSaleStaffID, CustomerID, CustomerCode, isCanceled, SaleDistrictID, WasteDataID," & vbCrLf
                tSql = tSql & "                      WasteName , TreatmentRate, TreatmentUnitID, ComRateAmt, RateType, ComRecName, ComRatePC, RateTypePC, DocCustWeight, '" & Result.ListItems(i).SubItems(10) & "' AS RENo, '" & Format(Result.ListItems(i).SubItems(11), "yyyy/MM/dd") & "' AS REDate, '" & GetBaseDate & "', '' AS PaidDate, 2 AS Status, isPriceIncTrans, ISNULL(TripTransFee_Inc, 0)" & vbCrLf
                tSql = tSql & "FROM         dbo.Com_vwComEx_Union" & vbCrLf
                tSql = tSql & "WHERE MenifestID IN ('" & Result.ListItems(i).SubItems(6) & "')" & vbCrLf
                tSql = tSql & "AND WasteDataID IN ('" & Result.ListItems(i).SubItems(8) & "')" & vbCrLf
                DBConnExc tSql
            End If
            k = i
            BaseCount = BaseCount + 1
            Result.ListItems(i).SubItems(37) = "2"
            Set tList = Result.ListItems(i)
            For j = 1 To Result.ColumnHeaders.Count - 1
                tList.ListSubItems(j).ForeColor = vbBlue
            Next
            Result.Refresh
        End If
    Next
    MsgBox "ตั้งยอดค่าคอมฯ นอกแล้ว " & BaseCount & " รายการ", vbInformation, "External Com"
    Exit Sub
    
ErrD:
    If k > 0 Then Result.ListItems(k).SubItems(37) = "1"
    MsgBox Err.Description
    Exit Sub
End Sub

Private Sub cmdDoComExt_Cancel_Click()
On Error Resume Next
Dim i, j As Integer, CN As Integer
Dim tSql As String, Msg As String
    CN = 0
    If MsgBox("ยืนยันการยกเลิกการตัดจ่ายค่าคอมฯ นอก", vbExclamation + vbYesNo, "Warning") = vbNo Then Exit Sub
    For i = 1 To Result.ListItems.Count
        If Result.ListItems.item(i).Checked = True Then
            If Result.ListItems(i).SubItems(37) = "3" Then
            
                tSql = "DELETE FROM tbCom_ExtCom_Paid" & vbCrLf
                tSql = tSql & "WHERE MenifestID = '" & Result.ListItems(i).SubItems(6) & "' AND WasteDataID = '" & Result.ListItems(i).SubItems(8) & "'"
                DBConnExc tSql
            
                tSql = "UPDATE tbCom_Commission_External SET Status = 2" & vbCrLf
                tSql = tSql & "WHERE MenifestID = '" & Result.ListItems(i).SubItems(6) & "' AND WasteDataID = '" & Result.ListItems(i).SubItems(8) & "'"
                DBConnExc tSql
                
                Result.ListItems(i).SubItems(37) = "2"
                For j = 1 To Result.ColumnHeaders.Count - 1
                    Result.ListItems.item(i).ListSubItems(j).ForeColor = vbBlue
                Next
                Result.Refresh
                CN = CN + 1
            Else
                Result.ListItems(i).Checked = False
                DoEvents
            End If
        End If
    Next
    Msg = "ยกเลิกการจัดจ่ายค่าคอมฯ " & CN & " รายการ"
    MsgBox Msg
    StatusBar1.Panels(1).Text = Msg
End Sub

Private Sub cmdDoComExt_Del_Click()
On Error Resume Next
Dim i, j As Integer, CN As Integer
Dim tSql As String, Msg As String
    CN = 0
    For i = 1 To Result.ListItems.Count
        If Result.ListItems.item(i).Checked = True Then
            If Result.ListItems(i).SubItems(37) = "2" Then
                tSql = "UPDATE tbCom_Commission_External SET Status = 1, BaseDate = ''" & vbCrLf
                tSql = tSql & "WHERE MenifestID = '" & Result.ListItems(i).SubItems(6) & "' AND WasteDataID = '" & Result.ListItems(i).SubItems(8) & "'"
                DBConnExc tSql
                
                tSql = "DELETE FROM tbCom_Commission_External " & vbCrLf
                tSql = tSql & "WHERE MenifestID = '" & Result.ListItems(i).SubItems(6) & "' AND WasteDataID = '" & Result.ListItems(i).SubItems(8) & "'"
                DBConnExc tSql
                
                Result.ListItems(i).SubItems(37) = "1"
                For j = 1 To Result.ColumnHeaders.Count - 1
                    Result.ListItems.item(i).ListSubItems(j).ForeColor = vbBlack
                Next
                Result.Refresh
                CN = CN + 1
            Else
                Msg = "(Menifest No " & Result.ListItems.item(i).SubItems(7) & ") " & "ไม่สามารถลบได้"
                StatusBar1.Panels(1).Text = Msg
                Result.ListItems(i).Checked = False
                DoEvents
            End If
        End If
    Next
    Msg = "ลบข้อมูลการตั้งยอด " & CN & " รายการ"
    MsgBox Msg
    StatusBar1.Panels(1).Text = Msg
End Sub

Private Sub cmdPrintExt2_Click()
    PrintMainExt2 1
End Sub

Private Sub cmdPrintExt1_Click()
    PrintMainExt1 0
End Sub

Private Sub cmdSearch_Click()
On Error Resume Next
Dim RS As ADODB.Recordset
Dim tSql As String, tSql1 As String, tSql2 As String
Dim CN As Long
Dim isBase As Boolean

    If cboDateType.ListIndex > 0 Then
        cmdDoComExt_Base.Enabled = False
    Else
        cmdDoComExt_Base.Enabled = True
    End If
    
    isBase = Check1.Value
    Result.ListItems.Clear
    DoEvents
    tmp_CommRE = "tbCom_CommissionRE_" & ComName
        
    StatusBar1.Panels(1).Text = "**กำลังค้นหาข้อมูล..."
    tSql = "SELECT  *  FROM tbCom_Commission_External" & vbCrLf
    tSql = tSql & SqlQuery(cboDateType.ListIndex)
    tSql = tSql & "ORDER BY WorkDate, TimeTableNo,InCollectFee, TripTransFee_Inc Desc"
    
    Set RS = New ADODB.Recordset
    Set RS = GetRS(tSql, adOpenForwardOnly, adLockReadOnly)
    If isBase = False Then
        If RS.RecordCount = 0 Then
            If cboDateType.ListIndex = 0 Then
                tSql = "SELECT     SDName, StaffFName, StaffLName, CompanyName, WorkDate, TimeTableNo, MenifestID, DocNo, Weight, InCollectFee, TreatmentCharge, " & vbCrLf
                tSql = tSql & "                      isPriceIncTrans, TripTransFee_Inc, NonOrHz, WasteType, ZoneName, CompanyID, SaleZoneID, UnderSaleStaffID, CustomerID, isCanceled, SaleDistrictID, WasteDataID," & vbCrLf
                tSql = tSql & "                      WasteName , TreatmentRate, TreatmentUnitID, ComRateAmt, RateType, ComRecName, ComRatePC, RateTypePC, DocCustWeight, '' AS RENo, '' AS REDate, '' AS PaidDate, 1 AS Status" & vbCrLf
                tSql = tSql & "INTO " & tmp_CommRE & vbCrLf
                tSql = tSql & "FROM         dbo.Com_vwComEx_Union" & vbCrLf
                tSql = tSql & SqlQuery(cboDateType.ListIndex)
                tSql = tSql & "ORDER BY WorkDate, TimeTableNo, InCollectFee,TripTransFee_Inc Desc"
            Else
                tSql = "SELECT     SDName, StaffFName, StaffLName, CompanyName, WorkDate, TimeTableNo, MenifestID, DocNo, Weight, InCollectFee, TreatmentCharge, " & vbCrLf
                tSql = tSql & "                      isPriceIncTrans, TripTransFee_Inc, NonOrHz, WasteType, ZoneName, CompanyID, SaleZoneID, UnderSaleStaffID, CustomerID, isCanceled, SaleDistrictID, WasteDataID," & vbCrLf
                tSql = tSql & "                      WasteName , TreatmentRate, TreatmentUnitID, ComRateAmt, RateType, ComRecName, ComRatePC, RateTypePC, DocCustWeight, RENo, REDate, '' AS PaidDate, 1 Status" & vbCrLf
                tSql = tSql & "INTO " & tmp_CommRE & vbCrLf
                tSql = tSql & "FROM         dbo.Com_vwComEx2_Union" & vbCrLf
                tSql = tSql & SqlQuery(cboDateType.ListIndex)
                tSql = tSql & "ORDER BY WorkDate, TimeTableNo, InCollectFee,TripTransFee_Inc Desc"
            End If
        Else
            CN = RS.RecordCount
            AddtoListView RS, 1
            Result.Refresh

            If cboDateType.ListIndex = 0 Then
                tSql = "SELECT     SDName, StaffFName, StaffLName, CompanyName, WorkDate, TimeTableNo, MenifestID, DocNo, Weight, InCollectFee, TreatmentCharge, " & vbCrLf
                tSql = tSql & "                      isPriceIncTrans, TripTransFee_Inc, NonOrHz, WasteType, ZoneName, CompanyID, SaleZoneID, UnderSaleStaffID, CustomerID, isCanceled, SaleDistrictID, WasteDataID," & vbCrLf
                tSql = tSql & "                      WasteName , TreatmentRate, TreatmentUnitID, ComRateAmt, RateType, ComRecName, ComRatePC, RateTypePC, DocCustWeight, '' AS RENo, '' AS REDate, '' AS PaidDate, 1 AS Status" & vbCrLf
                tSql = tSql & "INTO " & tmp_CommRE & vbCrLf
                tSql = tSql & "FROM         dbo.Com_vwComEx_Union" & vbCrLf
                tSql = tSql & SqlQuery(cboDateType.ListIndex) & vbCrLf
                tSql = tSql & "AND MenifestID + '_' + WasteDataID NOT IN ("
                tSql = tSql & "SELECT DISTINCT  MenifestID + '_' + WasteDataID  FROM tbCom_Commission_External" & vbCrLf
                tSql = tSql & SqlQuery(cboDateType.ListIndex) & ")"
                tSql = tSql & "ORDER BY WorkDate, TimeTableNo, InCollectFee,TripTransFee_Inc Desc"
            Else
                tSql = "SELECT     SDName, StaffFName, StaffLName, CompanyName, WorkDate, TimeTableNo, MenifestID, DocNo, Weight, InCollectFee, TreatmentCharge, " & vbCrLf
                tSql = tSql & "                      isPriceIncTrans, TripTransFee_Inc, NonOrHz, WasteType, ZoneName, CompanyID, SaleZoneID, UnderSaleStaffID, CustomerID, isCanceled, SaleDistrictID, WasteDataID," & vbCrLf
                tSql = tSql & "                      WasteName , TreatmentRate, TreatmentUnitID, ComRateAmt, RateType, ComRecName, ComRatePC, RateTypePC, DocCustWeight, '' AS RENo, '' AS REDate, '' AS PaidDate, 1 AS Status" & vbCrLf
                tSql = tSql & "INTO " & tmp_CommRE & vbCrLf
                tSql = tSql & "FROM         dbo.Com_vwComEx_Union" & vbCrLf
                tSql = tSql & SqlQuery(cboDateType.ListIndex)
                tSql = tSql & "ORDER WorkDate, TimeTableNo, InCollectFee,TripTransFee_Inc Desc"
            End If
        End If
    Else
        CN = RS.RecordCount
        AddtoListView RS, 1
        Result.Refresh
        StatusBar1.Panels(1).Text = "ผลการค้นหา พบ " & CN & " รายการ"
'        MsgBox StatusBar1.Panels(1).Text, vbInformation + vbOKOnly, "Result"
    End If
    
    If isBase = False Then
        tSql1 = Replace(tSql, "'", "''")
        tSql2 = "EXEC sp_GetCommission_Pay '" & tmp_CommRE & "', '" & tSql1 & "'"
        Set RS = New ADODB.Recordset
        DBConnExc tSql2
        tSql = "SELECT * FROM " & tmp_CommRE & vbCrLf
        tSql = tSql & "ORDER BY WorkDate, TimeTableNo, InCollectFee,TripTransFee_Inc Desc"
        Set RS = GetRS(tSql, adOpenForwardOnly, adLockReadOnly)
        CN = CN + RS.RecordCount
        StatusBar1.Panels(1).Text = "ผลการค้นหา พบ " & CN & " รายการ"
        If RS.RecordCount > 0 Then Call AddtoListView(RS, 1)
        Set RS = Nothing
        DoEvents
    End If
    
    If cboDateType.ListIndex < 2 Then CheckRE
    MsgBox StatusBar1.Panels(1).Text, vbInformation + vbOKOnly, "Result"
    
    If Result.ListItems.Count > 0 Then
        cmdDoComExt_Base.Enabled = True
        cmdDoComExt_Pay.Enabled = True
        cmdDoComExt_Del.Enabled = True
        If LCase(CurrentUser.LoginName) = "admin" Then
            cmdDoComExt_Cancel.Enabled = True
        Else
            cmdDoComExt_Cancel.Enabled = False
        End If
        cmdPrintExt1.Enabled = True
        cmdPrintExt2.Enabled = True
    Else
        cmdDoComExt_Base.Enabled = False
        cmdDoComExt_Pay.Enabled = False
        cmdDoComExt_Del.Enabled = False
        cmdDoComExt_Cancel.Enabled = False
        cmdPrintExt1.Enabled = False
        cmdPrintExt2.Enabled = False
    End If
End Sub


Private Sub Form_Activate()
    Call ActivateForm
End Sub

Private Sub Form_Resize()
On Error Resume Next
    Call ActivateForm
End Sub

Private Sub ActivateForm()
    Frame1.Height = Me.Height - StatusBar1.Height - 350
    
    Frame2.Height = Me.Height - StatusBar1.Height - 350
    Frame2.Width = Me.Width - Frame2.Left - 150
    With SSTab1
        .Top = 120
        .Left = 0
        .Height = Frame2.Height - 120
        .Width = Frame2.Width
    End With
    With Result
        .Checkboxes = True
        .Top = 350
        .Left = 0
        .Height = Frame2.Height - 450
        .Width = Frame2.Width
    End With
    
    StatusBar1.Height = 435
    StatusBar1.Panels(1).Width = Me.Width
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDoComExt_Pay_Click()
    If Result.ListItems.Count <= 0 Then
        MsgBox "ยังไม่ได้เลือกรายการที่ตัดจ่ายค่าคอมฯ", vbInformation, "Warning"
        Exit Sub
    Else
        If MsgBox("รายการที่เลือกจะถูกบันทึกในฐานข้อมูล ค่าคอมฯ ที่จ่ายแล้ว กด [Yes] เพื่อยืนยันการบันทึกข้อมูล", vbExclamation + vbYesNo, "Excuting") = vbNo Then Exit Sub
        Call SaveData
    End If
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
    InitWaste
    dtFrom.DayValue = "01"
    SSTab1.Tab = 0
    StatusBar1.Panels(1).Text = ""

    With cboDateType
        .AddItem "วันที่ขน"
        .AddItem "วันที่ RE"
        .AddItem "วันที่ตัดจ่าย"
        .ListIndex = 0
    End With
    
    cmdDoComExt_Base.Enabled = False
    cmdDoComExt_Pay.Enabled = False
    cmdDoComExt_Del.Enabled = False
    cmdDoComExt_Cancel.Enabled = False
    cmdPrintExt1.Enabled = False
    cmdPrintExt2.Enabled = True
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
    DBConnExc "Delete tbPrintOrd01 Where SessionID='" & sessionID_ & "'"
End Sub

Private Sub MNU_DeSelect_All_Click()
Dim i As Integer
    If lvIndex = 0 Then
        For i = 1 To Result.ListItems.Count
            Result.ListItems(i).Checked = False
        Next
    End If
End Sub

Private Sub MNU_NotApprove_Click()
Dim i, j As Integer, CN As Integer
Dim tSql As String
    If lvIndex = 0 Then
        CN = 0
        If MsgBox("รายการที่ถูกเลือกจะถูกกำหนดสถานะเป็นไม่อนุมัติจ่าย", vbInformation + vbOKCancel) = vbCancel Then Exit Sub
        For i = 1 To Result.ListItems.Count
            If Result.ListItems(i).Checked = True Then
                If Result.ListItems(i).SubItems(37) = "2" Then
                    tSql = "UPDATE tbCom_Commission_External SET Status = 9" & vbCrLf
                    tSql = tSql & "WHERE MenifestID = '" & Result.ListItems(i).SubItems(6) & "' AND WasteDataID = '" & Result.ListItems(i).SubItems(8) & "'"
                    DBConnExc tSql
                    
                    Result.ListItems(i).SubItems(37) = "9"
                    For j = 1 To Result.ColumnHeaders.Count - 1
                        Result.ListItems.item(i).ListSubItems(j).ForeColor = Shape1(4).FillColor
                    Next
                    Result.Refresh
                    CN = CN + 1
                Else
                    Result.ListItems(i).Checked = False
                End If
            End If
        Next
        MsgBox "ไม่อนุมัติการจ่ายค่าคอมฯ นอก " & CN & " รายการ", vbInformation
    End If
End Sub

Private Sub MNU_SetApprove_Click()
Dim i, j As Integer, CN As Integer
Dim tSql As String
    If lvIndex = 0 Then
        CN = 0
        If MsgBox("รายการที่ถูกเลือกจะถูกกำหนดสถานะเป็นอนุมัติจ่าย", vbInformation + vbOKCancel) = vbCancel Then Exit Sub
        For i = 1 To Result.ListItems.Count
            If Result.ListItems(i).Checked = True Then
                If Result.ListItems(i).SubItems(37) = "9" Then
                    tSql = "UPDATE tbCom_Commission_External SET Status = 2" & vbCrLf
                    tSql = tSql & "WHERE MenifestID = '" & Result.ListItems(i).SubItems(6) & "' AND WasteDataID = '" & Result.ListItems(i).SubItems(8) & "'"
                    DBConnExc tSql
                    
                    Result.ListItems(i).SubItems(37) = "2"
                    For j = 1 To Result.ColumnHeaders.Count - 1
                        Result.ListItems.item(i).ListSubItems(j).ForeColor = Shape1(0).FillColor
                    Next
                    Result.Refresh
                    CN = CN + 1
                Else
                    Result.ListItems(i).Checked = False
                End If
            End If
        Next
        MsgBox "เปลี่ยนสถานะเป็นอนุมัติการจ่ายค่าคอมฯ นอก " & CN & " รายการ", vbInformation
    End If
End Sub

Private Sub MNU_Select_All_Click()
Dim i As Integer
    If lvIndex = 0 Then
        For i = 1 To Result.ListItems.Count
            If Result.ListItems(i).SubItems(31) = "Y" Then
                Result.ListItems(i).Checked = False
            Else
                If CDbl(Result.ListItems.item(i).SubItems(17)) = 0 Then
                    Result.ListItems(i).Checked = False
                Else
                    Result.ListItems(i).Checked = True
                End If
            End If
        Next
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim tSql As String
    If tmp_CommRE <> "" Then
        tSql = "IF  EXISTS (SELECT * FROM sys.tables WHERE object_id = OBJECT_ID(N'[dbo].[" & tmp_CommRE & "]'))" & vbCrLf
        tSql = tSql & "DROP TABLE [dbo].[" & tmp_CommRE & "]"
        DBConn.Execute tSql
    End If
End Sub

Private Sub Result_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Result.SortKey = ColumnHeader.Index - 1
    If Result.SortOrder = lvwAscending Then
        Result.SortOrder = lvwDescending
    Else
        Result.SortOrder = lvwAscending
    End If
    Result.Sorted = True
End Sub

Private Sub SettingResult(ByVal Index As Long, ByVal SKey As String, ByVal SText As String)
Dim i, j As Integer
    For i = 1 To Result.ListItems.Count
        If SKey = Result.ListItems.item(i).Key Then
            Result.ListItems.item(i).ForeColor = vbBlack
            For j = 1 To Result.ListItems.item(i).ListSubItems.Count - 1
                Result.ListItems.item(i).ListSubItems(j).ForeColor = vbBlack
            Next
            Exit For
        End If
    Next
End Sub

Private Sub Result_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Dim Pt As POINTAPI
    Dim ret As BTN_STYLE

    If Button = vbRightButton Then
        hMenu = CreatePopupMenu()
        AppendMenu hMenu, MF_STRING, 1, "เลือกทั้งหมด"
        AppendMenu hMenu, MF_SEPARATOR, -1, "SEP1"
        AppendMenu hMenu, MF_STRING, 2, "ไม่เลือกทั้งหมด"
        AppendMenu hMenu, MF_SEPARATOR, -1, "SEP2"
        AppendMenu hMenu, MF_STRING, 3, "ไม่อนุมัติจ่าย"
        If LCase(CurrentUser.LoginName) = "admin" Then AppendMenu hMenu, MF_STRING, 4, "อนุมัติจ่าย"
        GetCursorPos Pt
        ret = TrackPopupMenuEx(hMenu, TPM_LEFTALIGN Or TPM_RETURNCMD, _
                                Pt.x, Pt.Y, Me.hwnd, ByVal 0&)
        DestroyMenu hMenu

        If ret = 1 Then
            MNU_Select_All_Click
        ElseIf ret = 2 Then
            MNU_DeSelect_All_Click
        ElseIf ret = 3 Then
            MNU_NotApprove_Click
        ElseIf ret = 4 Then
            MNU_SetApprove_Click
        End If
    End If
    
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    If SSTab1.Tab = 2 Then
        Label1(5).Caption = "วันที่ตัดจ่าย"
    Else
        Label1(5).Caption = "วันที่รับชำระ"
    End If
End Sub

Private Sub txtCustomer_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then txtCustomer_Validate False
End Sub

Public Sub ValidateCustomer()
    txtCustomer_Validate False
End Sub

Private Sub txtCustomer_Validate(Cancel As Boolean)
    If txtCustomer.Text = "" Then Exit Sub
    If txtCustomer.Text = "*" Then Exit Sub
    If curCustomer_ Is Nothing Then Set curCustomer_ = New BWGCustomer
    txtCustomer.Tag = ""
    If Trim(txtCustomer.Text) = "" And Trim(curCustomer_.CustomerName) <> Trim(txtCustomer.Text) Then
        Set curCustomer_ = Nothing
        Set curCustomer_ = New BWGCustomer
    Else
        If Trim(curCustomer_.CustomerName) <> Trim(txtCustomer.Text) Then
            Set curCustomer_ = Nothing
            Set curCustomer_ = New BWGCustomer
            Load frCustomerSearch
            With frCustomerSearch
                .getCriteria (txtCustomer.Text)
                .setCallerForm Me
                .Show 1
            End With
            If curCustomer_.ID = "" Then txtCustomer.Text = ""
        End If
    End If
End Sub

Private Sub txtSale_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then txtSale_Validate False
End Sub

Private Sub txtSale_Validate(Cancel As Boolean)
    If txtSale.Text = "" Then Exit Sub
    If txtSale.Text = "*" Then Exit Sub
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

Private Sub InitWaste()
Dim x%
Dim strWaste As String
Dim tWaste() As String
    strWaste = "ไม่ระบุ,Haz,Non,HBL,NBL"
    tWaste = Split(strWaste, ",")
    cboWaste.Clear
    For x = 0 To 4
        cboWaste.AddItem tWaste(x)
    Next
    If cboWaste.ListCount > 0 Then
        cboWaste.ListIndex = 0
    End If
End Sub

Public Sub SetDataCustomer(selCustomer As BWGCustomer, Optional isAssigned As Boolean = True)
Dim tmpRec As New ADODB.Recordset
Dim x%
    Set curCustomer_ = Nothing
    Set curCustomer_ = selCustomer
    If curCustomer_ Is Nothing Then Exit Sub
    txtCustomer.Tag = ""
    With curCustomer_
        txtCustomer.Text = .CustomerName
        txtCustomer.Tag = .ID
    End With
End Sub

Private Function SqlQuery(Optional ByVal UseDate As Integer = 1) As String
Dim tmpSql As String, iCount As Long
Dim CurDate As String
Dim FirstDate As String
    CurDate = Format(Date, "dd/mm/yyyy")
    FirstDate = Right(CurDate, 4) & "/01/01"
    
    If UseDate = 0 Then
        tmpSql = tmpSql & "WHERE WorkDate >= '" & dtFrom.ValueYMD & "' And WorkDate <='" & dtTo.ValueYMD & "'" & vbCrLf
    ElseIf UseDate = 1 Then
        tmpSql = tmpSql & "WHERE REDate >= '" & dtFrom.ValueYMD & "' And REDate <='" & dtTo.ValueYMD & "'" & vbCrLf
    ElseIf UseDate = 2 Then
        tmpSql = tmpSql & "WHERE BaseDate >= '" & dtFrom.ValueYMD & "' AND BaseDate <= '" & dtTo.ValueYMD & "'" & vbCrLf
    End If
    
    tmpSql = tmpSql & "AND CompanyID = '" & CurrentUser.WorkCompanies(cboCompany.ListIndex + 1).ID & "'" & vbCrLf
    
    If cboZone.ListIndex > 0 Then
        tmpSql = tmpSql & "AND SaleZoneID='" & colZone_(cboZone.ListIndex).ID & "'" & vbCrLf
    End If
    If Trim(txtSale.Text) <> "" Then
        If Trim(txtSale.Text) <> "*" Then
            tmpSql = tmpSql & "AND UnderSaleStaffID='" & curStaff_.ID & "'" & vbCrLf
        End If
    End If
    If cboWaste.ListIndex > 0 Then
        tmpSql = tmpSql & "AND NonOrHz='" & Trim(cboWaste.Text) & "'" & vbCrLf
    End If
    
    If Trim(txtCustomer.Text) <> "" Then
        If Trim(txtCustomer.Text) <> "*" Then
            tmpSql = tmpSql & "AND CustomerID = '" & curCustomer_.ID & "'" & vbCrLf
        Else
            tmpSql = tmpSql & "AND ComRateAmt > 0"
        End If
    Else
        tmpSql = tmpSql & "AND ComRateAmt > 0"
    End If
    
    tmpSql = tmpSql & " "
    SqlQuery = tmpSql
End Function

Private Function SqlQueryList() As String
    Dim Index As Long
    Dim menifestID As String
    menifestID = ""
    For Index = 1 To Result.ListItems.Count
        If Result.ListItems.item(Index).Checked = True Then
            If menifestID = "" Then
                menifestID = "'" & Result.ListItems.item(Index).SubItems(6) & "'"
            Else
                menifestID = menifestID & ", '" & Result.ListItems.item(Index).SubItems(6) & "'"
            End If
        End If
    Next Index
    SqlQueryList = "AND MenifestID IN (" & menifestID & ")" & vbCrLf
End Function

Private Sub SaveData()
On Error GoTo ErrD
Dim RS As ADODB.Recordset
Dim tSql As String
Dim CurrentDay As String
Dim i, j As Long
    CurrentDay = Format(Date, "yyyy/mm/dd")
    Counts = 0
    If Result.ListItems.Count <= 0 Then Exit Sub
'    DBConn.BeginTrans
    For i = 1 To Result.ListItems.Count
        If CDbl(Result.ListItems.item(i).SubItems(17)) = 0 Then Result.ListItems.item(i).Checked = False
        If Result.ListItems.item(i).Checked = True Then
            If Trim(Result.ListItems.item(i).SubItems(4)) <> "" Then
                If Result.ListItems(i).SubItems(37) = "1" Then
                    StatusBar1.Panels(1).Text = "(Menifest No " & Result.ListItems.item(i).SubItems(7) & ") " & "ไม่ได้ทำการตั้งยอด ไม่สามารถตัดจ่ายได้"
                    Result.ListItems(i).Checked = False
                    DoEvents
                ElseIf Result.ListItems(i).SubItems(37) = "3" Then
                    StatusBar1.Panels(1).Text = "(Menifest No " & Result.ListItems.item(i).SubItems(7) & ") " & "ตัดจ่ายไปแล้ว ไม่สามารถตัดจ่ายได้อีก"
                    Result.ListItems(i).Checked = False
                    DoEvents
                ElseIf Result.ListItems(i).SubItems(37) = "9" Then
                    StatusBar1.Panels(1).Text = "(Menifest No " & Result.ListItems.item(i).SubItems(7) & ") " & "ไม่อนุมัติจ่าย ไม่สามารถตัดจ่ายได้อีก"
                    Result.ListItems(i).Checked = False
                    DoEvents
                ElseIf Trim(Result.ListItems(i).SubItems(10)) = "" Then
                    StatusBar1.Panels(1).Text = "(Menifest No " & Result.ListItems.item(i).SubItems(7) & ") " & "ยังไม่ได้เก็บเงิน ไม่สามารถตัดจ่ายได้"
                    Result.ListItems(i).Checked = False
                    DoEvents
                Else
                    Counts = Counts + 1
                    'Delete Old Transection
'                    tSql = "DELETE FROM tbCom_Commission_External" & vbCrLf
'                    tSql = tSql & "WHERE MenifestID = '" & Result.ListItems(i).SubItems(6) & "' AND WasteDataID = '" & Result.ListItems(i).SubItems(8) & "'"
'                    DBConnExc tSql
                    
                    tSql = "INSERT INTO tbCom_ExtCom_Paid(SDName , UnderSaleStaffID, StaffFName, CustomerID, CompanyName, "
                    tSql = tSql & "WorkDate, MenifestID, DocNo, WasteDataID, WasteName, ReNo, REDate, Weight, TreatmentRate, "
                    tSql = tSql & "TreatMentUnitID, InCollectFee, TransportFee, isPriceIncTrans, TripTransFee_Inc, ComRateAmt, RateType, CommValue, "
                    tSql = tSql & "VatRate, VatValue, NetComm, "
                    tSql = tSql & "NonOrHz, SaleZoneID, CompanyID, "
                    tSql = tSql & " StaffLName,TimeTableNo,TreatmentCharge, WasteType, ZoneName, isCanceled, SaleDistrictID, ComRecName, ComRatePC, RateTypePC, DocCustWeight, PaidDate,issuedate,billno)" & vbCrLf
                    tSql = tSql & "Values('"
                    tSql = tSql & Result.ListItems.item(i) & "', '" ' SDName
                    tSql = tSql & Result.ListItems.item(i).SubItems(1) & "', '" ' SaleSaleID
                    tSql = tSql & Result.ListItems.item(i).SubItems(2) & "', '" ' StaffFName
                    tSql = tSql & Result.ListItems.item(i).SubItems(3) & "', '" ' CustomerID
                    tSql = tSql & Result.ListItems.item(i).SubItems(4) & "', '" ' CustomerName
                    tSql = tSql & Format(Result.ListItems.item(i).SubItems(5), "yyyy/mm/dd") & "', '" ' WorkDate
                    tSql = tSql & Result.ListItems.item(i).SubItems(6) & "', '" ' MenifestID
                    tSql = tSql & Result.ListItems.item(i).SubItems(7) & "', '" ' DocNo
                    tSql = tSql & Result.ListItems.item(i).SubItems(8) & "', '" ' WasteDataID
                    tSql = tSql & Result.ListItems.item(i).SubItems(9) & "', '" ' WasteName
                    tSql = tSql & Result.ListItems.item(i).SubItems(10) & "', '" ' ReNo
                    tSql = tSql & Format(Result.ListItems.item(i).SubItems(11), "yyyy/mm/dd") & "', " ' ReDate
                    tSql = tSql & CDbl(Result.ListItems.item(i).SubItems(12)) & ", " ' Weight
                    tSql = tSql & CDbl(Result.ListItems.item(i).SubItems(13)) & ", '" ' TreatmentRate
                    tSql = tSql & Result.ListItems.item(i).SubItems(14) & "', " ' TreatmentUnit
                    tSql = tSql & CDbl(Result.ListItems.item(i).SubItems(15)) & ", " ' TreatmentCharg
                    tSql = tSql & "0" & ", " ' TransportFee ใส่ค่าขนส่ง กรณีแยกค่าขนส่งเป็น 0
                    If InStr(Result.ListItems.item(i).SubItems(16), "รวมค่าขนส่ง") > 0 Then
                        tSql = tSql & "'Y', "
                    Else
                        tSql = tSql & "'N', "
                    End If
                    tSql = tSql & CDbl(Trim(Replace(Result.ListItems.item(i).SubItems(16), "  (รวมค่าขนส่ง)", ""))) & ", " ' TripTransFee_Inc
                    tSql = tSql & CDbl(Result.ListItems.item(i).SubItems(17)) & ", '" ' ComRate
                    tSql = tSql & Result.ListItems.item(i).SubItems(18) & "', " ' RateType
                    tSql = tSql & CDbl(Result.ListItems.item(i).SubItems(19)) & ", " ' ComValue
                    tSql = tSql & CDbl(Result.ListItems.item(i).SubItems(20)) & ", " 'อัตราภาษี
                    tSql = tSql & CDbl(Result.ListItems.item(i).SubItems(21)) & ", " 'ภาษี
                    tSql = tSql & CDbl(Result.ListItems.item(i).SubItems(22)) & ", '" 'ค่าคอมสุทธิ
                    tSql = tSql & Result.ListItems.item(i).SubItems(23) & "', '" ' WasteType
                    tSql = tSql & Result.ListItems.item(i).SubItems(24) & "', '" ' ZoneID
                    tSql = tSql & Result.ListItems.item(i).SubItems(25) & "', '" ' CompanyID
                    tSql = tSql & Result.ListItems.item(i).SubItems(26) & "', '" ' StaffLName
                    tSql = tSql & Result.ListItems.item(i).SubItems(27) & "', " ' TimeTableNo
                    tSql = tSql & CDbl(Result.ListItems.item(i).SubItems(28)) & ", " ' TreatmentCharge
                    tSql = tSql & CDbl(Result.ListItems.item(i).SubItems(29)) & ", '" ' WasteType
                    tSql = tSql & Result.ListItems.item(i).SubItems(30) & "', '" ' ZoneName
                    tSql = tSql & Result.ListItems.item(i).SubItems(31) & "', '" ' IsCancel
                    tSql = tSql & Result.ListItems.item(i).SubItems(32) & "', '" ' SaleDistict
                    tSql = tSql & Result.ListItems.item(i).SubItems(33) & "', " ' ComRecName
                    tSql = tSql & CDbl(Result.ListItems.item(i).SubItems(34)) & ", '" ' ComRatePC
                    tSql = tSql & Result.ListItems.item(i).SubItems(35) & "', '" ' RateTypePC
                    tSql = tSql & Result.ListItems.item(i).SubItems(36) & "', '" ' DocCustweight
                    tSql = tSql & dtTo.ValueYMD & "','"
                    tSql = tSql & Result.ListItems.item(i).SubItems(38) & "', '" ' Issuedate
                    tSql = tSql & Result.ListItems.item(i).SubItems(39) & "' " ' BillNo
                    tSql = tSql & ")"
                    DBConnExc tSql
                    
                    tSql = "UPDATE tbCom_Commission_External SET RENo = '" & Result.ListItems(i).SubItems(10) & "' , REDate = '" & Format(Result.ListItems.item(i).SubItems(11), "yyyy/mm/dd") & "' , Status = 3" & vbCrLf
                    tSql = tSql & "WHERE MenifestID = '" & Result.ListItems(i).SubItems(6) & "' AND WasteDataID = '" & Result.ListItems(i).SubItems(8) & "'"
                    DBConnExc tSql
                    
                    Result.ListItems(i).SubItems(37) = "3"
                    For j = 1 To Result.ColumnHeaders.Count - 1
                        Result.ListItems.item(i).ListSubItems(j).ForeColor = vbGreen
                    Next
                    Result.Refresh
                End If
            End If
        End If
    Next
'    DBConn.CommitTrans
    MsgBox "บันทึกการตัดจ่ายค่าคอมฯ เสร็จแล้ว " & Counts & " รายการ", vbInformation, "Completed"
    StatusBar1.Panels(1).Text = "บันทึกการตัดจ่ายแล้ว " & Counts & " รายการ"
    Exit Sub
ErrD:
    MsgBox Err.Description
    StatusBar1.Panels(1).Text = Err.Description
'    DBConn.RollbackTrans
    Exit Sub
End Sub

Private Function LoadCommPaid() As Long
On Error GoTo ErrD
Dim RS As ADODB.Recordset
Dim tSql As String
Dim tList As ListItem
Dim InCollectFee As Double
Dim CommValue As Double
Dim RateType As String
Dim VatValue As Double
    tSql = SqlQuery(cboDateType.ListIndex)
    Set RS = New ADODB.Recordset
    Set RS = GetRS(tSql, adOpenForwardOnly, adLockReadOnly)
    LoadCommPaid = RS.RecordCount
    Call AddtoListView(RS, 3)
    Set RS = Nothing
    Exit Function
ErrD:
    MsgBox Err.Description
    LoadCommPaid = 0
    Set RS = Nothing
    Exit Function
End Function

Private Sub AddtoListView(ByRef RS As ADODB.Recordset, ByVal Index As Integer)
On Error Resume Next
Dim RSPay As ADODB.Recordset
Dim tSql As String
Dim InCollectFee As Double, TransportFee As Double
Dim CommValue As Double
Dim RateType As String, TransportType As String
Dim i As Integer

Dim OldCustID As String
Dim CurCustID As String
Dim Cust_Com As Double
Dim VatValue As Double

Dim LastKey As String
Dim LastRecord As Boolean
Dim MinIndex As Integer
Dim MaxIndex As Integer

'ตัวแปรสำหรับเปรียบเทียบการขนแบบเที่ยว
Dim SameTimeTableNo As String, FirstWaste As Boolean, WasteCount As Integer

    If RS.RecordCount = 0 Then Exit Sub
    
    OldCustID = ""
    Cust_Com = 0
    MinIndex = 0: MaxIndex = 0
    
    SameTimeTableNo = ""
    WasteCount = 0
    Do While Not RS.EOF
'        DoEvents
        CurCustID = RS!CustomerID
        If OldCustID = "" Then OldCustID = RS!CustomerID: MinIndex = 1

        If SameTimeTableNo = RS!TimeTableNo Then
            WasteCount = WasteCount + 1
            FirstWaste = False
        ElseIf SameTimeTableNo <> RS!TimeTableNo Then
            SameTimeTableNo = RS!TimeTableNo
            WasteCount = 1
            FirstWaste = True
        End If
        
        If Index = 1 Then
            Set tList = Result.ListItems.Add(, RS!UnderSaleStaffID & "_" & OldCustID & "_" & RS!RENo & "_" & RS!docNo & "_" & RS!wastedataID & "_" & IIf(IsNull(RS!Weight), 0, RS!Weight), RS!SDName)
        End If
        
        LastKey = RS!UnderSaleStaffID & "_" & OldCustID
        LastRecord = False
        MaxIndex = tList.Index
        If CurCustID <> OldCustID Then
            '******** ค้นหาค่าคอมเก่าภายในปีปัจจุบัน ***********
            MaxIndex = MaxIndex - 1
            Call CalcCom(OldCustID, Cust_Com, Index, LastRecord, MinIndex, MaxIndex, LastKey)
            OldCustID = CurCustID
            Cust_Com = 0
            MinIndex = tList.Index
            MaxIndex = tList.Index
        End If
        '************
            tList.SubItems(1) = RS!UnderSaleStaffID
            tList.SubItems(2) = IIf(IsNull(RS!StaffFName), "", RS!StaffFName)
            tList.SubItems(3) = RS!CustomerID
            tList.SubItems(4) = RS!CompanyName
            tList.SubItems(5) = Format(RS!WorkDate, "dd/mm/yyyy")
            tList.SubItems(6) = RS!menifestID
            tList.SubItems(7) = RS!docNo
            tList.SubItems(8) = RS!wastedataID
            tList.SubItems(9) = RS!WasteName
            tList.SubItems(10) = RS!RENo
            tList.SubItems(11) = Format(RS!REDate, "dd/mm/yyyy")
            tList.SubItems(12) = IIf(IsNull(RS!Weight), 0#, Format(RS!Weight, "0.000"))
            tList.SubItems(13) = IIf(IsNull(RS!TreatmentRate), 0, RS!TreatmentRate)
            tList.SubItems(14) = IIf(IsNull(RS!TreatmentUnitID), "", RS!TreatmentUnitID)
            If RS!TreatmentUnitID = "บาท/เที่ยว" Then
                If WasteCount = 1 Then
                    InCollectFee = IIf(IsNull(RS!TreatmentRate), 0, RS!TreatmentRate)
                Else
                    InCollectFee = 0
                End If
            Else
                InCollectFee = IIf(IsNull(RS!TreatmentRate), 0, RS!TreatmentRate) * IIf(IsNull(RS!Weight), 0, RS!Weight)
            End If
            
            If IIf(IsNull(RS!isPriceIncTrans), "N", RS!isPriceIncTrans) = "Y" Then
                If WasteCount = 1 Then
                    InCollectFee = InCollectFee - IIf(IsNull(RS!TripTransFee_Inc), 0, RS!TripTransFee_Inc)
                    TransportFee = IIf(IsNull(RS!TripTransFee_Inc), 0, RS!TripTransFee_Inc)
                    TransportType = "  (รวมค่าขนส่ง)"
                Else
                    TransportFee = 0
                    TransportType = ""
'                    TransportFee = IIf(IsNull(RS!TripTransFee_Inc), 0, RS!TripTransFee_Inc)
'                     TransportType = "  (รวมค่าขนส่ง)"

                End If
            Else
                TransportFee = 0
                TransportType = ""
            End If
            
            tList.SubItems(15) = Format(InCollectFee, "#,##0.000")
            tList.SubItems(16) = Format(TransportFee, "#,##0") & TransportType
            tList.SubItems(17) = Format(RS!ComRateAmt, "#,##0")
            
            RateType = IIf(IsNull(RS!RateType), "", RS!RateType)
            tList.SubItems(18) = RateType
            
            If Trim(RateType) = "%" Then
                CommValue = InCollectFee * RS!ComRateAmt / 100
            Else
                If Trim(RateType) = "บาท/เที่ยว" Then
                    If FirstWaste = True Then
                        CommValue = RS!ComRateAmt
                    Else
                        CommValue = 0
                    End If
                Else
                    CommValue = IIf(IsNull(RS!ComRateAmt), 0, RS!ComRateAmt) * IIf(IsNull(RS!Weight), 0, RS!Weight)
                End If
            End If
            
            If RS!IsCanceled = "Y" Then
                tList.SubItems(19) = "0.000"
            Else
                tList.SubItems(19) = Format(CommValue, "#,##0.000")
                Cust_Com = Cust_Com + CommValue
            End If
'            tList.SubItems(20) = "5" 'อัตราภาษี
'            VatValue = CommValue * CDbl(tList.SubItems(20)) / 100
'            tList.SubItems(21) = Format(VatValue, "#,#00.00")
'            tList.SubItems(22) = Format(CommValue - VatValue, "#,#00.00")
            tList.SubItems(23) = IIf(IsNull(RS!NonOrHz), "", RS!NonOrHz)
            tList.SubItems(24) = IIf(IsNull(RS!SaleZoneID), "", RS!SaleZoneID)
            tList.SubItems(25) = IIf(IsNull(RS!CompanyID), "", RS!CompanyID)
            tList.SubItems(26) = IIf(IsNull(RS!StaffLName), "", RS!StaffLName)
            tList.SubItems(27) = IIf(IsNull(RS!TimeTableNo), "", RS!TimeTableNo)
            tList.SubItems(28) = IIf(IsNull(RS!TreatmentCharge), 0, RS!TreatmentCharge)
            tList.SubItems(29) = IIf(IsNull(RS!WasteType), "", RS!WasteType)
            tList.SubItems(30) = IIf(IsNull(RS!ZoneName), "", RS!ZoneName)
            tList.SubItems(31) = IIf(IsNull(RS!IsCanceled), "", RS!IsCanceled)
            tList.SubItems(32) = IIf(IsNull(RS!SaleDistrictID), "", RS!SaleDistrictID)
            tList.SubItems(33) = IIf(IsNull(RS!ComRecName), "", RS!ComRecName)
            tList.SubItems(34) = IIf(IsNull(RS!ComRatePC), 0, RS!ComRatePC)
            tList.SubItems(35) = IIf(IsNull(RS!RateTypePC), "", RS!RateTypePC)
            tList.SubItems(36) = IIf(IsNull(RS!DocCustWeight), "", RS!DocCustWeight)
            tList.SubItems(37) = IIf(IsNull(RS!Status), "", RS!Status)
            If Index = 3 Then
                tList.SubItems(37) = RS!RunPaidNo
            End If
            'Check for paid
            If Index = 1 Then
                Dim vColor As ColorConstants
                If tList.SubItems(37) = "1" Then
                    vColor = vbBlack
                ElseIf tList.SubItems(37) = "2" Then
                    vColor = vbBlue
                ElseIf tList.SubItems(37) = "3" Then
                    vColor = vbGreen
                ElseIf tList.SubItems(37) = "8" Then
                    vColor = vbBlack
                ElseIf tList.SubItems(37) = "9" Then
                    vColor = Shape1(4).FillColor
                Else
                    vColor = vbBlack
                End If
                If tList.SubItems(31) = "Y" Then
                    vColor = vbRed
                End If
                
                For i = 1 To Result.ColumnHeaders.Count - 1
                    tList.ListSubItems(i).ForeColor = vColor
                Next
            
'                tSql = SqlQuery("tbCom_ExtCom_Paid", 0)
'                Set RSPay = New ADODB.Recordset
'                RSPay.Open tSql, DBConn, adOpenKeyset, adLockOptimistic
'                Do While Not RSPay.EOF
'                    If RSPay!CompanyID & "_" & RSPay!UnderSaleStaffID & "_" & RSPay!CustomerID & "_" & RSPay!MenifestID & "_" & RSPay!WasteDataID = RS!CompanyID & "_" & RS!UnderSaleStaffID & "_" & RS!CustomerID & "_" & RS!MenifestID & "_" & RS!WasteDataID Then
'                        tList.ForeColor = vbBlue
'                        For i = 1 To Result.ColumnHeaders.Count - 1
'                            tList.ListSubItems(i).ForeColor = vbBlue
'                        Next
'                        Exit Do
'                    End If
'                    RSPay.MoveNext
'                Loop
            End If
            RS.MoveNext
    Loop
    Result.Refresh
    LastRecord = True
    Call CalcCom(OldCustID, Cust_Com, Index, LastRecord, MinIndex, MaxIndex, LastKey)
    
    Set RS = Nothing
    Exit Sub
ErrD:
    MsgBox Err.Description
    Set RS = Nothing
    Exit Sub
End Sub

Private Sub CalcCom(ByVal CustID As String, ByVal Comm As Double, ByVal Index As Integer, ByVal LastRecord As Boolean, ByVal MinIndex As Integer, ByVal MaxIndex As Integer, Optional ByVal Key As String = "")
On Error GoTo ErrD
Dim Cust_TotalCom As Double
Dim ComValue As Double
Dim ComRate As Double
Dim VatValue As Double

Dim RS As ADODB.Recordset
Dim tSql As String
Dim CurDate As String
Dim CurYear As String
Dim FirstDate As String
Dim LastDate As String
Dim i As Integer

    CurDate = dtFrom.ValueDMY
    CurYear = dtFrom.YearValue
    FirstDate = CurYear & "/01/01"
    LastDate = CurYear & "/12/31"

'    If LastRecord = True Then
'        If Key <> "" Then
'            If Index = 1 Then
'                Set tList = Result.ListItems.Add(, Key, "")
'            ElseIf Index = 2 Then
'                Set tList = Result2.ListItems.Add(, Key, "")
'            ElseIf Index = 3 Then
'                Set tList = Result3.ListItems.Add(, Key, "")
'            End If
'        End If
'    End If
    tSql = "SELECT ISNULL(SUM(ISNULL(CommValue, 0)), 0) AS SumCommValue FROM tbCom_ExtCom_Paid" & vbCrLf
    tSql = tSql & "WHERE CompanyID = '" & CurrentUser.WorkCompanies(cboCompany.ListIndex + 1).ID & "'" & vbCrLf
    tSql = tSql & "AND CustomerID = '" & CustID & "'" & vbCrLf
    tSql = tSql & "AND PaidDate >='" & FirstDate & "'"
    Set RS = New ADODB.Recordset
    Set RS = GetRS(tSql, adOpenForwardOnly, adLockReadOnly)
    If Not RS.EOF Then
        Cust_TotalCom = IIf(IsNull(RS!SumCommValue), 0, RS!SumCommValue)
    Else
        Cust_TotalCom = 0
    End If
    Set RS = Nothing
'            tList.Text = ""
'            tList.SubItems(19) = Format(Comm, "#,#00.00")
            Cust_TotalCom = Cust_TotalCom + Comm
'            If Cust_TotalCom > 150000 Then MsgBox "More 150,000"
            ComRate = GetComRate(FirstDate, LastDate, Cust_TotalCom)
'            If Cust_TotalCom <= 100000 Then
'                ComRate = 5
'            ElseIf Cust_TotalCom > 100000 And Cust_TotalCom <= 500000 Then
'                ComRate = 10
'            ElseIf Cust_TotalCom > 500000 And Cust_TotalCom <= 1000000 Then
'                ComRate = 20
'            ElseIf Cust_TotalCom > 1000000 And Cust_TotalCom <= 4000000 Then
'                ComRate = 30
'            ElseIf Cust_TotalCom > 4000000 Then
'                ComRate = 40
'            End If
            
            For i = MinIndex To MaxIndex
                If Index = 1 Then
                    ComValue = CDbl(Result.ListItems(i).SubItems(19))
                    VatValue = ComValue * ComRate / 100
                                    
                    Result.ListItems(i).SubItems(20) = ComRate
                    Result.ListItems(i).SubItems(21) = Format(VatValue, "#,#00.000")
                    Result.ListItems(i).SubItems(22) = Format(ComValue - VatValue, "#,#00.000")
                End If
'                VatValue = Comm * ComRate / 100
'                tList.SubItems(20) = ComRate
'                tList.SubItems(21) = Format(VatValue, "#,#00.00")
'                tList.SubItems(22) = Format(Comm - VatValue, "#,#00.00")
                
'                tList.ListSubItems(19).ForeColor = vbBlue
'                tList.ListSubItems(20).ForeColor = vbBlue
'                tList.ListSubItems(21).ForeColor = vbBlue
'                tList.ListSubItems(22).ForeColor = vbBlue
            Next
            Exit Sub
ErrD:
            MsgBox Err.Description
            Exit Sub
End Sub

Private Function GetComRate(ByVal SDate As String, ByVal EDate As String, ByVal SummaryValue As Double) As Double
Dim RS As New ADODB.Recordset
Dim tSql As String
    GetComRate = 0
    tSql = "SELECT     '" & SummaryValue & "' AS SummaryValue, MAX(CASE WHEN " & SummaryValue & " > SummaryValue THEN PercentRate ELSE 0 END)  AS PercentRate" & vbCrLf
    tSql = tSql & "FROM dbo.Com_tbComRateExternal" & vbCrLf
    tSql = tSql & "WHERE StartDate = '" & SDate & "' AND EndDate = '" & EDate & "'"
    Set RS = New ADODB.Recordset
    Set RS = GetRS(tSql, adOpenForwardOnly, adLockReadOnly)
    If Not RS.EOF Then
        GetComRate = Trim("0" & RS!PercentRate)
    Else
        GetComRate = 0
    End If
End Function

Private Sub PrintMainExt1(printType As Integer)
Dim tmpSql As String, iCount As Long

    tmpSql = "SELECT * FROM tbCom_Commission_External" & vbCrLf
    tmpSql = tmpSql & SqlQuery(cboDateType.ListIndex)
    tmpSql = tmpSql & SqlQueryList
    tmpSql = tmpSql & "AND (Status = '2' OR Status = '3')" & vbCrLf
    With rpt1
        .Reset
        .WindowTitle = "รายงานการตั้งค่าคอมฯ ภายนอก"
        .WindowShowCloseBtn = True
        .WindowShowExportBtn = True
        .WindowShowGroupTree = True
        .WindowShowPrintBtn = True
        .WindowShowPrintSetupBtn = True
        .WindowShowSearchBtn = True
        .WindowState = crptMaximized
        .LogonInfo(0) = "DSN = " & DSN_NAME & ";UID = " & strUid & ";PWD = " & strPwd
        .ReportFileName = App.Path & "\Reports\ComEx.rpt"
        tmpSql = tmpSql & "ORDER BY  WorkDate, TimeTableNo, TripTransFee_Inc Desc"
        .Formulas(0) = "Company='บริษัท " & Trim(cboCompany.Text) & "'"
        .Formulas(1) = "WorkDate='" & "วันที่ให้บริการ  " & dtFrom.ValueDMY & "  ถึงวันที่  " & dtTo.ValueDMY & "'"
        .Formulas(2) = "Sales='" & IIf(Trim(txtSale.Text) <> "", IIf(Trim(txtSale) <> "*", txtSale.Text, ""), "") & "'"
        .SqlQuery = tmpSql
        .Destination = printType
        .PrintReport
        If Trim(.LastErrorString) <> "" Then MsgBox .LastErrorString
    End With
End Sub

Private Sub PrintMainExt2(ByVal printType As Integer)
Dim Dates As String
Dim tSql As String, i As Integer
Dim ManifestID1 As String, WasteDataID1 As String
Dim ManifestID2 As String, WasteDataID2 As String

'    Dates = Format(Date, "yyyy/mm/dd")
'    tSql = "SELECT * FROM tbCom_Commission_External" & vbCrLf
'    tSql = tSql & SqlQuery(cboDateType.ListIndex)
'    tSql = tSql & "AND Status = '2' OR Status = '3'"
    
    ManifestID1 = ""
    WasteDataID1 = ""
    ManifestID2 = ""
    WasteDataID2 = ""
        
    For i = 1 To Result.ListItems.Count
        If Result.ListItems(i).SubItems(37) = "3" Then
            If ManifestID1 = "" Then
                ManifestID1 = "'" & Result.ListItems(i).SubItems(6) & "'"
                WasteDataID1 = "'" & Result.ListItems(i).SubItems(8) & "'"
            Else
                ManifestID1 = ManifestID1 & ", '" & Result.ListItems(i).SubItems(6) & "'"
                WasteDataID1 = WasteDataID1 & ", '" & Result.ListItems(i).SubItems(8) & "'"
            End If
        End If
    Next
    
    For i = 1 To Result.ListItems.Count
        If Result.ListItems(i).Checked = True Then
            If Result.ListItems(i).SubItems(37) = "3" Then
                If ManifestID2 = "" Then
                    ManifestID2 = "'" & Result.ListItems(i).SubItems(6) & "'"
                    WasteDataID2 = "'" & Result.ListItems(i).SubItems(8) & "'"
                Else
                    ManifestID2 = ManifestID2 & ", '" & Result.ListItems(i).SubItems(6) & "'"
                    WasteDataID2 = WasteDataID2 & ", '" & Result.ListItems(i).SubItems(8) & "'"
                End If
            End If
        End If
    Next
    
    tSql = "SELECT * FROM tbCom_ExtCom_Paid" & vbCrLf
    If cboDateType.ListIndex = 2 Then
        tSql = tSql & SqlQuery(cboDateType.ListIndex)
    ElseIf cboDateType.ListIndex = 0 Then
        tSql = tSql & "WHERE MenifestID IN (" & ManifestID2 & ")" & vbCrLf
        tSql = tSql & "AND WasteDataID IN (" & WasteDataID2 & ")" & vbCrLf
    Else
        tSql = tSql & "WHERE MenifestID IN (" & ManifestID2 & ")" & vbCrLf
        tSql = tSql & "AND WasteDataID IN (" & WasteDataID2 & ")" & vbCrLf
    End If
    
    With rpt1
        .Reset
        .WindowTitle = "รายงานการรับชำระค่าบริการ / การตัดจ่ายค่าคอมฯ ภายนอก"
        .WindowShowCloseBtn = True
        .WindowShowExportBtn = True
        .WindowShowGroupTree = True
        .WindowShowPrintBtn = True
        .WindowShowPrintSetupBtn = True
        .WindowShowSearchBtn = True
        .WindowState = crptMaximized
        .LogonInfo(0) = "DSN = " & DSN_NAME & ";UID = " & strUid & ";PWD = " & strPwd
        .ReportFileName = App.Path & "\Reports\ComEx2.rpt"
        .Formulas(0) = "Company='บริษัท " & Trim(cboCompany.Text) & "'"
        If cboDateType.ListIndex = 0 Then
            .Formulas(1) = "WorkDate='" & "วันที่ให้บริการ  " & dtFrom.ValueDMY & "  ถึงวันที่  " & dtTo.ValueDMY & "'"
        ElseIf cboDateType.ListIndex = 1 Then
            .Formulas(1) = "WorkDate='" & "วันที่รับชำระ  " & dtFrom.ValueDMY & "  ถึงวันที่  " & dtTo.ValueDMY & "'"
        ElseIf cboDateType.ListIndex = 2 Then
            .ReportFileName = App.Path & "\Reports\ComEx2_Paid.rpt"
        End If
        .Formulas(2) = "Sales='" & IIf(Trim(txtSale.Text) <> "", IIf(Trim(txtSale) <> "*", txtSale.Text, ""), "") & "'"
        .SqlQuery = tSql & "ORDER BY  WorkDate, TimeTableNo, TripTransFee_Inc Desc"
        .Destination = crptToWindow
        .PrintReport
        If Trim(.LastErrorString) <> "" Then MsgBox .LastErrorString
    End With
End Sub
