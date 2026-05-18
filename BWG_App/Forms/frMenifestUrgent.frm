VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMenifestUrgent 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "F063"
   ClientHeight    =   9135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11580
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9135
   ScaleWidth      =   11580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "เพิ่มรายการของเสีย"
      Height          =   795
      Left            =   1110
      Style           =   1  'Graphical
      TabIndex        =   171
      Top             =   8310
      Width           =   1605
   End
   Begin VB.TextBox txtMenifestNo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   9600
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   450
      Width           =   1935
   End
   Begin VB.ComboBox cboCompany 
      Height          =   315
      Left            =   6630
      Style           =   2  'Dropdown List
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   60
      Width           =   4905
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ปิดงาน"
      Height          =   795
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   75
      Top             =   8310
      Width           =   1005
   End
   Begin VB.TextBox txtDocNo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   6630
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   450
      Width           =   1935
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "เก็บข้อมูล"
      Height          =   795
      Left            =   6300
      Style           =   1  'Graphical
      TabIndex        =   49
      Top             =   8310
      Width           =   1005
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "เปิดข้อมูลเก่า"
      Height          =   795
      Left            =   7350
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   8310
      Width           =   1005
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "สร้างใหม่"
      Height          =   795
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   8310
      Width           =   1005
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "ปิด"
      Height          =   795
      Left            =   10500
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   8310
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ยกเลิก"
      Height          =   795
      Left            =   9450
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   8310
      Width           =   1005
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7455
      Left            =   30
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   795
      Width           =   11565
      _ExtentX        =   20399
      _ExtentY        =   13150
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "ส่วนของผู้ก่อกำเนิดของเสียอันตราย"
      TabPicture(0)   =   "frMenifestUrgent.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label2(2)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label2(4)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label2(5)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label2(8)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label2(10)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label2(11)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Image1(1)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label1(2)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label2(20)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label2(23)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label2(25)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label1(8)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label2(3)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label2(37)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label2(42)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label2(56)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label2(57)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Label2(68)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Label2(69)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "dtSign1"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "dtWorkDate"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "grItem"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "txtCustomer"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "txtCustTel"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "txtCustFax"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "txtSolidQty"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "txtLiquidQty"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "txtTreatment"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "txtSigneName1"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "txtCustER"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "Frame4"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "txtCollector"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "txtCollectorID"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "cboLiquidUnit"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "cboSolidUnit"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "cboSiteName"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "cboWasteType"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "cboGenerator"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "txtCollectorID2"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "txtCollector2"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "txtJobNo"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "cboWType"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).ControlCount=   44
      TabCaption(1)   =   "ส่วนของผู้ขนส่ง และผู้เก็บรวบรวมของเสียอันตราย"
      TabPicture(1)   =   "frMenifestUrgent.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label2(26)"
      Tab(1).Control(1)=   "Label2(27)"
      Tab(1).Control(2)=   "Label2(29)"
      Tab(1).Control(3)=   "Label2(30)"
      Tab(1).Control(4)=   "Label2(31)"
      Tab(1).Control(5)=   "Label2(28)"
      Tab(1).Control(6)=   "Image1(0)"
      Tab(1).Control(7)=   "Image1(2)"
      Tab(1).Control(8)=   "Label2(38)"
      Tab(1).Control(9)=   "Label2(39)"
      Tab(1).Control(10)=   "Label2(40)"
      Tab(1).Control(11)=   "Label2(41)"
      Tab(1).Control(12)=   "Label2(43)"
      Tab(1).Control(13)=   "Label1(0)"
      Tab(1).Control(14)=   "Label2(7)"
      Tab(1).Control(15)=   "Label1(3)"
      Tab(1).Control(16)=   "Label2(49)"
      Tab(1).Control(17)=   "Label2(48)"
      Tab(1).Control(18)=   "Label2(47)"
      Tab(1).Control(19)=   "Label2(46)"
      Tab(1).Control(20)=   "Label2(45)"
      Tab(1).Control(21)=   "Label2(44)"
      Tab(1).Control(22)=   "Label2(19)"
      Tab(1).Control(23)=   "Label2(18)"
      Tab(1).Control(24)=   "Label2(17)"
      Tab(1).Control(25)=   "Label2(16)"
      Tab(1).Control(26)=   "Label2(15)"
      Tab(1).Control(27)=   "Label2(9)"
      Tab(1).Control(28)=   "Label2(12)"
      Tab(1).Control(29)=   "dtSign3"
      Tab(1).Control(30)=   "dtSign2"
      Tab(1).Control(31)=   "txtTrans1B"
      Tab(1).Control(32)=   "txtTransID1B"
      Tab(1).Control(33)=   "txtTransER"
      Tab(1).Control(33).Enabled=   0   'False
      Tab(1).Control(34)=   "txtTransFax"
      Tab(1).Control(34).Enabled=   0   'False
      Tab(1).Control(35)=   "txtTransTel"
      Tab(1).Control(35).Enabled=   0   'False
      Tab(1).Control(36)=   "txtCarRegisID1"
      Tab(1).Control(36).Enabled=   0   'False
      Tab(1).Control(37)=   "txtFromProvince"
      Tab(1).Control(38)=   "txtToProvince"
      Tab(1).Control(39)=   "txtTransTime"
      Tab(1).Control(40)=   "txtSign1B"
      Tab(1).Control(41)=   "cboTransUnit"
      Tab(1).Control(42)=   "Frame1"
      Tab(1).Control(43)=   "txtDestName"
      Tab(1).Control(44)=   "cboTimeUnit"
      Tab(1).Control(45)=   "txtTreatTime"
      Tab(1).Control(45).Enabled=   0   'False
      Tab(1).Control(46)=   "txtDestTel"
      Tab(1).Control(46).Enabled=   0   'False
      Tab(1).Control(47)=   "txtDestFax"
      Tab(1).Control(47).Enabled=   0   'False
      Tab(1).Control(48)=   "txtDestER"
      Tab(1).Control(48).Enabled=   0   'False
      Tab(1).Control(49)=   "txtRealWeight"
      Tab(1).Control(50)=   "txtDestroyPlace"
      Tab(1).Control(51)=   "txtDestroyID"
      Tab(1).Control(52)=   "txtDestroyer"
      Tab(1).Control(53)=   "txtTumbol"
      Tab(1).Control(54)=   "txtAumphur"
      Tab(1).Control(54).Enabled=   0   'False
      Tab(1).Control(55)=   "txtTruckType"
      Tab(1).Control(56)=   "cboActualUnit"
      Tab(1).ControlCount=   57
      TabCaption(2)   =   "ข้อมูลเพิ่มเติมส่วนงานบัญชี"
      TabPicture(2)   =   "frMenifestUrgent.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label2(59)"
      Tab(2).Control(1)=   "Frame2(0)"
      Tab(2).Control(2)=   "Frame2(1)"
      Tab(2).Control(3)=   "Frame2(2)"
      Tab(2).Control(4)=   "txtAccStaff"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).ControlCount=   5
      Begin VB.ComboBox cboWType 
         Height          =   315
         Left            =   8400
         Style           =   2  'Dropdown List
         TabIndex        =   163
         Top             =   3750
         Width           =   2595
      End
      Begin VB.ComboBox cboActualUnit 
         Height          =   315
         ItemData        =   "frMenifestUrgent.frx":0054
         Left            =   -65640
         List            =   "frMenifestUrgent.frx":0061
         Style           =   2  'Dropdown List
         TabIndex        =   95
         Top             =   4020
         Width           =   1455
      End
      Begin VB.TextBox txtJobNo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C00000&
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   9390
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   540
         Width           =   1755
      End
      Begin VB.TextBox txtTruckType 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   -73080
         TabIndex        =   28
         Top             =   1200
         Width           =   5475
      End
      Begin VB.TextBox txtAccStaff 
         Appearance      =   0  'Flat
         BackColor       =   &H00C00000&
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   -66510
         TabIndex        =   155
         TabStop         =   0   'False
         Top             =   6870
         Width           =   2805
      End
      Begin VB.Frame Frame2 
         Caption         =   "ต้นทุนค่าใช้จ่ายสำหรับใบขนนี้"
         Height          =   1365
         Index           =   2
         Left            =   -74880
         TabIndex        =   141
         Top             =   5430
         Width           =   11205
         Begin VB.TextBox txtTotalOut 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C00000&
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   9540
            Locked          =   -1  'True
            TabIndex        =   168
            TabStop         =   0   'False
            Top             =   780
            Width           =   1005
         End
         Begin VB.TextBox txtOutWorker 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   6960
            TabIndex        =   157
            TabStop         =   0   'False
            Top             =   390
            Width           =   1005
         End
         Begin VB.TextBox txtOutEquip 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   9540
            TabIndex        =   152
            TabStop         =   0   'False
            Top             =   390
            Width           =   1005
         End
         Begin VB.TextBox txtWorkerCount 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   5430
            TabIndex        =   149
            TabStop         =   0   'False
            Top             =   390
            Width           =   465
         End
         Begin VB.TextBox txtOutTrans 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   1110
            TabIndex        =   146
            TabStop         =   0   'False
            Top             =   390
            Width           =   1005
         End
         Begin VB.TextBox txtOutOther 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   1110
            MaxLength       =   60
            TabIndex        =   143
            TabStop         =   0   'False
            Top             =   780
            Width           =   3705
         End
         Begin VB.TextBox txtOutOtherFee 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   4890
            TabIndex        =   142
            TabStop         =   0   'False
            Top             =   780
            Width           =   1005
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "รวมค่าใช้จ่ายทั้งหมด"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   73
            Left            =   8040
            TabIndex        =   170
            Top             =   840
            Width           =   1425
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "บาท"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   72
            Left            =   10620
            TabIndex        =   169
            Top             =   840
            Width           =   285
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ค่าแรง"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   67
            Left            =   6420
            TabIndex        =   159
            Top             =   450
            Width           =   465
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "บาท/คน"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   60
            Left            =   8040
            TabIndex        =   158
            Top             =   450
            Width           =   585
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ค่าอุปกรณ์"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   62
            Left            =   8730
            TabIndex        =   154
            Top             =   450
            Width           =   705
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "บาท"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   61
            Left            =   10620
            TabIndex        =   153
            Top             =   450
            Width           =   285
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "จำนวนคนงาน"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   64
            Left            =   4380
            TabIndex        =   151
            Top             =   450
            Width           =   990
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "คน"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   63
            Left            =   5970
            TabIndex        =   150
            Top             =   450
            Width           =   225
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ค่าขนส่ง"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   66
            Left            =   480
            TabIndex        =   148
            Top             =   450
            Width           =   555
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "บาท (ที่ต้องจ่ายให้ Vendor)"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   65
            Left            =   2190
            TabIndex        =   147
            Top             =   450
            Width           =   1845
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "อื่นๆ ระบุ"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   58
            Left            =   420
            TabIndex        =   145
            Top             =   840
            Width           =   615
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "บาท"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   55
            Left            =   6000
            TabIndex        =   144
            Top             =   840
            Width           =   285
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "ค่าบริการที่เรียกเก็บลูกค้า"
         Height          =   3855
         Index           =   1
         Left            =   -74880
         TabIndex        =   124
         Top             =   1470
         Width           =   11205
         Begin VB.TextBox txtTotalIn 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C00000&
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   9540
            Locked          =   -1  'True
            TabIndex        =   165
            TabStop         =   0   'False
            Top             =   3360
            Width           =   1005
         End
         Begin WasteManagment.ctlGrid grCustRate 
            Height          =   2595
            Left            =   90
            TabIndex        =   160
            Top             =   240
            Width           =   10995
            _ExtentX        =   19394
            _ExtentY        =   4577
         End
         Begin VB.TextBox txtOtherFee 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   4890
            TabIndex        =   139
            TabStop         =   0   'False
            Top             =   3360
            Width           =   1005
         End
         Begin VB.TextBox txtOtherDesc 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   1110
            MaxLength       =   60
            TabIndex        =   137
            TabStop         =   0   'False
            Top             =   3360
            Width           =   3735
         End
         Begin VB.TextBox txtInEquip 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   9540
            TabIndex        =   134
            TabStop         =   0   'False
            Top             =   2970
            Width           =   1005
         End
         Begin VB.TextBox txtInWorker 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   6960
            TabIndex        =   131
            TabStop         =   0   'False
            Top             =   2970
            Width           =   1005
         End
         Begin VB.TextBox txtIntrans 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   3840
            TabIndex        =   128
            TabStop         =   0   'False
            Top             =   2970
            Width           =   1005
         End
         Begin VB.TextBox txtInCollect 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C00000&
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   1110
            Locked          =   -1  'True
            TabIndex        =   125
            TabStop         =   0   'False
            Top             =   2970
            Width           =   1005
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "รวมค่าบริการที่เรียกเก็บ"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   71
            Left            =   7800
            TabIndex        =   167
            Top             =   3420
            Width           =   1695
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "บาท"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   70
            Left            =   10620
            TabIndex        =   166
            Top             =   3420
            Width           =   285
         End
         Begin VB.Shape Shape1 
            Height          =   2655
            Left            =   60
            Top             =   210
            Width           =   11055
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "บาท"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   54
            Left            =   6000
            TabIndex        =   140
            Top             =   3420
            Width           =   285
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "อื่นๆ ระบุ"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   53
            Left            =   420
            TabIndex        =   138
            Top             =   3420
            Width           =   615
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "บาท"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   52
            Left            =   10620
            TabIndex        =   136
            Top             =   3030
            Width           =   285
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ค่าอุปกรณ์"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   51
            Left            =   8790
            TabIndex        =   135
            Top             =   3030
            Width           =   705
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "บาท"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   50
            Left            =   8070
            TabIndex        =   133
            Top             =   3030
            Width           =   285
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ค่าคนงานเพิ่ม"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   35
            Left            =   5910
            TabIndex        =   132
            Top             =   3030
            Width           =   975
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "บาท"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   34
            Left            =   4950
            TabIndex        =   130
            Top             =   3030
            Width           =   285
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ค่าขนส่ง"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   33
            Left            =   3210
            TabIndex        =   129
            Top             =   3030
            Width           =   555
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "บาท"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   32
            Left            =   2190
            TabIndex        =   127
            Top             =   3030
            Width           =   285
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ค่าเก็บขน"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   24
            Left            =   390
            TabIndex        =   126
            Top             =   3030
            Width           =   675
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "ตรวจรับเอกสาร"
         Height          =   975
         Index           =   0
         Left            =   -74850
         TabIndex        =   118
         Top             =   450
         Width           =   11205
         Begin VB.CheckBox chkDoc5 
            Caption         =   "ใบควบคุมผู้ว่าจ้าง"
            Height          =   255
            Left            =   8880
            TabIndex        =   123
            Top             =   450
            Width           =   1605
         End
         Begin VB.CheckBox chkDoc4 
            Caption         =   "ใบน้ำหนักผู้ว่าจ้าง"
            Height          =   255
            Left            =   6816
            TabIndex        =   122
            Top             =   450
            Width           =   1545
         End
         Begin VB.CheckBox chkDoc3 
            Caption         =   "ใบน้ำหนักเข้าศูนย์"
            Height          =   255
            Left            =   4634
            TabIndex        =   121
            Top             =   450
            Width           =   1665
         End
         Begin VB.CheckBox chkDoc2 
            Caption         =   "ชุดใบกำกับ"
            Height          =   255
            Left            =   2932
            TabIndex        =   120
            Top             =   450
            Width           =   1185
         End
         Begin VB.CheckBox chkDoc1 
            Caption         =   "ใบบันทึกการปฏิบัติงาน"
            Height          =   255
            Left            =   450
            TabIndex        =   119
            Top             =   450
            Width           =   1965
         End
      End
      Begin VB.TextBox txtAumphur 
         Appearance      =   0  'Flat
         BackColor       =   &H00C00000&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   -66870
         TabIndex        =   113
         TabStop         =   0   'False
         Top             =   3300
         Width           =   2685
      End
      Begin VB.TextBox txtTumbol 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   -70050
         TabIndex        =   112
         Top             =   3300
         Width           =   1995
      End
      Begin VB.TextBox txtDestroyer 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   -73350
         TabIndex        =   98
         Top             =   2940
         Width           =   3375
      End
      Begin VB.TextBox txtDestroyID 
         Appearance      =   0  'Flat
         BackColor       =   &H00C00000&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   -67830
         Locked          =   -1  'True
         TabIndex        =   97
         Top             =   2940
         Width           =   3645
      End
      Begin VB.TextBox txtDestroyPlace 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   -73350
         TabIndex        =   96
         Top             =   3300
         Width           =   2475
      End
      Begin VB.TextBox txtRealWeight 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   -67260
         TabIndex        =   94
         Top             =   4020
         Width           =   1575
      End
      Begin VB.TextBox txtDestER 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   -66270
         TabIndex        =   93
         TabStop         =   0   'False
         Top             =   3660
         Width           =   2085
      End
      Begin VB.TextBox txtDestFax 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   -70050
         TabIndex        =   92
         TabStop         =   0   'False
         Top             =   3660
         Width           =   1995
      End
      Begin VB.TextBox txtDestTel 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   -73350
         TabIndex        =   91
         TabStop         =   0   'False
         Top             =   3660
         Width           =   2475
      End
      Begin VB.TextBox txtTreatTime 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   -70050
         TabIndex        =   90
         TabStop         =   0   'False
         Top             =   4380
         Width           =   1395
      End
      Begin VB.ComboBox cboTimeUnit 
         Height          =   315
         ItemData        =   "frMenifestUrgent.frx":0075
         Left            =   -68610
         List            =   "frMenifestUrgent.frx":0082
         Style           =   2  'Dropdown List
         TabIndex        =   89
         Top             =   4380
         Width           =   855
      End
      Begin VB.TextBox txtDestName 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   -73350
         TabIndex        =   88
         Top             =   4740
         Width           =   3825
      End
      Begin VB.Frame Frame1 
         Caption         =   "กรณีของเสียอันตรายไม่ตรงตามที่แจ้ง"
         Height          =   2235
         Left            =   -74940
         TabIndex        =   87
         Top             =   5130
         Width           =   11415
         Begin VB.CommandButton Command2 
            Caption         =   "เพิ่ม"
            Height          =   315
            Left            =   10650
            TabIndex        =   117
            Top             =   240
            Width           =   705
         End
         Begin MSComctlLib.ListView lvResult 
            Height          =   1905
            Left            =   90
            TabIndex        =   116
            Top             =   240
            Width           =   10545
            _ExtentX        =   18600
            _ExtentY        =   3360
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "ของเสีย"
               Object.Width           =   8819
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "ปริมาณ"
               Object.Width           =   4410
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "การดำเนินการ"
               Object.Width           =   4410
            EndProperty
         End
      End
      Begin VB.TextBox txtCollector2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1770
         TabIndex        =   82
         Top             =   3390
         Width           =   3375
      End
      Begin VB.TextBox txtCollectorID2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   8400
         TabIndex        =   81
         Top             =   3390
         Width           =   2595
      End
      Begin VB.ComboBox cboGenerator 
         Height          =   315
         Left            =   5880
         Style           =   2  'Dropdown List
         TabIndex        =   80
         Top             =   900
         Width           =   2715
      End
      Begin VB.ComboBox cboWasteType 
         Height          =   315
         ItemData        =   "frMenifestUrgent.frx":0096
         Left            =   9360
         List            =   "frMenifestUrgent.frx":00A0
         Style           =   2  'Dropdown List
         TabIndex        =   79
         Top             =   1620
         Width           =   915
      End
      Begin VB.ComboBox cboSiteName 
         Height          =   315
         Left            =   1770
         Style           =   2  'Dropdown List
         TabIndex        =   77
         Top             =   1260
         Width           =   9375
      End
      Begin VB.ComboBox cboTransUnit 
         Height          =   315
         ItemData        =   "frMenifestUrgent.frx":00AD
         Left            =   -65100
         List            =   "frMenifestUrgent.frx":00B7
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   1950
         Width           =   885
      End
      Begin VB.ComboBox cboSolidUnit 
         Height          =   315
         Left            =   10110
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   6210
         Width           =   1305
      End
      Begin VB.ComboBox cboLiquidUnit 
         Height          =   315
         Left            =   5340
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   6210
         Width           =   1305
      End
      Begin VB.TextBox txtCollectorID 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   8400
         TabIndex        =   15
         Top             =   3030
         Width           =   2595
      End
      Begin VB.TextBox txtCollector 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1770
         TabIndex        =   14
         Top             =   3030
         Width           =   3375
      End
      Begin VB.Frame Frame4 
         Caption         =   "ผู้ขนส่งของเสียอัตราย"
         Height          =   1005
         Left            =   90
         TabIndex        =   68
         Top             =   1980
         Width           =   11025
         Begin VB.TextBox txtTrans1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C00000&
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   240
            Width           =   3375
         End
         Begin VB.TextBox txtTransID1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C00000&
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   8310
            Locked          =   -1  'True
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   240
            Width           =   2595
         End
         Begin VB.TextBox txtTransID2 
            Appearance      =   0  'Flat
            BackColor       =   &H00C00000&
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   8310
            Locked          =   -1  'True
            TabIndex        =   13
            Top             =   600
            Width           =   2595
         End
         Begin VB.TextBox txtTrans2 
            Appearance      =   0  'Flat
            BackColor       =   &H00FF8080&
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   1680
            TabIndex        =   12
            Top             =   600
            Width           =   3375
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ชื่อบริษัท 1"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   6
            Left            =   840
            TabIndex        =   72
            Top             =   300
            Width           =   735
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "เลขประจำตัวผู้ขนส่งของเสียอันตรายรายที่ 1"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   14
            Left            =   5250
            TabIndex        =   71
            Top             =   300
            Width           =   2985
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "เลขประจำตัวผู้ขนส่งของเสียอันตรายรายที่ 2"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   21
            Left            =   5250
            TabIndex        =   70
            Top             =   660
            Width           =   2985
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ชื่อบริษัท 2"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   22
            Left            =   840
            TabIndex        =   69
            Top             =   660
            Width           =   765
         End
      End
      Begin VB.TextBox txtSign1B 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   -73080
         TabIndex        =   34
         Top             =   2340
         Width           =   4905
      End
      Begin VB.TextBox txtTransTime 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   -66480
         TabIndex        =   32
         Top             =   1950
         Width           =   1335
      End
      Begin VB.TextBox txtToProvince 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   -70080
         TabIndex        =   31
         Top             =   1950
         Width           =   1905
      End
      Begin VB.TextBox txtFromProvince 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   -73080
         TabIndex        =   30
         Top             =   1950
         Width           =   1875
      End
      Begin VB.TextBox txtCarRegisID1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   -66600
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   1200
         Width           =   2040
      End
      Begin VB.TextBox txtTransTel 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   -73080
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   840
         Width           =   2355
      End
      Begin VB.TextBox txtTransFax 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   -69720
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   840
         Width           =   2130
      End
      Begin VB.TextBox txtTransER 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   -66600
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   840
         Width           =   2025
      End
      Begin VB.TextBox txtTransID1B 
         Appearance      =   0  'Flat
         BackColor       =   &H00C00000&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   -67440
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   480
         Width           =   2865
      End
      Begin VB.TextBox txtTrans1B 
         Appearance      =   0  'Flat
         BackColor       =   &H00C00000&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   -73080
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   480
         Width           =   3345
      End
      Begin VB.TextBox txtCustER 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   6450
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   1620
         Width           =   1425
      End
      Begin VB.TextBox txtSigneName1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   3360
         TabIndex        =   22
         Top             =   6990
         Width           =   3765
      End
      Begin VB.TextBox txtTreatment 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   3360
         TabIndex        =   21
         Top             =   6600
         Width           =   8055
      End
      Begin VB.TextBox txtLiquidQty 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   3360
         TabIndex        =   17
         Top             =   6210
         Width           =   1935
      End
      Begin VB.TextBox txtSolidQty 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   8040
         TabIndex        =   19
         Top             =   6210
         Width           =   2025
      End
      Begin VB.TextBox txtCustFax 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   3870
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   1620
         Width           =   1545
      End
      Begin VB.TextBox txtCustTel 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1770
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1620
         Width           =   1395
      End
      Begin VB.TextBox txtCustomer 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         Height          =   315
         Left            =   1770
         TabIndex        =   4
         Top             =   510
         Width           =   6825
      End
      Begin WasteManagment.ctlGrid grItem 
         Height          =   1935
         Left            =   120
         TabIndex        =   16
         Top             =   4170
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   3413
      End
      Begin WasteManagment.ctlDate dtWorkDate 
         Height          =   315
         Left            =   9390
         TabIndex        =   6
         Top             =   900
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   556
         BackColor       =   -2147483633
      End
      Begin WasteManagment.ctlDate dtSign2 
         Height          =   315
         Left            =   -66480
         TabIndex        =   35
         Top             =   2310
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   556
         BackColor       =   -2147483633
      End
      Begin WasteManagment.ctlDate dtSign1 
         Height          =   315
         Left            =   8040
         TabIndex        =   85
         Top             =   6990
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   556
         BackColor       =   -2147483633
      End
      Begin WasteManagment.ctlDate dtSign3 
         Height          =   315
         Left            =   -67260
         TabIndex        =   99
         Top             =   4740
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   556
         BackColor       =   -2147483633
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "คิดค่าบริการจากปริมาณของ"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   69
         Left            =   6450
         TabIndex        =   164
         Top             =   3810
         Width           =   1890
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ใบสั่งงาน"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   68
         Left            =   8700
         TabIndex        =   161
         Top             =   600
         Width           =   630
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "เจ้าหน้าที่บัญชีผู้อัปเดทข้อมูล"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   59
         Left            =   -68520
         TabIndex        =   156
         Top             =   6930
         Width           =   1950
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "อำเภอ/จังหวัด"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   12
         Left            =   -67890
         TabIndex        =   115
         Top             =   3360
         Width           =   975
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ตำบล"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   9
         Left            =   -70530
         TabIndex        =   114
         Top             =   3360
         Width           =   375
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "สถานที่กำจัด"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   15
         Left            =   -74280
         TabIndex        =   111
         Top             =   3360
         Width           =   870
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "เลขประจำตัวผู้รับกำจัด"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   16
         Left            =   -69480
         TabIndex        =   110
         Top             =   3000
         Width           =   1575
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "นับจากวันที่ได้รับของเสีย"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   17
         Left            =   -67650
         TabIndex        =   109
         Top             =   4440
         Width           =   1725
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ปริมาณที่รับเข้าจริง"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   18
         Left            =   -68640
         TabIndex        =   108
         Top             =   4080
         Width           =   1305
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "และสามารถกำจัดของเสียที่รับมานี้ได้ภายในระยะเวลา"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   19
         Left            =   -73830
         TabIndex        =   107
         Top             =   4440
         Width           =   3705
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ชื่อผู้รับกำจัด"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   44
         Left            =   -74280
         TabIndex        =   106
         Top             =   3000
         Width           =   885
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "กรณีฉุกเฉิน"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   45
         Left            =   -67200
         TabIndex        =   105
         Top             =   3720
         Width           =   885
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "โทรศัพท์"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   46
         Left            =   -74040
         TabIndex        =   104
         Top             =   3720
         Width           =   615
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "โทรสาร"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   47
         Left            =   -70650
         TabIndex        =   103
         Top             =   3720
         Width           =   525
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "คำรับรอง : ข้าพเจ้าขอรับรองว่าได้รับวัสดุที่ไม่ใช้แล้วตามปริมาณที่ระบุดังนี้"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   48
         Left            =   -74160
         TabIndex        =   102
         Top             =   4080
         Width           =   5025
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ลงชื่อผู้รับกำจัด"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   49
         Left            =   -74460
         TabIndex        =   101
         Top             =   4800
         Width           =   1035
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "วันที่"
         Height          =   195
         Index           =   3
         Left            =   -67680
         TabIndex        =   100
         Top             =   4800
         Width           =   315
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "วันที่ลงชื่อ"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   57
         Left            =   7320
         TabIndex        =   86
         Top             =   7050
         Width           =   675
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "เลขประจำตัว TSDF รายที่ 2"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   56
         Left            =   6450
         TabIndex        =   84
         Top             =   3450
         Width           =   1905
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "บริษัท TSDF 2"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   42
         Left            =   690
         TabIndex        =   83
         Top             =   3450
         Width           =   1020
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ประเภทของเสีย"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   37
         Left            =   8220
         TabIndex        =   78
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ประเภทรถขนส่ง"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   7
         Left            =   -74280
         TabIndex        =   74
         Top             =   1260
         Width           =   1125
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "บริษัท TSDF 1"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   3
         Left            =   690
         TabIndex        =   73
         Top             =   3090
         Width           =   1005
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "วันที่"
         Height          =   195
         Index           =   0
         Left            =   -66840
         TabIndex        =   67
         Top             =   2370
         Width           =   315
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ลงชื่อผู้ขนส่งรายที่1"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   43
         Left            =   -74430
         TabIndex        =   66
         Top             =   2400
         Width           =   1290
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ใช้ระยะเวลาประมาณ"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   41
         Left            =   -68010
         TabIndex        =   65
         Top             =   2010
         Width           =   1455
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ไปยังจังหวัด"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   40
         Left            =   -70950
         TabIndex        =   64
         Top             =   2010
         Width           =   840
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "โดยขนส่งจากจังหวัด"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   39
         Left            =   -74550
         TabIndex        =   63
         Top             =   2010
         Width           =   1395
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "คำรับรอง : ข้าพเจ้าขอรับรองว่าได้ส่งมอบของเสียอันตรายแล้วตามที่ระบุข้างต้น และการขนส่งเป็นไปตามข้อกำหนดของกฎหมายทุกประการ"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   38
         Left            =   -73920
         TabIndex        =   62
         Top             =   1650
         Width           =   9330
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   30
         Index           =   2
         Left            =   -75000
         Top             =   2790
         Width           =   11565
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   30
         Index           =   0
         Left            =   -75000
         Top             =   1560
         Width           =   11565
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "เลขทะเบียน"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   28
         Left            =   -67500
         TabIndex        =   61
         Top             =   1260
         Width           =   840
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "โทรสาร"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   31
         Left            =   -70320
         TabIndex        =   60
         Top             =   900
         Width           =   525
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "โทรศัพท์"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   30
         Left            =   -73770
         TabIndex        =   59
         Top             =   900
         Width           =   615
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ฉุกเฉิน"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   29
         Left            =   -67200
         TabIndex        =   58
         Top             =   900
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "เลขประจำตัวผู้ขนส่ง"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   27
         Left            =   -68880
         TabIndex        =   57
         Top             =   540
         Width           =   1365
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "รายชื่อผู้ขนส่ง"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   26
         Left            =   -74100
         TabIndex        =   56
         Top             =   540
         Width           =   945
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "วันที่ขน"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   8
         Left            =   8790
         TabIndex        =   55
         Top             =   960
         Width           =   525
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ลงชื่อ"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   25
         Left            =   2940
         TabIndex        =   54
         Top             =   7050
         Width           =   375
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "การปฏิบัติที่มีลักษณะพิเศษ และข้อมูลเพิ่มเติม"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   23
         Left            =   150
         TabIndex        =   53
         Top             =   6660
         Width           =   3165
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "กรณีฉุกเฉิน"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   20
         Left            =   5520
         TabIndex        =   52
         Top             =   1680
         Width           =   885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "รายละเอียดของของเสียอันตรายที่ขนส่งเคลื่อนย้าย"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   50
         Top             =   3870
         Width           =   3435
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   45
         Index           =   1
         Left            =   120
         Top             =   4110
         Width           =   11325
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ของแข็ง"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   11
         Left            =   7410
         TabIndex        =   44
         Top             =   6270
         Width           =   555
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "รวมปริมาตรของเสียอันตรายทั้งหมด  ของเหลว"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   10
         Left            =   150
         TabIndex        =   43
         Top             =   6270
         Width           =   3165
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "เลขประจำตัว TSDF รายที่ 1"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   8
         Left            =   6450
         TabIndex        =   42
         Top             =   3090
         Width           =   1890
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "โทรศัพท์"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   5
         Left            =   1080
         TabIndex        =   41
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "โทรสาร"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   4
         Left            =   3270
         TabIndex        =   40
         Top             =   1680
         Width           =   525
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "เลขที่ผู้กำเนิดของเสีย (13 หลัก)"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   3690
         TabIndex        =   39
         Top             =   960
         Width           =   2115
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "สถานที่ก่อกำเนิด"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   540
         TabIndex        =   38
         Top             =   1320
         Width           =   1185
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ชื่อลูกค้า"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   0
         Left            =   1140
         TabIndex        =   37
         Top             =   570
         Width           =   570
      End
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
      Left            =   4530
      TabIndex        =   162
      Top             =   0
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Menifest No."
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   36
      Left            =   8640
      TabIndex        =   76
      Top             =   510
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "เลขที่"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   13
      Left            =   6210
      TabIndex        =   51
      Top             =   510
      Width           =   360
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   4
      Index           =   0
      X1              =   90
      X2              =   5770
      Y1              =   510
      Y2              =   510
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      Index           =   1
      X1              =   180
      X2              =   5890
      Y1              =   585
      Y2              =   585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ข้อมูลใบ Menifest"
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
      Left            =   240
      TabIndex        =   36
      Top             =   60
      Width           =   1560
   End
   Begin VB.Image Image2 
      Height          =   795
      Left            =   0
      Picture         =   "frMenifestUrgent.frx":00C9
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11655
   End
End
Attribute VB_Name = "frmMenifestUrgent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim myMenifest_ As BWGMenifest
Dim curCustomer_ As BWGCustomer
Dim curTrans2_ As BWGTransportVendor
'''
Dim curTruckSubType_ As BWGTruckSubType
Dim curFromProvince_ As IMIProvince
Dim curToProvince_ As IMIProvince
Dim provLocation_ As String
Dim curTSDF_ As BWGTSDF
Dim curTSDFTumbol_ As IMITumbol
Dim isDetailChange_ As Boolean

Private Sub SetMSFlexGrid()
    With grItem.getGridObj
        .Cols = 9
        .ColWidth(0) = 0
        .ColWidth(1) = 495
        .ColWidth(2) = 3540
        .ColWidth(3) = 1110
        .ColWidth(4) = 795
        .ColWidth(5) = 1110
        .ColWidth(6) = 1440
        .ColWidth(7) = 1350
        .ColWidth(8) = 1200
        .Rows = 20
        .FixedAlignment(0) = 4
        .FixedAlignment(1) = 4
        .FixedAlignment(2) = 4
        .FixedAlignment(3) = 4
        .FixedAlignment(4) = 4
        .FixedAlignment(5) = 4
        .FixedAlignment(6) = 4
        .FixedAlignment(7) = 4
        .FixedAlignment(8) = 4
        .ColAlignment(0) = 4
        .ColAlignment(1) = 7
        .ColAlignment(2) = 1
        .ColAlignment(3) = 4
        .ColAlignment(4) = 7
        .ColAlignment(5) = 4
        .ColAlignment(6) = 7
        .ColAlignment(7) = 7
        .ColAlignment(8) = 1
        .TextMatrix(0, 1) = "ลำดับ"
        .TextMatrix(0, 2) = "รายละเอียด"
        .TextMatrix(0, 3) = "รหัสของเสีย"
        .TextMatrix(0, 4) = "จำนวน"
        .TextMatrix(0, 5) = "ชนิด"
        .TextMatrix(0, 6) = "ปริมาณผู้ก่อกำเนิด"
        .TextMatrix(0, 7) = "ปริมาณผู้รับกำจัด"
        .TextMatrix(0, 8) = "หน่วย"
    End With

    With grCustRate.getGridObj
        .Cols = 5
        .Rows = 50
        .ColWidth(0) = 0
        .ColWidth(1) = 5475
        .ColWidth(2) = 1755
        .ColWidth(3) = 1605
        .ColWidth(4) = 1860
        .TextMatrix(0, 1) = "ชื่อกากของเสีย"
        .TextMatrix(0, 2) = "ค่ากำจัดกาก (Quote)"
        .TextMatrix(0, 3) = "ค่าขนส่ง (Quote)"
        .TextMatrix(0, 4) = "ค่ากำจัดกาก (เรียกเก็บ)"
    End With
End Sub

Private Sub cboActualUnit_Validate(Cancel As Boolean)
    If cboActualUnit.ListIndex > -1 Then
        myMenifest_.QtyUnitID = colUnit_(cboActualUnit.ListIndex + 1).id
        myMenifest_.QtyUnitName = cboActualUnit.Text
    Else
        myMenifest_.QtyUnitID = ""
        myMenifest_.QtyUnitName = ""
    End If
End Sub

Private Sub cboCompany_Validate(Cancel As Boolean)
    If cboCompany.ListIndex >= 0 Then
        If cboCompany.ListCount > 1 Then
            If cboCompany.ListIndex = 0 Then Exit Sub
            myMenifest_.CompanyID = CurrentUser.WorkCompanies(cboCompany.ListIndex).id
        Else
            myMenifest_.CompanyID = CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).id
        End If
        myMenifest_.CompanyName = cboCompany.Text
    Else
        myMenifest_.CompanyID = ""
        myMenifest_.CompanyName = ""
    End If
End Sub

Private Sub cboGenerator_Validate(Cancel As Boolean)
    myMenifest_.WasteGenNo = cboGenerator.Text
End Sub

Private Sub cboLiquidUnit_Validate(Cancel As Boolean)
    If cboLiquidUnit.ListIndex > -1 Then
        myMenifest_.LiquidUnitID = colUnit_(cboLiquidUnit.ListIndex + 1).id
        myMenifest_.LiquidUnitName = cboLiquidUnit.Text
    Else
        myMenifest_.LiquidUnitID = ""
        myMenifest_.LiquidUnitName = ""
    End If
End Sub

Private Sub cboSiteName_Validate(Cancel As Boolean)
Dim tmpCustSite As BWGCustomerSite
    With myMenifest_
        If cboSiteName.ListIndex = -1 Then
            .CustSiteAddr = ""
            .CustSiteID = ""
            .CustSiteName = ""
            .FromProvinceID = ""
            .FromProvinceName = ""
        Else
            Set tmpCustSite = curCustomer_.SiteAddresses(cboSiteName.ListIndex + 1)
            .CustSiteAddr = cboSiteName.Text
            .CustSiteID = tmpCustSite.id
            .CustSiteName = tmpCustSite.SiteName
            .FromProvinceID = tmpCustSite.ProvinceID
            .FromProvinceName = tmpCustSite.ProvinceName
        End If
    End With
End Sub

Private Sub cboSolidUnit_Validate(Cancel As Boolean)
    If cboSolidUnit.ListIndex > -1 Then
        myMenifest_.SolidUnitID = colUnit_(cboSolidUnit.ListIndex + 1).id
        myMenifest_.SolidUnitName = cboSolidUnit.Text
    Else
        myMenifest_.SolidUnitID = ""
        myMenifest_.SolidUnitName = ""
    End If
End Sub

Private Sub cboTimeUnit_Validate(Cancel As Boolean)
    If cboTimeUnit.ListIndex = 0 Then
        myMenifest_.PeriodUnit = "D"
    ElseIf cboTimeUnit.ListIndex = 1 Then
        myMenifest_.PeriodUnit = "M"
    ElseIf cboTimeUnit.ListIndex = 2 Then
        myMenifest_.PeriodUnit = "Y"
    End If
End Sub

Private Sub cboTransUnit_Validate(Cancel As Boolean)
    If cboTransUnit.ListIndex = 0 Then
        myMenifest_.DayOrHour = "H"
    ElseIf cboTransUnit.ListIndex = 1 Then
        myMenifest_.DayOrHour = "D"
    End If
End Sub

Private Sub cboWasteType_Validate(Cancel As Boolean)
    myMenifest_.NonOrHz = cboWasteType.Text
End Sub

Private Sub cboWType_Validate(Cancel As Boolean)
    If cboWType.ListIndex = 0 Then
        myMenifest_.ChargeWType = "A"
    ElseIf cboWType.ListIndex = 1 Then
        myMenifest_.ChargeWType = "B"
    End If
End Sub

Private Sub chkDoc1_Validate(Cancel As Boolean)
    myMenifest_.DocJobRecord = chkDoc1.Value = 1
End Sub

Private Sub chkDoc2_Validate(Cancel As Boolean)
    myMenifest_.DocControl = chkDoc2.Value = 1
End Sub

Private Sub chkDoc3_Validate(Cancel As Boolean)
    myMenifest_.DocSiteWeight = chkDoc3.Value = 1
End Sub

Private Sub chkDoc4_Validate(Cancel As Boolean)
    myMenifest_.DocCustWeight = chkDoc4.Value = 1
End Sub

Private Sub chkDoc5_Validate(Cancel As Boolean)
    myMenifest_.DocCustControl = chkDoc5.Value = 1
End Sub

Private Sub cmdCancel_Click()
Dim tmpStr As String, CancelDetail As String
    If Trim(myMenifest_.BillingID) <> "" Then
        MsgBox "ไม่สามารถยกเลิกใบ Menifest นี้ได้เนื่องจากได้ทำการวางบิลเก็บเงินลูกค้าแล้ว", vbExclamation
        Exit Sub
    End If
    If Trim(myMenifest_.BillPayID) <> "" Then
        MsgBox "ไม่สามารถยกเลิกใบ Menifest นี้ได้เนื่องจากได้ทำการจ่ายเงิน Vendor แล้ว", vbExclamation
        Exit Sub
    End If
    If MsgBox("ต้องการยกเลิกใบ Menifest นี้หรือไม่?", vbQuestion + vbYesNo) = vbYes Then
        CancelDetail = InputBox("ใส่รายละเอียดในการยกเลิก", "ใส่รายละเอียดการยกเลิก", "")
        If Trim(Replace(CancelDetail, "*", "")) = "" Then
            MsgBox "กรุณาระบุเหตุผลของการยกเลิก", vbExclamation + vbOKOnly, "Warning"
            Exit Sub
        End If
        tmpStr = myMenifest_.Cancel(CancelDetail)
        If Trim(tmpStr) = "" Then
            'ClearScreen
            'CreateNewMenifest
            'txtCustomer.SetFocus
            SetCancelScreen myMenifest_.IsCanceled
        Else
            MsgBox tmpStr
        End If
    End If
End Sub

Private Sub cmdClear_Click()
Dim selResult As Integer
    If myMenifest_.isChanged And cmdSave.Enabled = True Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนสร้างใหม่หรือไม่?", vbQuestion + vbYesNoCancel)
        If selResult = vbYes Then
            CmdSave_Click
            If myMenifest_.isChanged Then Exit Sub
        ElseIf selResult = vbCancel Then
            Exit Sub
        End If
    End If
    ClearScreen
    CreateNewMenifest
    txtCustomer.SetFocus
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub CmdSave_Click()
    If Trim(txtCustomer.Text) = "" Then
        MsgBox "กรุณาระบุลูกค้า", vbExclamation
        txtCustomer.SetFocus
        Exit Sub
    End If
'    If Trim(cboGenerator.Text) = "" Then
'        MsgBox "กรุณาระบุเลขประจำตัวผู้กำเนิดของเสีย", vbExclamation
'        cboGenerator.SetFocus
'        Exit Sub
'    End If
    If myMenifest_.isChanged And cmdSave.Enabled = True Then
        myMenifest_.Save
        txtDocNo.Text = myMenifest_.DocNo
        txtMenifestNo.Text = myMenifest_.MenifestNo
    End If
    If myMenifest_.isChanged = False Then MsgBox "จัดเก็บข้อมูลเรียบร้อยแล้ว", vbInformation
End Sub

Private Sub cmdSearch_Click()
Dim selResult As Integer
    If myMenifest_.isChanged And cmdSave.Enabled = True Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนทำงานต่อหรือไม่?", vbQuestion + vbYesNoCancel)
        If selResult = vbYes Then
            CmdSave_Click
            If myMenifest_.isChanged Then Exit Sub
        ElseIf selResult = vbCancel Then
            Exit Sub
        End If
    End If
    Load frmMenifestSearch
    With frmMenifestSearch
        .SetCallerForm Me
        .Show 1
    End With
End Sub

Private Sub Command1_Click()
Dim selResult As Integer
Dim x%
Dim tmpJobDet As BWGJobDataDetail
    If Trim(txtCustomer.Text) = "" Then
        MsgBox "กรุณาระบุลูกค้า", vbExclamation
        txtCustomer.SetFocus
        Exit Sub
    End If
    
    If myMenifest_.isClosed = False Then
        For x = 1 To myMenifest_.JobDetails.Count
            Set tmpJobDet = myMenifest_.JobDetails(x)
            If tmpJobDet.DisposerWeight = 0 Then
                MsgBox "กรุณาระบุปริมาณผู้รับกำจัดให้ครบถ้วนทุกรายการของเสีย", vbExclamation
                SSTab1.Tab = 0
                grItem.SetFocus
                Set tmpJobDet = Nothing
                Exit Sub
            Else
                If tmpJobDet.QuoTreatmentUnitName = "บาท/เที่ยว" Then
                    tmpJobDet.TreatmentCharge = tmpJobDet.QuoTreatmentRate
                Else
                    tmpJobDet.TreatmentCharge = tmpJobDet.QuoTreatmentRate * tmpJobDet.DisposerWeight
                End If
                grCustRate.getGridObj.TextMatrix(x, 4) = tmpJobDet.TreatmentCharge
            End If
            Set tmpJobDet = Nothing
        Next
        SumTreatmentCharge
    End If
    
    If Trim(txtRealWeight.Text) = "" Then
        MsgBox "กรุณาระบุปริมาณที่รับเข้าจริง", vbExclamation
        SSTab1.Tab = 1
        txtRealWeight.SetFocus
        Exit Sub
    End If
    If IsNumeric(Trim(txtRealWeight.Text)) = False Then
        MsgBox "กรุณาใส่ปริมาณที่รับเข้าจริงให้ถูกต้อง", vbExclamation
        SSTab1.Tab = 1
        txtRealWeight.SetFocus
        Exit Sub
    End If
    
    If myMenifest_.isChanged And cmdSave.Enabled = True Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนทำงานต่อหรือไม่?", vbQuestion + vbYesNo)
        If selResult = vbYes Then
            CmdSave_Click
            If myMenifest_.isChanged Then Exit Sub
        Else 'If selResult = vbCancel Then
            Exit Sub
        End If
    End If
    Load frmMenifestClose
    With frmMenifestClose
        .SetMenifestData myMenifest_
        .Show 1
    End With
End Sub

Private Sub Command2_Click()
Dim selResult As Integer
    If Trim(txtCustomer.Text) = "" Then
        MsgBox "กรุณาระบุชื่อลูกค้า", vbExclamation
        txtCustomer.SetFocus
        Exit Sub
    End If
    If myMenifest_.isChanged Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนสร้างใหม่หรือไม่?", vbQuestion + vbYesNo)
        If selResult = vbYes Then
            CmdSave_Click
            If myMenifest_.isChanged Then Exit Sub
        Else 'If selResult = vbCancel Then
            Exit Sub
        End If
    End If
    Load frmMenifestDiscrep
    With frmMenifestDiscrep
        .SetCallerForm Me, myMenifest_
        .Show 1
        myMenifest_.RequeryDiscrepWaste
        ShowDiscrepWasteData
    End With
End Sub

Private Sub Command3_Click()
    If Trim(txtCustomer.Text) = "" Then
        MsgBox "กรุณาเลือกลูกค้าก่อนทำการเพิ่มรายการกากของเสีย", vbExclamation
        txtCustomer.SetFocus
        Exit Sub
    End If
    If cboCompany.ListIndex < 0 Then
        MsgBox "กรุณาระบุบริษัทที่ออกใบ Manifest ก่อนการเพิ่มกากของเสีย", vbExclamation
        cboCompany.SetFocus
        Exit Sub
    End If
    isDetailChange_ = False
    Load frmJobWasteDataAdd
    With frmJobWasteDataAdd
        If cboGenerator.ListIndex = 0 Then
            .SetCallerForm Me, curCustomer_, "", CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).id, "", ""
        Else
            .SetCallerForm Me, curCustomer_, cboGenerator.Text, CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).id, "", ""
        End If
        .Show 1
        If isDetailChange_ Then
            ShowWasteData
        End If
    End With
End Sub

Private Sub dtSign1_Validate(Cancel As Boolean)
    myMenifest_.SignDate1 = dtSign1.ValueYMD
End Sub

Private Sub dtSign2_Validate(Cancel As Boolean)
    myMenifest_.SignDate2 = dtSign2.ValueYMD
End Sub

Private Sub dtSign3_Validate(Cancel As Boolean)
    myMenifest_.SignDate3 = dtSign3.ValueYMD
End Sub

Private Sub dtWorkDate_Validate(Cancel As Boolean)
    myMenifest_.WorkDate = dtWorkDate.ValueYMD
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 And Me.ActiveControl.Name <> "grItem" And Me.ActiveControl.Name <> "grCustRate" Then SendKeys "{Tab}"
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

    With frmPicture
        cmdSave.Picture = .picSave.Picture
        cmdSearch.Picture = .picSearch.Picture
        cmdClear.Picture = .picNew.Picture
        cmdCancel.Picture = .picDelete.Picture
        cmdClose.Picture = .picExit.Picture
    End With
    With cboWType
        .AddItem "ปริมาณผู้รับกำจัด"
        .AddItem "ปริมาณผู้ก่อกำเนิด"
        .ListIndex = 0
    End With
    txtAccStaff.Text = CurrentUser.StaffFName & " " & CurrentUser.StaffLName
    Call SetMSFlexGrid
    'cmdPeriod.ListIndex = 0
    PopulateCompany cboCompany, cboIndex
    SSTab1.Tab = 0
    PopulateUnit
    cboWasteType.ListIndex = 0
    CreateNewMenifest
    Set curCustomer_ = New BWGCustomer
End Sub

Private Sub PopulateUnit()
On Error Resume Next
Dim x%
    cboLiquidUnit.Clear
    cboSolidUnit.Clear
    If colUnit_ Is Nothing Then Exit Sub
    For x = 1 To colUnit_.Count
        cboLiquidUnit.AddItem colUnit_(x).UnitName
        cboSolidUnit.AddItem colUnit_(x).UnitName
        cboActualUnit.AddItem colUnit_(x).UnitName
    Next
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim selResult As Integer
    If myMenifest_ Is Nothing Then Exit Sub
    If myMenifest_.isChanged And cmdSave.Enabled = True Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนปิดหน้าจอหรือไม่?", vbQuestion + vbYesNoCancel)
        If selResult = vbYes Then
            CmdSave_Click
            If myMenifest_.isChanged Then Cancel = 1
        ElseIf selResult = vbCancel Then
            Cancel = 1
        End If
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
Set myMenifest_ = Nothing
Set colUnit_ = Nothing
Set curCustomer_ = Nothing
Set curTrans2_ = Nothing
Set curTruckSubType_ = Nothing
Set curFromProvince_ = Nothing
Set curToProvince_ = Nothing
Set curTSDFTumbol_ = Nothing
End Sub
Private Sub grCustRate_BeforeUpdateValue(NewValue As String, row As Integer, col As Integer, Cancel As Boolean)
    If col = 4 And IsNumeric(NewValue) = False Then
        MsgBox "กรุณาระบุข้อมูลเป็นตัวเลข", vbExclamation
        Cancel = True
    End If
End Sub
Private Sub grCustRate_Click(row As Integer, col As Integer)
    'MsgBox grCustRate.getGridObj.ColWidth(col)
End Sub
Private Sub grCustRate_DblClick(row As Integer, col As Integer)
Dim tmpKey As String
Dim qID As String
Dim tmpS As New BWGSearchManager, tmpCol As Collection
     tmpKey = "'" & grCustRate.getGridObj.TextMatrix(row, 0) & "'"
     If isExist(myMenifest_.JobDetails, tmpKey) Then
        If col <> 4 Then
            qID = myMenifest_.JobDetails(tmpKey).QuotationID
            Set tmpCol = tmpS.QuotationHeaderSearch("QuotationID='" & qID & "'")
            If tmpCol.Count > 0 Then
                Load frmQuotation
                With frmQuotation
                    .SetDataQuotation tmpCol(1)
                    .Show 1
                End With
            End If
            Set tmpS = Nothing
            Set tmpCol = Nothing
        End If
    End If
End Sub
Private Sub grCustRate_RolColChanged(row As Integer, col As Integer)
    If Trim(grCustRate.getGridObj.TextMatrix(row, 0)) <> "" Then
        grCustRate.ColEnabled(col) = col = 4
    Else
        grCustRate.ColEnabled(col) = False
    End If
End Sub
Private Sub grCustRate_UpdateValue(row As Integer, col As Integer)
Dim tmpKey As String, tmpVal As String
Dim tmpDet As BWGJobDataDetail
    tmpKey = "'" & Trim(grCustRate.getGridObj.TextMatrix(row, 0)) & "'"
    tmpVal = Trim(grCustRate.getGridObj.TextMatrix(row, col))
    If isExist(myMenifest_.JobDetails, tmpKey) Then
        Set tmpDet = myMenifest_.JobDetails(tmpKey)
        Select Case col
            Case 4
                If IsNumeric(tmpVal) Then
                    tmpDet.TreatmentCharge = tmpVal
                Else
                    tmpDet.TreatmentCharge = 0
                End If
        End Select
        SumTreatmentCharge
        Set tmpDet = Nothing
        myMenifest_.isChanged = True
    End If
End Sub
Private Sub grItem_BeforeUpdateValue(NewValue As String, row As Integer, col As Integer, Cancel As Boolean)
    If (col = 4 Or col = 6 Or col = 7) And IsNumeric(NewValue) = False Then
        MsgBox "กรุณาระบุข้อมูลเป็นตัวเลขเท่านั้น", vbExclamation
        Cancel = True
    End If
End Sub

Private Sub grItem_Click(row As Integer, col As Integer)
    'MsgBox grItem.getGridObj.ColWidth(col)
End Sub

Private Sub grItem_RolColChanged(row As Integer, col As Integer)
    If Trim(grItem.getGridObj.TextMatrix(row, 1)) = "" Then
        grItem.ColEnabled(col) = False
    Else
        grItem.ColEnabled(col) = (col = 4 Or col = 5 Or col = 6 Or col = 7)
    End If
End Sub

Private Sub grItem_UpdateValue(row As Integer, col As Integer)
Dim tmpKey As String, tmpVal As String
Dim tmpDet As BWGJobDataDetail
    tmpKey = "'" & Trim(grItem.getGridObj.TextMatrix(row, 0)) & "'"
    tmpVal = Trim(grItem.getGridObj.TextMatrix(row, col))
    If isExist(myMenifest_.JobDetails, tmpKey) Then
        Set tmpDet = myMenifest_.JobDetails(tmpKey)
        Select Case col
            Case 4
                If IsNumeric(tmpVal) Then
                    tmpDet.ContainerCount = tmpVal
                Else
                    tmpDet.ContainerCount = 0
                End If
            Case 5
                tmpDet.ContainerTypeDesc = tmpVal
            Case 6
                If IsNumeric(tmpVal) Then
                    tmpDet.EstWasteQty = tmpVal
                Else
                    tmpDet.EstWasteQty = 0
                End If
            Case 7
                If IsNumeric(tmpVal) Then
                    tmpDet.DisposerWeight = tmpVal
                Else
                    tmpDet.DisposerWeight = 0
                End If
        End Select
        If col = 6 Or col = 7 Then
            If tmpDet.QuoTreatmentUnitName = "บาท/เที่ยว" Then
                tmpDet.TreatmentCharge = tmpDet.QuoTreatmentRate
            Else
                If myMenifest_.ChargeWType = "A" Then
                    tmpDet.TreatmentCharge = tmpDet.QuoTreatmentRate * tmpDet.DisposerWeight
                Else
                    tmpDet.TreatmentCharge = tmpDet.QuoTreatmentRate * tmpDet.EstWasteQty
                End If
            End If
            grCustRate.getGridObj.TextMatrix(row, 4) = tmpDet.TreatmentCharge
            SumTreatmentCharge
        End If
        myMenifest_.isChanged = True
        Set tmpDet = Nothing
    End If
End Sub

Private Sub lvResult_DblClick()
    If lvResult.SelectedItem Is Nothing Then Exit Sub
    Load frmMenifestDiscrep
    With frmMenifestDiscrep
        .SetCallerForm Me, myMenifest_
        .SetDataMenifestDiscrep myMenifest_.DiscrepWastes(lvResult.SelectedItem.Key)
        .Show 1
        myMenifest_.RequeryDiscrepWaste
        ShowDiscrepWasteData
    End With
End Sub

Private Sub txtCarRegisID1_Validate(Cancel As Boolean)
    myMenifest_.CarRegisID = txtCarRegisID1.Text
End Sub

Private Sub txtCollector_Validate(Cancel As Boolean)
    myMenifest_.TSDFName1 = Trim(txtCollector.Text)
End Sub

Private Sub txtCollector2_Validate(Cancel As Boolean)
    myMenifest_.TSDFName2 = Trim(txtCollector2.Text)
End Sub

Private Sub txtCollectorID_Validate(Cancel As Boolean)
    myMenifest_.TSDFRegisNo1 = Trim(txtCollectorID.Text)
End Sub

Private Sub txtCollectorID2_Validate(Cancel As Boolean)
    myMenifest_.TSDFRegisNo2 = txtCollectorID2.Text
End Sub

Private Sub txtCustER_Validate(Cancel As Boolean)
    myMenifest_.ErContact = Trim(txtCustER.Text)
End Sub

Private Sub txtCustFax_Validate(Cancel As Boolean)
    myMenifest_.FaxNo = Trim(txtCustFax.Text)
End Sub

Private Sub txtCustomer_Validate(Cancel As Boolean)
    If curCustomer_ Is Nothing Then Set curCustomer_ = New BWGCustomer
    If Trim(txtCustomer.Text) = "" And Trim(curCustomer_.CustomerName) <> Trim(txtCustomer.Text) Then
        Set curCustomer_ = Nothing
        Set curCustomer_ = New BWGCustomer
        With myMenifest_
            .CustomerID = ""
            .CustomerName = ""
        End With
    Else
        If Trim(curCustomer_.CustomerName) <> Trim(txtCustomer.Text) Then
            Set curCustomer_ = Nothing
            Set curCustomer_ = New BWGCustomer
            With myMenifest_
                .CustomerID = ""
                .CustomerName = ""
            End With
            Load frmCustomerSearch
            With frmCustomerSearch
                .getCriteria (txtCustomer.Text)
                .SetCallerForm Me
                .Show 1
            End With
            If myMenifest_.CustomerID = "" Then txtCustomer.Text = ""
        End If
    End If
End Sub

Public Sub SetDataCustomer(selCustomer As BWGCustomer, Optional isAssigned As Boolean = True)
On Error GoTo ErrD
Dim tmpRec As New ADODB.Recordset
Dim x%, sqlStr As String
Dim tmpWaste As BWGWasteDataCR
    Set curCustomer_ = Nothing
    Set curCustomer_ = selCustomer
    If curCustomer_ Is Nothing Then Exit Sub
    cboSiteName.Clear
    
    With curCustomer_
        For x = 1 To .SiteAddresses.Count
            cboSiteName.AddItem .SiteAddresses(x).Address & " " & .SiteAddresses(x).TumbolName & " " & .SiteAddresses(x).AumphurName & " " & .SiteAddresses(x).ProvinceName
        Next
        sqlStr = "Select FactoryRegisNo from tbWasteDataCR WHERE CustomerID='" & .id & "' Group By FactoryRegisNo"
        Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
        If tmpRec Is Nothing Then Exit Sub
        cboGenerator.Clear
        cboGenerator.AddItem "ไม่ระบุเลขทะเบียนโรงงาน"
        Do Until tmpRec.EOF
            If Trim("" & tmpRec("FactoryRegisNo")) <> "" Then cboGenerator.AddItem Trim("" & tmpRec("FactoryRegisNo"))
            tmpRec.MoveNext
        Loop
        If cboGenerator.ListCount = 2 Then
            cboGenerator.ListIndex = 1
        Else
            cboGenerator.ListIndex = 0
        End If
        tmpRec.Close
        Set tmpRec = Nothing
        
        txtCustomer.Text = .CustomerName
        If isAssigned Then
            myMenifest_.CustomerID = .id
            myMenifest_.CustomerName = .CustomerName
        End If
    End With
'    With grItem.getGridObj
'        grItem.ClearAllData
'        .Rows = curCustomer_.WasteItems.Count + 20
'        For x = 1 To curCustomer_.WasteItems.Count
'            Set tmpWaste = curCustomer_.WasteItems(x)
'            .TextMatrix(x, 0) = tmpWaste.ID
'            .TextMatrix(x, 1) = x
'            .TextMatrix(x, 2) = tmpWaste.WasteName
'            .TextMatrix(x, 3) = tmpWaste.EUCode1 & "-" & tmpWaste.EUCode2 & "-" & tmpWaste.EUCode3
'            .TextMatrix(x, 4) = tmpWaste.Moud
'            .TextMatrix(x, 5) = tmpWaste.Kor
'            Set tmpWaste = Nothing
'        Next
'    End With
Exit Sub
ErrD:
    GetMsg "", vbCritical
    Exit Sub
End Sub

Private Sub CreateNewMenifest()
On Error GoTo ErrD
    Set myMenifest_ = New BWGMenifest
    With myMenifest_
        .id = GetGUID
        .OpenDate = TodayDate
        .WorkDate = dtWorkDate.ValueYMD
        .NonOrHz = cboWasteType.Text
        .OpenStaffID = CurrentUser.id
        .TransportName1 = "บริษัท  อัคคีปราการ  จำกัด (มหาชน)"
        .TransRegisNo1 = "DIW-T-085800068"
        txtTrans1.Text = "บริษัท  อัคคีปราการ  จำกัด (มหาชน)"
        txtTransID1.Text = "DIW-T-085800068"
        If cboWType.ListIndex = 0 Then
            .ChargeWType = "A"
        ElseIf cboWType.ListIndex = 1 Then
            .ChargeWType = "B"
        End If
        If cboCompany.ListCount > 1 Then
            If cboCompany.ListIndex > 0 Then
                .CompanyID = CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).id
                .CompanyName = CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).CompanyNameTH
            End If
        Else
            If cboCompany.ListIndex >= 0 Then
                .CompanyID = CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).id
                .CompanyName = CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).CompanyNameTH
            End If
        End If
        .isChanged = False
    End With
    Exit Sub
ErrD:
    GetMsg "", vbCritical
    Exit Sub
End Sub



Private Sub ClearScreen()
    txtDocNo.Text = "== AUTO =="
    txtMenifestNo.Text = "== AUTO =="
    txtCustomer.Text = ""
    cboWType.ListIndex = 0
    SetCancelScreen False
    Set curCustomer_ = Nothing
    Set curCustomer_ = New BWGCustomer
    cboGenerator.ListIndex = -1
    dtWorkDate.ValueYMD = TodayDate
    cboSiteName.Clear
    txtCustTel.Text = ""
    txtCustFax.Text = ""
    txtCustER.Text = ""
    cboWasteType.ListIndex = 0
    txtTrans1.Text = "เบตเตอร์ เวิลด์ ทรานสปอร์ต จำกัด"
    txtTransID1.Text = "DIW-T-050200740"
    txtTrans2.Text = ""
    txtTransID2.Text = ""
    txtCollector.Text = ""
    txtCollectorID.Text = ""
    txtCollector2.Text = ""
    txtCollectorID2.Text = ""
    grItem.ClearAllData
    grItem.getGridObj.Rows = 20
    txtLiquidQty.Text = ""
    cboLiquidUnit.ListIndex = -1
    txtSolidQty.Text = ""
    cboSolidUnit.ListIndex = -1
    txtTreatment.Text = ""
    txtSigneName1.Text = ""
    dtSign1.ValueYMD = TodayDate
    
    txtTrans1B.Text = ""
    txtTransID1B.Text = ""
    txtTransTel.Text = ""
    txtTransFax.Text = ""
    txtTransER.Text = ""
    txtTruckType.Text = ""
    txtCarRegisID1.Text = ""
    txtFromProvince.Text = ""
    txtToProvince.Text = ""
    txtTransTime.Text = ""
    cboTransUnit.ListIndex = -1
    txtSign1B.Text = ""
    dtSign2.ValueYMD = TodayDate
    txtDestroyer.Text = ""
    txtDestroyID.Text = ""
    txtDestroyPlace.Text = ""
    txtTumbol.Text = ""
    txtAumphur.Text = ""
    txtDestTel.Text = ""
    txtDestFax.Text = ""
    txtDestER.Text = ""
    txtRealWeight.Text = ""
    txtTreatTime.Text = ""
    cboTimeUnit.ListIndex = -1
    txtDestName.Text = ""
    dtSign3.ValueYMD = TodayDate
    lvResult.ListItems.Clear
    chkDoc1.Value = 0
    chkDoc2.Value = 0
    chkDoc3.Value = 0
    chkDoc4.Value = 0
    chkDoc5.Value = 0
    txtInCollect.Text = ""
    txtIntrans.Text = ""
    txtInWorker.Text = ""
    txtInEquip.Text = ""
    txtOtherDesc.Text = ""
    txtOtherFee.Text = ""
    txtOutTrans.Text = ""
    txtWorkerCount.Text = ""
    txtOutWorker.Text = ""
    txtOutEquip.Text = ""
    txtOutOther.Text = ""
    txtOutOtherFee.Text = ""
    grCustRate.ClearAllData
    grCustRate.getGridObj.Rows = 50
End Sub

Public Sub SetDataMenifest(selMenifest As BWGMenifest)
Dim tmpS As New BWGSearchManager
Dim tmpCol As Collection, tmpJobDet As BWGJobDataDetail
Dim x%, tmpGrid As MSFlexGrid
Dim maxTrans As Double
Dim tmpChargeRate As BWGTransVendorChargeRate
    Set myMenifest_ = selMenifest
    With myMenifest_
        SetCancelScreen .IsCanceled
        txtDocNo.Text = .DocNo
        txtMenifestNo.Text = .MenifestNo
        txtJobNo.Text = .JobNo
        Set tmpCol = tmpS.CustomerSearch("CustomerID='" & .CustomerID & "'")
        If tmpCol.Count > 0 Then SetDataCustomer tmpCol(1), False
        SetTextToCombo .CompanyName, cboCompany
        If Trim(.WasteGenNo) = "" Then
            cboGenerator.ListIndex = 0
        Else
            SetTextToCombo .WasteGenNo, cboGenerator
        End If
        If .ChargeWType = "A" Then
            cboWType.ListIndex = 0
        ElseIf .ChargeWType = "B" Then
            cboWType.ListIndex = 1
        End If
        dtWorkDate.ValueYMD = .WorkDate
        SetTextToCombo .CustSiteAddr, cboSiteName
        txtCustTel.Text = .TelNo
        txtCustFax.Text = .FaxNo
        txtCustER.Text = .ErContact
        SetTextToCombo .NonOrHz, cboWasteType
        If Trim(.TruckTypeName) <> "รถโรงงานลูกค้า" Then
            txtTrans1.Text = .TransportName1 ' "บริษัท  อัคคีปราการ  จำกัด (มหาชน)"
            txtTransID1.Text = .TransRegisNo1  '"DIW-T-085800068"
        Else
            txtTrans1.Text = ""
            txtTransID1.Text = ""
        End If
        If Trim(.Transport2ID) = "" Then
            txtTrans2.Text = .TransportName2
            txtTransID2.Text = .TransRegisNo2
        Else
            Set tmpCol = tmpS.TransportVenderSearch("VendorID='" & .Transport2ID & "'")
            If tmpCol.Count > 0 Then SetDataTransportVender tmpCol(1), False
        End If
        
        txtCollector.Text = .TSDFName1
        txtCollectorID.Text = .TSDFRegisNo1
        txtCollector2.Text = .TSDFName2
        txtCollectorID2.Text = .TSDFRegisNo2
        SetTextToCombo .QtyUnitName, cboActualUnit
        txtLiquidQty.Text = .TotalLiquid
        SetTextToCombo .LiquidUnitName, cboLiquidUnit
        txtSolidQty.Text = .TotalSolid
        SetTextToCombo .SolidUnitName, cboSolidUnit
        txtTreatment.Text = .SpecialComment
        txtSigneName1.Text = .SignName1
        dtSign1.ValueYMD = .SignDate1

        txtTrans1B.Text = .TransportName2
        txtTransID1B.Text = .TransRegisNo2
        
        txtTransTel.Text = .TransTelNo1
        txtTransFax.Text = .TransFaxNo1
        txtTransER.Text = .TransErContact1
        txtTruckType.Text = .TruckSubTypeName & ", " & .TruckTypeName
        txtCarRegisID1.Text = .CarRegisID
        txtFromProvince.Text = .FromProvinceName
        txtToProvince.Text = .ToProvinceName
        txtTransTime.Text = .TimeSpending
        If Trim(.DayOrHour) = "H" Then
            cboTransUnit.ListIndex = 0
        ElseIf Trim(.DayOrHour) = "D" Then
            cboTransUnit.ListIndex = 1
        End If
        txtSign1B.Text = .SignName2
        dtSign2.ValueYMD = .SignDate2
        txtDestroyer.Text = .TSDFName1
        txtDestroyID.Text = .TSDFRegisNo1
        txtDestroyPlace.Text = .TSDFAddr
        txtTumbol.Text = .TSDFTumbolName
        txtAumphur.Text = .TSDFAumphurName & " " & .TSDFProvinceName
        txtDestTel.Text = .TSDFTelNo
        txtDestFax.Text = .TSDFFaxNo
        txtDestER.Text = .TSDFErContact
        txtRealWeight.Text = .ActualQty
        txtTreatTime.Text = .TreatmentPeriod
        If .PeriodUnit = "D" Then
            cboTimeUnit.ListIndex = 0
        ElseIf .PeriodUnit = "M" Then
            cboTimeUnit.ListIndex = 1
        ElseIf .PeriodUnit = "Y" Then
            cboTimeUnit.ListIndex = 2
        End If
        txtDestName.Text = .SignName3
        dtSign3.ValueYMD = .SignDate3
        
        If .DocJobRecord = True Then
            chkDoc1.Value = 1
        Else
            chkDoc1.Value = 0
        End If
        If .DocControl = True Then
            chkDoc2.Value = 1
        Else
            chkDoc2.Value = 0
        End If
        If .DocSiteWeight = True Then
            chkDoc3.Value = 1
        Else
            chkDoc3.Value = 0
        End If
        If .DocCustWeight = True Then
            chkDoc4.Value = 1
        Else
            chkDoc4.Value = 0
        End If
        If .DocCustControl = True Then
            chkDoc5.Value = 1
        Else
            chkDoc5.Value = 0
        End If
        
        'รอทำ Menifest Discrep
        'lvResult.ListItems.Clear
        grItem.ClearAllData
        Set tmpGrid = grItem.getGridObj
        tmpGrid.Rows = .JobDetails.Count + 20
        For x = 1 To .JobDetails.Count
            Set tmpJobDet = .JobDetails(x)
            tmpGrid.TextMatrix(x, 0) = tmpJobDet.id
            tmpGrid.TextMatrix(x, 1) = x
            tmpGrid.TextMatrix(x, 2) = tmpJobDet.WasteName & " [" & tmpJobDet.WasteQuotationNo & "]"
            tmpGrid.TextMatrix(x, 3) = tmpJobDet.WasteEUCode
            tmpGrid.TextMatrix(x, 4) = tmpJobDet.ContainerCount
            tmpGrid.TextMatrix(x, 5) = tmpJobDet.ContainerTypeDesc
            tmpGrid.TextMatrix(x, 6) = tmpJobDet.EstWasteQty
            tmpGrid.TextMatrix(x, 7) = tmpJobDet.DisposerWeight
            If Trim(tmpJobDet.QuoTreatmentUnitName) = "บาท/ถัง" Then
                tmpGrid.TextMatrix(x, 8) = "ถัง"
            ElseIf Trim(tmpJobDet.QuoTreatmentUnitName) = "บาท/ลูกบาศก์เมตร" Then
                tmpGrid.TextMatrix(x, 8) = "ลูกบาศก์เมตร"
            Else
                tmpGrid.TextMatrix(x, 8) = "ตัน"
            End If
            Set tmpJobDet = Nothing
        Next
        Set tmpGrid = grCustRate.getGridObj
        tmpGrid.Rows = .JobDetails.Count + 20
        For x = 1 To .JobDetails.Count
            Set tmpJobDet = .JobDetails(x)
            tmpGrid.TextMatrix(x, 0) = tmpJobDet.id
            tmpGrid.TextMatrix(x, 1) = tmpJobDet.WasteName & " [" & tmpJobDet.WasteQuotationNo & "]"
            tmpGrid.TextMatrix(x, 2) = Format(tmpJobDet.QuoTreatmentRate, "#,##0.00") & " " & tmpJobDet.QuoTreatmentUnitName
            If tmpJobDet.isPriceIncTrans Then
                tmpGrid.TextMatrix(x, 3) = "รวมค่าขนส่ง"
            Else
                tmpGrid.TextMatrix(x, 3) = Format(tmpJobDet.QuoTransportFee, "#,##0.00") & " " & tmpJobDet.QuoTransUnitName
                If maxTrans < tmpJobDet.QuoTransportFee Then maxTrans = tmpJobDet.QuoTransportFee
            End If
            If tmpJobDet.QuoTreatmentUnitName = "บาท/เที่ยว" And tmpJobDet.TreatmentCharge = 0 Then
                tmpJobDet.TreatmentCharge = tmpJobDet.QuoTreatmentRate
            End If
            tmpGrid.TextMatrix(x, 4) = tmpJobDet.TreatmentCharge
            Set tmpJobDet = Nothing
        Next
        If .InCollectFee > 0 Then
            txtInCollect.Text = .InCollectFee
            txtIntrans.Text = .IntransFee
            txtInWorker.Text = .InWorkerFee
            txtInEquip.Text = .InEquipFee
            txtOtherDesc.Text = .InOtherDesc
            txtOtherFee.Text = .InOtherFee
        Else
            txtIntrans.Text = maxTrans
            .IntransFee = maxTrans
        End If
        txtWorkerCount.Text = .OutWorkerCount
        If .OutTransFee > 0 Then
            txtOutTrans.Text = .OutTransFee
            txtOutEquip.Text = .OutEquipFee
            txtOutOther.Text = .OutOtherDesc
            txtOutOtherFee.Text = .OutOtherFee
        Else
            Set tmpChargeRate = CalTransFeeFromDistance(curCustomer_.SiteAddresses(cboSiteName.ListIndex + 1).AumphurID, .TSDFAumphurID, .truckTypeID, .WorkDate)
            If Not tmpChargeRate Is Nothing Then
                txtOutTrans.Text = tmpChargeRate.RatePerTrip
            Else
                txtOutTrans.Text = ""
            End If
        End If
        txtOutWorker.Text = .OutWorkerFee
    End With
    Set tmpCol = Nothing
    Set tmpS = Nothing
    SumTreatmentCharge
    ShowDiscrepWasteData
End Sub

Private Sub txtCustTel_Validate(Cancel As Boolean)
    myMenifest_.TelNo = Trim(txtCustTel.Text)
End Sub

Private Sub txtDestER_Validate(Cancel As Boolean)
    myMenifest_.TSDFErContact = txtDestER.Text
End Sub

Private Sub txtDestFax_Validate(Cancel As Boolean)
    myMenifest_.TSDFFaxNo = txtDestFax.Text
End Sub

Private Sub txtDestName_Validate(Cancel As Boolean)
    myMenifest_.SignName3 = txtDestName.Text
End Sub

Private Sub txtDestroyer_Validate(Cancel As Boolean)
    If curTSDF_ Is Nothing Then Set curTSDF_ = New BWGTSDF
    If Trim(txtDestroyer.Text) = "" And Trim(curTSDF_.CompanyName) <> Trim(txtDestroyer.Text) Then
        Set curTSDF_ = Nothing
        Set curTSDF_ = New BWGTSDF
        With myMenifest_
            .TSDF1ID = ""
            .TSDFRegisNo1 = ""
            .TSDFName1 = ""
        End With
    Else
        If Trim(curTSDF_.CompanyName) <> Trim(txtDestroyer.Text) Then
            Set curTSDF_ = Nothing
            Set curTSDF_ = New BWGTSDF
            With myMenifest_
                .TSDF1ID = ""
                .TSDFRegisNo1 = ""
                .TSDFName1 = ""
            End With
            Load frmTSDFSearch
            With frmTSDFSearch
                .getCriteria (txtDestroyer.Text)
                .SetCallerForm Me
                .Show 1
            End With
            If myMenifest_.TSDF1ID = "" Then txtDestroyer.Text = ""
        End If
    End If
End Sub

Private Sub txtDestroyPlace_Validate(Cancel As Boolean)
    myMenifest_.TSDFAddr = txtDestroyPlace.Text
End Sub

Private Sub txtDestTel_Validate(Cancel As Boolean)
    myMenifest_.TSDFTelNo = txtDestTel.Text
End Sub

Private Sub txtDocNo_Validate(Cancel As Boolean)
    myMenifest_.DocNo = txtDocNo.Text
End Sub

Private Sub txtFromProvince_Validate(Cancel As Boolean)
    If curFromProvince_ Is Nothing Then Set curFromProvince_ = New IMIProvince
    If Trim(txtFromProvince.Text) = "" And Trim(curFromProvince_.ProvinceName) <> Trim(txtFromProvince.Text) Then
        Set curFromProvince_ = Nothing
        Set curFromProvince_ = New IMIProvince
        With myMenifest_
            .FromProvinceID = ""
            .FromProvinceName = ""
        End With
    Else
        If Trim(curFromProvince_.ProvinceName) <> Trim(txtFromProvince.Text) Then
            Set curFromProvince_ = Nothing
            Set curFromProvince_ = New IMIProvince
            With myMenifest_
                .FromProvinceID = ""
                .FromProvinceName = ""
            End With
            provLocation_ = "F"
            Load frmProvinceSearch
            With frmProvinceSearch
                .getCriteria (txtFromProvince.Text)
                .SetCallerForm Me
                .Show 1
            End With
            If myMenifest_.FromProvinceID = "" Then txtFromProvince.Text = ""
        End If
    End If
End Sub

Private Sub txtInCollect_Change()
    SumTotalIn
End Sub

Private Sub txtInCollect_Validate(Cancel As Boolean)
    If IsNumeric(Trim(txtInCollect.Text)) Then
        myMenifest_.InCollectFee = Trim(txtInCollect.Text)
    Else
        myMenifest_.InCollectFee = 0
    End If
End Sub

Private Sub txtInEquip_Change()
    SumTotalIn
End Sub

Private Sub txtInEquip_Validate(Cancel As Boolean)
    If IsNumeric(Trim(txtInEquip.Text)) Then
        myMenifest_.InEquipFee = Trim(txtInEquip.Text)
    Else
        myMenifest_.InEquipFee = 0
    End If
End Sub

Private Sub txtIntrans_Change()
    SumTotalIn
End Sub

Private Sub txtIntrans_Validate(Cancel As Boolean)
    If IsNumeric(Trim(txtIntrans.Text)) Then
        myMenifest_.IntransFee = Trim(txtIntrans.Text)
    Else
        myMenifest_.IntransFee = 0
    End If
End Sub

Private Sub txtInWorker_Change()
    SumTotalIn
End Sub

Private Sub txtInWorker_Validate(Cancel As Boolean)
    If IsNumeric(Trim(txtInWorker.Text)) Then
        myMenifest_.InWorkerFee = Trim(txtInWorker.Text)
    Else
        myMenifest_.InWorkerFee = 0
    End If
End Sub

Private Sub txtLiquidQty_Validate(Cancel As Boolean)
    If IsNumeric(txtLiquidQty.Text) Then
        myMenifest_.TotalLiquid = txtLiquidQty.Text
    Else
        myMenifest_.TotalLiquid = 0
    End If
End Sub

Private Sub txtMenifestNo_Validate(Cancel As Boolean)
    myMenifest_.MenifestNo = txtMenifestNo.Text
End Sub

Private Sub txtOtherDesc_Validate(Cancel As Boolean)
    myMenifest_.InOtherDesc = Trim(txtOtherDesc.Text)
End Sub

Private Sub txtOtherFee_Change()
    SumTotalIn
End Sub

Private Sub txtOtherFee_Validate(Cancel As Boolean)
    If IsNumeric(Trim(txtOtherFee.Text)) Then
        myMenifest_.InOtherFee = Trim(txtOtherFee.Text)
    Else
        myMenifest_.InOtherFee = 0
    End If
End Sub

Private Sub txtOutEquip_Change()
    SumTotalOut
End Sub

Private Sub txtOutEquip_Validate(Cancel As Boolean)
    If IsNumeric(Trim(txtOutEquip.Text)) Then
        myMenifest_.OutEquipFee = Trim(txtOutEquip.Text)
    Else
        myMenifest_.OutEquipFee = 0
    End If
End Sub

Private Sub txtOutOther_Validate(Cancel As Boolean)
    myMenifest_.OutOtherDesc = Trim(txtOutOther.Text)
End Sub

Private Sub txtOutOtherFee_Change()
    SumTotalOut
End Sub

Private Sub txtOutOtherFee_Validate(Cancel As Boolean)
    If IsNumeric(Trim(txtOutOtherFee.Text)) Then
        myMenifest_.OutOtherFee = Trim(txtOutOtherFee.Text)
    Else
        myMenifest_.OutOtherFee = 0
    End If
End Sub

Private Sub txtOutTrans_Change()
    SumTotalOut
End Sub

Private Sub txtOutTrans_Validate(Cancel As Boolean)
    If IsNumeric(Trim(txtOutTrans.Text)) Then
        myMenifest_.OutTransFee = Trim(txtOutTrans.Text)
    Else
        myMenifest_.OutTransFee = 0
    End If
End Sub

Private Sub txtOutWorker_Change()
    SumTotalOut
End Sub

Private Sub txtOutWorker_Validate(Cancel As Boolean)
    If IsNumeric(Trim(txtOutWorker.Text)) Then
        myMenifest_.OutWorkerFee = Trim(txtOutWorker.Text)
    Else
        myMenifest_.OutWorkerFee = 0
    End If
End Sub

Private Sub txtRealWeight_Validate(Cancel As Boolean)
    If IsNumeric(txtRealWeight.Text) Then
        myMenifest_.ActualQty = txtRealWeight.Text
    Else
        myMenifest_.ActualQty = 0
    End If
End Sub

Private Sub txtSign1B_Validate(Cancel As Boolean)
    myMenifest_.SignName2 = txtSign1B.Text
End Sub

Private Sub txtSigneName1_Validate(Cancel As Boolean)
    myMenifest_.SignName1 = txtSigneName1.Text
End Sub

Private Sub txtSolidQty_Validate(Cancel As Boolean)
    If IsNumeric(txtSolidQty.Text) Then
        myMenifest_.TotalSolid = txtSolidQty.Text
    Else
        myMenifest_.TotalSolid = 0
    End If
End Sub

Private Sub txtToProvince_Validate(Cancel As Boolean)
    If curToProvince_ Is Nothing Then Set curToProvince_ = New IMIProvince
    If Trim(txtToProvince.Text) = "" And Trim(curToProvince_.ProvinceName) <> Trim(txtToProvince.Text) Then
        Set curToProvince_ = Nothing
        Set curToProvince_ = New IMIProvince
        With myMenifest_
            .ToProvinceID = ""
            .ToProvinceName = ""
        End With
    Else
        If Trim(curToProvince_.ProvinceName) <> Trim(txtToProvince.Text) Then
            Set curToProvince_ = Nothing
            Set curToProvince_ = New IMIProvince
            With myMenifest_
                .ToProvinceID = ""
                .ToProvinceName = ""
            End With
            provLocation_ = "T"
            Load frmProvinceSearch
            With frmProvinceSearch
                .getCriteria (txtToProvince.Text)
                .SetCallerForm Me
                .Show 1
            End With
            If myMenifest_.ToProvinceID = "" Then txtToProvince.Text = ""
        End If
    End If
End Sub

Private Sub txtTrans2_Validate(Cancel As Boolean)
    If curTrans2_ Is Nothing Then Set curTrans2_ = New BWGTransportVendor
    If Trim(txtTrans2.Text) = "" And Trim(curTrans2_.NameTH) <> Trim(txtTrans2.Text) Then
        Set curTrans2_ = Nothing
        Set curTrans2_ = New BWGTransportVendor
        With myMenifest_
            .Transport2ID = ""
            .TransportName2 = ""
            .TransRegisNo2 = ""
        End With
    Else
        If Trim(curTrans2_.NameTH) <> Trim(txtTrans2.Text) Then
            Set curTrans2_ = Nothing
            Set curTrans2_ = New BWGTransportVendor
            With myMenifest_
                .Transport2ID = ""
                .TransportName2 = ""
                .TransRegisNo2 = ""
            End With
            Load frmTransportVenderSearch
            With frmTransportVenderSearch
                .getCriteria (txtTrans2.Text)
                .SetCallerForm Me
                .Show 1
            End With
            If myMenifest_.Transport2ID = "" Then txtTrans2.Text = ""
        End If
    End If
End Sub

Private Sub txtTransER_Validate(Cancel As Boolean)
    myMenifest_.TransErContact1 = txtTransER.Text
End Sub

Private Sub txtTransFax_Validate(Cancel As Boolean)
    myMenifest_.TransFaxNo1 = txtTransFax.Text
End Sub

Private Sub txtTransTel_Validate(Cancel As Boolean)
    myMenifest_.TransTelNo1 = txtTransTel.Text
End Sub

Private Sub txtTransTime_Validate(Cancel As Boolean)
    If IsNumeric(txtTransTime.Text) Then
        myMenifest_.TimeSpending = txtTransTime.Text
    Else
        myMenifest_.TimeSpending = 0
    End If
End Sub

Private Sub txtTreatment_Validate(Cancel As Boolean)
    myMenifest_.SpecialComment = txtTreatment.Text
End Sub

Private Sub txtTreatTime_Validate(Cancel As Boolean)
    If IsNumeric(txtTreatTime.Text) Then
        myMenifest_.TreatmentPeriod = txtTreatTime.Text
    Else
        myMenifest_.TreatmentPeriod = 0
    End If
End Sub

Private Sub txtTruckType_Validate(Cancel As Boolean)
    If curTruckSubType_ Is Nothing Then Set curTruckSubType_ = New BWGTruckSubType
    If Trim(txtTruckType.Text) = "" And Trim(curTruckSubType_.SubTypeDesc & ", " & curTruckSubType_.TypeDesc) <> Trim(txtTruckType.Text) Then
        Set curTruckSubType_ = Nothing
        Set curTruckSubType_ = New BWGTruckSubType
        With myMenifest_
            .TruckSubTypeID = ""
            .TruckSubTypeName = ""
            .truckTypeID = ""
            .TruckTypeName = ""
        End With
    Else
        If Trim(curTruckSubType_.SubTypeDesc & ", " & curTruckSubType_.TypeDesc) <> Trim(txtTruckType.Text) Then
            Set curTruckSubType_ = Nothing
            Set curTruckSubType_ = New BWGTruckSubType
            With myMenifest_
                .TruckSubTypeID = ""
                .TruckSubTypeName = ""
                .truckTypeID = ""
                .TruckTypeName = ""
            End With
            Load frmCarTypeSearch
            With frmCarTypeSearch
                .getCriteria (txtTruckType.Text)
                .SetCallerForm Me
                .Show 1
            End With
            If myMenifest_.TruckSubTypeID = "" Then txtTruckType.Text = ""
        End If
    End If
End Sub

Private Sub txtTumbol_Validate(Cancel As Boolean)
    If curTSDFTumbol_ Is Nothing Then Set curTSDFTumbol_ = New IMITumbol
    If Trim(txtTumbol.Text) = "" And Trim(curTSDFTumbol_.Name) <> Trim(txtTumbol.Text) Then
        Set curTSDFTumbol_ = Nothing
        Set curTSDFTumbol_ = New IMITumbol
        txtAumphur.Text = ""
        With myMenifest_
            .TSDFTumbolID = ""
            .TSDFTumbolName = ""
            .TSDFAumphurID = ""
            .TSDFAumphurName = ""
            .TSDFProvinceName = ""
        End With
    Else
        If Trim(curTSDFTumbol_.Name) <> Trim(txtTumbol.Text) Then
            Set curTSDFTumbol_ = Nothing
            Set curTSDFTumbol_ = New IMITumbol
            txtAumphur.Text = ""
            With myMenifest_
                .TSDFTumbolID = ""
                .TSDFTumbolName = ""
                .TSDFAumphurID = ""
                .TSDFAumphurName = ""
                .TSDFProvinceName = ""
            End With
            Load frmTumbolSearch
            With frmTumbolSearch
                .getCriteria (txtTumbol.Text)
                .SetCallerForm Me
                .Show 1
            End With
            If myMenifest_.TSDFTumbolID = "" Then txtTumbol.Text = ""
        End If
    End If
End Sub

Private Sub txtWorkerCount_Change()
    SumTotalOut
End Sub

Private Sub txtWorkerCount_Validate(Cancel As Boolean)
    If IsNumeric(Trim(txtWorkerCount.Text)) Then
        myMenifest_.OutWorkerCount = Trim(txtWorkerCount.Text)
    Else
        myMenifest_.OutWorkerCount = 0
    End If
End Sub

Private Sub SumTreatmentCharge()
Dim x%
Dim tmpAmt As Double
    For x = 1 To myMenifest_.JobDetails.Count
        tmpAmt = tmpAmt + myMenifest_.JobDetails(x).TreatmentCharge
    Next
    txtInCollect.Text = Format(tmpAmt, "#,##0.00")
    myMenifest_.InCollectFee = txtInCollect.Text
End Sub

Public Sub SetDataTransportVender(selVendor As BWGTransportVendor, Optional isAssigned As Boolean = True)
    Set curTrans2_ = selVendor
    With curTrans2_
        txtTrans2.Text = .NameTH
        txtTransID2.Text = .TransLicenseNo
        txtTrans1B.Text = .NameTH
        txtTransID1B.Text = .TransLicenseNo
        If isAssigned Then
            myMenifest_.Transport2ID = .id
            myMenifest_.TransportName2 = .NameTH
            myMenifest_.TransRegisNo2 = .TransLicenseNo
        End If
    End With
End Sub

Private Sub SetCancelScreen(cFlag As Boolean)
    Command1.Enabled = Not cFlag
    cmdSave.Enabled = Not cFlag
    cmdCancel.Enabled = Not cFlag
    lbCancel.Visible = cFlag
End Sub

Private Sub ShowDiscrepWasteData()
Dim x%
Dim tmpDiscrep As BWGDiscrepWaste
Dim tmpList As ListItem
Dim tmpName As String
    lvResult.ListItems.Clear
    For x = 1 To myMenifest_.DiscrepWastes.Count
        Set tmpDiscrep = myMenifest_.DiscrepWastes(x)
        If Trim(tmpDiscrep.NewWasteName) = "" Then
            tmpName = tmpDiscrep.DiscrepDesc
        Else
            tmpName = tmpDiscrep.NewWasteName & " [" & tmpDiscrep.NewWasteNo & "]"
        End If
        Set tmpList = lvResult.ListItems.Add(, "'" & tmpDiscrep.id & "'", tmpName)
        tmpList.SubItems(1) = tmpDiscrep.DiscrepQty & " " & tmpDiscrep.QtyUnitName
        If tmpDiscrep.ActionTaken = "A" Then
            tmpList.SubItems(2) = "ส่งคืน"
        ElseIf tmpDiscrep.ActionTaken = "B" Then
            tmpList.SubItems(2) = "จัดประเภทใหม่"
        ElseIf tmpDiscrep.ActionTaken = "C" Then
            tmpList.SubItems(2) = "รับกำจัด"
        End If
        Set tmpList = Nothing
        Set tmpDiscrep = Nothing
    Next
End Sub

Public Sub SetDataTruckSubType(selSubType As BWGTruckSubType, Optional isTrailer As Boolean = True)
    Set curTruckSubType_ = Nothing
    Set curTruckSubType_ = selSubType
    With curTruckSubType_
        txtTruckType.Text = .SubTypeDesc & ", " & .TypeDesc
        'If isAssigned Then
            myMenifest_.TruckSubTypeID = .id
            myMenifest_.TruckSubTypeName = .SubTypeDesc
            myMenifest_.truckTypeID = .truckTypeID
            myMenifest_.TruckTypeName = .TypeDesc
        'End If
    End With
End Sub

Public Sub SetDataProvince(selProvince As IMIProvince, Optional isAssigned As Boolean = True)
    If provLocation_ = "F" Then
        Set curFromProvince_ = selProvince
        With curFromProvince_
            txtFromProvince.Text = .ProvinceName
            If isAssigned Then
                myMenifest_.FromProvinceID = .id
                myMenifest_.FromProvinceName = .ProvinceName
            End If
        End With
    ElseIf provLocation_ = "T" Then
        Set curToProvince_ = selProvince
        With curToProvince_
            txtToProvince.Text = .ProvinceName
            If isAssigned Then
                myMenifest_.ToProvinceID = .id
                myMenifest_.ToProvinceName = .ProvinceName
            End If
        End With
    End If
End Sub

Public Sub SetDataTumbol(selTumbol As IMITumbol, Optional isAssigned As Boolean = True)
    Set curTSDFTumbol_ = Nothing
    Set curTSDFTumbol_ = selTumbol
    If curTSDFTumbol_ Is Nothing Then Exit Sub
    With curTSDFTumbol_
        txtTumbol.Text = .Name
        txtAumphur.Text = .AumphurName & " " & .ProvinceName
        If isAssigned Then
            myMenifest_.TSDFAumphurID = .AumphurID
            myMenifest_.TSDFAumphurName = .AumphurName
            myMenifest_.TSDFTumbolID = .id
            myMenifest_.TSDFTumbolName = .Name
            myMenifest_.TSDFProvinceName = .ProvinceName
        End If
    End With
End Sub

Private Sub SumTotalIn()
Dim tmpAmt1 As Double, tmpAmt2 As Double, tmpAmt3 As Double, tmpAmt4 As Double, tmpAmt5 As Double
    If IsNumeric(txtInCollect.Text) Then
        tmpAmt1 = txtInCollect.Text
    Else
        tmpAmt1 = 0
    End If
    If IsNumeric(txtIntrans.Text) Then
        tmpAmt2 = txtIntrans.Text
    Else
        tmpAmt2 = 0
    End If
    If IsNumeric(txtInWorker.Text) Then
        tmpAmt3 = txtInWorker.Text
    Else
        tmpAmt3 = 0
    End If
    If IsNumeric(txtInEquip.Text) Then
        tmpAmt4 = txtInEquip.Text
    Else
        tmpAmt4 = 0
    End If
    If IsNumeric(txtOtherFee.Text) Then
        tmpAmt5 = txtOtherFee.Text
    Else
        tmpAmt5 = 0
    End If
    txtTotalIn.Text = Format(tmpAmt1 + tmpAmt2 + tmpAmt3 + tmpAmt4 + tmpAmt5, "#,##0.00")
End Sub

Private Sub SumTotalOut()
Dim tmpAmt1 As Double, tmpAmt2 As Double, tmpAmt3 As Double, tmpAmt4 As Double, tmpAmt5 As Double
    If IsNumeric(txtOutTrans.Text) Then
        tmpAmt1 = txtOutTrans.Text
    Else
        tmpAmt1 = 0
    End If
    If IsNumeric(txtWorkerCount.Text) Then
        tmpAmt2 = txtWorkerCount.Text
    Else
        tmpAmt2 = 0
    End If
    If IsNumeric(txtOutWorker.Text) Then
        tmpAmt3 = txtOutWorker.Text
    Else
        tmpAmt3 = 0
    End If
    If IsNumeric(txtOutEquip.Text) Then
        tmpAmt4 = txtOutEquip.Text
    Else
        tmpAmt4 = 0
    End If
    If IsNumeric(txtOutOtherFee.Text) Then
        tmpAmt5 = txtOutOtherFee.Text
    Else
        tmpAmt5 = 0
    End If
    txtTotalOut.Text = Format(tmpAmt1 + (tmpAmt2 * tmpAmt3) + tmpAmt4 + tmpAmt5, "#,##0.00")
End Sub

Public Sub SetDataTSDF(selTsdf As BWGTSDF, Optional isAssigned As Boolean = True)
    Set curTSDF_ = selTsdf
    With curTSDF_
        txtDestroyer.Text = .CompanyName
        txtDestroyID.Text = .TSDFRegisNo
        If isAssigned Then
            myMenifest_.TSDF1ID = .id
            myMenifest_.TSDFName1 = .CompanyName
            myMenifest_.TSDFRegisNo1 = .TSDFRegisNo
        End If
    End With
End Sub

Private Sub ShowWasteData()
Dim x%, tmpDet As BWGJobDataDetail
    With grItem.getGridObj
        grItem.ClearAllData
        .Rows = myMenifest_.JobDetails.Count + 20
        For x = 1 To myMenifest_.JobDetails.Count
            Set tmpDet = myMenifest_.JobDetails(x)
            .TextMatrix(x, 0) = tmpDet.id
            .TextMatrix(x, 1) = x
            .TextMatrix(x, 2) = tmpDet.WasteName & " [" & tmpDet.WasteQuotationNo & "]"
            .TextMatrix(x, 3) = tmpDet.WasteEUCode
            .TextMatrix(x, 4) = tmpDet.ContainerCount
            .TextMatrix(x, 5) = tmpDet.ContainerTypeDesc
            .TextMatrix(x, 6) = tmpDet.EstWasteQty
            .TextMatrix(x, 7) = tmpDet.DisposerWeight
            If Trim(tmpDet.QuoTreatmentUnitName) = "บาท/ถัง" Then
                .TextMatrix(x, 8) = "ถัง"
            ElseIf Trim(tmpDet.QuoTreatmentUnitName) = "บาท/ลูกบาศก์เมตร" Then
                .TextMatrix(x, 8) = "ลูกบาศก์เมตร"
            Else
                .TextMatrix(x, 8) = "ตัน"
            End If
            Set tmpDet = Nothing
        Next
    End With
End Sub
