VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmTripControl 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ข้อมูลใบคุม - F131"
   ClientHeight    =   8640
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10485
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8640
   ScaleWidth      =   10485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtWasteType 
      Appearance      =   0  'Flat
      BackColor       =   &H00C00000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   5550
      Locked          =   -1  'True
      TabIndex        =   70
      TabStop         =   0   'False
      Top             =   840
      Width           =   1425
   End
   Begin VB.TextBox txtAllFee 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C00000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   6420
      Locked          =   -1  'True
      TabIndex        =   67
      TabStop         =   0   'False
      Top             =   7830
      Width           =   1425
   End
   Begin VB.TextBox txtAllMenifest 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C00000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   2190
      Locked          =   -1  'True
      TabIndex        =   64
      TabStop         =   0   'False
      Top             =   7830
      Width           =   1425
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5985
      Left            =   0
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1800
      Width           =   10425
      _ExtentX        =   18389
      _ExtentY        =   10557
      _Version        =   393216
      Style           =   1
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "ค่าบริการที่เรียกเก็บลูกค้า"
      TabPicture(0)   =   "frTripControl.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Line1(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(2)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label2(10)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label2(9)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label2(7)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label2(6)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label2(5)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label2(4)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label2(3)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label2(1)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label2(0)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label2(2)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label2(38)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label2(39)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label2(40)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "grItem"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtOtherFee"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtOtherDesc"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtEquip"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txtMoreWorkerFee"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "txtMoreWorker"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "txtTransFee"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "txtCalTransFee"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "txtCalStaffCount"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "txtIncTransFee"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).ControlCount=   25
      TabCaption(1)   =   "ค่าใช้จ่ายที่ต้องจ่ายให้ Vendor"
      TabPicture(1)   =   "frTripControl.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtAllVendor"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "txtVDTransFee"
      Tab(1).Control(2)=   "txtVDWorker"
      Tab(1).Control(3)=   "txtVDWorkerFee"
      Tab(1).Control(4)=   "txtVDEquipFee"
      Tab(1).Control(5)=   "txtVDOtherDesc"
      Tab(1).Control(6)=   "txtVDOtherFee"
      Tab(1).Control(7)=   "LbDream"
      Tab(1).Control(8)=   "LbAllowance"
      Tab(1).Control(9)=   "Label2(32)"
      Tab(1).Control(10)=   "Label2(31)"
      Tab(1).Control(11)=   "Line1(4)"
      Tab(1).Control(12)=   "Label2(30)"
      Tab(1).Control(13)=   "Label2(29)"
      Tab(1).Control(14)=   "Label2(28)"
      Tab(1).Control(15)=   "Label2(27)"
      Tab(1).Control(16)=   "Label2(26)"
      Tab(1).Control(17)=   "Label2(25)"
      Tab(1).Control(18)=   "Label2(24)"
      Tab(1).Control(19)=   "Label2(23)"
      Tab(1).Control(20)=   "Label2(22)"
      Tab(1).Control(21)=   "Label2(21)"
      Tab(1).ControlCount=   22
      TabCaption(2)   =   "ค่าใช้จ่ายอื่นๆ"
      TabPicture(2)   =   "frTripControl.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtBagPrice"
      Tab(2).Control(1)=   "txtTankPrice"
      Tab(2).Control(2)=   "txtBagCount"
      Tab(2).Control(3)=   "txtTankCount"
      Tab(2).Control(4)=   "txtTankPriceS"
      Tab(2).Control(5)=   "txtTankTypeSCount"
      Tab(2).Control(6)=   "Label2(51)"
      Tab(2).Control(7)=   "Label2(50)"
      Tab(2).Control(8)=   "Label2(49)"
      Tab(2).Control(9)=   "Label2(48)"
      Tab(2).Control(10)=   "Label2(43)"
      Tab(2).Control(11)=   "Label2(41)"
      Tab(2).ControlCount=   12
      Begin VB.TextBox txtBagPrice 
         BackColor       =   &H80000003&
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   -72720
         TabIndex        =   86
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox txtTankPrice 
         BackColor       =   &H80000003&
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   -72720
         TabIndex        =   85
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox txtBagCount 
         BackColor       =   &H80000003&
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   -73680
         TabIndex        =   84
         Top             =   1680
         Width           =   855
      End
      Begin VB.TextBox txtTankCount 
         BackColor       =   &H80000003&
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   -73680
         TabIndex        =   83
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox txtTankPriceS 
         BackColor       =   &H80000003&
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   -72720
         TabIndex        =   82
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox txtTankTypeSCount 
         BackColor       =   &H80000003&
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   -73680
         TabIndex        =   81
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox txtIncTransFee 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   960
         TabIndex        =   72
         Top             =   1260
         Width           =   975
      End
      Begin VB.TextBox txtCalStaffCount 
         Appearance      =   0  'Flat
         BackColor       =   &H00C00000&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   900
         Width           =   735
      End
      Begin VB.TextBox txtCalTransFee 
         Appearance      =   0  'Flat
         BackColor       =   &H00C00000&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   900
         Width           =   975
      End
      Begin VB.TextBox txtAllVendor 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C00000&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   -68310
         Locked          =   -1  'True
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   1440
         Width           =   1425
      End
      Begin VB.TextBox txtVDTransFee 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   -73680
         TabIndex        =   21
         Top             =   510
         Width           =   945
      End
      Begin VB.TextBox txtVDWorker 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   -71250
         TabIndex        =   22
         Top             =   510
         Width           =   705
      End
      Begin VB.TextBox txtVDWorkerFee 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   -67860
         TabIndex        =   23
         Top             =   510
         Width           =   945
      End
      Begin VB.TextBox txtVDEquipFee 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   -73680
         TabIndex        =   24
         Top             =   870
         Width           =   945
      End
      Begin VB.TextBox txtVDOtherDesc 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   -71250
         TabIndex        =   25
         Top             =   870
         Width           =   3345
      End
      Begin VB.TextBox txtVDOtherFee 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   -67860
         TabIndex        =   26
         Top             =   870
         Width           =   945
      End
      Begin VB.TextBox txtTransFee 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   960
         TabIndex        =   12
         Top             =   540
         Width           =   975
      End
      Begin VB.TextBox txtMoreWorker 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   3360
         TabIndex        =   13
         Top             =   540
         Width           =   735
      End
      Begin VB.TextBox txtMoreWorkerFee 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   5730
         TabIndex        =   14
         Top             =   540
         Width           =   945
      End
      Begin VB.TextBox txtEquip 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   8940
         TabIndex        =   15
         Top             =   540
         Width           =   945
      End
      Begin VB.TextBox txtOtherDesc 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   5730
         TabIndex        =   18
         Top             =   900
         Width           =   3195
      End
      Begin VB.TextBox txtOtherFee 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   8940
         TabIndex        =   19
         Top             =   900
         Width           =   945
      End
      Begin WasteManagment.ctlGrid grItem 
         Height          =   4185
         Left            =   60
         TabIndex        =   20
         Top             =   1710
         Width           =   10275
         _ExtentX        =   18124
         _ExtentY        =   7382
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "บาท"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   51
         Left            =   -71520
         TabIndex        =   90
         Top             =   1200
         Width           =   285
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ถังปากจู๋  :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   50
         Left            =   -74505
         TabIndex        =   89
         Top             =   1200
         Width           =   690
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "บาท"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   49
         Left            =   -71520
         TabIndex        =   88
         Top             =   1680
         Width           =   285
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "บาท"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   48
         Left            =   -71520
         TabIndex        =   87
         Top             =   720
         Width           =   285
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ถุงบิ๊กแบ็ค  :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   43
         Left            =   -74760
         TabIndex        =   80
         Top             =   1680
         Width           =   945
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ถังปากกว้าง  :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   41
         Left            =   -74760
         TabIndex        =   78
         Top             =   720
         Width           =   950
      End
      Begin VB.Label LbDream 
         Caption         =   "."
         Height          =   375
         Left            =   -66060
         TabIndex        =   77
         Top             =   900
         Width           =   1335
      End
      Begin VB.Label LbAllowance 
         Caption         =   "."
         Height          =   315
         Left            =   -66060
         TabIndex        =   76
         Top             =   540
         Width           =   975
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "คน"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   40
         Left            =   4140
         TabIndex        =   75
         Top             =   960
         Width           =   255
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "คนงานทั้งหมด"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   39
         Left            =   2280
         TabIndex        =   74
         Top             =   960
         Width           =   1035
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ค่าขนส่งแฝง"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   38
         Left            =   0
         TabIndex        =   73
         Top             =   1320
         Width           =   915
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "บาท"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   32
         Left            =   -66840
         TabIndex        =   63
         Top             =   1500
         Width           =   285
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "รวมต้องจ่าย Vendor"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   31
         Left            =   -69810
         TabIndex        =   62
         Top             =   1500
         Width           =   1410
      End
      Begin VB.Line Line1 
         Index           =   4
         X1              =   -74970
         X2              =   -64620
         Y1              =   1350
         Y2              =   1350
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "บาท"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   30
         Left            =   -72690
         TabIndex        =   61
         Top             =   570
         Width           =   285
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ค่าขนส่ง"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   29
         Left            =   -74310
         TabIndex        =   60
         Top             =   570
         Width           =   555
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "คน"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   28
         Left            =   -70500
         TabIndex        =   59
         Top             =   570
         Width           =   225
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "คนงาน"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   27
         Left            =   -71820
         TabIndex        =   58
         Top             =   570
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "บาท"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   26
         Left            =   -66870
         TabIndex        =   57
         Top             =   570
         Width           =   285
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "รวมค่าคนงาน"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   25
         Left            =   -68910
         TabIndex        =   56
         Top             =   570
         Width           =   960
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "บาท"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   24
         Left            =   -72690
         TabIndex        =   55
         Top             =   930
         Width           =   285
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ค่าอุปกรณ์"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   23
         Left            =   -74460
         TabIndex        =   54
         Top             =   930
         Width           =   705
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ค่าอื่นๆ"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   22
         Left            =   -71790
         TabIndex        =   53
         Top             =   930
         Width           =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "บาท"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   21
         Left            =   -66870
         TabIndex        =   52
         Top             =   930
         Width           =   285
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "บาท"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   1980
         TabIndex        =   51
         Top             =   600
         Width           =   315
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ค่าขนส่ง"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   300
         TabIndex        =   50
         Top             =   600
         Width           =   555
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "คน"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   4140
         TabIndex        =   49
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "คนงานเพิ่ม"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   3
         Left            =   2520
         TabIndex        =   48
         Top             =   600
         Width           =   795
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "บาท"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   4
         Left            =   6720
         TabIndex        =   47
         Top             =   600
         Width           =   285
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ค่าคนงานเพิ่ม"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   5
         Left            =   4650
         TabIndex        =   46
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "บาท"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   6
         Left            =   9930
         TabIndex        =   45
         Top             =   600
         Width           =   285
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ค่าอุปกรณ์"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   7
         Left            =   8160
         TabIndex        =   44
         Top             =   600
         Width           =   705
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ค่าอื่นๆ"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   9
         Left            =   5190
         TabIndex        =   43
         Top             =   960
         Width           =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "บาท"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   10
         Left            =   9930
         TabIndex        =   42
         Top             =   960
         Width           =   285
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Manifest ที่ขนและค่ากำจัดกาก"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   60
         TabIndex        =   41
         Top             =   1440
         Width           =   10305
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   2040
         X2              =   12360
         Y1              =   1380
         Y2              =   1380
      End
   End
   Begin WasteManagment.ctlDate dtIn 
      Height          =   345
      Left            =   2520
      TabIndex        =   8
      Top             =   1320
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   609
      BackColor       =   -2147483633
   End
   Begin VB.TextBox txtTimeOut 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   6960
      TabIndex        =   9
      Top             =   1320
      Width           =   705
   End
   Begin VB.TextBox txtTimeIn 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1260
      TabIndex        =   7
      Top             =   1320
      Width           =   705
   End
   Begin VB.TextBox txtDriver 
      Appearance      =   0  'Flat
      BackColor       =   &H00C00000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   7980
      Locked          =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   840
      Width           =   2055
   End
   Begin VB.TextBox txtWorkDate 
      Appearance      =   0  'Flat
      BackColor       =   &H00C00000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1260
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   840
      Width           =   1845
   End
   Begin VB.TextBox txtCarType 
      Appearance      =   0  'Flat
      BackColor       =   &H00C00000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   7980
      Locked          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   450
      Width           =   2055
   End
   Begin VB.TextBox txtCarRegisNo 
      Appearance      =   0  'Flat
      BackColor       =   &H00C00000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   5550
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   450
      Width           =   1425
   End
   Begin VB.TextBox txtVendor 
      Appearance      =   0  'Flat
      BackColor       =   &H00C00000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1260
      Locked          =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   450
      Width           =   3345
   End
   Begin VB.TextBox txtCustomer 
      Appearance      =   0  'Flat
      BackColor       =   &H00C00000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   3750
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   60
      Width           =   6285
   End
   Begin VB.TextBox txtDocNo 
      Appearance      =   0  'Flat
      BackColor       =   &H00C00000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1260
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   60
      Width           =   1845
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "ปิด"
      Height          =   795
      Left            =   9405
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   7800
      Width           =   1035
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "เก็บข้อมูล"
      Height          =   795
      Left            =   8370
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   7800
      Width           =   1005
   End
   Begin WasteManagment.ctlDate dtOut 
      Height          =   345
      Left            =   8250
      TabIndex        =   10
      Top             =   1320
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   609
      BackColor       =   -2147483633
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ถัง :"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   42
      Left            =   120
      TabIndex        =   79
      Top             =   3000
      Width           =   270
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ประเภทเที่ยวที่ขน"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   37
      Left            =   4230
      TabIndex        =   71
      Top             =   900
      Width           =   1245
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "รวมค่าบริการทั้งหมดที่เรียกเก็บ"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   35
      Left            =   4140
      TabIndex        =   69
      Top             =   7890
      Width           =   2190
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "บาท"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   36
      Left            =   7890
      TabIndex        =   68
      Top             =   7890
      Width           =   285
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "รวมค่าบริการตามใบ Menifest"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   33
      Left            =   60
      TabIndex        =   66
      Top             =   7890
      Width           =   2055
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "บาท"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   34
      Left            =   3660
      TabIndex        =   65
      Top             =   7890
      Width           =   285
   End
   Begin VB.Line Line1 
      Index           =   3
      X1              =   60
      X2              =   10620
      Y1              =   1710
      Y2              =   1710
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   0
      X2              =   10560
      Y1              =   1230
      Y2              =   1230
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "วันที่"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   20
      Left            =   7830
      TabIndex        =   40
      Top             =   1380
      Width           =   315
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "วันที่"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   19
      Left            =   2100
      TabIndex        =   39
      Top             =   1380
      Width           =   315
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "เวลาออก"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   17
      Left            =   6270
      TabIndex        =   38
      Top             =   1380
      Width           =   615
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "เวลาเข้า"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   18
      Left            =   630
      TabIndex        =   37
      Top             =   1380
      Width           =   570
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "พนักงานขับ"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   16
      Left            =   7110
      TabIndex        =   36
      Top             =   900
      Width           =   810
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "วันที่ขน"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   15
      Left            =   630
      TabIndex        =   35
      Top             =   900
      Width           =   525
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ประเภทรถ"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   14
      Left            =   7140
      TabIndex        =   34
      Top             =   510
      Width           =   750
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ทะเบียนรถ"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   13
      Left            =   4710
      TabIndex        =   33
      Top             =   510
      Width           =   780
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vendor"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   12
      Left            =   600
      TabIndex        =   32
      Top             =   510
      Width           =   540
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ลูกค้า"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   11
      Left            =   3270
      TabIndex        =   31
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "เลขที่ใบคุม"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   8
      Left            =   420
      TabIndex        =   30
      Top             =   120
      Width           =   765
   End
End
Attribute VB_Name = "frmTripControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim myTripControl_ As BWGJobDataTimeTable
Dim fromAumphurID_ As String
Dim toAumphurID_ As String
Dim TruckTypeID_ As String

Private Type EditType
    EditType As String
    EditbyUserID As String
    EditbyUserName As String
End Type

Private EditType As EditType

Private Sub cmdClose_Click()
    Unload Me
End Sub

Public Sub SaveData()
    CmdSave_Click
End Sub

Private Sub CmdSave_Click()
Dim x%, tmpAmt1 As Double, tmpAmt2 As Double, tmpCost As Double
Dim tmpCon As ADODB.Connection, curID As String
Dim tSql As String, rs As ADODB.Recordset
Dim CurrentDate As String
    If AutoCalcMF = True Then GoTo SaveTrip
    If Trim(txtTransFee.Text) <> "" And IsNumeric(txtTransFee.Text) = False Then
        MsgBox "กรุณาใส่ข้อมูลค่าขนส่งเป็นตัวเลขเท่านั้น", vbExclamation
        txtTransFee.SetFocus
        Exit Sub
    End If
    If Trim(txtMoreWorker.Text) <> "" And IsNumeric(txtMoreWorker.Text) = False Then
        MsgBox "กรุณาใส่ข้อมูลคนงานเพิ่มเป็นตัวเลขเท่านั้น", vbExclamation
        txtMoreWorker.SetFocus
        Exit Sub
    End If
    If Trim(txtMoreWorkerFee.Text) <> "" And IsNumeric(txtMoreWorkerFee.Text) = False Then
        MsgBox "กรุณาใส่ข้อมูลค่าบริการคนงานเพิ่มเป็นตัวเลขเท่านั้น", vbExclamation
        txtMoreWorkerFee.SetFocus
        Exit Sub
    End If
    If Trim(txtEquip.Text) <> "" And IsNumeric(txtEquip.Text) = False Then
        MsgBox "กรุณาใส่ข้อมูลค่าอุปกรณ์เป็นตัวเลขเท่านั้น", vbExclamation
        txtEquip.SetFocus
        Exit Sub
    End If
    If Trim(txtOtherFee.Text) <> "" And IsNumeric(txtOtherFee.Text) = False Then
        MsgBox "กรุณาใส่ข้อมูลค่าบริการอื่นๆ เป็นตัวเลขเท่านั้น", vbExclamation
        txtOtherFee.SetFocus
        Exit Sub
    End If
    If Trim(txtOtherFee.Text) <> "" And Trim(txtOtherDesc.Text) = "" And Val(txtOtherFee.Text) <> 0 Then
        MsgBox "กรุณาระบุรายละเอียดค่าใช้จ่ายอื่นๆ", vbExclamation
        txtOtherDesc.SetFocus
        Exit Sub
    End If
    
SaveTrip:
    If AutoCalcMF = True Then myTripControl_.isChanged = True
    myTripControl_.isChanged = True
    If myTripControl_.isChanged Then

        myTripControl_.Save "", AutoCalcMF
        With grItem.getGridObj
            curID = Trim(.TextMatrix(1, 7))
            For x = 1 To .Rows - 1
                If Trim(curID) = Trim(.TextMatrix(x, 7)) Then
        'คิดราคา IncollectFee
                    If IsNumeric(Trim(.TextMatrix(x, 6))) Then
                        tmpAmt1 = tmpAmt1 + CDbl(Trim(.TextMatrix(x, 6)))
                    End If
                    If IsNumeric(Trim(.TextMatrix(x, 8))) Then
                        tmpCost = tmpCost + CDbl(Trim(.TextMatrix(x, 8)))
                    End If
                Else
'                    DBConnExc "Update tbMenifestHeader Set InCollectFee=" & tmpAmt1 & " WHERE MenifestID='" & Trim(curID) & "'"
                    tSql = "SELECT CONVERT(varchar(6), GETDATE(), 103) + CONVERT(varchar(4), YEAR(GETDATE()) + 543) AS CurrentDate" & vbCrLf
                    Set rs = GetRS(tSql, adOpenForwardOnly, adLockOptimistic)
                    CurrentDate = Format(rs!CurrentDate, "yyyy/MM/dd")
                    If Format(rs!CurrentDate, "yyyy/MM") = Format(Left(myTripControl_.WorkDate, 7), "yyyy/MM") Then
                        If AutoCalcMF = False Then
                            tSql = "UPDATE tbMenifestHeader SET InCollectFee = " & tmpAmt1 & ",SumCost=" & tmpCost & ", DiffFee = " & tmpAmt1 & ", ModifyDate = GetDate(), ModifyUser = '" & Trim(CurrentUser.StaffFName & " " & CurrentUser.StaffLName) & "'  WHERE MenifestID = '" & Trim(curID) & "'"
                        Else
                            tSql = "UPDATE tbMenifestHeader SET InCollectFee = " & tmpAmt1 & ",SumCost=" & tmpCost & ", DiffFee = " & tmpAmt1 & "  WHERE MenifestID = '" & Trim(curID) & "'"
                        End If
                        DBConnExc tSql
                    ElseIf Format(CurrentDate, "yyyy/MM") > Format(Left(myTripControl_.WorkDate, 7), "yyyy/MM") Then
                        Dim cMonth As Integer, wMonth As Integer
                        cMonth = Month(CurrentDate)
                        wMonth = CInt(Format$(Trim(myTripControl_.WorkDate), "MM"))
                        If DateDiff("m", myTripControl_.WorkDate, CurrentDate) = 1 Then
                            If Format(CurrentDate, "dd") > "10" Then
'                                If AutoCalcMF = False Then MsgBox "ถ้ามีการบันทึกรายการใบเมนิเฟส หลังวันที่ 10 ของเดือนถัดไป" & vbCrLf & "รายการจะไม่อัพเดทให้เห็นแต่จะบันทึกไปที่รายการวางวิล และคำนวณค่าคอมฯ ปกติ", vbInformation
                                If AutoCalcMF = False Then
                                    tSql = "IF ((SELECT ISNULL(IncollectFee, 0) FROM tbMenifestHeader WHERE MenifestID = '" & Trim(curID) & "') = 0)" & vbCrLf
                                    tSql = tSql & "BEGIN" & vbCrLf
                                    tSql = tSql & "UPDATE tbMenifestHeader SET InCollectFee = " & tmpAmt1 & ",SumCost=" & tmpCost & " WHERE MenifestID = '" & Trim(curID) & "'" & vbCrLf
                                    tSql = tSql & "UPDATE tbMenifestHeader SET DiffFee = 0, ModifyDate = GetDate(), ModifyUser = '" & Trim(CurrentUser.StaffFName & " " & CurrentUser.StaffLName) & "' WHERE MenifestID = '" & Trim(curID) & "'" & vbCrLf
                                    tSql = tSql & "END" & vbCrLf
                                    tSql = tSql & "ELSE" & vbCrLf
                                    tSql = tSql & "UPDATE tbMenifestHeader SET DiffFee = InCollectFee, InCollectFee = " & tmpAmt1 & ",SumCost=" & tmpCost & ", ModifyDate = GetDate(), ModifyUser = '" & Trim(CurrentUser.StaffFName & " " & CurrentUser.StaffLName) & "' WHERE MenifestID = '" & Trim(curID) & "'" & vbCrLf
                                Else
                                    tSql = "IF ((SELECT ISNULL(IncollectFee, 0) FROM tbMenifestHeader WHERE MenifestID = '" & Trim(curID) & "') = 0)" & vbCrLf
                                    tSql = tSql & "BEGIN" & vbCrLf
                                    tSql = tSql & "UPDATE tbMenifestHeader SET InCollectFee = " & tmpAmt1 & ",SumCost=" & tmpCost & " WHERE MenifestID = '" & Trim(curID) & "'" & vbCrLf
                                    tSql = tSql & "UPDATE tbMenifestHeader SET DiffFee = 0 WHERE MenifestID = '" & Trim(curID) & "'" & vbCrLf
                                    tSql = tSql & "END" & vbCrLf
                                    tSql = tSql & "ELSE" & vbCrLf
                                    tSql = tSql & "UPDATE tbMenifestHeader SET DiffFee = InCollectFee, InCollectFee = " & tmpAmt1 & ",SumCost=" & tmpCost & " WHERE MenifestID = '" & Trim(curID) & "'" & vbCrLf
                                End If
                            Else
                                If AutoCalcMF = False Then
                                    tSql = "UPDATE tbMenifestHeader SET InCollectFee = " & tmpAmt1 & ",SumCost=" & tmpCost & ", DiffFee = " & tmpAmt1 & ", ModifyDate = GetDate(), ModifyUser = '" & Trim(CurrentUser.StaffFName & " " & CurrentUser.StaffLName) & "' WHERE MenifestID = '" & Trim(curID) & "'"
                                Else
                                    tSql = "UPDATE tbMenifestHeader SET InCollectFee = " & tmpAmt1 & ",SumCost=" & tmpCost & ", DiffFee = " & tmpAmt1 & " WHERE MenifestID = '" & Trim(curID) & "'"
                                End If
                            End If
                        Else
'                            If AutoCalcMF = False Then MsgBox "ถ้ามีการบันทึกรายการใบเมนิเฟส เกิน 1 เดือน รายการจะไม่อัพเดทให้เห็นแต่จะบันทึกไปที่รายการวางวิล และคำนวณค่าคอมฯ ปกติ", vbInformation
                            If AutoCalcMF = False Then
                                tSql = "IF ((SELECT ISNULL(IncollectFee, 0) FROM tbMenifestHeader WHERE MenifestID = '" & Trim(curID) & "') = 0)" & vbCrLf
                                tSql = tSql & "BEGIN" & vbCrLf
                                tSql = tSql & "UPDATE tbMenifestHeader SET InCollectFee = " & tmpAmt1 & ",SumCost=" & tmpCost & " WHERE MenifestID = '" & Trim(curID) & "'" & vbCrLf
                                tSql = tSql & "UPDATE tbMenifestHeader SET DiffFee = 0, ModifyDate = GetDate(), ModifyUser = '" & Trim(CurrentUser.StaffFName & " " & CurrentUser.StaffLName) & "' WHERE MenifestID = '" & Trim(curID) & "'" & vbCrLf
                                tSql = tSql & "END" & vbCrLf
                                tSql = tSql & "ELSE" & vbCrLf
                                tSql = tSql & "UPDATE tbMenifestHeader SET DiffFee = InCollectFee, InCollectFee = " & tmpAmt1 & ",SumCost=" & tmpCost & ", ModifyDate = GetDate(), ModifyUser = '" & Trim(CurrentUser.StaffFName & " " & CurrentUser.StaffLName) & "' WHERE MenifestID = '" & Trim(curID) & "'" & vbCrLf
                            Else
                                tSql = "IF ((SELECT ISNULL(IncollectFee, 0) FROM tbMenifestHeader WHERE MenifestID = '" & Trim(curID) & "') = 0)" & vbCrLf
                                tSql = tSql & "BEGIN" & vbCrLf
                                tSql = tSql & "UPDATE tbMenifestHeader SET InCollectFee = " & tmpAmt1 & ",SumCost=" & tmpCost & " WHERE MenifestID = '" & Trim(curID) & "'" & vbCrLf
                                tSql = tSql & "UPDATE tbMenifestHeader SET DiffFee = 0 WHERE MenifestID = '" & Trim(curID) & "'" & vbCrLf
                                tSql = tSql & "END" & vbCrLf
                                tSql = tSql & "ELSE" & vbCrLf
                                tSql = tSql & "UPDATE tbMenifestHeader SET DiffFee = InCollectFee, InCollectFee = " & tmpAmt1 & ",SumCost=" & tmpCost & " WHERE MenifestID = '" & Trim(curID) & "'" & vbCrLf
                            End If
                        End If
                        DBConnExc tSql
                    End If
                    
                    If IsNumeric(Trim(.TextMatrix(x, 6))) Then
                        tmpAmt1 = CDbl(Trim(.TextMatrix(x, 6)))
                    Else
                        tmpAmt1 = 0
                    End If
                    curID = Trim(.TextMatrix(x, 7))
                End If
            Next
            If Trim(curID) <> "" Then
'                DBConnExc "Update tbMenifestHeader Set InCollectFee=" & tmpAmt1 & " WHERE MenifestID='" & Trim(curID) & "'"
                tSql = "SELECT CONVERT(varchar(6), GETDATE(), 103) + CONVERT(varchar(4), YEAR(GETDATE()) + 543) AS CurrentDate" & vbCrLf
                Set rs = GetRS(tSql, adOpenForwardOnly, adLockOptimistic)
                CurrentDate = Format(rs!CurrentDate, "yyyy/MM/dd")
                If Format(CurrentDate, "yyyy/MM") = Format(myTripControl_.WorkDate, "yyyy/MM") Then
                    If AutoCalcMF = False Then
                        tSql = "UPDATE tbMenifestHeader SET InCollectFee = " & tmpAmt1 & ",SumCost=" & tmpCost & ", DiffFee = " & tmpAmt1 & ", ModifyDate = GetDate(), ModifyUser = '" & Trim(CurrentUser.StaffFName & " " & CurrentUser.StaffLName) & "' WHERE MenifestID = '" & Trim(curID) & "'"
                    Else
                        tSql = "UPDATE tbMenifestHeader SET InCollectFee = " & tmpAmt1 & ",SumCost=" & tmpCost & ", DiffFee = " & tmpAmt1 & " WHERE MenifestID = '" & Trim(curID) & "'"
                    End If
                    DBConnExc tSql
                ElseIf Format(CurrentDate, "yyyy/MM") > Format(myTripControl_.WorkDate, "yyyy/MM") Then
                    If DateDiff("m", myTripControl_.WorkDate, CurrentDate) = 1 Then
                        If Format(CurrentDate, "dd") > "10" Then
'                            If AutoCalcMF = False Then MsgBox "ถ้ามีการบันทึกรายการใบเมนิเฟส หลังวันที่ 10 ของเดือนถัดไป" & vbCrLf & "รายการจะไม่อัพเดทให้เห็นแต่จะบันทึกไปที่รายการวางวิล และคำนวณค่าคอมฯ ปกติ", vbInformation
                            If AutoCalcMF = False Then
                                tSql = "IF ((SELECT ISNULL(IncollectFee, 0) FROM tbMenifestHeader WHERE MenifestID = '" & Trim(curID) & "') = 0)" & vbCrLf
                                tSql = tSql & "BEGIN" & vbCrLf
                                tSql = tSql & "UPDATE tbMenifestHeader SET InCollectFee = " & tmpAmt1 & ",SumCost=" & tmpCost & " WHERE MenifestID = '" & Trim(curID) & "'" & vbCrLf
                                tSql = tSql & "UPDATE tbMenifestHeader SET DiffFee = 0, ModifyDate = GetDate(), ModifyUser = '" & Trim(CurrentUser.StaffFName & " " & CurrentUser.StaffLName) & "' WHERE MenifestID = '" & Trim(curID) & "'" & vbCrLf
                                tSql = tSql & "END" & vbCrLf
                                tSql = tSql & "ELSE" & vbCrLf
                                tSql = tSql & "UPDATE tbMenifestHeader SET DiffFee = InCollectFee, InCollectFee = " & tmpAmt1 & ",SumCost=" & tmpCost & ", ModifyDate = GetDate(), ModifyUser = '" & Trim(CurrentUser.StaffFName & " " & CurrentUser.StaffLName) & "' WHERE MenifestID = '" & Trim(curID) & "'" & vbCrLf
                            Else
                                tSql = "IF ((SELECT ISNULL(IncollectFee, 0) FROM tbMenifestHeader WHERE MenifestID = '" & Trim(curID) & "') = 0)" & vbCrLf
                                tSql = tSql & "BEGIN" & vbCrLf
                                tSql = tSql & "UPDATE tbMenifestHeader SET InCollectFee = " & tmpAmt1 & ",SumCost=" & tmpCost & " WHERE MenifestID = '" & Trim(curID) & "'" & vbCrLf
                                tSql = tSql & "UPDATE tbMenifestHeader SET DiffFee = 0 WHERE MenifestID = '" & Trim(curID) & "'" & vbCrLf
                                tSql = tSql & "END" & vbCrLf
                                tSql = tSql & "ELSE" & vbCrLf
                                tSql = tSql & "UPDATE tbMenifestHeader SET DiffFee = InCollectFee, InCollectFee = " & tmpAmt1 & ",SumCost=" & tmpCost & " WHERE MenifestID = '" & Trim(curID) & "'" & vbCrLf
                            End If
                        Else
                            If AutoCalcMF = False Then
                                tSql = "UPDATE tbMenifestHeader SET InCollectFee = " & tmpAmt1 & ",SumCost=" & tmpCost & ", DiffFee = " & tmpAmt1 & ", ModifyDate = GetDate(), ModifyUser = '" & Trim(CurrentUser.StaffFName & " " & CurrentUser.StaffLName) & "' WHERE MenifestID = '" & Trim(curID) & "'"
                            Else
                                tSql = "UPDATE tbMenifestHeader SET InCollectFee = " & tmpAmt1 & ",SumCost=" & tmpCost & ", DiffFee = " & tmpAmt1 & " WHERE MenifestID = '" & Trim(curID) & "'"
                            End If
                        End If
                    Else
'                        If AutoCalcMF = False Then MsgBox "ถ้ามีการบันทึกรายการใบเมนิเฟส เกิน 1 เดือน รายการจะไม่อัพเดทให้เห็นแต่จะบันทึกไปที่รายการวางวิล และคำนวณค่าคอมฯ ปกติ", vbInformation
                        If AutoCalcMF = False Then
                            tSql = "IF ((SELECT ISNULL(IncollectFee, 0) FROM tbMenifestHeader WHERE MenifestID = '" & Trim(curID) & "') = 0)" & vbCrLf
                            tSql = tSql & "BEGIN" & vbCrLf
                            tSql = tSql & "UPDATE tbMenifestHeader SET InCollectFee = " & tmpAmt1 & ",SumCost=" & tmpCost & " WHERE MenifestID = '" & Trim(curID) & "'" & vbCrLf
                            tSql = tSql & "UPDATE tbMenifestHeader SET DiffFee = 0, ModifyDate = GetDate(), ModifyUser = '" & Trim(CurrentUser.StaffFName & " " & CurrentUser.StaffLName) & "' WHERE MenifestID = '" & Trim(curID) & "'" & vbCrLf
                            tSql = tSql & "END" & vbCrLf
                            tSql = tSql & "ELSE" & vbCrLf
                            tSql = tSql & "UPDATE tbMenifestHeader SET DiffFee = InCollectFee, InCollectFee = " & tmpAmt1 & ",SumCost=" & tmpCost & ", ModifyDate = GetDate(), ModifyUser = '" & Trim(CurrentUser.StaffFName & " " & CurrentUser.StaffLName) & "' WHERE MenifestID = '" & Trim(curID) & "'" & vbCrLf
                        Else
                            tSql = "IF ((SELECT ISNULL(IncollectFee, 0) FROM tbMenifestHeader WHERE MenifestID = '" & Trim(curID) & "') = 0)" & vbCrLf
                            tSql = tSql & "BEGIN" & vbCrLf
                            tSql = tSql & "UPDATE tbMenifestHeader SET InCollectFee = " & tmpAmt1 & ",SumCost=" & tmpCost & " WHERE MenifestID = '" & Trim(curID) & "'" & vbCrLf
                            tSql = tSql & "UPDATE tbMenifestHeader SET DiffFee = 0 WHERE MenifestID = '" & Trim(curID) & "'" & vbCrLf
                            tSql = tSql & "END" & vbCrLf
                            tSql = tSql & "ELSE" & vbCrLf
                            tSql = tSql & "UPDATE tbMenifestHeader SET DiffFee = InCollectFee, InCollectFee = " & tmpAmt1 & ",SumCost=" & tmpCost & " WHERE MenifestID = '" & Trim(curID) & "'" & vbCrLf
                        End If
                    End If
                    DBConnExc tSql
                End If
            End If
        End With
        Set tmpCon = Nothing
        Set rs = Nothing
    End If
    
    If myTripControl_.isChanged = False Then
        If AutoCalcMF = False Then
            MsgBox "จัดเก็บข้อมูลเรียบร้อยแล้ว", vbInformation
        End If
        Unload Me
    End If
End Sub

Private Sub dtIn_Validate(Cancel As Boolean)
    myTripControl_.CustInDate = dtIn.ValueYMD
End Sub

Private Sub dtOut_Validate(Cancel As Boolean)
    myTripControl_.CustOutDate = dtOut.ValueYMD
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{Tab}"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    With grItem.getGridObj
        .Cols = 9
        .Rows = 50
        .TextMatrix(0, 1) = "รายการ"
        .TextMatrix(0, 2) = "จำนวน"
        .TextMatrix(0, 3) = "หน่วย"
        .TextMatrix(0, 4) = "อัตราค่าบริการ"
        .TextMatrix(0, 5) = "หน่วย"
        .TextMatrix(0, 6) = "รวมเป็นเงิน"
        .TextMatrix(0, 8) = "ต้นทุน"
        .ColWidth(0) = 0
        .ColWidth(1) = 4695
        .ColWidth(2) = 870
        .ColWidth(3) = 915
        .ColWidth(4) = 1200
        .ColWidth(5) = 1095
        .ColWidth(6) = 1245
        .ColWidth(8) = 1245
    End With
'        .TextMatrix(0, 1) = "เลขที่ Manifest"
'        .TextMatrix(0, 2) = "Waste No"
'        .TextMatrix(0, 3) = "ชื่อกาก"
'        .TextMatrix(0, 4) = "ใบเสนอราคา"
'        .TextMatrix(0, 5) = "ค่าขนส่ง"
'        .TextMatrix(0, 5) = "หน่วย"
'        .TextMatrix(0, 5) = "ขั้นต่ำ"
'        .TextMatrix(0, 5) = "ปริมาณ"
End Sub

Public Sub SetTripControl(selControl As BWGJobDataTimeTable, custName As String, CarType As String, CarRegisNo As String, FromAumphurID As String, ToAumphurID As String, TruckTypeID As String)
 Dim x%, tmpItem As BWGMenifest, tmpList As ListItem
Dim tmpCol As Collection, tmpAmt As Double, tmpS As New BWGSearchManager
Dim tmpJobDet As BWGJobDataDetail, tmpCol2 As Collection, tmpAddWaste As BWGMenifestWasteAdded
Dim tripType As String, minType$, maxType$, tmpTripused As BWGJobDataCarUsed, tmpChargeRate As BWGTransVendorChargeRate, tmpChargeRateWT As BWGTransVendorChargeRate
Dim curTreatUnitID$, curMfID$, curRow%, tmpTotalAmt As Double, y%, tmpTranUnit$, nonGroupWeight As Double, maxNonTreatRate As Double, TotalCost As Double
Dim maxNonUnit$, curGroupWeight As Double, tmpWeight As Double, tmpWasteType As String
Dim nonWeight As Double, hazWeight As Double, nblWeight As Double, hblWeight As Double, ColWasteWeight As Collection, tmpWaste As BWGWasteDataCR
Dim MaxWeight As Double, MaxWasteID As String, SumTotalWeight As Double, isUnderWeight As Boolean, colTrip As Collection, z%, minMenifestID$
Dim curUnitTimeTable$, selMID$

Dim curTimeTableNo$, colMinWeightCheck As New Collection, curMenifest$, tmpTimeTableNo$
Dim selMenifestID As String, curTripID As String, ExtraCharge$
Dim TWeight As Double
Dim CurCustomerID As String

    Set myTripControl_ = selControl
    minType = "ZZZ"
    maxType = "AAA"
    fromAumphurID_ = FromAumphurID
    toAumphurID_ = ToAumphurID
    TruckTypeID_ = TruckTypeID
    
    With myTripControl_
        txtDocNo.Text = .TimeTableNo
        txtCustomer.Text = custName
        txtVendor.Text = .VendorName
        If .TrailerTruckNo = "" Then
            txtCarRegisNo.Text = CarRegisNo ' .carRegisNo
        Else
            txtCarRegisNo.Text = CarRegisNo & " - " & .TrailerTruckNo & " (พ่วง)"
        End If
        txtCarType.Text = CarType
        txtWorkDate.Text = FormatYMD_to_DMY(.WorkDate, "/", "/")
        txtDriver.Text = .DriverName
        txtTimeIn.Text = .CustInTime
        If Trim(.CustInDate) <> "" Then
            dtIn.ValueYMD = FormatYMD_to_DMY(.CustInDate, "/", "/")
        End If
        txtTimeOut.Text = .CustOutTime
        If Trim(.CustOutDate) <> "" Then
            dtOut.ValueYMD = FormatYMD_to_DMY(.CustOutDate, "/", "/")
        End If
        
        CalCustChargeFee
        
        If .TripAllFee > 0 Then
            txtTransFee.Text = .TripTranFee
            txtMoreWorker.Text = .TripMoreWorker
            txtMoreWorkerFee.Text = .TripMoreWorkerFee
            txtEquip.Text = .TripEquipFee
            txtOtherDesc.Text = .TripOtherDesc
            txtOtherFee.Text = .TripOtherFee
        End If
                
        Set tmpChargeRate = Nothing
        If .isJobs <> "D" Then
            If .EditTypeTrip <> "M" Then
                Set tmpChargeRate = CalTransFeeFromDistance_Cust(fromAumphurID_, toAumphurID_, TruckTypeID_, .WorkDate)
                If Not tmpChargeRate Is Nothing Then
                    If .isTrailer Then
                        .IncTripTransFee = tmpChargeRate.TrailerRateTrip
                    Else
                        .IncTripTransFee = tmpChargeRate.RatePerTrip
                    End If
                Else
                    .IncTripTransFee = 0
                End If
                txtIncTransFee.Text = Format(.IncTripTransFee, "#,##0")
            Else
                txtIncTransFee.Text = Format(.IncTripTransFee, "#,##0")
            End If
        Else
            .IncTripTransFee = 0
            txtIncTransFee.Text = Format(.IncTripTransFee, "#,##0")
        End If
        
        
        Set tmpChargeRate = Nothing
        
        If .EditTypeVD <> "M" Then
'            If Trim(txtVDTransFee.Text) = "" Or Trim(txtVDTransFee.Text) = "0" Then
                DoEvents
                If .VendorID = "185510566D9D4813907BDE7D216A36FE" Or .VendorID = "C19B2C69A99C422EA28C8C1A33517905" Then  ' ถ้าเป็นรถ ค่าเบี้ยง ดรีม BWT ให้ใช้Rat BWT ทั้งหมด
                    Set tmpChargeRateWT = CalTransFeeFromDistance(fromAumphurID_, "AP2005062715572615", TruckTypeID_, .WorkDate)
                Else
                    Set tmpChargeRateWT = CalTransFeeFromDistance(fromAumphurID_, toAumphurID_, TruckTypeID_, .WorkDate)
                End If

                Set tmpChargeRate = CalTransFeeFromDistance(fromAumphurID_, toAumphurID_, TruckTypeID_, .WorkDate)
                If Not tmpChargeRate Is Nothing Then
                
                    If .isTrailer Then
                        .VDTranFee = tmpChargeRate.TrailerRateTrip
                         If .VendorID = "185510566D9D4813907BDE7D216A36FE" Or .VendorID = "C19B2C69A99C422EA28C8C1A33517905" Then ' กรณี รถ BWT คิดเบี้ยเลี้ยง
                            .Allowance = tmpChargeRateWT.Allowance
                            .Dream = tmpChargeRateWT.Dream
                         Else
                            .Allowance = tmpChargeRate.Allowance
                            .Dream = tmpChargeRate.Dream
                         End If

                    Else
                        .VDTranFee = tmpChargeRate.RatePerTrip
                        If .VendorID = "185510566D9D4813907BDE7D216A36FE" Or .VendorID = "C19B2C69A99C422EA28C8C1A33517905" Then ' กรณี รถ BWT คิดเบี้ยเลี้ยง
                            .Allowance = tmpChargeRateWT.Allowance
                            .Dream = tmpChargeRateWT.Dream
                         Else
                            .Allowance = tmpChargeRate.Allowance
                            .Dream = tmpChargeRate.Dream
                         End If
                    End If
                Else
                    .VDTranFee = 0
                End If
'            End If
        End If
        Set tmpChargeRate = Nothing
        txtVDTransFee.Text = .VDTranFee
        txtVDWorker.Text = .VDWorker
        txtVDWorkerFee.Text = .VDWorkerFee
        txtVDEquipFee.Text = .VDEquipFee
        txtVDOtherDesc.Text = .VDOtherDesc
        txtVDOtherFee.Text = .VDOtherFee
        
        Set tmpTripused = CalJobDataCarUsed(.JobDataCarID)

        .TankType = tmpTripused.TankType
        .Tankcount = tmpTripused.Tankcount
        .TankPrice = tmpTripused.TankPrice
        .TankPriceS = tmpTripused.TankPriceS
        .TankPriceV = tmpTripused.TankPriceV
        .TankTypeB = tmpTripused.TankTypeB
        txtTankCount.Text = .Tankcount
        txtTankPrice.Text = .TankPrice
        .TankTypeS = tmpTripused.TankTypeS
        .TankTypeSCount = tmpTripused.TankTypeSCount
        .TankTypeSB = tmpTripused.TankTypeSB
        txtTankTypeSCount.Text = .TankTypeSCount
        txtTankPriceS.Text = .TankPriceS
        .TankPriceSV = tmpTripused.TankPriceSV
        .BagCount = tmpTripused.BagCount
        .BagPrice = tmpTripused.BagPrice
        .BagPriceV = tmpTripused.BagPriceV
        
        txtBagCount.Text = .BagCount
        txtBagPrice.Text = .BagPrice
        .isChanged = False
    End With
    SumPayVendor
    
'    แก้ส่วนนี้ให้เหมือนกับการพิมพ์ใบกำกับภาษี
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    grItem.ClearAllData
    curRow = 1
    CurCustomerID = ""
    '''''''''''''เริ่มการหาขั้นต่ำ
    Set colTrip = tmpS.JobDetailWasteInTripSearch("TimeTableNo='" & myTripControl_.TimeTableNo & "'", "TimeTableNo,DocNo,WorkDate,WasteType,TreatmentUnitID")
'    Set colTrip = tmpS.JobDetailWasteInTripSearch("MenifestID in (" & myTripControl_.menifestID & ")", "WorkDate,TimeTableNo,DocNo,WasteType,TreatmentUnitID")
    DoEvents
    Set ColWasteWeight = New Collection
    
    If Not colTrip Is Nothing Then
        For x = 1 To colTrip.Count '******** วบเช็ค Waste วนตามใบคุม, ใบเมนิเฟส, ประเภทเวสต์
            
            Set tmpJobDet = colTrip(x)
            If (curTimeTableNo <> tmpJobDet.TimeTableNo) Then
                If Trim(curTimeTableNo) <> "" Then
                    '*** เพิ่มข้อมูลลงใน colMinWeightCheck ในกรณีที่มีใบคุมหลายใบ
                    Call AddMinWeight(colMinWeightCheck, ColWasteWeight, curTimeTableNo)
                End If
                CurCustomerID = tmpJobDet.CustomerID
                curTimeTableNo = tmpJobDet.TimeTableNo
                curMenifest = tmpJobDet.menifestID
            End If
            
            '***** Calculate min rate with waste type ********************
            If LCase(tmpJobDet.WasteType) = "non" Then
                If isExist(ColWasteWeight, "'NON'") Then
                    Set tmpWaste = ColWasteWeight("'NON'")
                    If tmpJobDet.ChargeWeightType = "A" Then
                        tmpWaste.MaxQty = tmpWaste.MaxQty + tmpJobDet.DisposerWeight
                    Else
                        tmpWaste.MaxQty = tmpWaste.MaxQty + tmpJobDet.EstWasteQty
                    End If
                    If tmpWaste.MinWeightPerCar = 0 Then tmpWaste.MinWeightPerCar = tmpJobDet.QuoMinWeightPerCar
                    If tmpWaste.NonMinRate = 0 Then tmpWaste.NonMinRate = tmpJobDet.NonMinRate
                    
                    If tmpWaste.MaxWeightPerCar = 0 Then tmpWaste.MaxWeightPerCar = tmpJobDet.QuoMaxWeightPerCar
                    If tmpWaste.MaxRate = 0 Then tmpWaste.MaxRate = tmpJobDet.QuoMaxRate
                    If tmpWaste.MaxRateType = "" Then tmpWaste.MaxRateType = tmpJobDet.QuoMaxRateType
                    Set tmpWaste = Nothing
                Else
                    Set tmpWaste = New BWGWasteDataCR
                    tmpWaste.id = "NON"
                    tmpWaste.WasteName = "NON"
                    tmpWaste.WasteType = "Non"
                    tmpWaste.menifestID = tmpJobDet.menifestID
                    tmpWaste.TreatmentUnitID = tmpJobDet.QuoTreatmentUnitID
                    tmpWaste.MinWeightPerCar = tmpJobDet.QuoMinWeightPerCar
                    tmpWaste.NonMinRate = tmpJobDet.NonMinRate
                    tmpWaste.MaxWeightPerCar = tmpJobDet.QuoMaxWeightPerCar
                    tmpWaste.MaxRate = tmpJobDet.QuoMaxRate
                    tmpWaste.MaxRateType = tmpJobDet.QuoMaxRateType
                    If tmpJobDet.ChargeWeightType = "A" Then
                        tmpWaste.MaxQty = tmpJobDet.DisposerWeight
                    Else
                        tmpWaste.MaxQty = tmpJobDet.EstWasteQty
                    End If
                    ColWasteWeight.Add tmpWaste, "'NON'"
                    Set tmpWaste = Nothing
                End If
            ElseIf LCase(tmpJobDet.WasteType) = "nbl" Then
                If isExist(ColWasteWeight, "'NBL'") Then
                    Set tmpWaste = ColWasteWeight("'NBL'")
                    If tmpJobDet.ChargeWeightType = "A" Then
                        tmpWaste.MaxQty = tmpWaste.MaxQty + tmpJobDet.DisposerWeight
                    Else
                        tmpWaste.MaxQty = tmpWaste.MaxQty + tmpJobDet.EstWasteQty
                    End If
                    If tmpWaste.MinWeightPerCar = 0 Then tmpWaste.MinWeightPerCar = tmpJobDet.QuoMinWeightPerCar
                    If tmpWaste.NBLMinRate = 0 Then tmpWaste.NBLMinRate = tmpJobDet.NBLMinRate
                    
                    If tmpWaste.MaxWeightPerCar = 0 Then tmpWaste.MaxWeightPerCar = tmpJobDet.QuoMaxWeightPerCar
                    If tmpWaste.MaxRate = 0 Then tmpWaste.MaxRate = tmpJobDet.QuoMaxRate
                    If tmpWaste.MaxRateType = "" Then tmpWaste.MaxRateType = tmpJobDet.QuoMaxRateType
                    Set tmpWaste = Nothing
                Else
                    Set tmpWaste = New BWGWasteDataCR
                    tmpWaste.id = "NBL"
                    tmpWaste.WasteName = "NBL"
                    tmpWaste.WasteType = "NBL"
                    tmpWaste.menifestID = tmpJobDet.menifestID
                    tmpWaste.TreatmentUnitID = tmpJobDet.QuoTreatmentUnitID
                    tmpWaste.MinWeightPerCar = tmpJobDet.QuoMinWeightPerCar
                    tmpWaste.NBLMinRate = tmpJobDet.NBLMinRate
                    tmpWaste.MaxWeightPerCar = tmpJobDet.QuoMaxWeightPerCar
                    tmpWaste.MaxRate = tmpJobDet.QuoMaxRate
                    tmpWaste.MaxRateType = tmpJobDet.QuoMaxRateType
                    If tmpJobDet.ChargeWeightType = "A" Then
                        tmpWaste.MaxQty = tmpJobDet.DisposerWeight
                    Else
                        tmpWaste.MaxQty = tmpJobDet.EstWasteQty
                    End If
                    ColWasteWeight.Add tmpWaste, "'NBL'"
                    Set tmpWaste = Nothing
                End If
            Else
                If isExist(ColWasteWeight, "'" & tmpJobDet.wastedataID & "'") Then
                    Set tmpWaste = ColWasteWeight("'" & tmpJobDet.wastedataID & "'")
                    If tmpJobDet.ChargeWeightType = "A" Then
                        tmpWaste.MaxQty = tmpWaste.MaxQty + tmpJobDet.DisposerWeight
                    Else
                        tmpWaste.MaxQty = tmpWaste.MaxQty + tmpJobDet.EstWasteQty
                    End If
                    If tmpWaste.MinWeightPerCar = 0 Then tmpWaste.MinWeightPerCar = tmpJobDet.QuoMinWeightPerCar
                    If tmpWaste.MaxWeightPerCar = 0 Then tmpWaste.MaxWeightPerCar = tmpJobDet.QuoMaxWeightPerCar
                    If tmpWaste.MaxRate = 0 Then tmpWaste.MaxRate = tmpJobDet.QuoMaxRate
                    If tmpWaste.MaxRateType = "" Then tmpWaste.MaxRateType = tmpJobDet.QuoMaxRateType
                    Set tmpWaste = Nothing
                Else
                    Set tmpWaste = New BWGWasteDataCR
                    tmpWaste.id = tmpJobDet.wastedataID
                    tmpWaste.WasteName = tmpJobDet.WasteName
                    tmpWaste.WasteType = tmpJobDet.WasteType
                    tmpWaste.menifestID = tmpJobDet.menifestID
                    tmpWaste.TreatmentUnitID = tmpJobDet.QuoTreatmentUnitID
                    tmpWaste.MaxWeightPerCar = tmpJobDet.QuoMaxWeightPerCar
                    tmpWaste.MaxRate = tmpJobDet.QuoMaxRate
                    tmpWaste.MaxRateType = tmpJobDet.QuoMaxRateType
                    tmpWaste.MinWeightPerCar = tmpJobDet.QuoMinWeightPerCar
                    If LCase(tmpJobDet.WasteType) = "haz" Then
                        If tmpJobDet.HazMinRate > 0 Then
                            tmpWaste.HazMinRate = tmpJobDet.HazMinRate
                        Else
                            tmpWaste.HazMinRate = tmpJobDet.QuoTreatmentRate
                        End If
                    ElseIf LCase(tmpJobDet.WasteType) = "hbl" Then
                        If tmpJobDet.HBLMinRate > 0 Then
                            tmpWaste.HBLMinRate = tmpJobDet.HBLMinRate
                        Else
                            tmpWaste.HBLMinRate = tmpJobDet.QuoTreatmentRate
                        End If
                    ElseIf LCase(tmpJobDet.WasteType) = "nbl" Then
                        If tmpJobDet.NBLMinRate > 0 Then
                            tmpWaste.NBLMinRate = tmpJobDet.NBLMinRate
                        Else
                            tmpWaste.NBLMinRate = tmpJobDet.QuoTreatmentRate
                        End If
                    Else
                        If tmpJobDet.HazMinRate > 0 Then
                            tmpWaste.HazMinRate = tmpJobDet.HazMinRate
                        Else
                            tmpWaste.HazMinRate = tmpJobDet.QuoTreatmentRate
                        End If
                    End If
                    If tmpJobDet.ChargeWeightType = "A" Then
                        
                            tmpWaste.MaxQty = tmpJobDet.DisposerWeight
                    
                    Else
                    
                            tmpWaste.MaxQty = tmpJobDet.EstWasteQty
                      

                    End If
                    ColWasteWeight.Add tmpWaste, "'" & tmpWaste.id & "'"
                    Set tmpWaste = Nothing
                End If
            End If
            tmpTimeTableNo = tmpJobDet.TimeTableNo
            selMID = tmpJobDet.menifestID
            Dim TType As String
            TType = tmpJobDet.TransRate
            Set tmpJobDet = Nothing
        Next '******** วบเช็คน้ำหนัก Waste วนตามใบคุม, ใบเมนิเฟส, ประเภทเวสต์
    End If ' colTrip is Nothing
    
    '*** เพิ่มข้อมูลลงใน colMinWeightCheck ในกรณีที่มีใบคุมใบเดียว หรือ ใบคุมใบสุดท้าย
    If Trim(tmpTimeTableNo) <> "" Then Call AddMinWeight(colMinWeightCheck, ColWasteWeight, tmpTimeTableNo)

'************* End Calculate min rate with waste type ***********************
    CurCustomerID = ""
    With grItem.getGridObj
        .Rows = 3
        curTreatUnitID = ""
        'Set tmpCol2 = myTripControl_.GetMenifests("T")  'Or "A"
'        Set tmpCol2 = tmpS.MenifestSearch("MenifestID in (" & selMenifestID & ") And (isCanceled<>'Y' Or isCanceled is null)")
        Set tmpCol2 = myTripControl_.GetMenifests("A")
        
        For x = 1 To tmpCol2.Count
            Set tmpItem = tmpCol2(x)
            'If Trim(curTripID) <> tmpItem.JobDataCarID & "_" & tmpItem.tripNo Then
            
            'If Trim(curTripID) <> tmpItem.TimeTableNo & "_" & tmpItem.TripNo Then 'กรณีรถพ่วงคิดแยกหัวกับหาง
            If Trim(curTripID) <> tmpItem.TimeTableNo Then  'กรณีรถพ่วงคิดแบบรวมทั้งคัน
                If Trim(curTripID) = "" Then
                    nonGroupWeight = 0
                    curGroupWeight = 0
                ElseIf tmpItem.TimeTableNo <> Replace(Trim(Split(curTripID, "_")(0)), "'", "") Then
                    nonGroupWeight = 0
                    curGroupWeight = 0
                End If
                curTreatUnitID = ""
                'curTripID = tmpItem.JobDataCarID & "_" & tmpItem.tripNo
                'curTripID = tmpItem.TimeTableNo & "_" & tmpItem.TripNo '*** กรณีรถพ่วงคิดแยกหัวกับหาง
                curTripID = tmpItem.TimeTableNo  '*** กรณีรถพ่วงคิดแบบรวมทั้งคัน
            End If
            If Trim(tmpItem.id) <> Trim(curMfID) Then
                curMfID = Trim(tmpItem.id)
                If maxType < LCase(tmpItem.NonOrHz) Then maxType = LCase(tmpItem.NonOrHz)
                If minType > LCase(tmpItem.NonOrHz) Then minType = LCase(tmpItem.NonOrHz)
                .TextMatrix(curRow, 7) = tmpItem.id
                .TextMatrix(curRow, 0) = tmpItem.JobDataCarID & "_" & tmpItem.TripNo
                .TextMatrix(curRow, 1) = "วันที่ " & FormatYMD_to_DMY(tmpItem.WorkDate, "/", "/") & " เลขที่ " & tmpItem.DocNo
                .TextMatrix(curRow, 2) = ""
                .TextMatrix(curRow, 3) = ""
                .TextMatrix(curRow, 4) = ""
                .TextMatrix(curRow, 5) = ""
                .TextMatrix(curRow, 6) = ""
                .TextMatrix(curRow, 8) = ""
                curRow = curRow + 1
                If .Rows - 1 = curRow Then .Rows = .Rows + 5
                If Trim(tmpItem.InOtherDesc) <> "" Then
                    .TextMatrix(curRow, 7) = tmpItem.id
                    .TextMatrix(curRow, 0) = tmpItem.JobDataCarID & "_" & tmpItem.TripNo
                    .TextMatrix(curRow, 1) = tmpItem.InOtherDesc
                    .TextMatrix(curRow, 2) = "1"
                    .TextMatrix(curRow, 3) = ""
                    .TextMatrix(curRow, 4) = tmpItem.InOtherFee
                    .TextMatrix(curRow, 5) = ""
                    .TextMatrix(curRow, 6) = tmpItem.InOtherFee
                    .TextMatrix(curRow, 8) = ""
                    tmpTotalAmt = tmpTotalAmt + tmpItem.InOtherFee
                    curRow = curRow + 1
                    If .Rows - 1 = curRow Then .Rows = .Rows + 5
                End If
            End If
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Set tmpCol = Nothing
            Set tmpCol = New Collection
            If Not colTrip Is Nothing Then
                For z = 1 To colTrip.Count
                    If Trim(colTrip(z).menifestID) = Trim(tmpItem.id) Then
                        tmpCol.Add colTrip(z), "'" & colTrip(z).id & "'"
                    End If
                Next
            End If
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            'Set tmpCol = tmpS.JobDetailInTripSearch("MenifestID='" & tmpItem.ID & "'", "WasteType,TreatmentUnitID")
            For y = 1 To tmpCol.Count
                Set tmpJobDet = tmpCol(y)
                tmpTranUnit = tmpJobDet.QuoTransUnitID
'                curTreatUnitID = Trim(tmpJobDet.QuoTreatmentUnitID)
                If curUnitTimeTable & "_" & CurCustomerID <> tmpJobDet.TimeTableNo & "_" & tmpJobDet.CustomerID Then If curUnitTimeTable <> "" Then maxNonTreatRate = 0: curGroupWeight = 0: nonGroupWeight = 0
                If Trim(tmpJobDet.FailType) = "" Then
                    .TextMatrix(curRow, 0) = tmpItem.JobDataCarID & "_" & tmpItem.TripNo
                    .TextMatrix(curRow, 1) = tmpJobDet.WasteName
                    .TextMatrix(curRow, 7) = tmpItem.id
                    If UCase(tmpJobDet.WasteType) = "NON" Or UCase(tmpJobDet.WasteType) = "NBL" Then
                        If tmpJobDet.ChargeWeightType = "A" Then
                            nonGroupWeight = nonGroupWeight + tmpJobDet.DisposerWeight
                        Else
                            nonGroupWeight = nonGroupWeight + tmpJobDet.EstWasteQty
                        End If
                        If maxNonTreatRate <= tmpJobDet.QuoTreatmentRate Then
                            maxNonTreatRate = tmpJobDet.QuoTreatmentRate
                            maxNonUnit = tmpJobDet.QuoTreatmentUnitID
                        End If
                    End If
                    If tmpJobDet.QuoTreatmentUnitID = "บาท/เที่ยว" Then
                        If Trim(curTreatUnitID) = "" Or (curUnitTimeTable & "_" & CurCustomerID <> tmpJobDet.TimeTableNo & "_" & tmpJobDet.CustomerID) Then
                            .TextMatrix(curRow - 1, 2) = "1"
                            .TextMatrix(curRow - 1, 3) = "เที่ยว"
                            .TextMatrix(curRow - 1, 4) = Format(tmpJobDet.QuoTreatmentRate, "#,##0.00")
                            .TextMatrix(curRow - 1, 5) = tmpJobDet.QuoTreatmentUnitID
                            .TextMatrix(curRow - 1, 6) = Format(tmpJobDet.QuoTreatmentRate, "#,##0.00")
                            .TextMatrix(curRow - 1, 7) = tmpItem.id
                            .TextMatrix(curRow - 1, 8) = Format(tmpJobDet.QuoCost, "#,##0.00")
                            tmpTotalAmt = tmpTotalAmt + tmpJobDet.QuoTreatmentRate
                            curUnitTimeTable = tmpJobDet.TimeTableNo
                        End If
                        If UCase(tmpItem.NonOrHz) = "NON" Or UCase(tmpItem.NonOrHz) = "NBL" Then
                            .TextMatrix(curRow, 2) = ""
                            .TextMatrix(curRow, 3) = ""
                            .TextMatrix(curRow, 4) = ""
                            .TextMatrix(curRow, 5) = ""
                            .TextMatrix(curRow, 6) = ""
                            .TextMatrix(curRow, 8) = Format(tmpJobDet.QuoCost, "#,##0.00")
                        Else
                            If tmpJobDet.ChargeWeightType = "A" Then
                                .TextMatrix(curRow, 2) = Format(tmpJobDet.DisposerWeight, "#,##0.000")
                            Else
                                .TextMatrix(curRow, 2) = Format(tmpJobDet.EstWasteQty, "#,##0.000")
                            End If
                            .TextMatrix(curRow, 3) = "ตัน"
                            .TextMatrix(curRow, 4) = ""
                            .TextMatrix(curRow, 5) = ""
                            .TextMatrix(curRow, 6) = ""
                            .TextMatrix(curRow, 8) = Format(tmpJobDet.QuoCost, "#,##0.00")
                        End If
                        If tmpJobDet.ChargeWeightType = "A" Then
                            curGroupWeight = curGroupWeight + tmpJobDet.DisposerWeight
                        Else
                            curGroupWeight = curGroupWeight + tmpJobDet.EstWasteQty
                        End If
                        curTreatUnitID = Trim(tmpJobDet.QuoTreatmentUnitID)
                    Else
                        If Trim(curTreatUnitID) <> "" Then
                            .TextMatrix(curRow, 7) = tmpItem.id
                            .TextMatrix(curRow, 0) = tmpItem.JobDataCarID & "_" & tmpItem.TripNo
                            .TextMatrix(curRow, 1) = "น้ำหนักรวม"
                            .TextMatrix(curRow, 2) = Format(tmpJobDet.MenifestActQty, "#,##0.000")
                            .TextMatrix(curRow, 3) = "ตัน"
                            .TextMatrix(curRow, 4) = ""
                            .TextMatrix(curRow, 5) = ""
                            .TextMatrix(curRow, 6) = ""
                            .TextMatrix(curRow, 8) = Format(tmpJobDet.QuoCost, "#,##0.00")
                            curRow = curRow + 1
                            If .Rows - 1 = curRow Then .Rows = .Rows + 5
                        End If
                        curTreatUnitID = ""
                        curGroupWeight = 0
                        If UCase(tmpJobDet.WasteType) <> "NON" And UCase(tmpJobDet.WasteType) <> "NBL" Then
                            If tmpJobDet.ChargeWeightType = "A" Then
                                .TextMatrix(curRow, 2) = Format(tmpJobDet.DisposerWeight, "#,##0.000")
                                .TextMatrix(curRow, 3) = Replace(tmpJobDet.QuoTreatmentUnitID, "บาท/", "")
                                tmpWeight = tmpJobDet.DisposerWeight
                            Else
                            If InStr(1, tmpJobDet.QuoTreatmentUnitID, "กิโลกรัม") > 0 Then
                                .TextMatrix(curRow, 2) = Format(tmpJobDet.EstWasteQty, "#,##0")
                            Else
                                .TextMatrix(curRow, 2) = Format(tmpJobDet.EstWasteQty, "#,##0.000")
                            End If
                                .TextMatrix(curRow, 3) = Replace(tmpJobDet.QuoTreatmentUnitID, "บาท/", "")
                                tmpWeight = tmpJobDet.EstWasteQty
                            End If
                            If colMinWeightCheck("'" & tmpJobDet.TimeTableNo & "'").isUnderWeight And LCase(Trim(colMinWeightCheck("'" & tmpJobDet.TimeTableNo & "'").MaxWasteID)) = LCase(Trim(tmpJobDet.wastedataID)) Then
                                If LCase(tmpJobDet.WasteType) = "haz" Then
                                    If tmpJobDet.HazMinRate > 0 Then
                                        .TextMatrix(curRow, 4) = Format(tmpJobDet.HazMinRate, "#,##0.00")
                                    Else
                                        .TextMatrix(curRow, 4) = Format(tmpJobDet.QuoTreatmentRate, "#,##0.00")
                                    End If
                                ElseIf LCase(tmpJobDet.WasteType) = "hbl" Then
                                    If tmpJobDet.HBLMinRate > 0 Then
                                        .TextMatrix(curRow, 4) = Format(tmpJobDet.HBLMinRate, "#,##0.00")
                                    Else
                                        .TextMatrix(curRow, 4) = Format(tmpJobDet.QuoTreatmentRate, "#,##0.00")
                                    End If
                                Else
                                    If tmpJobDet.HazMinRate > 0 Then
                                        .TextMatrix(curRow, 4) = Format(tmpJobDet.HazMinRate, "#,##0.00")
                                    Else
                                        .TextMatrix(curRow, 4) = Format(tmpJobDet.QuoTreatmentRate, "#,##0.00")
                                    End If
                                End If
                                .row = curRow
                                .col = 4
                                .CellForeColor = vbRed
                            Else
                                .TextMatrix(curRow, 4) = Format(tmpJobDet.QuoTreatmentRate, "#,##0.00")
                            End If
                            .TextMatrix(curRow, 5) = tmpJobDet.QuoTreatmentUnitID
                            .TextMatrix(curRow, 6) = Format(tmpWeight * CDbl(.TextMatrix(curRow, 4)), "#,##0.00")
                            .TextMatrix(curRow, 8) = Format(tmpWeight * CDbl(tmpJobDet.QuoCost), "#,##0.00")
                            tmpTotalAmt = tmpTotalAmt + CDbl(.TextMatrix(curRow, 6))
                            TotalCost = TotalCost + CDbl(.TextMatrix(curRow, 8))
                        End If
                    End If
                    curRow = curRow + 1
                    If .Rows - 1 = curRow Then .Rows = .Rows + 5
                End If
                tmpWasteType = tmpJobDet.WasteType
                tmpTimeTableNo = tmpJobDet.TimeTableNo
                CurCustomerID = tmpJobDet.CustomerID
                Set tmpJobDet = Nothing
            Next y
            
           Select Case UCase(tmpWasteType)
                Case "NON", "NBL"
                    If InStr(1, maxNonUnit, "ตัน") > 0 Then
                        If colMinWeightCheck("'" & tmpTimeTableNo & "'").isUnderWeight And LCase(Trim(colMinWeightCheck("'" & tmpTimeTableNo & "'").MaxWasteType)) = LCase(Trim(tmpWasteType)) Then
                            .TextMatrix(curRow, 7) = tmpItem.id
                            .TextMatrix(curRow, 0) = tmpItem.JobDataCarID & "_" & tmpItem.TripNo
                            .TextMatrix(curRow, 1) = "น้ำหนักรวม"
                            .TextMatrix(curRow, 2) = Format(nonGroupWeight, "#,##0.000")
                            .TextMatrix(curRow, 3) = "ตัน"
                            '=================================
                            .TextMatrix(curRow, 4) = Format(maxNonTreatRate, "#,##0.00")
                            .TextMatrix(curRow, 5) = maxNonUnit
                            .TextMatrix(curRow, 6) = Format(nonGroupWeight * maxNonTreatRate, "#,##0.00")
                            tmpTotalAmt = tmpTotalAmt + Format(nonGroupWeight * maxNonTreatRate, "0.00") * 1
                            '=================================
                            curRow = curRow + 1
                            If .Rows - 1 = curRow Then .Rows = .Rows + 5
                            If Trim(colMinWeightCheck("'" & tmpTimeTableNo & "'").menifestID) = Trim(tmpItem.id) Then
                                Set tmpWaste = colMinWeightCheck("'" & tmpTimeTableNo & "'").ColWasteWeight("'" & colMinWeightCheck("'" & tmpTimeTableNo & "'").MaxWasteID & "'")
                                .TextMatrix(curRow, 7) = tmpItem.id
                                .TextMatrix(curRow, 0) = tmpItem.JobDataCarID & "_" & tmpItem.TripNo
                                ExtraCharge = "น้ำหนักไม่ถึง " & tmpWaste.MinWeightPerCar & " ตัน คิดน้ำหนักเพิ่ม"
                                .TextMatrix(curRow, 1) = ExtraCharge
        '                        .TextMatrix(curRow, 2) = Format(tmpWaste.MinWeightPerCar, "#,##0.000")
                                '===============================
                                .TextMatrix(curRow, 2) = Format(tmpWaste.MinWeightPerCar - colMinWeightCheck("'" & tmpTimeTableNo & "'").TotalWeight, "#,##0.000")
                                TWeight = CDbl(.TextMatrix(curRow, 2))
                                '===============================
                                .TextMatrix(curRow, 3) = "ตัน"
                                If LCase(tmpWaste.WasteType) = "non" Then
                                    If tmpWaste.NonMinRate > 0 Then
                                        .TextMatrix(curRow, 4) = Format(tmpWaste.NonMinRate, "#,##0.00")
                                    Else
                                        .TextMatrix(curRow, 4) = Format(tmpWaste.TreatmentFee, "#,##0.00")
                                    End If
                                ElseIf LCase(tmpWaste.WasteType) = "nbl" Then
                                    If tmpWaste.NBLMinRate > 0 Then
                                        .TextMatrix(curRow, 4) = Format(tmpWaste.NBLMinRate, "#,##0.00")
                                    Else
                                        .TextMatrix(curRow, 4) = Format(tmpWaste.TreatmentFee, "#,##0.00")
                                    End If
                                End If
                                .TextMatrix(curRow, 5) = "บาท/ตัน"
        '                        .TextMatrix(curRow, 6) = Format(CDbl(.TextMatrix(curRow, 4)) * CDbl(.TextMatrix(curRow, 2)), "#,##0.00")
                                '===============
                                .TextMatrix(curRow, 6) = Format(CDbl(.TextMatrix(curRow, 4)) * TWeight, "#,##0.00")
                                '===============
                                tmpTotalAmt = tmpTotalAmt + CDbl(.TextMatrix(curRow, 6))
                                curRow = curRow + 1
                                If .Rows - 1 = curRow Then .Rows = .Rows + 5
                            End If
                        ElseIf colMinWeightCheck("'" & tmpTimeTableNo & "'").isOverWeight And LCase(Trim(colMinWeightCheck("'" & tmpTimeTableNo & "'").MaxWasteType)) = LCase(Trim(tmpWasteType)) Then
                            .TextMatrix(curRow, 7) = tmpItem.id
                            .TextMatrix(curRow, 0) = tmpItem.JobDataCarID & "_" & tmpItem.TripNo
                            .TextMatrix(curRow, 1) = "น้ำหนักรวม"
                            .TextMatrix(curRow, 2) = Format(nonGroupWeight, "#,##0.000")
                            .TextMatrix(curRow, 3) = "ตัน"
                            '=================================
                            .TextMatrix(curRow, 4) = Format(maxNonTreatRate, "#,##0.00")
                            .TextMatrix(curRow, 5) = maxNonUnit
                            .TextMatrix(curRow, 6) = Format(nonGroupWeight * maxNonTreatRate, "#,##0.00")
                            tmpTotalAmt = tmpTotalAmt + Format(nonGroupWeight * maxNonTreatRate, "0.00") * 1
                            '=================================
                            curRow = curRow + 1
                            If .Rows - 1 = curRow Then .Rows = .Rows + 5
                            If Trim(colMinWeightCheck("'" & tmpTimeTableNo & "'").menifestID) = Trim(tmpItem.id) Then
                                Set tmpWaste = colMinWeightCheck("'" & tmpTimeTableNo & "'").ColWasteWeight("'" & colMinWeightCheck("'" & tmpTimeTableNo & "'").MaxWasteID & "'")
                                .TextMatrix(curRow, 7) = tmpItem.id
                                .TextMatrix(curRow, 0) = tmpItem.JobDataCarID & "_" & tmpItem.TripNo
                                .TextMatrix(curRow, 1) = "น้ำหนักเกินที่กำหนด " & tmpWaste.MaxWeightPerCar & " ตัน คิดน้ำหนักส่วนเกิน"
                                .TextMatrix(curRow, 2) = Format(colMinWeightCheck("'" & tmpTimeTableNo & "'").TotalWeight - tmpWaste.MaxWeightPerCar, "#,##0.000")
                                .TextMatrix(curRow, 3) = "ตัน"
                                .TextMatrix(curRow, 4) = Format(tmpWaste.MaxRate, "#,##0.00")
                                .TextMatrix(curRow, 5) = tmpWaste.MaxRateType
                                If InStr(1, tmpWaste.MaxRateType, "เที่ยว") > 0 Then
                                    .TextMatrix(curRow, 6) = Format(tmpWaste.MaxRate, "#,##0.00")
                                Else
                                    .TextMatrix(curRow, 6) = Format(CDbl(.TextMatrix(curRow, 4)) * CDbl(.TextMatrix(curRow, 2)), "#,##0.00")
                                End If
                                tmpTotalAmt = tmpTotalAmt + CDbl(.TextMatrix(curRow, 6))
                                curRow = curRow + 1
                                If .Rows - 1 = curRow Then .Rows = .Rows + 5
                            End If
                        Else
                            .TextMatrix(curRow, 7) = tmpItem.id
                            .TextMatrix(curRow, 0) = tmpItem.JobDataCarID & "_" & tmpItem.TripNo
                            .TextMatrix(curRow, 1) = "น้ำหนักรวม"
                            .TextMatrix(curRow, 2) = Format(nonGroupWeight, "#,##0.000")
                            .TextMatrix(curRow, 3) = "ตัน"
                            .TextMatrix(curRow, 4) = Format(maxNonTreatRate, "#,##0.00")
                            .TextMatrix(curRow, 5) = maxNonUnit
                            .TextMatrix(curRow, 6) = Format(nonGroupWeight * maxNonTreatRate, "#,##0.00")
                            tmpTotalAmt = tmpTotalAmt + Format(nonGroupWeight * maxNonTreatRate, "0.00") * 1
                            curRow = curRow + 1
                            If .Rows - 1 = curRow Then .Rows = .Rows + 5
                        End If
                        nonGroupWeight = 0
                        maxNonTreatRate = 0
                        ElseIf InStr(1, maxNonUnit, "กิโลกรัม") > 0 Then
                        If colMinWeightCheck("'" & tmpTimeTableNo & "'").isUnderWeight And LCase(Trim(colMinWeightCheck("'" & tmpTimeTableNo & "'").MaxWasteType)) = LCase(Trim(tmpWasteType)) Then
                            .TextMatrix(curRow, 7) = tmpItem.id
                            .TextMatrix(curRow, 0) = tmpItem.JobDataCarID & "_" & tmpItem.TripNo
                            .TextMatrix(curRow, 1) = "น้ำหนักรวม"
                            .TextMatrix(curRow, 2) = Format(nonGroupWeight, "#,##0")
                            .TextMatrix(curRow, 3) = "กิโลกรัม"
                            '=================================
                            .TextMatrix(curRow, 4) = Format(maxNonTreatRate, "#,##0.00")
                            .TextMatrix(curRow, 5) = maxNonUnit
                            .TextMatrix(curRow, 6) = Format(nonGroupWeight * maxNonTreatRate, "#,##0.00")
                            tmpTotalAmt = tmpTotalAmt + Format(nonGroupWeight * maxNonTreatRate, "0.00") * 1
                            '=================================
                            curRow = curRow + 1
                            If .Rows - 1 = curRow Then .Rows = .Rows + 5
                            If Trim(colMinWeightCheck("'" & tmpTimeTableNo & "'").menifestID) = Trim(tmpItem.id) Then
                                Set tmpWaste = colMinWeightCheck("'" & tmpTimeTableNo & "'").ColWasteWeight("'" & colMinWeightCheck("'" & tmpTimeTableNo & "'").MaxWasteID & "'")
                                .TextMatrix(curRow, 7) = tmpItem.id
                                .TextMatrix(curRow, 0) = tmpItem.JobDataCarID & "_" & tmpItem.TripNo
                                ExtraCharge = "น้ำหนักไม่ถึง " & tmpWaste.MinWeightPerCar & " กิโลกรัม คิดน้ำหนักเพิ่ม"
                                .TextMatrix(curRow, 1) = ExtraCharge
        '                        .TextMatrix(curRow, 2) = Format(tmpWaste.MinWeightPerCar, "#,##0.000")
                                '===============================
                                .TextMatrix(curRow, 2) = Format((tmpWaste.MinWeightPerCar - colMinWeightCheck("'" & tmpTimeTableNo & "'").TotalWeight) * 1000, "#,##0")
                                TWeight = CDbl(.TextMatrix(curRow, 2))
                                '===============================
                                .TextMatrix(curRow, 3) = "กิโลกรัม"
                                If LCase(tmpWaste.WasteType) = "non" Then
                                    If tmpWaste.NonMinRate > 0 Then
                                        .TextMatrix(curRow, 4) = Format(tmpWaste.NonMinRate, "#,##0.00")
                                    Else
                                        .TextMatrix(curRow, 4) = Format(tmpWaste.TreatmentFee, "#,##0.00")
                                    End If
                                ElseIf LCase(tmpWaste.WasteType) = "nbl" Then
                                    If tmpWaste.NBLMinRate > 0 Then
                                        .TextMatrix(curRow, 4) = Format(tmpWaste.NBLMinRate, "#,##0.00")
                                    Else
                                        .TextMatrix(curRow, 4) = Format(tmpWaste.TreatmentFee, "#,##0.00")
                                    End If
                                End If
                                .TextMatrix(curRow, 5) = "บาท/ตัน"
        '                        .TextMatrix(curRow, 6) = Format(CDbl(.TextMatrix(curRow, 4)) * CDbl(.TextMatrix(curRow, 2)), "#,##0.00")
                                '===============
                                .TextMatrix(curRow, 6) = Format(CDbl(.TextMatrix(curRow, 4)) * TWeight, "#,##0.00")
                                '===============
                                tmpTotalAmt = tmpTotalAmt + CDbl(.TextMatrix(curRow, 6))
                                curRow = curRow + 1
                                If .Rows - 1 = curRow Then .Rows = .Rows + 5
                            End If
                        ElseIf colMinWeightCheck("'" & tmpTimeTableNo & "'").isOverWeight And LCase(Trim(colMinWeightCheck("'" & tmpTimeTableNo & "'").MaxWasteType)) = LCase(Trim(tmpWasteType)) Then
                            .TextMatrix(curRow, 7) = tmpItem.id
                            .TextMatrix(curRow, 0) = tmpItem.JobDataCarID & "_" & tmpItem.TripNo
                            .TextMatrix(curRow, 1) = "น้ำหนักรวม"
                            .TextMatrix(curRow, 2) = Format(nonGroupWeight, "#,##0")
                            .TextMatrix(curRow, 3) = "กิโลกรัม"
                            '=================================
                            .TextMatrix(curRow, 4) = Format(maxNonTreatRate, "#,##0.00")
                            .TextMatrix(curRow, 5) = maxNonUnit
                            .TextMatrix(curRow, 6) = Format(nonGroupWeight * maxNonTreatRate, "#,##0.00")
                            tmpTotalAmt = tmpTotalAmt + Format(nonGroupWeight * maxNonTreatRate, "0.00") * 1
                            '=================================
                            curRow = curRow + 1
                            If .Rows - 1 = curRow Then .Rows = .Rows + 5
                            If Trim(colMinWeightCheck("'" & tmpTimeTableNo & "'").menifestID) = Trim(tmpItem.id) Then
                                Set tmpWaste = colMinWeightCheck("'" & tmpTimeTableNo & "'").ColWasteWeight("'" & colMinWeightCheck("'" & tmpTimeTableNo & "'").MaxWasteID & "'")
                                .TextMatrix(curRow, 7) = tmpItem.id
                                .TextMatrix(curRow, 0) = tmpItem.JobDataCarID & "_" & tmpItem.TripNo
                                .TextMatrix(curRow, 1) = "น้ำหนักเกินที่กำหนด " & tmpWaste.MaxWeightPerCar & " กิโลกรัม คิดน้ำหนักส่วนเกิน"
                                .TextMatrix(curRow, 2) = Format((colMinWeightCheck("'" & tmpTimeTableNo & "'").TotalWeight - tmpWaste.MaxWeightPerCar) * 1000, "#,##0")
                                .TextMatrix(curRow, 3) = "กิโลกรัม"
                                .TextMatrix(curRow, 4) = Format(tmpWaste.MaxRate, "#,##0.00")
                                .TextMatrix(curRow, 5) = tmpWaste.MaxRateType
                                If InStr(1, tmpWaste.MaxRateType, "เที่ยว") > 0 Then
                                    .TextMatrix(curRow, 6) = Format(tmpWaste.MaxRate, "#,##0.00")
                                Else
                                    .TextMatrix(curRow, 6) = Format(CDbl(.TextMatrix(curRow, 4)) * CDbl(.TextMatrix(curRow, 2)), "#,##0.00")
                                End If
                                tmpTotalAmt = tmpTotalAmt + CDbl(.TextMatrix(curRow, 6))
                                curRow = curRow + 1
                                If .Rows - 1 = curRow Then .Rows = .Rows + 5
                            End If
                        Else
                            .TextMatrix(curRow, 7) = tmpItem.id
                            .TextMatrix(curRow, 0) = tmpItem.JobDataCarID & "_" & tmpItem.TripNo
                            .TextMatrix(curRow, 1) = "น้ำหนักรวม"
                            .TextMatrix(curRow, 2) = Format(nonGroupWeight, "#,##0")
                            .TextMatrix(curRow, 3) = "กิโลกรัม"
                            .TextMatrix(curRow, 4) = Format(maxNonTreatRate, "#,##0.00")
                            .TextMatrix(curRow, 5) = maxNonUnit
                            .TextMatrix(curRow, 6) = Format(nonGroupWeight * maxNonTreatRate, "#,##0.00")
                            tmpTotalAmt = tmpTotalAmt + Format(nonGroupWeight * maxNonTreatRate, "0.00") * 1
                            curRow = curRow + 1
                            If .Rows - 1 = curRow Then .Rows = .Rows + 5
                        End If
                        nonGroupWeight = 0
                        maxNonTreatRate = 0
                    Else
                        GoTo RatePerTrip
                    End If
                    
                Case "HAZ", "HBL", "HIC", "HBI", "BBH", "BBL", "NIC", "NBI"
                    If InStr(1, Trim(curTreatUnitID), "เที่ยว") = 0 Then
                        If Trim(colMinWeightCheck("'" & tmpTimeTableNo & "'").MaxWasteID) <> "" Then
                            Set tmpWaste = colMinWeightCheck("'" & tmpTimeTableNo & "'").ColWasteWeight("'" & colMinWeightCheck("'" & tmpTimeTableNo & "'").MaxWasteID & "'")
                            If InStr(1, tmpWaste.TreatmentUnitID, "ตัน") > 0 Then
                                If LCase(Trim(tmpWaste.WasteType)) = LCase(Trim(tmpWasteType)) And Trim(colMinWeightCheck("'" & tmpTimeTableNo & "'").menifestID) = Trim(tmpItem.id) Then
                                    If colMinWeightCheck("'" & tmpTimeTableNo & "'").isUnderWeight Then
                                        .TextMatrix(curRow, 7) = tmpItem.id
                                        .TextMatrix(curRow, 0) = tmpItem.JobDataCarID & "_" & tmpItem.TripNo
                                        If Trim(.TextMatrix(curRow, 1)) <> "" Then
                                            curRow = curRow + 1
                                            If .Rows - 1 = curRow Then .Rows = .Rows + 5
                                        End If
                                        .TextMatrix(curRow, 1) = "น้ำหนักไม่ถึง " & tmpWaste.MinWeightPerCar & " ตัน คิดน้ำหนักเพิ่มอีก"
                                        .TextMatrix(curRow, 2) = Format(tmpWaste.MinWeightPerCar - colMinWeightCheck("'" & tmpTimeTableNo & "'").TotalWeight, "#,##0.000")
                                        .TextMatrix(curRow, 3) = "ตัน"
                                        If LCase(tmpWaste.WasteType) = "haz" Then
                                            If tmpWaste.HazMinRate > 0 Then
                                                .TextMatrix(curRow, 4) = Format(tmpWaste.HazMinRate, "#,##0.00")
                                            Else
                                                .TextMatrix(curRow, 4) = Format(tmpWaste.TreatmentFee, "#,##0.00")
                                            End If
                                        ElseIf LCase(tmpWaste.WasteType) = "hbl" Then
                                            If tmpWaste.HBLMinRate > 0 Then
                                                .TextMatrix(curRow, 4) = Format(tmpWaste.HBLMinRate, "#,##0.00")
                                            Else
                                                .TextMatrix(curRow, 4) = Format(tmpWaste.TreatmentFee, "#,##0.00")
                                            End If
                                        Else
                                            If tmpWaste.HazMinRate > 0 Then
                                                .TextMatrix(curRow, 4) = Format(tmpWaste.HazMinRate, "#,##0.00")
                                            Else
                                                .TextMatrix(curRow, 4) = Format(tmpWaste.TreatmentFee, "#,##0.00")
                                            End If
                                        End If
                                        .TextMatrix(curRow, 5) = "บาท/ตัน"
                                        .TextMatrix(curRow, 6) = Format(CDbl(.TextMatrix(curRow, 4)) * CDbl(.TextMatrix(curRow, 2)), "#,##0.00")
                                        tmpTotalAmt = tmpTotalAmt + CDbl(.TextMatrix(curRow, 6))
                                        curRow = curRow + 1
                                        If .Rows - 1 = curRow Then .Rows = .Rows + 5
                                    ElseIf colMinWeightCheck("'" & tmpTimeTableNo & "'").isOverWeight Then
                                        .TextMatrix(curRow, 7) = tmpItem.id
                                        .TextMatrix(curRow, 0) = tmpItem.JobDataCarID & "_" & tmpItem.TripNo
                                        If Trim(.TextMatrix(curRow, 1)) <> "" Then
                                            curRow = curRow + 1
                                            If .Rows - 1 = curRow Then .Rows = .Rows + 5
                                        End If
                                        .TextMatrix(curRow, 1) = "น้ำหนักเกินที่กำหนด " & tmpWaste.MaxWeightPerCar & " ตัน คิดน้ำหนักส่วนเกิน"
                                        .TextMatrix(curRow, 2) = Format(colMinWeightCheck("'" & tmpTimeTableNo & "'").TotalWeight - tmpWaste.MaxWeightPerCar, "#,##0.000")
                                        .TextMatrix(curRow, 3) = "ตัน"
                                        .TextMatrix(curRow, 4) = Format(tmpWaste.MaxRate, "#,##0.00")
                                        .TextMatrix(curRow, 5) = tmpWaste.MaxRateType
                                        If InStr(1, tmpWaste.MaxRateType, "เที่ยว") > 0 Then
                                            .TextMatrix(curRow, 6) = Format(tmpWaste.MaxRate, "#,##0.00")
                                        Else
                                            .TextMatrix(curRow, 6) = Format(CDbl(.TextMatrix(curRow, 4)) * CDbl(.TextMatrix(curRow, 2)), "#,##0.00")
                                        End If
                                        tmpTotalAmt = tmpTotalAmt + CDbl(.TextMatrix(curRow, 6))
                                        curRow = curRow + 1
                                        If .Rows - 1 = curRow Then .Rows = .Rows + 5
                                    End If
                                End If
                            ElseIf InStr(1, tmpWaste.TreatmentUnitID, "กิโลกรัม") Then
                                If LCase(Trim(tmpWaste.WasteType)) = LCase(Trim(tmpWasteType)) And Trim(colMinWeightCheck("'" & tmpTimeTableNo & "'").menifestID) = Trim(tmpItem.id) Then
                                    If colMinWeightCheck("'" & tmpTimeTableNo & "'").isUnderWeight Then
                                        .TextMatrix(curRow, 7) = tmpItem.id
                                        .TextMatrix(curRow, 0) = tmpItem.JobDataCarID & "_" & tmpItem.TripNo
                                        If Trim(.TextMatrix(curRow, 1)) <> "" Then
                                            curRow = curRow + 1
                                            If .Rows - 1 = curRow Then .Rows = .Rows + 5
                                        End If
                                        .TextMatrix(curRow, 1) = "น้ำหนักไม่ถึง " & tmpWaste.MinWeightPerCar & " กิโลกรัม คิดน้ำหนักเพิ่มอีก"
                                        .TextMatrix(curRow, 2) = Format((tmpWaste.MinWeightPerCar - colMinWeightCheck("'" & tmpTimeTableNo & "'").TotalWeight) * 1000, "#,##0")
                                        .TextMatrix(curRow, 3) = "กิโลกรัม"
                                        If LCase(tmpWaste.WasteType) = "haz" Then
                                            If tmpWaste.HazMinRate > 0 Then
                                                .TextMatrix(curRow, 4) = Format(tmpWaste.HazMinRate, "#,##0.00")
                                            Else
                                                .TextMatrix(curRow, 4) = Format(tmpWaste.TreatmentFee, "#,##0.00")
                                            End If
                                        ElseIf LCase(tmpWaste.WasteType) = "hbl" Then
                                            If tmpWaste.HBLMinRate > 0 Then
                                                .TextMatrix(curRow, 4) = Format(tmpWaste.HBLMinRate, "#,##0.00")
                                            Else
                                                .TextMatrix(curRow, 4) = Format(tmpWaste.TreatmentFee, "#,##0.00")
                                            End If
                                        Else
                                            If tmpWaste.HazMinRate > 0 Then
                                                .TextMatrix(curRow, 4) = Format(tmpWaste.HazMinRate, "#,##0.00")
                                            Else
                                                .TextMatrix(curRow, 4) = Format(tmpWaste.TreatmentFee, "#,##0.00")
                                            End If
                                        End If
                                        .TextMatrix(curRow, 5) = "บาท/กิโลกรัม"
                                        .TextMatrix(curRow, 6) = Format(CDbl(.TextMatrix(curRow, 4)) * CDbl(.TextMatrix(curRow, 2)), "#,##0.00")
                                        tmpTotalAmt = tmpTotalAmt + CDbl(.TextMatrix(curRow, 6))
                                        curRow = curRow + 1
                                        If .Rows - 1 = curRow Then .Rows = .Rows + 5
                                    ElseIf colMinWeightCheck("'" & tmpTimeTableNo & "'").isOverWeight Then
                                        .TextMatrix(curRow, 7) = tmpItem.id
                                        .TextMatrix(curRow, 0) = tmpItem.JobDataCarID & "_" & tmpItem.TripNo
                                        If Trim(.TextMatrix(curRow, 1)) <> "" Then
                                            curRow = curRow + 1
                                            If .Rows - 1 = curRow Then .Rows = .Rows + 5
                                        End If
                                        .TextMatrix(curRow, 1) = "น้ำหนักเกินที่กำหนด " & tmpWaste.MaxWeightPerCar & "กิโลกรัม  คิดน้ำหนักส่วนเกิน"
                                        .TextMatrix(curRow, 2) = Format((colMinWeightCheck("'" & tmpTimeTableNo & "'").TotalWeight - tmpWaste.MaxWeightPerCar) * 1000, "#,##0")
                                        .TextMatrix(curRow, 3) = "กิโลกรัม"
                                        .TextMatrix(curRow, 4) = Format(tmpWaste.MaxRate, "#,##0.00")
                                        .TextMatrix(curRow, 5) = tmpWaste.MaxRateType
                                        If InStr(1, tmpWaste.MaxRateType, "เที่ยว") > 0 Then
                                            .TextMatrix(curRow, 6) = Format(tmpWaste.MaxRate, "#,##0.00")
                                        Else
                                            .TextMatrix(curRow, 6) = Format(CDbl(.TextMatrix(curRow, 4)) * CDbl(.TextMatrix(curRow, 2)), "#,##0.00")
                                        End If
                                        tmpTotalAmt = tmpTotalAmt + CDbl(.TextMatrix(curRow, 6))
                                        curRow = curRow + 1
                                        If .Rows - 1 = curRow Then .Rows = .Rows + 5
                                    End If
                                End If
                            End If
                            Set tmpWaste = Nothing
                        Else
                            GoTo RatePerTrip
                        End If
                    Else
                        GoTo RatePerTrip
                    End If
                    
                Case Else
RatePerTrip:
                    If Trim(curTreatUnitID) <> "" Then
                     If InStr(1, curTreatUnitID, "กิโลกรัม") Then
                      If colMinWeightCheck("'" & tmpTimeTableNo & "'").isOverWeight And LCase(Trim(colMinWeightCheck("'" & tmpTimeTableNo & "'").MaxWasteType)) = LCase(Trim(tmpWasteType)) Then
                             Set tmpWaste = colMinWeightCheck("'" & tmpTimeTableNo & "'").ColWasteWeight("'" & colMinWeightCheck("'" & tmpTimeTableNo & "'").MaxWasteID & "'")
                             .TextMatrix(curRow, 0) = tmpItem.JobDataCarID & "_" & tmpItem.TripNo
                             .TextMatrix(curRow, 1) = "น้ำหนักรวม"
                             .TextMatrix(curRow, 2) = Format(curGroupWeight, "#,##0")
                             .TextMatrix(curRow, 3) = "กิโลกรัม"
                             .TextMatrix(curRow, 4) = ""
                             .TextMatrix(curRow, 5) = ""
                             .TextMatrix(curRow, 6) = ""
                             .TextMatrix(curRow, 7) = tmpItem.id
                             If Trim(.TextMatrix(curRow, 1)) <> "" Then
                                 curRow = curRow + 1
                                 If .Rows - 1 = curRow Then .Rows = .Rows + 5
                             End If
                             If Trim(.TextMatrix(curRow, 1)) <> "" Then
                                 curRow = curRow + 1
                                 If .Rows - 1 = curRow Then .Rows = .Rows + 5
                             End If
                             If tmpWaste.menifestID = tmpItem.id Then
                                 .TextMatrix(curRow, 0) = tmpItem.JobDataCarID & "_" & tmpItem.TripNo
                                 .TextMatrix(curRow, 1) = "น้ำหนักเกินที่กำหนด " & tmpWaste.MaxWeightPerCar & " กิโลกรัม คิดน้ำหนักส่วนเกิน"
                                 .TextMatrix(curRow, 2) = Format((colMinWeightCheck("'" & tmpTimeTableNo & "'").TotalWeight - tmpWaste.MaxWeightPerCar) * 1000, "#,##0")
                                 .TextMatrix(curRow, 3) = "กิโลกรัม"
                                 .TextMatrix(curRow, 4) = Format(tmpWaste.MaxRate, "#,##0.00")
                                 .TextMatrix(curRow, 5) = tmpWaste.MaxRateType
                                 If InStr(1, tmpWaste.MaxRateType, "เที่ยว") > 0 Then
                                     .TextMatrix(curRow, 6) = Format(tmpWaste.MaxRate, "#,##0.00")
                                 Else
                                     .TextMatrix(curRow, 6) = Format(CDbl(.TextMatrix(curRow, 4)) * CDbl(.TextMatrix(curRow, 2)), "#,##0.00")
                                 End If
                                 .TextMatrix(curRow, 7) = tmpItem.id
                                 tmpTotalAmt = tmpTotalAmt + CDbl(.TextMatrix(curRow, 6))
                             End If
                             curRow = curRow + 1
                             If .Rows - 1 = curRow Then .Rows = .Rows + 5
                         Else
                             .TextMatrix(curRow, 0) = tmpItem.JobDataCarID & "_" & tmpItem.TripNo
                             .TextMatrix(curRow, 1) = "น้ำหนักรวม"
                             .TextMatrix(curRow, 2) = Format(curGroupWeight, "#,##0")
                             .TextMatrix(curRow, 3) = "กิโลกรัม"
                             .TextMatrix(curRow, 4) = ""
                             .TextMatrix(curRow, 5) = ""
                             .TextMatrix(curRow, 6) = ""
                             .TextMatrix(curRow, 7) = tmpItem.id
                             curRow = curRow + 1
                             If .Rows - 1 = curRow Then .Rows = .Rows + 5
                             curGroupWeight = 0
                         End If
                     Else
                         If colMinWeightCheck("'" & tmpTimeTableNo & "'").isOverWeight And LCase(Trim(colMinWeightCheck("'" & tmpTimeTableNo & "'").MaxWasteType)) = LCase(Trim(tmpWasteType)) Then
                             Set tmpWaste = colMinWeightCheck("'" & tmpTimeTableNo & "'").ColWasteWeight("'" & colMinWeightCheck("'" & tmpTimeTableNo & "'").MaxWasteID & "'")
                             .TextMatrix(curRow, 0) = tmpItem.JobDataCarID & "_" & tmpItem.TripNo
                             .TextMatrix(curRow, 1) = "น้ำหนักรวม"
                             .TextMatrix(curRow, 2) = Format(curGroupWeight, "#,##0.000")
                             .TextMatrix(curRow, 3) = "ตัน"
                             .TextMatrix(curRow, 4) = ""
                             .TextMatrix(curRow, 5) = ""
                             .TextMatrix(curRow, 6) = ""
                             .TextMatrix(curRow, 7) = tmpItem.id
                             If Trim(.TextMatrix(curRow, 1)) <> "" Then
                                 curRow = curRow + 1
                                 If .Rows - 1 = curRow Then .Rows = .Rows + 5
                             End If
                             If Trim(.TextMatrix(curRow, 1)) <> "" Then
                                 curRow = curRow + 1
                                 If .Rows - 1 = curRow Then .Rows = .Rows + 5
                             End If
                             If tmpWaste.menifestID = tmpItem.id Then
                                 .TextMatrix(curRow, 0) = tmpItem.JobDataCarID & "_" & tmpItem.TripNo
                                 .TextMatrix(curRow, 1) = "น้ำหนักเกินที่กำหนด " & tmpWaste.MaxWeightPerCar & " ตัน คิดน้ำหนักส่วนเกิน"
                                 .TextMatrix(curRow, 2) = Format((colMinWeightCheck("'" & tmpTimeTableNo & "'").TotalWeight - tmpWaste.MaxWeightPerCar), "#,##0.000")
                                 .TextMatrix(curRow, 3) = "ตัน"
                                 .TextMatrix(curRow, 4) = Format(tmpWaste.MaxRate, "#,##0.00")
                                 .TextMatrix(curRow, 5) = tmpWaste.MaxRateType
                                 If InStr(1, tmpWaste.MaxRateType, "เที่ยว") > 0 Then
                                     .TextMatrix(curRow, 6) = Format(tmpWaste.MaxRate, "#,##0.00")
                                 Else
                                     .TextMatrix(curRow, 6) = Format(CDbl(.TextMatrix(curRow, 4)) * CDbl(.TextMatrix(curRow, 2)), "#,##0.00")
                                 End If
                                 .TextMatrix(curRow, 7) = tmpItem.id
                                 tmpTotalAmt = tmpTotalAmt + CDbl(.TextMatrix(curRow, 6))
                             End If
                             curRow = curRow + 1
                             If .Rows - 1 = curRow Then .Rows = .Rows + 5
                         Else
                             .TextMatrix(curRow, 0) = tmpItem.JobDataCarID & "_" & tmpItem.TripNo
                             .TextMatrix(curRow, 1) = "น้ำหนักรวม"
                             .TextMatrix(curRow, 2) = Format(curGroupWeight, "#,##0.000")
                             .TextMatrix(curRow, 3) = "ตัน"
                             .TextMatrix(curRow, 4) = ""
                             .TextMatrix(curRow, 5) = ""
                             .TextMatrix(curRow, 6) = ""
                             .TextMatrix(curRow, 7) = tmpItem.id
                             curRow = curRow + 1
                             If .Rows - 1 = curRow Then .Rows = .Rows + 5
                             curGroupWeight = 0
                         End If
                     End If
                    End If
            End Select
            Set tmpItem = Nothing
        Next x
    End With
    
    'เงื่อนไขค่าขนส่ง
    If TType = "H" Then
        txtTransFee.Text = CDbl(txtTransFee.Text) / 2
    ElseIf TType = "F" Then
        txtTransFee.Text = "0"
    End If
     
    SumAllFee
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'        txtAllMenifest.Text = Format(tmpAmt, "#,##0.00")
'        txtAllFee.Text = Format(tmpAmt + myTripControl_.TripTranFee + myTripControl_.TripMoreWorkerFee + myTripControl_.TripEquipFee + myTripControl_.TripOtherFee, "#,##0.00")
'
'        Set tmpCol2 = tmpS.JobDetailInTripSearch("JobDataCarID='" & myTripControl_.JobDataCarID & "' And TripNo=" & myTripControl_.tripNo)
'
        Select Case LCase(maxType)
            Case "non": txtWasteType.Text = "Non"
            Case "nbl": txtWasteType.Text = "NBL"
            Case "haz": txtWasteType.Text = "Haz"
            Case "hbl": txtWasteType.Text = "HBL"
            Case "hic": txtWasteType.Text = "HIC"
            Case "hbi": txtWasteType.Text = "HBI"
            Case "bbh": txtWasteType.Text = "BBH"
            Case "bbl": txtWasteType.Text = "BBL"
            Case Else
                txtWasteType.Text = UCase(maxType)
        End Select
        
        Set tmpCol = Nothing
        Set tmpCol2 = Nothing
        'myTripControl_.isChanged = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim selResult As Integer
    If myTripControl_ Is Nothing Then Exit Sub
    If myTripControl_.isChanged Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนปิดหน้าจอหรือไม่?", vbQuestion + vbYesNoCancel)
        If selResult = vbYes Then
            CmdSave_Click
            If myTripControl_.isChanged Then Cancel = 1
        ElseIf selResult = vbCancel Then
            Cancel = 1
        End If
    End If
End Sub

Private Sub grItem_BeforeUpdateValue(NewValue As String, row As Integer, col As Integer, Cancel As Boolean)
    If col = 2 Or col = 4 Then
        If IsNumeric(NewValue) = False Then
            MsgBox "กรุณาใส่ข้อมูลเป็นตัวเลขเท่านั้น", vbExclamation
            Cancel = True
        End If
    End If
End Sub

Private Sub grItem_Click(row As Integer, col As Integer)
    'grItem.getGridObj.Rows = 50
    'MsgBox grItem.getGridObj.ColWidth(col)
End Sub

Private Sub grItem_RolColChanged(row As Integer, col As Integer)
    grItem.ColEnabled(col) = False
End Sub

Private Sub grItem_UpdateValue(row As Integer, col As Integer)
Dim tmpItem As BWGMenifest
Dim tmpKey As String, tmpVal As String
Dim tmpCol As Collection
    tmpKey = Trim(grItem.getGridObj.TextMatrix(row, 0))
    tmpVal = Trim(grItem.getGridObj.TextMatrix(row, col))
    Set tmpCol = myTripControl_.GetMenifests("A")
    If isExist(tmpCol, "'" & tmpKey & "'") Then
        Set tmpItem = tmpCol("'" & tmpKey & "'")
        Select Case col
            Case 2
                If IsNumeric(tmpVal) Then
                    tmpItem.InCollectFee = tmpVal
                Else
                    tmpItem.InCollectFee = 0
                End If
            Case 3
                tmpItem.InOtherDesc = Trim(tmpVal)
            Case 4
                If IsNumeric(tmpVal) Then
                    tmpItem.InOtherFee = tmpVal
                Else
                    tmpItem.InOtherFee = 0
                End If
        End Select
        myTripControl_.isChanged = True
    End If
    Set tmpItem = Nothing
    Set tmpCol = Nothing
    SumAllFee
    If IsNumeric(Trim(txtAllFee.Text)) Then myTripControl_.TripAllFee = Trim(txtAllFee.Text)
End Sub

Private Sub txtAllFee_Change()
    If IsNumeric(Trim(txtAllFee.Text)) And Not myTripControl_ Is Nothing Then myTripControl_.TripAllFee = Trim(txtAllFee.Text)
End Sub

Private Sub txtAllMenifest_Change()
    If IsNumeric(Trim(txtAllMenifest.Text)) And Not myTripControl_ Is Nothing Then myTripControl_.TripTreatFee = Trim(txtAllMenifest.Text)
End Sub



Private Sub txtEquip_Change()
    If Me.Visible Then
        If Me.ActiveControl.Name = "txtEquip" Then
            SumAllFee
        End If
    End If
End Sub

Private Sub txtEquip_Validate(Cancel As Boolean)
    If IsNumeric(txtEquip.Text) Then
        myTripControl_.TripEquipFee = txtEquip.Text
    Else
        myTripControl_.TripEquipFee = 0
    End If
    If IsNumeric(Trim(txtAllFee.Text)) Then myTripControl_.TripAllFee = Trim(txtAllFee.Text)
End Sub

Private Sub txtIncTransFee_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then txtIncTransFee_Validate False
End Sub

Private Sub txtIncTransFee_Validate(Cancel As Boolean)
    If IsNumeric(txtIncTransFee.Text) Then
        myTripControl_.IncTripTransFee = txtIncTransFee
    Else
        myTripControl_.IncTripTransFee = 0
    End If
    If IsNumeric(Trim(txtAllFee.Text)) Then myTripControl_.TripAllFee = Trim(txtAllFee.Text)
    
    'Set Current for last edited data.
    EditType.EditbyUserID = CurrentUser.id
    EditType.EditbyUserName = Trim(CurrentUser.StaffFName & " " & CurrentUser.StaffLName)
    EditType.EditType = "M"
    myTripControl_.EditTypeTrip = "M"
End Sub

Private Sub txtMoreWorker_Validate(Cancel As Boolean)
    If IsNumeric(txtMoreWorker.Text) Then
        myTripControl_.TripMoreWorker = txtMoreWorker.Text
    Else
        myTripControl_.TripMoreWorker = 0
    End If
End Sub

Private Sub txtMoreWorkerFee_Change()
    If Me.Visible Then
        If Me.ActiveControl.Name = "txtMoreWorkerFee" Then
            SumAllFee
        End If
    End If
End Sub

Private Sub txtMoreWorkerFee_Validate(Cancel As Boolean)
    If IsNumeric(txtMoreWorkerFee.Text) Then
        myTripControl_.TripMoreWorkerFee = txtMoreWorkerFee.Text
    Else
        myTripControl_.TripMoreWorkerFee = 0
    End If
    If IsNumeric(Trim(txtAllFee.Text)) Then myTripControl_.TripAllFee = Trim(txtAllFee.Text)
End Sub

Private Sub txtOtherDesc_Validate(Cancel As Boolean)
    myTripControl_.TripOtherDesc = Trim(txtOtherDesc.Text)
End Sub

Private Sub txtOtherFee_Change()
    If Me.Visible Then
        If Me.ActiveControl.Name = "txtOtherFee" Then
            SumAllFee
        End If
    End If
End Sub

Private Sub txtOtherFee_Validate(Cancel As Boolean)
    If IsNumeric(txtOtherFee.Text) Then
        myTripControl_.TripOtherFee = txtOtherFee.Text
    Else
        myTripControl_.TripOtherFee = 0
    End If
    If IsNumeric(Trim(txtAllFee.Text)) Then myTripControl_.TripAllFee = Trim(txtAllFee.Text)
End Sub

Private Sub txtTimeIn_Validate(Cancel As Boolean)
    myTripControl_.CustInTime = Trim(txtTimeIn.Text)
End Sub

Private Sub txtTimeOut_Validate(Cancel As Boolean)
    myTripControl_.CustOutTime = txtTimeOut.Text
End Sub

Private Sub txtTransFee_Change()
    If Me.Visible Then
        If Me.ActiveControl.Name = "txtTransFee" Then
            SumAllFee
        End If
    End If
End Sub

Private Sub txtTransFee_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKey0, vbKey1, vbKey2, vbKey3, vbKey4, vbKey5, vbKey6, vbKey7, vbKey8, vbKey9
            myTripControl_.EditTypeTrip = "M"
    End Select
End Sub

Private Sub txtTransFee_Validate(Cancel As Boolean)
    If IsNumeric(txtTransFee.Text) Then
        myTripControl_.TripTranFee = txtTransFee.Text
    Else
        myTripControl_.TripTranFee = 0
    End If
    If IsNumeric(Trim(txtAllFee.Text)) Then myTripControl_.TripAllFee = Trim(txtAllFee.Text)
    
    'Set Current for last edited data.
    EditType.EditbyUserID = CurrentUser.id
    EditType.EditbyUserName = Trim(CurrentUser.StaffFName & " " & CurrentUser.StaffLName)
    EditType.EditType = "M"
End Sub

Private Sub CalCustChargeFee()
Dim maxWorker%, maxTransFee As Double
Dim DiscountTrans As String
    CalCustTripFee myTripControl_, maxWorker, maxTransFee, DiscountTrans
    
    txtTransFee.Text = myTripControl_.TripTranFee
    txtMoreWorker.Text = myTripControl_.TripMoreWorker
    txtMoreWorkerFee.Text = myTripControl_.TripMoreWorkerFee
    
    ' *** คำนวนจำนวนคนงานจ่ายให้ vendor
'    myTripControl_.TruckTypeID = ""
    txtVDWorker.Text = myTripControl_.WorkerCount
    txtVDWorkerFee.Text = Format(300 * myTripControl_.WorkerCount, "#,##0.00")
    
    myTripControl_.VDWorker = myTripControl_.WorkerCount
    myTripControl_.VDWorkerFee = myTripControl_.WorkerCount * 300
    ' ***
    
    txtCalTransFee.Text = Format(maxTransFee, "#,##0.00")

'** Remark Date 21/04/2011 09:26
    txtCalStaffCount.Text = myTripControl_.WorkerCount + myTripControl_.TripMoreWorker
    
'    If maxWorker < myTripControl_.WorkerCount Then
'        txtCalStaffCount.Text = myTripControl_.WorkerCount - maxWorker
'    Else
'        txtCalStaffCount.Text = ""
'    End If
'*** End

End Sub

Private Sub SumMenifestFee()
Dim x%, amt1 As Double, amt2 As Double, amt3 As Double
    With grItem.getGridObj
        For x = 1 To .Rows - 1
            If IsNumeric(.TextMatrix(x, 6)) Then
                amt1 = amt1 + CDbl(.TextMatrix(x, 6))
            End If
        Next
    End With
    txtAllMenifest.Text = Format(amt1, "#,##0.00")
End Sub

Private Sub SumAllFee()
Dim tmpAmt As Double
    SumMenifestFee
        If IsNumeric(txtTransFee.Text) Then tmpAmt = CDbl(txtTransFee.Text)
        If IsNumeric(txtMoreWorkerFee.Text) Then tmpAmt = tmpAmt + CDbl(txtMoreWorkerFee.Text)
        If IsNumeric(txtEquip.Text) Then tmpAmt = tmpAmt + CDbl(txtEquip.Text)
        If IsNumeric(txtOtherFee.Text) Then tmpAmt = tmpAmt + CDbl(txtOtherFee.Text)
        
        If IsNumeric(txtAllMenifest.Text) Then
            tmpAmt = tmpAmt + CDbl(txtAllMenifest.Text)
        End If
        txtAllFee.Text = Format(tmpAmt, "#,##0.00")
End Sub

Private Sub txtVDEquipFee_Change()
    If Me.Visible Then
        If Me.ActiveControl.Name = "txtVDEquipFee" Then
            SumPayVendor
        End If
    End If
End Sub

Private Sub txtVDEquipFee_Validate(Cancel As Boolean)
    If IsNumeric(txtVDEquipFee.Text) Then
        myTripControl_.VDEquipFee = txtVDEquipFee.Text
    Else
        myTripControl_.VDEquipFee = 0
    End If
End Sub

Private Sub txtVDOtherDesc_Validate(Cancel As Boolean)
    myTripControl_.VDOtherDesc = txtVDOtherDesc.Text
End Sub

Private Sub txtVDOtherFee_Change()
    If Me.Visible Then
        If Me.ActiveControl.Name = "txtVDOtherFee" Then
            SumPayVendor
        End If
    End If
End Sub

Private Sub txtVDOtherFee_Validate(Cancel As Boolean)
    If IsNumeric(txtVDOtherFee.Text) Then
        myTripControl_.VDOtherFee = txtVDOtherFee.Text
    Else
        myTripControl_.VDOtherFee = 0
    End If
End Sub

Private Sub txtVDTransFee_Change()
    If Me.Visible Then
        If Me.ActiveControl.Name = "txtVDTransFee" Then
            SumPayVendor
        End If
    End If
End Sub

Private Sub txtVDTransFee_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKey0, vbKey1, vbKey2, vbKey3, vbKey4, vbKey5, vbKey6, vbKey7, vbKey8, vbKey9
            myTripControl_.EditTypeVD = "M"
    End Select
End Sub

Private Sub txtVDTransFee_Validate(Cancel As Boolean)
    If IsNumeric(txtVDTransFee.Text) Then
        myTripControl_.VDTranFee = txtVDTransFee.Text
    Else
        myTripControl_.VDTranFee = 0
    End If
End Sub

Private Sub txtVDWorker_Validate(Cancel As Boolean)
    If IsNumeric(txtVDWorker.Text) Then
        myTripControl_.VDWorker = txtVDWorker.Text
    Else
        myTripControl_.VDWorker = 0
    End If
End Sub

Private Sub txtVDWorkerFee_Change()
    If Me.Visible Then
        If Me.ActiveControl.Name = "txtVDWorkerFee" Then
            SumPayVendor
        End If
    End If
End Sub

Private Sub txtVDWorkerFee_Validate(Cancel As Boolean)
    If IsNumeric(txtVDWorkerFee.Text) Then
        myTripControl_.VDWorkerFee = txtVDWorkerFee.Text
    Else
        myTripControl_.VDWorkerFee = 0
    End If
End Sub

Private Sub SumPayVendor()
    Dim tmpAmt1 As Double, tmpAmt2 As Double, tmpAmt3 As Double, tmpAmt4 As Double, tmpAmt5 As Double
    If IsNumeric(txtVDTransFee.Text) Then
        tmpAmt1 = txtVDTransFee.Text
    Else
        tmpAmt1 = 0
    End If
    If IsNumeric(txtVDWorkerFee.Text) Then
        tmpAmt3 = txtVDWorkerFee.Text
    Else
        tmpAmt3 = 0
    End If
    If IsNumeric(txtVDEquipFee.Text) Then
        tmpAmt4 = txtVDEquipFee.Text
    Else
        tmpAmt4 = 0
    End If
    If IsNumeric(txtVDOtherFee.Text) Then
        tmpAmt5 = txtVDOtherFee.Text
    Else
        tmpAmt5 = 0
    End If
    txtAllVendor.Text = Format(tmpAmt1 + tmpAmt2 + tmpAmt3 + tmpAmt4 + tmpAmt5, "#,##0.00")
End Sub

Public Function Save() As Boolean
On Error GoTo ErrD
    Save = True
    
    myTripControl_.isChanged = True
    CmdSave_Click
    Exit Function
ErrD:
    Save = False
End Function
