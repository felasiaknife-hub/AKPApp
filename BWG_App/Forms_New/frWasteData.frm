VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmWasteData 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Waste Data"
   ClientHeight    =   9195
   ClientLeft      =   495
   ClientTop       =   330
   ClientWidth     =   11370
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9195
   ScaleWidth      =   11370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtQuoPrice 
      Height          =   345
      Left            =   2040
      TabIndex        =   172
      Top             =   3900
      Width           =   9255
   End
   Begin VB.CommandButton cmdReword 
      Caption         =   "ปรับ Waste"
      Height          =   795
      Left            =   2520
      TabIndex        =   169
      Top             =   8400
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ดูข้อมูล"
      Height          =   300
      Left            =   2160
      TabIndex        =   168
      Top             =   1080
      Width           =   800
   End
   Begin VB.ComboBox CboWasteYear 
      Height          =   315
      ItemData        =   "frWasteData.frx":0000
      Left            =   1440
      List            =   "frWasteData.frx":0002
      TabIndex        =   166
      Text            =   "2500"
      Top             =   1080
      Width           =   735
   End
   Begin VB.CommandButton cmdAccepdate 
      Caption         =   "กำหนดวันที่ตรวจรับ"
      Height          =   315
      Left            =   9780
      TabIndex        =   164
      Top             =   2160
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Frame Frame0 
      Height          =   1245
      Left            =   0
      TabIndex        =   155
      Top             =   -165
      Width           =   11295
      Begin VB.TextBox TxtMECode 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   6345
         Locked          =   -1  'True
         TabIndex        =   175
         TabStop         =   0   'False
         Top             =   900
         Width           =   1995
      End
      Begin VB.TextBox TxtRefME 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   2385
         Locked          =   -1  'True
         TabIndex        =   173
         TabStop         =   0   'False
         Top             =   900
         Width           =   1995
      End
      Begin VB.ComboBox cboWasteAnalysis 
         BackColor       =   &H0080FFFF&
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   7080
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   540
         Width           =   3855
      End
      Begin VB.TextBox txtWasteNo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000000C0&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1170
         Locked          =   -1  'True
         TabIndex        =   156
         TabStop         =   0   'False
         Text            =   "== AUTO =="
         Top             =   540
         Width           =   1455
      End
      Begin VB.TextBox txtEUCode1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   3630
         TabIndex        =   1
         Top             =   540
         Width           =   435
      End
      Begin VB.TextBox txtEUCode2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   4290
         TabIndex        =   2
         Top             =   540
         Width           =   435
      End
      Begin VB.TextBox txtEUCode3 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   4890
         TabIndex        =   3
         Top             =   540
         Width           =   435
      End
      Begin VB.TextBox txtCustomer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000000C0&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1170
         Locked          =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   180
         Width           =   4695
      End
      Begin VB.TextBox txtHAType 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   5430
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   540
         Width           =   435
      End
      Begin VB.ComboBox cboCompany 
         BackColor       =   &H0080FFFF&
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   7080
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   180
         Width           =   3855
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         Caption         =   "รหัสลูกค้า Better ME"
         Height          =   195
         Index           =   35
         Left            =   4860
         TabIndex        =   176
         Top             =   960
         Width           =   1410
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         Caption         =   "เลขที่ของเสีย (อ้างอิง Better ME)"
         Height          =   195
         Index           =   34
         Left            =   180
         TabIndex        =   174
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "การวิเคราะห์"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   33
         Left            =   6060
         TabIndex        =   163
         Top             =   600
         Width           =   885
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "เลขที่ของเสีย"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   162
         Top             =   600
         Width           =   885
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "รหัสของเสีย"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   2790
         TabIndex        =   161
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         Height          =   195
         Index           =   2
         Left            =   4170
         TabIndex        =   160
         Top             =   540
         Width           =   75
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         Height          =   195
         Index           =   3
         Left            =   4770
         TabIndex        =   159
         Top             =   540
         Width           =   75
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ชื่อลูกค้า"
         Height          =   195
         Index           =   24
         Left            =   510
         TabIndex        =   158
         Top             =   240
         Width           =   570
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "บริษัท"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   32
         Left            =   6540
         TabIndex        =   157
         Top             =   240
         Width           =   435
      End
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H000080FF&
      Caption         =   "ไม่มีตัวอย่าง"
      Height          =   195
      Left            =   4320
      TabIndex        =   8
      Top             =   1140
      Width           =   1155
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H000080FF&
      Caption         =   "มีตัวอย่าง"
      Height          =   195
      Left            =   3240
      TabIndex        =   7
      Top             =   1140
      Width           =   975
   End
   Begin VB.CommandButton cmdApprove 
      Caption         =   "ตรวจรับโดย Envi"
      Height          =   795
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   99
      Top             =   8400
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.TextBox txtIndustryType 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1755
      TabIndex        =   19
      Top             =   2130
      Width           =   5235
   End
   Begin VB.ComboBox cboLABRes 
      Height          =   315
      Left            =   6600
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   1080
      Width           =   1275
   End
   Begin VB.CommandButton Command1 
      Caption         =   "..."
      Height          =   285
      Left            =   11040
      TabIndex        =   18
      Top             =   1800
      Width           =   345
   End
   Begin VB.TextBox txtExpDate 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   9480
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   1800
      Width           =   1545
   End
   Begin VB.TextBox txtWasteGenNo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5940
      TabIndex        =   16
      Top             =   1800
      Width           =   1815
   End
   Begin VB.TextBox txtFactoryNo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1560
      TabIndex        =   15
      Top             =   1800
      Width           =   1695
   End
   Begin VB.TextBox txtContractNo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   9360
      TabIndex        =   14
      Top             =   1440
      Width           =   1575
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "ปิด"
      Height          =   795
      Left            =   10260
      Style           =   1  'Graphical
      TabIndex        =   104
      Top             =   8400
      Width           =   945
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "เก็บข้อมูล"
      Height          =   795
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   101
      Top             =   8400
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.CommandButton cmdReject 
      Caption         =   "ตีกลับ"
      Height          =   795
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   100
      Top             =   8400
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "ลบข้อมูล"
      Height          =   795
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   103
      Top             =   8400
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "สร้างใหม่"
      Height          =   795
      Left            =   8220
      Style           =   1  'Graphical
      TabIndex        =   102
      Top             =   8400
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.ComboBox cboSummary 
      Height          =   315
      ItemData        =   "frWasteData.frx":0004
      Left            =   8820
      List            =   "frWasteData.frx":0011
      Style           =   2  'Dropdown List
      TabIndex        =   98
      Top             =   7920
      Width           =   2385
   End
   Begin VB.TextBox txtEnviStaff 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   5370
      TabIndex        =   97
      Top             =   7950
      Width           =   2685
   End
   Begin VB.TextBox txtMktStaff 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1440
      TabIndex        =   96
      Top             =   7950
      Width           =   2685
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3495
      Left            =   60
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   4320
      Width           =   11265
      _ExtentX        =   19870
      _ExtentY        =   6165
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "คุณสมบัติของเสีย"
      TabPicture(0)   =   "frWasteData.frx":0042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "cboWasteChar"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cboWasteCharDesc"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "chkColor"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "chkOdour"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtOdour"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtColor"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "วิธีการจัดเก็บ"
      TabPicture(1)   =   "frWasteData.frx":005E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "chkBottle"
      Tab(1).Control(1)=   "txtBottle"
      Tab(1).Control(2)=   "chkPaperBox"
      Tab(1).Control(3)=   "txtPaperBox"
      Tab(1).Control(4)=   "chkPlasticGallon"
      Tab(1).Control(5)=   "txtPlasticGallon"
      Tab(1).Control(6)=   "chkPlasticDrum"
      Tab(1).Control(7)=   "txtPlasticDrum"
      Tab(1).Control(8)=   "chkPlasticBag"
      Tab(1).Control(9)=   "txtPlasticBag"
      Tab(1).Control(10)=   "cboTreatmentType"
      Tab(1).Control(11)=   "chkYard"
      Tab(1).Control(12)=   "chkTank"
      Tab(1).Control(13)=   "txtOther2"
      Tab(1).Control(14)=   "chkOther2"
      Tab(1).Control(15)=   "txtDrumDesc"
      Tab(1).Control(16)=   "chkDrum"
      Tab(1).Control(17)=   "txtBag"
      Tab(1).Control(18)=   "chkBag"
      Tab(1).Control(19)=   "cboIndoor"
      Tab(1).Control(20)=   "cboStoreType"
      Tab(1).Control(21)=   "Label7(5)"
      Tab(1).Control(22)=   "Label7(4)"
      Tab(1).Control(23)=   "Label7(3)"
      Tab(1).Control(24)=   "Label7(2)"
      Tab(1).Control(25)=   "Label7(1)"
      Tab(1).Control(26)=   "Label7(0)"
      Tab(1).Control(27)=   "Label2(31)"
      Tab(1).Control(28)=   "Label2(19)"
      Tab(1).Control(29)=   "Label2(18)"
      Tab(1).Control(30)=   "Label2(17)"
      Tab(1).ControlCount=   31
      TabCaption(2)   =   "วิธีการกำจัด"
      TabPicture(2)   =   "frWasteData.frx":007A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame4"
      Tab(2).Control(1)=   "Frame3"
      Tab(2).Control(2)=   "lvNewTreatment"
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "ไฟล์แนบ"
      TabPicture(3)   =   "frWasteData.frx":0096
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "cmdAttach(0)"
      Tab(3).Control(1)=   "cmdCls(0)"
      Tab(3).Control(2)=   "cmdOpenFile(0)"
      Tab(3).Control(3)=   "cmdOpenFile(1)"
      Tab(3).Control(4)=   "cmdOpenFile(2)"
      Tab(3).Control(5)=   "cmdOpenFile(3)"
      Tab(3).Control(6)=   "cmdCls(1)"
      Tab(3).Control(7)=   "cmdCls(2)"
      Tab(3).Control(8)=   "cmdCls(3)"
      Tab(3).Control(9)=   "cmdAttach(1)"
      Tab(3).Control(10)=   "cmdAttach(2)"
      Tab(3).Control(11)=   "cmdAttach(3)"
      Tab(3).Control(12)=   "OpenDialog"
      Tab(3).Control(13)=   "Label5"
      Tab(3).Control(14)=   "lblAttach(0)"
      Tab(3).Control(15)=   "Label6"
      Tab(3).Control(16)=   "Label8"
      Tab(3).Control(17)=   "Label10"
      Tab(3).Control(18)=   "lblAttach(1)"
      Tab(3).Control(19)=   "lblAttach(2)"
      Tab(3).Control(20)=   "lblAttach(3)"
      Tab(3).ControlCount=   21
      TabCaption(4)   =   "Tab 4"
      TabPicture(4)   =   "frWasteData.frx":00B2
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "txtOldStorePlace"
      Tab(4).Control(1)=   "cboMaxUnit"
      Tab(4).Control(2)=   "txtMaxQty"
      Tab(4).Control(3)=   "lvOldTreatment"
      Tab(4).Control(4)=   "Label2(20)"
      Tab(4).Control(5)=   "Label2(9)"
      Tab(4).ControlCount=   6
      Begin VB.Frame Frame4 
         Caption         =   "Analysis Type"
         Height          =   1395
         Left            =   -68160
         TabIndex        =   153
         Top             =   540
         Width           =   4095
         Begin VB.TextBox txtAnalysDesc 
            Appearance      =   0  'Flat
            ForeColor       =   &H00FF0000&
            Height          =   645
            Left            =   1200
            Locked          =   -1  'True
            ScrollBars      =   1  'Horizontal
            TabIndex        =   76
            Top             =   660
            Width           =   2715
         End
         Begin VB.OptionButton OptAnalysType3 
            Caption         =   "Others"
            Height          =   195
            Left            =   300
            TabIndex        =   75
            Top             =   720
            Width           =   795
         End
         Begin VB.OptionButton OptAnalysType2 
            Caption         =   "Finger Print"
            Height          =   195
            Left            =   2280
            TabIndex        =   74
            Top             =   360
            Width           =   1275
         End
         Begin VB.OptionButton OptAnalysType1 
            Caption         =   "Full Analysis"
            Height          =   195
            Left            =   300
            TabIndex        =   73
            Top             =   360
            Width           =   1275
         End
      End
      Begin VB.CheckBox chkBottle 
         Caption         =   "ขวดแก้ว : Bottle ขนาด"
         Height          =   255
         Left            =   -74280
         TabIndex        =   60
         Top             =   2040
         Width           =   1995
      End
      Begin VB.TextBox txtBottle 
         Appearance      =   0  'Flat
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   -71640
         Locked          =   -1  'True
         TabIndex        =   61
         Top             =   1980
         Width           =   1575
      End
      Begin VB.CheckBox chkPaperBox 
         Caption         =   "กล่องกระดาษ ขนาด"
         Height          =   255
         Left            =   -68820
         TabIndex        =   67
         Top             =   1650
         Width           =   2565
      End
      Begin VB.TextBox txtPaperBox 
         Appearance      =   0  'Flat
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   -66120
         Locked          =   -1  'True
         TabIndex        =   68
         Top             =   1620
         Width           =   1575
      End
      Begin VB.CheckBox chkPlasticGallon 
         Caption         =   "แกนลอนพลาสติก : Plastic Gallon ขนาด"
         Height          =   255
         Left            =   -74280
         TabIndex        =   58
         Top             =   1680
         Width           =   3135
      End
      Begin VB.TextBox txtPlasticGallon 
         Appearance      =   0  'Flat
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   -71100
         Locked          =   -1  'True
         TabIndex        =   59
         Top             =   1620
         Width           =   1035
      End
      Begin VB.CheckBox chkPlasticDrum 
         Caption         =   "ถังพลาสติก : Plastic Drum ขนาด"
         Height          =   255
         Left            =   -68820
         TabIndex        =   65
         Top             =   1290
         Width           =   2565
      End
      Begin VB.TextBox txtPlasticDrum 
         Appearance      =   0  'Flat
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   -66120
         Locked          =   -1  'True
         TabIndex        =   66
         Top             =   1260
         Width           =   1575
      End
      Begin VB.CheckBox chkPlasticBag 
         Caption         =   "ถุงพลาสติก : Plastic Bag ขนาด"
         Height          =   255
         Left            =   -68820
         TabIndex        =   63
         Top             =   930
         Width           =   2565
      End
      Begin VB.TextBox txtPlasticBag 
         Appearance      =   0  'Flat
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   -66120
         Locked          =   -1  'True
         TabIndex        =   64
         Top             =   900
         Width           =   1575
      End
      Begin VB.Frame Frame3 
         Caption         =   "เลือกเฉพาะ 075"
         Enabled         =   0   'False
         Height          =   1275
         Left            =   -68160
         TabIndex        =   146
         Top             =   1980
         Width           =   4095
         Begin VB.OptionButton OptBlend0 
            Caption         =   "ไม่เลือก"
            Height          =   195
            Left            =   300
            TabIndex        =   77
            Top             =   360
            Width           =   855
         End
         Begin VB.OptionButton OptBlend1 
            Caption         =   "ต้องทำ Waste Blending"
            Height          =   195
            Left            =   300
            TabIndex        =   78
            Top             =   660
            Width           =   2235
         End
         Begin VB.OptionButton OptBlend2 
            Caption         =   "ไม่ต้องทำ Waste Blending"
            Height          =   195
            Left            =   300
            TabIndex        =   79
            Top             =   960
            Width           =   2235
         End
      End
      Begin VB.CommandButton cmdAttach 
         Caption         =   "..."
         Height          =   315
         Index           =   0
         Left            =   -67440
         TabIndex        =   81
         Top             =   540
         Width           =   435
      End
      Begin VB.CommandButton cmdCls 
         Caption         =   "ลบไฟล์"
         Height          =   315
         Index           =   0
         Left            =   -65940
         TabIndex        =   83
         Top             =   540
         Width           =   795
      End
      Begin VB.CommandButton cmdOpenFile 
         Caption         =   "เปิดไฟล์"
         Height          =   315
         Index           =   0
         Left            =   -66900
         TabIndex        =   82
         Top             =   540
         Width           =   795
      End
      Begin VB.CommandButton cmdOpenFile 
         Caption         =   "เปิดไฟล์"
         Height          =   315
         Index           =   1
         Left            =   -66900
         TabIndex        =   86
         Top             =   900
         Width           =   795
      End
      Begin VB.CommandButton cmdOpenFile 
         Caption         =   "เปิดไฟล์"
         Height          =   315
         Index           =   2
         Left            =   -66900
         TabIndex        =   90
         Top             =   1260
         Width           =   795
      End
      Begin VB.CommandButton cmdOpenFile 
         Caption         =   "เปิดไฟล์"
         Height          =   315
         Index           =   3
         Left            =   -66900
         TabIndex        =   94
         Top             =   1620
         Width           =   795
      End
      Begin VB.CommandButton cmdCls 
         Caption         =   "ลบไฟล์"
         Height          =   315
         Index           =   1
         Left            =   -65940
         TabIndex        =   87
         Top             =   900
         Width           =   795
      End
      Begin VB.CommandButton cmdCls 
         Caption         =   "ลบไฟล์"
         Height          =   315
         Index           =   2
         Left            =   -65940
         TabIndex        =   91
         Top             =   1260
         Width           =   795
      End
      Begin VB.CommandButton cmdCls 
         Caption         =   "ลบไฟล์"
         Height          =   315
         Index           =   3
         Left            =   -65940
         TabIndex        =   95
         Top             =   1620
         Width           =   795
      End
      Begin VB.CommandButton cmdAttach 
         Caption         =   "..."
         Height          =   315
         Index           =   1
         Left            =   -67440
         TabIndex        =   85
         Top             =   900
         Width           =   435
      End
      Begin VB.CommandButton cmdAttach 
         Caption         =   "..."
         Height          =   315
         Index           =   2
         Left            =   -67440
         TabIndex        =   89
         Top             =   1260
         Width           =   435
      End
      Begin VB.CommandButton cmdAttach 
         Caption         =   "..."
         Height          =   315
         Index           =   3
         Left            =   -67440
         TabIndex        =   93
         Top             =   1620
         Width           =   435
      End
      Begin VB.TextBox txtOldStorePlace 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -73320
         TabIndex        =   138
         Top             =   1320
         Width           =   3165
      End
      Begin VB.ComboBox cboMaxUnit 
         Height          =   315
         Left            =   -72270
         Style           =   2  'Dropdown List
         TabIndex        =   137
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox txtMaxQty 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -73530
         TabIndex        =   136
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox txtColor 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   9060
         TabIndex        =   32
         Top             =   480
         Width           =   1995
      End
      Begin VB.TextBox txtOdour 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   5580
         TabIndex        =   30
         Top             =   480
         Width           =   2295
      End
      Begin VB.CheckBox chkOdour 
         Caption         =   "กลิ่น : Odour"
         Height          =   255
         Left            =   4320
         TabIndex        =   29
         Top             =   540
         Width           =   1275
      End
      Begin VB.CheckBox chkColor 
         Caption         =   "สี : Color"
         Height          =   255
         Left            =   8100
         TabIndex        =   31
         Top             =   540
         Width           =   915
      End
      Begin VB.ComboBox cboTreatmentType 
         Height          =   315
         ItemData        =   "frWasteData.frx":00CE
         Left            =   -66000
         List            =   "frWasteData.frx":00D8
         Style           =   2  'Dropdown List
         TabIndex        =   105
         Top             =   480
         Visible         =   0   'False
         Width           =   1995
      End
      Begin VB.CheckBox chkYard 
         Caption         =   "ลานกอง : Stock yard"
         Height          =   255
         Left            =   -68820
         TabIndex        =   69
         Top             =   2040
         Width           =   1845
      End
      Begin VB.CheckBox chkTank 
         Caption         =   "บ่อรวบรวม : Collecting Tank"
         Height          =   255
         Left            =   -74280
         TabIndex        =   62
         Top             =   2400
         Width           =   2355
      End
      Begin VB.TextBox txtOther2 
         Appearance      =   0  'Flat
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   -68040
         Locked          =   -1  'True
         TabIndex        =   71
         Top             =   2340
         Width           =   3495
      End
      Begin VB.CheckBox chkOther2 
         Caption         =   "อื่นๆ"
         Height          =   255
         Left            =   -68820
         TabIndex        =   70
         Top             =   2400
         Width           =   675
      End
      Begin VB.TextBox txtDrumDesc 
         Appearance      =   0  'Flat
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   -71640
         Locked          =   -1  'True
         TabIndex        =   57
         Top             =   1260
         Width           =   1575
      End
      Begin VB.CheckBox chkDrum 
         Caption         =   "ถังเหล็ก : Metal Drum ขนาด"
         Height          =   255
         Left            =   -74280
         TabIndex        =   56
         Top             =   1320
         Width           =   2475
      End
      Begin VB.TextBox txtBag 
         Appearance      =   0  'Flat
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   -71640
         Locked          =   -1  'True
         TabIndex        =   55
         Top             =   900
         Width           =   1575
      End
      Begin VB.CheckBox chkBag 
         Caption         =   "ถุงขนาดใหญ่ : Big Bag ขนาด"
         Height          =   255
         Left            =   -74280
         TabIndex        =   54
         Top             =   960
         Width           =   2475
      End
      Begin VB.ComboBox cboIndoor 
         Height          =   315
         Left            =   -70440
         Style           =   2  'Dropdown List
         TabIndex        =   53
         Top             =   480
         Width           =   3015
      End
      Begin VB.ComboBox cboStoreType 
         Height          =   315
         ItemData        =   "frWasteData.frx":00EE
         Left            =   -74280
         List            =   "frWasteData.frx":00F0
         Style           =   2  'Dropdown List
         TabIndex        =   52
         Top             =   480
         Width           =   3015
      End
      Begin VB.Frame Frame2 
         Caption         =   "คุณสมบัติของเสียอันตราย"
         Height          =   2535
         Left            =   120
         TabIndex        =   116
         Top             =   900
         Width           =   10935
         Begin WasteManagment.ctlGrid grKorOrNo 
            Height          =   1755
            Left            =   7200
            TabIndex        =   170
            Top             =   600
            Width           =   3555
            _ExtentX        =   6271
            _ExtentY        =   3096
         End
         Begin VB.TextBox txtCharDesc 
            Appearance      =   0  'Flat
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   2040
            TabIndex        =   154
            Top             =   0
            Visible         =   0   'False
            Width           =   1035
         End
         Begin VB.CheckBox chkHospital 
            Caption         =   "ขยะติดเชื้อ : Hospital Waste"
            Height          =   255
            Left            =   300
            TabIndex        =   49
            Top             =   2100
            Width           =   2295
         End
         Begin VB.TextBox txtOther 
            Appearance      =   0  'Flat
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   3780
            TabIndex        =   51
            Top             =   2040
            Width           =   3315
         End
         Begin VB.CheckBox chkOther 
            Caption         =   "อื่นๆ ระบุ"
            Height          =   255
            Left            =   2760
            TabIndex        =   50
            Top             =   2100
            Width           =   975
         End
         Begin VB.TextBox txtPesticide 
            Appearance      =   0  'Flat
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   2760
            TabIndex        =   48
            Top             =   1680
            Width           =   4335
         End
         Begin VB.CheckBox chkPesticide 
            Caption         =   "สารกำจัดศัตรูพืช/วัชพืช   ระบุ"
            Height          =   255
            Left            =   300
            TabIndex        =   47
            Top             =   1740
            Width           =   2415
         End
         Begin VB.TextBox txtRadioActive 
            Appearance      =   0  'Flat
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   2280
            TabIndex        =   46
            Top             =   1320
            Width           =   4815
         End
         Begin VB.CheckBox chkRadioActive 
            Caption         =   "กัมมันตภาพรังสี   ระบุ"
            Height          =   255
            Left            =   300
            TabIndex        =   45
            Top             =   1380
            Width           =   1815
         End
         Begin VB.ComboBox cboToxic 
            Height          =   315
            ItemData        =   "frWasteData.frx":00F2
            Left            =   2280
            List            =   "frWasteData.frx":00FC
            Style           =   2  'Dropdown List
            TabIndex        =   43
            Top             =   960
            Width           =   1875
         End
         Begin VB.TextBox txtToxicDesc 
            Appearance      =   0  'Flat
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   4560
            TabIndex        =   44
            Top             =   960
            Width           =   2535
         End
         Begin VB.CheckBox chkToxic 
            Caption         =   "มีสารเป็นพิษต่อร่างกาย"
            Height          =   255
            Left            =   300
            TabIndex        =   42
            Top             =   1020
            Width           =   1995
         End
         Begin VB.TextBox txtCancerDesc 
            Appearance      =   0  'Flat
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   5520
            TabIndex        =   41
            Top             =   600
            Width           =   1575
         End
         Begin VB.CheckBox chkCancer 
            Caption         =   "สารก่อมะเร็ง   ระบุ"
            Height          =   255
            Left            =   3900
            TabIndex        =   40
            Top             =   600
            Width           =   1575
         End
         Begin VB.TextBox txtReactDesc 
            Appearance      =   0  'Flat
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   1620
            TabIndex        =   39
            Top             =   600
            Width           =   1755
         End
         Begin VB.CheckBox chkReact 
            Caption         =   "ทำปฏิกริยากับ"
            Height          =   255
            Left            =   300
            TabIndex        =   38
            Top             =   660
            Width           =   1335
         End
         Begin VB.TextBox txtpH 
            Appearance      =   0  'Flat
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   9600
            TabIndex        =   37
            Top             =   240
            Width           =   975
         End
         Begin VB.CheckBox chkCorrosive 
            Caption         =   "กัดกร่อน    pH"
            Height          =   255
            Left            =   8160
            TabIndex        =   36
            Top             =   240
            Width           =   1335
         End
         Begin VB.TextBox txtFlashPoint 
            Appearance      =   0  'Flat
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   6720
            TabIndex        =   35
            Top             =   180
            Width           =   435
         End
         Begin VB.CheckBox chkFlame 
            Caption         =   "ติดไฟ  จุดวาบไฟ"
            Height          =   255
            Left            =   5160
            TabIndex        =   34
            Top             =   240
            Width           =   1575
         End
         Begin VB.CheckBox chkExplode 
            Caption         =   "ระเบิดได้"
            Height          =   255
            Left            =   300
            TabIndex        =   33
            Top             =   300
            Width           =   1095
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ระบุ"
            Height          =   195
            Index           =   16
            Left            =   4200
            TabIndex        =   119
            Top             =   1020
            Width           =   315
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "C"
            Height          =   195
            Index           =   15
            Left            =   7380
            TabIndex        =   118
            Top             =   300
            Width           =   135
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "o"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   150
            Index           =   14
            Left            =   7260
            TabIndex        =   117
            Top             =   180
            Width           =   75
         End
      End
      Begin VB.ComboBox cboWasteCharDesc 
         Height          =   315
         ItemData        =   "frWasteData.frx":0129
         Left            =   2280
         List            =   "frWasteData.frx":012B
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   480
         Width           =   1695
      End
      Begin VB.ComboBox cboWasteChar 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   480
         Width           =   1935
      End
      Begin MSComctlLib.ListView lvOldTreatment 
         Height          =   1335
         Left            =   -72660
         TabIndex        =   139
         Top             =   1860
         Visible         =   0   'False
         Width           =   5085
         _ExtentX        =   8969
         _ExtentY        =   2355
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
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "วิธีการกำจัดเดิม"
            Object.Width           =   7320
         EndProperty
      End
      Begin MSComDlg.CommonDialog OpenDialog 
         Left            =   -66600
         Top             =   2220
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DialogTitle     =   "เลือกไฟล์"
      End
      Begin MSComctlLib.ListView lvNewTreatment 
         Height          =   2655
         Left            =   -74760
         TabIndex        =   72
         Top             =   600
         Width           =   6405
         _ExtentX        =   11298
         _ExtentY        =   4683
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
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "วิธีการกำจัดที่ควรใช้"
            Object.Width           =   11298
         EndProperty
      End
      Begin VB.Label Label7 
         Caption         =   "ลิตร"
         Height          =   195
         Index           =   5
         Left            =   -69900
         TabIndex        =   152
         Top             =   2040
         Width           =   315
      End
      Begin VB.Label Label7 
         Caption         =   "ลิตร"
         Height          =   195
         Index           =   4
         Left            =   -69900
         TabIndex        =   151
         Top             =   1680
         Width           =   315
      End
      Begin VB.Label Label7 
         Caption         =   "ลิตร"
         Height          =   195
         Index           =   3
         Left            =   -64380
         TabIndex        =   150
         Top             =   1320
         Width           =   315
      End
      Begin VB.Label Label7 
         Caption         =   "ลิตร"
         Height          =   195
         Index           =   2
         Left            =   -69900
         TabIndex        =   149
         Top             =   1320
         Width           =   315
      End
      Begin VB.Label Label7 
         Caption         =   "กก."
         Height          =   195
         Index           =   1
         Left            =   -64380
         TabIndex        =   148
         Top             =   960
         Width           =   315
      End
      Begin VB.Label Label7 
         Caption         =   "กก."
         Height          =   195
         Index           =   0
         Left            =   -69900
         TabIndex        =   147
         Top             =   960
         Width           =   315
      End
      Begin VB.Label Label5 
         Caption         =   "ไฟล์แนบที่ 1"
         Height          =   195
         Left            =   -74760
         TabIndex        =   145
         Top             =   600
         Width           =   975
      End
      Begin VB.Label lblAttach 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   -73740
         TabIndex        =   80
         Top             =   600
         Width           =   6135
      End
      Begin VB.Label Label6 
         Caption         =   "ไฟล์แนบที่ 2"
         Height          =   195
         Left            =   -74760
         TabIndex        =   144
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label8 
         Caption         =   "ไฟล์แนบที่ 3"
         Height          =   195
         Left            =   -74760
         TabIndex        =   143
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label10 
         Caption         =   "ไฟล์แนบที่ 4"
         Height          =   195
         Left            =   -74760
         TabIndex        =   142
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label lblAttach 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   -73740
         TabIndex        =   84
         Top             =   960
         Width           =   6135
      End
      Begin VB.Label lblAttach 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   -73740
         TabIndex        =   88
         Top             =   1320
         Width           =   6135
      End
      Begin VB.Label lblAttach 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   -73740
         TabIndex        =   92
         Top             =   1680
         Width           =   6135
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "สถานที่กำจัดเดิม"
         Height          =   195
         Index           =   20
         Left            =   -74640
         TabIndex        =   141
         Top             =   1350
         Width           =   1155
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "เก็บได้มากที่สุด"
         Height          =   195
         Index           =   9
         Left            =   -74640
         TabIndex        =   140
         Top             =   510
         Width           =   1065
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "รูปแบบการบำบัด"
         Height          =   195
         Index           =   31
         Left            =   -67260
         TabIndex        =   133
         Top             =   540
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "10.3)"
         Height          =   195
         Index           =   19
         Left            =   -74760
         TabIndex        =   122
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "10.2)"
         Height          =   195
         Index           =   18
         Left            =   -70920
         TabIndex        =   121
         Top             =   540
         Width           =   375
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "10.1)"
         Height          =   195
         Index           =   17
         Left            =   -74760
         TabIndex        =   120
         Top             =   540
         Width           =   375
      End
   End
   Begin VB.TextBox txtGenerateDesc 
      Appearance      =   0  'Flat
      Height          =   615
      Left            =   1440
      MultiLine       =   -1  'True
      TabIndex        =   25
      Top             =   3240
      Width           =   9855
   End
   Begin VB.Frame Frame1 
      Caption         =   "ปริมาณของเสีย"
      Height          =   735
      Left            =   60
      TabIndex        =   110
      Top             =   2460
      Width           =   11235
      Begin VB.ComboBox cboTimeRate 
         Height          =   315
         ItemData        =   "frWasteData.frx":012D
         Left            =   7290
         List            =   "frWasteData.frx":012F
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   300
         Width           =   795
      End
      Begin VB.ComboBox cboUnitRate 
         Height          =   315
         Left            =   6360
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   300
         Width           =   735
      End
      Begin VB.TextBox txtRate 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5310
         TabIndex        =   22
         Top             =   300
         Width           =   1005
      End
      Begin VB.ComboBox cboQtyUnit 
         Height          =   315
         Left            =   2700
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   300
         Width           =   735
      End
      Begin VB.TextBox txtQty 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         TabIndex        =   20
         Top             =   300
         Width           =   1215
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "/"
         Height          =   195
         Index           =   11
         Left            =   7140
         TabIndex        =   113
         Top             =   330
         Width           =   75
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "อัตราการเกิดของเสีย"
         Height          =   195
         Index           =   10
         Left            =   3810
         TabIndex        =   112
         Top             =   330
         Width           =   1440
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "จัดเก็บ ณ ปัจจุบัน"
         Height          =   195
         Index           =   8
         Left            =   150
         TabIndex        =   111
         Top             =   330
         Width           =   1215
      End
   End
   Begin VB.TextBox txtWasteName 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1380
      TabIndex        =   13
      Top             =   1440
      Width           =   6975
   End
   Begin VB.ComboBox cboWasteType 
      Height          =   315
      Left            =   9360
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   1080
      Width           =   1575
   End
   Begin VB.TextBox txtKor 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   11160
      TabIndex        =   12
      Top             =   1500
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.TextBox txtMoud 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   11160
      TabIndex        =   11
      Top             =   1140
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Label Label12 
      Caption         =   "หมายเหตุ: ใบเสนอราคา"
      Height          =   255
      Left            =   180
      TabIndex        =   171
      Top             =   3960
      Width           =   1755
   End
   Begin VB.Label Label11 
      BackColor       =   &H00C0E0FF&
      Caption         =   " Waste   ปี :"
      Height          =   255
      Left            =   480
      TabIndex        =   167
      Top             =   1140
      Width           =   855
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C0FFFF&
      Caption         =   "กรุณาตรวจสอบ"
      Height          =   195
      Left            =   3300
      TabIndex        =   165
      Top             =   1860
      Width           =   1155
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "30/04/2551"
      BeginProperty Font 
         Name            =   "EucrosiaUPC"
         Size            =   18
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   3840
      TabIndex        =   135
      Top             =   8760
      Visible         =   0   'False
      Width           =   3075
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "Caption"
      BeginProperty Font 
         Name            =   "EucrosiaUPC"
         Size            =   18
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   435
      Left            =   3840
      TabIndex        =   134
      Top             =   8400
      Visible         =   0   'False
      Width           =   3075
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   7020
      TabIndex        =   132
      Top             =   2160
      Visible         =   0   'False
      Width           =   2685
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ประเภทอุตสาหกรรม"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   30
      Left            =   300
      TabIndex        =   131
      Top             =   2160
      Width           =   1410
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ผล LAB"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   29
      Left            =   5880
      TabIndex        =   130
      Top             =   1140
      Width           =   555
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "วันที่ใบอนุญาตหมดอายุ"
      Height          =   195
      Index           =   28
      Left            =   7800
      TabIndex        =   129
      Top             =   1860
      Width           =   1620
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "เลขที่ผู้กำเนิดของเสีย"
      Height          =   195
      Index           =   27
      Left            =   4440
      TabIndex        =   128
      Top             =   1860
      Width           =   1470
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "เลขทะเบียนโรงงาน"
      Height          =   195
      Index           =   26
      Left            =   180
      TabIndex        =   127
      Top             =   1800
      Width           =   1350
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "เลขที่สัญญา"
      Height          =   195
      Index           =   25
      Left            =   8460
      TabIndex        =   126
      Top             =   1500
      Width           =   795
   End
   Begin VB.Line Line1 
      X1              =   30
      X2              =   11250
      Y1              =   8310
      Y2              =   8310
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ผลสรุป"
      Height          =   195
      Index           =   23
      Left            =   8220
      TabIndex        =   125
      Top             =   7980
      Width           =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "เจ้าหน้าที่ Envi"
      Height          =   195
      Index           =   22
      Left            =   4290
      TabIndex        =   124
      Top             =   7980
      Width           =   1005
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "เจ้าหน้าที่การตลาด"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   21
      Left            =   90
      TabIndex        =   123
      Top             =   7980
      Width           =   1290
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ทำให้เกิดของเสีย"
      Height          =   195
      Index           =   13
      Left            =   180
      TabIndex        =   115
      Top             =   3540
      Width           =   1185
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "กระบวนการที่"
      Height          =   195
      Index           =   12
      Left            =   420
      TabIndex        =   114
      Top             =   3300
      Width           =   960
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ชื่อกากของเสีย"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   7
      Left            =   360
      TabIndex        =   109
      Top             =   1500
      Width           =   1035
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ประเภทของเสีย"
      Height          =   195
      Index           =   6
      Left            =   8040
      TabIndex        =   108
      Top             =   1140
      Width           =   1095
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ข้อ"
      Height          =   195
      Index           =   5
      Left            =   3960
      TabIndex        =   107
      Top             =   1140
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "หมวด"
      Height          =   195
      Index           =   4
      Left            =   4800
      TabIndex        =   106
      Top             =   1140
      Visible         =   0   'False
      Width           =   435
   End
End
Attribute VB_Name = "frmWasteData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim callerForm_ As Form
Dim curCustomer_ As BWGCustomer
Dim curEUCode1_ As BWGEUCode1
Dim curEUCode2_ As BWGEUCode2
Dim curEUCode3_ As BWGEUCode3
Dim curMktStaff_ As BWGCompanyStaff
Dim curEnviStaff_ As BWGCompanyStaff
Dim myWasteDataCR_ As BWGWasteDataCR
Dim myWasteDataLab_ As BWGWasteDataLAB
Dim colTreatment_ As Collection
Dim callStaffType_ As String
Dim curIndustryType_ As BWGIndustryType
Dim StrGrp As String
Dim isNew As Boolean
Dim isEnv As Boolean
Private Type AttacheFile
    FileName As String
    FullPath As String
End Type
Dim Attaches() As AttacheFile
Dim isSave As Boolean
Dim WasteID As String
Dim WasteY As String
Public Accepdate As String

Private Function AttacheFiles(ByVal Index As Integer, ByRef FileName As String, ByRef FullPath As String) As String
Dim StrFileName As String
Dim StrPath As String, StrFullPath As String
    OpenDialog.Filter = "PDF Files (*.pdf)|*.pdf|Pictures (*.jpg, *.bmp, *.gif, *.png| *.jpg; *.bmp; *.gif; *.png|" & _
    "World Document (*.doc)|*.doc|Zip (*.rar)|Work Sheet (*.xls)|*.xls|Visio (*.vsd)|*.vsd"
    OpenDialog.FileName = ""
    OpenDialog.ShowOpen
    If OpenDialog.FileName = "" Then
        AttacheFiles = ""
        Exit Function
    End If
    StrFullPath = OpenDialog.FileName
    StrFileName = OpenDialog.FileTitle
    StrPath = Mid(StrFullPath, 1, Len(StrFullPath) - Len(StrFileName))
    
    AttacheFiles = StrFileName
    FileName = StrFileName
    FullPath = StrFullPath
End Function

Private Function Chk10Detail() As Boolean
    Chk10Detail = False
    If chkBag.Value = 1 Then Chk10Detail = True: Exit Function
    If chkDrum.Value = 1 Then Chk10Detail = True: Exit Function
    If chkYard.Value = 1 Then Chk10Detail = True: Exit Function
    If chkOther2.Value = 1 Then Chk10Detail = True: Exit Function
    If chkTank.Value = 1 Then Chk10Detail = True: Exit Function
    If chkPlasticBag.Value = 1 Then Chk10Detail = True: Exit Function
    If chkPlasticDrum.Value = 1 Then Chk10Detail = True: Exit Function
    If chkPlasticGallon.Value = 1 Then Chk10Detail = True: Exit Function
    If chkPaperBox.Value = 1 Then Chk10Detail = True: Exit Function
    If chkBottle.Value = 1 Then Chk10Detail = True: Exit Function
End Function

Public Sub setAccepDate(ByVal Accdate As String)
    Label1.Caption = Label1.Caption + " " + Accdate
    myWasteDataCR_.isChanged = True
End Sub

Private Sub cboCompany_Click()
On Error Resume Next
    If CboCompany.ListIndex >= 0 Then
        myWasteDataCR_.CompanyID = CurrentUser.WasteCompany(CboCompany.ListIndex + cboIndex).ID
    Else
        myWasteDataCR_.CompanyID = ""
    End If
    myWasteDataCR_.isChanged = True
End Sub

Private Sub cboCompany_Validate(Cancel As Boolean)
    If CboCompany.ListIndex >= 0 Then
        If CboCompany.ListCount > 1 Then
            If CboCompany.ListIndex = 0 Then Exit Sub
            myWasteDataCR_.CompanyID = CurrentUser.WorkCompanies(CboCompany.ListIndex).ID
        Else
            myWasteDataCR_.CompanyID = CurrentUser.WorkCompanies(CboCompany.ListIndex + cboIndex).ID
        End If
    Else
        myWasteDataCR_.CompanyID = ""
    End If
    myWasteDataCR_.isChanged = True
End Sub


Private Sub cboIndoor_Click()
On Error Resume Next
    Select Case cboIndoor.ListIndex
        Case 0
            myWasteDataLab_.InOrOutDoor = ""
        Case 1
            myWasteDataLab_.InOrOutDoor = "I" 'In door
        Case 2
            myWasteDataLab_.InOrOutDoor = "O" 'Out door
    End Select
End Sub

Private Sub cboIndoor_Validate(Cancel As Boolean)
    Select Case cboIndoor.ListIndex
        Case 0
            myWasteDataLab_.InOrOutDoor = ""
        Case 1
            myWasteDataLab_.InOrOutDoor = "I" 'In door
        Case 2
            myWasteDataLab_.InOrOutDoor = "O" 'Out door
    End Select
End Sub

Private Sub cboLABRes_Click()
On Error Resume Next
    If cboLABRes.ListIndex = 0 Then
        myWasteDataCR_.isLabResultNeeded = False
    Else
        myWasteDataCR_.isLabResultNeeded = True
    End If
End Sub

Private Sub cboLABRes_Validate(Cancel As Boolean)
    If cboLABRes.ListIndex = 0 Then
        myWasteDataCR_.isLabResultNeeded = False
    Else
        myWasteDataCR_.isLabResultNeeded = True
    End If
End Sub

Private Sub cboMaxUnit_Validate(Cancel As Boolean)
    Select Case cboMaxUnit.ListIndex
        Case 0
            myWasteDataCR_.MaxUnit = "T" 'ตัน
        Case 1
            myWasteDataCR_.MaxUnit = "L" 'ลิตร
    End Select
End Sub

Private Sub cboQtyUnit_Click()
On Error Resume Next
    Select Case cboQtyUnit.ListIndex
        Case 0
            myWasteDataCR_.NormalUnit = "T" 'ตัน
        Case 1
            myWasteDataCR_.NormalUnit = "L" 'ลิตร
    End Select
End Sub

Private Sub cboQtyUnit_Validate(Cancel As Boolean)
    Select Case cboQtyUnit.ListIndex
        Case 0
            myWasteDataCR_.NormalUnit = "T" 'ตัน
        Case 1
            myWasteDataCR_.NormalUnit = "L" 'ลิตร
    End Select
End Sub

Private Sub cboStoreType_Click()
On Error Resume Next
    Select Case cboStoreType.ListIndex
        Case 0
            myWasteDataLab_.MixOrSeperate = ""
        Case 1
            myWasteDataLab_.MixOrSeperate = "S"  'Separate
        Case 2
            myWasteDataLab_.MixOrSeperate = "M"  'Mix
    End Select
End Sub

Private Sub cboStoreType_Validate(Cancel As Boolean)
    Select Case cboStoreType.ListIndex
        Case 0
            myWasteDataLab_.MixOrSeperate = ""
        Case 1
            myWasteDataLab_.MixOrSeperate = "S"  'Separate
        Case 2
            myWasteDataLab_.MixOrSeperate = "M"  'Mix
    End Select
End Sub

Private Sub cboSummary_Click()
On Error Resume Next
    Select Case cboSummary.ListIndex
        Case 0
            myWasteDataCR_.WasteApproveStatus = "P" 'Pending
            myWasteDataCR_.Status = "P"
        Case 1
            myWasteDataCR_.WasteApproveStatus = "E" 'Accept
            myWasteDataCR_.Status = "A"
            myWasteDataCR_.isTreatAble = True
        Case 2
            myWasteDataCR_.WasteApproveStatus = "R" 'Reject
            myWasteDataCR_.Status = "E"
            myWasteDataCR_.isTreatAble = False
    End Select
End Sub

Private Sub cboSummary_Validate(Cancel As Boolean)
    Select Case cboSummary.ListIndex
        Case 0
            myWasteDataCR_.WasteApproveStatus = "P" 'Pending
            myWasteDataCR_.Status = "P"
        Case 1
            myWasteDataCR_.WasteApproveStatus = "E" 'Except
            myWasteDataCR_.Status = "A"
            myWasteDataCR_.isTreatAble = True
        Case 2
            myWasteDataCR_.WasteApproveStatus = "R" 'Reject
            myWasteDataCR_.Status = "E"
            myWasteDataCR_.isTreatAble = False
    End Select
End Sub

Private Sub cboTimeRate_Click()
On Error Resume Next
    Select Case cboTimeRate.ListIndex
        Case 0
            myWasteDataCR_.WasteRateTime = "M" 'ตัน
        Case 1
            myWasteDataCR_.WasteRateTime = "Y" 'ลิตร
    End Select
End Sub

Private Sub cboTimeRate_Validate(Cancel As Boolean)
    Select Case cboTimeRate.ListIndex
        Case 0
            myWasteDataCR_.WasteRateTime = "M" 'ตัน
        Case 1
            myWasteDataCR_.WasteRateTime = "Y" 'ลิตร
    End Select
End Sub

Private Sub cboToxic_Click()
On Error Resume Next
    Select Case cboToxic.ListIndex
        Case -1
            myWasteDataLab_.ToxicType = "-"
        Case 0
            myWasteDataLab_.ToxicType = "A" 'ออกฤทธิ์เฉียบพลัน
        Case 1
            myWasteDataLab_.ToxicType = "C" 'ออกฤทธิ์ไม่เฉียบพลัน
    End Select
End Sub

Private Sub cboToxic_Validate(Cancel As Boolean)
    Select Case cboToxic.ListIndex
        Case -1
            myWasteDataLab_.ToxicType = "-"
        Case 0
            myWasteDataLab_.ToxicType = "A" 'ออกฤทธิ์เฉียบพลัน
        Case 1
            myWasteDataLab_.ToxicType = "C" 'ออกฤทธิ์ไม่เฉียบพลัน
    End Select
End Sub

Private Sub cboTreatmentType_Click()
On Error Resume Next

End Sub

Private Sub cboTreatmentType_Validate(Cancel As Boolean)
    Select Case cboTreatmentType.ListIndex
        Case -1
            myWasteDataCR_.TreatmentType = ""
        Case 0
            myWasteDataCR_.TreatmentType = "B"  'Blending
        Case 1
            myWasteDataCR_.TreatmentType = "L" 'Landfill
    End Select
End Sub

Private Sub cboUnitRate_Click()
On Error Resume Next
    Select Case cboUnitRate.ListIndex
        Case 0
            myWasteDataCR_.WasteRateUnit = "T" 'ตัน
        Case 1
            myWasteDataCR_.WasteRateUnit = "L" 'ลิตร
    End Select
End Sub

Private Sub cboUnitRate_Validate(Cancel As Boolean)
    Select Case cboUnitRate.ListIndex
        Case 0
            myWasteDataCR_.WasteRateUnit = "T" 'ตัน
        Case 1
            myWasteDataCR_.WasteRateUnit = "L" 'ลิตร
    End Select
End Sub

Private Sub cboWasteAnalysis_Validate(Cancel As Boolean)
    If cboWasteAnalysis.ListIndex = 0 Then
        myWasteDataCR_.isQuotationReq = False
        If isEnv = False Then
'            cmdApprove.Visible = True
        End If
    ElseIf cboWasteAnalysis.ListIndex = 1 Then
        myWasteDataCR_.isQuotationReq = True
        If isEnv = False Then
            cmdApprove.Visible = False
        Else
'            cmdApprove.Visible=true
        End If
    End If
    myWasteDataCR_.isChanged = True
End Sub


Private Sub cboWasteChar_Click()
On Error Resume Next
    Select Case cboWasteChar.ListIndex
        Case 0
            cboWasteCharDesc.Clear
            cboWasteCharDesc.Enabled = False
        Case 1
            With cboWasteCharDesc
                .Clear
                .Enabled = True
                .AddItem "เป็นผง"
                .AddItem "เป็นก้อน"
                .AddItem "เป็นเส้น"
                .AddItem "อื่น ๆ"
            End With
        Case 2
            With cboWasteCharDesc
                .Clear
                .Enabled = True
                .AddItem "สามารถปั๊มได้"
                .AddItem "ไม่สามารถปั๊มได้"
            End With
        Case 3
            cboWasteCharDesc.Clear
            cboWasteCharDesc.Enabled = False
    End Select
    
    Select Case cboWasteChar.ListIndex
        Case 0
            myWasteDataLab_.WasteCharacter = "NA"
        Case 1
            myWasteDataLab_.WasteCharacter = "SO" 'Solid
        Case 2
            myWasteDataLab_.WasteCharacter = "SL" 'Sludge
        Case 3
            myWasteDataLab_.WasteCharacter = "LI" 'Liquid
    End Select
    myWasteDataLab_.isChanged = True
End Sub

Private Sub cboWasteChar_Validate(Cancel As Boolean)
    Select Case cboWasteChar.ListIndex
        Case 0
            myWasteDataLab_.WasteCharacter = "NA"
        Case 1
            myWasteDataLab_.WasteCharacter = "SO" 'Solid
        Case 2
            myWasteDataLab_.WasteCharacter = "SL" 'Sludge
        Case 3
            myWasteDataLab_.WasteCharacter = "LI" 'Liquid
    End Select
    myWasteDataLab_.isChanged = True
End Sub

Private Sub cboWasteCharDesc_Click()
On Error Resume Next
    txtCharDesc.Visible = False
    If cboWasteChar.ListIndex = 1 Then
        If cboWasteCharDesc.ListIndex = 0 Then
            myWasteDataLab_.PowderOrLump = "P"
        ElseIf cboWasteCharDesc.ListIndex = 1 Then
            myWasteDataLab_.PowderOrLump = "L"
        ElseIf cboWasteCharDesc.ListIndex = 2 Then
            myWasteDataLab_.PowderOrLump = "N"
        ElseIf cboWasteCharDesc.ListIndex = 3 Then
            myWasteDataLab_.PowderOrLump = "O"
            txtCharDesc.Visible = True
            myWasteDataLab_.CharDescOther = Trim(txtCharDesc)
        End If
    ElseIf cboWasteChar.ListIndex = 2 Then
        If cboWasteCharDesc.ListIndex = 0 Then
            myWasteDataLab_.isPumpAble = True
        ElseIf cboWasteCharDesc.ListIndex = 1 Then
            myWasteDataLab_.isPumpAble = False
        End If
    End If
    myWasteDataLab_.isChanged = True
End Sub

Private Sub cboWasteCharDesc_Validate(Cancel As Boolean)
    If cboWasteChar.ListIndex = 1 Then
        If cboWasteCharDesc.ListIndex = 0 Then
            myWasteDataLab_.PowderOrLump = "P"
        ElseIf cboWasteCharDesc.ListIndex = 1 Then
            myWasteDataLab_.PowderOrLump = "L"
        End If
    ElseIf cboWasteChar.ListIndex = 2 Then
        If cboWasteCharDesc.ListIndex = 0 Then
            myWasteDataLab_.isPumpAble = True
        ElseIf cboWasteCharDesc.ListIndex = 1 Then
            myWasteDataLab_.isPumpAble = False
        End If
    End If
    myWasteDataLab_.isChanged = True
End Sub

Private Sub cboWasteType_Click()
On Error Resume Next
    If cboWasteType.ListIndex = 0 Then
        myWasteDataCR_.WasteType = "N/A"
    Else
        myWasteDataCR_.WasteType = cboWasteType.Text
    End If
End Sub

Private Sub cboWasteType_Validate(Cancel As Boolean)
    If cboWasteType.ListIndex = 0 Then
        myWasteDataCR_.WasteType = "N/A"
    Else
        myWasteDataCR_.WasteType = cboWasteType.Text
    End If
End Sub

Private Sub CboWasteYear_Validate(Cancel As Boolean)
myWasteDataCR_.isChanged = True
End Sub

Private Sub chkBag_Click()
On Error Resume Next
    txtBag.Locked = chkBag.Value = 0
    If chkBag.Value = 1 Then
        If Me.Visible = True Then
            If Not Me.ActiveControl Is Nothing Then
               If Me.ActiveControl.Name = "chkBag" Then chkBag_Validate False
            End If
            txtBag.SetFocus
        End If
        chkDrum.Value = 0: txtDrumDesc.Text = ""
        chkYard.Value = 0
        chkOther2.Value = 0: txtOther2.Text = ""
        chkTank.Value = 0
        chkPlasticBag.Value = 0: txtPlasticBag.Text = ""
        chkPlasticDrum.Value = 0: txtPlasticDrum.Text = ""
        chkPlasticGallon.Value = 0: txtPlasticGallon.Text = ""
        chkPaperBox.Value = 0: txtPaperBox.Text = ""
        chkBottle.Value = 0: txtBottle.Text = ""
        myWasteDataLab_.StorageType = "BA" 'Bag
    Else
        myWasteDataLab_.StorageType = ""
    End If
End Sub

Private Sub chkBag_Validate(Cancel As Boolean)
    If chkBag.Value = 1 Then
        myWasteDataLab_.StorageType = "BA" 'Bag
        myWasteDataLab_.StorageDesc = Trim(txtBag.Text)
    End If
End Sub

Private Sub chkBottle_Click()
On Error Resume Next
    txtBottle.Locked = chkBottle.Value = 0
    If chkBottle.Value = 1 Then
        txtBottle.SetFocus
        chkBag.Value = 0: txtBag.Text = ""
        chkDrum.Value = 0: txtDrumDesc.Text = ""
        chkYard.Value = 0
        chkOther2.Value = 0: txtOther2.Text = ""
        chkTank.Value = 0
        chkPlasticBag.Value = 0: txtPlasticBag.Text = ""
        chkPlasticDrum.Value = 0: txtPlasticDrum.Text = ""
        chkPlasticGallon.Value = 0: txtPlasticGallon.Text = ""
        chkPaperBox.Value = 0: txtPaperBox.Text = ""
    End If
    myWasteDataLab_.isBottle = chkBottle.Value = 1
End Sub

Private Sub chkBottle_Validate(Cancel As Boolean)
    myWasteDataLab_.isBottle = chkBottle.Value = 1
End Sub

Private Sub chkCancer_Validate(Cancel As Boolean)
    myWasteDataLab_.isCarcinoGenic = chkCancer.Value = 1
End Sub

Private Sub chkColor_Validate(Cancel As Boolean)
    myWasteDataLab_.isColor = chkColor.Value = 1
End Sub

Private Sub chkCorrosive_Validate(Cancel As Boolean)
    myWasteDataLab_.isCorrosive = chkCorrosive.Value = 1
End Sub

Private Sub chkDrum_Click()
On Error Resume Next
    txtDrumDesc.Locked = chkDrum.Value = 0
    If chkDrum.Value = 1 Then
        If Me.Visible = True Then
            If Not Me.ActiveControl Is Nothing Then
               If Me.ActiveControl.Name = "chkDrum" Then chkDrum_Validate False
            End If
            txtDrumDesc.SetFocus
        End If
        chkBag.Value = 0: txtBag.Text = ""
        chkYard.Value = 0
        chkOther2.Value = 0: txtOther2.Text = ""
        chkTank.Value = 0
        chkPlasticBag.Value = 0: txtPlasticBag.Text = ""
        chkPlasticDrum.Value = 0: txtPlasticDrum.Text = ""
        chkPlasticGallon.Value = 0: txtPlasticGallon.Text = ""
        chkPaperBox.Value = 0: txtPaperBox.Text = ""
        chkBottle.Value = 0: txtBottle.Text = ""
        myWasteDataLab_.StorageType = "DR" 'Drum
    Else
        myWasteDataLab_.StorageType = ""
    End If
End Sub

Private Sub chkDrum_Validate(Cancel As Boolean)
    If chkDrum.Value = 1 Then
        myWasteDataLab_.StorageType = "DR" 'Drum
        myWasteDataLab_.StorageDesc = Trim(txtDrumDesc.Text)
    End If
End Sub

Private Sub chkExplode_Validate(Cancel As Boolean)
    myWasteDataLab_.isExplode = chkExplode.Value = 1
End Sub

Private Sub chkFlame_Validate(Cancel As Boolean)
    myWasteDataLab_.isFlameAble = chkFlame.Value = 1
End Sub

Private Sub chkHospital_Validate(Cancel As Boolean)
    myWasteDataLab_.isHospital = chkHospital.Value = 1
End Sub

Private Sub chkOdour_Validate(Cancel As Boolean)
    myWasteDataLab_.isOdour = chkOdour.Value = 1
End Sub

Private Sub chkOther2_Click()
On Error Resume Next
    txtOther2.Locked = chkOther2.Value = 0
    If chkOther2.Value = 1 Then
        If Me.Visible = True Then
            If Not Me.ActiveControl Is Nothing Then
               If Me.ActiveControl.Name = "chkOther2" Then chkOther2_Validate False
            End If
            txtOther2.SetFocus
        End If
        chkBag.Value = 0: txtBag.Text = ""
        chkDrum.Value = 0: txtDrumDesc.Text = ""
        chkYard.Value = 0
        chkTank.Value = 0
        chkPlasticBag.Value = 0: txtPlasticBag.Text = ""
        chkPlasticDrum.Value = 0: txtPlasticDrum.Text = ""
        chkPlasticGallon.Value = 0: txtPlasticGallon.Text = ""
        chkPaperBox.Value = 0: txtPaperBox.Text = ""
        chkBottle.Value = 0: txtBottle.Text = ""
        myWasteDataLab_.StorageType = "OT" 'Other
    Else
        myWasteDataLab_.StorageType = ""
    End If
End Sub

Private Sub chkOther2_Validate(Cancel As Boolean)
    If chkOther2.Value = 1 Then
        myWasteDataLab_.StorageType = "OT" 'Other
        myWasteDataLab_.StorageDesc = Trim(txtOther2.Text)
    End If
End Sub

Private Sub chkPaperBox_Click()
On Error Resume Next
    txtPaperBox.Locked = chkPaperBox.Value = 0
    If chkPaperBox.Value = 1 Then
        txtPaperBox.SetFocus
        chkBag.Value = 0: txtBag.Text = ""
        chkDrum.Value = 0: txtDrumDesc.Text = ""
        chkYard.Value = 0
        chkOther2.Value = 0: txtOther2.Text = ""
        chkTank.Value = 0
        chkPlasticBag.Value = 0: txtPlasticBag.Text = ""
        chkPlasticDrum.Value = 0: txtPlasticDrum.Text = ""
        chkPlasticGallon.Value = 0: txtPlasticGallon.Text = ""
        chkBottle.Value = 0: txtBottle.Text = ""
    End If
    myWasteDataLab_.isPaperBox = chkPaperBox.Value = 1
End Sub

Private Sub chkPaperBox_Validate(Cancel As Boolean)
    myWasteDataLab_.isPaperBox = chkPaperBox.Value = 1
End Sub


Private Sub chkPesticide_Validate(Cancel As Boolean)
    myWasteDataLab_.isPesticide = chkPesticide.Value = 1
End Sub


Private Sub chkPlasticBag_Click()
On Error Resume Next
    txtPlasticBag.Locked = chkPlasticBag.Value = 0
    If chkPlasticBag.Value = 1 Then
        txtPlasticBag.SetFocus
        chkBag.Value = 0: txtBag.Text = ""
        chkDrum.Value = 0: txtDrumDesc.Text = ""
        chkYard.Value = 0
        chkOther2.Value = 0: txtOther2.Text = ""
        chkTank.Value = 0
        chkPlasticDrum.Value = 0: txtPlasticDrum.Text = ""
        chkPlasticGallon.Value = 0: txtPlasticGallon.Text = ""
        chkPaperBox.Value = 0: txtPaperBox.Text = ""
        chkBottle.Value = 0: txtBottle.Text = ""
    End If
    myWasteDataLab_.isPlasticBag = chkPlasticBag.Value = 1
End Sub

Private Sub chkPlasticBag_Validate(Cancel As Boolean)
    If chkPlasticBag.Value = 0 Then txtPlasticBag.Text = ""
    myWasteDataLab_.isPlasticBag = chkPlasticBag.Value = 1
End Sub

Private Sub chkPlasticDrum_Click()
On Error Resume Next
    txtPlasticDrum.Locked = chkPlasticDrum.Value = 0
    If chkPlasticDrum.Value = 1 Then
        txtPlasticDrum.SetFocus
        chkBag.Value = 0: txtBag.Text = ""
        chkDrum.Value = 0: txtDrumDesc.Text = ""
        chkYard.Value = 0
        chkOther2.Value = 0: txtOther2.Text = ""
        chkTank.Value = 0
        chkPlasticBag.Value = 0: txtPlasticBag.Text = ""
        chkPlasticGallon.Value = 0: txtPlasticGallon.Text = ""
        chkPaperBox.Value = 0: txtPaperBox.Text = ""
        chkBottle.Value = 0: txtBottle.Text = ""
    End If
    myWasteDataLab_.isPlasticDrum = chkPlasticDrum.Value = 1
End Sub

Private Sub chkPlasticDrum_Validate(Cancel As Boolean)
    myWasteDataLab_.isPlasticDrum = chkPlasticDrum.Value = 1
End Sub

Private Sub chkPlasticGallon_Click()
On Error Resume Next
    txtPlasticGallon.Locked = chkPlasticGallon.Value = 0
    If chkPlasticGallon.Value = 1 Then
        txtPlasticGallon.SetFocus
        chkBag.Value = 0: txtBag.Text = ""
        chkDrum.Value = 0: txtDrumDesc.Text = ""
        chkYard.Value = 0
        chkOther2.Value = 0: txtOther2.Text = ""
        chkTank.Value = 0
        chkPlasticBag.Value = 0: txtPlasticBag.Text = ""
        chkPlasticDrum.Value = 0: txtPlasticDrum.Text = ""
        chkPaperBox.Value = 0: txtPaperBox.Text = ""
        chkBottle.Value = 0: txtBottle.Text = ""
    Else
    myWasteDataLab_.isPlasticGallon = ""
    myWasteDataLab_.PlasticGallon = ""
    End If
    myWasteDataLab_.isPlasticGallon = chkPlasticGallon.Value = 1
End Sub

Private Sub chkPlasticGallon_Validate(Cancel As Boolean)
    myWasteDataLab_.isPlasticGallon = chkPlasticGallon.Value = 1
End Sub

Private Sub chkRadioActive_Validate(Cancel As Boolean)
    myWasteDataLab_.isRadioActive = chkRadioActive.Value = 1
End Sub

Private Sub chkReact_Validate(Cancel As Boolean)
    myWasteDataLab_.isReact = chkReact.Value = 1
End Sub

Private Sub chkTank_Click()
On Error Resume Next
    If chkTank.Value = 1 Then
        chkBag.Value = 0: txtBag.Text = ""
        chkDrum.Value = 0: txtDrumDesc.Text = ""
        chkYard.Value = 0
        chkOther2.Value = 0: txtOther2.Text = ""
        chkPlasticBag.Value = 0: txtPlasticBag.Text = ""
        chkPlasticDrum.Value = 0: txtPlasticDrum.Text = ""
        chkPlasticGallon.Value = 0: txtPlasticGallon.Text = ""
        chkPaperBox.Value = 0: txtPaperBox.Text = ""
        chkBottle.Value = 0: txtBottle.Text = ""
        myWasteDataLab_.StorageType = "TA" 'Tank
    Else
        myWasteDataLab_.StorageType = ""
    End If
End Sub

Private Sub chkTank_Validate(Cancel As Boolean)
    If chkTank.Value = 1 Then
        myWasteDataLab_.StorageType = "TA" 'Tank
        myWasteDataLab_.StorageDesc = ""
    End If
End Sub

Private Sub chkToxic_Validate(Cancel As Boolean)
    myWasteDataLab_.isToxic = chkToxic.Value = 1
End Sub

Private Sub chkYard_Click()
On Error Resume Next
    If chkYard.Value = 1 Then
        chkBag.Value = 0: txtBag.Text = ""
        chkDrum.Value = 0: txtDrumDesc.Text = ""
        chkOther2.Value = 0: txtOther2.Text = ""
        chkTank.Value = 0
        chkPlasticBag.Value = 0: txtPlasticBag.Text = ""
        chkPlasticDrum.Value = 0: txtPlasticDrum.Text = ""
        chkPlasticGallon.Value = 0: txtPlasticGallon.Text = ""
        chkPaperBox.Value = 0: txtPaperBox.Text = ""
        chkBottle.Value = 0: txtBottle.Text = ""
        myWasteDataLab_.StorageType = "ST" 'Stock yard
    Else
        myWasteDataLab_.StorageType = ""
    End If
End Sub

Private Sub chkYard_Validate(Cancel As Boolean)
    If chkYard.Value = 1 Then
        myWasteDataLab_.StorageType = "ST" 'Stock yard
        myWasteDataLab_.StorageDesc = ""
    End If
End Sub

Private Sub cmdAccepdate_Click()
Load frmWastedataAccepdate
frmWastedataAccepdate.Show 1
End Sub

Private Sub cmdApprove_Click()
    If isEnv = True Then
        If MsgBox("ยืนยันการตรวจรับ?", vbQuestion + vbYesNo) = vbYes Then
            myWasteDataCR_.isChanged = True
            myWasteDataLab_.isChanged = True
            If curCustomer_.CustomerCode = "==NEW==" Then
                MsgBox "ข้อมูลลูกค้ายังไม่ได้ตรวจรับจาก CR ไม่สามารถตรวจรับ Waste ได้", vbExclamation, "Warning"
                Exit Sub
            End If
            isNew = False
            isSave = True
            CmdSave_Click
            If isSave = False Then Exit Sub
            myWasteDataCR_.EnviAccepted
            Label1.Caption = "Envi ตรวจรับแล้วเมื่อ " & FormatYMD_to_DMY(myWasteDataCR_.AcceptedDate, "/", "/")
            Label1.Visible = True
            Label3.Caption = "ตรวจรับแล้ว"
            Label3.Visible = True
            cmdReject.Visible = False
            cmdApprove.Visible = False
            cmdSave.Enabled = True
            txtWasteNo.Text = myWasteDataCR_.WasteNo
            isNew = False
            myWasteDataCR_.isAcceptedByEnvi = True
        End If
    Else
        If myWasteDataCR_.isAcceptedByEnvi = True Then
            MsgBox myWasteDataCR_.WasteName & " ตรวจรับแล้ว ไม่สามารถส่งให้ตรวจรับได้อีก", vbExclamation, "Warning"
            Exit Sub
        End If
        If MsgBox("ยืนยันการส่งข้อมูลกากฯ", vbQuestion + vbYesNo) = vbYes Then
            If myWasteDataCR_.Status = "W" Then
                CmdSave_Click
                If isSave = False Then Exit Sub
            End If
            myWasteDataCR_.Request
            cmdApprove.Visible = False
            Label1.Caption = ""
            Label3.Caption = "รอการตรวจรับ"
            Label3.Visible = True
            Label4.Caption = ""
            Label4.Visible = False
        End If
    End If
End Sub

Private Sub cmdAttach_Click(Index As Integer)
Dim FileName As String
Dim FullPath As String
    If lblAttach(Index).Caption <> "" Then KillFile Index, myWasteDataCR_.CustomerID, myWasteDataCR_.ID
    AttacheFiles Index, FileName, FullPath
    lblAttach(Index).Caption = FileName
    Attaches(Index).FileName = FileName
    Attaches(Index).FullPath = FullPath
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdCls_Click(Index As Integer)
    KillFile Index, myWasteDataCR_.CustomerID, myWasteDataCR_.ID
End Sub

Private Sub cmdOpenFile_Click(Index As Integer)
Dim OpenFile As Double
    If lblAttach(Index).Caption <> "" Then
        ShellExecute 0, vbNullString, Attaches(Index).FullPath, vbNullString, vbNullString, vbNormalFocus
    End If
End Sub

Private Sub cmdReject_Click()
Dim RejectDetail As String
    If MsgBox("ยืนยันการตีกลับ", vbInformation + vbYesNo, "ตีกลับ") = vbNo Then Exit Sub
    RejectDetail = InputBox("ใส่รายละเอียดในการตีกลับ", "ใส่รายละเอียด", "")
    myWasteDataCR_.Reject RejectDetail
    cmdReject.Visible = False
    cmdApprove.Visible = False
    cmdSave.Enabled = False
    cmdDel.Enabled = False
    Label3.Caption = "ตีกลับ"
    Label3.Visible = True
    Label4.Visible = False
    
End Sub

Private Sub cmdReword_Click()
'myWasteDataCR_.WasteYears = "2552"
'myWasteDataCR_.Add
'myWasteDataLab_.Add
'    If SaveAttache(myWasteDataCR_.CustomerID, myWasteDataCR_.ID) = False Then
'            MsgBox "การบันทึกไฟล์แนบไม่สมบูรณ์", vbInformation, "ไฟล์แนบ"
'        End If
'    MsgBox " ปรับข้อมูล Waste สำเร็จ", vbInformation, "การดำเนินการ"
End Sub

Private Sub CmdSave_Click()
    isSave = True
    If Trim(txtCustomer.Text) = "" Then
        MsgBox "กรุณาระบุลูกค้าก่อนทำการจัดเก็บข้อมูล", vbExclamation
        txtCustomer.SetFocus
        isSave = False
        cmdApprove.Visible = False
        Exit Sub
    End If
    
    If Trim(txtMktStaff.Text) = "" Then
        MsgBox "กรุณาระบุชื่อเจ้าหน้าที่การตลาด", vbExclamation
        txtMktStaff.SetFocus
        isSave = False
        cmdApprove.Visible = False
        Exit Sub
    End If
    
    If CboCompany.ListIndex < 0 Then
        MsgBox "กรุณาเลือกบริษัทที่รับกำจัด", vbExclamation, "Warning"
        isSave = False
        cmdApprove.Visible = False
        Exit Sub
    End If
    
    If Trim(txtWasteName.Text) = "" Then
        MsgBox "กรุณาระบุชื่อกากของเสีย", vbExclamation
        txtWasteName.SetFocus
        isSave = False
        cmdApprove.Visible = False
        Exit Sub
    End If
    
    If myWasteDataCR_.isQuotationReq = False Then
        If Trim(txtFactoryNo.Text) = "" Then
            MsgBox "กรุณาระบุเลขทะเบียนโรงงาน", vbExclamation
            txtFactoryNo.SetFocus
            isSave = False
            cmdApprove.Visible = False
            Exit Sub
        End If
    End If
    
    If Trim(txtWasteGenNo.Text) = "" Then
        MsgBox "กรุณาระบุเลขที่ผู้กำเนิดของเสีย", vbExclamation
        txtWasteGenNo.SetFocus
        isSave = False
        cmdApprove.Visible = False
        Exit Sub
    End If
    
'    If Trim(txtIndustryType.Text) = "" Then
'        MsgBox "กรุณาระบุประเภทอุตสาหกรรม", vbExclamation
'        txtIndustryType.SetFocus
'        isSave = False
'        cmdApprove.Visible = False
'        Exit Sub
'    End If
    
    If Trim(txtQty.Text) = "" Or Trim(txtQty.Text) = "0" Then
        MsgBox "กรุณาระบุปริมาณที่จัดเก็บปัจจบัน และมากกว่า 0 ", vbExclamation
        txtQty.SetFocus
        isSave = False
        cmdApprove.Visible = False
        Exit Sub
    End If
    
    If Trim(txtRate.Text) = "" Or Trim(txtRate.Text) = "0" Then
        MsgBox "กรุณาระบุอัตราการเกิดของเสีย และมากกว่า 0", vbExclamation
        txtRate.SetFocus
        isSave = False
        cmdApprove.Visible = False
        Exit Sub
    End If
    
    If Trim(txtGenerateDesc.Text) = "" Then
        MsgBox "กรุณาระบุกระบวนการเกิดของเสีย", vbExclamation
        txtGenerateDesc.SetFocus
        isSave = False
        cmdApprove.Visible = False
        Exit Sub
    End If
    
    If cboWasteChar.ListIndex = 0 Then
        MsgBox "กรุณาระบุคุณสมบัติ", vbExclamation
        cboWasteChar.SetFocus
        isSave = False
        cmdApprove.Visible = False
        Exit Sub
    End If
        
    If cboWasteChar.ListIndex >= 1 And cboWasteChar.ListIndex <= 2 Then
        If cboWasteCharDesc.ListIndex < 0 Then
            MsgBox "กรุณาระบุรายละเอียดคุณสมบัติ", vbExclamation
            cboWasteCharDesc.SetFocus
            isSave = False
            cmdApprove.Visible = False
            Exit Sub
        End If
      End If
      
        If cboStoreType.ListIndex = 0 Then
            MsgBox "กรุณาระบุวิธีการจัดเก็บ", vbExclamation
            cboStoreType.SetFocus
            isSave = False
            cmdApprove.Visible = False
            Exit Sub
        End If
        
        If cboIndoor.ListIndex = 0 Then
            MsgBox "กรุณาระบุสถานที่ในการจัดเก็บ", vbExclamation
            cboIndoor.SetFocus
            isSave = False
            cmdApprove.Visible = False
            Exit Sub
        End If
        
        If Chk10Detail = False Then
            MsgBox "กรุณาระบุข้อ 10.3", vbExclamation
            isSave = False
            cmdApprove.Visible = False
            Exit Sub
        End If
        
     If chkBag.Value = 1 Then
        If Trim(txtBag.Text) = "" Or Trim(txtBag.Text) = 0 Then
            MsgBox "ระบุขนาดจัดเก็บ ของถุง  Big Bag และขนาด มากกว่า 0", vbExclamation
            cmdApprove.Visible = False
            Exit Sub
        End If
    End If
    If chkDrum.Value = 1 Then
        If Trim(txtDrumDesc.Text) = "" Or Trim(txtDrumDesc.Text) = 0 Then
            MsgBox "ระบุขนาดจัดเก็บ Metal Drum และขนาด มากกว่า 0 ", vbExclamation
            cmdApprove.Visible = False
            Exit Sub
        End If
    End If
    
    If chkPlasticGallon.Value = 1 Then
        If Trim(txtPlasticGallon.Text) = "" And Trim(txtPlasticGallon.Text) = 0 Then
            MsgBox "ระบุขนาดจัดเก็บ แกนลอนพลาสติก : Plastic Gallon  และขนาด มากกว่า 0 ", vbExclamation
            cmdApprove.Visible = False
            Exit Sub
        End If
    End If
    
    If chkBottle.Value = 1 Then
        If Trim(txtBottle.Text) = "" Or Trim(txtBottle.Text) = 0 Then
            MsgBox "ระบุขนาดจัดเก็บ ใน ขวดแก้ว : Bottle และระบุขนาดจัดเก็บมากกว่า 0", vbExclamation
            cmdApprove.Visible = False
            Exit Sub
        End If
    End If
    
    If chkPlasticBag.Value = 1 Then
        If Trim(txtPlasticBag.Text) = "" Or Trim(txtPlasticBag.Text) = 0 Then
            MsgBox "ระบุขนาดจัดเก็บ ถุงพลาสติก : Plastic Bag และระบุขนาดจัดเก็บมากกว่า 0 ", vbExclamation
            cmdApprove.Visible = False
            Exit Sub
        End If
    End If
    
    If chkPlasticDrum.Value = 1 Then
        If Trim(txtPlasticDrum.Text) = "" Or Trim(txtPlasticDrum.Text) = 0 Then
            MsgBox "ระบุขนาดจัดเก็บ ถังพลาสติก : Plastic Drum และระบุขนาดจัดเก็บมากกว่า 0", vbExclamation
            cmdApprove.Visible = False
            Exit Sub
        End If
    End If
    
    If chkPaperBox.Value = 1 Then
        If Trim(txtPaperBox.Text) = "" Or Trim(txtPaperBox.Text) = 0 Then
            MsgBox "ระบุขนาดจัดเก็บ ของ กล่องกระดาษ และระบุขนาดจัดเก็บมากกว่า 0", vbExclamation
            cmdApprove.Visible = False
            Exit Sub
        End If
    End If
    
    If chkOther2.Value = 1 Then
        If Trim(txtOther2.Text) = "" Then
            MsgBox "กรุณา ระบุรายละเอียด อื่น ๆ"
            cmdApprove.Visible = False
            Exit Sub
        End If
    End If
        If Option1.Value = True Then
            myWasteDataCR_.isSample = "Y"
        ElseIf Option2.Value = True Then
            myWasteDataCR_.isSample = "N"
        Else
            myWasteDataCR_.isSample = "N"
        End If
    
    If IsDate(Accepdate) = True Then
        myWasteDataCR_.AcceptedDate = Accepdate
    End If
    
    myWasteDataCR_.isNew = isNew
    myWasteDataLab_.isNew = isNew
    If isNew = True Then
        If myWasteDataCR_.Status = "W" Then
            myWasteDataCR_.Status = "W"
        Else
            myWasteDataCR_.Status = "D"  'new waste
        End If
        myWasteDataCR_.isAcceptedByEnvi = False
        myWasteDataCR_.isRequest = False
    End If

    If myWasteDataCR_.isChanged Then
        myWasteDataCR_.Save
        If SaveAttache(myWasteDataCR_.CustomerID, myWasteDataCR_.ID) = False Then
            MsgBox "การบันทึกไฟล์แนบไม่สมบูรณ์", vbInformation, "ไฟล์แนบ"
        End If
        If isNew = False Then
            txtWasteNo.Text = myWasteDataCR_.WasteNo
        Else
            txtWasteNo.Text = "NEW No."
        End If
        If isEnv = False And isNew = True Then
            If txtFactoryNo.Text = "" Or txtIndustryType.Text = "" Or txtQty.Text = "" Or txtRate.Text = "" Or txtGenerateDesc.Text = "" Or cboStoreType.Text = "" Or cboQtyUnit.Text = "" Then
                cmdApprove.Visible = False
            Else
'                cmdApprove.Visible = True
            End If
        End If
    End If
    If myWasteDataLab_.isChanged Then
        myWasteDataLab_.Save
        If isEnv = False And isNew = True Then cmdApprove.Visible = True
    End If
    If txtFactoryNo.Text = "" Or txtIndustryType.Text = "" Or txtQty.Text = "" Or txtRate.Text = "" Or txtGenerateDesc.Text = "" Or cboStoreType.Text = "" Or cboQtyUnit.Text = "" Then
        cmdApprove.Visible = False
    Else
'        cmdApprove.Visible = True
    End If
    If chkBag.Value = 0 And chkDrum.Value = 0 And chkPlasticGallon.Value = 0 And chkBottle.Value = 0 And chkTank.Value = 0 And chkPlasticBag.Value = 0 And chkPlasticDrum.Value = 0 And chkPaperBox.Value = 0 And chkYard.Value = 0 And chkOther2.Value = 0 Then
        cmdApprove.Visible = False
    Else
'        cmdApprove.Visible = True
    End If
    If myWasteDataCR_.isChanged = False And myWasteDataLab_.isChanged = False Then
        MsgBox "จัดเก็บข้อมูลเรียบร้อยแล้ว", vbExclamation
    End If
End Sub
Private Sub ck10_3()
End Sub
Private Sub Command1_Click()
    Load frmDateAssign
    With frmDateAssign
        .SetCallerForm Me
        .Show 1
    End With
End Sub

Private Sub Command3_Click()
On Error Resume Next
Dim tmpS As New BWGSearchManager
Dim tmpCol As Collection
    Set tmpCol = tmpS.WasteDataCRSearch("WasteDataID = '" & myWasteDataCR_.ID & "' Or WasteReferrent='" & myWasteDataCR_.ID & "') And (WasteYears='" & CboWasteYear.Text & "'")
    If tmpCol.Count > 0 Then
            SetDataWasteData tmpCol(1)
            Me.Show 1
    Else
        Set tmpCol = tmpS.WasteDataCRSearch("WasteDataID = '" & myWasteDataCR_.WasteReferrent & "') And (WasteYears='" & CboWasteYear.Text & "'")
        If tmpCol.Count > 0 Then
            SetDataWasteData tmpCol(1)
            Me.Show 1
        Else
            MsgBox "ไม่มีข้อมูล Waste ในปีที่ต้องการหานี้ กรุณาหาข้อมูลปีก่อนหน้านี้ 1 ปี หรือ หลังข้อมูลนี้ 1 ปี", vbInformation, "การค้นหา"
            CboWasteYear.Text = myWasteDataCR_.WasteYears
            Exit Sub
        End If
    End If
    Set tmpCol = Nothing
    Set tmpS = Nothing

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 And Me.ActiveControl.Name <> "grKorOrNo" Then SendKeys "{Tab}"
End Sub
Private Sub PopulateCompany()
Dim x%
    CboCompany.Clear
    For x = 1 To CurrentUser.WasteCompany.Count
        If CurrentUser.WasteCompany(x).CompanyCode = "AKP" Then CboCompany.AddItem CurrentUser.WasteCompany(x).CompanyNameTH
        If CurrentUser.WasteCompany(x).CompanyCode = "BWG" Then CboCompany.AddItem CurrentUser.WasteCompany(x).CompanyNameTH
    Next
    If CboCompany.ListCount > 0 Then
        CboCompany.ListIndex = 0
    End If
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
On Error Resume Next
Dim x%
    isNew = True
    PopulateCompany
    With cboWasteAnalysis
        .AddItem "1.Waste ที่ต้องผ่านการตรวจสอบจากฝ่ายสิ่งแวดล้อม"
        .AddItem "2.Waste ที่ต้องการออกใบเสนอราคาแบบแร่งด่วน"
        .ListIndex = 0
    End With
    With CboWasteYear
        .AddItem "2549"
        .AddItem "2550"
        .AddItem "2551"
        .AddItem "2552"
        .AddItem "2553"
        .AddItem "2554"
        .AddItem "2555"
        .ListIndex = 3
    End With
    With cboWasteType
        .AddItem "ยังไม่ระบุ"
        For x = 1 To WasteTypes.Count
            .AddItem WasteTypes(x).WasteType
        Next
    End With
    With frmPicture
        cmdSave.Picture = .picSave.Picture
        cmdClear.Picture = .picNew.Picture
        cmdReject.Picture = .picDel.Picture
        cmdClose.Picture = .picExit.Picture
        cmdDel.Picture = .picDelete.Picture
    End With
    With grKorOrNo.getGridObj
        .Cols = 3
        .Rows = 20
        .ColWidth(0) = 0
        .ColWidth(1) = 2235
        .ColWidth(2) = 1200
        .TextMatrix(0, 1) = "เลขที่ใบ ก.อ."
        .TextMatrix(0, 2) = "วันที่"
        .FixedAlignment(1) = 1
        .ColAlignment(1) = 1
    End With
    With cboWasteChar
        .Clear
        .AddItem "==ลักษณะของเสีย=="
        .AddItem "ของแข็ง"
        .AddItem "โคลน"
        .AddItem "ของเหลว"
    End With
    With cboLABRes
        .Clear
        .AddItem "ไม่ต้องการ"
        .AddItem "ต้องการ"
    End With
    With cboQtyUnit
        .Clear
        .AddItem "ตัน"
        .AddItem "ลิตร"
    End With
    With cboMaxUnit
        .Clear
        .AddItem "ตัน"
        .AddItem "ลิตร"
    End With
    With cboUnitRate
        .Clear
        .AddItem "ตัน"
        .AddItem "ลิตร"
    End With
    With cboTimeRate
        .Clear
        .AddItem "เดือน"
        .AddItem "ปี"
        .ListIndex = 0
    End With
    With cboSummary
        .Clear
        .AddItem "ยังไม่สรุปผล"
        .AddItem "รับดำเนินการ"
        .AddItem "ไม่รับดำเนินการ"
    End With
    With cboStoreType
        .Clear
        .AddItem ""
        .AddItem "แยกตามประเภทของเสีย"
        .AddItem "รวมกับของเสียชนิดอื่น"
    End With
    With cboIndoor
        .Clear
        .AddItem ""
        .AddItem "กลางแจ้ง"
        .AddItem "ในร่ม"
    End With
    datetime
    cboWasteType.ListIndex = 0
    cboWasteChar.ListIndex = 0
    'cboWasteCharDesc.ListIndex = 0
    cboSummary.ListIndex = 0
    cboStoreType.ListIndex = 0
    cboIndoor.ListIndex = 0
    Set curEUCode1_ = New BWGEUCode1
    Set curEUCode2_ = New BWGEUCode2
    Set curEUCode3_ = New BWGEUCode3
    Set curMktStaff_ = New BWGCompanyStaff
    Set curEnviStaff_ = New BWGCompanyStaff
    PopulateTreatment
    CreateNewWasteData
    SetUserGroup
    ReDim Attaches(4)
    SSTab1.Tab = 0
    SSTab1.TabVisible(4) = False
End Sub
Public Sub datetime()
    Dim today
    today = Now
    Dim myday
    Dim mymounth
    Dim myyear
Dim AcceptedDate As Date
AcceptedDate = myyear & mymounth & myday
End Sub

Private Sub grKorOrNo_KeyDown(KeyCode As Integer, Shift As Integer)
Dim tmpKey As String
    With grKorOrNo.getGridObj
        If KeyCode = 46 And Trim(.TextMatrix(.row, 0)) <> "" Then
            If MsgBox("ท่านต้องการลบเลข กอ. '" & .TextMatrix(.row, 1) & "' หรือไม่?", vbQuestion + vbYesNo) = vbYes Then
                tmpKey = "'" & Trim(.TextMatrix(.row, 0)) & "'"
                If isExist(myWasteDataCR_.KorOrNumbers, tmpKey) Then myWasteDataCR_.KorOrNumbers.Remove tmpKey
                .RemoveItem .row
                myWasteDataCR_.isChanged = True
            End If
        End If
    End With
End Sub

Private Sub grKorOrNo_UpdateValue(row As Integer, col As Integer)
Dim tmpKey As String
Dim tmpDet As BWGKorOrNumber
    With grKorOrNo.getGridObj
        tmpKey = "'" & Trim(.TextMatrix(row, 0)) & "'"
        If isExist(myWasteDataCR_.KorOrNumbers, tmpKey) Then
            Set tmpDet = myWasteDataCR_.KorOrNumbers(tmpKey)
        Else
            Set tmpDet = New BWGKorOrNumber
            tmpDet.ID = GetGUID
            tmpDet.wastedataID = myWasteDataCR_.ID
            myWasteDataCR_.KorOrNumbers.Add tmpDet, "'" & tmpDet.ID & "'"
            .TextMatrix(row, 0) = tmpDet.ID
        End If
        If col = 2 Then
            tmpDet.KorOrNumber = Trim(.TextMatrix(row, 1))
            tmpDet.KorDate = Trim(.TextMatrix(row, 2))
            myWasteDataCR_.isChanged = True
        End If
        Set tmpDet = Nothing
    End With
End Sub

Private Sub lblAttach_Change(Index As Integer)
    myWasteDataCR_.isChanged = True
End Sub

Private Sub lvNewTreatment_ItemCheck(ByVal item As MSComctlLib.ListItem)
    If item.Checked Then
        If isExist(myWasteDataLab_.NewTreatments, item.Key) = False Then
            myWasteDataLab_.NewTreatments.Add colTreatment_(item.Key), item.Key
            myWasteDataLab_.isChanged = True
            If item.Index = 9 Then
                Frame3.Enabled = True
                OptBlend2.Value = True
            Else
                Frame3.Enabled = False
                OptBlend1.Value = False
                OptBlend2.Value = False
            End If
        End If
    Else
        OptBlend1.Value = False
        OptBlend2.Value = False
        Frame3.Enabled = False
        If isExist(myWasteDataLab_.NewTreatments, item.Key) Then
            myWasteDataLab_.NewTreatments.Remove item.Key
            myWasteDataLab_.isChanged = True
        End If
    End If
End Sub
Private Sub lvOldTreatment_ItemCheck(ByVal item As MSComctlLib.ListItem)
    If item.Checked Then
        If isExist(myWasteDataLab_.OldTreatments, item.Key) = False Then
            myWasteDataLab_.OldTreatments.Add colTreatment_(item.Key), item.Key
            myWasteDataLab_.isChanged = True
        End If
    Else
        If isExist(myWasteDataLab_.OldTreatments, item.Key) Then
            myWasteDataLab_.OldTreatments.Remove item.Key
            myWasteDataLab_.isChanged = True
        End If
    End If
End Sub
Private Sub OptAnalysType1_Click()
    On Error Resume Next
    myWasteDataLab_.AnalysisType = 1
    myWasteDataLab_.AnalysisDesc = ""
    txtAnalysDesc.Locked = True
    txtAnalysDesc.Text = ""
    myWasteDataLab_.isChanged = True
End Sub
Private Sub OptAnalysType2_Click()
    On Error Resume Next
    myWasteDataLab_.AnalysisType = 2
    myWasteDataLab_.AnalysisDesc = ""
    txtAnalysDesc.Locked = True
    txtAnalysDesc.Text = ""
    myWasteDataLab_.isChanged = True
End Sub
Private Sub OptAnalysType3_Click()
    On Error Resume Next
    myWasteDataLab_.AnalysisType = 3
    txtAnalysDesc.Locked = False
    txtAnalysDesc.SetFocus
    myWasteDataLab_.AnalysisDesc = Trim(txtAnalysDesc)
    myWasteDataLab_.isChanged = True
End Sub
Private Sub OptBlend0_Click()
    myWasteDataLab_.isBlending = ""
    myWasteDataLab_.isChanged = True
End Sub
Private Sub OptBlend1_Click()
    myWasteDataLab_.isBlending = "Y"
    myWasteDataLab_.isChanged = True
End Sub
Private Sub OptBlend2_Click()
    myWasteDataLab_.isBlending = "N"
    myWasteDataLab_.isChanged = True
End Sub
Private Sub Option1_Click()
    If Option1.Value = True Then
        myWasteDataCR_.isSample = "Y"
    ElseIf Option2.Value = True Then
        myWasteDataCR_.isSample = "N"
    Else
        myWasteDataCR_.isSample = "N"
    End If
    myWasteDataCR_.isChanged = True
End Sub
Private Sub Option2_Click()
    If Option1.Value = True Then
        myWasteDataCR_.isSample = "Y"
    ElseIf Option2.Value = True Then
        myWasteDataCR_.isSample = "N"
    Else
        myWasteDataCR_.isSample = "N"
    End If
    myWasteDataCR_.isChanged = True
End Sub

Private Sub txtAnalysDesc_Validate(Cancel As Boolean)
    myWasteDataLab_.AnalysisDesc = Trim(txtAnalysDesc)
    myWasteDataLab_.isChanged = True
End Sub

Private Sub txtBag_GotFocus()
    txtBag.Locked = chkBag.Value = 0
End Sub
Private Sub txtBag_Validate(Cancel As Boolean)
    If chkBag.Value = 1 Then myWasteDataLab_.StorageDesc = Trim(txtBag.Text)
    myWasteDataLab_.isChanged = True
End Sub

Private Sub txtBottle_Validate(Cancel As Boolean)
    myWasteDataLab_.Bottle = Trim(txtBottle)
    myWasteDataLab_.isChanged = True
End Sub

Private Sub txtCancerDesc_KeyPress(KeyAscii As Integer)
    myWasteDataLab_.CarcinoGenicDesc = Trim(txtCancerDesc.Text)
    myWasteDataLab_.isChanged = True
End Sub

Private Sub txtCancerDesc_Validate(Cancel As Boolean)
    myWasteDataLab_.CarcinoGenicDesc = Trim(txtCancerDesc.Text)
    myWasteDataLab_.isChanged = True
End Sub

Private Sub txtCharDesc_KeyPress(KeyAscii As Integer)
    myWasteDataLab_.CharDescOther = Trim(txtCharDesc)
    myWasteDataLab_.isChanged = True
End Sub

Private Sub txtCharDesc_Validate(Cancel As Boolean)
    myWasteDataLab_.CharDescOther = Trim(txtCharDesc)
    myWasteDataLab_.isChanged = True
End Sub

Private Sub txtColor_Validate(Cancel As Boolean)
    myWasteDataLab_.Color = Trim(txtColor)
    myWasteDataLab_.isChanged = True
End Sub

Private Sub txtContractNo_Validate(Cancel As Boolean)
    If Trim(txtContractNo.Text) = "" And Trim(myWasteDataCR_.Contract.DocNo) <> Trim(txtContractNo.Text) Then
        Set myWasteDataCR_.Contract = New BWGContract
    Else
        If Trim(myWasteDataCR_.Contract.DocNo) <> Trim(txtContractNo.Text) Then
            Set myWasteDataCR_.Contract = New BWGContract
            Load frmContractSearch
            With frmContractSearch
                .getCriteria txtContractNo.Text
                .SetCallerForm Me
                .Show 1
            End With
            If myWasteDataCR_.Contract.DocNo = "" Then txtContractNo.Text = ""
        End If
    End If
End Sub

Private Sub txtDrumDesc_GotFocus()
    txtDrumDesc.Locked = chkDrum.Value = 0
End Sub

Private Sub txtDrumDesc_Validate(Cancel As Boolean)
    If chkDrum.Value = 1 Then myWasteDataLab_.StorageDesc = Trim(txtDrumDesc.Text)
    myWasteDataLab_.isChanged = True
End Sub

Private Sub txtEnviStaff_GotFocus()
    callStaffType_ = "ENV"
End Sub

Private Sub txtEnviStaff_Validate(Cancel As Boolean)
    If curEnviStaff_ Is Nothing Then Set curEnviStaff_ = New BWGCompanyStaff
    If Trim(txtEnviStaff.Text) = "" And Trim(curEnviStaff_.StaffFName & " " & curEnviStaff_.StaffLName) <> Trim(txtEnviStaff.Text) Then
        Set curEnviStaff_ = Nothing
        Set curEnviStaff_ = New BWGCompanyStaff
        myWasteDataCR_.EnviStaffID = ""
        myWasteDataCR_.EnviStaffName = ""
    Else
        If Trim(curEnviStaff_.StaffFName & " " & curEnviStaff_.StaffLName) <> Trim(txtEnviStaff.Text) Then
            Set curEnviStaff_ = Nothing
            Set curEnviStaff_ = New BWGCompanyStaff
            myWasteDataCR_.EnviStaffID = ""
            myWasteDataCR_.EnviStaffName = ""
            Load frmStaffSearch
            With frmStaffSearch
                .getCriteria txtEnviStaff.Text
                .SetCallerForm Me
                .Show 1
            End With
            If myWasteDataCR_.EnviStaffID = "" Then
                txtEnviStaff.Text = ""
            End If
        End If
    End If
End Sub

Private Sub txtEUCode1_Validate(Cancel As Boolean)
    If Trim(txtEUCode1.Text) = "" And Trim(curEUCode1_.Code1) <> Trim(txtEUCode1.Text) Then
        Set curEUCode1_ = Nothing
        Set curEUCode1_ = New BWGEUCode1
        With myWasteDataCR_
            .EUCode1ID = ""
            .EUCode1 = ""
        End With
    Else
        If Trim(curEUCode1_.Code1) <> Trim(txtEUCode1.Text) Then
            Set curEUCode1_ = Nothing
            Set curEUCode1_ = New BWGEUCode1
            With myWasteDataCR_
                .EUCode1ID = ""
                .EUCode1 = ""
            End With
            Load frmEUCode1Search
            With frmEUCode1Search
                .getCriteria txtEUCode1.Text
                .SetCallerForm Me
                .Show 1
            End With
            If myWasteDataCR_.EUCode1ID = "" Then txtEUCode1.Text = ""
        End If
    End If
End Sub

Private Sub txtEUCode2_Validate(Cancel As Boolean)
    If Trim(txtEUCode2.Text) = "" And Trim(curEUCode2_.Code2) <> Trim(txtEUCode2.Text) Then
        Set curEUCode2_ = Nothing
        Set curEUCode2_ = New BWGEUCode2
        With myWasteDataCR_
            .EUCode2 = ""
            .EUCode2ID = ""
        End With
    Else
        If Trim(curEUCode2_.Code2) <> Trim(txtEUCode2.Text) Then
            Set curEUCode2_ = Nothing
            Set curEUCode2_ = New BWGEUCode2
            With myWasteDataCR_
                .EUCode2 = ""
                .EUCode2ID = ""
            End With
            Load frmEUCode2Search
            With frmEUCode2Search
                .getCriteria txtEUCode2.Text
                .SetCallerForm Me
                .Show 1
            End With
            If myWasteDataCR_.EUCode2ID = "" Then txtEUCode2.Text = ""
        End If
    End If
End Sub

Private Sub txtEUCode3_Validate(Cancel As Boolean)
    If Trim(txtEUCode3.Text) = "" And Trim(curEUCode3_.Code3) <> Trim(txtEUCode3.Text) Then
        Set curEUCode3_ = Nothing
        Set curEUCode3_ = New BWGEUCode3
        With myWasteDataCR_
            .EUCode3 = ""
            .EUCode3ID = ""
        End With
    Else
        If Trim(curEUCode3_.Code3) <> Trim(txtEUCode3.Text) Then
            Set curEUCode3_ = Nothing
            Set curEUCode3_ = New BWGEUCode3
            With myWasteDataCR_
                .EUCode3 = ""
                .EUCode3ID = ""
            End With
            Load frmEUCode3Search
            With frmEUCode3Search
                .getCriteria txtEUCode3.Text
                .SetCallerForm Me
                .Show 1
            End With
            If myWasteDataCR_.EUCode3ID = "" Then txtEUCode3.Text = ""
        End If
    End If
End Sub

Public Sub SetDataEUCode1(selEU1 As BWGEUCode1, Optional isAssigned As Boolean = True)
    Set curEUCode1_ = selEU1
    With curEUCode1_
        txtEUCode1.Text = .Code1
        If isAssigned Then
            myWasteDataCR_.EUCode1 = .Code1
            myWasteDataCR_.EUCode1ID = .ID
        End If
    End With
End Sub

Public Sub SetDataEUCode2(selEU2 As BWGEUCode2, Optional isAssigned As Boolean = True)
    Set curEUCode2_ = selEU2
    With curEUCode2_
        txtEUCode2.Text = .Code2
        If isAssigned Then
            myWasteDataCR_.EUCode2 = .Code2
            myWasteDataCR_.EUCode2ID = .ID
        End If
    End With
End Sub

Public Sub SetDataEUCode3(selEU3 As BWGEUCode3, Optional isAssigned As Boolean = True)
    Set curEUCode3_ = selEU3
    With curEUCode3_
        txtEUCode3.Text = .Code3
        txtHAType.Text = .HAType
        If isAssigned Then
            myWasteDataCR_.EUCode3 = .Code3
            myWasteDataCR_.EUCode3ID = .ID
        End If
    End With
End Sub

Private Sub CreateNewWasteData()
    Set myWasteDataCR_ = Nothing
    Set myWasteDataCR_ = New BWGWasteDataCR
    Set myWasteDataLab_ = Nothing
    Set myWasteDataLab_ = New BWGWasteDataLAB
    
    With myWasteDataCR_
        .ID = GetGUID
        If Not curCustomer_ Is Nothing Then
            .CustomerID = curCustomer_.ID
            .CustomerName = curCustomer_.CustomerName
        End If
        If Not curMktStaff_ Is Nothing Then
            .MarketingStaffID = curMktStaff_.ID
            .MarketingStaffName = curMktStaff_.StaffFName & " " & curMktStaff_.StaffLName
        End If
        If Not curEnviStaff_ Is Nothing Then
            .EnviStaffID = curEnviStaff_.ID
            .EnviStaffName = curEnviStaff_.StaffFName & " " & curEnviStaff_.StaffLName
        End If
        .isChanged = False
    End With
    With myWasteDataLab_
        .ID = GetGUID
        .wastedataID = myWasteDataCR_.ID
        .isChanged = False
    End With
    
End Sub

Public Function getEUCode1() As BWGEUCode1
    If curEUCode1_ Is Nothing Then Set curEUCode1_ = New BWGEUCode1
    Set getEUCode1 = curEUCode1_
End Function

Public Function getEUCode2() As BWGEUCode2
    If curEUCode2_ Is Nothing Then Set curEUCode2_ = New BWGEUCode2
    Set getEUCode2 = curEUCode2_
End Function

Public Function getEUCode3() As BWGEUCode3
    If curEUCode3_ Is Nothing Then Set curEUCode3_ = New BWGEUCode3
    Set getEUCode3 = curEUCode3_
End Function

Private Sub txtExpDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 Then
        If MsgBox("ท่านต้องการลบวันที่ใบอนุญาตหมดอายุหรือไม่?", vbQuestion + vbYesNo) = vbYes Then
            myWasteDataCR_.LicenseExpDate = ""
            txtExpDate.Text = ""
        End If
    End If
End Sub

Private Sub txtFactoryNo_Validate(Cancel As Boolean)
    myWasteDataCR_.FactoryRegisNo = Trim(txtFactoryNo.Text)
End Sub

Private Sub txtFlashPoint_Validate(Cancel As Boolean)
    If Trim(txtFlashPoint.Text) <> "" Then
        myWasteDataLab_.FlashPoint = Trim(txtFlashPoint.Text)
    Else
        myWasteDataLab_.FlashPoint = 0
    End If
    myWasteDataLab_.isChanged = True
End Sub

Private Sub txtGenerateDesc_Change()
    myWasteDataLab_.GenerateDetail = Trim(txtGenerateDesc)
    myWasteDataLab_.isChanged = True
End Sub

Private Sub txtGenerateDesc_KeyPress(KeyAscii As Integer)
    myWasteDataLab_.GenerateDetail = Trim(txtGenerateDesc)
    myWasteDataLab_.isChanged = True
End Sub

Private Sub txtGenerateDesc_Validate(Cancel As Boolean)
    myWasteDataLab_.GenerateDetail = Trim(txtGenerateDesc)
    myWasteDataLab_.isChanged = True
End Sub

Private Sub txtIndustryType_Validate(Cancel As Boolean)
'    If curIndustryType_ Is Nothing Then Set curIndustryType_ = New BWGIndustryType
'    If Trim(txtIndustryType.Text) = "" And Trim(curIndustryType_.TypeName) <> Trim(txtIndustryType.Text) Then
'        Set curIndustryType_ = Nothing
'        Set curIndustryType_ = New BWGIndustryType
'        With myWasteDataCR_
'            .IndustryTypeID = ""
'            .IndustryTypeName = ""
'        End With
'    Else
'        If Trim(curIndustryType_.TypeName) <> Trim(txtIndustryType.Text) Then
'            Set curIndustryType_ = Nothing
'            Set curIndustryType_ = New BWGIndustryType
'            With myWasteDataCR_
'                .IndustryTypeID = ""
'                .IndustryTypeName = ""
'            End With
'            Load frmIndustryType
'            With frmIndustryType
'                .SetCallerForm Me
'                .Show 1
'            End With
'            If myWasteDataCR_.IndustryTypeID = "" Then txtIndustryType.Text = ""
'        End If
'    End If
End Sub

Private Sub txtKor_Validate(Cancel As Boolean)
    myWasteDataCR_.Kor = txtKor.Text
End Sub

Private Sub txtMaxQty_Validate(Cancel As Boolean)
    If IsNumeric(Trim(txtMaxQty.Text)) Then
        myWasteDataCR_.MaxQty = CDbl(Trim(txtMaxQty.Text))
    Else
        myWasteDataCR_.MaxQty = 0
    End If
End Sub

Private Sub txtMktStaff_GotFocus()
    callStaffType_ = "MKT"
End Sub

Private Sub txtMktStaff_Validate(Cancel As Boolean)
    If curMktStaff_ Is Nothing Then Set curMktStaff_ = New BWGCompanyStaff
    If Trim(txtMktStaff.Text) = "" And Trim(curMktStaff_.StaffFName & " " & curMktStaff_.StaffLName) <> Trim(txtMktStaff.Text) Then
        Set curMktStaff_ = Nothing
        Set curMktStaff_ = New BWGCompanyStaff
        myWasteDataCR_.MarketingStaffID = ""
        myWasteDataCR_.MarketingStaffName = ""
    Else
        If Trim(curMktStaff_.StaffFName & " " & curMktStaff_.StaffLName) <> Trim(txtMktStaff.Text) Then
            Set curMktStaff_ = Nothing
            Set curMktStaff_ = New BWGCompanyStaff
            myWasteDataCR_.MarketingStaffID = ""
            myWasteDataCR_.MarketingStaffName = ""
            Load frmStaffSearch
            With frmStaffSearch
                .getCriteria txtMktStaff.Text
                .SetCallerForm Me
                .Show 1
            End With
            If myWasteDataCR_.MarketingStaffID = "" Then
                txtMktStaff.Text = ""
            End If
        End If
    End If
End Sub

Public Sub SetDataStaff(selStaff As BWGCompanyStaff, Optional isAssigned As Boolean = True)
    If callStaffType_ = "MKT" Then
        Set curMktStaff_ = selStaff
        With curMktStaff_
            txtMktStaff.Text = .StaffFName & " " & .StaffLName
            If isAssigned Then
                myWasteDataCR_.MarketingStaffID = .ID
                myWasteDataCR_.MarketingStaffName = .StaffFName & " " & .StaffLName
            End If
        End With
    ElseIf callStaffType_ = "ENV" Then
        Set curEnviStaff_ = selStaff
        With curEnviStaff_
            txtEnviStaff.Text = .StaffFName & " " & .StaffLName
            If isAssigned Then
                myWasteDataCR_.EnviStaffID = .ID
                myWasteDataCR_.EnviStaffName = .StaffFName & " " & .StaffLName
            End If
        End With
    End If
End Sub

Private Sub txtMoud_Validate(Cancel As Boolean)
    myWasteDataCR_.Moud = txtMoud.Text
End Sub

Private Sub txtOdour_Validate(Cancel As Boolean)
    myWasteDataLab_.Odour = Trim(txtOdour)
    myWasteDataLab_.isChanged = True
End Sub

Private Sub txtOldStorePlace_Validate(Cancel As Boolean)
    myWasteDataLab_.OldEliminator = Trim(txtOldStorePlace.Text)
End Sub

Private Sub txtOther_Change()
    If Trim(txtOther.Text) <> "" Then
        chkOther.Value = 1
    Else
        chkOther.Value = 0
    End If
    myWasteDataLab_.isChanged = True
End Sub

Private Sub txtOther_Validate(Cancel As Boolean)
    myWasteDataLab_.OtherDesc = Trim(txtOther.Text)
    myWasteDataLab_.isChanged = True
End Sub

Private Sub txtOther2_GotFocus()
    txtOther2.Locked = chkOther2.Value = 0
End Sub

Private Sub txtOther2_Validate(Cancel As Boolean)
    If chkOther2.Value = 1 Then myWasteDataLab_.StorageDesc = Trim(txtOther2.Text)
    myWasteDataLab_.isChanged = True
End Sub

Private Sub txtPaperBox_Validate(Cancel As Boolean)
    myWasteDataLab_.PaperBox = Trim(txtPaperBox)
    myWasteDataLab_.isChanged = True
End Sub

Private Sub txtPesticide_Validate(Cancel As Boolean)
    myWasteDataLab_.PesticideDesc = Trim(txtPesticide.Text)
    myWasteDataLab_.isChanged = True
End Sub

Private Sub txtpH_Validate(Cancel As Boolean)
    If Trim(txtpH.Text) <> "" Then
        myWasteDataLab_.pHValue = Trim(txtpH.Text)
    Else
        myWasteDataLab_.pHValue = 0
    End If
    myWasteDataLab_.isChanged = True
End Sub

Private Sub txtPlasticBag_Validate(Cancel As Boolean)
    myWasteDataLab_.PlasticBag = Trim(txtPlasticBag)
    myWasteDataLab_.isChanged = True
End Sub


Private Sub txtPlasticDrum_Validate(Cancel As Boolean)
    myWasteDataLab_.PlasticDrum = Trim(txtPlasticDrum)
    myWasteDataLab_.isChanged = True
End Sub

Private Sub txtPlasticGallon_Validate(Cancel As Boolean)
    myWasteDataLab_.PlasticGallon = Trim(txtPlasticGallon)
    myWasteDataLab_.isChanged = True
End Sub

Private Sub txtQty_Validate(Cancel As Boolean)
    If IsNumeric(Trim(txtQty.Text)) Then
        myWasteDataCR_.NormalQty = CDbl(Trim(txtQty.Text))
    Else
        myWasteDataCR_.NormalQty = 0
    End If
End Sub

Private Sub txtRadioActive_Validate(Cancel As Boolean)
    myWasteDataLab_.RadioActiveDesc = Trim(txtRadioActive.Text)
    myWasteDataLab_.isChanged = True
End Sub

Private Sub txtRate_Validate(Cancel As Boolean)
    If IsNumeric(Trim(txtRate.Text)) Then
        myWasteDataCR_.WasteRate = CDbl(Trim(txtRate.Text))
    Else
        myWasteDataCR_.WasteRate = 0
    End If
End Sub

Private Sub txtReactDesc_Validate(Cancel As Boolean)
    myWasteDataLab_.ReActWith = Trim(txtReactDesc.Text)
    myWasteDataLab_.isChanged = True
End Sub
Private Sub TxtRefME_Validate(Cancel As Boolean)
    myWasteDataCR_.RefME = Trim(TxtRefME.Text)
End Sub

Private Sub txtToxicDesc_Validate(Cancel As Boolean)
    myWasteDataLab_.ToxicDesc = Trim(txtToxicDesc.Text)
    myWasteDataLab_.isChanged = True
End Sub

Private Sub txtWasteGenNo_Validate(Cancel As Boolean)
    myWasteDataCR_.WasteCreateNo = Trim(txtWasteGenNo.Text)
End Sub

Private Sub txtWasteName_Validate(Cancel As Boolean)
    myWasteDataCR_.WasteName = Trim(txtWasteName.Text)
End Sub

Private Sub PopulateTreatment()
Dim tmpS As New BWGSearchManager
Dim x%
    Set colTreatment_ = tmpS.TreatmentDataSearch
    lvOldTreatment.ListItems.Clear
    lvNewTreatment.ListItems.Clear
    For x = 1 To colTreatment_.Count
        lvOldTreatment.ListItems.Add , "'" & colTreatment_(x).ID & "'", colTreatment_(x).treatmentDesc
        lvNewTreatment.ListItems.Add , "'" & colTreatment_(x).ID & "'", colTreatment_(x).treatmentDesc
    Next
End Sub

Public Sub SetCallerForm(frm As Form, selCust As BWGCustomer)
    Set callerForm_ = frm
    Set curCustomer_ = selCust
    With curCustomer_
        txtCustomer.Text = .CustomerCode & " - " & .CustomerName
        TxtMECode.Text = .Refcode
        myWasteDataCR_.CustomerID = .ID
        myWasteDataCR_.CustomerName = .CustomerName
    End With
    '******
    If myWasteDataCR_.FactoryRegisNo = "" Then
        If txtFactoryNo.Text = "" Then
             With curCustomer_
                 txtFactoryNo.Text = .RegisterNo
             End With
        End If
   End If
End Sub

Private Sub cmdClear_Click()
Dim selResult As Integer
    
    If myWasteDataCR_.isChanged Or myWasteDataLab_.isChanged Then
        If cmdSave.Enabled = True Then
            selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนสร้างใหม่หรือไม่?", vbQuestion + vbYesNoCancel)
            If selResult = vbYes Then
                CmdSave_Click
                If myWasteDataCR_.isChanged Or myWasteDataLab_.isChanged Then Exit Sub
            ElseIf selResult = vbCancel Then
                Exit Sub
            End If
        End If
    End If
    ClearScreen
    CreateNewWasteData
    isNew = True
    txtEUCode1.SetFocus
End Sub

Private Sub cmdDel_Click()
Dim tmpStr As String
Dim tmpGUID As String
Dim tmpS As New BWGSearchManager, tmpCol As Collection
    
    Set tmpCol = tmpS.QuotationDetailSearch("WasteDataID='" & myWasteDataCR_.ID & "'")
    If tmpCol.Count > 0 Then
        MsgBox "ไม่สามารถลบข้อมูลกากของเสียนี้ได้เพราะได้นำไปออกใบเสนอราคาแล้ว", vbExclamation
        Set tmpCol = Nothing
        Set tmpS = Nothing
        Exit Sub
    End If
    Set tmpCol = Nothing
    Set tmpS = Nothing

    If MsgBox("ต้องการลบข้อมูลกากของเสียนี้หรือไม่?", vbQuestion + vbYesNo) = vbYes Then
        tmpGUID = GetGUID
        If isNew = False Then
            MsgBox "ข้อมูลนี้เป็นข้อมูลเก่าไม่สามารถลบได้", vbInformation, "Warning"
            Exit Sub
        End If
        If isNew = True Then
            tmpStr = myWasteDataCR_.DeleteNew(WasteID)
            ClearScreen
            CreateNewWasteData
            txtEUCode1.SetFocus
        End If
    End If
    isNew = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim selResult As Integer
    If myWasteDataCR_ Is Nothing Or myWasteDataLab_ Is Nothing Then Exit Sub
    If myWasteDataCR_.isChanged Or myWasteDataLab_.isChanged Then
        If cmdSave.Visible = False Then Exit Sub
        If cmdSave.Enabled = True Then
            selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนปิดหน้าจอหรือไม่?", vbQuestion + vbYesNoCancel)
            If selResult = vbYes Then
                CmdSave_Click
                If myWasteDataCR_.isChanged Or myWasteDataLab_.isChanged Then Cancel = 1
            ElseIf selResult = vbCancel Then
                Cancel = 1
            End If
        End If
    End If
End Sub

Private Sub ClearScreen()
Dim x%
    cmdApprove.Visible = False
    cmdSave.Enabled = True
    cmdDel.Enabled = True
    Label1.Visible = False
    Label3.Visible = False
    txtWasteNo.Text = "== AUTO =="
    txtEUCode1.Text = ""
    txtEUCode2.Text = ""
    txtEUCode3.Text = ""
    txtHAType.Text = ""
    cboWasteType.ListIndex = 0
    txtMoud.Text = ""
    txtKor.Text = ""
    cboLABRes.ListIndex = -1
    txtWasteName.Text = ""
    txtQty.Text = ""
    cboQtyUnit.ListIndex = -1
    txtMaxQty.Text = ""
    cboMaxUnit.ListIndex = -1
    txtRate.Text = ""
    cboUnitRate.ListIndex = -1
    cboTimeRate.ListIndex = -1
    txtGenerateDesc.Text = ""
    cboWasteChar.ListIndex = 0
    cboWasteCharDesc.ListIndex = -1
    chkExplode.Value = 0
    chkFlame.Value = 0
    txtFlashPoint.Text = ""
    chkCorrosive.Value = 0
    txtpH.Text = ""
    chkReact.Value = 0
    txtReactDesc.Text = ""
    chkCancer.Value = 0
    txtCancerDesc.Text = ""
    chkToxic.Value = 0
    cboToxic.ListIndex = -1
    txtToxicDesc.Text = ""
    chkRadioActive.Value = 0
    txtRadioActive.Text = ""
    chkPesticide.Value = 0
    txtPesticide.Text = ""
    chkOther.Value = 0
    txtOther.Text = ""
    cboSummary.ListIndex = 0
    cboStoreType.ListIndex = 0
    cboIndoor.ListIndex = 0
    chkBag.Value = 0
    txtBag.Text = ""
    chkYard.Value = 0
    chkDrum.Value = 0
    txtDrumDesc.Text = ""
    chkOther2.Value = 0
    txtOther2.Text = ""
    chkTank.Value = 0
    txtOldStorePlace.Text = ""
    txtFactoryNo.Text = ""
    txtWasteGenNo.Text = ""
    txtExpDate.Text = ""
    txtIndustryType.Text = ""
    cboTreatmentType.ListIndex = -1
    Set curIndustryType_ = Nothing
    Set curIndustryType_ = New BWGIndustryType
    grKorOrNo.ClearAllData
    grKorOrNo.getGridObj.Rows = 20
    For x = 1 To lvOldTreatment.ListItems.Count
        lvOldTreatment.ListItems(x).Checked = False
    Next
    Set curEUCode1_ = Nothing
    Set curEUCode2_ = Nothing
    Set curEUCode3_ = Nothing
    Set curEUCode1_ = New BWGEUCode1
    Set curEUCode2_ = New BWGEUCode2
    Set curEUCode3_ = New BWGEUCode3
    
    '***** New Variable
    cboWasteAnalysis.ListIndex = 0
End Sub

Public Sub SetDataWasteData(selWaste As BWGWasteDataCR)
Dim tmpS As New BWGSearchManager
Dim tmpCol As Collection
Dim x%
    cmdAccepdate.Visible = False
    Set myWasteDataCR_ = selWaste
    With myWasteDataCR_
        WasteID = .ID
        If isEnv = True Then
            If curCustomer_.isNew = True Then
                cmdApprove.Visible = False
                cmdReject.Visible = False
                cmdSave.Enabled = False
                cmdDel.Enabled = False
            Else
                If .Status <> "R" Then
                    cmdApprove.Visible = False
                    cmdReject.Visible = False
                    If .Status = "J" Then
                        cmdSave.Enabled = False
                        cmdDel.Enabled = False
                        cmdAccepdate.Visible = False
                        isNew = True
                    ElseIf .Status = "W" Then
                        cmdSave.Enabled = False
                        cmdDel.Enabled = False
                        cmdAccepdate.Visible = False
                        isNew = True
                    Else
                        cmdSave.Enabled = True
                        cmdAccepdate.Visible = True
                        isNew = False
                    End If
                ElseIf .Status = "R" Then
'                    cmdApprove.Visible = True
'                    cmdReject.Visible = True
                    cmdSave.Enabled = False
                     cmdDel.Enabled = False
                     cmdAccepdate.Visible = False
                    isNew = True
                End If
            End If
            cboWasteAnalysis.Enabled = False
        Else
            cmdAccepdate.Visible = False
            If .isQuotationReq = False Then
                If .Status = "D" Or .Status = "R" Then
'                    cmdApprove.Visible = True
                    cmdSave.Enabled = True
                    cmdDel.Enabled = True
                Else
                    cmdApprove.Visible = False
                    cmdSave.Enabled = False
                    cmdDel.Enabled = False
                End If
            Else
                cmdApprove.Visible = False
                cmdSave.Enabled = True
                cmdDel.Enabled = True
            End If
            isNew = .isNew
        End If
        Select Case .Status
        Case "P", "A", "E", "IA"
            If .AcceptedDate = "" Then
                cmdAccepdate.Visible = True
            End If
            Label1.Caption = "Envi ตรวจรับแล้วเมื่อ " & FormatYMD_to_DMY(.AcceptedDate, "/", "/")
            Label1.Visible = True
            Label3.Caption = "ตรวจรับแล้ว"
            Label3.Visible = True
            If isEnv = True Then
                cmdSave.Enabled = True
                cmdDel.Enabled = True
            Else
                cmdSave.Enabled = False
                cmdDel.Enabled = False
            End If
        Case "D"
            Label1.Caption = ""
            Label1.Visible = True
            If .isQuotationReq = False Then
                Label3.Caption = "ยังไม่ส่งข้อมูล"
            Else
                Label3.Caption = "ขอออกใบเสนอราคา"
            End If
            Label3.Visible = True
            Label4.Caption = ""
            Label4.Visible = False
        Case "J"
            Label1.Caption = "Envi ตีกลับเมื่อ " & FormatYMD_to_DMY(.AcceptedDate, "/", "/")
            Label1.Visible = True
            Label3.Caption = "ตีกลับ"
            Label3.Visible = True
            Label4.Caption = .RejectDesc
            Label4.Visible = True
            cmdSave.Enabled = True
            cmdDel.Enabled = True
        Case "R"
            If isEnv = True Then
                Label3.Caption = "รอการตรวจรับ"
                Label3.Visible = True
                Label4.Caption = "ส่งให้ Env เมื่อ" & Format(.RequestDate, "dd/mm/yyyy")
                Label4.Visible = True
            Else
                Label3.Caption = "ส่งให้ Env เมื่อ"
                Label4.Caption = Format(.RequestDate, "dd/mm/yyyy")
                Label3.Visible = True
                Label4.Visible = True
            End If
        Case "W"
            Label3.Caption = "รออนุมัติใบเสนอราคา"
            Label4.Caption = "ให้ Sales เก็บข้อมูลเพิ่ม"
            Label3.Visible = True
            Label4.Visible = True
            If isEnv = False Then
                'cmdApprove.Visible = True
            End If
        End Select
        
        WasteReferent = .ID
        CboWasteYear.Text = .WasteYears
        txtWasteNo.Text = .WasteNo
        txtContractNo.Text = .Contract.DocNo
        If isExist(CurrentUser.WorkCompanies, "'" & .CompanyID & "'") Then
            SetTextToCombo CurrentUser.WorkCompanies("'" & .CompanyID & "'").CompanyNameTH, CboCompany
            If .CompanyID = "" Then CboCompany.ListIndex = 0
        End If
        Set tmpCol = tmpS.EUCode1Search("EUCode1ID='" & .EUCode1ID & "'")
        If tmpCol.Count > 0 Then SetDataEUCode1 tmpCol(1), False
        Set tmpCol = tmpS.EUCode2Search("EUCode2ID='" & .EUCode2ID & "'")
        If tmpCol.Count > 0 Then SetDataEUCode2 tmpCol(1), False
        Set tmpCol = tmpS.EUCode3Search("EUCode3ID='" & .EUCode3ID & "'")
        If tmpCol.Count > 0 Then SetDataEUCode3 tmpCol(1), False
             
        SetTextToCombo .WasteType, cboWasteType
        If cboWasteType.ListIndex = -1 Then cboWasteType.ListIndex = 0
        
        If .isLabResultNeeded Then
            cboLABRes.ListIndex = 1
        Else
            cboLABRes.ListIndex = 0
        End If
        
        If .isSample = "Y" Then
            Option1.Value = True
        Else
            Option2.Value = True
        End If
        
        txtMoud.Text = .Moud
        txtKor.Text = .Kor
        txtWasteName.Text = .WasteName
        txtQty.Text = .NormalQty
        txtMaxQty.Text = .MaxQty
        txtRate.Text = .WasteRate
        txtFactoryNo.Text = .FactoryRegisNo
        txtWasteGenNo.Text = .WasteCreateNo
        txtExpDate.Text = FormatYMD_to_DMY(.LicenseExpDate, "/", "/")
        If .NormalUnit = "T" Then
            cboQtyUnit.ListIndex = 0
        ElseIf .NormalUnit = "L" Then
            cboQtyUnit.ListIndex = 1
        End If
        If .MaxUnit = "T" Then
            cboMaxUnit.ListIndex = 0
        ElseIf .MaxUnit = "L" Then
            cboMaxUnit.ListIndex = 1
        End If
        If .WasteRateUnit = "T" Then
            cboUnitRate.ListIndex = 0
        ElseIf .WasteRateUnit = "L" Then
            cboUnitRate.ListIndex = 1
        End If
        If .WasteRateTime = "M" Then
            cboTimeRate.ListIndex = 0
        ElseIf .WasteRateTime = "Y" Then
            cboTimeRate.ListIndex = 1
        End If
        callStaffType_ = "MKT"
        Set tmpCol = tmpS.UserSearch("StaffID='" & .MarketingStaffID & "'")
        If tmpCol.Count > 0 Then SetDataStaff tmpCol(1), False
        Set tmpCol = tmpS.IndustryTypeSearch("IndustryTypeID='" & .IndustryTypeID & "'")
        If tmpCol.Count > 0 Then SetDataIndustryType tmpCol(1), False
        callStaffType_ = "ENV"
        Set tmpCol = tmpS.UserSearch("StaffID='" & .EnviStaffID & "'")
        
        If tmpCol.Count > 0 Then SetDataStaff tmpCol(1), False
        Select Case .Status
            Case "P"
                cboSummary.ListIndex = 0
            Case "A", "IA"
                cboSummary.ListIndex = 1
            Case "E"
                cboSummary.ListIndex = 2
        End Select
        
        '*************
        If .isQuotationReq = True Then
            cboWasteAnalysis.ListIndex = 1
        Else
            cboWasteAnalysis.ListIndex = 0
        End If
        '*************
        
        If Trim(.TreatmentType) = "" Then
            cboTreatmentType.ListIndex = -1
        ElseIf Trim(.TreatmentType) = "B" Then
            cboTreatmentType.ListIndex = 0
        ElseIf Trim(.TreatmentType) = "L" Then
            cboTreatmentType.ListIndex = 1
        End If
        
        txtQuoPrice.Text = .QuotationPrice
        
        Set myWasteDataLab_ = Nothing
        Set myWasteDataLab_ = New BWGWasteDataLAB
        myWasteDataLab_.ID = GetGUID
        myWasteDataLab_.wastedataID = myWasteDataCR_.ID
                Set tmpCol = tmpS.WasteDataLABSearch("WasteDataID='" & .ID & "'", isNew)
        If tmpCol.Count > 0 Then Set myWasteDataLab_ = tmpCol(1)
        .isChanged = False
        
    End With
    
    With grKorOrNo.getGridObj
        grKorOrNo.ClearAllData
        .Rows = myWasteDataCR_.KorOrNumbers.Count + 20
        For x = 1 To myWasteDataCR_.KorOrNumbers.Count
            .TextMatrix(x, 0) = myWasteDataCR_.KorOrNumbers(x).ID
            .TextMatrix(x, 1) = myWasteDataCR_.KorOrNumbers(x).KorOrNumber
            .TextMatrix(x, 2) = myWasteDataCR_.KorOrNumbers(x).KorDate
        Next
    End With
    With myWasteDataLab_
        txtGenerateDesc.Text = Trim(.GenerateDetail)
        Select Case .WasteCharacter
            Case "NA"
                cboWasteChar.ListIndex = 0
            Case "SO"
                cboWasteChar.ListIndex = 1
                If .PowderOrLump = "P" Then
                    cboWasteCharDesc.ListIndex = 0
                ElseIf .PowderOrLump = "L" Then
                    cboWasteCharDesc.ListIndex = 1
                ElseIf .PowderOrLump = "N" Then
                    cboWasteCharDesc.ListIndex = 2
                ElseIf .PowderOrLump = "O" Then
                    txtCharDesc.Text = .CharDescOther
                    cboWasteCharDesc.ListIndex = 3
                End If
            Case "SL"
                cboWasteChar.ListIndex = 2
                If .isPumpAble Then
                    cboWasteCharDesc.ListIndex = 0
                Else
                    cboWasteCharDesc.ListIndex = 1
                End If
            Case "LI"
                cboWasteChar.ListIndex = 3
        End Select
        chkExplode.Value = Abs(CInt(.isExplode))
        txtFlashPoint.Text = .FlashPoint
        chkFlame.Value = Abs(CInt(.isFlameAble))
        txtpH.Text = .pHValue
        chkCorrosive = Abs(CInt(.isCorrosive))
        chkReact = Abs(CInt(.isReact))
        chkCancer = Abs(CInt(.isCarcinoGenic))
        chkToxic = Abs(CInt(.isToxic))
        chkRadioActive = Abs(CInt(.isRadioActive))
        chkPesticide = Abs(CInt(.isPesticide))
        txtReactDesc.Text = .ReActWith
        txtCancerDesc.Text = .CarcinoGenicDesc
        If .ToxicType = "-" Then
            cboToxic.ListIndex = -1
        ElseIf .ToxicType = "A" Then
            cboToxic.ListIndex = 0
        ElseIf .ToxicType = "C" Then
            cboToxic.ListIndex = 1
        End If
        txtToxicDesc.Text = .ToxicDesc
        txtRadioActive.Text = .RadioActiveDesc
        txtPesticide.Text = .PesticideDesc
        txtOther.Text = .OtherDesc
        If Trim(.OtherDesc) <> "" Then
            chkOther.Value = 1
        Else
            chkOther.Value = 0
        End If
        If .MixOrSeperate = "" Then
            cboStoreType.ListIndex = 0
        ElseIf .MixOrSeperate = "S" Then
            cboStoreType.ListIndex = 1
        ElseIf .MixOrSeperate = "M" Then
            cboStoreType.ListIndex = 2
        End If
        If .InOrOutDoor = "" Then
            cboIndoor.ListIndex = 0
        ElseIf .InOrOutDoor = "I" Then
            cboIndoor.ListIndex = 1
        ElseIf .InOrOutDoor = "O" Then
            cboIndoor.ListIndex = 2
        End If
        Select Case .StorageType
            Case "BA"
                chkBag.Value = 1
                txtBag.Text = .StorageDesc
            Case "ST"
                chkYard.Value = 1
            Case "DR"
                chkDrum.Value = 1
                txtDrumDesc.Text = .StorageDesc
            Case "OT"
                chkOther2.Value = 1
                txtOther2.Text = .StorageDesc
            Case "TA"
                chkTank.Value = 1
        End Select
        .isNew = isNew
        txtOldStorePlace.Text = .OldEliminator
        For x = 1 To .OldTreatments.Count
            lvOldTreatment.ListItems("'" & .OldTreatments(x).ID & "'").Checked = isExist(lvOldTreatment.ListItems, "'" & .OldTreatments(x).ID & "'")
        Next
        For x = 1 To .NewTreatments.Count
            lvNewTreatment.ListItems("'" & .NewTreatments(x).ID & "'").Checked = isExist(lvNewTreatment.ListItems, "'" & .NewTreatments(x).ID & "'")
        Next
    
        '********** for new Varliable **************
        chkHospital = IIf(.isHospital = True, 1, 0)
        txtOdour = .Odour
        chkOdour = IIf(.isOdour = True, 1, 0)
        txtColor = .Color
        chkColor = IIf(.isColor = True, 1, 0)
        txtPlasticBag = .PlasticBag
        chkPlasticBag = IIf(.isPlasticBag = True, 1, 0)
        txtPlasticDrum = .PlasticDrum
        chkPlasticDrum = IIf(.isPlasticDrum = True, 1, 0)
        txtPlasticGallon = .PlasticGallon
        chkPlasticGallon = IIf(.isPlasticGallon = True, 1, 0)
        txtPaperBox = .PaperBox
        chkPaperBox = IIf(.isPaperBox = True, 1, 0)
        txtBottle = .Bottle
        chkBottle = IIf(.isBottle = True, 1, 0)
        If .AnalysisType = 1 Then
            OptAnalysType1.Value = True
        ElseIf .AnalysisType = 2 Then
            OptAnalysType2.Value = True
        ElseIf .AnalysisType = 3 Then
            txtAnalysDesc = .AnalysisDesc
            OptAnalysType3.Value = True
        End If
        If .isBlending = "Y" Then
            OptBlend1.Value = True
        ElseIf .isBlending = "N" Then
            OptBlend2.Value = True
        Else
            OptBlend0.Value = True
        End If
    End With
    LoadAttacheFile myWasteDataCR_.CustomerID, myWasteDataCR_.ID
    myWasteDataCR_.isChanged = False
    myWasteDataLab_.isChanged = False
End Sub

Public Sub SetDataContract(selContract As BWGContract, Optional isAssigned As Boolean = True)
    txtContractNo.Text = selContract.DocNo
    If isAssigned Then
        Set myWasteDataCR_.Contract = selContract
    End If
End Sub

Public Sub SetDate(selDate As String)
    myWasteDataCR_.LicenseExpDate = selDate
    txtExpDate.Text = FormatYMD_to_DMY(selDate, "/", "/")
End Sub

Public Sub SetDataIndustryType(selIndustry As BWGIndustryType, Optional isAssigned As Boolean = True)
'    Set curIndustryType_ = selIndustry
'    With curIndustryType_
'        txtIndustryType.Text = .TypeName
'        If isAssigned Then
'            myWasteDataCR_.IndustryTypeID = .ID
'            myWasteDataCR_.IndustryTypeName = .TypeName
'        End If
'    End With
End Sub

Private Sub SetUserGroup()
Dim cVal As Boolean
Dim UserGrp As Collection
Dim Group As New AWSUserGroup
Dim i As Integer
    Set UserGrp = CurrentUser.getUserGroups
    For i = 1 To UserGrp.Count
        Set Group = UserGrp(i)
        If Group.Name = "Administrator" Then
            StrGrp = Group.Name
            Exit For
        Else
            StrGrp = Group.Name
        End If
    Next
    
    If LCase(StrGrp) = "environment" Then
        isEnv = True
        cboLABRes.Enabled = True
'        cmdReword.Visible = True
        cmdApprove.Caption = "ตรวจรับโดย Envi"
'        cmdReject.Visible = True
        SSTab1.TabEnabled(2) = True
        lvNewTreatment.Enabled = True
        txtEnviStaff.Enabled = True
        cboSummary.Enabled = True
    Else
        isEnv = False
        cboLABRes.Enabled = False
        cmdReword.Visible = False
        cboWasteType.Enabled = True
        cmdApprove.Caption = "ส่งให้สิ่งแวดล้อม"
        cmdReject.Visible = False
        'SSTab1.TabEnabled(2) = False
        lvNewTreatment.Enabled = False
        txtEnviStaff.Enabled = False
        cboSummary.Enabled = False
        cboWasteType.Enabled = False
    End If
    cmdApprove.Picture = frmPicture.picApprove.Picture
End Sub

Private Function SaveAttache(ByVal CustomerID As String, ByVal wastedataID As String) As Boolean
On Error GoTo ErrD
Dim i As Integer
Dim ServerPath As String
Dim SourceFile As String
    SaveAttache = True
    ServerPath = DefaultPath & CustomerID & "\"
    '======= Check Path ===================
    If Dir(ServerPath, vbDirectory) = "" Then
      MkDir ServerPath
    End If
    ServerPath = ServerPath & wastedataID & "\"
    If Dir(ServerPath, vbDirectory) = "" Then
      MkDir ServerPath
    End If
    
    For i = 0 To 3
        If Attaches(i).FileName <> "" Then
            FileCopy Attaches(i).FullPath, ServerPath & Attaches(i).FileName
        End If
    Next
    Exit Function
ErrD:
    SaveAttache = False
    Resume Next
End Function

Private Function KillFile(ByVal Index As Integer, ByVal CustomerID As String, ByVal wastedataID As String) As Boolean
On Error GoTo ErrD
Dim FileName As String
    KillFile = True
    FileName = DefaultPath & CustomerID & "\" & wastedataID & "\" & lblAttach(Index).Caption
    If Dir(FileName, vbNormal) <> "" Then
        Kill FileName
        lblAttach(Index).Caption = ""
    End If
    Exit Function
ErrD:
    KillFile = False
End Function

Private Sub LoadAttacheFile(ByVal CustomerID As String, ByVal wastedataID As String)
On Error GoTo ErrD
Dim FSO As New FileSystemObject
Dim Filex As File
Dim FileObject As Files
Dim FolderObject As Folder
Dim ServerFilePath As String
Dim i As Integer, j As Integer, x As Integer, Files As String
Dim Extension As String

    ServerFilePath = DefaultPath & CustomerID & "\" & wastedataID
    If FSO.FolderExists(ServerFilePath) = True Then
        Set FolderObject = FSO.GetFolder(ServerFilePath)
        Set FileObject = FolderObject.Files
        i = FileObject.Count
        x = 0
        ReDim Attaches(4)
        For Each Filex In FileObject
            For j = Len(Filex.Name) To 1 Step -1
                Extension = Right(Filex.Name, Len(Filex.Name) - j)
                If Mid(Filex.Name, j, 1) = "." Then Exit For
            Next
            If (LCase(Extension) <> "db") And (LCase(Extension) <> "dll") And (LCase(Extension) <> "exe") And (LCase(Extension) <> "com") And (LCase(Extension) <> "bat") Then
                Attaches(x).FileName = Filex.Name
                Attaches(x).FullPath = ServerFilePath & "\" & Filex.Name
                lblAttach(x).Caption = Filex.Name
                x = x + 1
            End If
            If x > 3 Then Exit For
        Next Filex
    End If
    
    Exit Sub
ErrD:
    MsgBox err.Description
End Sub
