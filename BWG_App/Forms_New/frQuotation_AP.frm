VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmQuotation 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "F111"
   ClientHeight    =   9585
   ClientLeft      =   75
   ClientTop       =   855
   ClientWidth     =   14385
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9585
   ScaleWidth      =   14385
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      BackColor       =   &H80000005&
      Caption         =   "ลูกค้า BME กลุ่มใหม่"
      Height          =   255
      Left            =   0
      TabIndex        =   164
      Top             =   0
      Width           =   1935
   End
   Begin TabDlg.SSTab SSTabQuo 
      Height          =   5280
      Left            =   0
      TabIndex        =   87
      Top             =   3120
      Width           =   14370
      _ExtentX        =   25347
      _ExtentY        =   9313
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "รายการของเสีย"
      TabPicture(0)   =   "frQuotation_AP.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "หมายเหตุ"
      TabPicture(1)   =   "frQuotation_AP.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "ประเภทรถ"
      TabPicture(2)   =   "frQuotation_AP.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label2(51)"
      Tab(2).Control(1)=   "Label2(50)"
      Tab(2).Control(2)=   "Label2(49)"
      Tab(2).Control(3)=   "Label2(48)"
      Tab(2).Control(4)=   "Label2(47)"
      Tab(2).Control(5)=   "Label2(46)"
      Tab(2).Control(6)=   "Label2(45)"
      Tab(2).Control(7)=   "Label2(43)"
      Tab(2).Control(8)=   "Label2(42)"
      Tab(2).Control(9)=   "Label2(41)"
      Tab(2).Control(10)=   "Label2(40)"
      Tab(2).Control(11)=   "Label2(39)"
      Tab(2).Control(12)=   "Label2(38)"
      Tab(2).Control(13)=   "Label2(33)"
      Tab(2).Control(14)=   "Label2(32)"
      Tab(2).Control(15)=   "Label2(44)"
      Tab(2).Control(16)=   "lblเวปไซท์(4)"
      Tab(2).Control(17)=   "lblเวปไซท์(3)"
      Tab(2).Control(18)=   "lblเวปไซท์(2)"
      Tab(2).Control(19)=   "lblCustomerSite(0)"
      Tab(2).Control(20)=   "cboIsTrailer9"
      Tab(2).Control(21)=   "cboTransUnit9"
      Tab(2).Control(22)=   "txtTransRate9"
      Tab(2).Control(23)=   "txtCarType9"
      Tab(2).Control(24)=   "txtCarType8"
      Tab(2).Control(25)=   "txtTransRate8"
      Tab(2).Control(26)=   "cboTransUnit8"
      Tab(2).Control(27)=   "cboIsTrailer8"
      Tab(2).Control(28)=   "txtCarType7"
      Tab(2).Control(29)=   "txtTransRate7"
      Tab(2).Control(30)=   "cboTransUnit7"
      Tab(2).Control(31)=   "cboIsTrailer7"
      Tab(2).Control(32)=   "txtCarType6"
      Tab(2).Control(33)=   "txtTransRate6"
      Tab(2).Control(34)=   "cboTransUnit6"
      Tab(2).Control(35)=   "cboIsTrailer6"
      Tab(2).Control(36)=   "txtCarType5"
      Tab(2).Control(37)=   "txtTransRate5"
      Tab(2).Control(38)=   "cboTransUnit5"
      Tab(2).Control(39)=   "cboIsTrailer5"
      Tab(2).Control(40)=   "txtCarType2"
      Tab(2).Control(41)=   "txtTransRate2"
      Tab(2).Control(42)=   "cboTransUnit2"
      Tab(2).Control(43)=   "cboIsTrailer2"
      Tab(2).Control(44)=   "txtCarType4"
      Tab(2).Control(45)=   "txtTransRate4"
      Tab(2).Control(46)=   "cboTransUnit4"
      Tab(2).Control(47)=   "cboIsTrailer4"
      Tab(2).Control(48)=   "txtCarType3"
      Tab(2).Control(49)=   "txtTransRate3"
      Tab(2).Control(50)=   "cboTransUnit3"
      Tab(2).Control(51)=   "cboIsTrailer3"
      Tab(2).Control(52)=   "GrCar"
      Tab(2).Control(53)=   "CmdSelTruck"
      Tab(2).Control(54)=   "txtProvince"
      Tab(2).Control(54).Enabled=   0   'False
      Tab(2).Control(55)=   "txtAumphur"
      Tab(2).Control(55).Enabled=   0   'False
      Tab(2).Control(56)=   "txtTumbol"
      Tab(2).ControlCount=   57
      Begin VB.TextBox txtTumbol 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   -72840
         TabIndex        =   168
         Top             =   360
         Width           =   1395
      End
      Begin VB.TextBox txtAumphur 
         Appearance      =   0  'Flat
         BackColor       =   &H000000C0&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   -70860
         TabIndex        =   167
         TabStop         =   0   'False
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox txtProvince 
         Appearance      =   0  'Flat
         BackColor       =   &H000000C0&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   -68880
         TabIndex        =   166
         TabStop         =   0   'False
         Top             =   360
         Width           =   1275
      End
      Begin VB.CommandButton CmdSelTruck 
         Caption         =   "เลือกรถ"
         Height          =   690
         Left            =   -61920
         TabIndex        =   154
         Top             =   840
         Width           =   1140
      End
      Begin WasteManagment.ctlGrid GrCar 
         Height          =   4365
         Left            =   -74955
         TabIndex        =   153
         Top             =   840
         Width           =   12930
         _ExtentX        =   22807
         _ExtentY        =   7699
      End
      Begin VB.ComboBox cboIsTrailer3 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frQuotation_AP.frx":0054
         Left            =   -72330
         List            =   "frQuotation_AP.frx":005E
         Style           =   2  'Dropdown List
         TabIndex        =   138
         Top             =   4455
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.ComboBox cboTransUnit3 
         Enabled         =   0   'False
         Height          =   315
         Left            =   -69375
         Style           =   2  'Dropdown List
         TabIndex        =   137
         Top             =   4455
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox txtTransRate3 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   315
         Left            =   -70455
         TabIndex        =   136
         Top             =   4455
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.TextBox txtCarType3 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Enabled         =   0   'False
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   -74010
         TabIndex        =   135
         Top             =   4455
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.ComboBox cboIsTrailer4 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frQuotation_AP.frx":0074
         Left            =   -72330
         List            =   "frQuotation_AP.frx":007E
         Style           =   2  'Dropdown List
         TabIndex        =   134
         Top             =   4800
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.ComboBox cboTransUnit4 
         Enabled         =   0   'False
         Height          =   315
         Left            =   -69375
         Style           =   2  'Dropdown List
         TabIndex        =   133
         Top             =   4800
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox txtTransRate4 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   315
         Left            =   -70455
         TabIndex        =   132
         Top             =   4800
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.TextBox txtCarType4 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Enabled         =   0   'False
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   -74010
         TabIndex        =   131
         Top             =   4800
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.ComboBox cboIsTrailer2 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frQuotation_AP.frx":0094
         Left            =   -72330
         List            =   "frQuotation_AP.frx":009E
         Style           =   2  'Dropdown List
         TabIndex        =   115
         Top             =   4095
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.ComboBox cboTransUnit2 
         Enabled         =   0   'False
         Height          =   315
         Left            =   -69375
         Style           =   2  'Dropdown List
         TabIndex        =   114
         Top             =   4095
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox txtTransRate2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   315
         Left            =   -70455
         TabIndex        =   113
         Top             =   4095
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.TextBox txtCarType2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   -74010
         TabIndex        =   112
         Top             =   4095
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.ComboBox cboIsTrailer5 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frQuotation_AP.frx":00B4
         Left            =   -72330
         List            =   "frQuotation_AP.frx":00BE
         Style           =   2  'Dropdown List
         TabIndex        =   111
         Top             =   5190
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.ComboBox cboTransUnit5 
         Enabled         =   0   'False
         Height          =   315
         Left            =   -69375
         Style           =   2  'Dropdown List
         TabIndex        =   110
         Top             =   5190
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox txtTransRate5 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   315
         Left            =   -70455
         TabIndex        =   109
         Top             =   5190
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.TextBox txtCarType5 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   -74010
         TabIndex        =   108
         Top             =   5190
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.ComboBox cboIsTrailer6 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frQuotation_AP.frx":00D4
         Left            =   -65085
         List            =   "frQuotation_AP.frx":00DE
         Style           =   2  'Dropdown List
         TabIndex        =   107
         Top             =   3735
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.ComboBox cboTransUnit6 
         Enabled         =   0   'False
         Height          =   315
         Left            =   -62130
         Style           =   2  'Dropdown List
         TabIndex        =   106
         Top             =   3735
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox txtTransRate6 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   315
         Left            =   -63210
         TabIndex        =   105
         Top             =   3735
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.TextBox txtCarType6 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   -66765
         TabIndex        =   104
         Top             =   3735
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.ComboBox cboIsTrailer7 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frQuotation_AP.frx":00F4
         Left            =   -65085
         List            =   "frQuotation_AP.frx":00FE
         Style           =   2  'Dropdown List
         TabIndex        =   103
         Top             =   4065
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.ComboBox cboTransUnit7 
         Enabled         =   0   'False
         Height          =   315
         Left            =   -62130
         Style           =   2  'Dropdown List
         TabIndex        =   102
         Top             =   4095
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox txtTransRate7 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   315
         Left            =   -63210
         TabIndex        =   101
         Top             =   4065
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.TextBox txtCarType7 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   -66765
         TabIndex        =   100
         Top             =   4065
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.ComboBox cboIsTrailer8 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frQuotation_AP.frx":0114
         Left            =   -65085
         List            =   "frQuotation_AP.frx":011E
         Style           =   2  'Dropdown List
         TabIndex        =   99
         Top             =   4455
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.ComboBox cboTransUnit8 
         Enabled         =   0   'False
         Height          =   315
         Left            =   -62130
         Style           =   2  'Dropdown List
         TabIndex        =   98
         Top             =   4455
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox txtTransRate8 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   315
         Left            =   -63210
         TabIndex        =   97
         Top             =   4455
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.TextBox txtCarType8 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   -66765
         TabIndex        =   96
         Top             =   4455
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.TextBox txtCarType9 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   -66765
         TabIndex        =   95
         Top             =   4815
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.TextBox txtTransRate9 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   315
         Left            =   -63210
         TabIndex        =   94
         Top             =   4815
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.ComboBox cboTransUnit9 
         Enabled         =   0   'False
         Height          =   315
         Left            =   -62130
         Style           =   2  'Dropdown List
         TabIndex        =   93
         Top             =   4815
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.ComboBox cboIsTrailer9 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frQuotation_AP.frx":0134
         Left            =   -65085
         List            =   "frQuotation_AP.frx":013E
         Style           =   2  'Dropdown List
         TabIndex        =   92
         Top             =   4815
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H8000000E&
         Height          =   5205
         Left            =   -74955
         TabIndex        =   90
         Top             =   360
         Width           =   14250
         Begin WasteManagment.ctlGrid grRemark 
            Height          =   4890
            Left            =   0
            TabIndex        =   91
            Top             =   0
            Width           =   14235
            _ExtentX        =   25109
            _ExtentY        =   8625
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H8000000E&
         Caption         =   "รายการการให้บริการการขนกากของเสีย"
         Height          =   4770
         Left            =   135
         TabIndex        =   88
         Top             =   450
         Width           =   14070
         Begin WasteManagment.ctlGrid grItem 
            Height          =   4470
            Left            =   0
            TabIndex        =   89
            Top             =   270
            Width           =   13965
            _ExtentX        =   24633
            _ExtentY        =   7885
         End
         Begin WasteManagment.ctlGrid grItemAP 
            Height          =   4470
            Left            =   0
            TabIndex        =   151
            Top             =   270
            Width           =   13995
            _ExtentX        =   24686
            _ExtentY        =   7885
         End
      End
      Begin VB.Label lblCustomerSite 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "สถานที่เก็บขน"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   -74760
         TabIndex        =   172
         Top             =   420
         Width           =   990
      End
      Begin VB.Label lblเวปไซท์ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ตำบล"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   -73320
         TabIndex        =   171
         Top             =   420
         Width           =   375
      End
      Begin VB.Label lblเวปไซท์ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "อำเภอ"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   3
         Left            =   -71340
         TabIndex        =   170
         Top             =   420
         Width           =   435
      End
      Begin VB.Label lblเวปไซท์ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "จังหวัด"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   4
         Left            =   -69480
         TabIndex        =   169
         Top             =   420
         Width           =   495
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ค่าขนส่ง"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   44
         Left            =   -63990
         TabIndex        =   152
         Top             =   3795
         Visible         =   0   'False
         Width           =   675
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ค่าขนส่ง"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   32
         Left            =   -71235
         TabIndex        =   130
         Top             =   4155
         Visible         =   0   'False
         Width           =   675
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ประเภทรถ"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   33
         Left            =   -75630
         TabIndex        =   129
         Top             =   4155
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ค่าขนส่ง"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   38
         Left            =   -71235
         TabIndex        =   128
         Top             =   4515
         Visible         =   0   'False
         Width           =   675
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ประเภทรถ"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   39
         Left            =   -74910
         TabIndex        =   127
         Top             =   4515
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ค่าขนส่ง"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   40
         Left            =   -71235
         TabIndex        =   126
         Top             =   4860
         Visible         =   0   'False
         Width           =   675
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ประเภทรถ"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   41
         Left            =   -75630
         TabIndex        =   125
         Top             =   4860
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ค่าขนส่ง"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   42
         Left            =   -71235
         TabIndex        =   124
         Top             =   5250
         Visible         =   0   'False
         Width           =   675
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ประเภทรถ"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   43
         Left            =   -74910
         TabIndex        =   123
         Top             =   5250
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ประเภทรถ"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   45
         Left            =   -68385
         TabIndex        =   122
         Top             =   3780
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ค่าขนส่ง"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   46
         Left            =   -63990
         TabIndex        =   121
         Top             =   4125
         Visible         =   0   'False
         Width           =   675
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ประเภทรถ"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   47
         Left            =   -67665
         TabIndex        =   120
         Top             =   4125
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ค่าขนส่ง"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   48
         Left            =   -63990
         TabIndex        =   119
         Top             =   4515
         Visible         =   0   'False
         Width           =   675
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ประเภทรถ"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   49
         Left            =   -68385
         TabIndex        =   118
         Top             =   4515
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ประเภทรถ"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   50
         Left            =   -68385
         TabIndex        =   117
         Top             =   4875
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ค่าขนส่ง"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   51
         Left            =   -63990
         TabIndex        =   116
         Top             =   4875
         Visible         =   0   'False
         Width           =   675
         WordWrap        =   -1  'True
      End
   End
   Begin VB.CommandButton cmdPrintCar 
      Caption         =   "พิมพ์สำหรับให้ลูกค้าเลือกรถ"
      Height          =   705
      Left            =   2025
      Style           =   1  'Graphical
      TabIndex        =   80
      Top             =   8865
      Width           =   2415
   End
   Begin VB.CommandButton cmdPrintPDF 
      Caption         =   "พิมพ์ PDF"
      Height          =   705
      Left            =   1035
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   8850
      Width           =   945
   End
   Begin VB.ComboBox cboPrintForm 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   9300
      Style           =   2  'Dropdown List
      TabIndex        =   67
      TabStop         =   0   'False
      Top             =   520
      Width           =   3570
   End
   Begin VB.TextBox txtRem 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   7260
      TabIndex        =   21
      Top             =   8430
      Width           =   3585
   End
   Begin VB.CommandButton cmdCloseBill 
      Caption         =   "หยุดใช้งานใบเสนอราคา"
      Height          =   375
      Left            =   10920
      TabIndex        =   32
      Top             =   8400
      Width           =   1965
   End
   Begin VB.TextBox txtPosition 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   4320
      TabIndex        =   20
      Top             =   8430
      Width           =   2085
   End
   Begin VB.ComboBox cboApprove 
      Height          =   315
      Left            =   960
      TabIndex        =   19
      Top             =   8430
      Width           =   2505
   End
   Begin VB.CommandButton Command1 
      Caption         =   "เลือกหมายเหตุ"
      Height          =   705
      Left            =   8025
      TabIndex        =   26
      Top             =   8850
      Width           =   1305
   End
   Begin VB.ComboBox cboCompany 
      Height          =   315
      Left            =   9300
      Style           =   2  'Dropdown List
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   200
      Width           =   3555
   End
   Begin Crystal.CrystalReport rpt1 
      Left            =   1830
      Top             =   30
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin VB.CommandButton cmdApproved 
      Caption         =   "อนุมัติใบเสนอราคา"
      Height          =   705
      Left            =   5130
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   8850
      Width           =   1515
   End
   Begin VB.TextBox txtGrandTotalAmt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C00000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   8970
      Locked          =   -1  'True
      TabIndex        =   52
      Top             =   7860
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.TextBox txtVATAmt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C00000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   6120
      Locked          =   -1  'True
      TabIndex        =   51
      Top             =   7860
      Visible         =   0   'False
      Width           =   1620
   End
   Begin VB.TextBox txtDiscountAmt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C00000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   3810
      Locked          =   -1  'True
      TabIndex        =   50
      Top             =   7860
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.TextBox txtTotalAmt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C00000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   900
      Locked          =   -1  'True
      TabIndex        =   49
      Top             =   7860
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.CommandButton cmdAddItem 
      Caption         =   "เลือกกาก"
      Height          =   705
      Left            =   6675
      TabIndex        =   25
      Top             =   8850
      Width           =   1305
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "พิมพ์"
      Height          =   705
      Left            =   45
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   8865
      Width           =   945
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ยกเลิก"
      Height          =   705
      Left            =   12345
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   8850
      Width           =   945
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "ปิด"
      Height          =   705
      Left            =   13335
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   8850
      Width           =   945
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "สร้างใหม่"
      Height          =   705
      Left            =   11355
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   8850
      Width           =   945
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "เปิดข้อมูล"
      Height          =   705
      Left            =   10365
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   8850
      Width           =   945
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "เก็บข้อมูล"
      Height          =   705
      Left            =   9375
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   8850
      Width           =   945
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000E&
      Caption         =   "รายละเอียดใบเสนอราคา"
      Height          =   2445
      Left            =   0
      TabIndex        =   35
      Top             =   675
      Width           =   14370
      Begin VB.CheckBox chkTrnDis 
         BackColor       =   &H80000005&
         Caption         =   "ฟรี่ค่าขนส่ง"
         Height          =   255
         Left            =   10800
         TabIndex        =   173
         Top             =   1680
         Width           =   1455
      End
      Begin VB.CheckBox chkcarbt 
         BackColor       =   &H80000005&
         Caption         =   "รถ BT"
         Height          =   255
         Left            =   10440
         TabIndex        =   165
         Top             =   1320
         Width           =   1335
      End
      Begin VB.CheckBox chkBmeNew 
         BackColor       =   &H80000005&
         Caption         =   "ลูกค้า BME กลุ่มใหม่"
         Height          =   255
         Left            =   10440
         TabIndex        =   163
         Top             =   1020
         Width           =   1935
      End
      Begin VB.ComboBox cbProject 
         Height          =   315
         Left            =   8460
         TabIndex        =   161
         Text            =   "Combo1"
         Top             =   1680
         Width           =   2175
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H8000000E&
         Height          =   780
         Left            =   13800
         TabIndex        =   155
         Top             =   1215
         Visible         =   0   'False
         Width           =   480
         Begin VB.CheckBox ChkHalfTrn 
            BackColor       =   &H8000000E&
            Caption         =   "ค่าขนส่งครึ่งราคา(ค่าคอมฯ)"
            ForeColor       =   &H000000C0&
            Height          =   195
            Left            =   135
            TabIndex        =   157
            Top             =   495
            Width           =   2175
         End
         Begin VB.CheckBox ChkFreeTrn 
            BackColor       =   &H8000000E&
            Caption         =   "ฟรีค่าขนส่ง(ค่าคอมฯ)"
            ForeColor       =   &H000000C0&
            Height          =   195
            Left            =   135
            TabIndex        =   156
            Top             =   180
            Width           =   1770
         End
      End
      Begin VB.TextBox txtBMEQuotation 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         Height          =   285
         Left            =   12780
         TabIndex        =   85
         Top             =   405
         Width           =   1455
      End
      Begin VB.CommandButton CmdQuoTruck 
         BackColor       =   &H00C0FFFF&
         Caption         =   "ใบเสนอราคาสำหรับเลือกรถ"
         Height          =   780
         Left            =   12060
         Style           =   1  'Graphical
         TabIndex        =   84
         Top             =   2340
         Visible         =   0   'False
         Width           =   2220
      End
      Begin VB.TextBox txtRefQuotationID 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         Height          =   285
         Left            =   12810
         TabIndex        =   79
         Top             =   945
         Width           =   1455
      End
      Begin VB.ComboBox cboMaxRateType 
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   7635
         TabIndex        =   77
         Top             =   2040
         Width           =   1215
      End
      Begin VB.TextBox txtMaxRate 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   6915
         TabIndex        =   18
         Top             =   2040
         Width           =   675
      End
      Begin VB.TextBox txtMaxWeight 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1485
         TabIndex        =   17
         Top             =   2040
         Width           =   675
      End
      Begin VB.CommandButton Command2 
         Caption         =   "บันทึกข้อตกลง"
         Height          =   315
         Left            =   9240
         TabIndex        =   73
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox txtLastUser 
         Appearance      =   0  'Flat
         BackColor       =   &H000040C0&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   5700
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   960
         Width           =   2475
      End
      Begin VB.ComboBox cboQuotationType 
         Height          =   315
         ItemData        =   "frQuotation_AP.frx":0154
         Left            =   9240
         List            =   "frQuotation_AP.frx":0156
         TabIndex        =   3
         Top             =   600
         Width           =   1995
      End
      Begin VB.ComboBox cboIsTrailer 
         Height          =   315
         ItemData        =   "frQuotation_AP.frx":0158
         Left            =   2625
         List            =   "frQuotation_AP.frx":0162
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1680
         Width           =   1215
      End
      Begin VB.ComboBox cboTransUnit 
         Height          =   315
         ItemData        =   "frQuotation_AP.frx":0178
         Left            =   5580
         List            =   "frQuotation_AP.frx":017A
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1680
         Width           =   1455
      End
      Begin VB.TextBox txtTransRate 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   4500
         TabIndex        =   12
         Top             =   1680
         Width           =   1035
      End
      Begin VB.TextBox txtCarType 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   945
         TabIndex        =   10
         Top             =   1680
         Width           =   1635
      End
      Begin VB.TextBox txtExtraCharge 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   12960
         TabIndex        =   15
         Top             =   1995
         Width           =   675
      End
      Begin VB.CheckBox chkIncVAT 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "ราคารวม VAT แล้ว"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   8460
         TabIndex        =   9
         Top             =   1380
         Width           =   1695
      End
      Begin VB.TextBox txtMinQty 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   4305
         TabIndex        =   16
         Top             =   2040
         Width           =   675
      End
      Begin VB.TextBox txtLaborCount 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   10155
         TabIndex        =   14
         Top             =   1995
         Width           =   675
      End
      Begin VB.CheckBox chkIncTrans 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "ราคารวมค่าขนส่งแล้ว"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   8460
         TabIndex        =   8
         Top             =   1020
         Width           =   1815
      End
      Begin WasteManagment.ctlDate dtIssuedDate 
         Height          =   315
         Left            =   5700
         TabIndex        =   2
         Top             =   600
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         BackColor       =   -2147483633
      End
      Begin VB.TextBox txtApprovedStaff 
         Appearance      =   0  'Flat
         BackColor       =   &H00C00000&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1740
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1320
         Width           =   2595
      End
      Begin VB.ComboBox cboCustStaff 
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   5700
         TabIndex        =   7
         Top             =   1320
         Width           =   2535
      End
      Begin VB.TextBox txtIssuedStaff 
         Appearance      =   0  'Flat
         BackColor       =   &H00C00000&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1740
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   960
         Width           =   2595
      End
      Begin VB.TextBox txtCustomer 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1740
         TabIndex        =   0
         Top             =   240
         Width           =   7395
      End
      Begin VB.TextBox txtQuoNo 
         Appearance      =   0  'Flat
         BackColor       =   &H00C00000&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1740
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Text            =   "== AUTO =="
         Top             =   600
         Width           =   2595
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "สำหรับโปรเจค"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   35
         Left            =   7365
         TabIndex        =   162
         Top             =   1740
         Width           =   1005
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblY 
         Height          =   240
         Left            =   11250
         TabIndex        =   160
         Top             =   1665
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.Label lblM 
         Height          =   240
         Left            =   10665
         TabIndex        =   159
         Top             =   1665
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.Label lblD 
         Height          =   285
         Left            =   10080
         TabIndex        =   158
         Top             =   1620
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "อ้างอิงใบเสนอราคา (ปลายทางอื่น)"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   52
         Left            =   10110
         TabIndex        =   86
         Top             =   180
         Width           =   4155
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "อ้างอิงใบเสนอราคา(อัคคี)"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   30
         Left            =   12500
         TabIndex        =   78
         Top             =   720
         Width           =   1750
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "อัตราส่วนเกิน"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   31
         Left            =   5550
         TabIndex        =   76
         Top             =   2100
         Width           =   1320
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ปริมาณขนสูงสุด"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   29
         Left            =   225
         TabIndex        =   75
         Top             =   2100
         Width           =   1215
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ตัน/คัน"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   28
         Left            =   2205
         TabIndex        =   74
         Top             =   2100
         Width           =   615
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ผู้บันทึกล่าสุด"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   27
         Left            =   4620
         TabIndex        =   69
         Top             =   1020
         Width           =   1035
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ประเภทใบเสนอราคา"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   26
         Left            =   7620
         TabIndex        =   68
         Top             =   660
         Width           =   1575
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ค่าขนส่ง"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   14
         Left            =   3720
         TabIndex        =   63
         Top             =   1740
         Width           =   675
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ประเภทรถ"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   13
         Left            =   45
         TabIndex        =   62
         Top             =   1740
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ค่าแรงคนงานเพิ่ม"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   10
         Left            =   11640
         TabIndex        =   61
         Top             =   2055
         Width           =   1275
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "บาท/คน"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   3
         Left            =   13680
         TabIndex        =   60
         Top             =   2055
         Width           =   615
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ตัน/คัน"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   9
         Left            =   5025
         TabIndex        =   59
         Top             =   2100
         Width           =   615
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ปริมาณขนขั้นต่ำ"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   7
         Left            =   3045
         TabIndex        =   58
         Top             =   2100
         Width           =   1215
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "คน/คัน"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   6
         Left            =   10875
         TabIndex        =   57
         Top             =   2055
         Width           =   615
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "รวมคนงาน"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   5
         Left            =   9315
         TabIndex        =   56
         Top             =   2055
         Width           =   795
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ผู้อนุมัติใบเสนอราคา"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   55
         Top             =   1380
         Width           =   1575
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ผู้รับใบเสนอราคา"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   11
         Left            =   4440
         TabIndex        =   41
         Top             =   1380
         Width           =   1215
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ผู้ออกใบเสนอราคา"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   39
         Top             =   1020
         Width           =   1575
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ชื่อลูกค้า"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   8
         Left            =   1125
         TabIndex        =   38
         Top             =   300
         Width           =   570
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "เลขที่ใบเสนอราคา"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   37
         Top             =   660
         Width           =   1575
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "วันที่ใบเสนอราคา"
         Height          =   195
         Index           =   8
         Left            =   4425
         TabIndex        =   36
         Top             =   660
         Width           =   1230
      End
   End
   Begin VB.TextBox txtDiscount 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1020
      TabIndex        =   71
      Top             =   1140
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.ComboBox cboDiscountType 
      Height          =   315
      Left            =   1740
      Style           =   2  'Dropdown List
      TabIndex        =   70
      Top             =   1140
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.TextBox txtHBLRate 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   4230
      TabIndex        =   139
      Top             =   2745
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.TextBox txtNBLRate 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   4230
      TabIndex        =   140
      Top             =   2385
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.TextBox txtHazRate 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1650
      TabIndex        =   141
      Top             =   2745
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.TextBox txtNonRate 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1650
      TabIndex        =   142
      Top             =   2385
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ประเภทรถ"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   37
      Left            =   6570
      TabIndex        =   83
      Top             =   2850
      Width           =   1575
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ค่าขนส่ง"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   36
      Left            =   10965
      TabIndex        =   82
      Top             =   2850
      Width           =   675
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ค่าขนส่ง"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   34
      Left            =   3720
      TabIndex        =   81
      Top             =   2865
      Width           =   675
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "หมายเหตุ"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   25
      Left            =   6510
      TabIndex        =   66
      Top             =   8490
      Width           =   780
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ตำแหน่ง"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   16
      Left            =   3660
      TabIndex        =   65
      Top             =   8490
      Width           =   780
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ผู้เซ็นอนุมัติ"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   15
      Left            =   90
      TabIndex        =   64
      Top             =   8490
      Width           =   915
      WordWrap        =   -1  'True
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
      Height          =   615
      Left            =   4020
      TabIndex        =   54
      Top             =   30
      Visible         =   0   'False
      Width           =   2865
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "บาท"
      Height          =   195
      Index           =   3
      Left            =   10980
      TabIndex        =   53
      Top             =   7920
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "รวมทั้งสิ้น"
      Height          =   195
      Index           =   4
      Left            =   8250
      TabIndex        =   48
      Top             =   7920
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VAT"
      Height          =   195
      Index           =   6
      Left            =   5730
      TabIndex        =   47
      Top             =   7920
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "มูลค่าสินค้า"
      Height          =   195
      Index           =   19
      Left            =   90
      TabIndex        =   46
      Top             =   7920
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "บาท"
      Height          =   195
      Index           =   20
      Left            =   2880
      TabIndex        =   45
      Top             =   7920
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ส่วนลด"
      Height          =   195
      Index           =   11
      Left            =   3270
      TabIndex        =   44
      Top             =   7920
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "บาท"
      Height          =   195
      Index           =   12
      Left            =   5340
      TabIndex        =   43
      Top             =   7920
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "บาท"
      Height          =   195
      Index           =   21
      Left            =   7800
      TabIndex        =   42
      Top             =   7920
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "รายการรออนุมัติ, หรืออนุมัติแล้วเมื่อวันที่..... โดย...."
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Index           =   4
      Left            =   9330
      TabIndex        =   40
      Top             =   0
      Width           =   3555
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   4
      Index           =   0
      X1              =   90
      X2              =   7560
      Y1              =   450
      Y2              =   450
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ใบเสนอราคา"
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
      Left            =   150
      TabIndex        =   34
      Top             =   60
      Width           =   1095
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      Index           =   1
      X1              =   180
      X2              =   7650
      Y1              =   525
      Y2              =   525
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   14310
      Y1              =   8790
      Y2              =   8790
   End
   Begin VB.Image Image2 
      Height          =   645
      Left            =   0
      Picture         =   "frQuotation_AP.frx":017C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   14400
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ส่วนลด"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   12
      Left            =   420
      TabIndex        =   72
      Top             =   1200
      Visible         =   0   'False
      Width           =   555
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ราคาขั้นต่ำ Non"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   18
      Left            =   0
      TabIndex        =   150
      Top             =   2415
      Visible         =   0   'False
      Width           =   1575
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "บาท/ตัน"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   17
      Left            =   2370
      TabIndex        =   149
      Top             =   2445
      Visible         =   0   'False
      Width           =   615
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "บาท/ตัน"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   19
      Left            =   2370
      TabIndex        =   148
      Top             =   2805
      Visible         =   0   'False
      Width           =   615
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ราคาขั้นต่ำ Haz"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   20
      Left            =   390
      TabIndex        =   147
      Top             =   2805
      Visible         =   0   'False
      Width           =   1215
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "บาท/ตัน"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   21
      Left            =   4950
      TabIndex        =   146
      Top             =   2445
      Visible         =   0   'False
      Width           =   615
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ราคาขั้นต่ำ NBL"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   22
      Left            =   3090
      TabIndex        =   145
      Top             =   2445
      Visible         =   0   'False
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "บาท/ตัน"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   23
      Left            =   4950
      TabIndex        =   144
      Top             =   2805
      Visible         =   0   'False
      Width           =   615
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ราคาขั้นต่ำ HBL"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   24
      Left            =   3090
      TabIndex        =   143
      Top             =   2805
      Visible         =   0   'False
      Width           =   1095
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmQuotation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim myQuotation_ As BWGQuotationHeader
Dim curCustomer_ As BWGCustomer
Dim tmpS As BWGSearchManager
Dim isDetailChange_ As Boolean
Dim curCarType_ As BWGTruckType
Dim curRole_ As AWSRoleType
Dim colPendingJob_ As Collection
Dim curTumbol_ As IMITumbol
Dim isSave As Boolean
Dim SignName As Collection
Dim ObjSel As String
Public QuotationPromies As String
Public QuotationStartDatePromies As String
Public QuotationEndDatePromies As String
Public frmPromiesChang As Boolean
Dim colSearchResult_ As Collection

Private Sub ValidateData()
    If cboPrintForm.ListIndex = 0 Then
        grItem.Visible = True
        grItemAP.Visible = False
    ElseIf cboPrintForm.ListIndex = 1 Then
        grItem.Visible = False
        grItemAP.Visible = True
    End If
    myQuotation_.QuotationType = cboPrintForm.ListIndex + 1
    myQuotation_.isChanged = True
    ShowQuoDetails
End Sub

Private Sub cboApprove_Change()
    If cboApprove.ListIndex > -1 Then
        txtPosition.Text = SignName(cboApprove.ListIndex + 1).SignPosition1
        myQuotation_.SignName = cboApprove.Text
        myQuotation_.SignPosition = Trim(txtPosition)
        myQuotation_.isChanged = True
    End If
End Sub

Private Sub cboApprove_Click()
    txtPosition.Text = SignName(cboApprove.ListIndex + 1).SignPosition1
'    If Trim(cboApprove.Text) = SystemConfig.SignName1 Then
'        txtPosition.Text = SystemConfig.SignPosition1
'    ElseIf Trim(cboApprove.Text) = SystemConfig.SignName2 Then
'        txtPosition.Text = SystemConfig.SignPosition2
'    ElseIf Trim(cboApprove.Text) = SystemConfig.SignName3 Then
'        txtPosition.Text = SystemConfig.SignPosition3
'    ElseIf Trim(cboApprove.Text) = SystemConfig.SignName4 Then
'        txtPosition.Text = SystemConfig.SignPosition4
'    ElseIf Trim(cboApprove.Text) = SystemConfig.SignName5 Then
'        txtPosition.Text = SystemConfig.SignPosition5
'    End If
End Sub

Private Sub cboApprove_Validate(Cancel As Boolean)
    myQuotation_.SignName = cboApprove.Text
    myQuotation_.SignPosition = Trim(txtPosition)
    myQuotation_.isChanged = True
End Sub

Private Sub cboCompany_Validate(Cancel As Boolean)
    If cboCompany.ListIndex >= 0 Then
        If cboCompany.ListCount > 1 Then
            If cboCompany.ListIndex = 0 Then Exit Sub
            myQuotation_.CompanyID = CurrentUser.WorkCompanies(cboCompany.ListIndex).id
        Else
            myQuotation_.CompanyID = CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).id
        End If
    Else
        myQuotation_.CompanyID = ""
    End If
End Sub

Private Sub cboCustStaff_Validate(Cancel As Boolean)
    myQuotation_.AttentionTo = Trim(cboCustStaff.Text)
End Sub

Private Sub cboDiscountType_Click()
    If Me.Visible Then
        If Me.ActiveControl.Name = cboDiscountType.Name Then
            If cboDiscountType.ListIndex = 0 Then
                myQuotation_.DiscountType = "P"
            ElseIf cboDiscountType.ListIndex = 1 Then
                myQuotation_.DiscountType = "B"
            End If
        End If
    End If
End Sub

Private Sub cboIsTrailer_Validate(Cancel As Boolean)
    If cboIsTrailer.ListIndex = 0 Then
        myQuotation_.isTrailer = False
    Else
        myQuotation_.isTrailer = True
    End If
End Sub

Private Sub cboPrintForm_Validate(Cancel As Boolean)
    ValidateData
End Sub

Private Sub cboQuotationType_Validate(Cancel As Boolean)
    If cboQuotationType.ListIndex = 0 Then
        myQuotation_.isJob = "N"
        myQuotation_.isCase = ""
    ElseIf cboQuotationType.ListIndex = 1 Then
        myQuotation_.isJob = "Y"
        myQuotation_.isCase = ""
    ElseIf cboQuotationType.ListIndex = 2 Then
        myQuotation_.isJob = "D"
'        myQuotation_.isCase = WasteCases(cboQuotationType.ListIndex - 1).WasteCaseID
    End If
    If Trim(myQuotation_.isJob) = "" Then myQuotation_.isJob = "N"
End Sub

Private Sub cboTransUnit_Validate(Cancel As Boolean)
    If cboTransUnit.ListIndex > -1 Then
        myQuotation_.TransferUnitID = cboTransUnit.Text
        myQuotation_.TransferUnitName = cboTransUnit.Text
        myQuotation_.Truck1Unit = cboTransUnit.Text
    Else
        myQuotation_.TransferUnitID = ""
        myQuotation_.TransferUnitName = ""
        myQuotation_.Truck1Unit = ""
    End If
End Sub
Private Sub cboTransUnit2_Validate(Cancel As Boolean)
    If cboTransUnit2.ListIndex > -1 Then
        myQuotation_.Truck2Unit = cboTransUnit2.Text
    Else
        myQuotation_.Truck2Unit = ""
    End If
End Sub
Private Sub cboTransUnit3_Validate(Cancel As Boolean)
    If cboTransUnit3.ListIndex > -1 Then
        myQuotation_.TRuck3Unit = cboTransUnit3.Text
    Else
        myQuotation_.TRuck3Unit = ""
    End If
End Sub
Private Sub cboTransUnit4_Validate(Cancel As Boolean)
    If cboTransUnit4.ListIndex > -1 Then
        myQuotation_.TRuck4Unit = cboTransUnit4.Text
    Else
        myQuotation_.TRuck4Unit = ""
    End If
End Sub
Private Sub cboTransUnit5_Validate(Cancel As Boolean)
    If cboTransUnit5.ListIndex > -1 Then
        myQuotation_.TRuck5Unit = cboTransUnit5.Text
    Else
        myQuotation_.TRuck5Unit = ""
    End If
End Sub
Private Sub cboTransUnit6_Validate(Cancel As Boolean)
    If cboTransUnit6.ListIndex > -1 Then
        myQuotation_.TRuck6Unit = cboTransUnit6.Text
    Else
        myQuotation_.TRuck6Unit = ""
    End If
End Sub
Private Sub cboTransUnit7_Validate(Cancel As Boolean)
    If cboTransUnit7.ListIndex > -1 Then
        myQuotation_.TRuck7Unit = cboTransUnit7.Text
    Else
        myQuotation_.TRuck7Unit = ""
    End If
End Sub
Private Sub cboTransUnit8_Validate(Cancel As Boolean)
    If cboTransUnit8.ListIndex > -1 Then
        myQuotation_.TRuck8Unit = cboTransUnit8.Text
    Else
        myQuotation_.TRuck8Unit = ""
    End If
End Sub
Private Sub cboTransUnit9_Validate(Cancel As Boolean)
    If cboTransUnit9.ListIndex > -1 Then
        myQuotation_.TRuck9Unit = cboTransUnit9.Text
    Else
        myQuotation_.TRuck9Unit = ""
    End If
End Sub

Private Sub cbProject_Validate(Cancel As Boolean)
Dim ProID
Dim id As String
    If cbProject.Text <> "" Then
        ProID = Split(cbProject.Text, "#", -1, vbTextCompare)
        id = ProID(1)
    End If
    myQuotation_.ProjectID = id
End Sub

Private Sub chkBmeNew_Click()
    If chkBmeNew.Value = 1 Then
        myQuotation_.isBmeNew = "Y"
        Else
         myQuotation_.isBmeNew = "N"
    End If
    myQuotation_.isChanged = True
End Sub

Private Sub chkcarbt_Click()
    If chkcarbt.Value = 1 Then
        myQuotation_.isBusBT = "Y"
        Else
         myQuotation_.isBusBT = "N"
    End If
    myQuotation_.isChanged = True
End Sub

Private Sub ChkFreeTrn_Click()
         If ChkFreeTrn.Value = 1 Then
            ChkHalfTrn.Value = 0
            myQuotation_.TransRate = "F"
        End If
End Sub
Private Sub ChkHalfTrn_Click()
        If ChkHalfTrn.Value = 1 Then
            ChkFreeTrn.Value = 0
            myQuotation_.TransRate = "H"
        End If
End Sub

Private Sub chkIncTrans_Click()
    If chkIncTrans.Value = 1 Then
        txtTransRate.Enabled = False
        cboTransUnit.Enabled = False
        Frame4.Visible = True
    Else
        txtTransRate.Enabled = True
        cboTransUnit.Enabled = True
        Frame4.Visible = False
    End If
End Sub

Private Sub chkIncTrans_Validate(Cancel As Boolean)
    myQuotation_.IsPriceIncTransport = (chkIncTrans.Value = 1)
    If myQuotation_.IsPriceIncTransport Then
        txtTransRate.Text = ""
        cboTransUnit.ListIndex = -1
'        txtTransRate.Enabled = False
'        cboTransUnit.Enabled = False
        myQuotation_.TransferRate = 0
        myQuotation_.TransferUnitID = ""
        myQuotation_.TransferUnitName = ""
'    Else
'        txtTransRate.Enabled = True
'        cboTransUnit.Enabled = True
    End If
End Sub

Private Sub chkIncVAT_Validate(Cancel As Boolean)
    myQuotation_.IsPriceIncVAT = (chkIncVAT.Value = 1)
End Sub

Private Sub chkTrnDis_Click()
myQuotation_.IsDiscountPriceTrans = (chkTrnDis.Value = 1)
End Sub

Private Sub cmdAddItem_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole(frmWasteDataAdd.Name, AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    If Trim(txtCustomer.Text) = "" Then
        MsgBox "กรุณาเลือกลูกค้าก่อนทำการเพิ่มรายการกากของเสีย", vbExclamation
        txtCustomer.SetFocus
        Exit Sub
    End If
    isDetailChange_ = False
    Load frmWasteDataAdd
    With frmWasteDataAdd
        .SetCallerForm Me, curCustomer_
        .Show 1
        If isDetailChange_ Then
            ShowQuoDetails
        End If
    End With
End Sub

Private Sub cmdApproved_Click()
Dim result
Dim selResult As Integer
Dim tmpStr As String
Dim msg As String
    tmpStr = CurrentUser.VerifyAccessRole(Me.Name, ApproveRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
If curCustomer_.CustomerType <> "R" Then
    If Trim(txtCustomer.Text) = "" Then
        MsgBox "กรุณาใส่ชื่อลูกค้า", vbExclamation
        Exit Sub
    End If
    If Trim(txtMinQty.Text) <> "0" Then
        If Trim(txtMinQty.Text) <> "" Then
        If myQuotation_.CompanyID <> "776EDA5830F744A6B3551348D851BC22" Then
                If Trim(txtNonRate.Text) = "" And Trim(txtNBLRate.Text) = "" And Trim(txtHazRate.Text) = "" And Trim(txtHBLRate.Text) = "" Then
                    MsgBox "กรุณาระบุ ราคาขั้นต่ำ Non,Haz,HBL,NBL", vbExclamation
                    Exit Sub
            End If
            End If
        End If
    End If
    
    If myQuotation_.QuotationDetails Is Nothing Then
        MsgBox "กรุณาระบุรายการสินค้าที่ต้องการสั่งซื้อ", vbExclamation
        Exit Sub
    End If
    If myQuotation_.QuotationDetails.Count = 0 Then
        MsgBox "กรุณาระบุรายการสินค้าที่ต้องการสั่งซื้อ", vbExclamation
        Exit Sub
    End If

'        If myQuotation_.isQuotationReq = True Then
'            If myQuotation_.isWasteApproveAll = False Then
'                MsgBox "ใบเสนอราคานี้เป็นใบเสนอราคาที่ออกแบบเร่งด่วน ยังไม่ผ่านการตรวจรับจากฝ่ายสิ่งแวดล้อม" & vbCrLf & "ระบบจะส่งข้อมูลให้ฝ่ายสิ่งแวดล้อมตรวจสอบก่อนจึงสามารถอนุมัติได้", vbExclamation, "Warning"
'                msg = myQuotation_.WantApprove
'                If Trim(msg) <> "" Then
'                    MsgBox msg, vbCritical, "Warning"
'                End If
'                Exit Sub
'            End If
'        End If
    
    If myQuotation_.isChanged And cmdSave.Enabled Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนสร้างใหม่หรือไม่?", vbQuestion + vbYesNo)
        If selResult = vbYes Then
            CmdSave_Click
            If myQuotation_.isChanged Then Exit Sub
        Else 'If selResult = vbCancel Then
            Exit Sub
        End If
    End If
    
'    If myQuotation_.isWasteApproveAll = False Then
'        MsgBox "ใบเสนอราคาใบนี้มีบางเวสต์ที่ยังไม่ผ่านการตรวจรับจากฝ่ายสิ่งแวดล้อม", vbExclamation, "Warning"
'        If Trim(msg) <> "" Then
'            MsgBox msg, vbCritical, "Warning"
'        End If
'        Exit Sub
'    End If
    End If
    
    result = MsgBox("ต้องการอนุมัติใบเสนอราคานี้ใช่หรือไม่", vbQuestion + vbYesNo)
    If result = vbYes Then
        myQuotation_.isApproved = True
        myQuotation_.Approveddate = TodayDate
        myQuotation_.ApprovedStaffID = CurrentUser.id
        myQuotation_.ApprovedStaffName = CurrentUser.StaffFName & " " & CurrentUser.StaffLName
        myQuotation_.Approve
'        myQuotation_.Save
        Label2(4).Caption = "อนุมัติแล้วเมื่อวันที่ " & FormatYMD_to_DMY(myQuotation_.Approveddate, "/", "/") & " โดย " & myQuotation_.ApprovedStaffName
        txtApprovedStaff.Text = myQuotation_.ApprovedStaffName
        SetEnabledScreen
    
        If myQuotation_.isApproved Then
            lbCancel.Caption = "Approved"
            lbCancel.Visible = True
        Else
            lbCancel.Visible = False
        End If
    End If
End Sub

Private Sub cmdCancel_Click()
Dim result
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole(Me.Name, DeleteRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    If InStr(1, LCase(Trim(txtQuoNo.Text)), "auto") > 0 Or Trim(txtQuoNo.Text) = "" Then
        MsgBox "กรุณาระบุใบเสนอราคาที่ต้องการยกเลิก", vbExclamation
        Exit Sub
    End If
    If Trim(txtRem.Text) = "" Then
        MsgBox "กรุณาระบุสาเหตุในการยกเลิกในช่องหมายเหตุ", vbExclamation
        txtRem.SetFocus
        Exit Sub
    End If
    
    result = MsgBox("ต้องการยกเลิกใบสั่งซื้อใช่หรือไม่", vbQuestion + vbYesNo)
    If result = vbYes Then
        If CheckPendingJob Then
            If MsgBox("มี Job ที่ยังไม่ได้ออก Manifest ใช้งานใบเสนอราคานี้อยู่ ท่านต้องการย้าย Job เหล่านี้ไปใช้งานใบเสนอราคาอื่นหรือไม่", vbQuestion + vbYesNo) = vbYes Then
                Load frmQuotationTransfer2
                With frmQuotationTransfer2
                    .SetDataQuotation myQuotation_
                    .SetPendingJob colPendingJob_, txtQuoNo.Text
                    .Show 1
                End With
            
'                Load frmQuotationTransfer
'                With frmQuotationTransfer
'                    .SetPendingJob colPendingJob_, txtQuoNo.Text
'                    .Show 1
'                End With
                If CheckPendingJob Then
                    MsgBox "ยังมี Job ที่ยังไม่ได้ออก Manifest ใช้งานใบเสนอราคานี้อยู่ ไม่สามารถยกเลิกใบเสนอราคานี้ได้", vbExclamation
                    Exit Sub
                End If
            Else
                MsgBox "ไม่สามารถทำการยกเลิกใบเสนอราคาได้", vbExclamation
                Exit Sub
            End If
        End If
    
        myQuotation_.CancelQuotation
        If myQuotation_.isCancel Then
            lbCancel.Caption = "ยกเลิก"
            lbCancel.Visible = True
        Else
            lbCancel.Visible = False
        End If
        SetEnabledScreen
    End If
End Sub

Private Sub cmdClear_Click()
Dim selResult As Integer
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole(Me.Name, AddNewRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    
    If myQuotation_.isChanged And cmdSave.Enabled Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนสร้างใหม่หรือไม่?", vbQuestion + vbYesNoCancel)
        If selResult = vbYes Then
            CmdSave_Click
            If myQuotation_.isChanged Then Exit Sub
        ElseIf selResult = vbCancel Then
            Exit Sub
        End If
    End If
    ClearScreen
    CreateNewQuotation
    txtCustomer.SetFocus
    isSave = False
    cboPrintForm.Enabled = True
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdCloseBill_Click()
Dim result
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole(Me.Name, DeleteRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    If InStr(1, LCase(Trim(txtQuoNo.Text)), "auto") > 0 Or Trim(txtQuoNo.Text) = "" Then
        MsgBox "กรุณาระบุใบเสนอราคาที่ต้องการยกเลิก", vbExclamation
        Exit Sub
    End If
    
    If Trim(txtRem.Text) = "" Then
        MsgBox "กรุณาระบุสาเหตุในการหยุดใช้งานใบเสนอราคาในช่องหมายเหตุ", vbExclamation
        txtRem.SetFocus
        Exit Sub
    End If
    
    
    result = MsgBox("ต้องการหยุดใช้งานใบเสนอราคานี้หรือไม่", vbQuestion + vbYesNo)
    If result = vbYes Then
'        If CheckPendingJob Then
'            If MsgBox("มี Job ที่ยังไม่ได้ออก Manifest ใช้งานใบเสนอราคานี้อยู่ ท่านต้องการย้าย Job เหล่านี้ไปใช้งานใบเสนอราคาอื่นหรือไม่", vbQuestion + vbYesNo) = vbYes Then
'                Load frmQuotationTransfer2
'                With frmQuotationTransfer2
'                    .SetDataQuotation myQuotation_
'                    .SetPendingJob colPendingJob_, txtQuoNo.Text
'                    .Show 1
'                End With
'                If CheckPendingJob Then
'                    MsgBox "ยังมี Job ที่ยังไม่ได้ออก Manifest ใช้งานใบเสนอราคานี้อยู่ ไม่สามารถหยุดใช้งานใบเสนอราคานี้ได้", vbExclamation
'                    Exit Sub
'                End If
'            Else
'                MsgBox "ไม่สามารถทำการหยุดใช้งานใบเสนอราคาได้", vbExclamation
'                Exit Sub
'            End If
'        End If
        
        myQuotation_.CloseQuotation
        If myQuotation_.isClosed Then
            lbCancel.Caption = "Closed"
            lbCancel.Visible = True
        Else
            lbCancel.Visible = False
        End If
        SetEnabledScreen
    End If
End Sub
Private Sub PrintQuotationOld()
Dim tmpSql As String
Dim selResult
Dim tmpStr As String
Dim AcceptState As Boolean, i As Integer
    tmpStr = CurrentUser.VerifyAccessRole(Me.Name, PrintRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If

    If Trim(txtCustomer.Text) = "" Then
        MsgBox "กรุณาใส่ชื่อลูกค้า", vbExclamation
        Exit Sub
    End If
    If Trim(cboApprove.Text) = "" Then
        MsgBox "กรุณาระบุชื่อผู้อนุมัติ", vbExclamation
        cboApprove.SetFocus
        Exit Sub
    End If
    If Trim(txtPosition.Text) = "" Then
        MsgBox "กรุณาระบุตำแหน่งผู้อนุมัติ", vbExclamation
        txtPosition.SetFocus
        Exit Sub
    End If

    If myQuotation_.isChanged And cmdSave.Enabled Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนสร้างใหม่หรือไม่?", vbQuestion + vbYesNo)
        If selResult = vbYes Then
            CmdSave_Click
            If myQuotation_.isChanged Then Exit Sub
        Else 'If selResult = vbCancel Then
            Exit Sub
        End If
    End If

    With rpt1
        .WindowShowCloseBtn = True
        .WindowShowExportBtn = True
        .WindowShowPrintSetupBtn = True
        .LogonInfo(0) = "DSN = " & DSN_NAME & ";UID = " & strUid & ";PWD = " & strPwd
        If myQuotation_.CompanyID = "A2B9D8E8DD344FD9BF64250AF00A07E3" Then
            If myQuotation_.QuotationType = 2 Then
                .ReportFileName = App.Path & "\reports\QuotationAP_Et.rpt"
            Else
                .ReportFileName = App.Path & "\reports\Quotation_Et.rpt"
            End If
        ElseIf myQuotation_.CompanyID = "3F764F1713974948A42CC0FAF926C730" Then
            If myQuotation_.QuotationType = 2 Then
                .ReportFileName = App.Path & "\reports\QuotationAP_PST2.rpt"
            Else
               .ReportFileName = App.Path & "\reports\Quotation_PST.rpt"
            End If
        Else
            If myQuotation_.QuotationType = 2 Then
                If myQuotation_.CompanyID = "776EDA5830F744A6B3551348D851BC22" Then
                    If myQuotation_.isQuotationReq = False Then
                        .ReportFileName = App.Path & "\Reports\QuotationAKP.rpt"
                    Else
                        .ReportFileName = App.Path & "\Reports\QuotationAKP2.rpt"
                    End If
                Else
                    If myQuotation_.isQuotationReq = False Then
                        .ReportFileName = App.Path & "\Reports\QuotationAP.rpt"
                    Else
                        .ReportFileName = App.Path & "\Reports\QuotationAP2.rpt"
                    End If
                End If
            Else
                AcceptState = True
                For i = 1 To myQuotation_.QuotationDetails.Count
                    If myQuotation_.QuotationDetails(i).WasteStatus <> "A" Then
                        AcceptState = False
                        Exit For
                    End If
                Next
                If AcceptState = True Then
                    If myQuotation_.CompanyID = "776EDA5830F744A6B3551348D851BC22" Then
                        .ReportFileName = App.Path & "\Reports\QuotationAKP.rpt"
                    Else
                        .ReportFileName = App.Path & "\Reports\Quotation.rpt"
                    End If
                Else
                    If myQuotation_.CompanyID = "776EDA5830F744A6B3551348D851BC22" Then
                        If myQuotation_.isQuotationReq = False Then
                            .ReportFileName = App.Path & "\Reports\QuotationAKP.rpt"
                        Else
                            .ReportFileName = App.Path & "\Reports\QuotationAKP2.rpt"
                        End If
                    Else
                        If myQuotation_.isQuotationReq = False Then
                            .ReportFileName = App.Path & "\Reports\Quotation.rpt"
                        Else
                            .ReportFileName = App.Path & "\Reports\Quotation2.rpt"
                        End If
                    End If
                End If
            End If
        End If
        If cboQuotationType.ListIndex = 2 Then
            If chkIncVAT.Value = 1 Then
            .ReportFileName = App.Path & "\Reports\QuotationAKP3.rpt"
            Else
            .ReportFileName = App.Path & "\Reports\QuotationAKP4.rpt"
            End If
        End If
        .Formulas(0) = "ShowPODate=" & Chr(34) & dtIssuedDate.ValueDMY & Chr(34)
        If cboQuotationType.ListIndex = 2 Then
        .Formulas(1) = "CustName=" & Chr(34) & txtCustomer.Text & "  " & "(" & curCustomer_.CustomerCode & ")" & Chr(34)
        Else
        .Formulas(1) = "CustName=" & Chr(34) & txtCustomer.Text & Chr(34)
        End If
        .Formulas(2) = "Addr1=" & Chr(34) & curCustomer_.Address & " " & curCustomer_.TumbolName & " " & curCustomer_.AumphurName & " " & curCustomer_.ProvinceName & Chr(34)
        .Formulas(3) = "TelNo=" & Chr(34) & curCustomer_.TelNo & Chr(34)
        .Formulas(4) = "FaxNo=" & Chr(34) & curCustomer_.FaxNo & Chr(34)
        If chkIncTrans.Value = 0 Then
            .Formulas(5) = "IsIncTransFee=" & Chr(34) & "ราคายังไม่รวมค่าขนส่ง" & Chr(34)
        Else
            .Formulas(5) = "IsIncTransFee=" & Chr(34) & "ฟรีค่าขนส่ง" & Chr(34)
        End If
        If curCustomer_.CustomerType = "M" Or curCustomer_.CustomerType = "MS" Then
        .Formulas(6) = "AttentionTo=" & Chr(34) & "คุณจารุวรรณ โพธิ์แจ้ง" & Chr(34)
        Else
        .Formulas(6) = "AttentionTo=" & Chr(34) & cboCustStaff.Text & Chr(34)
        End If
        .Formulas(7) = "QuotationNo=" & Chr(34) & Replace(txtQuoNo.Text, "(CR)", "") & Chr(34)
        .Formulas(8) = "ApprovalName=" & Chr(34) & "(  " & cboApprove.Text & "  )" & Chr(34)
        .Formulas(9) = "ApprovePosition=" & Chr(34) & txtPosition.Text & Chr(34)
        .Formulas(11) = "PDay=" & Chr(34) & lblD.Caption & Chr(34)
        .Formulas(12) = "PMonth=" & Chr(34) & lblM.Caption & Chr(34)
        .Formulas(13) = "PYear=" & Chr(34) & lblY.Caption & Chr(34)
'        For i = 1 To myQuotation_.QuotationCar.Count
'            If i <= 9 Then
'             .Formulas(14 + i) = "truckDesc" & i & "=" & Chr(34) & i & ". " & myQuotation_.QuotationCar(i).TypeDesc & Chr(34)
'            .Formulas((14 + 9) + i) = "truckRate" & i & "=" & Chr(34) & Format(myQuotation_.QuotationCar(i).Rate, "#,##0.00") & Chr(34)
'            End If
'
'        Next i
        
        If myQuotation_.CompanyID = "776EDA5830F744A6B3551348D851BC22" Then
            If myQuotation_.SaleAP <> "" Then
                .Formulas(10) = "SaleName=" & Chr(34) & curCustomer_.UnderSaleName & " (" & myQuotation_.SaleAP & ") โทร. " & curCustomer_.UnderSaleTelNo & Chr(34)
            Else
                .Formulas(10) = "SaleName=" & Chr(34) & curCustomer_.UnderSaleName & " โทร. " & curCustomer_.UnderSaleTelNo & Chr(34)
            End If
        Else
            .Formulas(10) = "SaleName=" & Chr(34) & curCustomer_.UnderSaleName & " โทร. " & curCustomer_.UnderSaleTelNo & Chr(34)
        End If
        tmpSql = ""
        If myQuotation_.CompanyID = "A2B9D8E8DD344FD9BF64250AF00A07E3" Then
            If myQuotation_.QuotationType = 2 Then
                If myQuotation_.isQuotationReq = False Then
                    tmpSql = "SELECT * FROM Com_vwQuotationDetailSearch2" & vbCrLf
                Else
                    tmpSql = "SELECT * FROM Com_vwQuotationDetailSearch2_New  AS Com_vwQuotationDetailSearch2" & vbCrLf
                End If
            Else
                If myQuotation_.isQuotationReq = False Then
                    tmpSql = "SELECT * FROM Com_vwQuotationDetailSearch " & vbCrLf
                Else
                    tmpSql = "SELECT * FROM Com_vwQuotationDetailSearch2_New  AS Com_vwQuotationDetailSearch" & vbCrLf
                End If
            End If
        ElseIf myQuotation_.CompanyID = "3F764F1713974948A42CC0FAF926C730" Then
            If myQuotation_.QuotationType = 2 Then
                If myQuotation_.isQuotationReq = False Then
                    tmpSql = "SELECT * FROM Com_vwQuotationDetailSearch2" & vbCrLf
                Else
                    tmpSql = "SELECT * FROM Com_vwQuotationDetailSearch2_New  AS Com_vwQuotationDetailSearch2" & vbCrLf
                End If
            Else
                If myQuotation_.isQuotationReq = False Then
                    tmpSql = "SELECT * FROM Com_vwQuotationDetailSearch " & vbCrLf
                Else
                    tmpSql = "SELECT * FROM Com_vwQuotationDetailSearch2_New  AS Com_vwQuotationDetailSearch" & vbCrLf
                End If
            End If
        Else
            If myQuotation_.QuotationType = 2 Then
                If myQuotation_.CompanyID = "776EDA5830F744A6B3551348D851BC22" Then
                    If myQuotation_.isQuotationReq = False Then
                        tmpSql = "SELECT * FROM Com_vwQuotationDetailSearch " & vbCrLf
                    Else
                        tmpSql = "SELECT * FROM Com_vwQuotationDetailSearch_New " & vbCrLf
                    End If
                Else
                    If myQuotation_.isQuotationReq = False Then
                        tmpSql = "SELECT * FROM Com_vwQuotationDetailSearch2 " & vbCrLf
                    Else
                        tmpSql = "SELECT * FROM Com_vwQuotationDetailSearch2_New as Com_vwQuotationDetailSearch2 " & vbCrLf
                    End If
                End If
            Else
                If AcceptState = True Then
                    tmpSql = "SELECT * FROM Com_vwQuotationDetailSearch " & vbCrLf
                Else
                    If myQuotation_.isQuotationReq = False Then
                        tmpSql = "SELECT * FROM Com_vwQuotationDetailSearch " & vbCrLf
                    Else
                    If .ReportFileName = App.Path & "\Reports\QuotationAKP2.rpt" Then
                        tmpSql = "SELECT * FROM Com_vwQuotationDetailSearch_New " & vbCrLf
                    Else
                        tmpSql = "SELECT * FROM Com_vwQuotationDetailSearch_New  as Com_vwQuotationDetailSearch" & vbCrLf
                    End If
                    End If
                End If
            End If
        End If
        tmpSql = tmpSql & "WHERE QuotationID='" & myQuotation_.id & "' " & vbCrLf
        tmpSql = tmpSql & "ORDER BY RowPos"
        .SQLQuery = tmpSql
        .Destination = crptToWindow
        .PrintReport
        If Trim(.LastErrorString) <> "" Then MsgBox .LastErrorString
    End With
End Sub
Private Sub PrintQuotationPdfOld()
On Error GoTo ErrD
Dim tmpSql As String
Dim selResult
Dim tmpStr As String
Dim AcceptState As Boolean, i As Integer

    tmpStr = CurrentUser.VerifyAccessRole(Me.Name, PrintRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If

    If Trim(txtCustomer.Text) = "" Then
        MsgBox "กรุณาใส่ชื่อลูกค้า", vbExclamation
        Exit Sub
    End If
    If Trim(cboApprove.Text) = "" Then
        MsgBox "กรุณาระบุชื่อผู้อนุมัติ", vbExclamation
        cboApprove.SetFocus
        Exit Sub
    End If
    If Trim(txtPosition.Text) = "" Then
        MsgBox "กรุณาระบุตำแหน่งผู้อนุมัติ", vbExclamation
        txtPosition.SetFocus
        Exit Sub
    End If

    If myQuotation_.isChanged And cmdSave.Enabled Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนสร้างใหม่หรือไม่?", vbQuestion + vbYesNo)
        If selResult = vbYes Then
            CmdSave_Click
            If myQuotation_.isChanged Then Exit Sub
        Else 'If selResult = vbCancel Then
            Exit Sub
        End If
    End If

    With rpt1
        .WindowShowCloseBtn = True
        .WindowShowExportBtn = True
        .WindowShowPrintSetupBtn = True
        .LogonInfo(0) = "DSN = " & DSN_NAME & ";UID = " & strUid & ";PWD = " & strPwd
        If myQuotation_.CompanyID = "A2B9D8E8DD344FD9BF64250AF00A07E3" Then
            If myQuotation_.QuotationType = 2 Then
                .ReportFileName = App.Path & "\reports\QuotationAP_Et.rpt"
            Else
                .ReportFileName = App.Path & "\reports\Quotation_Et.rpt"
            End If
        Else
            If myQuotation_.QuotationType = 2 Then
                If myQuotation_.CompanyID = "776EDA5830F744A6B3551348D851BC22" Then
                    If myQuotation_.isQuotationReq = False Then
                        .ReportFileName = App.Path & "\Reports\QuotationAKP.rpt"
                    Else
                        .ReportFileName = App.Path & "\Reports\QuotationAKP2.rpt"
                    End If
                Else
                    If myQuotation_.isQuotationReq = False Then
                        .ReportFileName = App.Path & "\Reports\QuotationAP.rpt"
                    Else
                        .ReportFileName = App.Path & "\Reports\QuotationAP2.rpt"
                    End If
                End If
            Else
                AcceptState = True
                For i = 1 To myQuotation_.QuotationDetails.Count
                    If myQuotation_.QuotationDetails(i).WasteStatus <> "A" Then
                        AcceptState = False
                        Exit For
                    End If
                Next
                If AcceptState = True Then
                    If myQuotation_.CompanyID = "776EDA5830F744A6B3551348D851BC22" Then
                        .ReportFileName = App.Path & "\Reports\QuotationAKP.rpt"
                    Else
                        .ReportFileName = App.Path & "\Reports\QuotationBWGForm1.rpt"
                    End If
                Else
                    If myQuotation_.CompanyID = "776EDA5830F744A6B3551348D851BC22" Then
                        If myQuotation_.isQuotationReq = False Then
                            .ReportFileName = App.Path & "\Reports\QuotationAKP.rpt"
                        Else
                            .ReportFileName = App.Path & "\Reports\QuotationAKP2.rpt"
                        End If
                    Else
                        If myQuotation_.isQuotationReq = False Then
                            .ReportFileName = App.Path & "\Reports\QuotationBWGForm1.Rpt"
                        Else
                            .ReportFileName = App.Path & "\Reports\QuotationBWGForm2.Rpt"
                        End If
                    End If
                End If

            End If
        End If
        .Formulas(0) = "ShowPODate=" & Chr(34) & dtIssuedDate.ValueDMY & Chr(34)
        .Formulas(1) = "CustName=" & Chr(34) & txtCustomer.Text & Chr(34)
        .Formulas(2) = "Addr1=" & Chr(34) & curCustomer_.Address & " " & curCustomer_.TumbolName & " " & curCustomer_.AumphurName & " " & curCustomer_.ProvinceName & Chr(34)
        .Formulas(3) = "TelNo=" & Chr(34) & curCustomer_.TelNo & Chr(34)
        .Formulas(4) = "FaxNo=" & Chr(34) & curCustomer_.FaxNo & Chr(34)
        If chkIncTrans.Value = 0 Then
            .Formulas(5) = "IsIncTransFee=" & Chr(34) & "ราคายังไม่รวมค่าขนส่ง" & Chr(34)
        Else
            .Formulas(5) = "IsIncTransFee=" & Chr(34) & "ฟรีค่าขนส่ง" & Chr(34)
        End If
        .Formulas(6) = "AttentionTo=" & Chr(34) & cboCustStaff.Text & Chr(34)
        .Formulas(7) = "QuotationNo=" & Chr(34) & Replace(txtQuoNo.Text, "(CR)", "") & Chr(34)
        .Formulas(8) = "ApprovalName=" & Chr(34) & "(  " & cboApprove.Text & "  )" & Chr(34)
        .Formulas(9) = "ApprovePosition=" & Chr(34) & txtPosition.Text & Chr(34)
        If myQuotation_.CompanyID = "776EDA5830F744A6B3551348D851BC22" Then
            If myQuotation_.SaleAP <> "" Then
                .Formulas(10) = "SaleName=" & Chr(34) & curCustomer_.UnderSaleName & " (" & myQuotation_.SaleAP & ") โทร. " & curCustomer_.UnderSaleTelNo & Chr(34)
            Else
                .Formulas(10) = "SaleName=" & Chr(34) & curCustomer_.UnderSaleName & " โทร. " & curCustomer_.UnderSaleTelNo & Chr(34)
            End If
        Else
            .Formulas(10) = "SaleName=" & Chr(34) & curCustomer_.UnderSaleName & " โทร. " & curCustomer_.UnderSaleTelNo & Chr(34)
        End If

        tmpSql = ""
        If myQuotation_.CompanyID = "A2B9D8E8DD344FD9BF64250AF00A07E3" Then
            If myQuotation_.QuotationType = 2 Then
                If myQuotation_.isQuotationReq = False Then
                    tmpSql = "SELECT * FROM Com_vwQuotationDetailSearch2" & vbCrLf
                Else
                    tmpSql = "SELECT * FROM Com_vwQuotationDetailSearch2_New  AS Com_vwQuotationDetailSearch2" & vbCrLf
                End If
            Else
                If myQuotation_.isQuotationReq = False Then
                    tmpSql = "SELECT * FROM Com_vwQuotationDetailSearch " & vbCrLf
                Else
                    tmpSql = "SELECT * FROM Com_vwQuotationDetailSearch2_New  AS Com_vwQuotationDetailSearch" & vbCrLf
                End If
            End If
        Else
            If myQuotation_.QuotationType = 2 Then
                If myQuotation_.CompanyID = "776EDA5830F744A6B3551348D851BC22" Then
                    If myQuotation_.isQuotationReq = False Then
                        tmpSql = "SELECT * FROM Com_vwQuotationDetailSearch " & vbCrLf
                    Else
                        tmpSql = "SELECT * FROM Com_vwQuotationDetailSearch_New " & vbCrLf
                    End If
                Else
                    If myQuotation_.isQuotationReq = False Then
                        tmpSql = "SELECT * FROM Com_vwQuotationDetailSearch2 " & vbCrLf
                    Else
                        tmpSql = "SELECT * FROM Com_vwQuotationDetailSearch2_New Com_vwQuotationDetailSearch2 " & vbCrLf
                    End If
                End If
            Else
                If AcceptState = True Then
                    tmpSql = "SELECT * FROM Com_vwQuotationDetailSearch " & vbCrLf
                Else
                    If myQuotation_.isQuotationReq = False Then
                        tmpSql = "SELECT * FROM Com_vwQuotationDetailSearch " & vbCrLf
                    Else
                        tmpSql = "SELECT * FROM Com_vwQuotationDetailSearch_New " & vbCrLf
                    End If
                End If
            End If
        End If
        tmpSql = tmpSql & "WHERE QuotationID='" & myQuotation_.id & "' " & vbCrLf
        tmpSql = tmpSql & "ORDER BY RowPos"
        .SQLQuery = tmpSql
'        .Destination = crptToWindow
'        .PrintReport
        If Trim(.LastErrorString) <> "" Then MsgBox .LastErrorString
    End With
    
    Dim myCRApp As New CRAXDRT.Application
    Dim myCRReport As CRAXDRT.Report
    Dim adoCmd As ADODB.Command, rs As New ADODB.Recordset
    Dim crxTable As CRAXDRT.DatabaseTable
    Dim AppPath As Scripting.FileSystemObject
    Dim PDFPath As String
    
    GetDBConnection
    Set adoCmd = New ADODB.Command
    Set adoCmd.ActiveConnection = DBConn
    adoCmd.CommandText = tmpSql
    adoCmd.CommandType = adCmdText

    Set rs = GetRS(tmpSql, adOpenKeyset, adLockOptimistic)
    GetDBConnection
    
    Set myCRApp = New CRAXDRT.Application
    Set myCRReport = myCRApp.OpenReport(rpt1.ReportFileName)
'    myCRApp.NewReport
'    myCRApp.LogOnServer "sqloledb.dll", dbServer, dbVolumn, uID, uPwd
    myCRReport.DiscardSavedData
    For Each crxTable In myCRReport.Database.Tables
        crxTable.SetLogOnInfo DSN_NAME, strDb, strUid, strPwd
'        crxTable.SetDataSource tmpSql
'        crxTable.SetDataSource RS
    Next
'    myCRReport.Database.AddADOCommand DBConn, adoCmd
    myCRReport.SQLQueryString = tmpSql
    Set AppPath = New Scripting.FileSystemObject
    PDFPath = "C:\BWG_PDFExport\"
    If AppPath.FolderExists(PDFPath) = False Then
        AppPath.CreateFolder PDFPath
    End If
    
    myCRReport.FormulaFields.GetItemByName("ShowPODate").Text = Chr(34) & dtIssuedDate.ValueDMY & Chr(34)
    myCRReport.FormulaFields.GetItemByName("CustName").Text = Chr(34) & txtCustomer.Text & Chr(34)
    myCRReport.FormulaFields.GetItemByName("Addr1").Text = Chr(34) & curCustomer_.Address & " " & curCustomer_.TumbolName & " " & curCustomer_.AumphurName & " " & curCustomer_.ProvinceName & Chr(34)
    myCRReport.FormulaFields.GetItemByName("TelNo").Text = Chr(34) & curCustomer_.TelNo & Chr(34)
    myCRReport.FormulaFields.GetItemByName("FaxNo").Text = Chr(34) & curCustomer_.FaxNo & Chr(34)
    If chkIncTrans.Value = 0 Then
        myCRReport.FormulaFields.GetItemByName("IsIncTransFee").Text = Chr(34) & "ราคายังไม่รวมค่าขนส่ง" & Chr(34)
    Else
        myCRReport.FormulaFields.GetItemByName("IsIncTransFee").Text = Chr(34) & "ฟรีค่าขนส่ง" & Chr(34)
    End If
    myCRReport.FormulaFields.GetItemByName("AttentionTo").Text = Chr(34) & cboCustStaff.Text & Chr(34)
    myCRReport.FormulaFields.GetItemByName("QuotationNo").Text = Chr(34) & txtQuoNo.Text & Chr(34)
    myCRReport.FormulaFields.GetItemByName("ApprovalName").Text = Chr(34) & "(  " & cboApprove.Text & "  )" & Chr(34)
    myCRReport.FormulaFields.GetItemByName("ApprovePosition").Text = Chr(34) & txtPosition.Text & Chr(34)
    If myQuotation_.CompanyID = "776EDA5830F744A6B3551348D851BC22" Then
        If myQuotation_.SaleAP <> "" Then
            myCRReport.FormulaFields.GetItemByName("SaleName").Text = Chr(34) & curCustomer_.UnderSaleName & " (" & myQuotation_.SaleAP & ") โทร. " & curCustomer_.UnderSaleTelNo & Chr(34)
        Else
            myCRReport.FormulaFields.GetItemByName("SaleName").Text = Chr(34) & curCustomer_.UnderSaleName & " โทร. " & curCustomer_.UnderSaleTelNo & Chr(34)
        End If
    Else
        myCRReport.FormulaFields.GetItemByName("SaleName").Text = Chr(34) & curCustomer_.UnderSaleName & " โทร. " & curCustomer_.UnderSaleTelNo & Chr(34)
    End If
        
    myCRReport.ExportOptions.FormatType = crEFTPortableDocFormat
    myCRReport.ExportOptions.DestinationType = crEDTApplication
    myCRReport.ExportOptions.DiskFileName = PDFPath & myQuotation_.QuotationNo & ".pdf"
    myCRReport.ExportOptions.ApplicationFileName = PDFPath & myQuotation_.QuotationNo & ".pdf"
    myCRReport.Export False
    Set myCRApp = Nothing
    Set myCRReport = Nothing
    CloseDBConnection
    Exit Sub
ErrD:
    MsgBox err.Description
    Exit Sub
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

    If Trim(txtCustomer.Text) = "" Then
        MsgBox "กรุณาใส่ชื่อลูกค้า", vbExclamation
        Exit Sub
    End If
    If Trim(cboApprove.Text) = "" Then
        MsgBox "กรุณาระบุชื่อผู้อนุมัติ", vbExclamation
        cboApprove.SetFocus
        Exit Sub
    End If
    If Trim(txtPosition.Text) = "" Then
        MsgBox "กรุณาระบุตำแหน่งผู้อนุมัติ", vbExclamation
        txtPosition.SetFocus
        Exit Sub
    End If

    If myQuotation_.isChanged And cmdSave.Enabled Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนสร้างใหม่หรือไม่?", vbQuestion + vbYesNo)
        If selResult = vbYes Then
            CmdSave_Click
            If myQuotation_.isChanged Then Exit Sub
        Else 'If selResult = vbCancel Then
            Exit Sub
        End If
    End If

    With rpt1
        .WindowShowCloseBtn = True
        .WindowShowExportBtn = True
        .WindowShowPrintSetupBtn = True
        .LogonInfo(0) = "DSN = " & DSN_NAME & ";UID = " & strUid & ";PWD = " & strPwd
       .ReportFileName = App.Path & "\Reports\QuotationAKPNew.rpt"
        .SelectionFormula = "{vw_QuotationHeaderSearch.QuotationID} = '" & myQuotation_.id & "'"
        '.SQLQuery = tmpSql
        
        .Destination = crptToWindow
        .PrintReport
        If Trim(.LastErrorString) <> "" Then MsgBox .LastErrorString
    End With
End Sub

Private Sub cmdPrintCar_Click()
Dim tmpSql As String
Dim selResult
Dim tmpStr As String
Dim AcceptState As Boolean, i As Integer
    tmpStr = CurrentUser.VerifyAccessRole(Me.Name, PrintRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If

    If Trim(txtCustomer.Text) = "" Then
        MsgBox "กรุณาใส่ชื่อลูกค้า", vbExclamation
        Exit Sub
    End If
'    If Trim(cboApprove.Text) = "" Then
'        MsgBox "กรุณาระบุชื่อผู้อนุมัติ", vbExclamation
'        cboApprove.SetFocus
'        Exit Sub
'    End If
'    If Trim(txtPosition.Text) = "" Then
'        MsgBox "กรุณาระบุตำแหน่งผู้อนุมัติ", vbExclamation
'        txtPosition.SetFocus
'        Exit Sub
'    End If

    If myQuotation_.isChanged And cmdSave.Enabled Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนสร้างใหม่หรือไม่?", vbQuestion + vbYesNo)
        If selResult = vbYes Then
            CmdSave_Click
            If myQuotation_.isChanged Then Exit Sub
        Else 'If selResult = vbCancel Then
            Exit Sub
        End If
    End If

    With rpt1
        .WindowShowCloseBtn = True
        .WindowShowExportBtn = True
        .WindowShowPrintSetupBtn = True
        .LogonInfo(0) = "DSN = " & DSN_NAME & ";UID = " & strUid & ";PWD = " & strPwd
        If myQuotation_.CompanyID = "A2B9D8E8DD344FD9BF64250AF00A07E3" Then
            If myQuotation_.QuotationType = 2 Then
                .ReportFileName = App.Path & "\reports\QuotationAP_Et.rpt"
            Else
                .ReportFileName = App.Path & "\reports\Quotation_Et.rpt"
            End If
        ElseIf myQuotation_.CompanyID = "3F764F1713974948A42CC0FAF926C730" Then
            If myQuotation_.QuotationType = 2 Then
                .ReportFileName = App.Path & "\reports\QuotationAP_PST2.rpt"
            Else
               .ReportFileName = App.Path & "\reports\Quotation_PST.rpt"
            End If
        Else
            If myQuotation_.QuotationType = 2 Then
                If myQuotation_.CompanyID = "776EDA5830F744A6B3551348D851BC22" Then
                    If myQuotation_.isQuotationReq = False Then
                        .ReportFileName = App.Path & "\Reports\QuotationAKPTruck.rpt"
                    Else
                        .ReportFileName = App.Path & "\Reports\QuotationAKPTruck.rpt"
                    End If
                Else
                    If myQuotation_.isQuotationReq = False Then
                        .ReportFileName = App.Path & "\Reports\QuotationAKPTruck.rpt"
                    Else
                        .ReportFileName = App.Path & "\Reports\QuotationAKPTruck.rpt"
                    End If
                End If
            Else
                AcceptState = True
                For i = 1 To myQuotation_.QuotationDetails.Count
                    If myQuotation_.QuotationDetails(i).WasteStatus <> "A" Then
                        AcceptState = False
                        Exit For
                    End If
                Next
                If AcceptState = True Then
                    If myQuotation_.CompanyID = "776EDA5830F744A6B3551348D851BC22" Then
                        .ReportFileName = App.Path & "\Reports\QuotationAKPTruck.rpt"
                    Else
                        .ReportFileName = App.Path & "\Reports\QuotationAKPTruck.rpt"
                    End If
                Else
                    If myQuotation_.CompanyID = "776EDA5830F744A6B3551348D851BC22" Then
                        If myQuotation_.isQuotationReq = False Then
                            .ReportFileName = App.Path & "\Reports\QuotationAKPTruck.rpt"
                        Else
                            .ReportFileName = App.Path & "\Reports\QuotationAKPTruck.rpt"
                        End If
                    Else
                        If myQuotation_.isQuotationReq = False Then
                            .ReportFileName = App.Path & "\Reports\QuotationAKPTruck.rpt"
                        Else
                            .ReportFileName = App.Path & "\Reports\QuotationAKPTruck.rpt"
                        End If
                    End If
                End If

            End If
        End If
        .Formulas(0) = "ShowPODate=" & Chr(34) & dtIssuedDate.ValueDMY & Chr(34)
        .Formulas(1) = "CustName=" & Chr(34) & txtCustomer.Text & Chr(34)
        .Formulas(2) = "Addr1=" & Chr(34) & curCustomer_.Address & " " & curCustomer_.TumbolName & " " & curCustomer_.AumphurName & " " & curCustomer_.ProvinceName & Chr(34)
        .Formulas(3) = "TelNo=" & Chr(34) & curCustomer_.TelNo & Chr(34)
        .Formulas(4) = "FaxNo=" & Chr(34) & curCustomer_.FaxNo & Chr(34)
        If chkIncTrans.Value = 0 Then
            .Formulas(5) = "IsIncTransFee=" & Chr(34) & "ราคายังไม่รวมค่าขนส่ง" & Chr(34)
        Else
            .Formulas(5) = "IsIncTransFee=" & Chr(34) & "ฟรีค่าขนส่ง" & Chr(34)
        End If
        If curCustomer_.CustomerType = "M" Or curCustomer_.CustomerType = "MS" Then
        .Formulas(6) = "AttentionTo=" & Chr(34) & "คุณจารุวรรณ โพธิ์แจ้ง" & Chr(34)
        Else
        .Formulas(6) = "AttentionTo=" & Chr(34) & cboCustStaff.Text & Chr(34)
        End If
        .Formulas(7) = "QuotationNo=" & Chr(34) & Replace(txtQuoNo.Text, "(CR)", "") & Chr(34)
        .Formulas(8) = "ApprovalName=" & Chr(34) & "(  " & cboApprove.Text & "  )" & Chr(34)
        .Formulas(9) = "ApprovePosition=" & Chr(34) & txtPosition.Text & Chr(34)
        If myQuotation_.CompanyID = "776EDA5830F744A6B3551348D851BC22" Then
            If myQuotation_.SaleAP <> "" Then
                .Formulas(10) = "SaleName=" & Chr(34) & curCustomer_.UnderSaleName & " (" & myQuotation_.SaleAP & ") โทร. " & curCustomer_.UnderSaleTelNo & Chr(34)
            Else
                .Formulas(10) = "SaleName=" & Chr(34) & curCustomer_.UnderSaleName & " โทร. " & curCustomer_.UnderSaleTelNo & Chr(34)
            End If
        Else
            .Formulas(10) = "SaleName=" & Chr(34) & curCustomer_.UnderSaleName & " โทร. " & curCustomer_.UnderSaleTelNo & Chr(34)
        End If
        Dim TrType, TrType2, TrType3, TrType4, TrType5, TrType6, TrType7, TrType8, TrType9, TrType10 As String
        Dim sqlSelTruck As String
        Dim tmpRecSelTruck As ADODB.Recordset
        Dim x As Integer
        TrType = ""
        TrType2 = ""
        TrType3 = ""
        TrType4 = ""
        TrType5 = ""
        TrType6 = ""
        TrType7 = ""
        TrType8 = ""
        TrType9 = ""
        TrType10 = ""
        Set tmpRecSelTruck = New ADODB.Recordset
        sqlSelTruck = " select * from vw_Seltruck  Where QuotationID = '" & myQuotation_.id & "' order by SelTruckID"
        Set tmpRecSelTruck = GetRS(sqlSelTruck, adOpenForwardOnly, adLockReadOnly)
        If tmpRecSelTruck.EOF = False Then
            For x = 1 To tmpRecSelTruck.RecordCount
                If TrType = "" Then
                    TrType = tmpRecSelTruck("TruckTypeDesc")
                    .Formulas(42) = "TruckType=" & Chr(34) & tmpRecSelTruck("TrailerType") & Chr(34)
                    .Formulas(43) = "TransFee=" & Chr(34) & Format(tmpRecSelTruck("TransportFee"), "#,##0.00") & Chr(34)
                    .Formulas(44) = "LW=" & Chr(34) & tmpRecSelTruck("MinWeight") & Chr(34)
                ElseIf TrType2 = "" Then
                    TrType2 = tmpRecSelTruck("TruckTypeDesc")
                    .Formulas(45) = "TruckType2=" & Chr(34) & tmpRecSelTruck("TrailerType") & Chr(34)
                    .Formulas(46) = "TransFee2=" & Chr(34) & Format(tmpRecSelTruck("TransportFee"), "#,##0.00") & Chr(34)
                    .Formulas(47) = "LW2=" & Chr(34) & tmpRecSelTruck("MinWeight") & Chr(34)
                ElseIf TrType3 = "" Then
                    TrType3 = tmpRecSelTruck("TruckTypeDesc")
                    .Formulas(48) = "TruckType3=" & Chr(34) & tmpRecSelTruck("TrailerType") & Chr(34)
                    .Formulas(49) = "TransFee3=" & Chr(34) & Format(tmpRecSelTruck("TransportFee"), "#,##0.00") & Chr(34)
                    .Formulas(50) = "LW3=" & Chr(34) & tmpRecSelTruck("MinWeight") & Chr(34)
                ElseIf TrType4 = "" Then
                    TrType4 = tmpRecSelTruck("TruckTypeDesc")
                    .Formulas(51) = "TruckType4=" & Chr(34) & tmpRecSelTruck("TrailerType") & Chr(34)
                    .Formulas(52) = "TransFee4=" & Chr(34) & Format(tmpRecSelTruck("TransportFee"), "#,##0.00") & Chr(34)
                    .Formulas(53) = "LW4=" & Chr(34) & tmpRecSelTruck("MinWeight") & Chr(34)
                ElseIf TrType5 = "" Then
                    TrType5 = tmpRecSelTruck("TruckTypeDesc")
                    .Formulas(54) = "TruckType5=" & Chr(34) & tmpRecSelTruck("TrailerType") & Chr(34)
                    .Formulas(55) = "TransFee5=" & Chr(34) & Format(tmpRecSelTruck("TransportFee"), "#,##0.00") & Chr(34)
                    .Formulas(56) = "LW5=" & Chr(34) & tmpRecSelTruck("MinWeight") & Chr(34)
                ElseIf TrType6 = "" Then
                    TrType6 = tmpRecSelTruck("TruckTypeDesc")
                    .Formulas(57) = "TruckType6=" & Chr(34) & tmpRecSelTruck("TrailerType") & Chr(34)
                    .Formulas(58) = "TransFee6=" & Chr(34) & Format(tmpRecSelTruck("TransportFee"), "#,##0.00") & Chr(34)
                    .Formulas(59) = "LW6=" & Chr(34) & tmpRecSelTruck("MinWeight") & Chr(34)
                ElseIf TrType7 = "" Then
                    TrType7 = tmpRecSelTruck("TruckTypeDesc")
                    .Formulas(60) = "TruckType7=" & Chr(34) & tmpRecSelTruck("TrailerType") & Chr(34)
                    .Formulas(61) = "TransFee7=" & Chr(34) & Format(tmpRecSelTruck("TransportFee"), "#,##0.00") & Chr(34)
                    .Formulas(62) = "LW7=" & Chr(34) & tmpRecSelTruck("MinWeight") & Chr(34)
                ElseIf TrType8 = "" Then
                    TrType8 = tmpRecSelTruck("TruckTypeDesc")
                    .Formulas(63) = "TruckType8=" & Chr(34) & tmpRecSelTruck("TrailerType") & Chr(34)
                    .Formulas(64) = "TransFee8=" & Chr(34) & Format(tmpRecSelTruck("TransportFee"), "#,##0.00") & Chr(34)
                    .Formulas(65) = "LW8=" & Chr(34) & tmpRecSelTruck("MinWeight") & Chr(34)
                ElseIf TrType9 = "" Then
                    TrType9 = tmpRecSelTruck("TruckTypeDesc")
                    .Formulas(66) = "TruckType9=" & Chr(34) & tmpRecSelTruck("TrailerType") & Chr(34)
                    .Formulas(67) = "TransFee9=" & Chr(34) & Format(tmpRecSelTruck("TransportFee"), "#,##0.00") & Chr(34)
                    .Formulas(68) = "LW9=" & Chr(34) & tmpRecSelTruck("MinWeight") & Chr(34)
                ElseIf TrType10 = "" Then
                    TrType10 = tmpRecSelTruck("TruckTypeDesc")
                    .Formulas(69) = "TruckType10=" & Chr(34) & tmpRecSelTruck("TrailerType") & Chr(34)
                    .Formulas(70) = "TransFee10=" & Chr(34) & Format(tmpRecSelTruck("TransportFee"), "#,##0.00") & Chr(34)
                    .Formulas(71) = "LW10=" & Chr(34) & tmpRecSelTruck("MinWeight") & Chr(34)
                End If
                tmpRecSelTruck.MoveNext
            Next
        End If
            .Formulas(32) = "Truckselect=" & Chr(34) & TrType & Chr(34)
            .Formulas(33) = "Truckselect2=" & Chr(34) & TrType2 & Chr(34)
            .Formulas(34) = "Truckselect3=" & Chr(34) & TrType3 & Chr(34)
            .Formulas(35) = "Truckselect4=" & Chr(34) & TrType4 & Chr(34)
            .Formulas(36) = "Truckselect5=" & Chr(34) & TrType5 & Chr(34)
            .Formulas(37) = "Truckselect6=" & Chr(34) & TrType6 & Chr(34)
            .Formulas(38) = "Truckselect7=" & Chr(34) & TrType7 & Chr(34)
            .Formulas(39) = "Truckselect8=" & Chr(34) & TrType8 & Chr(34)
            .Formulas(40) = "Truckselect9=" & Chr(34) & TrType9 & Chr(34)
            .Formulas(41) = "Truckselect10=" & Chr(34) & TrType10 & Chr(34)
            
        tmpSql = ""
        tmpSql = "SELECT * FROM Com_vwQuotationDetailSearch " & vbCrLf
        tmpSql = tmpSql & "WHERE QuotationID='" & myQuotation_.id & "' " & vbCrLf
        tmpSql = tmpSql & "ORDER BY RowPos"
        .SQLQuery = tmpSql
        .Destination = crptToWindow
        .PrintReport
        If Trim(.LastErrorString) <> "" Then MsgBox .LastErrorString
    End With
End Sub

Private Sub cmdPrintPDF_Click()
Dim tmpSql As String
Dim selResult
Dim tmpStr As String
Dim AcceptState As Boolean, i As Integer
    tmpStr = CurrentUser.VerifyAccessRole(Me.Name, PrintRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If

    If Trim(txtCustomer.Text) = "" Then
        MsgBox "กรุณาใส่ชื่อลูกค้า", vbExclamation
        Exit Sub
    End If
    If Trim(cboApprove.Text) = "" Then
        MsgBox "กรุณาระบุชื่อผู้อนุมัติ", vbExclamation
        cboApprove.SetFocus
        Exit Sub
    End If
    If Trim(txtPosition.Text) = "" Then
        MsgBox "กรุณาระบุตำแหน่งผู้อนุมัติ", vbExclamation
        txtPosition.SetFocus
        Exit Sub
    End If

    If myQuotation_.isChanged And cmdSave.Enabled Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนสร้างใหม่หรือไม่?", vbQuestion + vbYesNo)
        If selResult = vbYes Then
            CmdSave_Click
            If myQuotation_.isChanged Then Exit Sub
        Else 'If selResult = vbCancel Then
            Exit Sub
        End If
    End If

    With rpt1
        .WindowShowCloseBtn = True
        .WindowShowExportBtn = True
        .WindowShowPrintSetupBtn = True
        .LogonInfo(0) = "DSN = " & DSN_NAME & ";UID = " & strUid & ";PWD = " & strPwd
       .ReportFileName = App.Path & "\Reports\QuotationAKPNew.rpt"
        .SelectionFormula = "{vw_QuotationHeaderSearch.QuotationID} = '" & myQuotation_.id & "'"
        '.SQLQuery = tmpSql
        
'        .Destination = crptToWindow
'        .PrintReport
        If Trim(.LastErrorString) <> "" Then MsgBox .LastErrorString
    End With
    
    Dim myCRApp As New CRAXDRT.Application
    Dim myCRReport As CRAXDRT.Report
    Dim adoCmd As ADODB.Command, rs As New ADODB.Recordset
    Dim crxTable As CRAXDRT.DatabaseTable
    Dim AppPath As Scripting.FileSystemObject
    Dim PDFPath As String
    
'    GetDBConnection
'    Set adoCmd = New ADODB.Command
'    Set adoCmd.ActiveConnection = DBConn
'    adoCmd.CommandText = tmpSql
'    adoCmd.CommandType = adCmdText
'
'    Set rs = GetRS(tmpSql, adOpenKeyset, adLockOptimistic)
'    GetDBConnection
    
    Set myCRApp = New CRAXDRT.Application
    Set myCRReport = myCRApp.OpenReport(rpt1.ReportFileName)
'    myCRApp.NewReport
'    myCRApp.LogOnServer "sqloledb.dll", dbServer, dbVolumn, uID, uPwd
    'myCRReport.DiscardSavedData
    For Each crxTable In myCRReport.Database.Tables
        crxTable.SetLogOnInfo DSN_NAME, strDb, strUid, strPwd
'        crxTable.SetDataSource tmpSql
'        crxTable.SetDataSource RS
    Next
'    myCRReport.Database.AddADOCommand DBConn, adoCmd
    myCRReport.SQLQueryString = tmpSql
    Set AppPath = New Scripting.FileSystemObject
    PDFPath = "C:\BWG_PDFExport\"
    If AppPath.FolderExists(PDFPath) = False Then
        AppPath.CreateFolder PDFPath
    End If
    
'    myCRReport.FormulaFields.GetItemByName("ShowPODate").Text = Chr(34) & dtIssuedDate.ValueDMY & Chr(34)
'    myCRReport.FormulaFields.GetItemByName("CustName").Text = Chr(34) & txtCustomer.Text & Chr(34)
'    myCRReport.FormulaFields.GetItemByName("Addr1").Text = Chr(34) & curCustomer_.Address & " " & curCustomer_.TumbolName & " " & curCustomer_.AumphurName & " " & curCustomer_.ProvinceName & Chr(34)
'    myCRReport.FormulaFields.GetItemByName("TelNo").Text = Chr(34) & curCustomer_.TelNo & Chr(34)
'    myCRReport.FormulaFields.GetItemByName("FaxNo").Text = Chr(34) & curCustomer_.FaxNo & Chr(34)
'    If chkIncTrans.Value = 0 Then
'        myCRReport.FormulaFields.GetItemByName("IsIncTransFee").Text = Chr(34) & "ราคายังไม่รวมค่าขนส่ง" & Chr(34)
'    Else
'        myCRReport.FormulaFields.GetItemByName("IsIncTransFee").Text = Chr(34) & "ฟรีค่าขนส่ง" & Chr(34)
'    End If
'    myCRReport.FormulaFields.GetItemByName("AttentionTo").Text = Chr(34) & cboCustStaff.Text & Chr(34)
'    myCRReport.FormulaFields.GetItemByName("QuotationNo").Text = Chr(34) & txtQuoNo.Text & Chr(34)
'    myCRReport.FormulaFields.GetItemByName("ApprovalName").Text = Chr(34) & "(  " & cboApprove.Text & "  )" & Chr(34)
'    myCRReport.FormulaFields.GetItemByName("ApprovePosition").Text = Chr(34) & TxtPosition.Text & Chr(34)
'    If myQuotation_.CompanyID = "776EDA5830F744A6B3551348D851BC22" Then
'        If myQuotation_.SaleAP <> "" Then
'            myCRReport.FormulaFields.GetItemByName("SaleName").Text = Chr(34) & curCustomer_.UnderSaleName & " (" & myQuotation_.SaleAP & ") โทร. " & curCustomer_.UnderSaleTelNo & Chr(34)
'        Else
'            myCRReport.FormulaFields.GetItemByName("SaleName").Text = Chr(34) & curCustomer_.UnderSaleName & " โทร. " & curCustomer_.UnderSaleTelNo & Chr(34)
'        End If
'    Else
'        myCRReport.FormulaFields.GetItemByName("SaleName").Text = Chr(34) & curCustomer_.UnderSaleName & " โทร. " & curCustomer_.UnderSaleTelNo & Chr(34)
'    End If
        
    myCRReport.ExportOptions.FormatType = crEFTPortableDocFormat
    myCRReport.ExportOptions.DestinationType = crEDTApplication
    myCRReport.ExportOptions.DiskFileName = PDFPath & myQuotation_.QuotationNo & ".pdf"
    myCRReport.ExportOptions.ApplicationFileName = PDFPath & myQuotation_.QuotationNo & ".pdf"
    myCRReport.Export False
    Set myCRApp = Nothing
    Set myCRReport = Nothing
    CloseDBConnection
    Exit Sub
ErrD:
    MsgBox err.Description
    Exit Sub
End Sub


Private Sub CmdQuoTruck_Click()
    With myQuotation_
        If CmdQuoTruck.Caption = "ใบเสนอราคาสำหรับเลือกรถ" Then
        
            txtTransRate2.Enabled = True
            cboTransUnit2.Enabled = True
            txtTransRate3.Enabled = True
            cboTransUnit3.Enabled = True
            txtTransRate4.Enabled = True
            cboTransUnit4.Enabled = True
            txtTransRate5.Enabled = True
            cboTransUnit5.Enabled = True
            txtTransRate6.Enabled = True
            cboTransUnit6.Enabled = True
            txtTransRate7.Enabled = True
            cboTransUnit7.Enabled = True
            txtTransRate8.Enabled = True
            cboTransUnit8.Enabled = True
            txtTransRate9.Enabled = True
            cboTransUnit9.Enabled = True
            txtCarType.Enabled = True
            
            txtTransRate2.Text = "0"
            txtTransRate3.Text = "0"
            txtTransRate4.Text = "0"
            txtTransRate5.Text = "0"
            txtTransRate6.Text = "0"
            txtTransRate7.Text = "0"
            txtTransRate8.Text = "0"
            txtTransRate9.Text = "0"
            
            txtCarType.Text = "รถ Lugger Box"
            cboIsTrailer.ListIndex = 0
            .TruckTypeID = "A0D27E874D9D4198B9E00DCF784515EA"
            .isTrailer = False
            .TRuck1ID = "A0D27E874D9D4198B9E00DCF784515EA"
            .Truck1Trailer = "N"

            txtCarType2.Text = "รถ Lugger Box - พ่วง"
            cboIsTrailer2.Text = "รถพ่วง"
            .TRuck2ID = "563D175CF91D485E90F84FB6143BB48E"
            .Truck2Trailer = "Y"

            txtCarType3.Text = "รถ Roll off Truck - เดี่ยว"
            cboIsTrailer3.Text = "รถเดี่ยว"
            .TRuck3ID = "F27CB5C50AFD4D198363035AE4BB529C"
            .TRuck3Trailer = "N"

            txtCarType4.Text = "รถ Roll off Truck - พ่วง"
            cboIsTrailer4.Text = "รถพ่วง"
            .TRuck4ID = "D60271D109074C51878317FE36E00C35"
            .TRuck4Trailer = "Y"

            txtCarType5.Text = "รถกระบะ"
            cboIsTrailer5.Text = "รถเดี่ยว"
            .TRuck5ID = "0392E12DBCD94CE093B4FF1B7BE240BE"
            .TRuck5Trailer = "N"

            txtCarType6.Text = "รถเทรลเลอร์"
            cboIsTrailer6.Text = "รถเดี่ยว"
            .TRuck6ID = "BF134C1258264AA6ABF0E9ECD4F0BD99"
            .TRuck6Trailer = "N"

              txtCarType7.Text = "รถแท็งค์ 15 ลูกบาศก์เมตร"
            cboIsTrailer7.Text = "รถเดี่ยว"
            .TRuck7ID = "67D1CCC9C8FF4782B89E3F56100A1592"
            .TRuck7Trailer = "N"

             txtCarType8.Text = "รถแท็งค์ ISO 24 ลูกบาศก์เมตร"
            cboIsTrailer8.Text = "รถเดี่ยว"
            .TRuck8ID = "7027516E42864192B470CA018350EF87"
            .TRuck8Trailer = "N"

              txtCarType9.Text = "รถโรงงาน"
            .TRuck9ID = "35A046631D8A4188BE0F6EA83A4D88E4"

         CmdQuoTruck.Caption = "ใบเสนอราคาใช้จริง"
        Else
            txtCarType.Enabled = True
            cboIsTrailer.Enabled = True
            txtCarType.Text = ""
            cboIsTrailer.ListIndex = -1
            txtTransRate.Text = ""
            cboTransUnit.ListIndex = -1
          
            txtCarType2.Text = ""
            cboIsTrailer2.ListIndex = -1
            txtTransRate2.Text = ""
            cboTransUnit2.ListIndex = -1
            txtTransRate2.Enabled = False
            cboTransUnit2.Enabled = False
            
            txtCarType3.Text = ""
            cboIsTrailer3.ListIndex = -1
            txtTransRate3.Text = ""
            cboTransUnit3.ListIndex = -1
            txtTransRate3.Enabled = False
            cboTransUnit3.Enabled = False

            txtCarType4.Text = ""
            cboIsTrailer4.ListIndex = -1
            txtTransRate4.Text = ""
            cboTransUnit4.ListIndex = -1
            txtTransRate4.Enabled = False
            cboTransUnit4.Enabled = False
            
            txtCarType5.Text = ""
            cboIsTrailer5.ListIndex = -1
            txtTransRate5.Text = ""
            cboTransUnit5.ListIndex = -1
            txtTransRate5.Enabled = False
            cboTransUnit5.Enabled = False
            
            txtCarType6.Text = ""
            cboIsTrailer6.ListIndex = -1
            txtTransRate6.Text = ""
            cboTransUnit6.ListIndex = -1
            txtTransRate6.Enabled = False
            cboTransUnit6.Enabled = False
            
            txtCarType7.Text = ""
            cboIsTrailer7.ListIndex = -1
            txtTransRate7.Text = ""
            cboTransUnit7.ListIndex = -1
            txtTransRate7.Enabled = False
            cboTransUnit7.Enabled = False
            
            txtCarType8.Text = ""
            cboIsTrailer8.ListIndex = -1
            txtTransRate8.Text = ""
            cboTransUnit8.ListIndex = -1
            txtTransRate8.Enabled = False
            cboTransUnit8.Enabled = False
        
            txtCarType9.Text = ""
            cboIsTrailer9.ListIndex = -1
            txtTransRate9.Text = ""
            cboTransUnit9.ListIndex = -1
            txtTransRate9.Enabled = False
            cboTransUnit9.Enabled = False
            
            .TruckTypeID = ""
            .TRuck1ID = ""
            .Truck1Trailer = ""
            .Truck1Rate = ""
            .Truck1Unit = ""
            
            .TRuck2ID = ""
            .Truck2Trailer = ""
            .Truck2Rate = ""
            .Truck2Unit = ""
            
            .TRuck3ID = ""
            .TRuck3Trailer = ""
            .TRuck3Rate = ""
            .TRuck3Unit = ""
            
            .TRuck4ID = ""
            .TRuck4Trailer = ""
            .TRuck4Rate = ""
            .TRuck4Unit = ""
            
            .TRuck5ID = ""
            .TRuck5Trailer = ""
            .TRuck5Rate = ""
            .TRuck5Unit = ""
            
            .TRuck6ID = ""
            .TRuck6Trailer = ""
            .TRuck6Rate = ""
            .TRuck6Unit = ""
            
            .TRuck7ID = ""
            .TRuck7Trailer = ""
            .TRuck7Rate = ""
            .TRuck7Unit = ""
            
            .TRuck8ID = ""
            .TRuck8Trailer = ""
            .TRuck8Rate = ""
            .TRuck8Unit = ""
            
            .TRuck9ID = ""
            .TRuck9Trailer = ""
            .TRuck9Rate = ""
            .TRuck9Unit = ""
            
            CmdQuoTruck.Caption = "ใบเสนอราคาสำหรับเลือกรถ"

        End If
    End With
End Sub

Private Sub CmdSave_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole(Me.Name, curRole_)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
         Dim SqlQuo As String
         Dim tmpRec As ADODB.Recordset
             Set tmpRec = New ADODB.Recordset
             SqlQuo = "select *  from vw_SaleDistrictSearch  Where SaleDistrictID = '" & curCustomer_.SaleDistrictID & "' "
             Set tmpRec = GetRS(SqlQuo, adOpenForwardOnly, adLockReadOnly)
                 If tmpRec.EOF = False Then
                    If tmpRec!ZoneActive <> "Y" Then
                        MsgBox "เขตการขายของลูกค้านี้ไม่ถูกต้อง กรุณาตรวจสอบข้อมูลเขตการขาย", vbExclamation
                        Exit Sub
                    End If
                    If tmpRec!isActive <> "Y" Then
                        MsgBox "เขตการขายของลูกค้านี้ไม่ถูกต้อง กรุณาตรวจสอบข้อมูลเขตการขาย", vbExclamation
                        Exit Sub
                    End If
                    
                 End If
    If cboCompany.ListIndex < 0 Then
        MsgBox "กรุณาเลือกบริษัทที่ออกใบเสนอราคา", vbExclamation
        cboCompany.SetFocus
        Exit Sub
    End If
    If Trim(txtCustomer.Text) = "" Then
        MsgBox "กรุณาใส่ชื่อลูกค้า", vbExclamation
        Exit Sub
    End If
    'If UCase(Left(myQuotation_.CustomerCode, 1)) = "M" And (UCase(Left(myQuotation_.BMEQuotation, 2)) <> "QT" Or Len(myQuotation_.BMEQuotation) <> 10) Then
    '    MsgBox "กรุณาระบุเลขที่ใบเสนอราคาของ BME", vbExclamation
    '    Exit Sub
    'End If
    If myQuotation_.QuotationDetails Is Nothing Then
        MsgBox "กรุณาระบุรายการกากของเสียอย่างน้อย 1 รายการ", vbExclamation
        Exit Sub
    End If
    If myQuotation_.QuotationDetails.Count = 0 Then
        MsgBox "กรุณาระบุรายการกากของเสียอย่างน้อย 1 รายการ", vbExclamation
        Exit Sub
    End If
    If IsNumeric(txtLaborCount.Text) = False And Trim(txtLaborCount.Text) <> "" Then
        MsgBox "กรุณาใส่จำนวนคนงานเป็นจำนวนตัวเลข", vbExclamation
        txtLaborCount.SetFocus
        Exit Sub
    End If
    If IsNumeric(txtMinQty.Text) = False And Trim(txtMinQty.Text) <> "" Then
        MsgBox "กรุณาใส่ปริมาณขั้นต่ำของของเสียที่ต้องขนเป็นจำนวนตัวเลข", vbExclamation
        txtMinQty.SetFocus
        Exit Sub
    End If
    
    If IsNumeric(txtDiscount.Text) = False And Trim(txtDiscount.Text) <> "" Then
        MsgBox "กรุณาใส่ส่วนลด เป็นจำนวนตัวเลข", vbExclamation
        txtDiscount.SetFocus
        Exit Sub
    End If
    If IsNumeric(txtExtraCharge.Text) = False And Trim(txtExtraCharge.Text) <> "" Then
        MsgBox "กรุณาใส่ค่าแรงคนงานเพิ่มเป็นตัวเลข", vbExclamation
        txtExtraCharge.SetFocus
        Exit Sub
    End If
    If Trim(txtCarType.Text) = "" Then
        MsgBox "กรุณาระบุประเภทรถที่ใช้ในการขนส่ง", vbExclamation
        txtCarType.SetFocus
        Exit Sub
    End If
    If chkIncTrans.Value = 0 Then
        If Trim(txtTransRate.Text) <> "" And IsNumeric(Trim(txtTransRate.Text)) = False Then
            MsgBox "กรุณาใส่อัตราค่าขนส่งเป็นตัวเลขเท่านั้น", vbExclamation
            txtTransRate.SetFocus
            Exit Sub
        Else
            Dim ExcTrans As Double
            If Trim(txtTransRate.Text) = "" Then
                ExcTrans = 0
            Else
                ExcTrans = CDbl(txtTransRate.Text)
            End If
            If ExcTrans = 0 Then
                '*** ตรวจสอบค่าขนส่งจากฐานข้อมูล
                Dim CustomerChargeRate As BWGTransVendorChargeRate
                Dim ToAumphurID As String
                Set CustomerChargeRate = Nothing
                If myQuotation_.CompanyID = "776EDA5830F744A6B3551348D851BC22" Then
                    ToAumphurID = "AP200506271557102"
                Else
                    ToAumphurID = "AP2005062715572615"
                End If
                Set CustomerChargeRate = CalTransFeeFromDistance_Cust(curCustomer_.AumphurID, ToAumphurID, myQuotation_.TruckTypeID, myQuotation_.IssuedDate)
                If Not CustomerChargeRate Is Nothing Then
                    If myQuotation_.isTrailer = True Then
                        ExcTrans = CustomerChargeRate.TrailerRateTrip
                    Else
                        ExcTrans = CustomerChargeRate.RatePerTrip
                    End If
                Else
                    ExcTrans = 0
                End If
'                If Trim(myQuotation_.TruckTypeName) <> "รถโรงงาน" Then
'                    MsgBox "กรุณาระบุอัตราค่าขนส่งที่เป็นไปได้ (ควรเก็บกับลูกค้าประมาณ = " & Format(ExcTrans, "#,##0.00") & " บาท/เที่ยว", vbExclamation
'                    txtTransRate.SetFocus
'                    txtTransRate.SelStart = 0
'                    txtTransRate.SelLength = Len(txtTransRate)
'                    Exit Sub
'                Else
'                    txtTransRate.Text = ""
'                End If
            End If
        End If
'        If Trim(cboTransUnit.Text) = "" Then
'            If Trim(myQuotation_.TruckTypeName) <> "รถโรงงาน" Then
'                MsgBox "กรุณาระบุหน่วยในการคิดค่าขนส่ง", vbExclamation
'                cboTransUnit.SetFocus
'                Exit Sub
'            End If
'        End If
    End If
    
    If Trim(txtMaxWeight.Text) <> "" Then
        If IsNumeric(txtMaxWeight.Text) Then
            If Trim(txtMaxRate.Text) = "" Then
                MsgBox "กรุณาระบุอัตราส่วนเกิน", vbCritical
                cboMaxRateType.ListIndex = -1
                txtMaxRate.SetFocus
                Exit Sub
            Else
                If IsNumeric(txtMaxRate.Text) = False Then
                    MsgBox "กรุณาระบุอัตราส่วนเกินเป็นตัวเลขเท่านั้น", vbCritical
                    txtMaxRate.Text = ""
                    cboMaxRateType.ListIndex = -1
                    txtMaxRate.SetFocus
                    Exit Sub
                Else
                    If cboMaxRateType.ListIndex < 0 Then
                        MsgBox "กรุณาเลือกประเภทราคาอัตรส่วนเกิน", vbCritical
                        cboMaxRateType.SetFocus
                        Exit Sub
                    End If
                End If
            End If
        Else
            MsgBox "ปริมาณการขนสูงสุด ให้ใส่เป็นตัวเลขเท่านั้น", vbCritical
            txtMaxWeight.Text = ""
            txtMaxRate.Text = ""
            cboMaxRateType.ListIndex = -1
            txtMaxWeight.SetFocus
            Exit Sub
        End If
    Else
        txtMaxRate.Text = "0"
        cboMaxRateType.ListIndex = 0
    End If
    
    If cboPrintForm.ListIndex < 0 Then
        MsgBox "กรุณาระบุแบบฟอร์ม", vbInformation, "Warning"
        Exit Sub
    End If
    
    '*** Check Status for IsJobs
    If cboQuotationType.ListIndex = 0 Then
        myQuotation_.isJob = "N"
    ElseIf cboQuotationType.ListIndex = 1 Then
        myQuotation_.isJob = "Y"
    ElseIf cboQuotationType.ListIndex = 2 Then
        myQuotation_.isJob = "D"
    End If
    If Trim(myQuotation_.isJob) = "" Then myQuotation_.isJob = "N"
    If frmPromiesChang = True Then myQuotation_.isChanged = True
    myQuotation_.Promies = QuotationPromies
    myQuotation_.StartPromiesDate = QuotationStartDatePromies
    myQuotation_.EndPromiesDate = QuotationEndDatePromies
    
    '*** Check Max Weight
    If IsNumeric(txtMaxWeight.Text) Then myQuotation_.MaxWeight = txtMaxWeight.Text
    If IsNumeric(txtMaxRate.Text) Then myQuotation_.MaxRate = txtMaxRate.Text
    If cboMaxRateType.ListIndex >= 0 Then myQuotation_.MaxRateType = cboMaxRateType.Text
    '***********************************
    
    '*** Check isPriceIncTransport
    If chkIncTrans.Value = 1 Then
        myQuotation_.IsPriceIncTransport = True
    Else
        myQuotation_.IsPriceIncTransport = False
    End If
      If chkTrnDis.Value = 1 Then
        myQuotation_.IsDiscountPriceTrans = True
    Else
        myQuotation_.IsDiscountPriceTrans = False
    End If
    
    
      If Trim(cboApprove.Text) = "" Then
        MsgBox "กรุณาระบุชื่อผู้อนุมัติ", vbExclamation
        cboApprove.SetFocus
        Exit Sub
        Else
            myQuotation_.SignName = cboApprove.Text
            myQuotation_.SignPosition = Trim(txtPosition)
            myQuotation_.isChanged = True
    End If
      If cboTransUnit.ListIndex > -1 Then
        myQuotation_.TransferUnitID = cboTransUnit.Text
        myQuotation_.TransferUnitName = cboTransUnit.Text
        myQuotation_.Truck1Unit = cboTransUnit.Text
    Else
        myQuotation_.TransferUnitID = ""
        myQuotation_.TransferUnitName = ""
        myQuotation_.Truck1Unit = ""
    End If
    If Trim(txtPosition.Text) = "" Then
        MsgBox "กรุณาระบุตำแหน่งผู้อนุมัติ", vbExclamation
        txtPosition.SetFocus
        Exit Sub
        Else
            myQuotation_.SignName = cboApprove.Text
            myQuotation_.SignPosition = Trim(txtPosition)
            myQuotation_.isChanged = True
    End If
    
    If myQuotation_.isChanged And cmdSave.Enabled Then
        If Trim(myQuotation_.CompanyID) = "" Then
            MsgBox "ข้อมูลบริษัทหายไป กรุณาตรวจสอบแล้วลองใหม่ (ปิดเปิดใหม่)", vbInformation
            Exit Sub
        End If
        myQuotation_.Save
        'SaveselTtruck
        If txtCarType2.Text <> "" Then
            myQuotation_.SaveQuoTruck
        End If
        txtLastUser.Text = myQuotation_.LastUpdateUser
        isSave = True
        cboPrintForm.Enabled = False
        If GrCar.getGridObj.TextMatrix(1, 0) <> "" Then
            txtQuoNo.Text = myQuotation_.QuotationNo & "(CR)"
        Else
            txtQuoNo.Text = myQuotation_.QuotationNo
        End If
        If myQuotation_.isApproved Then
            Label2(4).Caption = "อนุมัติแล้วเมื่อวันที่ " & FormatYMD_to_DMY(myQuotation_.Approveddate, "/", "/") & " โดย " & myQuotation_.ApprovedStaffName
        Else
            Label2(4).Caption = "รายการรออนุมัติ"
        End If
    End If
    If myQuotation_.isChanged = False Then
        MsgBox "จัดเก็บข้อมูลเรียบร้อยแล้ว", vbExclamation
    End If
End Sub

Private Sub cmdSearch_Click()
On Error Resume Next
Dim selResult As Integer
Dim x%
Dim tmpKey As String
Dim tmpVal As String
Dim tmpDet As BWGQuotationDetail
Dim tmpCol As Collection
    If myQuotation_.isChanged And cmdSave.Enabled Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนทำงานต่อหรือไม่?", vbQuestion + vbYesNoCancel)
        If selResult = vbYes Then
            CmdSave_Click
            If myQuotation_.isChanged Then Exit Sub
        ElseIf selResult = vbCancel Then
            Exit Sub
        End If
    End If
    Load frmQuotationSearch
    With frmQuotationSearch
        .SetCallerForm Me
        .Show 1
    End With
End Sub

Private Sub CmdSelTruck_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole(frmRemDesc.Name, AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    If myQuotation_.TumbolID = "" Then
       MsgBox "ท่านยังไม่ได้เลือกสถานที่เก็บขน", vbExclamation, "Warning"
       Exit Sub
    End If
    Load FrmSelTruck
    With FrmSelTruck
        .SetTumbolID myQuotation_.TumbolID
        .lvResult.Checkboxes = True
        .lvResult.Refresh
        .cmdOK.Visible = True
        .SetCallerForm Me
        
        .Show 1
        ShowCar
    End With
End Sub
Public Sub AddCar(cartem As BWGTruckRate)
    With cartem
        
    End With
End Sub

Private Sub Command1_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole(frmRemDesc.Name, AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    
    Load frmRemDesc
    With frmRemDesc
        .lvList.Checkboxes = True
        .lvList.Refresh
        .cmdOK.Visible = True
        .SetCallerForm Me
        .Show 1
    End With
End Sub

Private Sub Command2_Click()
frmPromist.Show 1
End Sub

Sub SaveselTtruck()
        Dim r As Integer
        Dim i As Integer
        Dim tSql As String
        Dim sqlSel As String
        Dim Tt As String
        Dim tmpRec As ADODB.Recordset
            With GrCar
                r = .getGridObj.Rows - 1
                For i = 1 To r - 1
                Set tmpRec = New ADODB.Recordset
                Tt = myQuotation_.TruckTypeID
                sqlSel = " select * from vw_Seltruck  Where QuotationID = '" & myQuotation_.id & "' and TruckTypeID='" & Tt & "' order by SelTruckID"
                Set tmpRec = GetRS(sqlSel, adOpenForwardOnly, adLockReadOnly)
                    If tmpRec.EOF = False Then
                        tSql = "update tbSelTruck set TrailerType='" & cboIsTrailer.Text & "',TransportFee=" & Val(txtTransRate.Text) & ",MinWeight=" & Val(txtMinQty.Text) & "" & vbCrLf
                        tSql = tSql & " Where SelTruckID = '" & tmpRec("SelTruckID") & "' "
                        DBConnExc tSql
                    Else
                        tSql = "INSERT INTO tbSelTruck(SelTruckID, QuotationID,TruckTypeID,TrailerType,TransportFee,MinWeight)" & vbCrLf
                        tSql = tSql & "VALUES ('" & GetGUID & "','" & myQuotation_.id & "', '" & myQuotation_.TruckTypeID & "','" & cboIsTrailer.Text & "'," & Val(txtTransRate.Text) & "," & Val(txtMinQty.Text) & ")" & vbCrLf
                        DBConnExc tSql
                    End If
                Next i
            End With
            
End Sub

Private Sub dtIssuedDate_Validate(Cancel As Boolean)
    myQuotation_.IssuedDate = dtIssuedDate.ValueYMD
End Sub

Private Sub Form_Load()
    Dim i As Integer, x As Integer
    CanAccess = False: CanUpdate = False
    SetPermission Me, CanAccess, CanUpdate
    If CanAccess = False Then
        MsgBox "ท่านไม่มีสิทธิใช้ระบบนี้", vbExclamation, "Warning"
        Unload Me
    End If
            txtTransRate2.Enabled = False
            cboTransUnit2.Enabled = False
            txtTransRate3.Enabled = False
            cboTransUnit3.Enabled = False
            txtTransRate4.Enabled = False
            cboTransUnit4.Enabled = False
            txtTransRate5.Enabled = False
            cboTransUnit5.Enabled = False
            txtTransRate6.Enabled = False
            cboTransUnit6.Enabled = False
            txtTransRate7.Enabled = False
            cboTransUnit7.Enabled = False
            txtTransRate8.Enabled = False
            cboTransUnit8.Enabled = False
            txtTransRate9.Enabled = False
            cboTransUnit9.Enabled = False
            txtTransRate2.Text = "0"
            txtTransRate3.Text = "0"
            txtTransRate4.Text = "0"
            txtTransRate5.Text = "0"
            txtTransRate6.Text = "0"
            txtTransRate7.Text = "0"
            txtTransRate8.Text = "0"
            txtTransRate9.Text = "0"
    isSave = False
    cboPrintForm.Enabled = True
    PopulateCompany cboCompany, cboIndex
    With cboPrintForm
        .AddItem "พิมพ์เอกสารฟอร์มทั่วไป"
        .AddItem "พิมพ์เอกสารฟอร์มส่งเผาเตา"
        .ListIndex = 0
    End With
    grItem.Visible = True
    grItemAP.Visible = False
    With cboTransUnit
        .AddItem "บาท/เที่ยว"
    End With
    With cboTransUnit2
        .AddItem "บาท/เที่ยว"
    End With
    With cboTransUnit3
        .AddItem "บาท/เที่ยว"
    End With
    With cboTransUnit4
        .AddItem "บาท/เที่ยว"
    End With
    With cboTransUnit5
        .AddItem "บาท/เที่ยว"
    End With
    With cboTransUnit6
        .AddItem "บาท/เที่ยว"
    End With
    With cboTransUnit7
        .AddItem "บาท/เที่ยว"
    End With
    With cboTransUnit8
        .AddItem "บาท/เที่ยว"
    End With
     With cboTransUnit9
        .AddItem "บาท/เที่ยว"
    End With
    cboTransUnit.ListIndex = 0
    With cboQuotationType
        .AddItem "ใบเสนอราคาแบบ ปกติ (Routine)"
        .AddItem "ใบเสนอราคาแบบ Jobs"
        .AddItem "งาน DOC (Document Other)"
'        For x = 1 To WasteCases.Count
'            .AddItem WasteCases(x).CaseName
'        Next
        .ListIndex = 0
    End With
    With cboMaxRateType
        .AddItem "บาท/ตัน"
        .AddItem "บาท/เที่ยว"
        .ListIndex = -1
    End With
    cboIsTrailer.ListIndex = 0
    With frmPicture
        cmdSave.Picture = .picSave.Picture
        cmdSearch.Picture = .picSearch.Picture
        cmdClear.Picture = .picNew.Picture
        cmdCancel.Picture = .picDelete.Picture
        cmdClose.Picture = .picExit.Picture
        cmdPrint.Picture = .picPrint
    End With
    With cboApprove
'        .AddItem SystemConfig.SignName1
'        If Trim(SystemConfig.SignName2) <> "" Then .AddItem SystemConfig.SignName2
'        If Trim(SystemConfig.SignName3) <> "" Then .AddItem SystemConfig.SignName3
'        If Trim(SystemConfig.SignName4) <> "" Then .AddItem SystemConfig.SignName4
'        If Trim(SystemConfig.SignName5) <> "" Then .AddItem SystemConfig.SignName5
        Dim tmpS As New BWGSearchManager
        Set SignName = tmpS.SignNameSearch("")
        For i = 1 To SignName.Count
            .AddItem SignName.item(i).SignName1
        Next
    End With
    cboApprove.ListIndex = 0
    PopulateUnit
    txtIssuedStaff.Text = CurrentUser.StaffFName & " " & CurrentUser.StaffLName
    Label2(4).Caption = "==สถานะการอนุมัติ=="
    cboDiscountType.AddItem "%"
    cboDiscountType.AddItem "บาท"
    cboDiscountType.ListIndex = 0
    Set curTumbol_ = Nothing
    Set curTumbol_ = New IMITumbol
    grItem.ColInputType(4) = ComboBoxEnum
    grItem.ColInputType(6) = ComboBoxEnum
    grItem.ColInputType(9) = ComboBoxEnum
    grItem.GridMoveType 1
    
    With grItem.getGridObj
        .Cols = 11
        .ColWidth(0) = 0
        .ColWidth(1) = 1200
        .ColWidth(2) = 5000
        .ColWidth(3) = 0
        .ColWidth(4) = 0
        .ColWidth(5) = 1200
        .ColWidth(6) = 1065
        .ColWidth(7) = 2070
        .ColWidth(8) = 1200
        .ColWidth(9) = 1200
         .ColWidth(10) = 1200
        .Rows = 20
        .FixedAlignment(1) = 4
        .FixedAlignment(2) = 4
        .FixedAlignment(3) = 4
        .FixedAlignment(4) = 4
        .FixedAlignment(5) = 4
        .FixedAlignment(6) = 4
        .FixedAlignment(7) = 4
        .FixedAlignment(8) = 4
        .FixedAlignment(9) = 4
        .FixedAlignment(10) = 4

        .ColAlignment(1) = 1
        .ColAlignment(2) = 1
        .ColAlignment(3) = 7
        .ColAlignment(4) = 1
        .ColAlignment(5) = 7
        .ColAlignment(6) = 1
        .ColAlignment(7) = 1
        .ColAlignment(8) = 1
        .ColAlignment(9) = 1
        .ColAlignment(10) = 1

        .TextMatrix(0, 1) = "รหัสของเสีย"
        .TextMatrix(0, 2) = "ชื่อกากของเสีย"
        .TextMatrix(0, 3) = "ปริมาณ"
        .TextMatrix(0, 4) = "หน่วย"
        .TextMatrix(0, 5) = "อัตราค่าบำบัด"
        .TextMatrix(0, 6) = "หน่วย"
        .TextMatrix(0, 7) = "หมายเหตุ"
        .TextMatrix(0, 8) = "ต้นทุน"
        .TextMatrix(0, 9) = "หน่วย"
        .TextMatrix(0, 10) = "ปลายทาง"
    End With
    
    grItemAP.ColInputType(4) = ComboBoxEnum
    grItemAP.GridMoveType 1
    With grItemAP.getGridObj
        .Cols = 8
        .ColWidth(0) = 0
        .ColWidth(1) = 1200
        .ColWidth(2) = 3150
        .ColWidth(3) = 1600
        .ColWidth(4) = 1065
        .ColWidth(5) = 2000
        .ColWidth(6) = 2000
        .ColWidth(7) = 1200
        .Rows = 20
        .FixedAlignment(1) = 4
        .FixedAlignment(2) = 4
        .FixedAlignment(3) = 4
        .FixedAlignment(4) = 4
        .FixedAlignment(5) = 4
        .FixedAlignment(6) = 4
        .FixedAlignment(7) = 4

        .ColAlignment(1) = 1
        .ColAlignment(2) = 1
        .ColAlignment(3) = 7
        .ColAlignment(4) = 1
        .ColAlignment(5) = 7
        .ColAlignment(6) = 7
        .ColAlignment(7) = 7

        .TextMatrix(0, 1) = "รหัสกาก"
        .TextMatrix(0, 2) = "ชื่อกากของเสีย"
        .TextMatrix(0, 3) = "ค่าบริการเผาทำลาย"
        .TextMatrix(0, 4) = "หน่วย"
        .TextMatrix(0, 5) = "ค่าบริการปรับปรุงคุณภาพ"
        .TextMatrix(0, 6) = "ค่าเกินเกณฑ์การออกแบบ"
        .TextMatrix(0, 7) = "รวมค่าบริการ"
    End With
    With grRemark.getGridObj
        .Cols = 2
        .ColWidth(0) = 0
        .ColWidth(1) = 10800
        .Rows = 20
        .FixedAlignment(1) = 4

        .ColAlignment(1) = 1

        .TextMatrix(0, 1) = "หมายเหตุ"
    End With
    SetGrCar

    CreateNewQuotation
    If Day(Now) < 10 Then
        lblD.Caption = "0" & Day(Now)
    Else
        lblD.Caption = Day(Now)
    End If
    If Month(Now) < 10 Then
        lblM.Caption = "0" & Month(Now)
    Else
        lblM.Caption = Month(Now)
    End If
    If Year(Now) <= "2500" Then
        lblY.Caption = Val(Year(Now)) + 543
    Else
        lblY.Caption = Year(Now)
    End If
    
End Sub
Sub SetGrCar()
    'grCar.ColInputType(2) = ComboBoxEnum
    GrCar.GridMoveType 1
        With GrCar.getGridObj
        .Clear
        .Cols = 4
        .ColWidth(0) = 0
        .ColWidth(1) = 6000
        .ColWidth(2) = 2000
        .ColWidth(3) = 6000
        '.ColWidth(4) = 2000
        .Rows = 2
        .FixedAlignment(1) = 4
        .FixedAlignment(2) = 4
        .FixedAlignment(3) = 4
        '.FixedAlignment(4) = 4

        .ColAlignment(1) = 1
        .ColAlignment(2) = 1
        .ColAlignment(3) = 1
        '.ColAlignment(4) = 1

        .TextMatrix(0, 1) = "ประเภทรถ"
        .TextMatrix(0, 2) = "ค่าขนส่ง"
         .TextMatrix(0, 3) = "ที่อยู่"
        '.TextMatrix(0, 3) = "ค่าขนส่ง"
        '.TextMatrix(0, 4) = "ขั้นต่ำ(ตัน/คัน)"
        End With
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 And Not TypeOf Me.ActiveControl Is ctlGrid Then SendKeys "{Tab}"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim selResult As Integer
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole(Me.Name, curRole_)
    If Trim(tmpStr) <> "" Then
        Exit Sub
    End If
    If myQuotation_ Is Nothing Then Exit Sub
    If myQuotation_.isChanged And cmdSave.Enabled Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนปิดหน้าจอหรือไม่?", vbQuestion + vbYesNoCancel)
        If selResult = vbYes Then
            CmdSave_Click
            If myQuotation_.isChanged Then Cancel = 1
        ElseIf selResult = vbCancel Then
            Cancel = 1
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set myQuotation_ = Nothing
Set curCustomer_ = Nothing
End Sub

Private Sub grCar_BeforeUpdateValue(NewValue As String, row As Integer, col As Integer, Cancel As Boolean)
    Dim tmpKey As String
    Dim x%
    tmpKey = "'" & GrCar.getGridObj.TextMatrix(row, 0) & "'"
    If col = 2 Then
            If isExist(myQuotation_.QuotationCar, tmpKey) Then
                myQuotation_.QuotationCar(tmpKey).Rate = NewValue
            End If
    End If
'    If col = 3 Or col = 4 Then
'        If NewValue <> "" Then
'            If IsNumeric(NewValue) = False Then
'                MsgBox "ข้อมูลไม่ถูกต้อง ต้องเป็นตัวเลขเท่านั้น", vbExclamation
'                Cancel = True
'            End If
'        End If
'        myQuotation_.isChanged = True
'    End If

End Sub

Private Sub grCar_KeyDown(KeyCode As Integer, Shift As Integer)
'Dim selResult As Integer
'Dim x%, tmpKey$, itemCount%
'Dim selCol As Collection
'Dim tSql As String
'Dim TrID As String
'    With GrCar.getGridObj
'        If KeyCode = 46 And Trim(.TextMatrix(.row, 0)) <> "" Then
'            selResult = MsgBox("ต้องการลบรายการรถนี้ออกจากใบเสนอราคาหรือไม่?", vbQuestion + vbYesNo)
'            If selResult = vbYes Then
'            TrID = .TextMatrix(.row, 0)
'                tmpKey = "'" & .TextMatrix(.row, 0) & "'"
'                .RemoveItem .row
'                        tSql = "delete tbSelTruck Where QuotationID = '" & myQuotation_.id & "' and TruckTypeID='" & TrID & "' "
'                        DBConnExc tSql
'
'                myQuotation_.isChanged = True
'            End If
'        End If
'    End With

End Sub

Private Sub grCar_RolColChanged(row As Integer, col As Integer)
Dim x%
Dim tmpCbo As ComboBox
Dim tmpKey As String
    tmpKey = "'" & GrCar.getGridObj.TextMatrix(row, 0) & "'"
    If Trim(tmpKey) = "''" Then
        GrCar.ColEnabled(col) = False
    End If
    If isExist(myQuotation_.QuotationCar, tmpKey) Then
    End If
'    If isExist(myQuotation_.QuotationDetails, tmpKey) Then
'        If col = 1 Or col = 2 Or col = 3 Or col = 4 Or col = 5 Or col = 6 Or col = 7 Then
'            grItemAP.ColEnabled(col) = True
'        Else
'            grItemAP.ColEnabled(col) = False
'        End If
'    End If
'    If col = 2 Then
'        Set tmpCbo = grCar.getColObject(col)
'        tmpCbo.Clear
'        With tmpCbo
'            .AddItem "รถเดี่ยว"
'            .AddItem "รถพ่วง"
'        End With
'    End If

End Sub

Private Sub grCar_UpdateValue(row As Integer, col As Integer)
        myQuotation_.isChanged = True
End Sub

Private Sub grItemAP_BeforeUpdateValue(NewValue As String, row As Integer, col As Integer, Cancel As Boolean)
    If col = 3 Or col = 5 Or col = 6 Then
        If NewValue <> "" Then
            If IsNumeric(NewValue) = False Then
                MsgBox "ข้อมูลไม่ถูกต้อง ต้องเป็นตัวเลขเท่านั้น", vbExclamation
                Cancel = True
            End If
        End If
    End If
End Sub

Private Sub grItemAP_Click(row As Integer, col As Integer)
'    MsgBox grItemAP.getGridObj.ColWidth(col)
End Sub

Private Sub grItemAP_KeyDown(KeyCode As Integer, Shift As Integer)
Dim selResult As Integer
Dim x%, tmpKey$, itemCount%
Dim selCol As Collection
    With grItemAP.getGridObj
        If KeyCode = 46 And Trim(.TextMatrix(.row, 0)) <> "" Then
            selResult = MsgBox("ต้องการลบรายการกากของเสียนี้ออกจากใบเสนอราคาหรือไม่?", vbQuestion + vbYesNo)
            If selResult = vbYes Then
                tmpKey = "'" & .TextMatrix(.row, 0) & "'"
                .RemoveItem .row
                Set selCol = myQuotation_.QuotationDetails
                selCol.Remove tmpKey

                itemCount = myQuotation_.QuotationDetails.Count
                For x = 1 To .Rows - 1
                    tmpKey = "'" & .TextMatrix(x, 0) & "'"
                    If isExist(myQuotation_.QuotationDetails, tmpKey) Then
                        myQuotation_.QuotationDetails(tmpKey).RowPos = x
                    End If
                Next
                myQuotation_.isChanged = True
            End If
        End If
    End With
End Sub

Private Sub grItemAP_RolColChanged(row As Integer, col As Integer)
Dim x%
Dim tmpCbo As ComboBox
Dim tmpKey As String
    tmpKey = "'" & grItemAP.getGridObj.TextMatrix(row, 0) & "'"
    If Trim(tmpKey) = "''" Then
        grItemAP.ColEnabled(col) = False
    End If
    If isExist(myQuotation_.QuotationDetails, tmpKey) Then
        If col = 1 Or col = 2 Or col = 3 Or col = 4 Or col = 5 Or col = 6 Or col = 7 Then
            grItemAP.ColEnabled(col) = True
        Else
            grItemAP.ColEnabled(col) = False
        End If
    End If
    If col = 4 Then
        Set tmpCbo = grItemAP.getColObject(col)
        tmpCbo.Clear
        With tmpCbo
            .AddItem "บาท/เที่ยว"
            .AddItem "บาท/ตัน"
            .AddItem "บาท/ถัง"
            .AddItem "บาท/ลบ.ม."
            .AddItem "บาท/กล่อง"
            .AddItem "บาท/งาน"
            .AddItem "บาท/กิโลกรัม"
            .AddItem "บาท/ถุง"
            .AddItem "บาท/ขวด"
            .AddItem "บาท/แล็ก"
            .AddItem "บาท/กระป๋อง"
            .AddItem "บาท/หลอด"
            .AddItem "บาท/แกลอน"
            .AddItem "บาท/ลูก"
        End With
    End If
End Sub

Private Sub grItemAP_UpdateValue(row As Integer, col As Integer)
Dim tmpDet As BWGQuotationDetail
Dim tmpKey As String
Dim tmpVal As String
    tmpKey = Trim(grItemAP.getGridObj.TextMatrix(row, 0))
    Set tmpDet = myQuotation_.QuotationDetails("'" & tmpKey & "'")
    tmpVal = Trim(grItemAP.getGridObj.TextMatrix(row, col))
    With tmpDet
        Select Case col
            Case 2
                .BillDesc = Trim(tmpVal)
            Case 3
                If IsNumeric(Trim(tmpVal)) Then
                    .TreatmentRate = CDbl(Trim(tmpVal))
                Else
                    .TreatmentRate = 0
                End If
            Case 4
                If Trim(tmpVal) <> "" Then
                    .TreatmentUnitID = Trim(tmpVal)
                    .TreatmentUnitName = Trim(tmpVal)
                Else
                    .TreatmentUnitID = ""
                    .TreatmentUnitName = ""
                End If
            Case 5
                .BlendingRate = tmpVal
            Case 6
                .ParameterRate = tmpVal
            
        End Select
        grItemAP.getGridObj.TextMatrix(row, 7) = CDbl(grItemAP.getGridObj.TextMatrix(row, 3)) + CDbl(grItemAP.getGridObj.TextMatrix(row, 5)) + CDbl(grItemAP.getGridObj.TextMatrix(row, 6))
        myQuotation_.isChanged = True
    End With
    Set tmpDet = Nothing
End Sub

Private Sub grRemark_KeyDown(KeyCode As Integer, Shift As Integer)
Dim selResult As Integer
Dim x%, tmpKey$, itemCount%
Dim selCol As Collection
    With grRemark.getGridObj
        If KeyCode = 46 And Trim(.TextMatrix(.row, 0)) <> "" Then
            selResult = MsgBox("ต้องการลบรายการหมายเหตุนี้ออกจากใบเสนอราคาหรือไม่?", vbQuestion + vbYesNo)
            If selResult = vbYes Then
                tmpKey = "'" & .TextMatrix(.row, 0) & "'"
                .RemoveItem .row
                Set selCol = myQuotation_.RemarkDetails
                selCol.Remove tmpKey

                itemCount = myQuotation_.RemarkDetails.Count
                For x = 1 To .Rows - 1
                    tmpKey = "'" & .TextMatrix(x, 0) & "'"
                    If isExist(myQuotation_.RemarkDetails, tmpKey) Then
                        myQuotation_.RemarkDetails(tmpKey).ItemNo = x
                    End If
                Next
                myQuotation_.isChanged = True
            End If
        End If
    End With
End Sub

Private Sub grRemark_UpdateValue(row As Integer, col As Integer)
Dim tmpDet As BWGQuotationRem
Dim tmpKey As String
Dim tmpVal As String
    tmpKey = Trim(grRemark.getGridObj.TextMatrix(row, 0))
    If isExist(myQuotation_.RemarkDetails, "'" & tmpKey & "'") = False Then
        Set tmpDet = New BWGQuotationRem
        tmpDet.id = GetGUID
        tmpDet.QuotationID = myQuotation_.id
        tmpDet.ItemNo = row
        grRemark.getGridObj.TextMatrix(row, 0) = tmpDet.id
        myQuotation_.RemarkDetails.Add tmpDet, "'" & tmpDet.id & "'"
    Else
        Set tmpDet = myQuotation_.RemarkDetails("'" & tmpKey & "'")
        myQuotation_.isChanged = True
    End If
    tmpVal = Trim(grRemark.getGridObj.TextMatrix(row, col))
    With tmpDet
        Select Case col
            Case 1
                .Remdesc = Trim(tmpVal)
                .ItemNo = row
        End Select
        myQuotation_.isChanged = True
    End With
    Set tmpDet = Nothing
End Sub

Private Sub txtBMEQuotation_Validate(Cancel As Boolean)
    myQuotation_.BMEQuotation = Trim(txtBMEQuotation.Text)
End Sub

Private Sub txtCarType_Validate(Cancel As Boolean)
    If curCarType_ Is Nothing Then Set curCarType_ = New BWGTruckType
    ObjSel = "1"
    If Trim(txtCarType.Text) = "" And Trim(curCarType_.TypeDesc) <> Trim(txtCarType.Text) Then
        Set curCarType_ = Nothing
        Set curCarType_ = New BWGTruckType
        With myQuotation_
            .TruckTypeID = ""
            .TruckTypeName = ""
        End With
    Else
        If Trim(curCarType_.TypeDesc) <> Trim(txtCarType.Text) Then
            Set curCarType_ = Nothing
            Set curCarType_ = New BWGTruckType
            With myQuotation_
                .TruckTypeID = ""
                .TruckTypeName = ""
            End With
            Load frmTruckType
            With frmTruckType
                '.getCriteria (txtCarType.Text)
                .SetCallerForm Me
                .Show 1
            End With
            If myQuotation_.TruckTypeID = "" Then txtCarType.Text = ""
        End If
    End If
End Sub

Private Sub txtCarType2_Validate(Cancel As Boolean)
'    If curCarType_ Is Nothing Then Set curCarType_ = New BWGTruckType
'        ObjSel = "2"
'    If Trim(txtCarType2.Text) = "" And Trim(curCarType_.TypeDesc) <> Trim(txtCarType2.Text) Then
'        Set curCarType_ = Nothing
'        Set curCarType_ = New BWGTruckType
'        With myQuotation_
'            .TRuck2ID = ""
'        End With
'    Else
'        If Trim(curCarType_.TypeDesc) <> Trim(txtCarType2.Text) Then
'            Set curCarType_ = Nothing
'            Set curCarType_ = New BWGTruckType
'            With myQuotation_
'                .TRuck2ID = ""
'            End With
'            Load frmTruckType
'            With frmTruckType
'                '.getCriteria (txtCarType.Text)
'                .SetCallerForm Me
'                .Show 1
'            End With
'            If myQuotation_.TRuck2ID = "" Then txtCarType2.Text = ""
'        End If
'    End If

End Sub

Private Sub txtCarType3_Validate(Cancel As Boolean)
'    If curCarType_ Is Nothing Then Set curCarType_ = New BWGTruckType
'        ObjSel = "3"
'    If Trim(txtCarType3.Text) = "" And Trim(curCarType_.TypeDesc) <> Trim(txtCarType3.Text) Then
'        Set curCarType_ = Nothing
'        Set curCarType_ = New BWGTruckType
'        With myQuotation_
'            .TRuck3ID = ""
'        End With
'    Else
'        If Trim(curCarType_.TypeDesc) <> Trim(txtCarType3.Text) Then
'            Set curCarType_ = Nothing
'            Set curCarType_ = New BWGTruckType
'            With myQuotation_
'                .TRuck3ID = ""
'            End With
'            Load frmTruckType
'            With frmTruckType
'                '.getCriteria (txtCarType.Text)
'                .SetCallerForm Me
'                .Show 1
'            End With
'            If myQuotation_.TRuck3ID = "" Then txtCarType3.Text = ""
'        End If
'    End If
'
End Sub
Private Sub txtCarType4_Validate(Cancel As Boolean)
'    If curCarType_ Is Nothing Then Set curCarType_ = New BWGTruckType
'        ObjSel = "4"
'    If Trim(txtCarType4.Text) = "" And Trim(curCarType_.TypeDesc) <> Trim(txtCarType4.Text) Then
'        Set curCarType_ = Nothing
'        Set curCarType_ = New BWGTruckType
'        With myQuotation_
'            .TRuck4ID = ""
'        End With
'    Else
'        If Trim(curCarType_.TypeDesc) <> Trim(txtCarType4.Text) Then
'            Set curCarType_ = Nothing
'            Set curCarType_ = New BWGTruckType
'            With myQuotation_
'                .TRuck4ID = ""
'            End With
'            Load frmTruckType
'            With frmTruckType
'                '.getCriteria (txtCarType.Text)
'                .SetCallerForm Me
'                .Show 1
'            End With
'            If myQuotation_.TRuck4ID = "" Then txtCarType4.Text = ""
'        End If
'    End If

End Sub
Private Sub txtCarType5_Validate(Cancel As Boolean)
'    If curCarType_ Is Nothing Then Set curCarType_ = New BWGTruckType
'        ObjSel = "5"
'    If Trim(txtCarType5.Text) = "" And Trim(curCarType_.TypeDesc) <> Trim(txtCarType5.Text) Then
'        Set curCarType_ = Nothing
'        Set curCarType_ = New BWGTruckType
'        With myQuotation_
'            .TRuck5ID = ""
'        End With
'    Else
'        If Trim(curCarType_.TypeDesc) <> Trim(txtCarType5.Text) Then
'            Set curCarType_ = Nothing
'            Set curCarType_ = New BWGTruckType
'            With myQuotation_
'                .TRuck5ID = ""
'            End With
'            Load frmTruckType
'            With frmTruckType
'                '.getCriteria (txtCarType.Text)
'                .SetCallerForm Me
'                .Show 1
'            End With
'            If myQuotation_.TRuck5ID = "" Then txtCarType5.Text = ""
'        End If
'    End If

End Sub

Private Sub txtDiscount_Change()
    If Me.Visible Then
        If Me.ActiveControl.Name = txtDiscount.Name Then
            If IsNumeric(Trim(txtDiscount.Text)) And Trim(txtDiscount.Text) <> "" Then
                myQuotation_.DiscountAmt = CSng(Trim(txtDiscount.Text))
            Else
                myQuotation_.DiscountAmt = 0
            End If
        End If
    End If
End Sub

Private Sub txtCustomer_Validate(Cancel As Boolean)
    If curCustomer_ Is Nothing Then Set curCustomer_ = New BWGCustomer
    If Trim(txtCustomer.Text) = "" And Trim(curCustomer_.CustomerName) <> Trim(txtCustomer.Text) Then
        Set curCustomer_ = Nothing
        Set curCustomer_ = New BWGCustomer
        With myQuotation_
            .CustomerID = ""
            .CustomerName = ""
        End With
    Else
        If Trim(curCustomer_.CustomerName) <> Trim(txtCustomer.Text) Then
            Set curCustomer_ = Nothing
            Set curCustomer_ = New BWGCustomer
            With myQuotation_
                .CustomerID = ""
                .CustomerName = ""
            End With
            Load frmCustomerSearch
            With frmCustomerSearch
                .CStatus = "N"
                .getCriteria (txtCustomer.Text)
                .GetCriteriaOptonal " AND CustGroupID <> '488F2E5C057F46869D1A56185B11D644'"
                .SetCallerForm Me
                .Show 1
            End With
            If myQuotation_.CustomerID = "" Then txtCustomer.Text = ""
        End If
    End If
    ProjectQuo
End Sub

Private Sub CreateNewQuotation()
Set myQuotation_ = Nothing
Set myQuotation_ = New BWGQuotationHeader
    curRole_ = AddNewRole
    With myQuotation_
        .id = GetGUID
        If cboCompany.ListCount > 1 Then
            If cboCompany.ListIndex > 0 Then
                .CompanyID = CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).id
            End If
        Else
            If cboCompany.ListIndex >= 0 Then
                .CompanyID = CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).id
            End If
        End If
        If Not curCustomer_ Is Nothing Then
            .CustomerID = curCustomer_.id
            .CustomerName = curCustomer_.CustomerName
        End If
        .IssuedDate = dtIssuedDate.ValueYMD
        .isApproved = False
        .isCancel = False
        .IssuedStaffID = CurrentUser.id
        .IssuedStaffName = CurrentUser.StaffFName & " " & CurrentUser.StaffLName
        If cboIsTrailer.ListIndex = 0 Then
            .isTrailer = False
        Else
            .isTrailer = True
        End If
        .isChanged = False
    End With
    SetEnabledScreen
End Sub

Private Sub ClearScreen()
    txtQuoNo.Text = "== AUTO =="
    txtCustomer.Text = ""
    dtIssuedDate.ValueYMD = TodayDate
    cboQuotationType.ListIndex = 0
    txtIssuedStaff.Text = CurrentUser.StaffFName & " " & CurrentUser.StaffLName
    cboCustStaff.Text = ""
    chkIncVAT.Value = 0
    txtExtraCharge.Text = ""
    txtCarType.Text = ""
    Set curCarType_ = Nothing
    Set curCarType_ = New BWGTruckType
    Set curTumbol_ = Nothing
    Set curTumbol_ = New IMITumbol
    txtTumbol.Text = ""
    txtAumphur.Text = ""
    txtProvince.Text = ""
    txtTransRate.Text = ""
    cboTransUnit.ListIndex = -1
    'cboApprove.Text = ""
    txtRem.Text = ""
    txtPosition.Text = ""
    txtRefQuotationID.Text = ""
    cboIsTrailer.ListIndex = 0
    cboIsTrailer.Enabled = False
    If cboCompany.ListCount > 0 Then
        cboCompany.ListIndex = 0
    End If
    txtApprovedStaff.Text = ""
    txtLaborCount.Text = ""
    chkIncTrans.Value = 0
    chkTrnDis.Value = 0
    txtMinQty.Text = ""
    txtNonRate.Text = ""
    txtHazRate.Text = ""
    txtNBLRate.Text = ""
    txtHBLRate.Text = ""
    lbCancel.Visible = False
    Label2(4).Caption = "==สถานะการอนุมัติ=="
    grItemAP.ClearAllData
    grItemAP.getGridObj.Rows = 20
    grRemark.ClearAllData
    grRemark.getGridObj.Rows = 20
    GrCar.ClearAllData
    GrCar.getGridObj.Rows = 2
    
    txtMaxWeight.Text = ""
    txtMaxRate.Text = ""
    cboMaxRateType.ListIndex = -1
    
    Set curCustomer_ = Nothing
    Set curCustomer_ = New BWGCustomer
End Sub

Public Sub SetDataQuotation(selItem As BWGQuotationHeader)
On Error Resume Next
Dim tmpS As New BWGSearchManager
Dim tmpCol As Collection
Dim tmpItem As IMITumbol
Dim i As Integer
    curRole_ = UpdateRole
    isSave = True
    cboPrintForm.Enabled = False
    
    Set myQuotation_ = Nothing
    Set myQuotation_ = selItem
    If myQuotation_ Is Nothing Then Exit Sub
    With myQuotation_
    
        Dim sqlSel As String
        Dim tmpSel As ADODB.Recordset
        Set tmpSel = New ADODB.Recordset
        sqlSel = " select * from tbSelTruck  Where QuotationID = '" & myQuotation_.id & "' "
        Set tmpSel = GetRS(sqlSel, adOpenForwardOnly, adLockReadOnly)
        If tmpSel.EOF = False Then
            txtQuoNo.Text = .QuotationNo & "(CR)"
        Else
            txtQuoNo.Text = .QuotationNo
        End If
        txtBMEQuotation.Text = .BMEQuotation
        
        Dim sqlPro As String
        Dim tmpPro As ADODB.Recordset
        Set tmpPro = New ADODB.Recordset
        sqlPro = " select * from tbJobProject  Where ProjectID = '" & .ProjectID & "' "
        Set tmpPro = GetRS(sqlPro, adOpenForwardOnly, adLockReadOnly)
        If tmpPro.EOF = False Then
            cbProject.Text = tmpPro!ProjectName & " # " & tmpPro!ProjectID
        End If
        If .TransRate = "F" Then
            ChkFreeTrn.Value = 1
            ElseIf .TransRate = "H" Then
            ChkHalfTrn.Value = 1
            Frame4.Visible = True
        Else
            ChkFreeTrn.Value = 0
            ChkHalfTrn.Value = 0
            Frame4.Visible = False
        End If
        
        dtIssuedDate.ValueYMD = .IssuedDate
        Set tmpCol = tmpS.CustomerSearch("CustomerID='" & .CustomerID & "'")
        If tmpCol.Count > 0 Then SetDataCustomer tmpCol(1), False
        txtIssuedStaff.Text = .IssuedStaffName
        txtLastUser.Text = .LastUpdateUser
        cboCustStaff.Text = .AttentionTo
        If .IsPriceIncVAT Then
            chkIncVAT.Value = 1
        Else
            chkIncVAT.Value = 0
        End If
        If isExist(CurrentUser.WorkCompanies, "'" & .CompanyID & "'") Then
            SetTextToCombo CurrentUser.WorkCompanies("'" & .CompanyID & "'").CompanyNameTH, cboCompany
        End If
        txtRem.Text = .Remdesc
        txtExtraCharge.Text = .ExtraLaborFee
        txtApprovedStaff.Text = .ApprovedStaffName
        If .StaffCountPerCar > 0 Then txtLaborCount.Text = .StaffCountPerCar
        If .IsPriceIncTransport Then
            chkIncTrans.Value = 1
        Else
            chkIncTrans.Value = 0
        End If
         If .IsDiscountPriceTrans Then
            chkTrnDis.Value = 1
        Else
            chkTrnDis.Value = 0
        End If
        If .isBmeNew = "Y" Then
            chkBmeNew.Value = 1
        Else
             chkBmeNew.Value = 0
        End If
        If .isBusBT = "Y" Then
            chkcarbt.Value = 1
        Else
            chkcarbt.Value = 0
        End If
        txtMinQty.Text = .MinWeightPerCar
        txtTransRate.Text = .TransferRate
        If .NonMinRate > 0 Then
            txtNonRate.Text = .NonMinRate
        Else
            txtNonRate.Text = ""
        End If
        If .HazMinRate > 0 Then
            txtHazRate.Text = .HazMinRate
        Else
            txtHazRate.Text = ""
        End If
        If .NBLMinRate > 0 Then
            txtNBLRate.Text = .NBLMinRate
        Else
            txtNBLRate.Text = ""
        End If
        If .HBLMinRate > 0 Then
            txtHBLRate.Text = .HBLMinRate
        Else
            txtHBLRate.Text = ""
        End If
        SetTextToCombo .TransferUnitName, cboTransUnit
        Set tmpCol = tmpS.TruckTypeSearch("TruckTypeID='" & .TruckTypeID & "'")
        If tmpCol.Count > 0 Then
            SetDataTruckType tmpCol(1), False
        End If
        If .TumbolID <> "" Then
        tmpItem = New IMITumbol
            Set tmpCol = tmpS.TumbolSearch("TumbolID='" & .TumbolID & "'")
             If tmpCol.Count > 0 Then
                    Set tmpItem = tmpCol(1)
                    .TumbolID = tmpItem.id
                    .TumbolName = tmpItem.Name
                    txtTumbol.Text = tmpItem.Name
                    .AumphurID = tmpItem.AumphurID
                    .AumphurName = tmpItem.AumphurName
                    txtAumphur.Text = tmpItem.AumphurName
                    .ProvinceID = tmpItem.ProvinceID
                    .ProvinceName = tmpItem.ProvinceName
                    txtProvince.Text = tmpItem.ProvinceName
            End If
        End If

        If .isTrailer Then
            cboIsTrailer.ListIndex = 1
        Else
            cboIsTrailer.ListIndex = 0
        End If
        If .isApproved = True Then
            Label2(4).Caption = "อนุมัติแล้วเมื่อวันที่ " & FormatYMD_to_DMY(myQuotation_.Approveddate, "/", "/") & " โดย " & myQuotation_.ApprovedStaffName
        Else
            Label2(4).Caption = "รายการรออนุมัติ"
        End If
        If .isClosed Then
            lbCancel.Caption = "Closed"
            lbCancel.Visible = True
        ElseIf .isCancel Then
            lbCancel.Caption = "ยกเลิก"
            lbCancel.Visible = True
        ElseIf .isApproved Then
            lbCancel.Caption = "Approved"
            lbCancel.Visible = True
        Else
            lbCancel.Visible = False
        End If

        cboApprove.Text = .SignName
        txtPosition.Text = .SignPosition
        SetEnabledScreen
        QuotationPromies = .Promies
        QuotationStartDatePromies = .StartPromiesDate
        QuotationEndDatePromies = .EndPromiesDate

        If CmdQuoTruck.Caption <> "ใบเสนอราคาสำหรับเลือกรถ" Then
            txtTransRate.Text = .Truck1Rate
            txtTransRate2.Text = .Truck2Rate
            txtTransRate3.Text = .TRuck3Rate
            txtTransRate4.Text = .TRuck4Rate
            txtTransRate5.Text = .TRuck5Rate
            txtTransRate6.Text = .TRuck6Rate
            txtTransRate7.Text = .TRuck7Rate
            txtTransRate8.Text = .TRuck8Rate
            txtTransRate9.Text = .TRuck9Rate
            txtCarType.Text = "รถ Lugger Box"
            cboIsTrailer.ListIndex = 0
            txtCarType2.Text = "รถ Lugger Box - พ่วง"
            cboIsTrailer2.Text = "รถพ่วง"
            txtCarType3.Text = "รถ Roll off Truck - เดี่ยว"
            cboIsTrailer3.Text = "รถเดี่ยว"
            txtCarType4.Text = "รถ Roll off Truck - พ่วง"
            cboIsTrailer4.Text = "รถพ่วง"
            txtCarType5.Text = "รถกระบะ"
            cboIsTrailer5.Text = "รถเดี่ยว"
            txtCarType6.Text = "รถเทรลเลอร์"
            cboIsTrailer6.Text = "รถเดี่ยว"
            txtCarType7.Text = "รถแท็งค์ 15 ลูกบาศก์เมตร"
            cboIsTrailer7.Text = "รถเดี่ยว"
            txtCarType8.Text = "รถแท็งค์ ISO 24 ลูกบาศก์เมตร"
            cboIsTrailer8.Text = "รถเดี่ยว"
            txtCarType9.Text = "รถโรงงาน"
            cboTransUnit.Text = .Truck1Unit
            cboTransUnit2.Text = .Truck2Unit
            cboTransUnit3.Text = .TRuck3Unit
            cboTransUnit4.Text = .TRuck4Unit
            cboTransUnit5.Text = .TRuck5Unit
            cboTransUnit6.Text = .TRuck6Unit
            cboTransUnit7.Text = .TRuck7Unit
            cboTransUnit8.Text = .TRuck8Unit
            cboTransUnit9.Text = .TRuck9Unit
        Else
            txtTransRate.Text = .TransferRate
            txtCarType.Text = .TruckTypeName
            cboIsTrailer.Text = .isTrailer
            cboTransUnit.Text = .TransferUnitID
        End If

        If .MaxWeight = 0 Then
            txtMaxWeight.Text = ""
        Else
            txtMaxWeight.Text = .MaxWeight
        End If
        If .MaxRate = 0 Then
            txtMaxRate.Text = ""
        Else
            txtMaxRate.Text = .MaxRate
        End If
        If .MaxRateType = "" Then
            cboMaxRateType.ListIndex = -1
        ElseIf .MaxRateType = "บาท/ตัน" Then
            cboMaxRateType.ListIndex = 0
        ElseIf .MaxRateType = "บาท/เที่ยว" Then
            cboMaxRateType.ListIndex = 1
        End If
    End With
    If myQuotation_.QuotationType <= 0 Then
        cboPrintForm.ListIndex = 0
    Else
        cboPrintForm.ListIndex = myQuotation_.QuotationType - 1
    End If
       If myQuotation_.isJob = "N" Then
            cboQuotationType.ListIndex = 0
        ElseIf myQuotation_.isJob = "Y" Then
            cboQuotationType.ListIndex = 1
        ElseIf myQuotation_.isJob = "D" Then
            cboQuotationType.ListIndex = 2
        End If
        If Trim("" & myQuotation_.isCase) <> "" Then
        Dim sqlStr As String
        Dim tmpRec As ADODB.Recordset
        Set tmpRec = New ADODB.Recordset
        sqlStr = "Select CaseName From tbWasteDataCase Where WasteCaseID = '" & myQuotation_.isCase & "' "
        Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
        cboQuotationType.Text = tmpRec(0)
        End If
        Dim sqlStrQ As String
        Dim tmpRecQ As ADODB.Recordset
        Set tmpRec = New ADODB.Recordset
        sqlStrQ = " select QuotationNo from tbQuotationHeader  Where QuotationID = '" & myQuotation_.RefQuotationID & "' "
        Set tmpRecQ = GetRS(sqlStrQ, adOpenForwardOnly, adLockReadOnly)
        If tmpRecQ.EOF = False Then
        txtRefQuotationID.Text = tmpRecQ(0)
        End If
    ValidateData
    ShowQuoDetails
    ShowQuoRemarks
    ShowSelTruck
    ShowCar
    myQuotation_.isChanged = False
    Set tmpCol = Nothing
    Set tmpS = Nothing
End Sub
Sub ProjectQuo()
    Dim ProID
    Dim Sql As String
    Dim rs As New ADODB.Recordset
    Dim i As Integer
    Dim ProjectList As Collection
    Sql = "SELECT * FROM vw_jobProject where CustomerID='" & curCustomer_.id & "' and ProjectStatus='W' " & vbCrLf
    Set rs = GetRS(Sql, adOpenForwardOnly, adLockOptimistic)
    cbProject.Clear
    If rs.RecordCount > 0 Then
    cbProject.AddItem "==เลือกโปรเจค=="
            For i = 1 To rs.RecordCount
                cbProject.AddItem rs!ProjectName & " # " & rs!ProjectID
                Set ProjectList = Nothing
            rs.MoveNext
            Next
            cbProject.ListIndex = 0
    End If
End Sub
Sub ShowSelTruck()
        Dim sqlSelTruck As String
        Dim tmpRecSelTruck As ADODB.Recordset
        Dim i As Integer
        Set tmpRecSelTruck = New ADODB.Recordset
        sqlSelTruck = " select * from vw_Seltruck  Where QuotationID = '" & myQuotation_.id & "' order by SelTruckID"
        Set tmpRecSelTruck = GetRS(sqlSelTruck, adOpenForwardOnly, adLockReadOnly)
        GrCar.getGridObj.Rows = 2
        If tmpRecSelTruck.EOF = False Then
            For i = 1 To tmpRecSelTruck.RecordCount
                With GrCar
                    .getGridObj.Rows = .getGridObj.Rows + 1
                    .getGridObj.TextMatrix(i, 0) = tmpRecSelTruck("TruckTypeID")
                    .getGridObj.TextMatrix(i, 1) = tmpRecSelTruck("TruckTypeDesc")
                    .getGridObj.TextMatrix(i, 2) = tmpRecSelTruck("TrailerType")
                    .getGridObj.TextMatrix(i, 3) = tmpRecSelTruck("TransportFee")
                    .getGridObj.TextMatrix(i, 4) = tmpRecSelTruck("MinWeight")
                End With
                tmpRecSelTruck.MoveNext
            Next
        End If

End Sub
Public Sub AddCartem(selItem As BWGTruckRate)
Dim tmpCol As Collection
Dim x%, isNew As Boolean
Dim tmpID As String
Dim tmpDetail As New BWGQuotationTruckRate
    Set tmpCol = myQuotation_.QuotationCar

    If tmpCol Is Nothing Then Set tmpCol = New Collection
    isNew = True
    tmpID = Trim(selItem.id)
    For x = 1 To tmpCol.Count
        'If tmpCol(x).selitem.ID = tmpID Then
        If tmpCol(x).TruckRateID = tmpID Then
            isNew = False
            Exit For
        End If
    Next
    
    If isNew Then
        If selItem.TumbolID = myQuotation_.TumbolID Then
            tmpDetail.id = GetGUID
            tmpDetail.QuotationID = myQuotation_.id
            tmpDetail.TypeDesc = selItem.TypeDesc
            tmpDetail.TruckRateID = selItem.id
            tmpDetail.Rate = selItem.Rate
            tmpDetail.TumbolName = selItem.TumbolName
            tmpDetail.AumphurName = selItem.AumphurName
            tmpDetail.ProvinceName = selItem.ProvinceName
            tmpCol.Add tmpDetail, "'" & tmpDetail.id & "'"
            isDetailChange_ = True
            myQuotation_.isChanged = True
            Set tmpDetail = Nothing
        End If
    End If
    If myQuotation_.QuotationCar Is Nothing Then
        Set myQuotation_.QuotationCar = tmpCol
        isDetailChange_ = True
        myQuotation_.isChanged = True
    End If
    Set tmpCol = Nothing
End Sub
Public Sub RemoveCar(selID As String)
Dim tmpCol As Collection
Dim x%
Dim tmpKey As String
    Set tmpCol = myQuotation_.QuotationCar
    If tmpCol Is Nothing Then Exit Sub
    For x = 1 To tmpCol.Count
        If tmpCol(x).TruckRateID = selID Then
            tmpKey = "'" & tmpCol(x).id & "'"
            Exit For
        End If
    Next
    If isExist(tmpCol, tmpKey) Then
        tmpCol.Remove tmpKey
        isDetailChange_ = True
        myQuotation_.isChanged = True
    End If
End Sub
Public Sub AddWasteDataItem(selItem As BWGWasteDataCR)
Dim tmpCol As Collection
Dim x%, isNew As Boolean
Dim tmpID As String
Dim tmpDetail As New BWGQuotationDetail
    Set tmpCol = myQuotation_.QuotationDetails

    If tmpCol Is Nothing Then Set tmpCol = New Collection
    isNew = True
    tmpID = Trim(selItem.id)
    For x = 1 To tmpCol.Count
        'If tmpCol(x).selitem.ID = tmpID Then
        If tmpCol(x).wastedataID = tmpID Then
            isNew = False
            Exit For
        End If
    Next
    
    myQuotation_.isQuotationReq = selItem.isQuotationReq
    
    If isNew Then
        tmpDetail.id = GetGUID
        tmpDetail.QuotationID = myQuotation_.id
        tmpDetail.wastedataID = selItem.id
        tmpDetail.WasteEUCode = selItem.EUCode1 & "-" & selItem.EUCode2 & "-" & selItem.EUCode3
        tmpDetail.WasteName = selItem.WasteName
        tmpDetail.WasteNo = selItem.WasteNo
        tmpDetail.BillDesc = selItem.WasteName
        tmpDetail.WasteType = selItem.WasteType
        tmpDetail.CompanyCode = selItem.CompanyCode
        tmpCol.Add tmpDetail, "'" & tmpDetail.id & "'"
        isDetailChange_ = True
        myQuotation_.isChanged = True
        Set tmpDetail = Nothing
    End If
    If myQuotation_.QuotationDetails Is Nothing Then
        Set myQuotation_.QuotationDetails = tmpCol
        isDetailChange_ = True
        myQuotation_.isChanged = True
    End If
    Set tmpCol = Nothing
End Sub

Public Sub SetDataCustomer(selCustomer As BWGCustomer, Optional isAssigned As Boolean = True)
On Error GoTo ErrD
Dim tmpRec As New ADODB.Recordset
Dim tmpCust As New ADODB.Recordset
Dim sqlStr As String
Dim Sql As String
    Set curCustomer_ = Nothing
    Set curCustomer_ = selCustomer
    If curCustomer_ Is Nothing Then Exit Sub
    cboCustStaff.Clear
    With curCustomer_
        sqlStr = "SELECT * FROM tbCustomerStaff WHERE CustomerID='" & .id & "'"
        Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
        If tmpRec Is Nothing Then Exit Sub
        If Trim(curCustomer_.ContactPerson) <> "" Then cboCustStaff.AddItem curCustomer_.ContactPerson
        Do Until tmpRec.EOF
            cboCustStaff.AddItem Trim("" & tmpRec("FName") & " " & tmpRec("LName"))
            tmpRec.MoveNext
        Loop
         Sql = "SELECT customercode FROM tbCustomer WHERE CustomerID='" & .id & "'"
        Set tmpCust = GetRS(Sql, adOpenForwardOnly, adLockReadOnly)
        If Left(tmpCust("customercode"), 1) = "M" Then
            cboCustStaff.Text = "คุณจารุวรรณ โพธิ์แจ้ง"
        End If
        
        If cboCustStaff.ListCount > 0 Then
            cboCustStaff.ListIndex = 0
            myQuotation_.AttentionTo = Trim(cboCustStaff.Text)
        End If
        tmpRec.Close
        Set tmpRec = Nothing
        txtCustomer.Text = .CustomerName
        If isAssigned Then
            myQuotation_.CustomerID = .id
            myQuotation_.CustomerName = .CustomerName
        End If
    End With
    Exit Sub
ErrD:
    GetMsg "", vbCritical
    Exit Sub
End Sub

Private Sub ShowQuoDetails()
Dim x%, y%
Dim tmpCol As Collection
Dim tmpDet As BWGQuotationDetail
Dim tmpKey As String, rowCount As Long
Dim forColor As String
Dim HeatValue As Double
Dim gr As ctlGrid
Dim i As Integer
    Set tmpCol = myQuotation_.QuotationDetails
    For i = 1 To tmpCol.Count
        If tmpCol(i).WasteStatus <> "A" Then
            myQuotation_.isWasteApproveAll = False
            Exit For
        End If
    Next
    If myQuotation_.QuotationType = 2 Then
        With grItemAP.getGridObj
            grItemAP.ClearAllData
            rowCount = 0
            forColor = vbBlack
            If .Rows - 1 <= tmpCol.Count Then .Rows = tmpCol.Count + 20
            For x = 1 To tmpCol.Count
                Set tmpDet = tmpCol(x)
                .TextMatrix(x + rowCount, 0) = tmpDet.id
                .row = x + rowCount
                .col = 1
                .CellForeColor = forColor
                .TextMatrix(x + rowCount, 1) = tmpDet.WasteNo
                .TextMatrix(x + rowCount, 2) = tmpDet.BillDesc
                HeatValue = tmpDet.HeatingValue
                .TextMatrix(x + rowCount, 3) = tmpDet.BurningRate
                .TextMatrix(x + rowCount, 4) = tmpDet.TreatmentUnitName
                .TextMatrix(x + rowCount, 5) = tmpDet.BlendingRate
                .TextMatrix(x + rowCount, 6) = tmpDet.ParameterRate
                .TextMatrix(x + rowCount, 7) = tmpDet.BurningRate + tmpDet.BlendingRate + tmpDet.ParameterRate
                tmpDet.RowPos = x
            Next
            .row = 1
            .col = 1
        End With
'    ElseIf myQuotation_.QuotationType = 1 Then
    Else
    With grItem.getGridObj
        grItem.ClearAllData
        rowCount = 0
        forColor = vbBlack
        If .Rows - 1 <= tmpCol.Count Then .Rows = tmpCol.Count + 20
        For x = 1 To tmpCol.Count
            Set tmpDet = tmpCol(x)
            .TextMatrix(x + rowCount, 0) = tmpDet.id
            .row = x + rowCount
            .col = 1
            .CellForeColor = forColor
            .TextMatrix(x + rowCount, 1) = tmpDet.WasteNo
            .TextMatrix(x + rowCount, 2) = Trim(tmpDet.BillDesc) & IIf(Trim(tmpDet.WasteType) <> "", "  [" & tmpDet.WasteType & "]", "")
            .TextMatrix(x + rowCount, 3) = tmpDet.YearQty
            .TextMatrix(x + rowCount, 4) = tmpDet.WasteQtyUnitName
            .TextMatrix(x + rowCount, 5) = tmpDet.TreatmentRate
            .TextMatrix(x + rowCount, 6) = tmpDet.TreatmentUnitName
            .TextMatrix(x + rowCount, 7) = tmpDet.Remdesc
            .TextMatrix(x + rowCount, 8) = tmpDet.Cost
            .TextMatrix(x + rowCount, 9) = tmpDet.CostUnit
            .TextMatrix(x + rowCount, 10) = tmpDet.CompanyCode
            tmpDet.RowPos = x
        Next
        .row = 1
        .col = 1
    End With
    End If
    Set tmpCol = Nothing
End Sub
Private Sub ShowCar()
Dim x%, y%
Dim tmpCol As Collection
Dim tmpDet As BWGQuotationTruckRate
Dim tmpKey As String, rowCount As Long
Dim forColor As String
Dim HeatValue As Double
Dim gr As ctlGrid
Dim i As Integer
    Set tmpCol = myQuotation_.QuotationCar
    With GrCar.getGridObj
        GrCar.ClearAllData
        rowCount = 0
        forColor = vbBlack
        If .Rows - 1 <= tmpCol.Count Then .Rows = tmpCol.Count + 20
        
        For x = 1 To tmpCol.Count
            Set tmpDet = tmpCol(x)
            .TextMatrix(x + rowCount, 0) = tmpDet.id
            .row = x + rowCount
            .col = 1
            .TextMatrix(x + rowCount, 1) = tmpDet.TypeDesc
            .TextMatrix(x + rowCount, 2) = Format(tmpDet.Rate, "#,###.00")
            .TextMatrix(x + rowCount, 3) = tmpDet.TumbolName & "," & tmpDet.AumphurName & "," & tmpDet.ProvinceName
        Next
        .row = 1
        .col = 1
    End With
    Set tmpCol = Nothing
End Sub
Public Function getWastDataItems() As Collection
On Error Resume Next
Dim x%
Dim tmpCol As New Collection
Dim tmpDet As BWGQuotationDetail
Dim tmpWaste As BWGWasteDataCR
Dim selCol As Collection, tmpStr As String, tmpArr
    Set selCol = myQuotation_.QuotationDetails
    With selCol
        For x = 1 To .Count
            Set tmpDet = .item(x)
            If Trim(tmpDet.wastedataID) <> "" Then
                Set tmpWaste = New BWGWasteDataCR
                tmpWaste.id = tmpDet.wastedataID
                tmpStr = tmpDet.WasteEUCode
                If InStr(1, tmpStr, "-") > 0 Then
                    tmpArr = Split(tmpStr, "-")
                    tmpWaste.EUCode1 = tmpArr(0)
                    tmpWaste.EUCode2 = tmpArr(1)
                    tmpWaste.EUCode3 = tmpArr(2)
                End If
                tmpWaste.WasteName = tmpDet.WasteName
                tmpWaste.WasteNo = tmpDet.WasteNo
                tmpCol.Add tmpWaste, "'" & tmpDet.wastedataID & "'"
                Set tmpWaste = Nothing
            End If
            Set tmpDet = Nothing
        Next
    End With

    Set getWastDataItems = tmpCol
    Set tmpCol = Nothing
End Function
Public Function getCartems() As Collection
On Error Resume Next
Dim x%
Dim tmpCol As New Collection
Dim tmpCarRate As BWGQuotationTruckRate
Dim tmpCar As BWGTruckRate
Dim searchM As New BWGSearchManager
Dim selCol As Collection, tmpStr As String, tmpArr
    Set selCol = myQuotation_.QuotationCar
    With selCol
        For x = 1 To .Count
            Set tmpCarRate = .item(x)
            If Trim(tmpCarRate.TruckRateID) <> "" Then
                Set tmpCar = searchM.TruckRateSearch("TruckRateID='" & tmpCarRate.TruckRateID & "'")
                tmpCol.Add tmpCar, "'" & tmpCarRate.TruckRateID & "'"
                Set tmpCar = Nothing
            End If
            Set tmpCarRate = Nothing
        Next
    End With

    Set getCartems = tmpCol
    Set tmpCol = Nothing
End Function

Private Sub SetEnabledScreen()
Dim eFlag As Boolean

Dim UserGrp As Collection
Dim Group As New AWSUserGroup
Dim StrGrp As String
Dim i As Integer
    eFlag = Not (myQuotation_.isCancel Or myQuotation_.isClosed Or myQuotation_.isApproved)
    
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
    
    If StrGrp = "Administrator" Then
        cmdApproved.Enabled = True
        cmdAddItem.Enabled = True
        cmdSave.Enabled = True
        Command1.Enabled = True
        cmdCancel.Enabled = True
        cmdCloseBill.Enabled = True
    Else
        cmdApproved.Enabled = eFlag
        cmdAddItem.Enabled = eFlag
        cmdSave.Enabled = eFlag
        Command1.Enabled = eFlag
        
        
        If myQuotation_.isCancel Or myQuotation_.isClosed Then
            cmdCancel.Enabled = False
            cmdCloseBill.Enabled = False
        Else
            cmdCancel.Enabled = True
            cmdCloseBill.Enabled = True
        End If
    End If
End Sub

Public Sub RemoveWasteData(selID As String)
Dim tmpCol As Collection
Dim x%
Dim tmpKey As String
    Set tmpCol = myQuotation_.QuotationDetails
    If tmpCol Is Nothing Then Exit Sub
    For x = 1 To tmpCol.Count
        If tmpCol(x).wastedataID = selID Then
            tmpKey = "'" & tmpCol(x).id & "'"
            Exit For
        End If
    Next
    If isExist(tmpCol, tmpKey) Then
        tmpCol.Remove tmpKey
        isDetailChange_ = True
        myQuotation_.isChanged = True
    End If
End Sub
'Private Sub ShowSelTruck()
'Dim x%, Y%
'Dim tmpCol As Collection
'Dim tmpDet As BWGTruckType
'Dim tmpKey As String, rowCount As Long
'Dim forColor As String
'    With GrCar.getGridObj
'        GrCar.ClearAllData
'        Set tmpCol = myQuotation_.selTruck(tmpDet.ID)
'        rowCount = 0
'        forColor = vbBlack
'        If .Rows - 1 <= tmpCol.Count Then .Rows = tmpCol.Count + 20
'        For x = 1 To tmpCol.Count
'            Set tmpDet = tmpCol(x)
'            .TextMatrix(x + rowCount, 0) = tmpDet.ID
'            .TextMatrix(x + rowCount, 1) = tmpDet.TypeDesc
'        Next
'        .row = 1
'        .col = 1
'        If .Rows < 3 Then .Rows = 20
'    End With
'    Set tmpCol = Nothing
'End Sub

Private Sub ShowQuoRemarks()
Dim x%, y%
Dim tmpCol As Collection
Dim tmpDet As BWGQuotationRem
Dim tmpKey As String, rowCount As Long
Dim forColor As String
    With grRemark.getGridObj
        grRemark.ClearAllData
                
        Set tmpCol = myQuotation_.RemarkDetails
        rowCount = 0
        forColor = vbBlack
        If .Rows - 1 <= tmpCol.Count Then .Rows = tmpCol.Count + 20
        For x = 1 To tmpCol.Count
            Set tmpDet = tmpCol(x)
            .TextMatrix(x + rowCount, 0) = tmpDet.id
            .TextMatrix(x + rowCount, 1) = tmpDet.Remdesc
            tmpDet.ItemNo = x
        Next
        .row = 1
        .col = 1
        If .Rows < 3 Then .Rows = 20
    End With
    Set tmpCol = Nothing
End Sub

Private Sub txtExtraCharge_Validate(Cancel As Boolean)
    If IsNumeric(txtExtraCharge.Text) Then
        myQuotation_.ExtraLaborFee = txtExtraCharge.Text
    Else
        myQuotation_.ExtraLaborFee = 0
        txtExtraCharge.Text = ""
        MsgBox "ข้อมูลไม่ถูกต้อง ต้องเป็นตัวเลขเท่านั้น", vbExclamation
    End If
End Sub

Private Sub txtHazRate_Validate(Cancel As Boolean)
    If IsNumeric(txtHazRate.Text) Then
        myQuotation_.HazMinRate = txtHazRate.Text
    Else
        myQuotation_.HazMinRate = 0
        txtHazRate.Text = ""
        MsgBox "ข้อมูลไม่ถูกต้อง ต้องเป็นตัวเลขเท่านั้น", vbExclamation
    End If
End Sub

Private Sub txtHBLRate_Validate(Cancel As Boolean)
    If IsNumeric(txtHBLRate.Text) Then
        myQuotation_.HBLMinRate = txtHBLRate.Text
    Else
        myQuotation_.HBLMinRate = 0
    End If
End Sub

Private Sub txtLaborCount_Validate(Cancel As Boolean)
'    If Trim(txtLaborCount.Text) <> "" Then
        If IsNumeric(txtLaborCount.Text) Then
'            If txtLaborCount.Text < 0 Then
'                MsgBox "ข้อมูลไม่ถูกต้อง ", vbExclamation
'                myQuotation_.StaffCountPerCar = 0
'                txtLaborCount.Text = ""
'                Exit Sub
                myQuotation_.StaffCountPerCar = txtLaborCount.Text
            Else
                myQuotation_.StaffCountPerCar = 0
'                txtLaborCount.Text = ""
'                MsgBox "ข้อมูลไม่ถูกต้อง ต้องเป็นตัวเลขเท่านั้น", vbExclamation
'            End If
        End If
'    End If
End Sub


Private Sub txtMinQty_Validate(Cancel As Boolean)
    If IsNumeric(txtMinQty.Text) Then
        myQuotation_.MinWeightPerCar = txtMinQty.Text
    Else
        myQuotation_.MinWeightPerCar = 0

    End If
End Sub

Public Sub SetDataRemark(selRem As BWGRemarkList)
Dim tmpDet As New BWGQuotationRem
    With tmpDet
        .id = GetGUID
        .QuotationID = myQuotation_.id
        .ItemNo = myQuotation_.RemarkDetails.Count + 1
        .Remdesc = selRem.Remdesc
        myQuotation_.RemarkDetails.Add tmpDet, "'" & .id & "'"
    End With
    Set tmpDet = Nothing
    ShowQuoRemarks
End Sub

Public Sub SetDataSelTruck(selTruck As BWGTruckType)
'Dim tmpDet As New BWGQuotationRem
'    With tmpDet
'        .ID = GetGUID
'        .QuotationID = myQuotation_.ID
'        .ItemNo = myQuotation_.RemarkDetails.Count + 1
'        .Remdesc = selRem.Remdesc
'        myQuotation_.RemarkDetails.Add tmpDet, "'" & .ID & "'"
'    End With
'    Set tmpDet = Nothing
    
End Sub


Public Sub SetDataTruckType(selTruck As BWGTruckType, Optional isAssigned As Boolean = True)
    Set curCarType_ = selTruck
        With curCarType_
            txtCarType.Text = .TypeDesc
            cboIsTrailer.Enabled = .CanBeTrailer
            If cboIsTrailer.Enabled = False Then cboIsTrailer.ListIndex = 0
            If isAssigned Then
                myQuotation_.TruckTypeID = .id
                myQuotation_.TruckTypeName = .TypeDesc
                myQuotation_.isTrailer = (cboIsTrailer.ListIndex = 1)
            End If
        End With
End Sub

Private Sub txtNBLRate_Validate(Cancel As Boolean)
'    If Trim(txtNBLRate.Text) <> "" Then
        If IsNumeric(txtNBLRate.Text) Then
'            If txtNBLRate.Text < 0 Then
'            MsgBox "ข้อมูลไม่ถูกต้อง ", vbExclamation
'            myQuotation_.NBLMinRate = 0
'            txtNBLRate.Text = ""
'            Exit Sub
'            End If
            myQuotation_.NBLMinRate = txtNBLRate.Text
        Else
            myQuotation_.NBLMinRate = 0
'            txtNBLRate.Text = ""
'            MsgBox "ข้อมูลไม่ถูกต้อง ต้องเป็นตัวเลขเท่านั้น", vbExclamation
        End If
'    End If
End Sub



Private Sub txtNonRate_Validate(Cancel As Boolean)
'    If Trim(txtNonRate.Text) <> "" Then
    If IsNumeric(txtNonRate.Text) And IsNumeric(txtNonRate.Text) < 0 Then
'        If txtNonRate.Text < 0 Then
'            MsgBox "ข้อมูลไม่ถูกต้อง ", vbExclamation
'            myQuotation_.NonMinRate = 0
'            txtNonRate.Text = ""
'            Exit Sub
'       End If
        myQuotation_.NonMinRate = txtNonRate.Text
    Else
        myQuotation_.NonMinRate = 0
'        txtNonRate.Text = ""
'        MsgBox "ข้อมูลไม่ถูกต้อง ต้องเป็นตัวเลขเท่านั้น", vbExclamation
    End If
'End If
End Sub

Private Sub TxtPosition_Validate(Cancel As Boolean)
    myQuotation_.SignPosition = txtPosition.Text
End Sub

Private Sub txtRefQuotationID_Validate(Cancel As Boolean)
 Dim SqlQuo As String
 Dim tmpRec As ADODB.Recordset
Set tmpRec = New ADODB.Recordset
 SqlQuo = "select QuotationID ,QuotationNo ,RefQuotationID from tbQuotationHeader  Where QuotationNo = '" & txtRefQuotationID.Text & "' And CustomerID = '" & myQuotation_.CustomerID & "' "
 Set tmpRec = GetRS(SqlQuo, adOpenForwardOnly, adLockReadOnly)
 If tmpRec.EOF = False Then
    myQuotation_.RefQuotationID = tmpRec(0)
Else
    MsgBox ("เลขที่ใบเสนอราคาไม่ถูกต้อง ตรวจสอบ")
 End If

End Sub

Private Sub txtRem_Validate(Cancel As Boolean)
    myQuotation_.Remdesc = txtRem.Text
End Sub

Private Sub txtTransRate_Validate(Cancel As Boolean)
    If IsNumeric(Trim(txtTransRate.Text)) Then
        myQuotation_.TransferRate = Trim(txtTransRate.Text)
        myQuotation_.Truck1Rate = Trim(txtTransRate.Text)
    Else
        myQuotation_.TransferRate = 0
        myQuotation_.Truck1Rate = 0
    End If
End Sub

Private Function CheckPendingJob() As Boolean
On Error GoTo ErrD
Dim tmpRec As New ADODB.Recordset
Dim tmpJobDet As BWGJobDataDetail
Dim sqlStr As String
    sqlStr = "SELECT * FROM vw_PendingJobOrder WHERE QuotationID='" & myQuotation_.id & "'"
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then Exit Function
    If tmpRec.EOF Then
        Set colPendingJob_ = Nothing
        CheckPendingJob = False
    Else
        CheckPendingJob = True
        Set colPendingJob_ = New Collection
        Do Until tmpRec.EOF
            Set tmpJobDet = New BWGJobDataDetail
            tmpJobDet.id = tmpRec("JobDetID")
            tmpJobDet.JobNo = Trim("" & tmpRec("JobNo"))
            tmpJobDet.JobID = Trim("" & tmpRec("JobID"))
            tmpJobDet.wastedataID = tmpRec("WasteDataID")
            tmpJobDet.WasteNo = tmpRec("WasteNo")
            tmpJobDet.WasteName = tmpRec("WasteName")
            tmpJobDet.QuoTruckTypeID = tmpRec("TruckTypeID")
            'tmpJobDet.QuoTruckTypeName = tmpRec("TruckTypeDesc")
            colPendingJob_.Add tmpJobDet, "'" & tmpJobDet.id & "'"
            Set tmpJobDet = Nothing
            tmpRec.MoveNext
        Loop
    End If
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function

Private Sub grItem_BeforeUpdateValue(NewValue As String, row As Integer, col As Integer, Cancel As Boolean)
    If col = 3 Or col = 5 Or col = 8 Then
        If IsNumeric(NewValue) = False Then
            MsgBox "ข้อมูลไม่ถูกต้อง ต้องเป็นตัวเลขเท่านั้น", vbExclamation
            Cancel = True
        End If
    End If
End Sub

Private Sub grItem_KeyDown(KeyCode As Integer, Shift As Integer)
Dim selResult As Integer
Dim x%, tmpKey$, itemCount%
Dim selCol As Collection
    With grItem.getGridObj
        If KeyCode = 46 And Trim(.TextMatrix(.row, 0)) <> "" Then
            selResult = MsgBox("ต้องการลบรายการกากของเสียนี้ออกจากใบเสนอราคาหรือไม่?", vbQuestion + vbYesNo)
            If selResult = vbYes Then
                tmpKey = "'" & .TextMatrix(.row, 0) & "'"
                .RemoveItem .row
                Set selCol = myQuotation_.QuotationDetails
                selCol.Remove tmpKey

                itemCount = myQuotation_.QuotationDetails.Count
                For x = 1 To .Rows - 1
                    tmpKey = "'" & .TextMatrix(x, 0) & "'"
                    If isExist(myQuotation_.QuotationDetails, tmpKey) Then
                        myQuotation_.QuotationDetails(tmpKey).RowPos = x
                    End If
                Next
                myQuotation_.isChanged = True
            End If
        End If
    End With
End Sub

Private Sub grItem_RolColChanged(row As Integer, col As Integer)
Dim x%
Dim tmpCbo As ComboBox
Dim tmpKey As String
    tmpKey = "'" & grItem.getGridObj.TextMatrix(row, 0) & "'"
    If Trim(tmpKey) = "''" Then
        grItem.ColEnabled(col) = False
    End If
    If isExist(myQuotation_.QuotationDetails, tmpKey) Then
        If col = 1 Or col = 2 Or col = 3 Or col = 4 Or col = 5 Or col = 6 Or col = 7 Or col = 8 Or col = 9 Then
            grItem.ColEnabled(col) = True
        Else
            grItem.ColEnabled(col) = False
        End If
    End If
    If (col = 4) And Not colUnit_ Is Nothing Then
        Set tmpCbo = grItem.getColObject(col)
        tmpCbo.Clear
        For x = 1 To colUnit_.Count
            tmpCbo.AddItem colUnit_(x).UnitName
        Next
        Set tmpCbo = Nothing
    ElseIf col = 6 Then
        Set tmpCbo = grItem.getColObject(col)
        tmpCbo.Clear
        With tmpCbo
            .AddItem "บาท/เที่ยว"
            .AddItem "บาท/ตัน"
            .AddItem "บาท/ถัง"
            .AddItem "บาท/ลบ.ม."
            .AddItem "บาท/กล่อง"
            .AddItem "บาท/งาน"
            .AddItem "บาท/Job"
            .AddItem "บาท/กิโลกรัม"
            .AddItem "บาท/ถุง"
            .AddItem "บาท/ขวด"
            .AddItem "บาท/แล็ก"
            .AddItem "บาท/กระป๋อง"
            .AddItem "บาท/แกลอน"
            .AddItem "บาท/วัน"
            .AddItem "บาท/หลอด"
            .AddItem "บาท/ลูก"
        End With
        ElseIf col = 9 Then
        Set tmpCbo = grItem.getColObject(col)
        tmpCbo.Clear
        With tmpCbo
            .AddItem "บาท/เที่ยว"
            .AddItem "บาท/ตัน"
            .AddItem "บาท/ถัง"
            .AddItem "บาท/ลบ.ม."
            .AddItem "บาท/กล่อง"
            .AddItem "บาท/งาน"
            .AddItem "บาท/Job"
            .AddItem "บาท/กิโลกรัม"
            .AddItem "บาท/ถุง"
            .AddItem "บาท/ขวด"
            .AddItem "บาท/แล็ก"
            .AddItem "บาท/กระป๋อง"
            .AddItem "บาท/แกลอน"
            .AddItem "บาท/วัน"
            .AddItem "บาท/หลอด"
            .AddItem "บาท/ลูก"
        End With

    End If
End Sub

Private Sub grItem_UpdateValue(row As Integer, col As Integer)
Dim tmpDet As BWGQuotationDetail
Dim tmpKey As String
Dim tmpVal As String
    tmpKey = Trim(grItem.getGridObj.TextMatrix(row, 0))
    Set tmpDet = myQuotation_.QuotationDetails("'" & tmpKey & "'")
    tmpVal = Trim(grItem.getGridObj.TextMatrix(row, col))
    With tmpDet
        Select Case col
            Case 2
                .BillDesc = Trim(tmpVal)
            Case 3
                If IsNumeric(Trim(tmpVal)) Then
                    .YearQty = CDbl(Trim(tmpVal))
                Else
                    .YearQty = 0
                End If
            Case 4
                If Trim(tmpVal) <> "" Then
                    .WasteQtyUnitID = colUnit_(grItem.getColObject(col).ListIndex + 1).id
                    .WasteQtyUnitName = Trim(tmpVal)
                Else
                    .WasteQtyUnitID = ""
                    .WasteQtyUnitName = ""
                End If
            Case 5
                If IsNumeric(Trim(tmpVal)) Then
                    .TreatmentRate = CDbl(Trim(tmpVal))
                Else
                    .TreatmentRate = 0
                End If
            Case 6
                If Trim(tmpVal) <> "" Then
                    .TreatmentUnitID = Trim(tmpVal)     'colUnit_(grItem.getColObject(7).ListIndex + 1).ID
                    .TreatmentUnitName = Trim(tmpVal)
                Else
                    .TreatmentUnitID = ""
                    .TreatmentUnitName = ""
                End If
            Case 7
                .Remdesc = Trim(tmpVal)
            Case 8
                If IsNumeric(Trim(tmpVal)) Then
                    .Cost = CDbl(Trim(tmpVal))
                Else
                    .Cost = 0
                End If
            Case 9
                If Trim(tmpVal) <> "" Then
                    .CostUnit = Trim(tmpVal)     'colUnit_(grItem.getColObject(7).ListIndex + 1).ID
                Else
                    .CostUnit = ""
                End If

        End Select
        myQuotation_.isChanged = True
    End With
    Set tmpDet = Nothing
End Sub
Private Sub txtTransRate2_Validate(Cancel As Boolean)
    If IsNumeric(Trim(txtTransRate2.Text)) Then
        myQuotation_.Truck2Rate = Trim(txtTransRate2.Text)
    Else
        myQuotation_.Truck2Rate = 0
    End If
End Sub
Private Sub txtTransRate3_Validate(Cancel As Boolean)
    If IsNumeric(Trim(txtTransRate3.Text)) Then
        myQuotation_.TRuck3Rate = Trim(txtTransRate3.Text)
    Else
        myQuotation_.TRuck3Rate = 0
    End If
End Sub
Private Sub txtTransRate4_Validate(Cancel As Boolean)
    If IsNumeric(Trim(txtTransRate4.Text)) Then
        myQuotation_.TRuck4Rate = Trim(txtTransRate4.Text)
    Else
        myQuotation_.TRuck4Rate = 0
    End If
End Sub
Private Sub txtTransRate5_Validate(Cancel As Boolean)
    If IsNumeric(Trim(txtTransRate5.Text)) Then
        myQuotation_.TRuck5Rate = Trim(txtTransRate5.Text)
    Else
        myQuotation_.TRuck5Rate = 0
    End If
End Sub
Private Sub txtTransRate6_Validate(Cancel As Boolean)
    If IsNumeric(Trim(txtTransRate6.Text)) Then
        myQuotation_.TRuck6Rate = Trim(txtTransRate6.Text)
    Else
        myQuotation_.TRuck6Rate = 0
    End If
End Sub
Private Sub txtTransRate7_Validate(Cancel As Boolean)
    If IsNumeric(Trim(txtTransRate7.Text)) Then
        myQuotation_.TRuck7Rate = Trim(txtTransRate7.Text)
    Else
        myQuotation_.TRuck7Rate = 0
    End If
End Sub
Private Sub txtTransRate8_Validate(Cancel As Boolean)
    If IsNumeric(Trim(txtTransRate8.Text)) Then
        myQuotation_.TRuck8Rate = Trim(txtTransRate8.Text)
    Else
        myQuotation_.TRuck8Rate = 0
    End If

End Sub
Private Sub txtTransRate9_Validate(Cancel As Boolean)
    If IsNumeric(Trim(txtTransRate9.Text)) Then
        myQuotation_.TRuck9Rate = Trim(txtTransRate9.Text)
    Else
        myQuotation_.TRuck9Rate = 0
    End If
End Sub
Private Sub txtTumbol_Validate(Cancel As Boolean)
Dim x%
Dim OldTumbolID As String
Dim tmpCol As Collection
Dim tmpDet As BWGQuotationTruckRate
 If Trim(txtTumbol.Text) = "" And Trim(curTumbol_.Name) <> Trim(txtTumbol.Text) Then
        Set curTumbol_ = Nothing
        Set curTumbol_ = New IMITumbol
        txtAumphur.Text = ""
        txtProvince.Text = ""
        With myQuotation_
            .TumbolID = ""
            .TumbolName = ""
            .AumphurID = ""
            .AumphurName = ""
            .ProvinceID = ""
            .ProvinceName = ""
        End With
    Else
        If Trim(curTumbol_.Name) <> Trim(txtTumbol.Text) Then
            Set curTumbol_ = Nothing
            Set curTumbol_ = New IMITumbol
            txtAumphur.Text = ""
            txtProvince.Text = ""
            With myQuotation_
            OldTumbolID = .TumbolID
                .TumbolID = ""
                .TumbolName = ""
                .AumphurID = ""
                .AumphurName = ""
                .ProvinceID = ""
                .ProvinceName = ""
            End With
            Load frmTumbolSearch
            With frmTumbolSearch
                .getCriteria (txtTumbol.Text)
                .SetCallerForm Me
                .Show 1
            End With
            If OldTumbolID <> myQuotation_.TumbolID Then
                Set tmpCol = myQuotation_.QuotationCar
             
                For x = 1 To tmpCol.Count
                    Set tmpDet = tmpCol(1)
                        RemoveCar tmpDet.TruckRateID
                Next x
                ShowCar
            End If
            If myQuotation_.TumbolID = "" Then txtTumbol.Text = ""
        End If
    End If
End Sub
Public Sub SetDataTumbol(selTumbol As IMITumbol, Optional isAssigned As Boolean = True)
    Set curTumbol_ = Nothing
    Set curTumbol_ = selTumbol
    If curTumbol_ Is Nothing Then Exit Sub
    With curTumbol_
        txtTumbol.Text = .Name
        txtAumphur.Text = .AumphurName
        txtProvince.Text = .ProvinceName
        If isAssigned Then
            myQuotation_.AumphurID = .AumphurID
            myQuotation_.AumphurName = .AumphurName
            myQuotation_.TumbolID = .id
            myQuotation_.TumbolName = .Name
            myQuotation_.ProvinceID = .ProvinceID
            myQuotation_.ProvinceName = .ProvinceName
        End If
    End With
End Sub

