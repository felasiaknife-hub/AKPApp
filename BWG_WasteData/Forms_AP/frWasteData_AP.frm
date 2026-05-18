VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmWasteData_AP 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   8280
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11280
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8280
   ScaleWidth      =   11280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAccepdate 
      Caption         =   "วันที่ตรวจ"
      Height          =   255
      Left            =   9780
      TabIndex        =   111
      Top             =   1560
      Width           =   975
   End
   Begin VB.ComboBox cboLABRes 
      Height          =   315
      ItemData        =   "frWasteData_AP.frx":0000
      Left            =   7080
      List            =   "frWasteData_AP.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   107
      Top             =   780
      Width           =   1275
   End
   Begin VB.CommandButton cmdApprove 
      Caption         =   "ตรวจรับโดย Envi"
      Height          =   795
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   102
      Top             =   7440
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtIndustryType 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1770
      TabIndex        =   15
      Top             =   1530
      Width           =   5235
   End
   Begin VB.CommandButton Command1 
      Caption         =   "..."
      Height          =   285
      Left            =   10590
      TabIndex        =   14
      Top             =   1170
      Width           =   345
   End
   Begin VB.TextBox txtExpDate 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   9030
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   1170
      Width           =   1545
   End
   Begin VB.TextBox txtWasteGenNo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5190
      TabIndex        =   12
      Top             =   1170
      Width           =   1815
   End
   Begin VB.TextBox txtFactoryNo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1770
      TabIndex        =   11
      Top             =   1170
      Width           =   1695
   End
   Begin VB.TextBox txtHAType 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   5640
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   450
      Width           =   435
   End
   Begin VB.TextBox txtContractNo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   9330
      TabIndex        =   10
      Top             =   810
      Width           =   1245
   End
   Begin VB.TextBox txtCustomer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000C0&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1380
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   90
      Width           =   9195
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "ปิด"
      Height          =   795
      Left            =   10260
      Style           =   1  'Graphical
      TabIndex        =   67
      Top             =   7440
      Width           =   945
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "เก็บข้อมูล"
      Height          =   795
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   63
      Top             =   7440
      Width           =   945
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "ค้นหา"
      Height          =   795
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   64
      Top             =   7470
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "ลบข้อมูล"
      Height          =   795
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   66
      Top             =   7440
      Width           =   945
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "สร้างใหม่"
      Height          =   795
      Left            =   8220
      Style           =   1  'Graphical
      TabIndex        =   65
      Top             =   7440
      Width           =   945
   End
   Begin VB.ComboBox cboSummary 
      Height          =   315
      ItemData        =   "frWasteData_AP.frx":0023
      Left            =   8790
      List            =   "frWasteData_AP.frx":0030
      Style           =   2  'Dropdown List
      TabIndex        =   62
      Top             =   6960
      Width           =   2385
   End
   Begin VB.TextBox txtEnviStaff 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   5370
      TabIndex        =   61
      Top             =   6990
      Width           =   2685
   End
   Begin VB.TextBox txtMktStaff 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1440
      TabIndex        =   60
      Top             =   6990
      Width           =   2685
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3495
      Left            =   30
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   3360
      Width           =   11205
      _ExtentX        =   19764
      _ExtentY        =   6165
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "คุณสมบัติของเสีย"
      TabPicture(0)   =   "frWasteData_AP.frx":0061
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "cboWasteChar"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cboWasteCharDesc"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "วิธีการจัดเก็บ"
      TabPicture(1)   =   "frWasteData_AP.frx":007D
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label2(17)"
      Tab(1).Control(1)=   "Label2(18)"
      Tab(1).Control(2)=   "Label2(19)"
      Tab(1).Control(3)=   "Label2(20)"
      Tab(1).Control(4)=   "Label2(31)"
      Tab(1).Control(5)=   "lvNewTreatment"
      Tab(1).Control(6)=   "lvOldTreatment"
      Tab(1).Control(7)=   "cboStoreType"
      Tab(1).Control(8)=   "cboIndoor"
      Tab(1).Control(9)=   "chkBag"
      Tab(1).Control(10)=   "txtBag"
      Tab(1).Control(11)=   "chkDrum"
      Tab(1).Control(12)=   "txtDrumDesc"
      Tab(1).Control(13)=   "chkOther2"
      Tab(1).Control(14)=   "txtOther2"
      Tab(1).Control(15)=   "chkTank"
      Tab(1).Control(16)=   "chkYard"
      Tab(1).Control(17)=   "txtOldStorePlace"
      Tab(1).Control(18)=   "cboTreatmentType"
      Tab(1).ControlCount=   19
      TabCaption(2)   =   "ค่าเกินเกณฑ์การออกแบบ"
      TabPicture(2)   =   "frWasteData_AP.frx":0099
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label2(32)"
      Tab(2).Control(1)=   "Label2(33)"
      Tab(2).Control(2)=   "Label2(34)"
      Tab(2).Control(3)=   "txtHeatValue"
      Tab(2).Control(4)=   "ctlDate1"
      Tab(2).Control(5)=   "grParameter"
      Tab(2).ControlCount=   6
      Begin WasteDataForm.ctlGrid grParameter 
         Height          =   2535
         Left            =   -74880
         TabIndex        =   110
         Top             =   780
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   4471
      End
      Begin WasteDataForm.ctlDate ctlDate1 
         Height          =   315
         Left            =   -73740
         TabIndex        =   109
         Top             =   420
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   556
         BackColor       =   -2147483633
      End
      Begin VB.TextBox txtHeatValue 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -70800
         TabIndex        =   59
         Top             =   420
         Width           =   1035
      End
      Begin VB.ComboBox cboTreatmentType 
         Height          =   315
         ItemData        =   "frWasteData_AP.frx":00B5
         Left            =   -66030
         List            =   "frWasteData_AP.frx":00BF
         Style           =   2  'Dropdown List
         TabIndex        =   47
         Top             =   480
         Width           =   1995
      End
      Begin VB.TextBox txtOldStorePlace 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -69780
         TabIndex        =   56
         Top             =   1590
         Width           =   3165
      End
      Begin VB.CheckBox chkYard 
         Caption         =   "ลานกอง"
         Height          =   255
         Left            =   -70440
         TabIndex        =   50
         Top             =   900
         Width           =   885
      End
      Begin VB.CheckBox chkTank 
         Caption         =   "บ่อรวบรวม"
         Height          =   255
         Left            =   -74280
         TabIndex        =   55
         Top             =   1620
         Width           =   1095
      End
      Begin VB.TextBox txtOther2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -69780
         Locked          =   -1  'True
         TabIndex        =   54
         Top             =   1230
         Width           =   5745
      End
      Begin VB.CheckBox chkOther2 
         Caption         =   "อื่นๆ"
         Height          =   255
         Left            =   -70440
         TabIndex        =   53
         Top             =   1260
         Width           =   675
      End
      Begin VB.TextBox txtDrumDesc 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -73170
         Locked          =   -1  'True
         TabIndex        =   52
         Top             =   1230
         Width           =   1875
      End
      Begin VB.CheckBox chkDrum 
         Caption         =   "ใส่ถังขนาด"
         Height          =   255
         Left            =   -74280
         TabIndex        =   51
         Top             =   1260
         Width           =   1125
      End
      Begin VB.TextBox txtBag 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -73170
         Locked          =   -1  'True
         TabIndex        =   49
         Top             =   870
         Width           =   1875
      End
      Begin VB.CheckBox chkBag 
         Caption         =   "ใส่ถุงขนาด"
         Height          =   255
         Left            =   -74280
         TabIndex        =   48
         Top             =   900
         Width           =   1065
      End
      Begin VB.ComboBox cboIndoor 
         Height          =   315
         ItemData        =   "frWasteData_AP.frx":00D5
         Left            =   -70440
         List            =   "frWasteData_AP.frx":00E2
         Style           =   2  'Dropdown List
         TabIndex        =   46
         Top             =   480
         Width           =   2985
      End
      Begin VB.ComboBox cboStoreType 
         Height          =   315
         ItemData        =   "frWasteData_AP.frx":00FA
         Left            =   -74280
         List            =   "frWasteData_AP.frx":0107
         Style           =   2  'Dropdown List
         TabIndex        =   45
         Top             =   480
         Width           =   2985
      End
      Begin VB.Frame Frame2 
         Caption         =   "คุณสมบัติของเสียอันตราย"
         Height          =   2505
         Left            =   0
         TabIndex        =   83
         Top             =   900
         Width           =   10815
         Begin WasteDataForm.ctlGrid grKorOrNo 
            Height          =   1755
            Left            =   6420
            TabIndex        =   108
            Top             =   660
            Width           =   4335
            _ExtentX        =   7646
            _ExtentY        =   3096
         End
         Begin VB.TextBox txtOther 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1290
            TabIndex        =   44
            Top             =   2040
            Width           =   5115
         End
         Begin VB.CheckBox chkOther 
            Caption         =   "อื่นๆ ระบุ"
            Height          =   255
            Left            =   300
            TabIndex        =   43
            Top             =   2070
            Width           =   1035
         End
         Begin VB.TextBox txtPesticide 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2730
            TabIndex        =   42
            Top             =   1710
            Width           =   3675
         End
         Begin VB.CheckBox chkPesticide 
            Caption         =   "สารกำจัดศัตรูพืช/วัชพืช   ระบุ"
            Height          =   255
            Left            =   300
            TabIndex        =   41
            Top             =   1740
            Width           =   2385
         End
         Begin VB.TextBox txtRadioActive 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2250
            TabIndex        =   40
            Top             =   1380
            Width           =   4155
         End
         Begin VB.CheckBox chkRadioActive 
            Caption         =   "กัมมันตภาพรังสี   ระบุ"
            Height          =   255
            Left            =   300
            TabIndex        =   39
            Top             =   1410
            Width           =   1815
         End
         Begin VB.ComboBox cboToxic 
            Height          =   315
            ItemData        =   "frWasteData_AP.frx":013A
            Left            =   2250
            List            =   "frWasteData_AP.frx":013C
            Style           =   2  'Dropdown List
            TabIndex        =   37
            Top             =   1020
            Width           =   1875
         End
         Begin VB.TextBox txtToxicDesc 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   4530
            TabIndex        =   38
            Top             =   1050
            Width           =   1875
         End
         Begin VB.CheckBox chkToxic 
            Caption         =   "มีสารเป็นพิษต่อร่างกาย"
            Height          =   255
            Left            =   300
            TabIndex        =   36
            Top             =   1050
            Width           =   1995
         End
         Begin VB.TextBox txtCancerDesc 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5220
            TabIndex        =   35
            Top             =   660
            Width           =   1155
         End
         Begin VB.CheckBox chkCancer 
            Caption         =   "สารก่อมะเร็ง   ระบุ"
            Height          =   255
            Left            =   3600
            TabIndex        =   34
            Top             =   660
            Width           =   1575
         End
         Begin VB.TextBox txtReactDesc 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1590
            TabIndex        =   33
            Top             =   660
            Width           =   1755
         End
         Begin VB.CheckBox chkReact 
            Caption         =   "ทำปฏิกริยากับ"
            Height          =   255
            Left            =   300
            TabIndex        =   32
            Top             =   690
            Width           =   1305
         End
         Begin VB.TextBox txtpH 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   9240
            TabIndex        =   31
            Top             =   240
            Width           =   975
         End
         Begin VB.CheckBox chkCorrosive 
            Caption         =   "กัดกร่อน    pH"
            Height          =   255
            Left            =   7740
            TabIndex        =   30
            Top             =   240
            Width           =   1305
         End
         Begin VB.TextBox txtFlashPoint 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3180
            TabIndex        =   29
            Top             =   300
            Width           =   435
         End
         Begin VB.CheckBox chkFlame 
            Caption         =   "ติดไฟ  จุดวาบไฟ"
            Height          =   255
            Left            =   1620
            TabIndex        =   28
            Top             =   300
            Width           =   1575
         End
         Begin VB.CheckBox chkExplode 
            Caption         =   "ระเบิดได้"
            Height          =   255
            Left            =   300
            TabIndex        =   27
            Top             =   330
            Width           =   1095
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ระบุ"
            Height          =   195
            Index           =   16
            Left            =   4200
            TabIndex        =   86
            Top             =   1080
            Width           =   270
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "C"
            Height          =   195
            Index           =   15
            Left            =   3780
            TabIndex        =   85
            Top             =   360
            Width           =   105
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
            Left            =   3720
            TabIndex        =   84
            Top             =   240
            Width           =   75
         End
      End
      Begin VB.ComboBox cboWasteCharDesc 
         Height          =   315
         ItemData        =   "frWasteData_AP.frx":013E
         Left            =   2310
         List            =   "frWasteData_AP.frx":014B
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   450
         Width           =   1695
      End
      Begin VB.ComboBox cboWasteChar 
         Height          =   315
         ItemData        =   "frWasteData_AP.frx":0164
         Left            =   330
         List            =   "frWasteData_AP.frx":0174
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   450
         Width           =   1905
      End
      Begin MSComctlLib.ListView lvOldTreatment 
         Height          =   1335
         Left            =   -74280
         TabIndex        =   57
         Top             =   1950
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
      Begin MSComctlLib.ListView lvNewTreatment 
         Height          =   1335
         Left            =   -69180
         TabIndex        =   58
         Top             =   1950
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
            Text            =   "วิธีการกำจัดที่ควรใช้"
            Object.Width           =   7320
         EndProperty
      End
      Begin VB.Label Label2 
         Caption         =   "วันที่วิเคราะห์"
         Height          =   195
         Index           =   34
         Left            =   -74820
         TabIndex        =   106
         Top             =   480
         Width           =   1035
      End
      Begin VB.Label Label2 
         Caption         =   "(k-cal/kg)"
         Height          =   195
         Index           =   33
         Left            =   -69660
         TabIndex        =   105
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ค่าความร้อน"
         Height          =   195
         Index           =   32
         Left            =   -71700
         TabIndex        =   104
         Top             =   450
         Width           =   855
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "รูปแบบการบำบัด"
         Height          =   195
         Index           =   31
         Left            =   -67290
         TabIndex        =   103
         Top             =   540
         Width           =   1185
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "สถานที่กำจัดเดิม"
         Height          =   195
         Index           =   20
         Left            =   -71100
         TabIndex        =   90
         Top             =   1620
         Width           =   1155
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "10.3)"
         Height          =   195
         Index           =   19
         Left            =   -74760
         TabIndex        =   89
         Top             =   930
         Width           =   345
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "10.2)"
         Height          =   195
         Index           =   18
         Left            =   -70920
         TabIndex        =   88
         Top             =   540
         Width           =   345
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "10.1)"
         Height          =   195
         Index           =   17
         Left            =   -74760
         TabIndex        =   87
         Top             =   540
         Width           =   330
      End
   End
   Begin VB.TextBox txtGenerateDesc 
      Appearance      =   0  'Flat
      Height          =   615
      Left            =   1380
      MultiLine       =   -1  'True
      TabIndex        =   23
      Top             =   2670
      Width           =   9855
   End
   Begin VB.Frame Frame1 
      Caption         =   "ปริมาณของเสีย"
      Height          =   705
      Left            =   30
      TabIndex        =   76
      Top             =   1890
      Width           =   11205
      Begin VB.ComboBox cboTimeRate 
         Height          =   315
         ItemData        =   "frWasteData_AP.frx":01A0
         Left            =   10290
         List            =   "frWasteData_AP.frx":01AA
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   300
         Width           =   795
      End
      Begin VB.ComboBox cboUnitRate 
         Height          =   315
         ItemData        =   "frWasteData_AP.frx":01B9
         Left            =   9360
         List            =   "frWasteData_AP.frx":01C3
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   300
         Width           =   735
      End
      Begin VB.TextBox txtRate 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   8310
         TabIndex        =   20
         Top             =   300
         Width           =   1005
      End
      Begin VB.ComboBox cboMaxUnit 
         Height          =   315
         ItemData        =   "frWasteData_AP.frx":01D2
         Left            =   5940
         List            =   "frWasteData_AP.frx":01DC
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   300
         Width           =   735
      End
      Begin VB.TextBox txtMaxQty 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4680
         TabIndex        =   18
         Top             =   300
         Width           =   1215
      End
      Begin VB.ComboBox cboQtyUnit 
         Height          =   315
         ItemData        =   "frWasteData_AP.frx":01EB
         Left            =   2700
         List            =   "frWasteData_AP.frx":01F5
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   300
         Width           =   735
      End
      Begin VB.TextBox txtQty 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         TabIndex        =   16
         Top             =   300
         Width           =   1215
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "/"
         Height          =   195
         Index           =   11
         Left            =   10140
         TabIndex        =   80
         Top             =   330
         Width           =   75
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "อัตราการเกิดของเสีย"
         Height          =   195
         Index           =   10
         Left            =   6810
         TabIndex        =   79
         Top             =   330
         Width           =   1440
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "เก็บได้มากที่สุด"
         Height          =   195
         Index           =   9
         Left            =   3570
         TabIndex        =   78
         Top             =   330
         Width           =   1065
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "จัดเก็บ ณ ปัจจุบัน"
         Height          =   195
         Index           =   8
         Left            =   150
         TabIndex        =   77
         Top             =   330
         Width           =   1215
      End
   End
   Begin VB.TextBox txtWasteName 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1380
      TabIndex        =   9
      Top             =   810
      Width           =   4725
   End
   Begin VB.ComboBox cboWasteType 
      Height          =   315
      Left            =   7320
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   420
      Width           =   1035
   End
   Begin VB.TextBox txtKor 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   10140
      TabIndex        =   8
      Top             =   450
      Width           =   435
   End
   Begin VB.TextBox txtMoud 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   9330
      TabIndex        =   7
      Top             =   450
      Width           =   435
   End
   Begin VB.TextBox txtEUCode3 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   5100
      TabIndex        =   4
      Top             =   450
      Width           =   435
   End
   Begin VB.TextBox txtEUCode2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   4470
      TabIndex        =   3
      Top             =   450
      Width           =   435
   End
   Begin VB.TextBox txtEUCode1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   3840
      TabIndex        =   2
      Top             =   450
      Width           =   435
   End
   Begin VB.TextBox txtWasteNo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000C0&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1380
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Text            =   "== AUTO =="
      Top             =   450
      Width           =   1455
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
      Left            =   7080
      TabIndex        =   101
      Top             =   1530
      Visible         =   0   'False
      Width           =   2565
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ประเภทอุตสาหกรรม"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   30
      Left            =   300
      TabIndex        =   100
      Top             =   1560
      Width           =   1410
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ผล LAB"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   29
      Left            =   6420
      TabIndex        =   99
      Top             =   840
      Width           =   540
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "วันที่ใบอนุญาตหมดอายุ"
      Height          =   195
      Index           =   28
      Left            =   7350
      TabIndex        =   98
      Top             =   1200
      Width           =   1620
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "เลขที่ผู้กำเนิดของเสีย"
      Height          =   195
      Index           =   27
      Left            =   3660
      TabIndex        =   97
      Top             =   1200
      Width           =   1470
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "เลขทะเบียนโรงงาน"
      Height          =   195
      Index           =   26
      Left            =   330
      TabIndex        =   96
      Top             =   1200
      Width           =   1350
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "เลขที่สัญญา"
      Height          =   195
      Index           =   25
      Left            =   8460
      TabIndex        =   95
      Top             =   870
      Width           =   795
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ชื่อลูกค้า"
      Height          =   195
      Index           =   24
      Left            =   720
      TabIndex        =   94
      Top             =   120
      Width           =   570
   End
   Begin VB.Line Line1 
      X1              =   30
      X2              =   11250
      Y1              =   7350
      Y2              =   7350
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ผลสรุป"
      Height          =   195
      Index           =   23
      Left            =   8220
      TabIndex        =   93
      Top             =   7020
      Width           =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "เจ้าหน้าที่ Envi"
      Height          =   195
      Index           =   22
      Left            =   4290
      TabIndex        =   92
      Top             =   7020
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
      TabIndex        =   91
      Top             =   7020
      Width           =   1290
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ทำให้เกิดของเสีย"
      Height          =   195
      Index           =   13
      Left            =   120
      TabIndex        =   82
      Top             =   2940
      Width           =   1185
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "กระบวนการที่"
      Height          =   195
      Index           =   12
      Left            =   360
      TabIndex        =   81
      Top             =   2670
      Width           =   960
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ชื่อกากของเสีย"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   7
      Left            =   330
      TabIndex        =   75
      Top             =   840
      Width           =   1005
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ประเภทของเสีย"
      Height          =   195
      Index           =   6
      Left            =   6180
      TabIndex        =   74
      Top             =   480
      Width           =   1080
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ข้อ"
      Height          =   195
      Index           =   5
      Left            =   9900
      TabIndex        =   73
      Top             =   480
      Width           =   180
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "หมวด"
      Height          =   195
      Index           =   4
      Left            =   8850
      TabIndex        =   72
      Top             =   480
      Width           =   405
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      Height          =   195
      Index           =   3
      Left            =   4980
      TabIndex        =   71
      Top             =   480
      Width           =   45
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      Height          =   195
      Index           =   2
      Left            =   4350
      TabIndex        =   70
      Top             =   480
      Width           =   45
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "รหัสของเสีย"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   0
      Left            =   2970
      TabIndex        =   69
      Top             =   480
      Width           =   810
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "เลขที่ของเสีย"
      Height          =   195
      Index           =   1
      Left            =   450
      TabIndex        =   68
      Top             =   480
      Width           =   885
   End
End
Attribute VB_Name = "frmWasteData_AP"
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
Dim myWasteParameter_ As BWGWasteParameter
Dim colTreatment_ As Collection
Dim callStaffType_ As String
Dim curIndustryType_ As BWGIndustryType

Private Sub LoadParameter(Optional ByVal wastedataID As String = "")
Dim rs As ADODB.Recordset
Dim tSql As String
    tSql = "SELECT * FROM vw_WasteDataParameter_Search" & vbCrLf
    If wastedataID <> "" Then
        tSql = tSql & "WHERE WasteDataID = '" & wastedataID & "'" & vbCrLf
    Else
        tSql = tSql & "WHERE WasteDataID IS NULL" & vbCrLf
    End If
    tSql = tSql & "ORDER BY Parameter_DefaultRate DESC, ParameterName"
    Set rs = New ADODB.Recordset
    Set rs = GetRS(tSql, adOpenForwardOnly, adLockReadOnly)
'    Set RS = GetRS(tSql, adOpenForwardOnly, adLockReadOnly)
    If Not rs.EOF Then
        grParameter.getGridObj.Rows = rs.RecordCount + 1
        Do While Not rs.EOF
            With grParameter.getGridObj
                .TextMatrix(rs.AbsolutePosition, 0) = rs!ParameterID
                .TextMatrix(rs.AbsolutePosition, 1) = rs!ParameterName
                .TextMatrix(rs.AbsolutePosition, 2) = rs!ParameterIndex
                .TextMatrix(rs.AbsolutePosition, 3) = rs!Parameter_DefaultRate
                .TextMatrix(rs.AbsolutePosition, 4) = rs!Parameter_ChargeMoreDefault
            End With
            rs.MoveNext
        Loop
    Else
        tSql = "SELECT     ParameterID, ParameterName, 0 AS ParameterIndex, Parameter_DefaultRate, Parameter_ChargeMoreDefault" & vbCrLf
        tSql = tSql & "From tbWasteParameter" & vbCrLf
        tSql = tSql & "ORDER BY Parameter_DefaultRate DESC, ParameterName"
        Set rs = New ADODB.Recordset
        Set rs = GetRS(tSql, adOpenForwardOnly, adLockReadOnly)
'        Set RS = GetRS(tSql, adOpenForwardOnly, adLockReadOnly)
        grParameter.getGridObj.Rows = rs.RecordCount + 1
        Do While Not rs.EOF
            With grParameter.getGridObj
                .TextMatrix(rs.AbsolutePosition, 0) = rs!ParameterID
                .TextMatrix(rs.AbsolutePosition, 1) = rs!ParameterName
                .TextMatrix(rs.AbsolutePosition, 2) = rs!ParameterIndex
                .TextMatrix(rs.AbsolutePosition, 3) = rs!Parameter_DefaultRate
                .TextMatrix(rs.AbsolutePosition, 4) = rs!Parameter_ChargeMoreDefault
            End With
            rs.MoveNext
        Loop
        Set rs = Nothing
    End If
    Set rs = Nothing
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

Private Sub cboQtyUnit_Validate(Cancel As Boolean)
    Select Case cboQtyUnit.ListIndex
        Case 0
            myWasteDataCR_.NormalUnit = "T" 'ตัน
        Case 1
            myWasteDataCR_.NormalUnit = "L" 'ลิตร
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

Private Sub cboTimeRate_Validate(Cancel As Boolean)
    Select Case cboTimeRate.ListIndex
        Case 0
            myWasteDataCR_.WasteRateTime = "M" 'ตัน
        Case 1
            myWasteDataCR_.WasteRateTime = "Y" 'ลิตร
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

Private Sub cboUnitRate_Validate(Cancel As Boolean)
    Select Case cboUnitRate.ListIndex
        Case 0
            myWasteDataCR_.WasteRateUnit = "T" 'ตัน
        Case 1
            myWasteDataCR_.WasteRateUnit = "L" 'ลิตร
    End Select
End Sub

Private Sub cboWasteChar_Click()
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
        Case 4
            cboWasteCharDesc.Clear
            cboWasteCharDesc.Enabled = False
    End Select
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
End Sub

Private Sub cboWasteType_Validate(Cancel As Boolean)
'    Select Case cboWasteType.ListIndex
'        Case 0
'            myWasteDataCR_.WasteType = "N/A"
'        Case 1
'            myWasteDataCR_.WasteType = "Non"
'        Case 2
'            myWasteDataCR_.WasteType = "Haz"
'        Case 3
'            myWasteDataCR_.WasteType = "NBL"
'        Case 4
'            myWasteDataCR_.WasteType = "HBL"
'    End Select
    If cboWasteType.ListIndex = 0 Then
        myWasteDataCR_.WasteType = "N/A"
    Else
        myWasteDataCR_.WasteType = cboWasteType.Text
    End If
End Sub

Private Sub chkBag_Click()
    txtBag.Locked = chkBag.Value = 0
    If chkBag.Value = 1 Then
        If Me.Visible = True Then
            If Not Me.ActiveControl Is Nothing Then
               If Me.ActiveControl.Name = "chkBag" Then chkBag_Validate False
            End If
            txtBag.SetFocus
        End If
        chkYard.Value = 0
        chkDrum.Value = 0
        chkOther2.Value = 0
        chkTank.Value = 0
    End If
End Sub

Private Sub chkBag_Validate(Cancel As Boolean)
    If chkBag.Value = 1 Then
        myWasteDataLab_.StorageType = "BA" 'Bag
        myWasteDataLab_.StorageDesc = Trim(txtBag.Text)
    End If
End Sub

Private Sub chkCancer_Validate(Cancel As Boolean)
    myWasteDataLab_.isCarcinoGenic = chkCancer.Value = 1
End Sub

Private Sub chkCorrosive_Validate(Cancel As Boolean)
    myWasteDataLab_.isCorrosive = chkCorrosive.Value = 1
End Sub

Private Sub chkDrum_Click()
    txtDrumDesc.Locked = chkDrum.Value = 0
    If chkDrum.Value = 1 Then
        If Me.Visible = True Then
            If Not Me.ActiveControl Is Nothing Then
               If Me.ActiveControl.Name = "chkDrum" Then chkDrum_Validate False
            End If
            txtDrumDesc.SetFocus
        End If
        chkBag.Value = 0
        chkYard.Value = 0
        chkOther2.Value = 0
        chkTank.Value = 0
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

Private Sub chkOther2_Click()
    txtOther2.Locked = chkOther2.Value = 0
    If chkOther2.Value = 1 Then
        If Me.Visible = True Then
            If Not Me.ActiveControl Is Nothing Then
               If Me.ActiveControl.Name = "chkOther2" Then chkOther2_Validate False
            End If
            txtOther2.SetFocus
        End If
        chkYard.Value = 0
        chkDrum.Value = 0
        chkBag.Value = 0
        chkTank.Value = 0
    End If
End Sub

Private Sub chkOther2_Validate(Cancel As Boolean)
    If chkOther2.Value = 1 Then
        myWasteDataLab_.StorageType = "OT" 'Other
        myWasteDataLab_.StorageDesc = Trim(txtOther2.Text)
    End If
End Sub

Private Sub chkPesticide_Validate(Cancel As Boolean)
    myWasteDataLab_.isPesticide = chkPesticide.Value = 1
End Sub

Private Sub chkRadioActive_Validate(Cancel As Boolean)
    myWasteDataLab_.isRadioActive = chkRadioActive.Value = 1
End Sub

Private Sub chkReact_Validate(Cancel As Boolean)
    myWasteDataLab_.isReact = chkReact.Value = 1
End Sub

Private Sub chkTank_Click()
    If chkTank.Value = 1 Then
        chkYard.Value = 0
        chkDrum.Value = 0
        chkOther2.Value = 0
        chkBag.Value = 0
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
    If chkYard.Value = 1 Then
        chkBag.Value = 0
        chkDrum.Value = 0
        chkOther2.Value = 0
        chkTank.Value = 0
    End If
End Sub

Private Sub chkYard_Validate(Cancel As Boolean)
    If chkYard.Value = 1 Then
        myWasteDataLab_.StorageType = "ST" 'Stock yard
        myWasteDataLab_.StorageDesc = ""
    End If
End Sub

Private Sub cmdAccepdate_Click()
frmWastedataAccepdate.Show
End Sub

Private Sub cmdApprove_Click()
    If MsgBox("ท่านต้องการตรวจรับของเสียรายการนี้หรือไม่?", vbQuestion + vbYesNo) = vbYes Then
        myWasteDataCR_.EnviAccepted
        Label1.Caption = "Envi ตรวจรับแล้วเมื่อ " & FormatYMD_to_DMY(myWasteDataCR_.acceptedDate, "/", "/")
        Label1.Visible = True
        cmdApprove.Visible = False
    End If
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
Dim i As Integer
    If Trim(txtCustomer.Text) = "" Then
        MsgBox "กรุณาระบุลูกค้าก่อนทำการจัดเก็บข้อมูล", vbExclamation
        txtCustomer.SetFocus
        Exit Sub
    End If
    If Trim(txtWasteName.Text) = "" Then
        MsgBox "กรุณาระบุชื่อของการของเสีย", vbExclamation
        txtWasteName.SetFocus
        Exit Sub
    End If
    If Trim(txtMktStaff.Text) = "" Then
        MsgBox "กรุณาระบุชื่อเจ้าหน้าที่การตลาด", vbExclamation
        txtMktStaff.SetFocus
        Exit Sub
    End If
    If cboLABRes.ListIndex = -1 Then
        MsgBox "กรุณาระบุผล LAB", vbExclamation
        cboLABRes.SetFocus
        Exit Sub
    End If
    If myWasteDataCR_.isChanged Then
        myWasteDataCR_.Save
        txtWasteNo.Text = myWasteDataCR_.WasteNo
    End If
    If myWasteDataLab_.isChanged Then
        myWasteDataLab_.Save
        cmdApprove.Visible = True
    End If
    If myWasteParameter_.isChanged = True Then
        For i = 1 To grParameter.getGridObj.Rows - 1
            myWasteParameter_.ID = myWasteDataCR_.ID
            myWasteParameter_.ParameterID = grParameter.getGridObj.TextMatrix(i, 0)
            myWasteParameter_.ParameterIndex = grParameter.getGridObj.TextMatrix(i, 2)
            myWasteParameter_.Save
        Next
    End If
    If myWasteDataCR_.isChanged = False And myWasteDataLab_.isChanged = False Then
        MsgBox "จัดเก็บข้อมูลเรียบร้อยแล้ว", vbExclamation
    End If
End Sub

Private Sub Command1_Click()
    Load frmDateAssign
    With frmDateAssign
        .SetCallerForm Me
        .Show 1
    End With
End Sub
Private Sub ctlDate1_GotFocus()
    SSTab1.Tab = 2
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If (KeyCode = 13 And Me.ActiveControl.Name <> "grKorOrNo") And (KeyCode = 13 And Me.ActiveControl.Name <> "grParameter") Then
        SendKeys "{Tab}"
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
Dim x%
    With cboWasteType
        .AddItem "ยังไม่ระบุ"
        For x = 1 To WasteTypes.Count
            .AddItem WasteTypes(x).WasteType
        Next
    End With
    With frmPicture
        cmdSave.Picture = .picSave.Picture
        cmdClear.Picture = .picNew.Picture
        cmdSearch.Picture = .picSearch.Picture
        cmdClose.Picture = .picExit.Picture
        cmdDel.Picture = .picDelete.Picture
    End With
    With grKorOrNo.getGridObj
        .Cols = 3
        .Rows = 20
        .ColWidth(0) = 0
        .ColWidth(1) = 2235
        .ColWidth(2) = 2230
        .TextMatrix(0, 1) = "เลขที่ใบ ก.อ."
        .TextMatrix(0, 2) = "วันที่"
        .FixedAlignment(1) = 1
        .ColAlignment(1) = 1
    End With
    With grParameter.getGridObj
        .Cols = 5
        .Rows = 20
        .ColWidth(0) = 0
        .ColWidth(1) = 4680
        .ColWidth(2) = 2000
        .ColWidth(3) = 2000
        .ColWidth(4) = 2000
        .TextMatrix(0, 1) = "พารามิเตอร์"
        .TextMatrix(0, 2) = "ค่าดัชนี"
        .TextMatrix(0, 3) = "ราคาเริ่มต้น"
        .TextMatrix(0, 4) = "อัตราส่วนเพิ่ม"
        .ColAlignment(1) = 3
        .BackColorFixed = &H8000000F
        .ForeColorFixed = &HFF0000
        For x = 1 To .Rows - 1
        .col = 1
        .row = x
        .CellBackColor = &HC0E0FF
        Next
    End With
    
    cboWasteType.ListIndex = 0
'    cboWasteChar.ListIndex = 0
    'cboWasteCharDesc.ListIndex = 0
'    cboSummary.ListIndex = 0
'    cboStoreType.ListIndex = 0
'    cboIndoor.ListIndex = 0
    Set curEUCode1_ = New BWGEUCode1
    Set curEUCode2_ = New BWGEUCode2
    Set curEUCode3_ = New BWGEUCode3
    Set curMktStaff_ = New BWGCompanyStaff
    Set curEnviStaff_ = New BWGCompanyStaff
    PopulateTreatment
    CreateNewWasteData
    LoadParameter
    SSTab1.Tab = 0
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
        If col = 1 Then
            tmpDet.KorOrNumber = Trim(.TextMatrix(row, 1))
            myWasteDataCR_.isChanged = True
        End If
        Set tmpDet = Nothing
    End With
End Sub

Private Sub grParameter_BeforeUpdateValue(newValue As String, row As Integer, col As Integer, Cancel As Boolean)
    If col = 1 Or col = 3 Or col = 4 Then Cancel = True
    If col = 1 Then
        If IsNumeric(newValue) = False Then
            MsgBox "ค่าดัชนี ใส่ได้เฉพาะตัวเลขเท่านั้น", vbExclamation, "Warning"
            Cancel = True
            newValue = 0
        End If
    End If
End Sub

Private Sub grParameter_RolColChanged(row As Integer, col As Integer)
    If col = 1 Or col = 3 Or col = 4 Then
        grParameter.ColEnabled(col) = False
    End If
End Sub

Private Sub grParameter_Validate(Cancel As Boolean)
    myWasteParameter_.isChanged = True
End Sub

Private Sub lvNewTreatment_ItemCheck(ByVal item As MSComctlLib.ListItem)
    If item.Checked Then
        If isExist(myWasteDataLab_.NewTreatments, item.Key) = False Then
            myWasteDataLab_.NewTreatments.Add colTreatment_(item.Key), item.Key
            myWasteDataLab_.isChanged = True
        End If
    Else
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

Private Sub Text1_Change()

End Sub

Private Sub Text1_Validate(Cancel As Boolean)

End Sub


Private Sub txtBag_GotFocus()
    txtBag.Locked = chkBag.Value = 0
    SSTab1.Tab = 1
End Sub

Private Sub txtBag_Validate(Cancel As Boolean)
    If chkBag.Value = 1 Then myWasteDataLab_.StorageDesc = Trim(txtBag.Text)
End Sub

Private Sub txtCancerDesc_Validate(Cancel As Boolean)
    myWasteDataLab_.CarcinoGenicDesc = Trim(txtCancerDesc.Text)
End Sub

Private Sub txtContractNo_Validate(Cancel As Boolean)
    If Trim(txtContractNo.Text) = "" And Trim(myWasteDataCR_.Contract.docNo) <> Trim(txtContractNo.Text) Then
        Set myWasteDataCR_.Contract = New BWGContract
    Else
        If Trim(myWasteDataCR_.Contract.docNo) <> Trim(txtContractNo.Text) Then
            Set myWasteDataCR_.Contract = New BWGContract
            Load frmContractSearch
            With frmContractSearch
                .getCriteria txtContractNo.Text
                .SetCallerForm Me
                .Show 1
            End With
            If myWasteDataCR_.Contract.docNo = "" Then txtContractNo.Text = ""
        End If
    End If
End Sub

Private Sub txtDrumDesc_GotFocus()
    txtDrumDesc.Locked = chkDrum.Value = 0
End Sub

Private Sub txtDrumDesc_Validate(Cancel As Boolean)
    If chkDrum.Value = 1 Then myWasteDataLab_.StorageDesc = Trim(txtDrumDesc.Text)
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
                .getCriteria (txtEnviStaff.Text)
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
    If IsNumeric(Trim(txtFlashPoint.Text)) Then
        myWasteDataLab_.FlashPoint = CDbl(Trim(txtFlashPoint.Text))
    Else
        myWasteDataLab_.FlashPoint = 0
    End If
End Sub

Private Sub txtGenerateDesc_Validate(Cancel As Boolean)
    myWasteDataLab_.GenerateDetail = Trim(txtGenerateDesc.Text)
End Sub

Private Sub txtHeatValue_Validate(Cancel As Boolean)
    If IsNumeric(txtHeatValue) = False Then
        MsgBox "ค่าความร้อนใส่ได้เฉพาะตัวเลขเท่านั้น", vbCritical, "Warning"
        txtHeatValue.Text = "0"
    End If
    myWasteDataCR_.HeatValue = CDbl(txtHeatValue)
    myWasteDataCR_.isChanged = True
End Sub


Private Sub txtIndustryType_Validate(Cancel As Boolean)
    If curIndustryType_ Is Nothing Then Set curIndustryType_ = New BWGIndustryType
    If Trim(txtIndustryType.Text) = "" And Trim(curIndustryType_.TypeName) <> Trim(txtIndustryType.Text) Then
        Set curIndustryType_ = Nothing
        Set curIndustryType_ = New BWGIndustryType
        With myWasteDataCR_
            .IndustryTypeID = ""
            .IndustryTypeName = ""
        End With
    Else
        If Trim(curIndustryType_.TypeName) <> Trim(txtIndustryType.Text) Then
            Set curIndustryType_ = Nothing
            Set curIndustryType_ = New BWGIndustryType
            With myWasteDataCR_
                .IndustryTypeID = ""
                .IndustryTypeName = ""
            End With
            Load frmIndustryType
            With frmIndustryType
                .SetCallerForm Me
                .Show 1
            End With
            If myWasteDataCR_.IndustryTypeID = "" Then txtIndustryType.Text = ""
        End If
    End If
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
                .getCriteria (txtMktStaff.Text)
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

Private Sub txtOldStorePlace_Validate(Cancel As Boolean)
    myWasteDataLab_.OldEliminator = Trim(txtOldStorePlace.Text)
End Sub

Private Sub txtOther_Change()
    If Trim(txtOther.Text) <> "" Then
        chkOther.Value = 1
    Else
        chkOther.Value = 0
    End If
End Sub

Private Sub txtOther_Validate(Cancel As Boolean)
    myWasteDataLab_.OtherDesc = Trim(txtOther.Text)
End Sub

Private Sub txtOther2_GotFocus()
    txtOther2.Locked = chkOther2.Value = 0
End Sub

Private Sub txtOther2_Validate(Cancel As Boolean)
    If chkOther2.Value = 1 Then myWasteDataLab_.StorageDesc = Trim(txtOther2.Text)
End Sub

Private Sub txtPesticide_Validate(Cancel As Boolean)
    myWasteDataLab_.PesticideDesc = Trim(txtPesticide.Text)
End Sub

Private Sub txtpH_Validate(Cancel As Boolean)
    If IsNumeric(Trim(txtpH.Text)) Then
        myWasteDataLab_.pHValue = CDbl(Trim(txtpH.Text))
    Else
        myWasteDataLab_.pHValue = 0
    End If
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
End Sub

Private Sub txtRate_Validate(Cancel As Boolean)
    If IsNumeric(Trim(txtRate.Text)) Then
        myWasteDataCR_.WasteRate = CDbl(Trim(txtRate.Text))
    Else
        myWasteDataCR_.WasteRate = 0
    End If
End Sub

Private Sub txtReactDesc_GotFocus()
    SSTab1.Tab = 0
End Sub

Private Sub txtReactDesc_Validate(Cancel As Boolean)
    myWasteDataLab_.ReActWith = Trim(txtReactDesc.Text)
End Sub

Private Sub txtToxicDesc_Validate(Cancel As Boolean)
    myWasteDataLab_.ToxicDesc = Trim(txtToxicDesc.Text)
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
        lvOldTreatment.ListItems.Add , "'" & colTreatment_(x).ID & "'", colTreatment_(x).TreatmentDesc
        lvNewTreatment.ListItems.Add , "'" & colTreatment_(x).ID & "'", colTreatment_(x).TreatmentDesc
    Next
End Sub

Public Sub SetCallerForm(frm As Form, selCust As BWGCustomer)
    Set callerForm_ = frm
    Set curCustomer_ = selCust
    With curCustomer_
        txtCustomer.Text = .CustomerCode & " - " & .CustomerName
        myWasteDataCR_.CustomerID = .ID
        myWasteDataCR_.CustomerName = .CustomerName
    End With
End Sub

Private Sub cmdClear_Click()
Dim selResult As Integer
    If myWasteDataCR_.isChanged Or myWasteDataLab_.isChanged Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนสร้างใหม่หรือไม่?", vbQuestion + vbYesNoCancel)
        If selResult = vbYes Then
            cmdSave_Click
            If myWasteDataCR_.isChanged Or myWasteDataLab_.isChanged Then Exit Sub
        ElseIf selResult = vbCancel Then
            Exit Sub
        End If
    End If
    ClearScreen
    CreateNewWasteData
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
'        DBManager.GetDBConnection.Execute "Update tbWasteDataLab Set WasteDataID='" & tmpGUID & "' Where WasteDataID='" & myWasteDataCR_.ID & "'"
        tmpStr = myWasteDataCR_.Delete
        If Trim(tmpStr) = "" Then
            myWasteDataLab_.Delete
            ClearScreen
            CreateNewWasteData
            txtEUCode1.SetFocus
        Else
            DBManager.GetDBConnection.Execute "Update tbWasteDataLab Set WasteDataID='" & myWasteDataCR_.ID & "' Where WasteDataID='" & tmpGUID & "'"
            MsgBox tmpStr, vbExclamation
        End If
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim selResult As Integer
    If myWasteDataCR_ Is Nothing Or myWasteDataLab_ Is Nothing Then Exit Sub
    If myWasteDataCR_.isChanged Or myWasteDataLab_.isChanged Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนปิดหน้าจอหรือไม่?", vbQuestion + vbYesNoCancel)
        If selResult = vbYes Then
            cmdSave_Click
            If myWasteDataCR_.isChanged Or myWasteDataLab_.isChanged Then Cancel = 1
        ElseIf selResult = vbCancel Then
            Cancel = 1
        End If
    End If
End Sub

Private Sub ClearScreen()
Dim x%
    cmdApprove.Visible = False
    Label1.Visible = False
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
End Sub
Public Sub SetDataWasteData(selWaste As BWGWasteDataCR)
Dim tmpS As New BWGSearchManager
Dim tmpCol As Collection
Dim x%
    Set myWasteDataCR_ = selWaste
    With myWasteDataCR_
        cmdApprove.Visible = Not .isAcceptedByEnvi
        If .isAcceptedByEnvi Then
            Label1.Caption = "Envi ตรวจรับแล้วเมื่อ " & FormatYMD_to_DMY(.acceptedDate, "/", "/")
            Label1.Visible = True
        End If
        txtWasteNo.Text = .WasteNo
        txtContractNo.Text = .Contract.docNo
        Set tmpCol = tmpS.EUCode1Search("EUCode1ID='" & .EUCode1ID & "'")
        If tmpCol.Count > 0 Then SetDataEUCode1 tmpCol(1), False
        Set tmpCol = tmpS.EUCode2Search("EUCode2ID='" & .EUCode2ID & "'")
        If tmpCol.Count > 0 Then SetDataEUCode2 tmpCol(1), False
        Set tmpCol = tmpS.EUCode3Search("EUCode3ID='" & .EUCode3ID & "'")
        If tmpCol.Count > 0 Then SetDataEUCode3 tmpCol(1), False
  
'        If .WasteType = "Non" Then
'            cboWasteType.ListIndex = 1
'        ElseIf .WasteType = "Haz" Then
'            cboWasteType.ListIndex = 2
'        ElseIf .WasteType = "NBL" Then
'            cboWasteType.ListIndex = 3
'        ElseIf .WasteType = "HBL" Then
'            cboWasteType.ListIndex = 4
'        Else
'            cboWasteType.ListIndex = 0
'        End If
        SetTextToCombo .WasteType, cboWasteType
        If cboWasteType.ListIndex = -1 Then cboWasteType.ListIndex = 0
        
        If .isLabResultNeeded Then
            cboLABRes.ListIndex = 1
        Else
            cboLABRes.ListIndex = 0
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
        
        If .Status = "P" Then
            cboSummary.ListIndex = 0
        ElseIf .Status = "A" Then
            cboSummary.ListIndex = 1
        ElseIf .Status = "E" Then
            cboSummary.ListIndex = 2
        End If
        
        If Trim(.TreatmentType) = "" Then
            cboTreatmentType.ListIndex = -1
        ElseIf Trim(.TreatmentType) = "B" Then
            cboTreatmentType.ListIndex = 0
        ElseIf Trim(.TreatmentType) = "L" Then
            cboTreatmentType.ListIndex = 1
        End If
        Set myWasteDataLab_ = Nothing
        Set myWasteDataLab_ = New BWGWasteDataLAB
        myWasteDataLab_.ID = GetGUID
        myWasteDataLab_.wastedataID = myWasteDataCR_.ID
        Set tmpCol = tmpS.WasteDataLABSearch("WasteDataID='" & .ID & "'")
        If tmpCol.Count > 0 Then Set myWasteDataLab_ = tmpCol(1)
        txtHeatValue.Text = Format(.HeatValue, "#,##0")
    End With
    With grKorOrNo.getGridObj
        grKorOrNo.ClearAllData
        .Rows = myWasteDataCR_.KorOrNumbers.Count + 20
        For x = 1 To myWasteDataCR_.KorOrNumbers.Count
            .TextMatrix(x, 0) = myWasteDataCR_.KorOrNumbers(x).ID
            .TextMatrix(x, 1) = myWasteDataCR_.KorOrNumbers(x).KorOrNumber
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
        chkFlame.Value = Abs(CInt(.isFlameAble))
        chkCorrosive = Abs(CInt(.isCorrosive))
        chkReact = Abs(CInt(.isReact))
        chkCancer = Abs(CInt(.isCarcinoGenic))
        chkToxic = Abs(CInt(.isToxic))
        chkRadioActive = Abs(CInt(.isRadioActive))
        chkPesticide = Abs(CInt(.isPesticide))
        If .FlashPoint <> 0 Then txtFlashPoint.Text = .FlashPoint
        If .pHValue <> 0 Then txtpH.Text = .pHValue
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
        txtOldStorePlace.Text = .OldEliminator
        For x = 1 To .OldTreatments.Count
            lvOldTreatment.ListItems("'" & .OldTreatments(x).ID & "'").Checked = isExist(lvOldTreatment.ListItems, "'" & .OldTreatments(x).ID & "'")
        Next
        For x = 1 To .NewTreatments.Count
            lvNewTreatment.ListItems("'" & .NewTreatments(x).ID & "'").Checked = isExist(lvNewTreatment.ListItems, "'" & .NewTreatments(x).ID & "'")
        Next
    End With
    Set myWasteParameter_ = New BWGWasteParameter
    LoadParameter myWasteDataCR_.ID
    myWasteDataLab_.isChanged = False
End Sub

Public Sub SetDataContract(selContract As BWGContract, Optional isAssigned As Boolean = True)
    txtContractNo.Text = selContract.docNo
    If isAssigned Then
        Set myWasteDataCR_.Contract = selContract
    End If
End Sub

Public Sub SetDate(selDate As String)
    myWasteDataCR_.LicenseExpDate = selDate
    txtExpDate.Text = FormatYMD_to_DMY(selDate, "/", "/")
End Sub
Public Sub SetDataIndustryType(selIndustry As BWGIndustryType, Optional isAssigned As Boolean = True)
    Set curIndustryType_ = selIndustry
    With curIndustryType_
        txtIndustryType.Text = .TypeName
        If isAssigned Then
            myWasteDataCR_.IndustryTypeID = .ID
            myWasteDataCR_.IndustryTypeName = .TypeName
        End If
    End With
End Sub

