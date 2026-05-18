VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCustomer 
   BackColor       =   &H00C0C0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "[F023]"
   ClientHeight    =   9435
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11250
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9435
   ScaleWidth      =   11250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtRefCode 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   9270
      TabIndex        =   131
      Top             =   405
      Width           =   1830
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   3210
      Locked          =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   121
      Top             =   8975
      Width           =   2535
   End
   Begin VB.CommandButton cmdReject 
      BackColor       =   &H000000FF&
      Caption         =   "ตีกลับ"
      Height          =   795
      Left            =   1950
      Style           =   1  'Graphical
      TabIndex        =   120
      Top             =   8600
      Width           =   1185
   End
   Begin VB.CommandButton cmdApprove 
      BackColor       =   &H00FFFFC0&
      Caption         =   "ส่งให้ CR ตรวจสอบ"
      Height          =   795
      Left            =   195
      Style           =   1  'Graphical
      TabIndex        =   118
      Top             =   8600
      Width           =   1725
   End
   Begin VB.TextBox txtCustCode 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   9285
      TabIndex        =   87
      Text            =   "AUTO"
      Top             =   90
      Width           =   1800
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H0000C000&
      Caption         =   "เก็บข้อมูล"
      Height          =   795
      Left            =   5820
      Style           =   1  'Graphical
      TabIndex        =   85
      Top             =   8600
      Width           =   1005
   End
   Begin VB.CommandButton cmdSearch 
      BackColor       =   &H00FFFFC0&
      Caption         =   "เปิดข้อมูลเก่า"
      Height          =   795
      Left            =   6870
      Style           =   1  'Graphical
      TabIndex        =   84
      Top             =   8600
      Width           =   1005
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00C0FFC0&
      Caption         =   "สร้างใหม่"
      Height          =   795
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   83
      Top             =   8600
      Width           =   1005
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H000040C0&
      Caption         =   "ปิด"
      Height          =   795
      Left            =   10020
      Style           =   1  'Graphical
      TabIndex        =   82
      Top             =   8600
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H000000FF&
      Caption         =   "ลบข้อมูล"
      Height          =   795
      Left            =   8970
      Style           =   1  'Graphical
      TabIndex        =   81
      Top             =   8600
      Width           =   1005
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7800
      Left            =   -45
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   765
      Width           =   11310
      _ExtentX        =   19950
      _ExtentY        =   13758
      _Version        =   393216
      Style           =   1
      Tabs            =   7
      TabsPerRow      =   7
      TabHeight       =   520
      BackColor       =   12648447
      ForeColor       =   12582912
      TabCaption(0)   =   "ข้อมูลทั่วไป"
      TabPicture(0)   =   "frCustomer.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label2(2)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label2(3)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label2(4)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label2(5)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label2(7)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label2(8)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label2(9)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label2(10)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label2(11)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Image1(1)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label1(2)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label2(14)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label2(20)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label2(21)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label2(23)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label2(26)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label2(27)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Label4"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Label6"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Label2(24)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Label2(35)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Label2(36)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Label2(25)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Label2(6)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Label7"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Label2(12)"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Label2(22)"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "Label49"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "lvWaste"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "txtCompanyNameTH"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "txtAddr"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "txtTumbol"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "txtAumphur"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "txtProvince"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "cboDistrict"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "txtFactoryID"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "txtContactName"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "txtDepartment"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "txtTelNo"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "txtFaxNo"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "cboCustType"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "txtParentCust"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "Command1"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "txtZipCode"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "txtSaleName"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "txtWebSite"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "txtEmail"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).Control(49)=   "cboCustGroup"
      Tab(0).Control(49).Enabled=   0   'False
      Tab(0).Control(50)=   "txtCustRemark"
      Tab(0).Control(50).Enabled=   0   'False
      Tab(0).Control(51)=   "txtSalepst"
      Tab(0).Control(51).Enabled=   0   'False
      Tab(0).Control(52)=   "cboDistrict2"
      Tab(0).Control(52).Enabled=   0   'False
      Tab(0).Control(53)=   "TxtWasteSK2Exp"
      Tab(0).Control(53).Enabled=   0   'False
      Tab(0).Control(54)=   "Txt14Digit"
      Tab(0).Control(54).Enabled=   0   'False
      Tab(0).Control(55)=   "txtIndustryType"
      Tab(0).Control(55).Enabled=   0   'False
      Tab(0).Control(56)=   "txtGenID"
      Tab(0).Control(56).Enabled=   0   'False
      Tab(0).Control(57)=   "txtShortName"
      Tab(0).Control(57).Enabled=   0   'False
      Tab(0).Control(58)=   "CbEstate"
      Tab(0).Control(58).Enabled=   0   'False
      Tab(0).Control(59)=   "cbDept"
      Tab(0).Control(59).Enabled=   0   'False
      Tab(0).Control(60)=   "cbCRStaff"
      Tab(0).Control(60).Enabled=   0   'False
      Tab(0).Control(61)=   "CbCp"
      Tab(0).Control(61).Enabled=   0   'False
      Tab(0).ControlCount=   62
      TabCaption(1)   =   "รายชื่อสถานที่เก็บของเสีย"
      TabPicture(1)   =   "frCustomer.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lvSiteAddress"
      Tab(1).Control(1)=   "Command2"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "ข้อมูลด้านบัญชี"
      TabPicture(2)   =   "frCustomer.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label2(15)"
      Tab(2).Control(1)=   "Label2(16)"
      Tab(2).Control(2)=   "Label2(17)"
      Tab(2).Control(3)=   "Label2(19)"
      Tab(2).Control(4)=   "lvBillAddr"
      Tab(2).Control(5)=   "Command3"
      Tab(2).Control(6)=   "txtAccContact"
      Tab(2).Control(7)=   "txtAccFax"
      Tab(2).Control(8)=   "txtAccTel"
      Tab(2).Control(9)=   "Frame1"
      Tab(2).Control(10)=   "txtRemdesc"
      Tab(2).Control(11)=   "Frame2"
      Tab(2).Control(12)=   "Frame3"
      Tab(2).ControlCount=   13
      TabCaption(3)   =   "สัญญาการให้บริการ"
      TabPicture(3)   =   "frCustomer.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "lvContract"
      Tab(3).Control(1)=   "Command4"
      Tab(3).ControlCount=   2
      TabCaption(4)   =   "รายชื่อเจ้าหน้าที่"
      TabPicture(4)   =   "frCustomer.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "lvStaff"
      Tab(4).Control(1)=   "Command5"
      Tab(4).ControlCount=   2
      TabCaption(5)   =   "หมายเหตุทั่วไป"
      TabPicture(5)   =   "frCustomer.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "lvCustRem"
      Tab(5).Control(1)=   "Command6"
      Tab(5).ControlCount=   2
      TabCaption(6)   =   "รายการใบเสนอราคา"
      TabPicture(6)   =   "frCustomer.frx":00A8
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "ListQuotation"
      Tab(6).ControlCount=   1
      Begin MSComctlLib.ListView ListQuotation 
         Height          =   6615
         Left            =   -74940
         TabIndex        =   145
         Top             =   660
         Width           =   11190
         _ExtentX        =   19738
         _ExtentY        =   11668
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "เลขที่ใบเสนอราคา"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "สถานะ"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "วันที่อนุมัติ"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "AttentionTo"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.ComboBox CbCp 
         Height          =   315
         Left            =   5445
         TabIndex        =   143
         Top             =   3855
         Width           =   2625
      End
      Begin VB.ComboBox cbCRStaff 
         Height          =   315
         Left            =   1545
         Style           =   2  'Dropdown List
         TabIndex        =   141
         Top             =   3855
         Width           =   2430
      End
      Begin VB.ComboBox cbDept 
         Height          =   315
         Left            =   9090
         TabIndex        =   139
         Top             =   2820
         Width           =   2085
      End
      Begin VB.ComboBox CbEstate 
         Height          =   315
         Left            =   1575
         TabIndex        =   137
         Text            =   "Combo1"
         Top             =   2820
         Width           =   2400
      End
      Begin VB.TextBox txtShortName 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   9630
         MaxLength       =   4
         TabIndex        =   18
         Top             =   7380
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.TextBox txtGenID 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   9070
         TabIndex        =   135
         Top             =   2100
         Width           =   2085
      End
      Begin VB.TextBox txtIndustryType 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1560
         TabIndex        =   134
         Top             =   3180
         Width           =   9615
      End
      Begin VB.TextBox Txt14Digit 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   9210
         TabIndex        =   128
         Top             =   1740
         Width           =   1945
      End
      Begin VB.TextBox TxtWasteSK2Exp 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1054
            SubFormatType   =   3
         EndProperty
         Height          =   315
         Left            =   9070
         TabIndex        =   126
         Top             =   2460
         Width           =   2085
      End
      Begin VB.ComboBox cboDistrict2 
         Height          =   315
         Left            =   6345
         TabIndex        =   124
         Top             =   7380
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox txtSalepst 
         BackColor       =   &H00C0E0FF&
         ForeColor       =   &H80000009&
         Height          =   350
         Left            =   7845
         TabIndex        =   123
         Top             =   7380
         Visible         =   0   'False
         Width           =   3060
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H00FFFFC0&
         Caption         =   "เพิ่มข้อมูลหมายเหตุ"
         Height          =   345
         Left            =   -74940
         Style           =   1  'Graphical
         TabIndex        =   116
         Top             =   6945
         Width           =   2205
      End
      Begin VB.Frame Frame3 
         Caption         =   "รายละเอียดการรับเช็ค"
         Height          =   1185
         Left            =   -74790
         TabIndex        =   111
         Top             =   2340
         Width           =   10755
         Begin VB.OptionButton optCheqType 
            Caption         =   "ทุกวัน"
            Height          =   255
            Index           =   1
            Left            =   390
            TabIndex        =   43
            Top             =   300
            Width           =   735
         End
         Begin VB.OptionButton optCheqType 
            Caption         =   "เฉพาะวันที่"
            Height          =   255
            Index           =   2
            Left            =   1980
            TabIndex        =   44
            Top             =   300
            Width           =   1095
         End
         Begin VB.TextBox txtCheqDate 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Index           =   1
            Left            =   3120
            TabIndex        =   45
            Top             =   240
            Width           =   315
         End
         Begin VB.TextBox txtCheqDate 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Index           =   2
            Left            =   3570
            TabIndex        =   46
            Top             =   240
            Width           =   315
         End
         Begin VB.TextBox txtCheqDate 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Index           =   3
            Left            =   4020
            TabIndex        =   47
            Top             =   240
            Width           =   315
         End
         Begin VB.TextBox txtCheqDate 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Index           =   4
            Left            =   4470
            TabIndex        =   48
            Top             =   240
            Width           =   315
         End
         Begin VB.TextBox txtCheqDate 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Index           =   5
            Left            =   4920
            TabIndex        =   49
            Top             =   240
            Width           =   315
         End
         Begin VB.OptionButton optCheqType 
            Caption         =   "ทุกวัน"
            Height          =   255
            Index           =   3
            Left            =   5970
            TabIndex        =   50
            Top             =   300
            Width           =   765
         End
         Begin VB.CheckBox chkCheqDay 
            Caption         =   "จ."
            Height          =   195
            Index           =   1
            Left            =   6900
            TabIndex        =   51
            Top             =   210
            Width           =   465
         End
         Begin VB.CheckBox chkCheqDay 
            Caption         =   "อ."
            Height          =   195
            Index           =   2
            Left            =   7440
            TabIndex        =   52
            Top             =   210
            Width           =   465
         End
         Begin VB.CheckBox chkCheqDay 
            Caption         =   "พ."
            Height          =   195
            Index           =   3
            Left            =   7980
            TabIndex        =   53
            Top             =   210
            Width           =   465
         End
         Begin VB.CheckBox chkCheqDay 
            Caption         =   "พฤ."
            Height          =   195
            Index           =   4
            Left            =   6900
            TabIndex        =   54
            Top             =   450
            Width           =   585
         End
         Begin VB.CheckBox chkCheqDay 
            Caption         =   "ศ."
            Height          =   195
            Index           =   5
            Left            =   7980
            TabIndex        =   55
            Top             =   450
            Width           =   465
         End
         Begin VB.ComboBox cboCheqWeekCount 
            Height          =   315
            Left            =   8640
            Style           =   2  'Dropdown List
            TabIndex        =   56
            Top             =   270
            Width           =   1845
         End
         Begin VB.OptionButton optCheqType 
            Caption         =   "อื่นๆ ระบุ"
            Height          =   255
            Index           =   4
            Left            =   390
            TabIndex        =   57
            Top             =   780
            Width           =   945
         End
         Begin VB.TextBox txtCheqDetail 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   1380
            TabIndex        =   58
            Top             =   750
            Width           =   9105
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   ","
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   34
            Left            =   3480
            TabIndex        =   115
            Top             =   300
            Width           =   45
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   ","
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   33
            Left            =   3930
            TabIndex        =   114
            Top             =   300
            Width           =   45
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   ","
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   32
            Left            =   4380
            TabIndex        =   113
            Top             =   300
            Width           =   45
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   ","
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   31
            Left            =   4830
            TabIndex        =   112
            Top             =   300
            Width           =   45
         End
         Begin VB.Line Line2 
            Index           =   3
            X1              =   1560
            X2              =   1560
            Y1              =   240
            Y2              =   630
         End
         Begin VB.Line Line2 
            Index           =   2
            X1              =   5670
            X2              =   5670
            Y1              =   240
            Y2              =   630
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "รายละเอียดการวางบิล"
         Height          =   1185
         Left            =   -74790
         TabIndex        =   106
         Top             =   1140
         Width           =   10755
         Begin VB.TextBox txtBillingDet 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   1380
            TabIndex        =   42
            Top             =   750
            Width           =   9105
         End
         Begin VB.OptionButton optBillType 
            Caption         =   "อื่นๆ ระบุ"
            Height          =   255
            Index           =   4
            Left            =   390
            TabIndex        =   41
            Top             =   780
            Width           =   945
         End
         Begin VB.ComboBox cboBillWeekCount 
            Height          =   315
            Left            =   8640
            Style           =   2  'Dropdown List
            TabIndex        =   40
            Top             =   270
            Width           =   1845
         End
         Begin VB.CheckBox chkBillDay 
            Caption         =   "ศ."
            Height          =   195
            Index           =   5
            Left            =   7980
            TabIndex        =   39
            Top             =   450
            Width           =   465
         End
         Begin VB.CheckBox chkBillDay 
            Caption         =   "พฤ."
            Height          =   195
            Index           =   4
            Left            =   6900
            TabIndex        =   38
            Top             =   450
            Width           =   585
         End
         Begin VB.CheckBox chkBillDay 
            Caption         =   "พ."
            Height          =   195
            Index           =   3
            Left            =   7980
            TabIndex        =   37
            Top             =   210
            Width           =   465
         End
         Begin VB.CheckBox chkBillDay 
            Caption         =   "อ."
            Height          =   195
            Index           =   2
            Left            =   7440
            TabIndex        =   36
            Top             =   210
            Width           =   465
         End
         Begin VB.CheckBox chkBillDay 
            Caption         =   "จ."
            Height          =   195
            Index           =   1
            Left            =   6900
            TabIndex        =   35
            Top             =   210
            Width           =   465
         End
         Begin VB.OptionButton optBillType 
            Caption         =   "ทุกวัน"
            Height          =   255
            Index           =   3
            Left            =   5970
            TabIndex        =   34
            Top             =   300
            Width           =   765
         End
         Begin VB.TextBox txtBillDate 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Index           =   5
            Left            =   4920
            TabIndex        =   33
            Top             =   240
            Width           =   315
         End
         Begin VB.TextBox txtBillDate 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Index           =   4
            Left            =   4470
            TabIndex        =   32
            Top             =   240
            Width           =   315
         End
         Begin VB.TextBox txtBillDate 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Index           =   3
            Left            =   4020
            TabIndex        =   31
            Top             =   240
            Width           =   315
         End
         Begin VB.TextBox txtBillDate 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Index           =   2
            Left            =   3570
            TabIndex        =   30
            Top             =   240
            Width           =   315
         End
         Begin VB.TextBox txtBillDate 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Index           =   1
            Left            =   3120
            TabIndex        =   29
            Top             =   240
            Width           =   315
         End
         Begin VB.OptionButton optBillType 
            Caption         =   "เฉพาะวันที่"
            Height          =   255
            Index           =   2
            Left            =   1980
            TabIndex        =   28
            Top             =   300
            Width           =   1095
         End
         Begin VB.OptionButton optBillType 
            Caption         =   "ทุกวัน"
            Height          =   255
            Index           =   1
            Left            =   390
            TabIndex        =   27
            Top             =   300
            Width           =   735
         End
         Begin VB.Line Line2 
            Index           =   1
            X1              =   5670
            X2              =   5670
            Y1              =   240
            Y2              =   630
         End
         Begin VB.Line Line2 
            Index           =   0
            X1              =   1560
            X2              =   1560
            Y1              =   240
            Y2              =   630
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   ","
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   30
            Left            =   4830
            TabIndex        =   110
            Top             =   300
            Width           =   45
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   ","
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   29
            Left            =   4380
            TabIndex        =   109
            Top             =   300
            Width           =   45
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   ","
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   28
            Left            =   3930
            TabIndex        =   108
            Top             =   300
            Width           =   45
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   ","
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   18
            Left            =   3480
            TabIndex        =   107
            Top             =   300
            Width           =   45
         End
      End
      Begin VB.TextBox txtCustRemark 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1560
         TabIndex        =   19
         Top             =   3525
         Width           =   9615
      End
      Begin VB.ComboBox cboCustGroup 
         Height          =   315
         ItemData        =   "frCustomer.frx":00C4
         Left            =   2790
         List            =   "frCustomer.frx":00C6
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   7335
         Width           =   2655
      End
      Begin VB.TextBox txtEmail 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   4965
         TabIndex        =   17
         Top             =   2820
         Width           =   2085
      End
      Begin VB.TextBox txtWebSite 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   7725
         TabIndex        =   16
         Top             =   7380
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.TextBox txtSaleName 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   2460
         Width           =   2370
      End
      Begin VB.TextBox txtZipCode 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   9675
         TabIndex        =   6
         Top             =   1035
         Width           =   1500
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00FFFFC0&
         Caption         =   "เพิ่มข้อมูลเจ้าหน้าที่"
         Height          =   300
         Left            =   -74940
         Style           =   1  'Graphical
         TabIndex        =   98
         Top             =   6990
         Width           =   2205
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00C0FFC0&
         Caption         =   "เพิ่มข้อมูลสัญญา"
         Height          =   345
         Left            =   -74940
         Style           =   1  'Graphical
         TabIndex        =   96
         Top             =   6945
         Width           =   2205
      End
      Begin VB.TextBox txtRemdesc 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   -73035
         TabIndex        =   66
         Top             =   4590
         Width           =   8985
      End
      Begin VB.Frame Frame1 
         Caption         =   "เอกสารประกอบที่จำเป็น"
         Height          =   945
         Left            =   -73020
         TabIndex        =   94
         Top             =   3570
         Width           =   8955
         Begin VB.TextBox txtOtherDoc 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   6900
            TabIndex        =   65
            Top             =   360
            Width           =   1755
         End
         Begin VB.CheckBox chkOtherDoc 
            Caption         =   "อื่นๆ"
            Height          =   285
            Left            =   6240
            TabIndex        =   64
            Top             =   390
            Width           =   645
         End
         Begin VB.CheckBox chkQuotation 
            Caption         =   "ใบเสนอราคา"
            Height          =   285
            Left            =   4860
            TabIndex        =   63
            Top             =   390
            Width           =   1215
         End
         Begin VB.CheckBox chkContract 
            Caption         =   "สัญญา"
            Height          =   285
            Left            =   3915
            TabIndex        =   62
            Top             =   390
            Width           =   765
         End
         Begin VB.CheckBox chkBilling 
            Caption         =   "ใบรับวางบิล"
            Height          =   285
            Left            =   2565
            TabIndex        =   61
            Top             =   390
            Width           =   1125
         End
         Begin VB.CheckBox chkInv 
            Caption         =   "ใบกำกับภาษี"
            Height          =   285
            Left            =   1110
            TabIndex        =   60
            Top             =   390
            Width           =   1215
         End
         Begin VB.CheckBox chkPO 
            Caption         =   "PO"
            Height          =   285
            Left            =   240
            TabIndex        =   59
            Top             =   390
            Width           =   585
         End
      End
      Begin VB.TextBox txtAccTel 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   -69780
         TabIndex        =   25
         Top             =   750
         Width           =   2655
      End
      Begin VB.TextBox txtAccFax 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   -66540
         TabIndex        =   26
         Top             =   750
         Width           =   2505
      End
      Begin VB.TextBox txtAccContact 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   -73845
         TabIndex        =   24
         Top             =   750
         Width           =   3375
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00C0FFC0&
         Caption         =   "เพิ่มข้อมูลสถานที่วางบิล"
         Height          =   345
         Left            =   -74940
         Style           =   1  'Graphical
         TabIndex        =   68
         Top             =   6960
         Width           =   2205
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "เพิ่มข้อมูลสถานที่เก็บของเสีย"
         Height          =   345
         Left            =   -74940
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   6990
         Width           =   2205
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFF80&
         Caption         =   "เพิ่มข้อมูลกากของเสีย"
         Height          =   345
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   90
         Top             =   7305
         Width           =   1815
      End
      Begin VB.TextBox txtParentCust 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1560
         TabIndex        =   9
         Top             =   1740
         Width           =   2400
      End
      Begin VB.ComboBox cboCustType 
         Height          =   315
         ItemData        =   "frCustomer.frx":00C8
         Left            =   1560
         List            =   "frCustomer.frx":00CA
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1395
         Width           =   2430
      End
      Begin VB.TextBox txtFaxNo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   4965
         TabIndex        =   14
         Top             =   2460
         Width           =   2085
      End
      Begin VB.TextBox txtTelNo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   4965
         TabIndex        =   13
         Top             =   2100
         Width           =   2085
      End
      Begin VB.TextBox txtDepartment 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   4965
         TabIndex        =   12
         Top             =   1740
         Width           =   2085
      End
      Begin VB.TextBox txtContactName 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   4965
         TabIndex        =   11
         Top             =   1395
         Width           =   2085
      End
      Begin VB.TextBox txtFactoryID 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   9075
         TabIndex        =   10
         Top             =   1395
         Width           =   2085
      End
      Begin VB.ComboBox cboDistrict 
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   2100
         Width           =   2430
      End
      Begin VB.TextBox txtProvince 
         Appearance      =   0  'Flat
         BackColor       =   &H000000C0&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   6600
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1035
         Width           =   1845
      End
      Begin VB.TextBox txtAumphur 
         Appearance      =   0  'Flat
         BackColor       =   &H000000C0&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   3990
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1035
         Width           =   1845
      End
      Begin VB.TextBox txtTumbol 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1560
         TabIndex        =   3
         Top             =   1035
         Width           =   1815
      End
      Begin VB.TextBox txtAddr 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   7515
         TabIndex        =   2
         Top             =   675
         Width           =   3660
      End
      Begin VB.TextBox txtCompanyNameTH 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1545
         TabIndex        =   1
         Top             =   675
         Width           =   4305
      End
      Begin MSComctlLib.ListView lvSiteAddress 
         Height          =   6315
         Left            =   -74940
         TabIndex        =   22
         Top             =   660
         Width           =   11190
         _ExtentX        =   19738
         _ExtentY        =   11139
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ชื่อสถานที่"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ที่อยู่"
            Object.Width           =   11465
         EndProperty
      End
      Begin MSComctlLib.ListView lvBillAddr 
         Height          =   1965
         Left            =   -74940
         TabIndex        =   67
         Top             =   4980
         Width           =   10875
         _ExtentX        =   19182
         _ExtentY        =   3466
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ชื่อสถานที่"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ที่อยู่"
            Object.Width           =   11465
         EndProperty
      End
      Begin MSComctlLib.ListView lvContract 
         Height          =   6270
         Left            =   -74940
         TabIndex        =   97
         Top             =   660
         Width           =   11190
         _ExtentX        =   19738
         _ExtentY        =   11060
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
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "วันที่เริ่มต้น"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "วันที่สิ้นสุด"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "เลขที่สัญญา"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "หมายเหตุ"
            Object.Width           =   7937
         EndProperty
      End
      Begin MSComctlLib.ListView lvStaff 
         Height          =   6315
         Left            =   -74940
         TabIndex        =   99
         Top             =   660
         Width           =   11190
         _ExtentX        =   19738
         _ExtentY        =   11139
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
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ชื่อ-นามสกุล"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ตำแหน่ง"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "แผนก"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "โทร"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Email"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "เรื่องที่ติดต่อ"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "ที่อยู่ในการจัดส่งเอกสาร"
            Object.Width           =   4480
         EndProperty
      End
      Begin MSComctlLib.ListView lvCustRem 
         Height          =   6270
         Left            =   -74940
         TabIndex        =   117
         Top             =   660
         Width           =   11190
         _ExtentX        =   19738
         _ExtentY        =   11060
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "หมายเหตุ"
            Object.Width           =   12347
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ลำดับที่"
            Object.Width           =   3528
         EndProperty
      End
      Begin MSComctlLib.ListView lvWaste 
         Height          =   2775
         Left            =   120
         TabIndex        =   20
         Top             =   4530
         Width           =   11070
         _ExtentX        =   19526
         _ExtentY        =   4895
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
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "เลขที่ของเสีย"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ชื่อของเสีย"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Waste Code"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Waste Type"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "เลขที่สัญญา"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "วันที่หมดอายุ"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label49 
         AutoSize        =   -1  'True
         Caption         =   "นโยบายบริษัท :"
         Height          =   195
         Left            =   4320
         TabIndex        =   144
         Top             =   3900
         Width           =   1095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "เจ้าหน้าที่ CR"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   22
         Left            =   585
         TabIndex        =   142
         Top             =   3930
         Width           =   915
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "แผนก"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   12
         Left            =   8550
         TabIndex        =   140
         Top             =   2865
         Width           =   450
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "อยู่ในนิคมฯ"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   675
         TabIndex        =   138
         Top             =   2865
         Width           =   825
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ประเภทลูกค้า"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   6
         Left            =   540
         TabIndex        =   136
         Top             =   1425
         Width           =   930
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ชื่อย่อ"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   25
         Left            =   9180
         TabIndex        =   103
         Top             =   6900
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "เลขที่ผู้กำเนิดของเสีย"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   36
         Left            =   7515
         TabIndex        =   133
         Top             =   2145
         Width           =   1470
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ประเภทอุตสาหกรรม"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   35
         Left            =   135
         TabIndex        =   130
         Top             =   3225
         Width           =   1500
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "เลขทะเบียนโรงงาน 14 หลัก"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   24
         Left            =   7245
         TabIndex        =   127
         Top             =   1785
         Width           =   1905
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "สก.2 หมดอายุ"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   8055
         TabIndex        =   125
         Top             =   2505
         Width           =   960
      End
      Begin VB.Label Label4 
         Caption         =   "ผู้แทนขาย"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   5505
         TabIndex        =   122
         Top             =   7380
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "หมายเหตุ"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   27
         Left            =   840
         TabIndex        =   105
         Top             =   3570
         Width           =   675
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "กลุ่มลูกค้า"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   26
         Left            =   2010
         TabIndex        =   104
         Top             =   7395
         Width           =   675
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "อีเมลล์"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   23
         Left            =   4455
         TabIndex        =   102
         Top             =   2850
         Width           =   450
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "เซลล์ที่รับผิดชอบ"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   21
         Left            =   285
         TabIndex        =   101
         Top             =   2535
         Width           =   1170
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "รหัสไปรษณีย์"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   20
         Left            =   8655
         TabIndex        =   100
         Top             =   1095
         Width           =   930
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "หมายเหตุ"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   19
         Left            =   -73770
         TabIndex        =   95
         Top             =   4635
         Width           =   675
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "โทร."
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   17
         Left            =   -70140
         TabIndex        =   93
         Top             =   810
         Width           =   315
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "แฟกซ์"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   16
         Left            =   -67050
         TabIndex        =   92
         Top             =   810
         Width           =   450
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ชื่อผู้ติดต่อ"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   15
         Left            =   -74610
         TabIndex        =   91
         Top             =   795
         Width           =   705
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Direct Customer"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   14
         Left            =   330
         TabIndex        =   89
         Top             =   1815
         Width           =   1125
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "รายการกากของเสีย"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   180
         Index           =   2
         Left            =   30
         TabIndex        =   86
         Top             =   4290
         Width           =   10935
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   30
         Index           =   1
         Left            =   0
         Top             =   4230
         Width           =   11325
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "แฟกซ์"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   11
         Left            =   4440
         TabIndex        =   80
         Top             =   2490
         Width           =   450
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "โทร."
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   10
         Left            =   4620
         TabIndex        =   79
         Top             =   2175
         Width           =   315
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "แผนก"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   9
         Left            =   4455
         TabIndex        =   78
         Top             =   1815
         Width           =   450
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ชื่อผู้ติดต่อ"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   8
         Left            =   4185
         TabIndex        =   77
         Top             =   1455
         Width           =   705
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "เลขทะเบียนโรงงาน"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   7
         Left            =   7665
         TabIndex        =   76
         Top             =   1455
         Width           =   1350
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "เขตพื้นที่การขาย"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   5
         Left            =   285
         TabIndex        =   75
         Top             =   2130
         Width           =   1200
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "จังหวัด"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   4
         Left            =   6090
         TabIndex        =   74
         Top             =   1095
         Width           =   465
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "อำเภอ"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   3
         Left            =   3510
         TabIndex        =   73
         Top             =   1095
         Width           =   435
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ตำบล"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   2
         Left            =   1080
         TabIndex        =   72
         Top             =   1095
         Width           =   375
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "เลขที่, หมู่, ถนน"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   1
         Left            =   6315
         TabIndex        =   71
         Top             =   735
         Width           =   1095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ชื่อบริษัท/โรงงาน"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   0
         Left            =   270
         TabIndex        =   70
         Top             =   720
         Width           =   1200
      End
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "รหัส Better ME"
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   8190
      TabIndex        =   132
      Top             =   495
      Width           =   1035
   End
   Begin VB.Label LblVIPCR 
      BackColor       =   &H00C0FFC0&
      Caption         =   "ลูกค้า VIP CR"
      BeginProperty Font 
         Name            =   "AngsanaUPC"
         Size            =   21.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   555
      Left            =   6210
      TabIndex        =   129
      Top             =   135
      Visible         =   0   'False
      Width           =   1770
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "EucrosiaUPC"
         Size            =   18
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   3210
      TabIndex        =   119
      Top             =   8600
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "รหัสลูกค้า"
      ForeColor       =   &H000000C0&
      Height          =   195
      Index           =   13
      Left            =   8490
      TabIndex        =   88
      Top             =   150
      Width           =   795
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
      Caption         =   "ข้อมูลลูกค้า"
      BeginProperty Font 
         Name            =   "Angsana New"
         Size            =   15.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   435
      Index           =   1
      Left            =   270
      TabIndex        =   69
      Top             =   60
      Width           =   1005
   End
   Begin VB.Image Image2 
      Height          =   795
      Left            =   0
      Picture         =   "frCustomer.frx":00CC
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11280
   End
End
Attribute VB_Name = "frmCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim curTumbol_ As IMITumbol
Dim curIndustryType_ As BWGIndustryType_New
Dim colDistrict_ As Collection
Dim colEstate_ As Collection
Dim colDistrict2_ As Collection
Dim curParentCust_ As BWGCustomer
Dim curSaleStaff_ As BWGCompanyStaff
Dim curCRstaff_ As BWGCompanyStaff
Dim myCustomer_ As BWGCustomer
Dim callLocation_ As String
Dim colCustGroup_ As Collection
Dim isNew As Boolean
Dim isCR As Boolean
Dim isMKT As Boolean
Public Capt As String
Dim colCR_ As Collection

Private Function RequestData() As Boolean
    RequestData = True
    If Trim(txtCompanyNameTH.Text) = "" Then
        MsgBox "กรุณาระบุชื่อบริษัทหรือโรงงานของลูกค้า", vbExclamation
        txtCompanyNameTH.SetFocus
        RequestData = False
        Exit Function
    End If
    If Trim(txtTumbol.Text) = "" Then
        MsgBox "กรุณาระบุตำบลที่ตั้งของลูกค้า", vbExclamation
        txtTumbol.SetFocus
        RequestData = False
        Exit Function
    End If
    If Trim(cboCustGroup.Text) = "" Then
        MsgBox "กรุณาระบุกลุ่มลูกค้า", vbExclamation
        cboCustGroup.SetFocus
        RequestData = False
        Exit Function
    End If
    If cboDistrict.ListIndex = 0 Then
        MsgBox "กรุณาระบุโซนพื้นที่ ที่ตั้งของลูกค้า", vbExclamation
        cboDistrict.SetFocus
        RequestData = False
        Exit Function
    End If
    If cboCustType.ListIndex = 0 Then
        MsgBox "กรุณาระบุประเภทของลูกค้า", vbExclamation
        cboCustType.SetFocus
        RequestData = False
        Exit Function
    End If
    If myCustomer_.CustomerType = "S" And Trim(txtParentCust.Text) = "" Then
        MsgBox "กรุณาระบุ Direct Customer ที่บริษัทนี้เกี่ยวข้องอยู่", vbExclamation
        txtParentCust.SetFocus
        RequestData = False
        Exit Function
    End If
End Function

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
    
    If LCase(StrGrp) = "customer relation" Or LCase(StrGrp) = "senior cr" Or LCase(StrGrp) = "Account" Or LCase(StrGrp) = "account" Or LCase(StrGrp) = "administrator" Or LCase(StrGrp) = "environment" Then
        isCR = True
        cmdApprove.Caption = "ตรวจรับโดย CR"
    Else
        isCR = False
        cmdApprove.Caption = "ส่งให้ CR ตรวจสอบ"
        cmdReject.Visible = False
    End If
    If LCase(StrGrp) = "incomes" Or LCase(StrGrp) = "marketing" Then
        isMKT = True
    End If
'    If LCase(StrGrp) = "customer relation" Or LCase(StrGrp) = "senior cr" Or LCase(StrGrp) = "account" Or LCase(StrGrp) = "administrator" Then
'        cmdSave.Enabled = True
'    Else
'        cmdSave.Enabled = False
'    End If
    cmdApprove.Picture = frmPicture.picApprove.Picture
    cmdReject.Picture = frmPicture.picDel.Picture
End Sub
Private Sub CbCp_Validate(Cancel As Boolean)
    If Trim(CbCp.Text) <> "" Then
        myCustomer_.CompanyPolicy = Trim(CbCp.Text)
    End If
End Sub

Private Sub cbCRStaff_Validate(Cancel As Boolean)
Dim tmpCRStaff As BWGCompanyStaff
    With myCustomer_
        If .CustomerCode = "" Then
            If cbCRStaff.ListIndex > 0 Then
                Set tmpCRStaff = colCR_(cbCRStaff.ListIndex)
                .CRStaffID = tmpCRStaff.ID
            Else
                .CRStaffID = ""
            End If
        Else
            If cbCRStaff.ListIndex > 0 Then
                Set tmpCRStaff = colCR_(cbCRStaff.ListIndex)
                .CRStaffID = tmpCRStaff.ID
            End If
        End If
    End With
End Sub

Private Sub cbDept_Validate(Cancel As Boolean)
    With myCustomer_
        If cbDept.ListIndex > 0 Then
            .CustdeptName = cbDept.Text
        End If
    End With
End Sub

Private Sub CbEstate_Validate(Cancel As Boolean)
Dim tmpEstate As New BWGEstate2
    With myCustomer_
        If CbEstate.ListIndex > 0 Then
            Set tmpEstate = colEstate_(CbEstate.ListIndex)
            .EstateID = tmpEstate.ID
        Else
            .EstateID = "0"
        End If
    End With

End Sub

Private Sub cboBillWeekCount_Validate(Cancel As Boolean)
    myCustomer_.BillingInfo.BillWeekCount = cboBillWeekCount.ListIndex
    If myCustomer_.BillingInfo.isChanged Then myCustomer_.isChanged = True
End Sub

Private Sub cboCheqWeekCount_Validate(Cancel As Boolean)
    myCustomer_.BillingInfo.CheqWeekCount = cboCheqWeekCount.ListIndex
    If myCustomer_.BillingInfo.isChanged Then myCustomer_.isChanged = True
End Sub

Private Sub cboCustGroup_Validate(Cancel As Boolean)
    If cboCustGroup.ListIndex = -1 Then
        myCustomer_.CustGroupID = ""
        myCustomer_.CustGroupName = ""
    Else
        myCustomer_.CustGroupID = colCustGroup_(cboCustGroup.ListIndex + 1).ID
        myCustomer_.CustGroupName = Trim(cboCustGroup.Text)
    End If
End Sub

Private Sub cboCustType_Click()
            On Error Resume Next
    If cboCustType.ListIndex = 2 Or cboCustType.ListIndex = 3 Or cboCustType.ListIndex = 4 Or cboCustType.ListIndex = 5 Or cboCustType.ListIndex = 6 Or cboCustType.ListIndex = 7 Or cboCustType.ListIndex = 8 Then
        txtParentCust.Enabled = True
    Else
        txtParentCust.Enabled = False
    End If
        If cboCustType.ListIndex = 3 Then
            txtParentCust.Text = "เบตเตอร์ มี จำกัด"
            txtParentCust.Enabled = True
            txtParentCust.SetFocus
          ElseIf cboCustType.ListIndex = 4 Then
            txtParentCust.Enabled = True
             txtParentCust.Text = ""
            txtParentCust.SetFocus
           Else
            txtParentCust.Text = ""
        End If
End Sub

Private Sub cboCustType_Validate(Cancel As Boolean)
    Select Case cboCustType.ListIndex
        Case 0
            myCustomer_.CustomerType = ""
        Case 1
            myCustomer_.CustomerType = "C"
        Case 2
            myCustomer_.CustomerType = "S"
          Case 3
          myCustomer_.CustomerType = "M"
          Case 4
          myCustomer_.CustomerType = "MS"
          Case 5
          myCustomer_.CustomerType = "B"
          Case 6
          myCustomer_.CustomerType = "E"
          Case 7
          myCustomer_.CustomerType = "MI"
          Case 8
          myCustomer_.CustomerType = "I"
    End Select

End Sub

Private Sub cboDistrict_Click()
On Error Resume Next
    Dim tmpDistrict As New BWGSaleDistrict
    If myCustomer_.CustomerCode <> "" Then
'        If LCase(CurrentUser.StaffFName) = "admin" Then
            With myCustomer_
                If cboDistrict.ListIndex > 0 Then
                    Set tmpDistrict = colDistrict_(cboDistrict.ListIndex)
                    txtSaleName.Text = tmpDistrict.SaleStaffName
        '                If tmpDistrict.ZoneName = "BME_A" Or tmpDistrict.ZoneName = "BME_B" Or tmpDistrict.ZoneName = "BME_C" Or tmpDistrict.ZoneName = "BME_D" Then
        '                        If Capt = "N" Then
        '                            cboCustType.ListIndex = 11
        '                            myCustomer_.CustomerType = "M"
        '                            txtParentCust.Enabled = True
        '                            txtParentCust.Text = "เบตเตอร์ มี จำกัด"
        '                            txtContactName.Text = "คุณจารุวรรณ โพธิ์แจ้ง"
        '                            myCustomer_.ContactPerson = txtContactName.Text
        '                            txtTelNo.Text = "0-2012-7888"
        '                            myCustomer_.TelNo = txtTelNo.Text
        '                            txtFaxNo.Text = "0-2012-7889"
        '                            myCustomer_.FaxNo = txtFaxNo.Text
        '                            txtParentCust.Enabled = True
        '                            txtParentCust.SetFocus
        '                       Else
        '                            txtContactName.Text = myCustomer_.ContactPerson
        '                            txtTelNo.Text = myCustomer_.TelNo
        '                            txtFaxNo.Text = myCustomer_.FaxNo
        '                        End If
        '                End If
                End If
                    .SaleDistrictID = tmpDistrict.ID
                    .UnderSaleStaffID = tmpDistrict.SaleStaffID
            End With
'            Else
''                MsgBox "คุณไม่สามารถแก้ไขข้อมูลนี้ได้ กรุณาติดต่อแผนกสารสนเทศ", vbOKOnly + vbInformation, "ไม่อนุญาตให้แก้ไขข้อมูล"
'                Exit Sub
'        End If
    Else
        With myCustomer_
            If cboDistrict.ListIndex > 0 Then
                Set tmpDistrict = colDistrict_(cboDistrict.ListIndex)
                txtSaleName.Text = tmpDistrict.SaleStaffName
            End If
                .SaleDistrictID = tmpDistrict.ID
                .UnderSaleStaffID = tmpDistrict.SaleStaffID
        End With
    End If
End Sub

Private Sub cboDistrict_Validate(Cancel As Boolean)
Dim tmpDistrict As New BWGSaleDistrict
    If myCustomer_.CustomerCode <> "" Then
'        If LCase(CurrentUser.StaffFName) = "admin" Then
            With myCustomer_
                If cboDistrict.ListIndex > 0 Then
                    Set tmpDistrict = colDistrict_(cboDistrict.ListIndex)
                    .SaleDistrictID = tmpDistrict.ID
                    .SaleDistrictName = tmpDistrict.DistrictName
                    .UnderSaleName = tmpDistrict.SaleStaffName
                    .UnderSaleStaffID = tmpDistrict.SaleStaffID
                    txtSaleName.Text = tmpDistrict.SaleStaffName
                Else
                    .SaleDistrictID = ""
                    .SaleDistrictName = ""
                    txtSaleName.Text = ""
                    .UnderSaleName = ""
                    .UnderSaleStaffID = ""
                End If
            End With
'        Else
''            MsgBox "คุณไม่สามารถแก้ไขข้อมูลนี้ได้ กรุณาติดต่อแผนกสารสนเทศ", vbOKOnly + vbInformation, "ไม่สามารถแก้ไขข้อมูล"
'            Exit Sub
'        End If
    Else
             With myCustomer_
                If cboDistrict.ListIndex > 0 Then
                    Set tmpDistrict = colDistrict_(cboDistrict.ListIndex)
                    .SaleDistrictID = tmpDistrict.ID
                    .SaleDistrictName = tmpDistrict.DistrictName
                    .UnderSaleName = tmpDistrict.SaleStaffName
                    .UnderSaleStaffID = tmpDistrict.SaleStaffID
                    txtSaleName.Text = tmpDistrict.SaleStaffName
                Else
                    .SaleDistrictID = ""
                    .SaleDistrictName = ""
                    txtSaleName.Text = ""
                    .UnderSaleName = ""
                    .UnderSaleStaffID = ""
                End If
            End With
    End If
End Sub

Private Sub cboDistrict2_Validate(Cancel As Boolean)
Dim tmpDistrict2 As New BWGSaleDistrict
    With myCustomer_
        If cboDistrict2.ListIndex > 0 Then
            Set tmpDistrict2 = colDistrict2_(cboDistrict2.ListIndex)
            .Sale2ID = tmpDistrict2.ID
            .Sale2Name = tmpDistrict2.SaleStaffName
            txtSalepst.Text = tmpDistrict2.SaleStaffName
        Else
            .Sale2ID = ""
            .Sale2Name = ""
            txtSaleName.Text = ""
        End If
    End With
End Sub

Private Sub chkBillDay_Validate(Index As Integer, Cancel As Boolean)
Dim tmpVal As Integer
Dim x%
    For x = 1 To chkBillDay.UBound
        If chkBillDay(x).Value = 1 Then
            tmpVal = x
        Else
            tmpVal = 0
        End If
        Select Case x
            Case 1
                myCustomer_.BillingInfo.BillDay1 = tmpVal
            Case 2
                myCustomer_.BillingInfo.BillDay2 = tmpVal
            Case 3
                myCustomer_.BillingInfo.BillDay3 = tmpVal
            Case 4
                myCustomer_.BillingInfo.BillDay4 = tmpVal
            Case 5
                myCustomer_.BillingInfo.BillDay5 = tmpVal
        End Select
    Next
    If myCustomer_.BillingInfo.isChanged Then myCustomer_.isChanged = True
End Sub

Private Sub chkBilling_Validate(Cancel As Boolean)
    myCustomer_.BillingInfo.NeedBillingNote = chkBilling.Value = 1
    If myCustomer_.BillingInfo.isChanged Then myCustomer_.isChanged = True
End Sub

Private Sub chkCheqDay_Validate(Index As Integer, Cancel As Boolean)
Dim tmpVal As Integer
Dim x%
    For x = 1 To chkCheqDay.UBound
        If chkCheqDay(x).Value = 1 Then
            tmpVal = x
        Else
            tmpVal = 0
        End If
        Select Case x
            Case 1
                myCustomer_.BillingInfo.CheqDay1 = tmpVal
            Case 2
                myCustomer_.BillingInfo.CheqDay2 = tmpVal
            Case 3
                myCustomer_.BillingInfo.CheqDay3 = tmpVal
            Case 4
                myCustomer_.BillingInfo.CheqDay4 = tmpVal
            Case 5
                myCustomer_.BillingInfo.CheqDay5 = tmpVal
        End Select
    Next
    If myCustomer_.BillingInfo.isChanged Then myCustomer_.isChanged = True
End Sub

Private Sub chkContract_Validate(Cancel As Boolean)
    myCustomer_.BillingInfo.NeedContract = chkContract.Value = 1
    If myCustomer_.BillingInfo.isChanged Then myCustomer_.isChanged = True
End Sub

Private Sub chkInv_Validate(Cancel As Boolean)
    myCustomer_.BillingInfo.NeedInvoice = chkInv.Value = 1
    If myCustomer_.BillingInfo.isChanged Then myCustomer_.isChanged = True
End Sub

Private Sub chkPO_Validate(Cancel As Boolean)
    myCustomer_.BillingInfo.NeedPO = chkPO.Value = 1
    If myCustomer_.BillingInfo.isChanged Then myCustomer_.isChanged = True
End Sub

Private Sub chkQuotation_Validate(Cancel As Boolean)
    myCustomer_.BillingInfo.NeedQuotation = chkQuotation.Value = 1
    If myCustomer_.BillingInfo.isChanged Then myCustomer_.isChanged = True
End Sub

Private Sub cmdApprove_Click()
    If RequestData = False Then Exit Sub
    If isCR = True Then
        If MsgBox("ยืนยันข้อมูลลูกค้า", vbInformation + vbYesNo, "ยืนยัน") = vbNo Then Exit Sub
        myCustomer_.isChanged = True
        isNew = False
        CmdSave_Click
        myCustomer_.Request
        myCustomer_.CRApprove
        cmdApprove.Visible = False
        cmdReject.Visible = False
        Label3.Caption = "รับเป็นลูกค้าแล้ว"
        Label3.Visible = True
        isNew = False
    Dim tSql As String
    Dim rs As ADODB.Recordset
    tSql = "SELECT * FROM tbCustomer" & vbCrLf
    tSql = tSql & " WHERE CustomerID = '" & myCustomer_.ID & "'"
    Set rs = New ADODB.Recordset
    Set rs = GetRS(tSql, adOpenForwardOnly, adLockReadOnly)
    txtCustCode.Text = rs!CustomerCodeDemo
    
    Else
        If MsgBox("ยืนยันการส่งข้อมูลลูกค้า", vbInformation + vbYesNo, "ยืนยัน") = vbNo Then Exit Sub
        myCustomer_.Request
        cmdApprove.Visible = False
        Label3.Caption = "รอการตรวจสอบ"
        Label3.Visible = True
        Text1 = ""
        Text1.Visible = False
        myCustomer_.isChanged = False
    End If
End Sub

Private Sub cmdCancel_Click()
Dim tmpStr As String
    If MsgBox("ข้อมูลนี้จะถูกลบเฉพาะในฐานของ Waste Data (โปรแกรมตัวใหม่) เท่านั้น" & vbCrLf & "ต้องการลบข้อมูลลูกค้านี้หรือไม่?", vbQuestion + vbYesNo) = vbYes Then
        tmpStr = myCustomer_.Delete
        If Trim(tmpStr) = "" Then
            ClearScreen
            CreateNewCustomer
            txtCompanyNameTH.SetFocus
        Else
            MsgBox tmpStr, vbExclamation
        End If
    End If
End Sub

Private Sub cmdClear_Click()
If Label3.Caption = "รับเป็นลูกค้าแล้ว" Then
        If isCR <> True Then
            Exit Sub
        End If
End If
Dim selResult As Integer
    isNew = True
    cmdApprove.Visible = True
    If myCustomer_.isChanged Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนสร้างใหม่หรือไม่?", vbQuestion + vbYesNoCancel)
        If selResult = vbYes Then
            CmdSave_Click
            If myCustomer_.isChanged Then Exit Sub
        ElseIf selResult = vbCancel Then
            Exit Sub
        End If
    End If
    ClearScreen
    CreateNewCustomer
    txtCompanyNameTH.SetFocus
    Capt = "N"
End Sub

Private Sub CmdClose_Click()
    Unload Me
End Sub

Private Sub cmdReject_Click()
Dim RejectDetail As String
    If isCR = True Then
        If MsgBox("ยืนยันการตีกลับ", vbInformation + vbYesNo, "ยืนยัน") = vbNo Then Exit Sub
        RejectDetail = InputBox("ใส่รายละเอียดในการตีกลับ", "ใส่รายละเอียด", "")
        myCustomer_.Reject Trim(RejectDetail)
        cmdReject.Visible = False
        cmdApprove.Visible = False
        cmdSave.Enabled = False
        cmdCancel.Enabled = False
        Label3.Caption = "ตีกลับ"
        Label3.Visible = True
        Text1 = RejectDetail
        Text1.Visible = True
    Else
    End If
End Sub

Private Sub CmdSave_Click()
If Label3.Caption = "รับเป็นลูกค้าแล้ว" Then
        If isCR <> True Then
            Exit Sub
        End If
End If

If myCustomer_.SaleDistrictID = "" Then
    MsgBox "กรุณาเลือกเขตการขายให้ถูกต้อง", vbOKOnly + vbExclamation, "ไม่พบข้อมูลเขตการขาย"
    Exit Sub
End If

If myCustomer_.UnderSaleStaffID = "" Then
    MsgBox "กรุณาเลือกพนักงานขายให้ถูกต้อง", vbOKOnly + vbExclamation, "ไม่พบข้อมูลพนักงานขาย"
    Exit Sub
End If

    If RequestData = False Then Exit Sub
    If myCustomer_.CustomerType = "L" Or myCustomer_.CustomerType = "P" Then
      MsgBox "ไม่สามารถคีย์ข้อมูลใน Waste Data ได้ ให้คีย์ในโปรแกรม BWG LAP "
      Exit Sub
    End If
    If cboCustType.ListIndex = 0 Then
      MsgBox "กรุณาระบุประเภทลูกค้า "
      Exit Sub
    End If
    If myCustomer_.IndustryTypeIDNew = "" Then
        MsgBox "กรุณาระบุประเภทอุตสาหกรรม", vbOKOnly + vbExclamation, "ไม่พบข้อมูลประเภทอุตสาหกรรม"
        txtIndustryType.SetFocus
        Exit Sub
    End If
    If txtFactoryID.Text = "" Then
        MsgBox "กรุณาระบุเลขทะเบียนโรงงาน", vbOKOnly + vbExclamation, "ไม่พบข้อมูลเลขทะเบียนโรงงาน"
        txtIndustryType.SetFocus
        Exit Sub
    End If
        Select Case cboCustType.ListIndex
'        Case 0
'            myCustomer_.CustomerType = ""
        Case 1
            myCustomer_.CustomerType = "C"
        Case 2
            myCustomer_.CustomerType = "S"
        Case 3
          myCustomer_.CustomerType = "M"
          Case 4
          myCustomer_.CustomerType = "MS"
          Case 5
          myCustomer_.CustomerType = "B"
          Case 6
          myCustomer_.CustomerType = "E"
           Case 7
          myCustomer_.CustomerType = "MI"
          Case 8
          myCustomer_.CustomerType = "I"
    End Select

    If myCustomer_.isChanged Then
        myCustomer_.isNew = isNew
              myCustomer_.Save
              If myCustomer_.CustomerType = "R" Then
                        myCustomer_.CRApprove
                        cmdApprove.Enabled = False
              End If
            If Trim(txtCustCode.Text) <> Trim(myCustomer_.CustomerCode) Then
                ShowSiteAddress
                ShowBillAddress
            End If
        If myCustomer_.isNew = True Then
            txtCustCode.Text = "AUTO"
        Else
            txtCustCode.Text = myCustomer_.CustomerCode
        End If
    End If
    If myCustomer_.isChanged = False Then MsgBox "จัดเก็บข้อมูลเรียบร้อยแล้ว", vbInformation
           cmdApprove.Visible = True
End Sub
Private Sub cmdSearch_Click()
Dim selResult As Integer
    If myCustomer_.isChanged Then
    If Label3.Caption = "รับเป็นลูกค้าแล้ว" Then
        If isCR <> True Then
            Exit Sub
        End If
        End If
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนทำงานต่อหรือไม่?", vbQuestion + vbYesNoCancel)
        If selResult = vbYes Then
            CmdSave_Click
            If myCustomer_.isChanged Then Exit Sub
        ElseIf selResult = vbCancel Then
            Exit Sub
        End If
    End If
    callLocation_ = "N" 'Normal
    Load frmCustomerSearch
    With frmCustomerSearch
        .SetCallerForm Me
        .Show 1
    End With
End Sub

Private Sub Combo1_Change()

End Sub

Private Sub Combo1_Validate(Cancel As Boolean)
Dim tmpDistrict As BWGSaleDistrict
    With myCustomer_
        If cboDistrict.ListIndex > 0 Then
            Set tmpDistrict = colDistrict_(cboDistrict.ListIndex)
            .SaleDistrictID = tmpDistrict.ID
            .SaleDistrictName = tmpDistrict.DistrictName
            .UnderSaleName = tmpDistrict.SaleStaffName
            .UnderSaleStaffID = tmpDistrict.SaleStaffID
            txtSaleName.Text = tmpDistrict.SaleStaffName
        Else
            .SaleDistrictID = ""
            .SaleDistrictName = ""
            txtSaleName.Text = ""
            .UnderSaleName = ""
            .UnderSaleStaffID = ""
        End If
    End With
End Sub

Private Sub Command1_Click()
Dim selResult As Integer
    If Trim(txtCompanyNameTH.Text) = "" Then
        MsgBox "กรุณาระบุชื่อบริษัท/โรงงาน", vbExclamation
        txtCompanyNameTH.SetFocus
        Exit Sub
    End If
    If myCustomer_.isChanged Then
    
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนสร้างใหม่หรือไม่?", vbQuestion + vbYesNo)
        If selResult = vbYes Then
            CmdSave_Click
            If myCustomer_.isChanged Then Exit Sub
        Else 'If selResult = vbCancel Then
            Exit Sub
        End If
    End If
    Load frmWasteData
    With frmWasteData
        .SetCallerForm Me, myCustomer_
        .Show 1
        myCustomer_.RequeryWasteItems
        ShowWasteData
    End With
End Sub

Private Sub Command2_Click()
If Label3.Caption = "รับเป็นลูกค้าแล้ว" Then
        If isCR <> True Then
            Exit Sub
        End If
End If
Dim selResult As Integer
    If Trim(txtCompanyNameTH.Text) = "" Then
        MsgBox "กรุณาระบุชื่อบริษัท/โรงงาน", vbExclamation
        SSTab1.Tab = 0
        txtCompanyNameTH.SetFocus
        Exit Sub
    End If
    If myCustomer_.isChanged Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนสร้างใหม่หรือไม่?", vbQuestion + vbYesNo)
        If selResult = vbYes Then
            CmdSave_Click
            If myCustomer_.isChanged Then Exit Sub
        Else 'If selResult = vbCancel Then
            Exit Sub
        End If
    End If
    Load frmCustomerSite
    With frmCustomerSite
        .SetCallerForm Me, myCustomer_
        .Show 1
        myCustomer_.RequerySiteAddress
        ShowSiteAddress
    End With
End Sub

Private Sub Command3_Click()
If Label3.Caption = "รับเป็นลูกค้าแล้ว" Then
        If isCR <> True Then
            Exit Sub
        End If
End If
Dim selResult As Integer
    If Trim(txtCompanyNameTH.Text) = "" Then
        MsgBox "กรุณาระบุชื่อบริษัท/โรงงาน", vbExclamation
        SSTab1.Tab = 0
        txtCompanyNameTH.SetFocus
        Exit Sub
    End If
    If myCustomer_.isChanged Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนสร้างใหม่หรือไม่?", vbQuestion + vbYesNo)
        If selResult = vbYes Then
            CmdSave_Click
            If myCustomer_.isChanged Then Exit Sub
        Else 'If selResult = vbCancel Then
            Exit Sub
        End If
    End If
'    Load frCustomerBillAddr
'    With frCustomerBillAddr
'        .SetCallerForm Me, myCustomer_
'        .Show 1
'        myCustomer_.RequeryBillAddress
'        ShowBillAddress
'    End With
End Sub

Private Sub Command4_Click()
If Label3.Caption = "รับเป็นลูกค้าแล้ว" Then
        If isCR <> True Then
            Exit Sub
        End If
End If
Dim selResult As Integer
    If Trim(txtCompanyNameTH.Text) = "" Then
        MsgBox "กรุณาระบุชื่อบริษัท/โรงงาน", vbExclamation
        SSTab1.Tab = 0
        txtCompanyNameTH.SetFocus
        Exit Sub
    End If
    If myCustomer_.isChanged Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนสร้างใหม่หรือไม่?", vbQuestion + vbYesNo)
        If selResult = vbYes Then
            CmdSave_Click
            If myCustomer_.isChanged Then Exit Sub
        Else 'If selResult = vbCancel Then
            Exit Sub
        End If
    End If
    Load frmContract
    With frmContract
        .SetCallerForm Me, myCustomer_
        .Show 1
        myCustomer_.RequeryContracts
        ShowContracts
    End With
End Sub

Private Sub Command5_Click()
If Label3.Caption = "รับเป็นลูกค้าแล้ว" Then
        If isCR <> True Then
            Exit Sub
        End If
End If
Dim selResult As Integer
    If Trim(txtCompanyNameTH.Text) = "" Then
        MsgBox "กรุณาระบุชื่อบริษัท/โรงงาน", vbExclamation
        SSTab1.Tab = 0
        txtCompanyNameTH.SetFocus
        Exit Sub
    End If
    If myCustomer_.isChanged Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนสร้างใหม่หรือไม่?", vbQuestion + vbYesNo)
        If selResult = vbYes Then
            CmdSave_Click
            If myCustomer_.isChanged Then Exit Sub
        Else 'If selResult = vbCancel Then
            Exit Sub
        End If
    End If
    Load frmCustomerStaff
    With frmCustomerStaff
        .SetCallerForm Me, myCustomer_
        .Show 1
        myCustomer_.RequeryCustStaff
        ShowCustStaff
    End With
End Sub

Private Sub Command6_Click()
If Label3.Caption = "รับเป็นลูกค้าแล้ว" Then
        If isCR <> True Then
            Exit Sub
        End If
End If
Dim selResult As Integer
    If Trim(txtCompanyNameTH.Text) = "" Then
        MsgBox "กรุณาระบุชื่อบริษัท/โรงงาน", vbExclamation
        SSTab1.Tab = 0
        txtCompanyNameTH.SetFocus
        Exit Sub
    End If
    If myCustomer_.isChanged Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนสร้างใหม่หรือไม่?", vbQuestion + vbYesNo)
        If selResult = vbYes Then
            CmdSave_Click
            If myCustomer_.isChanged Then Exit Sub
        Else 'If selResult = vbCancel Then
            Exit Sub
        End If
    End If
'    Load frCustOtherRemark
'    With frCustOtherRemark
'        .SetCallerForm Me, myCustomer_
'        .Show 1
'        myCustomer_.RequeryOtherRemark
'        ShowOtherRemark
'    End With
End Sub

Private Sub Form_Load()
Dim i As Integer
    Dim tmpRec As New ADODB.Recordset
    cbDept.Clear
    Set tmpRec = GetRS("select * from custdept ORDER BY DeptID", adOpenForwardOnly, adLockReadOnly)
    cbDept.AddItem "== เลือกแผนก =="
    Do Until tmpRec.EOF
        cbDept.AddItem tmpRec!DeptName
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    cbDept.ListIndex = 0
Capt = "N"
    With CbCp
        .Clear
        .AddItem "ไม่มี"
        .AddItem "No landfill"
        .AddItem "ESG"
        .AddItem "ต้องการเผาเท่านั้น"
        .AddItem "ใช้ขยะเพื่อเป็นพลังงาน"
        .ListIndex = 0
    End With
    With frmPicture
        cmdSave.Picture = .picSave.Picture
        cmdSearch.Picture = .picSearch.Picture
        cmdClear.Picture = .picNew.Picture
        cmdCancel.Picture = .picDelete.Picture
        CmdClose.Picture = .picExit.Picture
        cmdApprove.Picture = .picApprove.Picture
        cmdReject.Picture = .picDel.Picture
    End With
    With cboBillWeekCount
        .AddItem "ของทุกสัปดาห์"
        .AddItem "ของสัปดาห์ที่ 1"
        .AddItem "ของสัปดาห์ที่ 2"
        .AddItem "ของสัปดาห์ที่ 3"
        .AddItem "ของสัปดาห์ที่ 4"
    End With
    With cboCheqWeekCount
        .AddItem "ของทุกสัปดาห์"
        .AddItem "ของสัปดาห์ที่ 1"
        .AddItem "ของสัปดาห์ที่ 2"
        .AddItem "ของสัปดาห์ที่ 3"
        .AddItem "ของสัปดาห์ที่ 4"
    End With
    With cboCustType
        .AddItem "== ประเภทลูกค้า =="
        .AddItem "(C) Direct Customer"
        .AddItem "(S) Sub Customer"
        .AddItem "(M) Better Me"
        .AddItem "(MS) Sub Better Me"
        .AddItem "(B) Broker"
        .AddItem "(E) Sub Broker"
        .AddItem "(MI) Better ME IS"
        .AddItem "(I) IS"
        .ListIndex = 0
    End With
    Set curTumbol_ = Nothing
    Set curTumbol_ = New IMITumbol
    Set curIndustryType_ = Nothing
    Set curIndustryType_ = New BWGIndustryType_New
    Set curParentCust_ = New BWGCustomer
    PopulateDistrict
    PopulateCustGroup
    PopulateCR
    CreateNewCustomer
    PopulateEstate ("")
    isNew = True
    SetUserGroup
  
    If isCR = False Then
        For i = 1 To SSTab1.Tabs - 1
            SSTab1.TabVisible(i) = False
        Next i
    Else
        For i = 1 To SSTab1.Tabs - 1
            SSTab1.TabVisible(i) = True
        Next i
    End If
        SSTab1.TabVisible(4) = True
    End Sub
Private Sub PopulateEstate(STR As String)
On Error Resume Next
    Dim tmpS As New BWGSearchManager
    Dim x%
'    Dim str As String
'    str = " sdname like '%_ME%' or sdname like '%_AK%'"
    Set colEstate_ = tmpS.EstateSearch(STR)
    CbEstate.Clear
    CbEstate.AddItem "== กรุณาระบุนิคมฯ =="
    For x = 1 To colEstate_.Count
        CbEstate.AddItem colEstate_(x).EstateName
    Next
    CbEstate.ListIndex = 0
    
'    Dim tmpRec As ADODB.Recordset
'    Dim x As Integer
'    Set tmpRec = New ADODB.Recordset
'    If Str = "" Then
'        tmpRec.Open "Select * from vw_Estate ", DBManager.GetDBConnection, adOpenDynamic, adLockOptimistic
'    Else
'        tmpRec.Open Str, DBManager.GetDBConnection, adOpenDynamic, adLockOptimistic
'    End If
'
'    CbEstate.Clear
'    If tmpRec.EOF = False Then
'    CbEstate.AddItem ("=== เลือกนิคม ===")
'    For x = 1 To tmpRec.RecordCount
'        CbEstate.AddItem tmpRec!EstateName
'        tmpRec.MoveNext
'    Next
'        CbEstate.ListIndex = 0
'    End If
'
'    tmpRec.Close
'    Set tmpRec = Nothing

End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If Label3.Caption = "รับเป็นลูกค้าแล้ว" Then
        If isCR <> True Then
            Exit Sub
        End If
End If
Dim selResult As Integer
    If myCustomer_ Is Nothing Then Exit Sub
    If myCustomer_.isChanged Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนปิดหน้าจอหรือไม่?", vbQuestion + vbYesNoCancel)
        If selResult = vbYes Then
            CmdSave_Click
            If myCustomer_.isChanged Then Cancel = 1
        ElseIf selResult = vbCancel Then
            Cancel = 1
        End If
    End If
End Sub





Private Sub lvBillAddr_DblClick()
    If lvBillAddr.SelectedItem Is Nothing Then Exit Sub
'    Load frCustomerBillAddr
'    With frCustomerBillAddr
'        .SetCallerForm Me, myCustomer_
'        .SetDataBillingAddr myCustomer_.BillAddresses(lvBillAddr.SelectedItem.Key)
'        .Show 1
'        myCustomer_.RequeryBillAddress
'        ShowBillAddress
'    End With
End Sub

Private Sub lvContract_DblClick()
    If lvContract.SelectedItem Is Nothing Then Exit Sub
    Load frmContract
    With frmContract
        .SetCallerForm Me, myCustomer_
        .SetDataContract myCustomer_.Contracts(lvContract.SelectedItem.Key)
        .Show 1
        myCustomer_.RequeryContracts
        ShowContracts
    End With
End Sub

Private Sub lvCustRem_DblClick()
    If lvCustRem.SelectedItem Is Nothing Then Exit Sub
'    Load frCustOtherRemark
'    With frCustOtherRemark
'        .SetCallerForm Me, myCustomer_
'        .SetDataBillingAddr myCustomer_.OtherRemarks(lvCustRem.SelectedItem.Key)
'        .Show 1
'        myCustomer_.RequeryOtherRemark
'        ShowOtherRemark
'    End With
End Sub

Private Sub lvSiteAddress_DblClick()
    If lvSiteAddress.SelectedItem Is Nothing Then Exit Sub
    Load frmCustomerSite
    With frmCustomerSite
        .SetCallerForm Me, myCustomer_
        .SetDataCustomerSite myCustomer_.SiteAddresses(lvSiteAddress.SelectedItem.Key)
        .Show 1
        myCustomer_.RequerySiteAddress
        ShowSiteAddress
    End With
End Sub

Private Sub lvStaff_DblClick()
    If lvStaff.SelectedItem Is Nothing Then Exit Sub
    Load frmCustomerStaff
    With frmCustomerStaff
        .SetCallerForm Me, myCustomer_
        .SetDataCustomerStaff myCustomer_.CustStaffs(lvStaff.SelectedItem.Key)
        .Show 1
        myCustomer_.RequeryCustStaff
        ShowCustStaff
    End With
End Sub

Private Sub lvWaste_DblClick()
    If lvWaste.SelectedItem Is Nothing Then Exit Sub
    Load frmWasteData
    With frmWasteData
        .SetCallerForm Me, myCustomer_
        .SetDataWasteData myCustomer_.WasteItems(lvWaste.SelectedItem.Key)
        .Show 1
        myCustomer_.RequeryWasteItems
        ShowWasteData
    End With
End Sub
Private Sub ListQuotation_DblClick()
  If ListQuotation.SelectedItem Is Nothing Then Exit Sub
  Dim tmpSite As BWGQuotationHeader
  Set tmpSite = myCustomer_.CustQuotations(ListQuotation.SelectedItem.Key)
  Dim R
  R = Shell(App.Path & "\UploadFile.exe  " & tmpSite.QuotationNo & "," & CurrentUser.ID, 1)
  'MsgBox tmpSite.QuotationNo, vbOKOnly + vbInformation, "Printer"
End Sub
Private Sub optBillType_Validate(Index As Integer, Cancel As Boolean)
    myCustomer_.BillingInfo.BillRecType = Index
    If myCustomer_.BillingInfo.isChanged Then myCustomer_.isChanged = True
End Sub

Private Sub optCheqType_Validate(Index As Integer, Cancel As Boolean)
    myCustomer_.BillingInfo.CheqRecType = Index
    If myCustomer_.BillingInfo.isChanged Then myCustomer_.isChanged = True
End Sub

Private Sub Text2_Validate(Cancel As Boolean)
    myCustomer_.Refcode = txtDepartment.Text
End Sub

Private Sub Txt14Digit_Validate(Cancel As Boolean)
    myCustomer_.Digit14 = txt14Digit.Text
End Sub

Private Sub txtAccContact_Validate(Cancel As Boolean)
    myCustomer_.BillingInfo.ContactName = Trim(txtAccContact.Text)
    If myCustomer_.BillingInfo.isChanged Then myCustomer_.isChanged = True
End Sub

Private Sub txtAccFax_Validate(Cancel As Boolean)
    myCustomer_.BillingInfo.FaxNo = Trim(txtAccFax.Text)
    If myCustomer_.BillingInfo.isChanged Then myCustomer_.isChanged = True
End Sub

Private Sub txtAccTel_Validate(Cancel As Boolean)
    myCustomer_.BillingInfo.TelNo = Trim(txtAccTel.Text)
    If myCustomer_.BillingInfo.isChanged Then myCustomer_.isChanged = True
End Sub

Private Sub txtAddr_Validate(Cancel As Boolean)
    myCustomer_.Address = Trim(txtAddr.Text)
End Sub

Private Sub txtBillDate_Validate(Index As Integer, Cancel As Boolean)
Dim tmpVal As Integer
    If IsNumeric(txtBillDate(Index).Text) Then
        tmpVal = Trim(txtBillDate(Index).Text)
    Else
        tmpVal = 0
    End If
    Select Case Index
        Case 1
            myCustomer_.BillingInfo.BillDay1 = tmpVal
        Case 2
            myCustomer_.BillingInfo.BillDay2 = tmpVal
        Case 3
            myCustomer_.BillingInfo.BillDay3 = tmpVal
        Case 4
            myCustomer_.BillingInfo.BillDay4 = tmpVal
        Case 5
            myCustomer_.BillingInfo.BillDay5 = tmpVal
    End Select
    If myCustomer_.BillingInfo.isChanged Then myCustomer_.isChanged = True
End Sub

Private Sub txtBillingDet_Validate(Cancel As Boolean)
    myCustomer_.BillingInfo.BillingDetail = Trim(txtBillingDet.Text)
    If myCustomer_.BillingInfo.isChanged Then myCustomer_.isChanged = True
End Sub

Private Sub txtCheqDate_Validate(Index As Integer, Cancel As Boolean)
Dim tmpVal As Integer
    If IsNumeric(txtCheqDate(Index).Text) Then
        tmpVal = Trim(txtCheqDate(Index).Text)
    Else
        tmpVal = 0
    End If
    Select Case Index
        Case 1
            myCustomer_.BillingInfo.CheqDay1 = tmpVal
        Case 2
            myCustomer_.BillingInfo.CheqDay2 = tmpVal
        Case 3
            myCustomer_.BillingInfo.CheqDay3 = tmpVal
        Case 4
            myCustomer_.BillingInfo.CheqDay4 = tmpVal
        Case 5
            myCustomer_.BillingInfo.CheqDay5 = tmpVal
    End Select
    If myCustomer_.BillingInfo.isChanged Then myCustomer_.isChanged = True
End Sub

Private Sub txtCheqDetail_Validate(Cancel As Boolean)
    myCustomer_.BillingInfo.CheqDetail = Trim(txtCheqDetail.Text)
    If myCustomer_.BillingInfo.isChanged Then myCustomer_.isChanged = True
End Sub

Private Sub txtCompanyNameTH_Validate(Cancel As Boolean)
    myCustomer_.CustomerName = Trim(txtCompanyNameTH.Text)
End Sub

Private Sub txtContactName_Validate(Cancel As Boolean)
    myCustomer_.ContactPerson = txtContactName.Text
End Sub

Private Sub txtCustRemark_Validate(Cancel As Boolean)
    myCustomer_.CustRemark = txtCustRemark.Text
End Sub

Private Sub txtDepartment_Validate(Cancel As Boolean)
    myCustomer_.DepartmentName = txtDepartment.Text
End Sub

Private Sub txtEmail_Validate(Cancel As Boolean)
    myCustomer_.EmailAddr = txtEmail.Text
End Sub

Private Sub txtFactoryID_Validate(Cancel As Boolean)
    myCustomer_.RegisterNo = txtFactoryID.Text
End Sub

Private Sub txtFaxNo_Validate(Cancel As Boolean)
    myCustomer_.FaxNo = txtFaxNo.Text
End Sub

Private Sub txtGenID_Validate(Cancel As Boolean)
    myCustomer_.WasteManageID = txtGenID.Text
End Sub

Private Sub txtIndustryType_Validate(Cancel As Boolean)
Dim CDet As String
CDet = Trim(curIndustryType_.TypeCode) & " - " & Trim(curIndustryType_.TypeDesc)
    If Trim(txtIndustryType.Text) = "" And CDet <> Trim(txtIndustryType.Text) Then
        Set curIndustryType_ = Nothing
        Set curIndustryType_ = New BWGIndustryType_New
        With myCustomer_
            .IndustryTypeID = ""
            .IndustryTypeName = ""
            .IndustryTypeIDNew = "0"
            .TypeDesc = ""
        End With
    Else
        If CDet <> Trim(txtIndustryType.Text) Then
            Set curIndustryType_ = Nothing
            Set curIndustryType_ = New BWGIndustryType_New
            With myCustomer_
            .IndustryTypeID = ""
            .IndustryTypeName = ""
            .IndustryTypeIDNew = "0"
            .TypeDesc = ""
            End With
            Load frmIndustryType
            With frmIndustryType
                .txtShortName.Text = Trim(Left(txtIndustryType.Text, 3))
                .SearchData
                .SetCallerForm Me
                .Show 1
            End With
            If myCustomer_.IndustryTypeIDNew = "" Then txtIndustryType.Text = ""
        End If
    End If
    Dim OldID As String
    OldID = myCustomer_.IndustryTypeIDNew
    If myCustomer_.IndustryTypeIDNew <> "" Then
        If myCustomer_.IndustryTypeIDNew <> "0" Then
            If myCustomer_.IndustryTypeIDNew = OldID Then Exit Sub
                    myCustomer_.UpdateIndustryType
        End If
    End If

End Sub

Private Sub txtOtherDoc_Validate(Cancel As Boolean)
    If Trim(txtOtherDoc.Text) = "" Then
        chkOtherDoc.Value = 0
        myCustomer_.BillingInfo.OtherDocDesc = ""
    Else
        chkOtherDoc.Value = 1
        myCustomer_.BillingInfo.OtherDocDesc = Trim(txtOtherDoc.Text)
    End If
End Sub

Private Sub txtParentCust_Validate(Cancel As Boolean)
    If Trim(txtParentCust.Text) = "" And Trim(curParentCust_.CustomerName) <> Trim(txtParentCust.Text) Then
        Set curParentCust_ = Nothing
        Set curParentCust_ = New BWGCustomer
        With myCustomer_
            .ParentCustCode = ""
            .ParentCustID = ""
            .ParentCustName = ""
        End With
    Else
        If Trim(curParentCust_.CustomerName) <> Trim(txtParentCust.Text) Then
            Set curParentCust_ = Nothing
            Set curParentCust_ = New BWGCustomer
            With myCustomer_
                .ParentCustCode = ""
                .ParentCustID = ""
                .ParentCustName = ""
            End With
            callLocation_ = "P" 'Parent
            Load frmCustomerSearch
            With frmCustomerSearch
                .getCriteria (txtParentCust.Text)
                .SetCallerForm Me
                .Show 1
            End With
            If myCustomer_.ParentCustID = "" Then txtParentCust.Text = ""
        End If
    End If
End Sub

Private Sub TxtRefCode_Validate(Cancel As Boolean)
    myCustomer_.Refcode = TxtRefCode.Text
End Sub

Private Sub txtRemDesc_Validate(Cancel As Boolean)
    myCustomer_.BillingInfo.RemDesc = Trim(txtRemDesc.Text)
    If myCustomer_.BillingInfo.isChanged Then myCustomer_.isChanged = True
End Sub

Private Sub txtSaleName_Validate(Cancel As Boolean)
'    If curSaleStaff_ Is Nothing Then Set curSaleStaff_ = New BWGCompanyStaff
'    If Trim(txtSaleName.Text) = "" And Trim(curSaleStaff_.StaffFName & " " & curSaleStaff_.StaffLName) <> Trim(txtSaleName.Text) Then
'        Set curSaleStaff_ = Nothing
'        Set curSaleStaff_ = New BWGCompanyStaff
'        With myCustomer_
'            .UnderSaleName = ""
'            .UnderSaleStaffID = ""
'        End With
'    Else
'        If Trim(curSaleStaff_.StaffFName & " " & curSaleStaff_.StaffLName) <> Trim(txtSaleName.Text) Then
'            Set curSaleStaff_ = Nothing
'            Set curSaleStaff_ = New BWGCompanyStaff
'            With myCustomer_
'                .UnderSaleName = ""
'                .UnderSaleStaffID = ""
'            End With
'            Load frStaffSearch
'            With frStaffSearch
'                .getCriteria (txtSaleName.Text)
'                .setCallerForm Me
'                .Show 1
'            End With
'            If myCustomer_.UnderSaleStaffID = "" Then txtSaleName.Text = ""
'        End If
'    End If
End Sub

Private Sub txtShortName_Validate(Cancel As Boolean)
    myCustomer_.CustomerShortName = txtShortName.Text
End Sub

Private Sub txtTelNo_Validate(Cancel As Boolean)
    myCustomer_.TelNo = txtTelNo.Text
End Sub

Private Sub txtTumbol_Validate(Cancel As Boolean)
    If Trim(txtTumbol.Text) = "" And Trim(curTumbol_.Name) <> Trim(txtTumbol.Text) Then
        Set curTumbol_ = Nothing
        Set curTumbol_ = New IMITumbol
        txtAumphur.Text = ""
        txtProvince.Text = ""
        With myCustomer_
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
            With myCustomer_
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
            If myCustomer_.TumbolID = "" Then txtTumbol.Text = ""
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
            myCustomer_.AumphurID = .AumphurID
            myCustomer_.AumphurName = .AumphurName
            myCustomer_.TumbolID = .ID
            myCustomer_.TumbolName = .Name
            myCustomer_.ProvinceID = .ProvinceID
            myCustomer_.ProvinceName = .ProvinceName
            PopulateEstate ("Select * from vw_Estate where ProvinceID in ('" & .ProvinceID & "','0') ")
        End If
        
    End With
End Sub

Private Sub CreateNewCustomer()
    Set myCustomer_ = Nothing
    Set myCustomer_ = New BWGCustomer
    With myCustomer_
        .ID = GetGUID
        If cboCustType.ListIndex = 1 Then
            .CustomerType = "C"
        ElseIf cboCustType.ListIndex = 2 Then
            .CustomerType = "S"
        ElseIf cboCustType.ListIndex = 3 Then
            .CustomerType = "M"
        ElseIf cboCustType.ListIndex = 4 Then
            .CustomerType = "MS"
        ElseIf cboCustType.ListIndex = 5 Then
            .CustomerType = "B"
        ElseIf cboCustType.ListIndex = 6 Then
            .CustomerType = "E"
        ElseIf cboCustType.ListIndex = 7 Then
            .CustomerType = "MI"
        ElseIf cboCustType.ListIndex = 8 Then
            .CustomerType = "I"
        End If
        .CreatedStaffID = CurrentUser.ID
        .CreatedStaffName = CurrentUser.StaffFName & " " & CurrentUser.StaffLName
        .CreatedDate = TodayDate
        .isChanged = False
    End With
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{Tab}"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0
End Sub
Private Sub PopulateCR()
Dim tmpS As New BWGSearchManager
Dim x%
    Set colCR_ = tmpS.UserSearch("PositionName = 'CR'")
    cbCRStaff.Clear
    cbCRStaff.AddItem "== เลือกเจ้าหน้าที่ CR =="
    For x = 1 To colCR_.Count
        cbCRStaff.AddItem colCR_(x).StaffFName & " " & colCR_(x).StaffLName
    Next
    cbCRStaff.ListIndex = 0
End Sub

Private Sub ClearScreen()
    txtCustCode.Text = "== AUTO =="
    Label3.Visible = False
    txtCompanyNameTH.Text = ""
    txtAddr.Text = ""
    TxtRefCode.Text = ""
    txtTumbol.Text = ""
    cboCustGroup.ListIndex = -1
    txtAumphur.Text = ""
    txtCustRemark.Text = ""
    txtProvince.Text = ""
    txtShortName.Text = ""
    TxtWasteSK2Exp.Text = ""
    cboDistrict.ListIndex = 0
    txtIndustryType.Text = ""
    txtFactoryID.Text = ""
    txtContactName.Text = ""
    txtDepartment.Text = ""
    txtTelNo.Text = ""
    txtFaxNo.Text = ""
    txtGenID.Text = ""
    txtCheqDetail.Text = ""
    chkOtherDoc.Value = 0
    lvWaste.ListItems.Clear
    lvCustRem.ListItems.Clear
    ListQuotation.ListItems.Clear
    Set curTumbol_ = Nothing
    Set curTumbol_ = New IMITumbol
    cboCustType.ListIndex = 0
    txtParentCust.Text = ""
    Set curParentCust_ = Nothing
    Set curParentCust_ = New BWGCustomer
    lvSiteAddress.ListItems.Clear
    txtAccContact.Text = ""
    txtAccTel.Text = ""
    txtAccFax.Text = ""
    txtBillingDet.Text = ""
    chkPO.Value = 0
    chkInv.Value = 0
    chkBilling.Value = 0
    chkContract.Value = 0
    chkQuotation.Value = 0
    txtRemDesc.Text = ""
    txtZipCode.Text = ""
    txtSaleName.Text = ""
    txtWebSite.Text = ""
    txtEmail.Text = ""
    txtOtherDoc.Text = ""
    lvContract.ListItems.Clear
    Me.lvBillAddr.ListItems.Clear
    lvStaff.ListItems.Clear
    optBillType(1).Value = False
    optBillType(2).Value = False
    optBillType(3).Value = False
    optBillType(4).Value = False
    txtBillDate(1).Text = ""
    txtBillDate(2).Text = ""
    txtBillDate(3).Text = ""
    txtBillDate(4).Text = ""
    txtBillDate(5).Text = ""
    chkBillDay(1).Value = 0
    chkBillDay(2).Value = 0
    chkBillDay(3).Value = 0
    chkBillDay(4).Value = 0
    chkBillDay(5).Value = 0
    cboBillWeekCount.ListIndex = 0
    optCheqType(1).Value = False
    optCheqType(2).Value = False
    optCheqType(3).Value = False
    optCheqType(4).Value = False
    txtCheqDate(1).Text = ""
    txtCheqDate(2).Text = ""
    txtCheqDate(3).Text = ""
    txtCheqDate(4).Text = ""
    txtCheqDate(5).Text = ""
    chkCheqDay(1).Value = 0
    chkCheqDay(2).Value = 0
    chkCheqDay(3).Value = 0
    chkCheqDay(4).Value = 0
    chkCheqDay(5).Value = 0
    cboCheqWeekCount.ListIndex = 0
    
    Set curIndustryType_ = Nothing
    Set curIndustryType_ = New BWGIndustryType_New
    Set curSaleStaff_ = Nothing
    Set curSaleStaff_ = New BWGCompanyStaff
End Sub

Public Sub SetDataIndustryType(selIndustry As BWGIndustryType_New, Optional isAssigned As Boolean = True)
    Set curIndustryType_ = selIndustry
    With curIndustryType_
        txtIndustryType.Text = .TypeCode & " - " & .TypeDesc
        If isAssigned Then
            myCustomer_.IndustryTypeIDNew = .TypeID
            myCustomer_.TypeDesc = .TypeDesc
        End If
    End With
End Sub

Private Sub PopulateDistrict()
On Error Resume Next
    Dim tmpS As New BWGSearchManager
    Dim x%
'    Dim str As String
'    str = " sdname like '%_ME%' or sdname like '%_AK%'"
    Set colDistrict_ = tmpS.SaleDistrictSearch
    cboDistrict.Clear
    cboDistrict.AddItem "== กรุณาระบุเขต Sale =="
    For x = 1 To colDistrict_.Count
        cboDistrict.AddItem colDistrict_(x).DistrictName
    Next
    cboDistrict.ListIndex = 0
    
    Dim tmpS2 As New BWGSearchManager
    Dim i%
    Set colDistrict2_ = tmpS2.SaleDistrictSearch
    cboDistrict2.Clear
    cboDistrict2.AddItem "ระบุผู้แทน "
       For i = 1 To colDistrict2_.Count
        cboDistrict2.AddItem colDistrict2_(i).DistrictName
    Next
    cboDistrict2.ListIndex = 0
End Sub
Public Sub SetDataCustomer(selCust As BWGCustomer, Optional isAssigned As Boolean = True)
On Error Resume Next
Dim tmpS As New BWGSearchManager
Dim tmpSt As New BWGCompanyStaff
Dim tmpCol As Collection
    cmdApprove.Visible = False
    Capt = "O"
    If callLocation_ = "N" Then
        Set myCustomer_ = selCust
        With myCustomer_
            If .IsVIPCR = "Y" Then
                LblVIPCR.Visible = True
            Else
                LblVIPCR.Visible = False
            End If
            isNew = .isNew
            If .Status = "N" Then
                    Label3.Caption = "รับเป็นลูกค้าแล้ว"
                    Label3.Visible = True
                    'Text1 = Format(.ApproveDate, "dd/mm/yyyy")
                    'Text1.Visible = True
                    If isCR = True Then
                        cmdApprove.Visible = False
                        cmdReject.Visible = False
                    End If
            ElseIf .Status = "J" Then
                    Label3.Caption = "ตีกลับ"
                    Label3.Visible = True
                    Text1 = .RejectDesc2
                    Text1.Visible = True
                    If isCR = True Then
                        cmdApprove.Visible = False
                        cmdReject.Visible = False
                    End If
            ElseIf .Status = "R" Then
                    Label3.Caption = "รอการตรวจรับ"
                    Label3.Visible = True
                    If isCR = True Then
                        cmdApprove.Visible = True
                        cmdReject.Visible = True
                    End If
            ElseIf .Status = "P" And .isApprove = True Then
                    Label3.Caption = "รับเป็นลูกค้าแล้ว"
                    Label3.Visible = True
                    'Text1 = Format(.ApproveDate, "dd/mm/yyyy")
                    'Text1.Visible = True
                    If isCR = True Then
                        cmdApprove.Visible = False
                        cmdReject.Visible = False
                    End If
            ElseIf .Status = "D" Then
                    Label3.Caption = "ข้อมูลยังไม่ได้ส่ง"
                    Label3.Visible = True
                    Text1 = "ข้อมูลยังไม่ได้ส่งให้ CR ตรวจ"
                    Text1.Visible = True
                    cmdApprove.Visible = False
                    cmdReject.Visible = False
            End If
            If .CompanyPolicy <> "" Then
                CbCp.Text = .CompanyPolicy
            End If
            txtCustCode.Text = .CustomerCode
            txtCompanyNameTH.Text = .CustomerName
            txtAddr.Text = .Address
            txtZipCode.Text = .ZipCode
            txtEmail.Text = .EmailAddr
            txtWebSite.Text = .Website
            txtCustRemark.Text = .CustRemark
            txtShortName.Text = .CustomerShortName
            txtSaleName.Text = .UnderSaleName
            txtSalepst.Text = .Sale2Name
            TxtRefCode.Text = .Refcode
            TxtWasteSK2Exp.Text = .WasteSK2Exp
            If .CustdeptName <> "0" Then
                cbDept.Text = .CustdeptName
            Else
                cbDept.ListIndex = 0
            End If
            SetTextToCombo .CustGroupName, cboCustGroup
            Set tmpCol = tmpS.TumbolSearch("TumbolID='" & .TumbolID & "'")
            If tmpCol.Count > 0 Then
                SetDataTumbol tmpCol(1), False
                'txtTumbol.Text = .TumbolName
                'txtAumphur.Text = .AumphurName
                'txtProvince.Text = .ProvinceName
            End If
            SetTextToCombo .SaleDistrictName, cboDistrict
            SetTextToCombo .SaleDistrictName2, cboDistrict2
            If .CRStaffID <> "" Then
                SetTextToCombo .CRStaffName, cbCRStaff
            Else
                cbCRStaff.ListIndex = 0
            End If
            If .IndustryTypeIDNew <> "" Then
            Set tmpCol = tmpS.IndustryTypeSearch("TypeID='" & .IndustryTypeIDNew & "'")
                If tmpCol.Count > 0 Then
                    SetDataIndustryType tmpCol(1), False
                    txtIndustryType.Text = .TypeCode & " - " & .TypeDesc
                Else
                    txtIndustryType.Text = ""
                End If
            Else
                txtIndustryType.Text = ""
            End If
            
            If .EstateID <> "0" Then
                    Dim x%
                    Set colEstate_ = tmpS.EstateSearch("Select * from vw_Estate where EstateID='" & .EstateID & "' or EstateID='112' ")
                    CbEstate.Clear
                    For x = 1 To colEstate_.Count
                        CbEstate.AddItem colEstate_(x).EstateName
                    Next
                    Set colEstate_ = tmpS.EstateSearch("Select * from vw_Estate where ProvinceID='" & .ProvinceID & "' and EstateID<>'" & .EstateID & "' ")
                    For x = 1 To colEstate_.Count
                        CbEstate.AddItem colEstate_(x).EstateName
                    Next
                    CbEstate.ListIndex = 0
            Else
                PopulateEstate ("Select * from vw_Estate where ProvinceID='" & .ProvinceID & "' or EstateID='112' ")
                CbEstate.ListIndex = 0
            End If
            
            Set tmpCol = Nothing
            Set tmpCol = tmpS.UserSearch("StaffID='" & .UnderSaleStaffID & "'")
            If tmpCol.Count > 0 Then
                SetDataStaff tmpCol(1), False
            End If
            Set tmpCol = Nothing
            txtFactoryID.Text = .RegisterNo
            txt14Digit.Text = .Digit14
            txtContactName.Text = .ContactPerson
            txtDepartment.Text = .DepartmentName
            txtTelNo.Text = .TelNo
            txtFaxNo.Text = .FaxNo
            If UCase(.WasteManageID) <> "NULL" Or Trim(.WasteManageID) <> "" Then
                txtGenID.Text = .WasteManageID
                Else
                txtGenID.Text = ""
            End If
            If .CustomerType = "C" Then
                cboCustType.ListIndex = 1
            ElseIf .CustomerType = "S" Then
                cboCustType.ListIndex = 2
            ElseIf .CustomerType = "M" Then
                cboCustType.ListIndex = 3
            ElseIf .CustomerType = "MS" Then
                cboCustType.ListIndex = 4
            ElseIf .CustomerType = "B" Then
                cboCustType.ListIndex = 5
            ElseIf .CustomerType = "E" Then
                cboCustType.ListIndex = 6
            ElseIf .CustomerType = "MI" Then
                cboCustType.ListIndex = 7
             ElseIf .CustomerType = "I" Then
                cboCustType.ListIndex = 8
            Else
                cboCustType.ListIndex = 0
            End If
            Set tmpCol = tmpS.CustomerSearch("CustomerID='" & .ParentCustID & "'")
            If tmpCol.Count > 0 Then
                Set curParentCust_ = tmpCol(1)
                txtParentCust.Text = curParentCust_.CustomerName
            Else
                txtParentCust.Text = ""
            End If
        End With
        With myCustomer_.BillingInfo
            txtAccContact.Text = .ContactName
            txtAccTel.Text = .TelNo
            txtAccFax.Text = .FaxNo
            txtBillingDet.Text = .BillingDetail
            txtCheqDetail.Text = .CheqDetail
            txtOtherDoc.Text = .OtherDocDesc
            If Trim(txtOtherDoc.Text) <> "" Then
                chkOtherDoc.Value = 1
            Else
                chkOtherDoc.Value = 0
            End If
            chkPO.Value = Abs(CInt(.NeedPO))
            chkInv.Value = Abs(CInt(.NeedInvoice))
            chkBilling.Value = Abs(CInt(.NeedBillingNote))
            chkContract.Value = Abs(CInt(.NeedContract))
            chkQuotation.Value = Abs(CInt(.NeedQuotation))
            txtRemDesc.Text = .RemDesc
            If IsNumeric(.BillRecType) Then
                If CInt(.BillRecType) >= 1 And CInt(.BillRecType) <= 4 Then
                    optBillType(CInt(.BillRecType)).Value = True
                    If CInt(.BillRecType) = 2 Then
                        If .BillDay1 > 0 Then txtBillDate(1).Text = .BillDay1
                        If .BillDay2 > 0 Then txtBillDate(2).Text = .BillDay2
                        If .BillDay3 > 0 Then txtBillDate(3).Text = .BillDay3
                        If .BillDay4 > 0 Then txtBillDate(4).Text = .BillDay4
                        If .BillDay5 > 0 Then txtBillDate(5).Text = .BillDay5
                    ElseIf CInt(.BillRecType) = 3 Then
                        If .BillDay1 > 0 Then chkBillDay(1).Value = 1
                        If .BillDay2 > 0 Then chkBillDay(2).Value = 1
                        If .BillDay3 > 0 Then chkBillDay(3).Value = 1
                        If .BillDay4 > 0 Then chkBillDay(4).Value = 1
                        If .BillDay5 > 0 Then chkBillDay(5).Value = 1
                    End If
                End If
            End If
            cboBillWeekCount.ListIndex = .BillWeekCount
            If IsNumeric(.CheqRecType) Then
                If CInt(.CheqRecType) >= 1 And CInt(.CheqRecType) <= 4 Then
                    optCheqType(CInt(.CheqRecType)).Value = True
                    If CInt(.CheqRecType) = 2 Then
                        If .CheqDay1 > 0 Then txtCheqDate(1).Text = .CheqDay1
                        If .CheqDay2 > 0 Then txtCheqDate(2).Text = .CheqDay2
                        If .CheqDay3 > 0 Then txtCheqDate(3).Text = .CheqDay3
                        If .CheqDay4 > 0 Then txtCheqDate(4).Text = .CheqDay4
                        If .CheqDay5 > 0 Then txtCheqDate(5).Text = .CheqDay5
                    ElseIf CInt(.CheqRecType) = 3 Then
                        If .CheqDay1 > 0 Then chkCheqDay(1).Value = 1
                        If .CheqDay2 > 0 Then chkCheqDay(2).Value = 1
                        If .CheqDay3 > 0 Then chkCheqDay(3).Value = 1
                        If .CheqDay4 > 0 Then chkCheqDay(4).Value = 1
                        If .CheqDay5 > 0 Then chkCheqDay(5).Value = 1
                    End If
                End If
            End If
            cboCheqWeekCount.ListIndex = .CheqWeekCount
        End With
        ShowSiteAddress
        ShowBillAddress
        ShowWasteData
        ShowContracts
        ShowCustStaff
        ShowOtherRemark
        ShowCustQuotation
    ElseIf callLocation_ = "P" Then
        Set curParentCust_ = selCust
        With curParentCust_
            txtParentCust.Text = .CustomerName
            If isAssigned Then
                myCustomer_.ParentCustCode = .CustomerCode
                myCustomer_.ParentCustID = .ID
                myCustomer_.ParentCustName = .CustomerName
            End If
        End With
    End If
    If Label3.Caption = "รับเป็นลูกค้าแล้ว" Then
        If isCR = True Then
            cmdSave.Enabled = True
        Else
            cmdSave.Enabled = False
        End If
    End If
    End Sub

Public Property Get CallLocation() As String
    CallLocation = callLocation_
End Property
Public Property Let CallLocation(ByVal newValue As String)
    callLocation_ = Trim(newValue)
End Property

Private Sub ShowWasteData()
On Error Resume Next
Dim x%
Dim tmpWaste As BWGWasteDataCR
Dim tmpList As ListItem
    lvWaste.ListItems.Clear
    For x = 1 To myCustomer_.WasteItems.Count
        Set tmpWaste = myCustomer_.WasteItems(x)
        Set tmpList = lvWaste.ListItems.Add(, "'" & tmpWaste.ID & "'", tmpWaste.WasteNo)
        tmpList.SubItems(1) = tmpWaste.WasteName
        tmpList.SubItems(2) = tmpWaste.EUCode1 & "-" & tmpWaste.EUCode2 & "-" & tmpWaste.EUCode3
        tmpList.SubItems(3) = tmpWaste.WasteType
        tmpList.SubItems(5) = tmpWaste.LicenseExpDate
        Set tmpList = Nothing
        Set tmpWaste = Nothing
    Next
End Sub
Private Sub ShowSiteAddress()
Dim x%
Dim tmpSite As BWGCustomerSite
Dim tmpList As ListItem
    lvSiteAddress.ListItems.Clear
    For x = 1 To myCustomer_.SiteAddresses.Count
        Set tmpSite = myCustomer_.SiteAddresses(x)
        Set tmpList = lvSiteAddress.ListItems.Add(, "'" & tmpSite.ID & "'", tmpSite.SiteName)
        tmpList.SubItems(1) = tmpSite.Address & " " & tmpSite.TumbolName & " " & tmpSite.AumphurName & " " & tmpSite.ProvinceName & " " & tmpSite.ZipCode
        Set tmpList = Nothing
        Set tmpSite = Nothing
    Next
End Sub
Private Sub ShowBillAddress()
Dim x%
Dim tmpSite As BWGCustomerBillAddr
Dim tmpList As ListItem
    lvBillAddr.ListItems.Clear
    For x = 1 To myCustomer_.BillAddresses.Count
        Set tmpSite = myCustomer_.BillAddresses(x)
        Set tmpList = lvBillAddr.ListItems.Add(, "'" & tmpSite.ID & "'", tmpSite.BillName)
        tmpList.SubItems(1) = tmpSite.Address & " " & tmpSite.TumbolName & " " & tmpSite.AumphurName & " " & tmpSite.ProvinceName & " " & tmpSite.ZipCode
        Set tmpList = Nothing
        Set tmpSite = Nothing
    Next
End Sub

Private Sub ShowContracts()
Dim x%
Dim tmpSite As BWGContract
Dim tmpList As ListItem
    lvContract.ListItems.Clear
    For x = 1 To myCustomer_.Contracts.Count
        Set tmpSite = myCustomer_.Contracts(x)
        Set tmpList = lvContract.ListItems.Add(, "'" & tmpSite.ID & "'", FormatYMD_to_DMY(tmpSite.StartDate, "/", "/"))
        tmpList.SubItems(1) = FormatYMD_to_DMY(tmpSite.EndDate, "/", "/")
        tmpList.SubItems(2) = tmpSite.docNo
        tmpList.SubItems(3) = tmpSite.RemDesc
        Set tmpList = Nothing
        Set tmpSite = Nothing
    Next
End Sub

Public Sub ShowCustStaff()
Dim x%
Dim tmpSite As BWGCustomerStaff
Dim tmpList As ListItem
    lvStaff.ListItems.Clear
    For x = 1 To myCustomer_.CustStaffs.Count
        Set tmpSite = myCustomer_.CustStaffs(x)
        Set tmpList = lvStaff.ListItems.Add(, "'" & tmpSite.ID & "'", tmpSite.FirstName & " " & tmpSite.LastName)
        tmpList.SubItems(1) = tmpSite.PositionName
        tmpList.SubItems(2) = tmpSite.DeptName
        tmpList.SubItems(3) = tmpSite.TelNo
        tmpList.SubItems(4) = tmpSite.Email
        tmpList.SubItems(5) = tmpSite.Subjact
        tmpList.SubItems(6) = tmpSite.Address
        Set tmpList = Nothing
        Set tmpSite = Nothing
    Next
End Sub
Public Sub ShowCustQuotation()
Dim x%
Dim tmpSite As BWGQuotationHeader
Dim tmpList As ListItem
    ListQuotation.ListItems.Clear
    For x = 1 To myCustomer_.CustQuotations.Count
        Set tmpSite = myCustomer_.CustQuotations(x)
        Set tmpList = ListQuotation.ListItems.Add(, "'" & tmpSite.ID & "'", tmpSite.QuotationNo)
        tmpList.SubItems(1) = tmpSite.IsApproved
        tmpList.SubItems(2) = tmpSite.ApprovedDate
        tmpList.SubItems(3) = tmpSite.AttentionTo
        Set tmpList = Nothing
        Set tmpSite = Nothing
    Next
End Sub

Private Sub TxtWasteSK2Exp_Validate(Cancel As Boolean)
    Dim UserGrp As New AWSUserGroup
    Dim tSql As String, rs As New ADODB.Recordset
    Dim i As Integer
        If Not CurrentUser Is Nothing Then
            For i = 1 To CurrentUser.getUserGroups.Count
                Set UserGrp = CurrentUser.getUserGroups(i)
                If Trim(UserGrp.Name) = "Administrator" Or Trim(UserGrp.Name) = "IT" Or Trim(UserGrp.Name) = "Environment" Then
                TxtWasteSK2Exp.Text = FormatDateTime(TxtWasteSK2Exp.Text)
                    myCustomer_.WasteSK2Exp = Trim(TxtWasteSK2Exp.Text)
                    myCustomer_.UpdateSK2Exp
                       Exit For
                Else
                    MsgBox "ท่านไม่มีสิทธิ์ในการแก้ไขข้อมูลนี้ กรุณาติดต่อแผนก Environment เพื่อทำการแก้ไขข้อมูล"
                    TxtWasteSK2Exp.Text = ""
                    Exit Sub
                End If
            Next i
        End If

End Sub

Private Sub txtWebSite_Validate(Cancel As Boolean)
    myCustomer_.Website = txtWebSite.Text
End Sub

Private Sub txtZipCode_Validate(Cancel As Boolean)
    myCustomer_.ZipCode = Trim(txtZipCode.Text)
End Sub

Public Sub SetDataStaff(selStaff As BWGCompanyStaff, Optional isAssigned As Boolean = True)
    Set curSaleStaff_ = selStaff
    With curSaleStaff_
        txtSaleName.Text = .StaffFName & " " & .StaffLName
            If isAssigned Then
            myCustomer_.UnderSaleName = .StaffFName & " " & .StaffLName
            myCustomer_.UnderSaleStaffID = .ID
        End If
    End With
End Sub
Public Sub SetCRstaff(CRstaff As BWGCompanyStaff, Optional isAssigned As Boolean = True)
Set curCRstaff_ = CRstaff
With curCRstaff_
myCustomer_.CreatedStaffName = .StaffFName & "  " & .StaffFName
myCustomer_.CreatedStaffID = .ID
End With
End Sub

Private Sub PopulateCustGroup()
Dim tmpS As New BWGSearchManager
Dim x%
    Set colCustGroup_ = tmpS.CustomerGroupSearch
    cboCustGroup.Clear
    For x = 1 To colCustGroup_.Count
        cboCustGroup.AddItem colCustGroup_(x).CustGroupName
    Next
        cboCustGroup.ListIndex = 2
End Sub

Private Sub ShowOtherRemark()
Dim x%
Dim tmpSite As BWGCustOrderRemark
Dim tmpList As ListItem
    lvCustRem.ListItems.Clear
    For x = 1 To myCustomer_.OtherRemarks.Count
        Set tmpSite = myCustomer_.OtherRemarks(x)
        Set tmpList = lvCustRem.ListItems.Add(, "'" & tmpSite.ID & "'", tmpSite.RemDesc)
        tmpList.SubItems(1) = tmpSite.OrdPos
        Set tmpList = Nothing
        Set tmpSite = Nothing
    Next
End Sub
