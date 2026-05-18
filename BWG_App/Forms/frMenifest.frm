VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMenifest 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "F056"
   ClientHeight    =   9315
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11580
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9315
   ScaleWidth      =   11580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtRefME 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   5910
      TabIndex        =   222
      TabStop         =   0   'False
      Top             =   720
      Width           =   1845
   End
   Begin VB.TextBox txtCancelDetail 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "Angsana New"
         Size            =   15.75
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   4080
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   200
      Top             =   6300
      Visible         =   0   'False
      Width           =   4755
   End
   Begin VB.TextBox txtAPNo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   9960
      Locked          =   -1  'True
      TabIndex        =   197
      TabStop         =   0   'False
      Text            =   "== AUTO =="
      Top             =   360
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtMenifestNo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   7800
      Locked          =   -1  'True
      TabIndex        =   96
      TabStop         =   0   'False
      Text            =   "== AUTO =="
      Top             =   360
      Width           =   1575
   End
   Begin VB.ComboBox cboCompany 
      Enabled         =   0   'False
      Height          =   315
      Left            =   5220
      Style           =   2  'Dropdown List
      TabIndex        =   94
      TabStop         =   0   'False
      Top             =   0
      Width           =   6315
   End
   Begin VB.CommandButton cmdMenifestClosed 
      Caption         =   "ปิดงาน"
      Height          =   795
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   88
      Top             =   8460
      Width           =   1005
   End
   Begin VB.TextBox txtDocNo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   5220
      Locked          =   -1  'True
      TabIndex        =   95
      TabStop         =   0   'False
      Text            =   "== AUTO =="
      Top             =   360
      Width           =   1575
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "เก็บข้อมูล"
      Height          =   795
      Left            =   6300
      Style           =   1  'Graphical
      TabIndex        =   89
      Top             =   8460
      Width           =   1005
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "เปิดข้อมูลเก่า"
      Height          =   795
      Left            =   7350
      Style           =   1  'Graphical
      TabIndex        =   90
      Top             =   8460
      Width           =   1005
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "สร้างใหม่"
      Height          =   795
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   91
      Top             =   8460
      Width           =   1005
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "ปิด"
      Height          =   795
      Left            =   10500
      Style           =   1  'Graphical
      TabIndex        =   93
      Top             =   8460
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ยกเลิก"
      Height          =   795
      Left            =   9450
      Style           =   1  'Graphical
      TabIndex        =   92
      Top             =   8460
      Width           =   1005
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7365
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1080
      Width           =   11565
      _ExtentX        =   20399
      _ExtentY        =   12991
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "ส่วนของผู้ก่อกำเนิดของเสียอันตราย"
      TabPicture(0)   =   "frMenifest.frx":0000
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
      Tab(0).Control(21)=   "Label2(74)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Label2(89)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Label4"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Label7"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "dtSign1"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "dtWorkDate"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "grItem"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "txtCustomer"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "txtCustTel"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "txtCustFax"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "txtSolidQty"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "txtLiquidQty"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "txtTreatment"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "txtSigneName1"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "txtCustER"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "Frame4"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "txtCollector"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "txtCollectorID"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "cboLiquidUnit"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "cboSolidUnit"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "cboSiteName"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "cboWasteType"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "cboGenerator"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "txtCollectorID2"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "txtCollector2"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "txtJobNo"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "cboWType"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "txtPlace"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).Control(49)=   "Command4"
      Tab(0).Control(49).Enabled=   0   'False
      Tab(0).Control(50)=   "cboRem"
      Tab(0).Control(50).Enabled=   0   'False
      Tab(0).Control(51)=   "txtRem"
      Tab(0).Control(51).Enabled=   0   'False
      Tab(0).Control(52)=   "txtLengthNum"
      Tab(0).Control(52).Enabled=   0   'False
      Tab(0).Control(53)=   "Text1"
      Tab(0).Control(53).Enabled=   0   'False
      Tab(0).ControlCount=   54
      TabCaption(1)   =   "ส่วนของผู้ขนส่ง และผู้เก็บรวบรวมของเสียอันตราย"
      TabPicture(1)   =   "frMenifest.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "TxtRate"
      Tab(1).Control(1)=   "Frame5"
      Tab(1).Control(2)=   "Check1"
      Tab(1).Control(3)=   "txtDriverName"
      Tab(1).Control(4)=   "cboCarRegisID"
      Tab(1).Control(5)=   "cboActualUnit"
      Tab(1).Control(6)=   "txtTruckType"
      Tab(1).Control(7)=   "txtAumphur"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "txtTumbol"
      Tab(1).Control(9)=   "txtDestroyer"
      Tab(1).Control(10)=   "txtDestroyID"
      Tab(1).Control(11)=   "txtDestroyPlace"
      Tab(1).Control(12)=   "txtRealWeight"
      Tab(1).Control(13)=   "txtDestER"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "txtDestFax"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "txtDestTel"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "txtTreatTime"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "cboTimeUnit"
      Tab(1).Control(18)=   "txtDestName"
      Tab(1).Control(19)=   "Frame1"
      Tab(1).Control(20)=   "cboTransUnit"
      Tab(1).Control(21)=   "txtSign1B"
      Tab(1).Control(22)=   "txtTransTime"
      Tab(1).Control(23)=   "txtToProvince"
      Tab(1).Control(24)=   "txtFromProvince"
      Tab(1).Control(25)=   "txtTransTel"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).Control(26)=   "txtTransFax"
      Tab(1).Control(26).Enabled=   0   'False
      Tab(1).Control(27)=   "txtTransER"
      Tab(1).Control(27).Enabled=   0   'False
      Tab(1).Control(28)=   "txtTransID1B"
      Tab(1).Control(29)=   "txtTrans1B"
      Tab(1).Control(30)=   "dtSign2"
      Tab(1).Control(31)=   "dtSign3"
      Tab(1).Control(32)=   "Label5"
      Tab(1).Control(33)=   "Label2(90)"
      Tab(1).Control(34)=   "Label2(12)"
      Tab(1).Control(35)=   "Label2(9)"
      Tab(1).Control(36)=   "Label2(15)"
      Tab(1).Control(37)=   "Label2(16)"
      Tab(1).Control(38)=   "Label2(17)"
      Tab(1).Control(39)=   "Label2(18)"
      Tab(1).Control(40)=   "Label2(19)"
      Tab(1).Control(41)=   "Label2(44)"
      Tab(1).Control(42)=   "Label2(45)"
      Tab(1).Control(43)=   "Label2(46)"
      Tab(1).Control(44)=   "Label2(47)"
      Tab(1).Control(45)=   "Label2(48)"
      Tab(1).Control(46)=   "Label2(49)"
      Tab(1).Control(47)=   "Label1(3)"
      Tab(1).Control(48)=   "Label2(7)"
      Tab(1).Control(49)=   "Label1(0)"
      Tab(1).Control(50)=   "Label2(43)"
      Tab(1).Control(51)=   "Label2(41)"
      Tab(1).Control(52)=   "Label2(40)"
      Tab(1).Control(53)=   "Label2(39)"
      Tab(1).Control(54)=   "Label2(38)"
      Tab(1).Control(55)=   "Image1(2)"
      Tab(1).Control(56)=   "Image1(0)"
      Tab(1).Control(57)=   "Label2(28)"
      Tab(1).Control(58)=   "Label2(31)"
      Tab(1).Control(59)=   "Label2(30)"
      Tab(1).Control(60)=   "Label2(29)"
      Tab(1).Control(61)=   "Label2(27)"
      Tab(1).Control(62)=   "Label2(26)"
      Tab(1).ControlCount=   63
      TabCaption(2)   =   "ข้อมูลเพิ่มเติมส่วนงานบัญชี"
      TabPicture(2)   =   "frMenifest.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame3"
      Tab(2).Control(1)=   "txtAccStaff"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Frame2(2)"
      Tab(2).Control(3)=   "Frame2(1)"
      Tab(2).Control(4)=   "Frame2(0)"
      Tab(2).Control(5)=   "Label2(59)"
      Tab(2).ControlCount=   6
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1800
         TabIndex        =   228
         Text            =   "Text1"
         Top             =   900
         Width           =   1815
      End
      Begin VB.TextBox TxtRate 
         Height          =   330
         Left            =   -67755
         TabIndex        =   221
         Top             =   4725
         Width           =   870
      End
      Begin VB.Frame Frame5 
         Caption         =   "การส่ง"
         Enabled         =   0   'False
         Height          =   735
         Left            =   -65820
         TabIndex        =   216
         Top             =   4410
         Width           =   2310
         Begin VB.OptionButton Option3 
            Caption         =   "E-Mail"
            Height          =   195
            Left            =   45
            TabIndex        =   219
            Top             =   495
            Width           =   1140
         End
         Begin VB.OptionButton Option2 
            Caption         =   "แนบวางบิล"
            Height          =   285
            Left            =   1170
            TabIndex        =   218
            Top             =   180
            Width           =   1095
         End
         Begin VB.OptionButton Option1 
            Caption         =   "แนบ ฉ.6"
            Height          =   195
            Left            =   45
            TabIndex        =   217
            Top             =   225
            Width           =   1005
         End
      End
      Begin VB.CheckBox Check1 
         Caption         =   "GPS"
         Height          =   375
         Left            =   -66495
         TabIndex        =   215
         Top             =   4635
         Width           =   690
      End
      Begin VB.TextBox txtDriverName 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   -66285
         TabIndex        =   214
         Top             =   1200
         Width           =   2370
      End
      Begin VB.TextBox txtLengthNum 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   9360
         TabIndex        =   212
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox txtRem 
         Height          =   315
         Left            =   3300
         TabIndex        =   210
         Top             =   7020
         Width           =   3075
      End
      Begin VB.ComboBox cboRem 
         Height          =   315
         Left            =   1080
         TabIndex        =   209
         Top             =   7020
         Width           =   2235
      End
      Begin VB.CommandButton Command4 
         Caption         =   "เพิ่ม/แก้ไข Waste"
         Height          =   405
         Left            =   9780
         TabIndex        =   196
         Top             =   6150
         Width           =   1635
      End
      Begin VB.Frame Frame3 
         Caption         =   "ค่าบริการที่เรียกเก็บของการขนครั้งนี้"
         Height          =   2025
         Left            =   -74880
         TabIndex        =   183
         Top             =   4890
         Width           =   11205
         Begin VB.TextBox txtSumTotal 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C00000&
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   6900
            Locked          =   -1  'True
            TabIndex        =   226
            TabStop         =   0   'False
            Top             =   1935
            Visible         =   0   'False
            Width           =   1005
         End
         Begin VB.CommandButton cmdViewTrip 
            Caption         =   "ดูข้อมูล"
            Height          =   315
            Left            =   3690
            TabIndex        =   69
            Top             =   330
            Width           =   735
         End
         Begin VB.TextBox txtControlDocNo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C00000&
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   1230
            Locked          =   -1  'True
            TabIndex        =   68
            TabStop         =   0   'False
            Top             =   330
            Width           =   2415
         End
         Begin VB.TextBox txtOtherFee2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C00000&
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   3690
            Locked          =   -1  'True
            TabIndex        =   75
            TabStop         =   0   'False
            Top             =   1500
            Width           =   1005
         End
         Begin VB.TextBox txtOtherDesc2 
            Appearance      =   0  'Flat
            BackColor       =   &H00C00000&
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   1230
            Locked          =   -1  'True
            TabIndex        =   74
            TabStop         =   0   'False
            Top             =   1500
            Width           =   2385
         End
         Begin VB.TextBox txtEquipFee 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C00000&
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   3690
            Locked          =   -1  'True
            TabIndex        =   73
            TabStop         =   0   'False
            Top             =   1110
            Width           =   1005
         End
         Begin VB.TextBox txtWorkerFee 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C00000&
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   1230
            Locked          =   -1  'True
            TabIndex        =   72
            TabStop         =   0   'False
            Top             =   1110
            Width           =   1005
         End
         Begin VB.TextBox txtMoreWorker 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C00000&
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   3690
            Locked          =   -1  'True
            TabIndex        =   71
            TabStop         =   0   'False
            Top             =   720
            Width           =   735
         End
         Begin VB.TextBox txtTransFee 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C00000&
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   1230
            Locked          =   -1  'True
            TabIndex        =   70
            TabStop         =   0   'False
            Top             =   720
            Width           =   1005
         End
         Begin MSComctlLib.ListView lvMenifest 
            Height          =   1815
            Left            =   5400
            TabIndex        =   76
            Top             =   150
            Width           =   5715
            _ExtentX        =   10081
            _ExtentY        =   3201
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
               Text            =   "เลขที่"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   1
               Text            =   "ค่าเก็บขน"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   2
               Text            =   "ค่าอื่นๆ"
               Object.Width           =   2646
            EndProperty
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "รวมค่าบริการทั้งหมด"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   86
            Left            =   5400
            TabIndex        =   227
            Top             =   1995
            Visible         =   0   'False
            Width           =   1425
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "บาท"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   87
            Left            =   7980
            TabIndex        =   195
            Top             =   1890
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "เลขที่ใบคุม"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   85
            Left            =   390
            TabIndex        =   194
            Top             =   390
            Width           =   765
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "บาท"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   84
            Left            =   4770
            TabIndex        =   193
            Top             =   1560
            Width           =   285
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ค่าอื่นๆ"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   83
            Left            =   690
            TabIndex        =   192
            Top             =   1560
            Width           =   480
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "บาท"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   82
            Left            =   4770
            TabIndex        =   191
            Top             =   1170
            Width           =   285
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ค่าอุปกรณ์"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   81
            Left            =   2910
            TabIndex        =   190
            Top             =   1170
            Width           =   705
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "บาท"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   80
            Left            =   2310
            TabIndex        =   189
            Top             =   1170
            Width           =   285
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ค่าคนงานเพิ่ม"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   79
            Left            =   210
            TabIndex        =   188
            Top             =   1170
            Width           =   975
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "คน"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   78
            Left            =   4500
            TabIndex        =   187
            Top             =   780
            Width           =   225
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "คนงานเพิ่ม"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   77
            Left            =   2820
            TabIndex        =   186
            Top             =   780
            Width           =   795
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "บาท"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   76
            Left            =   2310
            TabIndex        =   185
            Top             =   780
            Width           =   285
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ค่าขนส่ง"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   75
            Left            =   600
            TabIndex        =   184
            Top             =   780
            Width           =   555
         End
      End
      Begin VB.ComboBox cboCarRegisID 
         Height          =   315
         ItemData        =   "frMenifest.frx":0054
         Left            =   -69165
         List            =   "frMenifest.frx":005E
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   1200
         Width           =   2025
      End
      Begin VB.TextBox txtPlace 
         Appearance      =   0  'Flat
         BackColor       =   &H00C00000&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1770
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1650
         Width           =   6735
      End
      Begin VB.ComboBox cboWType 
         Height          =   315
         ItemData        =   "frMenifest.frx":0070
         Left            =   8400
         List            =   "frMenifest.frx":0072
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   4170
         Width           =   2595
      End
      Begin VB.ComboBox cboActualUnit 
         Height          =   315
         ItemData        =   "frMenifest.frx":0074
         Left            =   -65640
         List            =   "frMenifest.frx":0076
         Style           =   2  'Dropdown List
         TabIndex        =   50
         Top             =   4020
         Width           =   1455
      End
      Begin VB.TextBox txtJobNo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   9390
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   540
         Width           =   1755
      End
      Begin VB.TextBox txtTruckType 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   -73485
         TabIndex        =   33
         Top             =   1200
         Width           =   3435
      End
      Begin VB.TextBox txtAccStaff 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   -66510
         TabIndex        =   87
         TabStop         =   0   'False
         Top             =   6960
         Width           =   2805
      End
      Begin VB.Frame Frame2 
         Caption         =   "ต้นทุนค่าใช้จ่ายสำหรับใบขนนี้"
         Height          =   1365
         Index           =   2
         Left            =   -74880
         TabIndex        =   163
         Top             =   7290
         Visible         =   0   'False
         Width           =   11205
         Begin VB.TextBox txtTotalOut 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C00000&
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   9540
            Locked          =   -1  'True
            TabIndex        =   86
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
            TabIndex        =   82
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
            TabIndex        =   83
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
            TabIndex        =   81
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
            TabIndex        =   80
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
            TabIndex        =   84
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
            TabIndex        =   85
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
            TabIndex        =   181
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
            TabIndex        =   180
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
            TabIndex        =   174
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
            TabIndex        =   173
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
            TabIndex        =   171
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
            TabIndex        =   170
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
            TabIndex        =   169
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
            TabIndex        =   168
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
            TabIndex        =   167
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
            TabIndex        =   166
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
            TabIndex        =   165
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
            TabIndex        =   164
            Top             =   840
            Width           =   285
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "ค่าบริการที่เรียกเก็บลูกค้าของ Manifest ใบนี้"
         Height          =   2955
         Index           =   1
         Left            =   -74880
         TabIndex        =   152
         Top             =   1830
         Width           =   11205
         Begin VB.TextBox txtTotalIn 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C00000&
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   9660
            Locked          =   -1  'True
            TabIndex        =   67
            TabStop         =   0   'False
            Top             =   2760
            Visible         =   0   'False
            Width           =   1005
         End
         Begin WasteManagment.ctlGrid grCustRate 
            Height          =   2145
            Left            =   120
            TabIndex        =   63
            Top             =   240
            Width           =   10995
            _ExtentX        =   19394
            _ExtentY        =   3784
         End
         Begin VB.TextBox txtOtherFee 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   9660
            TabIndex        =   66
            TabStop         =   0   'False
            Top             =   2520
            Width           =   1005
         End
         Begin VB.TextBox txtOtherDesc 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   5880
            MaxLength       =   60
            TabIndex        =   65
            TabStop         =   0   'False
            Top             =   2520
            Width           =   3735
         End
         Begin VB.TextBox txtInEquip 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   6570
            TabIndex        =   79
            TabStop         =   0   'False
            Top             =   2940
            Visible         =   0   'False
            Width           =   1005
         End
         Begin VB.TextBox txtInWorker 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   3990
            TabIndex        =   78
            TabStop         =   0   'False
            Top             =   2940
            Visible         =   0   'False
            Width           =   1005
         End
         Begin VB.TextBox txtIntrans 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   870
            TabIndex        =   77
            TabStop         =   0   'False
            Top             =   2940
            Visible         =   0   'False
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
            TabIndex        =   64
            TabStop         =   0   'False
            Top             =   2520
            Visible         =   0   'False
            Width           =   1005
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "รวม"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   71
            Left            =   9240
            TabIndex        =   179
            Top             =   2820
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "บาท"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   70
            Left            =   10740
            TabIndex        =   178
            Top             =   2820
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Shape Shape1 
            Height          =   2205
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
            Left            =   10770
            TabIndex        =   162
            Top             =   2580
            Width           =   285
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "อื่นๆ ระบุ"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   53
            Left            =   5190
            TabIndex        =   161
            Top             =   2580
            Width           =   615
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "บาท"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   52
            Left            =   7650
            TabIndex        =   160
            Top             =   3450
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ค่าอุปกรณ์"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   51
            Left            =   5820
            TabIndex        =   159
            Top             =   3450
            Visible         =   0   'False
            Width           =   705
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "บาท"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   50
            Left            =   5100
            TabIndex        =   158
            Top             =   3450
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ค่าคนงานเพิ่ม"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   35
            Left            =   2940
            TabIndex        =   157
            Top             =   3450
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "บาท"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   34
            Left            =   1980
            TabIndex        =   156
            Top             =   3450
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ค่าขนส่ง"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   33
            Left            =   240
            TabIndex        =   155
            Top             =   3450
            Visible         =   0   'False
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
            TabIndex        =   154
            Top             =   2580
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ค่าเก็บขน"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   24
            Left            =   360
            TabIndex        =   153
            Top             =   2580
            Visible         =   0   'False
            Width           =   675
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "ตรวจรับเอกสาร"
         Height          =   1395
         Index           =   0
         Left            =   -74880
         TabIndex        =   151
         Top             =   420
         Width           =   11205
         Begin VB.ComboBox CbReState 
            Height          =   315
            ItemData        =   "frMenifest.frx":0078
            Left            =   1605
            List            =   "frMenifest.frx":007A
            TabIndex        =   224
            Text            =   "cboInvState"
            Top             =   990
            Width           =   2955
         End
         Begin VB.CommandButton cmdSaveInv 
            BackColor       =   &H00C0E0FF&
            Caption         =   "บันทึกสาเหตุ"
            Height          =   600
            Left            =   10260
            MaskColor       =   &H00FFFFC0&
            TabIndex        =   207
            Top             =   675
            Width           =   675
         End
         Begin VB.TextBox txtInvDetail 
            Height          =   600
            Left            =   4560
            MultiLine       =   -1  'True
            TabIndex        =   206
            Top             =   675
            Width           =   5595
         End
         Begin VB.ComboBox cboInvState 
            Height          =   315
            ItemData        =   "frMenifest.frx":007C
            Left            =   1605
            List            =   "frMenifest.frx":007E
            TabIndex        =   204
            Text            =   "cboInvState"
            Top             =   675
            Width           =   2955
         End
         Begin VB.ComboBox cboComplete 
            Height          =   315
            Left            =   240
            Style           =   2  'Dropdown List
            TabIndex        =   62
            Top             =   315
            Width           =   2205
         End
         Begin VB.CheckBox chkDoc5 
            Caption         =   "ใบควบคุมผู้ว่าจ้าง"
            Height          =   255
            Left            =   8760
            TabIndex        =   61
            Top             =   315
            Width           =   1605
         End
         Begin VB.CheckBox chkDoc4 
            Caption         =   "ใบน้ำหนักผู้ว่าจ้าง"
            Height          =   255
            Left            =   7200
            TabIndex        =   60
            Top             =   315
            Width           =   1545
         End
         Begin VB.CheckBox chkDoc3 
            Caption         =   "ใบน้ำหนักเข้าศูนย์"
            Height          =   255
            Left            =   5520
            TabIndex        =   59
            Top             =   315
            Width           =   1665
         End
         Begin VB.CheckBox chkDoc2 
            Caption         =   "ชุดใบกำกับ"
            Height          =   255
            Left            =   4380
            TabIndex        =   58
            Top             =   315
            Width           =   1185
         End
         Begin VB.CheckBox chkDoc1 
            Caption         =   "ใบบันทึกการปฏิบัติงาน"
            Height          =   255
            Left            =   2460
            TabIndex        =   57
            Top             =   315
            Width           =   1965
         End
         Begin VB.Label Label6 
            Caption         =   "สาเหตุยังไม่ได้รับเงิน"
            Height          =   315
            Left            =   45
            TabIndex        =   225
            Top             =   1050
            Width           =   1560
         End
         Begin VB.Label Label3 
            Caption         =   "สาเหตุยังไม่วางบิล"
            Height          =   315
            Left            =   240
            TabIndex        =   205
            Top             =   735
            Width           =   1335
         End
      End
      Begin VB.TextBox txtAumphur 
         Appearance      =   0  'Flat
         BackColor       =   &H00C00000&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   -66870
         TabIndex        =   45
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
         TabIndex        =   44
         Top             =   3300
         Width           =   1995
      End
      Begin VB.TextBox txtDestroyer 
         Appearance      =   0  'Flat
         BackColor       =   &H00C00000&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   -73350
         Locked          =   -1  'True
         TabIndex        =   41
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
         TabIndex        =   42
         Top             =   2940
         Width           =   3645
      End
      Begin VB.TextBox txtDestroyPlace 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   -73350
         TabIndex        =   43
         Top             =   3300
         Width           =   2475
      End
      Begin VB.TextBox txtRealWeight 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   -67260
         TabIndex        =   49
         Top             =   4020
         Width           =   1575
      End
      Begin VB.TextBox txtDestER 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   -66270
         TabIndex        =   48
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
         TabIndex        =   47
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
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   3660
         Width           =   2475
      End
      Begin VB.TextBox txtTreatTime 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   -70905
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   4380
         Width           =   1395
      End
      Begin VB.ComboBox cboTimeUnit 
         Height          =   315
         ItemData        =   "frMenifest.frx":0080
         Left            =   -69465
         List            =   "frMenifest.frx":008D
         Style           =   2  'Dropdown List
         TabIndex        =   52
         Top             =   4380
         Width           =   855
      End
      Begin VB.TextBox txtDestName 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   -73350
         TabIndex        =   53
         Top             =   4740
         Width           =   2430
      End
      Begin VB.Frame Frame1 
         Caption         =   "กรณีของเสียอันตรายไม่ตรงตามที่แจ้ง"
         Height          =   2235
         Left            =   -74940
         TabIndex        =   136
         Top             =   5130
         Width           =   11415
         Begin VB.CommandButton cmdDisCrep 
            Caption         =   "เพิ่ม"
            Height          =   315
            Left            =   10650
            TabIndex        =   56
            Top             =   240
            Width           =   705
         End
         Begin MSComctlLib.ListView lvResult 
            Height          =   1905
            Left            =   90
            TabIndex        =   55
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
         TabIndex        =   17
         Top             =   3810
         Width           =   3375
      End
      Begin VB.TextBox txtCollectorID2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   8400
         TabIndex        =   18
         Top             =   3810
         Width           =   2595
      End
      Begin VB.ComboBox cboGenerator 
         Height          =   315
         Left            =   5880
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   900
         Width           =   2715
      End
      Begin VB.ComboBox cboWasteType 
         Height          =   315
         ItemData        =   "frMenifest.frx":00A1
         Left            =   9360
         List            =   "frMenifest.frx":00B1
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   2040
         Width           =   915
      End
      Begin VB.ComboBox cboSiteName 
         Height          =   315
         Left            =   1770
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1260
         Width           =   9375
      End
      Begin VB.ComboBox cboTransUnit 
         Height          =   315
         ItemData        =   "frMenifest.frx":00C9
         Left            =   -65100
         List            =   "frMenifest.frx":00D3
         Style           =   2  'Dropdown List
         TabIndex        =   38
         Top             =   1950
         Width           =   885
      End
      Begin VB.ComboBox cboSolidUnit 
         Height          =   315
         Left            =   8400
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   6210
         Width           =   1305
      End
      Begin VB.ComboBox cboLiquidUnit 
         Height          =   315
         Left            =   4830
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   6210
         Width           =   1305
      End
      Begin VB.TextBox txtCollectorID 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   8400
         TabIndex        =   16
         Top             =   3450
         Width           =   2595
      End
      Begin VB.TextBox txtCollector 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1770
         TabIndex        =   15
         Top             =   3450
         Width           =   3375
      End
      Begin VB.Frame Frame4 
         Caption         =   "ผู้ขนส่งของเสียอัตราย"
         Height          =   1005
         Left            =   90
         TabIndex        =   124
         Top             =   2400
         Width           =   11025
         Begin VB.TextBox txtTrans1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C00000&
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   11
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
            TabIndex        =   12
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
            TabIndex        =   14
            Top             =   600
            Width           =   2595
         End
         Begin VB.TextBox txtTrans2 
            Appearance      =   0  'Flat
            BackColor       =   &H00FF8080&
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   1680
            TabIndex        =   13
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
            TabIndex        =   128
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
            TabIndex        =   127
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
            TabIndex        =   126
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
            TabIndex        =   125
            Top             =   660
            Width           =   765
         End
      End
      Begin VB.TextBox txtSign1B 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   -73080
         TabIndex        =   39
         Top             =   2340
         Width           =   4905
      End
      Begin VB.TextBox txtTransTime 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   -66480
         TabIndex        =   37
         Top             =   1950
         Width           =   1335
      End
      Begin VB.TextBox txtToProvince 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   -70080
         TabIndex        =   36
         Top             =   1950
         Width           =   1905
      End
      Begin VB.TextBox txtFromProvince 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   -73080
         TabIndex        =   35
         Top             =   1950
         Width           =   1875
      End
      Begin VB.TextBox txtTransTel 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   -73485
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   840
         Width           =   2445
      End
      Begin VB.TextBox txtTransFax 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   -70125
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   840
         Width           =   2130
      End
      Begin VB.TextBox txtTransER 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   -67005
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   840
         Width           =   3105
      End
      Begin VB.TextBox txtTransID1B 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   -67755
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   480
         Width           =   3855
      End
      Begin VB.TextBox txtTrans1B 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   -73485
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   480
         Width           =   3795
      End
      Begin VB.TextBox txtCustER 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   6450
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   2040
         Width           =   1425
      End
      Begin VB.TextBox txtSigneName1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   7080
         TabIndex        =   26
         Top             =   7020
         Width           =   1845
      End
      Begin VB.TextBox txtTreatment 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   3360
         TabIndex        =   25
         Top             =   6600
         Width           =   8055
      End
      Begin VB.TextBox txtLiquidQty 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   3360
         TabIndex        =   21
         Top             =   6210
         Width           =   1425
      End
      Begin VB.TextBox txtSolidQty 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   6990
         TabIndex        =   23
         Top             =   6210
         Width           =   1365
      End
      Begin VB.TextBox txtCustFax 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   3870
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   2040
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
         Top             =   2040
         Width           =   1395
      End
      Begin VB.TextBox txtCustomer 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         Height          =   315
         Left            =   1800
         TabIndex        =   1
         Top             =   540
         Width           =   6825
      End
      Begin WasteManagment.ctlGrid grItem 
         Height          =   1395
         Left            =   120
         TabIndex        =   20
         Top             =   4680
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   2461
      End
      Begin WasteManagment.ctlDate dtWorkDate 
         Height          =   315
         Left            =   9390
         TabIndex        =   4
         Top             =   900
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   556
         BackColor       =   -2147483633
      End
      Begin WasteManagment.ctlDate dtSign2 
         Height          =   315
         Left            =   -66480
         TabIndex        =   40
         Top             =   2310
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   556
         BackColor       =   -2147483633
      End
      Begin WasteManagment.ctlDate dtSign1 
         Height          =   315
         Left            =   9660
         TabIndex        =   27
         Top             =   7020
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   556
         BackColor       =   -2147483633
      End
      Begin WasteManagment.ctlDate dtSign3 
         Height          =   315
         Left            =   -70410
         TabIndex        =   54
         Top             =   4740
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   556
         BackColor       =   -2147483633
      End
      Begin VB.Label Label7 
         Caption         =   "เลขที่ 14 หลัก"
         Height          =   195
         Left            =   720
         TabIndex        =   229
         Top             =   900
         Width           =   975
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "เรท"
         Height          =   195
         Left            =   -68160
         TabIndex        =   220
         Top             =   4770
         Width           =   270
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "พนักงานขับ"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   90
         Left            =   -67125
         TabIndex        =   213
         Top             =   1260
         Width           =   810
      End
      Begin VB.Label Label4 
         Caption         =   "ระยะทาง :"
         Height          =   255
         Left            =   8640
         TabIndex        =   211
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "หมายเหตุ :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   89
         Left            =   180
         TabIndex        =   208
         Top             =   7020
         Width           =   765
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "สถานที่ก่อกำเนิด"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   74
         Left            =   510
         TabIndex        =   182
         Top             =   1710
         Width           =   1170
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "คิดค่าบริการจากปริมาณของ"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   69
         Left            =   6450
         TabIndex        =   177
         Top             =   4230
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
         TabIndex        =   175
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
         TabIndex        =   172
         Top             =   6975
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
         TabIndex        =   150
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
         TabIndex        =   149
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
         TabIndex        =   148
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
         TabIndex        =   147
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
         Left            =   -68505
         TabIndex        =   146
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
         TabIndex        =   145
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
         Left            =   -74685
         TabIndex        =   144
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
         TabIndex        =   143
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
         TabIndex        =   142
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
         TabIndex        =   141
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
         TabIndex        =   140
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
         TabIndex        =   139
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
         TabIndex        =   138
         Top             =   4800
         Width           =   1035
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "วันที่"
         Height          =   195
         Index           =   3
         Left            =   -70830
         TabIndex        =   137
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
         Left            =   8940
         TabIndex        =   135
         Top             =   7080
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
         TabIndex        =   134
         Top             =   3870
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
         TabIndex        =   133
         Top             =   3870
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
         TabIndex        =   132
         Top             =   2100
         Width           =   1095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ประเภทรถขนส่ง"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   7
         Left            =   -74685
         TabIndex        =   130
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
         TabIndex        =   129
         Top             =   3510
         Width           =   1005
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "วันที่"
         Height          =   195
         Index           =   0
         Left            =   -66840
         TabIndex        =   123
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
         TabIndex        =   122
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
         TabIndex        =   121
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
         TabIndex        =   120
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
         TabIndex        =   119
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
         TabIndex        =   118
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
         Left            =   -70005
         TabIndex        =   117
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
         Left            =   -70725
         TabIndex        =   116
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
         Left            =   -74175
         TabIndex        =   115
         Top             =   900
         Width           =   705
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ฉุกเฉิน"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   29
         Left            =   -67605
         TabIndex        =   114
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
         Left            =   -69195
         TabIndex        =   113
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
         Left            =   -74505
         TabIndex        =   112
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
         TabIndex        =   111
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
         Left            =   6660
         TabIndex        =   110
         Top             =   7080
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
         TabIndex        =   109
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
         TabIndex        =   108
         Top             =   2100
         Width           =   885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "รายละเอียดของของเสียอันตรายที่ขนส่งเคลื่อนย้าย"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   106
         Top             =   4470
         Width           =   3435
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   45
         Index           =   1
         Left            =   120
         Top             =   4560
         Width           =   11325
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ของแข็ง"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   11
         Left            =   6360
         TabIndex        =   105
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
         TabIndex        =   104
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
         TabIndex        =   103
         Top             =   3510
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
         TabIndex        =   102
         Top             =   2100
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
         TabIndex        =   101
         Top             =   2100
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
         TabIndex        =   100
         Top             =   960
         Width           =   2115
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "สถานที่ขนกาก"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   690
         TabIndex        =   99
         Top             =   1320
         Width           =   990
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ชื่อลูกค้า"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   0
         Left            =   1140
         TabIndex        =   98
         Top             =   570
         Width           =   570
      End
   End
   Begin VB.TextBox txtTR_Remark 
      Appearance      =   0  'Flat
      ForeColor       =   &H00FF0000&
      Height          =   735
      Left            =   1200
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   199
      Top             =   8490
      Width           =   4935
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "อ้างอิง Better ME"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   91
      Left            =   4725
      TabIndex        =   223
      Top             =   780
      Width           =   1155
   End
   Begin VB.Label lblLastApprove 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      Caption         =   "อนุมัติโดย :"
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   10725
      TabIndex        =   203
      Top             =   840
      Width           =   795
   End
   Begin VB.Label lbApprove 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "APPROVED"
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
      Height          =   555
      Left            =   1920
      TabIndex        =   202
      Top             =   540
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Label lblLastUpdate 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "แก้ไขล่าสุดโดย :"
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   8040
      TabIndex        =   201
      Top             =   660
      Width           =   3480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "AP"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   88
      Left            =   9600
      TabIndex        =   198
      Top             =   420
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Label lbCancel 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "CANCELED"
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
      Height          =   555
      Left            =   1920
      TabIndex        =   176
      Top             =   0
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Menifest No."
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   36
      Left            =   6840
      TabIndex        =   131
      Top             =   420
      Width           =   915
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "เลขที่"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   13
      Left            =   4800
      TabIndex        =   107
      Top             =   420
      Width           =   375
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   4
      Index           =   0
      X1              =   90
      X2              =   4620
      Y1              =   510
      Y2              =   540
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      Index           =   1
      X1              =   60
      X2              =   4560
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
      TabIndex        =   97
      Top             =   60
      Width           =   1560
   End
   Begin VB.Image Image2 
      Height          =   1080
      Left            =   0
      Picture         =   "frMenifest.frx":00E5
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11655
   End
End
Attribute VB_Name = "frmMenifest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim myMenifest_ As BWGMenifest
Dim curCustomer_ As BWGCustomer
Dim curCustomerSite_ As BWGCustomerSite
Dim curRem_ As BWGMenifestHeaderRem
Dim curTrans2_ As BWGTransportVendor
Dim curTruckSubType_ As BWGTruckSubType
Dim curFromProvince_ As IMIProvince
Dim curToProvince_ As IMIProvince
Dim provLocation_ As String
Dim curTSDFTumbol_ As IMITumbol
Dim isDetailChange_ As Boolean
Dim curRole_ As AWSRoleType

Private Sub SetMSFlexGrid()
On Error Resume Next
    With grItem.getGridObj
        grItem.ColInputType(9) = ComboBoxEnum
        grItem.ColInputType(11) = ComboBoxEnum
        .Cols = 12
        .ColWidth(0) = 0
        .ColWidth(1) = 495
        .ColWidth(2) = 3540
        .ColWidth(3) = 1110
        .ColWidth(4) = 795
        .ColWidth(5) = 1110
        .ColWidth(6) = 1440
        .ColWidth(7) = 1350
        .ColWidth(8) = 1200
        .ColWidth(9) = 1000
        .ColWidth(10) = 1000
        .ColWidth(11) = 1000
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
        .FixedAlignment(9) = 4
        .FixedAlignment(10) = 4
         .FixedAlignment(11) = 4
        .ColAlignment(0) = 4
        .ColAlignment(1) = 7
        .ColAlignment(2) = 1
        .ColAlignment(3) = 4
        .ColAlignment(4) = 7
        .ColAlignment(5) = 4
        .ColAlignment(6) = 7
        .ColAlignment(7) = 7
        .ColAlignment(8) = 1
        .ColAlignment(9) = 1
        .ColAlignment(10) = 1
        .ColAlignment(11) = 1
        .TextMatrix(0, 1) = "ลำดับ"
        .TextMatrix(0, 2) = "รายละเอียด"
        .TextMatrix(0, 3) = "รหัสของเสีย"
        .TextMatrix(0, 4) = "จำนวน"
        .TextMatrix(0, 5) = "ชนิด"
        .TextMatrix(0, 6) = "ปริมาณผู้ก่อกำเนิด"
        .TextMatrix(0, 7) = "ปริมาณผู้รับกำจัด"
        .TextMatrix(0, 8) = "หน่วย"
        .TextMatrix(0, 9) = "Waste ใหม่"
        .TextMatrix(0, 10) = "น้ำหนักเข้า"
        .TextMatrix(0, 11) = "Waste ใหม่(CR)"
    End With

    With grCustRate.getGridObj
        .Cols = 5
        .Rows = 50
        .ColWidth(0) = 0
        .ColWidth(1) = 5475
        .ColWidth(2) = 1755
        .ColWidth(3) = 1605
        .ColWidth(4) = 0
        .TextMatrix(0, 1) = "ชื่อกากของเสีย"
        .TextMatrix(0, 2) = "ค่ากำจัดกาก (Quote)"
        .TextMatrix(0, 3) = "ค่าขนส่ง (Quote)"
        .TextMatrix(0, 4) = "" '"ค่ากำจัดกาก (เรียกเก็บ)"
    End With
End Sub

Private Sub ShowTR_Remark()
On Error Resume Next
Dim rs As ADODB.Recordset
Dim tSql As String, i As Integer
    tSql = "SELECT * FROM tbCustomer_TR_Remark" & vbCrLf
    tSql = tSql & "WHERE CustomerID = '" & myMenifest_.CustomerID & "'" & vbCrLf
    tSql = tSql & "ORDER BY RemarkNo"
    Set rs = New ADODB.Recordset
    Set rs = GetRS(tSql, adOpenForwardOnly, adLockReadOnly)
    If rs.RecordCount > 0 Then
        For i = 1 To rs.RecordCount
            If i > 1 Then
                txtTR_Remark.Text = txtTR_Remark.Text & vbCrLf & i & "." & rs!RemarkText
            Else
                 txtTR_Remark.Text = i & "." & rs!RemarkText
            End If
            rs.MoveNext
        Next
    End If
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

Private Sub cboCarRegisID_Validate(Cancel As Boolean)
Dim tmpRecT As ADODB.Recordset
    Dim sqlStrT As String
    sqlStrT = " select VendorTruckID,TripNo,DriverName from tbJobDataTimeTable where JobDataCarID = '" & myMenifest_.JobDataCarID & "' and TripNo = '" & myMenifest_.TripNo & "' "
    Set tmpRecT = GetRS(sqlStrT, adOpenForwardOnly, adLockReadOnly) 'ค้นหา vendor
    Dim tmpRecV As ADODB.Recordset
    Dim sqlStrV As String
    sqlStrV = " SELECT  VendorID,VendorTruckID ,ContactID,ContactName ,SubTypeID  FROM vw_VendorCarSearch " & vbCrLf
    sqlStrV = sqlStrV & " Where  VendorTruckID = '" & tmpRecT(0) & "' " & vbCrLf
    Set tmpRecV = GetRS(sqlStrV, adOpenForwardOnly, adLockReadOnly) 'ค้นหา vendorID
    
    
    Dim tmpRec As ADODB.Recordset
    myMenifest_.CarRegisID = cboCarRegisID.Text
    Dim sqlStr As String
    sqlStr = "SELECT ContactID,ContactName ,SubTypeID  FROM vw_VendorCarSearch where  SubTypeID = '" & myMenifest_.TruckSubTypeID & "' and VendorID = '" & tmpRecV(0) & "' " & vbCrLf
    sqlStr = sqlStr & " And CarRegisID = '" & cboCarRegisID.Text & "' " & vbCrLf
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec.EOF = False Then
     myMenifest_.DriverName = tmpRec(1)
     txtDriverName.Text = myMenifest_.DriverName
    myMenifest_.isChanged = True
     Else
     MsgBox "ไม่พบชื่อคนขับโปรดตรวจสอบ ทะเทียนกับคนขับในข้อมูล Vendor", vbExclamation
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

Private Sub cboComplete_Validate(Cancel As Boolean)
    If cboComplete.ListIndex = 1 Then
        myMenifest_.isDocCompleted = True
    Else
        myMenifest_.isDocCompleted = False
    End If
End Sub
Private Sub cboInvState_Validate(Cancel As Boolean)
'    If cboInvState.ListIndex > 0 Then
'        myMenifest_.InvState = cboInvState.Text
'    Else
'        myMenifest_.InvState = ""
'    End If
End Sub


Private Sub cboGenerator_Validate(Cancel As Boolean)
    If cboGenerator.ListIndex > 0 Then
        myMenifest_.WasteGenNo = cboGenerator.Text
    Else
        myMenifest_.WasteGenNo = ""
    End If
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

Private Sub cboRem_Validate(Cancel As Boolean)
    If cboRem.ListIndex > -1 Then
        myMenifest_.RemID = cboRem.Text
    Else
        myMenifest_.RemID = ""
    End If
End Sub

Private Sub cboSiteName_Click()
Dim tmpSite As BWGCustomerSite
    If curCustomer_ Is Nothing Then
        txtPlace.Text = ""
    Else
        If Trim(curCustomer_.id) = "" Then
            txtPlace.Text = ""
        Else
            If cboSiteName.ListIndex = -1 Then
                txtPlace.Text = curCustomer_.Address & " " & curCustomer_.TumbolName & " " & curCustomer_.AumphurName & " " & curCustomer_.ProvinceName & " " & curCustomer_.ZipCode
            Else
                Set tmpSite = curCustomer_.SiteAddresses(cboSiteName.ListIndex + 1)
                If tmpSite.isManifestAddr Then
                    txtPlace.Text = tmpSite.Address & " " & tmpSite.TumbolName & " " & tmpSite.AumphurName & " " & tmpSite.ProvinceName & " " & tmpSite.ZipCode
                Else
                    txtPlace.Text = curCustomer_.Address & " " & curCustomer_.TumbolName & " " & curCustomer_.AumphurName & " " & curCustomer_.ProvinceName & " " & curCustomer_.ZipCode
                End If
                txtLengthNum.Text = tmpSite.LengthNum
                Set tmpSite = Nothing
            End If
        End If
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
Dim x%
    With grItem.getGridObj
        If cboWType.ListIndex = 0 Then
            myMenifest_.ChargeWType = "A" 'ผู้รับกำจัด
            For x = 1 To .Rows - 1
                If Trim(.TextMatrix(x, 0)) <> "" Then
                    grItem_UpdateValue x, 7
                End If
            Next
        ElseIf cboWType.ListIndex = 1 Then
            myMenifest_.ChargeWType = "B"   'ผู้ก่อกำเนิด
            For x = 1 To .Rows - 1
                If Trim(.TextMatrix(x, 0)) <> "" Then
                    grItem_UpdateValue x, 6
                End If
            Next
        End If
    End With
End Sub

Private Sub Check1_Click()
    If Check1.Value = 1 Then
       sqlStr = "UPDATE tbJobDataTimeTable SET gpssend='" & Trim(Check1.Value) & "' " & vbCrLf
        sqlStr = sqlStr & "WHERE JobDataCarID = '" & myMenifest_.JobDataCarID & "' AND TripNO = '" & myMenifest_.TripNo & "'"
        DBConnExc sqlStr
         Frame5.Enabled = True
    Else
       sqlStr = "UPDATE tbJobDataTimeTable SET gpssend='" & Trim(Check1.Value) & "',SendDesc=''  " & vbCrLf
        sqlStr = sqlStr & "WHERE JobDataCarID = '" & myMenifest_.JobDataCarID & "' AND TripNO = '" & myMenifest_.TripNo & "'"
        DBConnExc sqlStr
        Frame5.Enabled = False
    End If
End Sub

Private Sub chkDoc1_Click()
    If Me.Visible Then
        If Me.ActiveControl.Name = "chkDoc1" Then
            If chkDoc1.Value = 1 And chkDoc2.Value = 1 And chkDoc3.Value = 1 And chkDoc4.Value = 1 And chkDoc5.Value = 1 Then
                cboComplete.ListIndex = 1
                myMenifest_.isDocCompleted = True
            End If
        End If
    End If
End Sub

Private Sub chkDoc1_Validate(Cancel As Boolean)
    myMenifest_.DocJobRecord = chkDoc1.Value = 1
End Sub

Private Sub chkDoc2_Click()
    If Me.Visible Then
        If Me.ActiveControl.Name = "chkDoc2" Then
            If chkDoc1.Value = 1 And chkDoc2.Value = 1 And chkDoc3.Value = 1 And chkDoc4.Value = 1 And chkDoc5.Value = 1 Then
                cboComplete.ListIndex = 1
                myMenifest_.isDocCompleted = True
            End If
        End If
    End If
End Sub

Private Sub chkDoc2_Validate(Cancel As Boolean)
    myMenifest_.DocControl = chkDoc2.Value = 1
End Sub

Private Sub chkDoc3_Click()
    If Me.Visible Then
        If Me.ActiveControl.Name = "chkDoc3" Then
            If chkDoc1.Value = 1 And chkDoc2.Value = 1 And chkDoc3.Value = 1 And chkDoc4.Value = 1 And chkDoc5.Value = 1 Then
                cboComplete.ListIndex = 1
                myMenifest_.isDocCompleted = True
            End If
        End If
    End If
End Sub

Private Sub chkDoc3_Validate(Cancel As Boolean)
    myMenifest_.DocSiteWeight = chkDoc3.Value = 1
End Sub

Private Sub chkDoc4_Click()
    If Me.Visible Then
        If Me.ActiveControl.Name = "chkDoc4" Then
            If chkDoc1.Value = 1 And chkDoc2.Value = 1 And chkDoc3.Value = 1 And chkDoc4.Value = 1 And chkDoc5.Value = 1 Then
                cboComplete.ListIndex = 1
                myMenifest_.isDocCompleted = True
            End If
        End If
    End If
End Sub

Private Sub chkDoc4_Validate(Cancel As Boolean)
    myMenifest_.DocCustWeight = chkDoc4.Value = 1
End Sub

Private Sub chkDoc5_Click()
    If Me.Visible Then
        If Me.ActiveControl.Name = "chkDoc5" Then
            If chkDoc1.Value = 1 And chkDoc2.Value = 1 And chkDoc3.Value = 1 And chkDoc4.Value = 1 And chkDoc5.Value = 1 Then
                cboComplete.ListIndex = 1
                myMenifest_.isDocCompleted = True
            End If
        End If
    End If
End Sub

Private Sub chkDoc5_Validate(Cancel As Boolean)
    myMenifest_.DocCustControl = chkDoc5.Value = 1
End Sub

Private Sub cmdCancel_Click()
Dim tmpStr As String, CancelDetail As String
    tmpStr = CurrentUser.VerifyAccessRole(Me.Name, DeleteRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    tmpStr = ""
    If Trim(myMenifest_.BillingID) <> "" Then
        MsgBox "ไม่สามารถยกเลิกใบ Menifest นี้ได้เนื่องจากได้ทำการวางบิลเก็บเงินลูกค้าแล้ว", vbExclamation
        Exit Sub
    End If
    If Trim(myMenifest_.BillPayID) <> "" Then
        MsgBox "ไม่สามารถยกเลิกใบ Menifest นี้ได้เนื่องจากได้ทำการจ่ายเงิน Vendor แล้ว", vbExclamation
        Exit Sub
    End If
    If MsgBox("ต้องการยกเลิกใบ Menifest นี้หรือไม่?", vbQuestion + vbYesNo) = vbYes Then
'        CancelDetail = InputBox("ใส่รายละเอียดในการยกเลิก", "ใส่รายละเอียดการยกเลิก")
        CancelDetail = frmCancelDetail.CancelText(myMenifest_.CancelDetail)
        If Trim(Replace(CancelDetail, "*", "")) = "" Then
            MsgBox "กรุณาระบุเหตุผลของการยกเลิก", vbExclamation + vbOKOnly, "Warning"
            Exit Sub
        End If
        myMenifest_.CancelDetail = CancelDetail
        tmpStr = myMenifest_.Cancel(CancelDetail)
        If Trim(tmpStr) = "" Then
            SetCancelScreen
        Else
            MsgBox tmpStr
        End If
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
    
    If myMenifest_.isChanged And cmdSave.Enabled = True Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนสร้างใหม่หรือไม่?", vbQuestion + vbYesNoCancel)
        If selResult = vbYes Then
            CmdSave_Click
            If myMenifest_.isChanged Then Exit Sub
        ElseIf selResult = vbCancel Then
            Exit Sub
        End If
    End If
    CreateNewMenifest
    ClearScreen
    txtCustomer.SetFocus
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub CmdSave_Click()
Dim tmpStr As String
    If myMenifest_.isClosed Then
        tmpStr = CurrentUser.VerifyAccessRole(Me.Name, ApproveRole)
    Else
        tmpStr = CurrentUser.VerifyAccessRole(Me.Name, curRole_)
    End If
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    
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
    If Trim(txtTrans2.Text) = "" Then
        MsgBox "กรุณาระบุชื่อบริษัทผู้ขนส่งของเสียอันตรายที่ 2", vbExclamation
        txtTrans2.SetFocus
        Exit Sub
    End If
    If Trim(txtTruckType.Text) = "" Then
        MsgBox "กรุณาระบุประเภทรถขนส่ง", vbExclamation
        txtTruckType.SetFocus
        Exit Sub
    End If
    If cboCarRegisID.ListIndex = -1 Then
        MsgBox "กรุณาระบุเลขทะเบียนรถที่ขนส่ง", vbExclamation
        cboCarRegisID.SetFocus
        Exit Sub
    End If
    If curCustomerSite_.isChanged And cmdSave.Enabled = True Then
     curCustomerSite_.LengthNum = Trim(txtLengthNum.Text)
     curCustomerSite_.Save
    End If
    If myMenifest_.isChanged And cmdSave.Enabled = True Then
        myMenifest_.SaleDistrictID = curCustomer_.SaleDistrictID
        myMenifest_.UnderStaffID = curCustomer_.UnderSaleStaffID
        myMenifest_.Save
        txtDocNo.Text = myMenifest_.DocNo
        txtMenifestNo.Text = myMenifest_.MenifestNo
        
    sqlStr = "UPDATE tbJobDataTimeTable SET  DriverName = '" & myMenifest_.DriverName & "'" & vbCrLf
    sqlStr = sqlStr & "WHERE JobDataCarID = '" & myMenifest_.JobDataCarID & "' "
    DBConnExc sqlStr
    Else
        MsgBox "ไม่มีการเปลี่ยนแปลงข้อมูล ข้อมูลไม่ถูกบันทึก", vbInformation, "Info"
    End If
    If myMenifest_.isSave = True Then MsgBox "จัดเก็บข้อมูลเรียบร้อยแล้ว", vbInformation
End Sub

Private Sub cmdSaveInv_Click()
    Dim strSql As String
    Dim rs As ADODB.Recordset
Dim tSql As String, i As Integer
Dim Doc As String
    tSql = "SELECT docno FROM vw_MenifestQue" & vbCrLf
    tSql = tSql & "WHERE jobno = '" & txtJobNo.Text & "'" & vbCrLf
    tSql = tSql & "ORDER BY docno"
    Set rs = New ADODB.Recordset
    Set rs = GetRS(tSql, adOpenForwardOnly, adLockReadOnly)
    If rs.RecordCount > 0 Then
        For i = 1 To rs.RecordCount
        Doc = rs!DocNo
            strSql = "UPDATE tbMenifestHeader Set InvState = '" & cboInvState.Text & "', restate='" & CbReState.Text & "' , InvDetail = '" & txtInvDetail.Text & "'  WHERE DocNo = '" & Doc & "' "
            myMenifest_.Update strSql
            rs.MoveNext
        Next
    End If
    MsgBox "จัดเก็บข้อมูลเรียบร้อยแล้ว", vbInformation
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

Private Sub cmdMenifestClosed_Click()
Dim selResult As Integer
Dim x%, tmpAddWaste As BWGMenifestWasteAdded
Dim tmpJobDet As BWGJobDataDetail
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole(frmMenifestClose.Name, AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    
    If Trim(txtCustomer.Text) = "" Then
        MsgBox "กรุณาระบุลูกค้า", vbExclamation
        txtCustomer.SetFocus
        Exit Sub
    End If
    
    If myMenifest_.isClosed = False Then
        For x = 1 To myMenifest_.JobDetails.Count
            Set tmpJobDet = myMenifest_.JobDetails(x)
            If tmpJobDet.DisposerWeight = 0 And Trim(tmpJobDet.FailType) = "" Then
                MsgBox "กรุณาระบุปริมาณผู้รับกำจัดให้ครบถ้วนทุกรายการของเสีย", vbExclamation
                SSTab1.Tab = 0
                grItem.SetFocus
                Set tmpJobDet = Nothing
                Exit Sub
            'Else
            '    If tmpJobDet.QuoTreatmentUnitName = "บาท/เที่ยว" Then
            '        tmpJobDet.TreatmentCharge = tmpJobDet.QuoTreatmentRate
            '    Else
            '        If cboWType.ListIndex = 0 Then
            '            tmpJobDet.TreatmentCharge = tmpJobDet.QuoTreatmentRate * tmpJobDet.DisposerWeight
            '        End If
            '    End If
            '    grCustRate.getGridObj.TextMatrix(x, 4) = tmpJobDet.TreatmentCharge
            End If
            Set tmpJobDet = Nothing
        Next
        For x = 1 To myMenifest_.AddedWastes.Count
            Set tmpAddWaste = myMenifest_.AddedWastes(x)
            If tmpAddWaste.DisposerWeight = 0 Then
                MsgBox "กรุณาระบุปริมาณผู้รับกำจัดให้ครบถ้วนทุกรายการของเสีย", vbExclamation
                SSTab1.Tab = 0
                grItem.SetFocus
                Set tmpAddWaste = Nothing
                Exit Sub
            'Else
            '    If tmpAddWaste.QuoTreatmentUnitName = "บาท/เที่ยว" Then
            '        tmpAddWaste.TreatmentCharge = tmpAddWaste.QuoTreatmentRate
            '    Else
            '        If cboWType.ListIndex = 1 Then
            '            tmpAddWaste.TreatmentCharge = tmpAddWaste.QuoTreatmentRate * tmpAddWaste.DisposerWeight
            '        End If
            '    End If
            '    grCustRate.getGridObj.TextMatrix(x, 4) = tmpAddWaste.TreatmentCharge
            End If
            Set tmpAddWaste = Nothing
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
    
    
    tmpStr = CurrentUser.VerifyAccessRole(frmMenifestClose.Name, UpdateRole)
    If Trim(tmpStr) <> "" Then
        Load frmMenifestClose
        With frmMenifestClose
            .SetMenifestData myMenifest_
            .Show 1
        End With
        myMenifest_.isChanged = False
        Exit Sub
    Else
        If myMenifest_.isChanged And cmdSave.Enabled = True Then
            selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนทำงานต่อหรือไม่?", vbQuestion + vbYesNo)
            If selResult = vbYes Then
                CmdSave_Click
                If myMenifest_.isChanged Then Exit Sub
            Else 'If selResult = vbCancel Then
                Exit Sub
            End If
        End If
    End If
    
    Load frmMenifestClose
    With frmMenifestClose
        .SetMenifestData myMenifest_
        .Show 1
    End With
End Sub

Private Sub cmdDisCrep_Click()
Dim selResult As Integer
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole(frmMenifestDiscrep.Name, AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    
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

Public Sub CalcTrip()
    cmdViewTrip_Click
End Sub
    
Private Sub cmdViewTrip_Click()
Dim tmpCol As Collection, tmpS As New BWGSearchManager
Dim tmpTime As BWGJobDataTimeTable, tmpAddWaste As BWGMenifestWasteAdded
Dim aumphurID1 As String, x%, tmpJobDet As BWGJobDataDetail
Dim selResult As Integer
Dim tmpStr As String
    If AutoCalcMF = True Then GoTo SetTrip
    tmpStr = CurrentUser.VerifyAccessRole(frmTripControl.Name, AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    
    For x = 1 To myMenifest_.JobDetails.Count
        Set tmpJobDet = myMenifest_.JobDetails(x)
        If tmpJobDet.DisposerWeight = 0 And Trim(tmpJobDet.FailType) = "" Then
            MsgBox "กรุณาระบุปริมาณผู้รับกำจัดให้ครบถ้วนทุกรายการของเสีย", vbExclamation
            SSTab1.Tab = 0
            grItem.SetFocus
            Set tmpJobDet = Nothing
            Exit Sub
        End If
        Set tmpJobDet = Nothing
    Next
    
    For x = 1 To myMenifest_.AddedWastes.Count
        Set tmpAddWaste = myMenifest_.AddedWastes(x)
        If tmpAddWaste.DisposerWeight = 0 Then
            MsgBox "กรุณาระบุปริมาณผู้รับกำจัดให้ครบถ้วนทุกรายการของเสีย", vbExclamation
            SSTab1.Tab = 0
            grItem.SetFocus
            Set tmpAddWaste = Nothing
            Exit Sub
        End If
        Set tmpAddWaste = Nothing
    Next
    
    If myMenifest_.isChanged And cmdSave.Enabled = True Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนทำงานต่อหรือไม่?", vbQuestion + vbYesNo)
        If selResult = vbYes Then
            CmdSave_Click
            If myMenifest_.isChanged Then Exit Sub
        ElseIf selResult = vbNo Then
            MsgBox "ไม่สามารถดูข้อมูลใบขนได้เพราะยังไม่ได้จัดเก็บข้อมูล", vbExclamation
            Exit Sub
        End If
    End If
    
SetTrip:
    'Set tmpCol = tmpS.JobDataTimeTableQuickSearch("JobDataCarID='" & myMenifest_.JobDataCarID & "' And TripNo=" & myMenifest_.tripNo)
    Set tmpCol = tmpS.JobDataTimeTableQuickSearch("TimeTableNo='" & txtControlDocNo.Text & "' And isMain='Y' And isCancel ='N' AND ISNULL(isCanceled, 'N') <> 'Y' AND TripNo = " & myMenifest_.TripNo & " AND JobDataCarID = '" & myMenifest_.JobDataCarID & "'")
    If tmpCol.Count > 0 Then
        Set tmpTime = tmpCol(1)
        Load frmTripControl
        With frmTripControl
            If cboSiteName.ListIndex > -1 Then
                aumphurID1 = Trim(curCustomer_.SiteAddresses(cboSiteName.ListIndex + 1).AumphurID)
            End If
            .SetTripControl tmpTime, txtCustomer.Text, txtTruckType, cboCarRegisID.Text, aumphurID1, myMenifest_.TSDFAumphurID, myMenifest_.TruckTypeID
            If AutoCalcMF = False Then
                .Show 1
            Else
                frmTripControl.SaveData
                frmMain.Stb1.Panels(1).Text = "จัดเก็บข้อมูลใบคุมเลขที่ " & txtControlDocNo.Text
            End If
        End With
        With tmpTime
            txtTransFee.Text = tmpTime.TripTranFee
            txtMoreWorker.Text = tmpTime.TripMoreWorker
            txtWorkerFee.Text = tmpTime.TripMoreWorkerFee
            txtEquipFee.Text = tmpTime.TripEquipFee
            txtOtherDesc2.Text = tmpTime.TripOtherDesc
            txtOtherFee2.Text = tmpTime.TripOtherFee
        End With
    End If
End Sub

Private Sub Command4_Click()
Dim tmpStr As String
Dim compID As String
    tmpStr = CurrentUser.VerifyAccessRole(frmJobWasteDataAdd.Name, AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    If cboCompany.ListIndex = 1 Then
        compID = "776EDA5830F744A6B3551348D851BC22"
    Else
        compID = ""
    End If
    
    If Trim(txtCustomer.Text) = "" Then
        MsgBox "กรุณาเลือกลูกค้าก่อนทำการเพิ่มรายการกากของเสีย", vbExclamation
        txtCustomer.SetFocus
        Exit Sub
    End If
    If cboCompany.ListIndex < 0 Then
        MsgBox "กรุณาระบุบริษัทที่ออกใบ Manifest ก่อนทำการเพิ่มกากของเสีย", vbExclamation
        cboCompany.SetFocus
        Exit Sub
    End If
    isDetailChange_ = False
    Load frmJobWasteDataAdd
    With frmJobWasteDataAdd
        If cboGenerator.ListIndex = 0 Then
            .SetCallerForm Me, curCustomer_, "", compID, "", myMenifest_.TruckTypeID
        Else
            .SetCallerForm Me, curCustomer_, cboGenerator.Text, compID, "", myMenifest_.TruckTypeID
        End If
        .Show 1
        If isDetailChange_ Then
            ShowWasteDetail
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
    With cboComplete
        .AddItem "เอกสารยังไม่ครบถ้วน"
        .AddItem "เอกสารครบถ้วนแล้ว"
        .ListIndex = 0
    End With
    With cboWType
        .AddItem "ปริมาณผู้รับกำจัด"
        .AddItem "ปริมาณผู้ก่อกำเนิด"
        .ListIndex = 0
    End With
    With cboInvState
        .AddItem ""
        .AddItem "ปัญหา PO"
        .AddItem "ปัญหาเรื่องใบเสนอราคา"
        .AddItem "ปัญหาเรื่องน้ำหนัก"
        .AddItem "งานราชการ ลูกค้าไม่ได้ตั้งงบประมาณ"
        .AddItem "ปัญหาภายในของลูกค้า"
        .AddItem "รอรวบรวมเอกสารวางบิล"
        .ListIndex = 0
    End With
    With CbReState
        .AddItem ""
        .AddItem "ปัญหา PO"
        .AddItem "เอกสารสูญหาย"
        .AddItem "งานราชการ ขั้นตอนล่าช้า"
        .AddItem "ปัญหาเรื่องสัญญา"
        .AddItem "ติดปัญหาสภาพคล่อง"
        .AddItem "รออนุมัติจากผู้บริหาร"
        .AddItem "ปัญหาเรื่อง น้ำหนัก ราคา ผิด"
        .AddItem "ลูกค้าต้องการเอกสารเพิ่มเติม"
        .AddItem "จนท รับเช็คไม่ทัน"
        .AddItem "มีกำหนดรับเงินแล้ว"
        .ListIndex = 0
    
    End With
    txtAccStaff.Text = CurrentUser.StaffFName & " " & CurrentUser.StaffLName
    Call SetMSFlexGrid
    'cmdPeriod.ListIndex = 0

    Set curCustomer_ = New BWGCustomer
    Set curCustomerSite_ = New BWGCustomerSite
    PopulateCompany cboCompany, cboIndex
    PopulateRem cboRem
    PopulateUnit
    cboWasteType.ListIndex = 0
    CreateNewMenifest
    SSTab1.Tab = 0
    TxtRefME.Text = Left(txtDocNo.Text, 2) & "BMEHIC"
End Sub

Private Sub PopulateUnit()
On Error Resume Next
Dim x%
Dim tmpS As New BWGSearchManager
    Set colUnit_ = tmpS.UnitGeneralSearch
    cboLiquidUnit.Clear
    cboSolidUnit.Clear
    If colUnit_ Is Nothing Then Exit Sub
    For x = 1 To colUnit_.Count
        cboLiquidUnit.AddItem colUnit_(x).UnitName
        cboSolidUnit.AddItem colUnit_(x).UnitName
        cboActualUnit.AddItem colUnit_(x).UnitName
    Next
    Set tmpS = Nothing
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim selResult As Integer
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole(Me.Name, curRole_)
    If Trim(tmpStr) <> "" Then
        Exit Sub
    End If
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
Set curCustomerSite_ = Nothing
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
'    If Trim(grCustRate.getGridObj.TextMatrix(row, 0)) <> "" Then
'        grCustRate.ColEnabled(col) = col = 4
'    Else
        grCustRate.ColEnabled(col) = False
'    End If
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
    If (col = 4 Or col = 6 Or col = 7 Or col = 10) And IsNumeric(NewValue) = False Then
        MsgBox "กรุณาระบุข้อมูลเป็นตัวเลขเท่านั้น", vbExclamation
        Cancel = True
    End If
End Sub

Private Sub grItem_Click(row As Integer, col As Integer)
    'MsgBox grItem.getGridObj.ColWidth(col)
End Sub

Private Sub grItem_DblClick(row As Integer, col As Integer)
Dim tmpDet As BWGJobDataDetail
    If grItem.ColEnabled(col) = False And Trim(txtCustomer.Text) <> "" And Trim(txtDocNo.Text) <> "" And Trim(grItem.getGridObj.TextMatrix(row, 1)) <> "" Then
        If isExist(myMenifest_.JobDetails, "'" & Trim(grItem.getGridObj.TextMatrix(row, 0)) & "'") Then
            Set tmpDet = myMenifest_.JobDetails("'" & Trim(grItem.getGridObj.TextMatrix(row, 0)) & "'")
            Load frmJobDetFail
            With frmJobDetFail
                .SetJobDetail tmpDet, txtDocNo.Text, txtCustomer.Text
                .Show 1
            End With
            myMenifest_.isChanged = tmpDet.isChanged
            With grItem.getGridObj
                PaintRow grItem.getGridObj, CLng(row), getColorFromError(tmpDet.FailType)
            End With
        Else
            Command4_Click
        End If
    End If
End Sub

Private Sub grItem_RolColChanged(row As Integer, col As Integer)
    If Trim(grItem.getGridObj.TextMatrix(row, 1)) = "" Then
        grItem.ColEnabled(col) = False
    Else
        grItem.ColEnabled(col) = (col = 4 Or col = 5 Or col = 6 Or col = 7 Or col = 9 Or col = 10 Or col = 11)
    End If
                    With grItem.getColObject(9)
                        .Clear
                        .AddItem "New"
                    End With
                    With grItem.getColObject(11)
                        .Clear
                        .AddItem "New"
                    End With

End Sub

Private Sub grItem_UpdateValue(row As Integer, col As Integer)
Dim tmpKey As String, tmpVal As String
Dim tmpDet As Object
    tmpKey = "'" & Trim(grItem.getGridObj.TextMatrix(row, 0)) & "'"
    tmpVal = Trim(grItem.getGridObj.TextMatrix(row, col))
    If isExist(myMenifest_.JobDetails, tmpKey) Then
        Set tmpDet = myMenifest_.JobDetails(tmpKey)
    ElseIf isExist(myMenifest_.AddedWastes, tmpKey) Then
        Set tmpDet = myMenifest_.AddedWastes(tmpKey)
    End If
    If Not tmpDet Is Nothing Then
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
            Case 9
                    tmpDet.IsNewWaste = tmpVal
            Case 10
                    tmpDet.InWeight = tmpVal
            Case 11
                    tmpDet.IsNewWasteCR = tmpVal
        End Select
        If col = 6 Or col = 7 Then
            If tmpDet.QuoTreatmentUnitName = "บาท/เที่ยว" Then
                tmpDet.TreatmentCharge = tmpDet.QuoTreatmentRate
            Else
                If myMenifest_.ChargeWType = "A" Then
                    'tmpDet.TreatmentCharge = tmpDet.QuoTreatmentRate * tmpDet.DisposerWeight
                    If tmpDet.DisposerWeight < tmpDet.QuoMinWeightPerCar Then
                        tmpDet.TreatmentCharge = tmpDet.QuoMinWeightPerCar * tmpDet.QuoTreatmentRate
                    Else
                        tmpDet.TreatmentCharge = tmpDet.DisposerWeight * tmpDet.QuoTreatmentRate
                    End If
                Else
                    'tmpDet.TreatmentCharge = tmpDet.QuoTreatmentRate * tmpDet.EstWasteQty
                    If tmpDet.EstWasteQty < tmpDet.QuoMinWeightPerCar Then
                        tmpDet.TreatmentCharge = tmpDet.QuoMinWeightPerCar * tmpDet.QuoTreatmentRate
                    Else
                        tmpDet.TreatmentCharge = tmpDet.EstWasteQty * tmpDet.QuoTreatmentRate
                    End If
                End If
            End If
            grCustRate.getGridObj.TextMatrix(row, 4) = tmpDet.TreatmentCharge
            SumTreatmentCharge
            SumTotalWeight
        End If
        myMenifest_.isChanged = True
        Set tmpDet = Nothing
    End If
End Sub

Private Sub lvMenifest_DblClick()
Dim tmpCol As Collection, tmpS As New BWGSearchManager
    If lvMenifest.SelectedItem Is Nothing Then Exit Sub
    Set tmpCol = tmpS.MenifestSearch("MenifestID=" & lvMenifest.SelectedItem.Key)
    If tmpCol.Count > 0 Then
        cmdClear_Click
        If Trim(txtCustomer.Text) = "" Then
            SetDataMenifest tmpCol(1)
        End If
    End If
    Set tmpCol = Nothing
    Set tmpS = Nothing
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
Private Sub Option1_Click()
    sqlStr = "UPDATE tbJobDataTimeTable SET SendDesc='" & Option1.Caption & "' " & vbCrLf
    sqlStr = sqlStr & "WHERE JobDataCarID = '" & myMenifest_.JobDataCarID & "' AND TripNO = '" & myMenifest_.TripNo & "'"
    DBConnExc sqlStr
End Sub

Private Sub Option2_Click()
    sqlStr = "UPDATE tbJobDataTimeTable SET SendDesc='" & Option2.Caption & "' " & vbCrLf
    sqlStr = sqlStr & "WHERE JobDataCarID = '" & myMenifest_.JobDataCarID & "' AND TripNO = '" & myMenifest_.TripNo & "'"
    DBConnExc sqlStr
End Sub

Private Sub Option3_Click()
    sqlStr = "UPDATE tbJobDataTimeTable SET SendDesc='" & Option3.Caption & "' " & vbCrLf
    sqlStr = sqlStr & "WHERE JobDataCarID = '" & myMenifest_.JobDataCarID & "' AND TripNO = '" & myMenifest_.TripNo & "'"
    DBConnExc sqlStr
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
            cboSiteName.AddItem .SiteAddresses(x).Address & " ต." & .SiteAddresses(x).TumbolName & " อ." & .SiteAddresses(x).AumphurName & " จ." & .SiteAddresses(x).ProvinceName & " " & .SiteAddresses(x).ZipCode
        Next
        sqlStr = "Select WasteManageID from tbCustomer WHERE CustomerID='" & .id & "' Group By WasteManageID"
        Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
        If tmpRec Is Nothing Then Exit Sub
        cboGenerator.Clear
        cboGenerator.AddItem "ไม่ระบุเลขผู้กำเนิดของเสีย"
        Do Until tmpRec.EOF
            If Trim("" & tmpRec("WasteManageID")) <> "" Then cboGenerator.AddItem Trim("" & tmpRec("WasteManageID"))
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
        Text1.Text = .Digit14
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
Public Sub SetDataCustomerSite(selSite As BWGCustomerSite)
Dim tmpCol As Collection
Dim tmpS As New BWGSearchManager
    Set curCustomerSite_ = Nothing
    Set curCustomerSite_ = selSite
    If curCustomerSite_ Is Nothing Then Exit Sub
    With curCustomerSite_
        txtLengthNum.Text = .LengthNum
    End With
End Sub
Private Sub CreateNewMenifest()
    Set myMenifest_ = New BWGMenifest
    curRole_ = AddNewRole
    With myMenifest_
        .id = GetGUID
        .OpenDate = TodayDate
        .WorkDate = dtWorkDate.ValueYMD
        .NonOrHz = cboWasteType.Text
        .OpenStaffID = CurrentUser.id
        If cboWType.ListIndex = 0 Then
            .ChargeWType = "A"
        ElseIf cboWType.ListIndex = 1 Then
            .ChargeWType = "B"
        End If
        .isDocCompleted = cboComplete.ListIndex = 1
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
    txtTR_Remark.Text = ""
End Sub



Private Sub ClearScreen()
    txtDocNo.Text = "== AUTO =="
    txtMenifestNo.Text = "== AUTO =="
    txtCustomer.Text = ""
    cboWType.ListIndex = 0
    SetCancelScreen
    Set curCustomer_ = Nothing
    Set curCustomer_ = New BWGCustomer
    cboGenerator.ListIndex = -1
    dtWorkDate.ValueYMD = TodayDate
    cboSiteName.Clear
    txtCustTel.Text = ""
    txtCustFax.Text = ""
    txtCustER.Text = ""
    cboComplete.ListIndex = 0
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
    TxtRefME.Text = Left(txtDocNo.Text, 2) & "BMEHIC"
    txtTrans1B.Text = ""
    txtTransID1B.Text = ""
    txtTransTel.Text = ""
    txtTransFax.Text = ""
    txtTransER.Text = ""
    txtTruckType.Text = ""
    cboCarRegisID.Clear
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
    
    txtControlDocNo.Text = ""
    txtTransFee.Text = ""
    txtMoreWorker.Text = ""
    txtWorkerFee.Text = ""
    txtEquipFee.Text = ""
    txtOtherDesc2.Text = ""
    txtOtherFee2.Text = ""
End Sub

Public Sub SetDataMenifest(selMenifest As BWGMenifest)
'On Error GoTo ErrD
Dim tmpS As New BWGSearchManager
Dim tmpCol As Collection
Dim tmpColsite As Collection
Dim x%, tmpGrid As MSFlexGrid
Dim tmpTime As BWGJobDataTimeTable
Dim tmpChargeRate As BWGTransVendorChargeRate
Dim tmpCustSite As BWGCustomerSite
Dim tmpList As ListItem, tmpMf As BWGMenifest
Dim UserGrp As Collection, Group As New AWSUserGroup, i As Integer
    Set myMenifest_ = selMenifest
    curRole_ = UpdateRole
    With myMenifest_
        SetCancelScreen
        lblLastUpdate.Caption = "อัพเดทล่าสุดโดย : " & .LastUser & " " & Format(.LastUpdate, "dd/MM/yyyy  HH:mm:ss")
        lblLastApprove.Caption = "อนุมัติโดย : " & .ApproveStaffName & " " & Format(.ApproveDate, "dd/MM/yyyy") & "  " & Format(.ApproveTime, "HH:mm:ss")
        txtDocNo.Text = .DocNo
        If Trim(.RefME) <> "" Then
            TxtRefME.Text = .RefME
        Else
                TxtRefME.Text = Left(txtDocNo.Text, 2) & "BMEHIC"
        End If
        txtMenifestNo.Text = .MenifestNo
        If .NonOrHz = "HIC" Or .NonOrHz = "HBI" Then
            Label2(88).Visible = True
            txtAPNo.Visible = True
            txtAPNo.Text = .APManifestNo
        Else
            Label2(88).Visible = False
            txtAPNo.Visible = False
            txtAPNo.Text = ""
        End If
        txtJobNo.Text = .JobNo
        'ค้นหาข้อมูลลูกค้า ไม่จำเป็นต้องค้นหาในหน้านี้
        Set tmpCol = tmpS.CustomerSearch("CustomerID='" & .CustomerID & "'")
        If tmpCol.Count > 0 Then SetDataCustomer tmpCol(1), False
        SetTextToCombo .CompanyName, cboCompany
        If Trim(.WasteGenNo) = "" Then
            cboGenerator.ListIndex = 0
        Else
            SetTextToCombo .WasteGenNo, cboGenerator
        End If
        Set tmpColsite = tmpS.CustomerSiteSearch("CustSiteID='" & .CustSiteID & "'")
        SetDataCustomerSite tmpColsite(1)
        SetTextToCombo .CustSiteAddr, cboSiteName
        
        If .ChargeWType = "A" Then
            cboWType.ListIndex = 0
        ElseIf .ChargeWType = "B" Then
            cboWType.ListIndex = 1
        End If
        dtWorkDate.ValueYMD = .WorkDate
        If cboSiteName.ListCount = 1 Then
            If cboSiteName.ListIndex < 0 Then cboSiteName.ListIndex = 0
        End If
        If cboSiteName.ListCount > 0 Then
            If cboSiteName.ListIndex = -1 Then cboSiteName.ListIndex = 0
        End If
        txtCustTel.Text = .TelNo
        txtCustFax.Text = .FaxNo
        txtCustER.Text = .ErContact
        If .isDocCompleted Then
            cboComplete.ListIndex = 1
        Else
            cboComplete.ListIndex = 0
        End If
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
        txtRem.Text = .RemComment
        cboRem.Text = .RemID
        If .GPSSend = "1" Then
            Check1.Value = 1
        Else
            Check1.Value = 0
        End If
        If .SendDesc = "แนบ ฉ.6" Then
            Option1.Value = True
            ElseIf .SendDesc = "แนบวางบิล" Then
            Option2.Value = True
            ElseIf .SendDesc = "E-Mail" Then
            Option3.Value = True
        End If
        
        '------------------------------------------------------------------------------------------------
        txtSolidQty.Text = .TotalSolid
        SetTextToCombo .SolidUnitName, cboSolidUnit
        txtTreatment.Text = .SpecialComment
        SetTextToCombo .InvState, cboInvState
        SetTextToCombo .REState, CbReState
        txtInvDetail.Text = .InvDetail
        
        Dim TRStaff As String
        Set UserGrp = CurrentUser.getUserGroups
        For i = 1 To UserGrp.Count
            Set Group = UserGrp(i)
            If InStr(1, Group.Name, "Transport") > 0 Then
                TRStaff = Trim(CurrentUser.StaffFName & " " & CurrentUser.StaffLName)
                Exit For
            End If
        Next
        
        If .isClosed = True Then
            txtSigneName1.Text = .SignName1
            txtDestName.Text = .SignName3
        Else
            If .SignName1 = "" Then
                txtSigneName1.Text = TRStaff
                txtSigneName1_Validate False
                myMenifest_.isChanged = False
            ElseIf Trim(.SignName1) <> Trim(CurrentUser.StaffFName & " " & CurrentUser.StaffLName) Then
                txtSigneName1.Text = TRStaff
                txtSigneName1_Validate False
                myMenifest_.isChanged = False
            Else
                txtSigneName1.Text = .SignName1
            End If
            
            If Trim(txtSigneName1.Text) <> "" Then
                txtDestName.Text = Trim(txtSigneName1.Text)
                txtDestName_Validate False
                myMenifest_.isChanged = False
            Else
                txtDestName.Text = .SignName3
            End If
            
        End If
        dtSign1.ValueYMD = .SignDate1
        dtSign3.ValueYMD = .SignDate3
        
        txtTrans1B.Text = .TransportName2
        txtTransID1B.Text = .TransRegisNo2
        txtDriverName.Text = .DriverName
        txtTransTel.Text = .TransTelNo1
        txtTransFax.Text = .TransFaxNo1
        txtTransER.Text = .TransErContact1
        
        Set tmpCol = tmpS.TruckSubTypeSearch("SubTypeID='" & .TruckSubTypeID & "'")
        If tmpCol.Count > 0 Then SetDataTruckSubType tmpCol(1), False
        'txtTruckType.Text = .TruckSubTypeName & ", " & .TruckTypeName
        SetTextToCombo .CarRegisID, cboCarRegisID
        
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
        txtRealWeight.Text = Format(.ActualQty, "#0.000")
        txtTreatTime.Text = .TreatmentPeriod
        If .PeriodUnit = "D" Then
            cboTimeUnit.ListIndex = 0
        ElseIf .PeriodUnit = "M" Then
            cboTimeUnit.ListIndex = 1
        ElseIf .PeriodUnit = "Y" Then
            cboTimeUnit.ListIndex = 2
        End If
        
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
        ShowWasteDetail
        
        If .InCollectFee > 0 Then
            txtInCollect.Text = .InCollectFee
            'txtIntrans.Text = .IntransFee
            'txtInWorker.Text = .InWorkerFee
            'txtInEquip.Text = .InEquipFee
            txtOtherDesc.Text = .InOtherDesc
            txtOtherFee.Text = .InOtherFee
        Else
            'txtIntrans.Text = maxTrans
            '.IntransFee = maxTrans
        End If
        txtWorkerCount.Text = .OutWorkerCount
        If .OutTransFee > 0 Then
            txtOutTrans.Text = .OutTransFee
            txtOutEquip.Text = .OutEquipFee
            txtOutOther.Text = .OutOtherDesc
            txtOutOtherFee.Text = .OutOtherFee
        Else
            If cboSiteName.ListIndex > -1 Then
                Set tmpChargeRate = CalTransFeeFromDistance(curCustomer_.SiteAddresses(cboSiteName.ListIndex + 1).AumphurID, .TSDFAumphurID, .TruckTypeID, .WorkDate)
            End If
            If Not tmpChargeRate Is Nothing Then
                txtOutTrans.Text = tmpChargeRate.RatePerTrip
            Else
                txtOutTrans.Text = ""
            End If
        End If
        txtOutWorker.Text = .OutWorkerFee
        
        txtControlDocNo.Text = .TimeTableNo
        'Set tmpCol = tmpS.JobDataTimeTableSearch("JobDataCarID='" & .JobDataCarID & "' And TripNo=" & .tripNo)
        Set tmpCol = tmpS.JobDataTimeTableQuickSearch("TimeTableNo='" & .TimeTableNo & "' And isMain='Y' And isCancel ='N' AND ISNULL(isCanceled, 'N') <> 'Y' AND TripNo = " & myMenifest_.TripNo)
        If tmpCol.Count = 0 Then
            DBConnExc "Update tbJobDataTimeTable Set isMain='Y' WHERE TimeTableNo='" & .TimeTableNo & "' And TripNo=1"
            Set tmpCol = tmpS.JobDataTimeTableQuickSearch("TimeTableNo='" & .TimeTableNo & "' And isMain='Y' And isCancel ='N' AND ISNULL(isCanceled, 'N') <> 'Y' AND TripNo = " & myMenifest_.TripNo)
        End If
        If tmpCol.Count > 0 Then
            Set tmpTime = tmpCol(1)
            txtTransFee.Text = tmpTime.TripTranFee
            txtMoreWorker.Text = tmpTime.TripMoreWorker
            txtWorkerFee.Text = tmpTime.TripMoreWorkerFee
            txtEquipFee.Text = tmpTime.TripEquipFee
            txtOtherDesc2.Text = tmpTime.TripOtherDesc
            txtOtherFee2.Text = tmpTime.TripOtherFee
            'Set tmpCol = tmpS.MenifestSearch("JobDataCarID='" & .JobDataCarID & "' And TripNo=" & .tripNo)
            Set tmpCol = tmpS.MenifestSearch("TimeTableNo='" & .TimeTableNo & "'")
            lvMenifest.ListItems.Clear
            For x = 1 To tmpCol.Count
                Set tmpMf = tmpCol(x)
                Set tmpList = lvMenifest.ListItems.Add(, "'" & tmpMf.id & "'", tmpMf.DocNo)
                tmpList.SubItems(1) = Format(tmpMf.InCollectFee, "#,##0.00")
                tmpList.SubItems(2) = Format(tmpMf.InOtherFee, "#,##0.00")
                If tmpMf.isClosed = False Then
                    tmpList.ForeColor = vbRed
                    tmpList.ListSubItems(1).ForeColor = vbRed
                    tmpList.ListSubItems(2).ForeColor = vbRed
                End If
                If tmpMf.id = myMenifest_.id Then
                    tmpList.Bold = True
                    tmpList.ListSubItems(1).Bold = True
                    tmpList.ListSubItems(2).Bold = True
                End If
                Set tmpList = Nothing
                Set tmpMf = Nothing
            Next
        End If
        .isChanged = False
        If .isClosed = False Then
            cboInvState.Enabled = False
            CbReState.Enabled = False
            txtInvDetail.Enabled = False
            TxtRefME.Enabled = True
        Else
            Label3.Visible = True
            cboInvState.Enabled = True
            CbReState.Enabled = True
            txtInvDetail.Enabled = True
            TxtRefME.Enabled = True
        End If
    End With
    Set tmpCol = Nothing
    Set tmpS = Nothing
    SumTreatmentCharge
    ShowDiscrepWasteData
    ShowTR_Remark
    TxtRefME.Enabled = True
    Exit Sub
ErrD:
    MsgBox err.Description & vbCrLf & "เกิดข้อผิดพลาดกรุณาลองใหม่อีกครั้ง", vbCritical, "Error"
    Exit Sub
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



Private Sub txtDestroyPlace_Validate(Cancel As Boolean)
    myMenifest_.TSDFAddr = txtDestroyPlace.Text
End Sub

Private Sub txtDestTel_Validate(Cancel As Boolean)
    myMenifest_.TSDFTelNo = txtDestTel.Text
End Sub

Private Sub txtEquipFee_Change()
    SumCustCharge
End Sub
Private Sub txtInvDetail_Validate(Cancel As Boolean)
'    myMenifest_.InvDetail = txtInvDetail.Text
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
'    If IsNumeric(Trim(txtInCollect.Text)) Then
'        myMenifest_.InCollectFee = Trim(txtInCollect.Text)
'    Else
'        myMenifest_.InCollectFee = 0
'    End If
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
Private Sub txtLengthNum_Validate(Cancel As Boolean)
    curCustomerSite_.LengthNum = Trim(txtLengthNum.Text)
    curCustomerSite_.isChanged = True
    myMenifest_.IsCanceled = True
End Sub

Private Sub txtLiquidQty_Validate(Cancel As Boolean)
    If IsNumeric(txtLiquidQty.Text) Then
        myMenifest_.TotalLiquid = txtLiquidQty.Text
    Else
        myMenifest_.TotalLiquid = 0
    End If
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

Private Sub txtRate_Validate(Cancel As Boolean)
    If IsNumeric(txtRate.Text) Then
        myMenifest_.CarRate = txtRate.Text
    Else
        myMenifest_.CarRate = 0
    End If
End Sub

Private Sub txtRealWeight_Validate(Cancel As Boolean)
    If IsNumeric(txtRealWeight.Text) Then
        myMenifest_.ActualQty = txtRealWeight.Text
    Else
        myMenifest_.ActualQty = 0
    End If
End Sub
Private Sub TxtRefME_Validate(Cancel As Boolean)
    If TxtRefME.Text <> Left(txtDocNo.Text, 2) & "BMEHIC" Then
        myMenifest_.RefME = Trim(TxtRefME.Text)
        sqlStr = "UPDATE tbMenifestHeader SET  RefME = '" & Trim(TxtRefME.Text) & "'" & vbCrLf
        sqlStr = sqlStr & "WHERE MenifestID = '" & myMenifest_.id & "' "
        DBConnExc sqlStr
    Else
        myMenifest_.RefME = ""
    End If

End Sub

Private Sub txtRem_Validate(Cancel As Boolean)
    myMenifest_.RemComment = txtRem.Text
    myMenifest_.isChanged = True
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

Private Sub txtTransFee_Change()
    SumCustCharge
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
            .TruckTypeID = ""
            .TruckTypeName = ""
        End With
    Else
        If Trim(curTruckSubType_.SubTypeDesc & ", " & curTruckSubType_.TypeDesc) <> Trim(txtTruckType.Text) Then
            Set curTruckSubType_ = Nothing
            Set curTruckSubType_ = New BWGTruckSubType
            With myMenifest_
                .TruckSubTypeID = ""
                .TruckSubTypeName = ""
                .TruckTypeID = ""
                .TruckTypeName = ""
            End With
            Load frmCarTypeSearch
            With frmCarTypeSearch
                .getCriteria (txtTruckType.Text)
                If curTrans2_ Is Nothing Then
                    .SetCallerForm Me
                Else
                    .SetCallerForm Me, , curTrans2_.id
                End If
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
On Error Resume Next
Dim x%, maxTripRate As Double
Dim tmpAmt As Double
Dim aa As BWGJobDataDetail
    For x = 1 To myMenifest_.JobDetails.Count
        If myMenifest_.JobDetails(x).QuoTreatmentUnitID = "บาท/เที่ยว" Then
            If maxTripRate < myMenifest_.JobDetails(x).TreatmentCharge Then maxTripRate = myMenifest_.JobDetails(x).TreatmentCharge
        Else
            tmpAmt = tmpAmt + myMenifest_.JobDetails(x).TreatmentCharge
        End If
    Next
    For x = 1 To myMenifest_.AddedWastes.Count
        If myMenifest_.AddedWastes(x).TreatmentUnitID = "บาท/เที่ยว" Then
            If maxTripRate < myMenifest_.AddedWastes(x).TreatmentCharge Then maxTripRate = myMenifest_.AddedWastes(x).TreatmentCharge
        Else
            tmpAmt = tmpAmt + myMenifest_.AddedWastes(x).TreatmentCharge
        End If
    Next
    txtInCollect.Text = Format(tmpAmt + maxTripRate, "#,##0.00")
    'myMenifest_.InCollectFee = txtInCollect.Text
End Sub

Public Sub SetDataTransportVender(selVendor As BWGTransportVendor, Optional isAssigned As Boolean = True)
Dim x%, hasCar As Boolean
    Set curTrans2_ = selVendor
    With curTrans2_
        txtTrans2.Text = .NameTH
        txtTransID2.Text = .TransLicenseNo
        txtTrans1B = .NameTH
        txtTransID1B = .TransLicenseNo
        If isAssigned Then
            myMenifest_.Transport2ID = .id
            myMenifest_.TransportName2 = .NameTH
            myMenifest_.TransRegisNo2 = .TransLicenseNo
            If Not curTruckSubType_ Is Nothing Then
                hasCar = False
                For x = 1 To .VenderCar.Count
                    If .VenderCar(x).SubTypeID = curTruckSubType_.id Then
                        hasCar = True
                        Exit For
                    End If
                Next
                Set curTruckSubType_ = Nothing
                txtTruckType.Text = ""
                myMenifest_.TruckSubTypeID = ""
                myMenifest_.TruckSubTypeName = ""
                myMenifest_.TruckTypeID = ""
                myMenifest_.TruckTypeName = ""
            End If
        End If
        ShowCarRegisID
    End With
End Sub

Public Sub SetCancelScreen()
Dim cFlag As Boolean, AFlag As Boolean
TxtRefME.Enabled = True
    cFlag = myMenifest_.IsCanceled
    AFlag = myMenifest_.isApproved
    If myMenifest_.IsCanceled Then
        cmdSave.Enabled = Not cFlag
        cmdCancel.Enabled = Not cFlag
        Command4.Enabled = Not AFlag
        lbCancel.Visible = cFlag
        lbCancel.Caption = "CANCELED"
        txtCancelDetail.Visible = cFlag
        txtCancelDetail.Text = myMenifest_.CancelDetail
    ElseIf myMenifest_.isClosed Then
        cFlag = myMenifest_.isClosed
        cmdSave.Enabled = Not cFlag
        cmdCancel.Enabled = Not cFlag
        lbCancel.Visible = cFlag
        lbCancel.Caption = "CLOSED"
        txtCancelDetail.Visible = Not cFlag
    Else
        cmdSave.Enabled = Not cFlag
        cmdCancel.Enabled = Not cFlag
        lbCancel.Visible = cFlag
        txtCancelDetail.Visible = cFlag
    End If
    
    lbApprove.Visible = AFlag
    lblLastApprove.Visible = AFlag
    If myMenifest_.isApproved Then
        cmdSave.Enabled = Not AFlag
        cmdCancel.Enabled = Not AFlag
        cmdMenifestClosed.Enabled = True
        Command4.Enabled = Not AFlag
        lbApprove.Caption = "APPROVED"
    End If
End Sub

Private Sub ShowDiscrepWasteData()
On Error Resume Next
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

Public Sub SetDataTruckSubType(selSubType As BWGTruckSubType, Optional isAssigned As Boolean = True)
On Error Resume Next
    Set curTruckSubType_ = Nothing
    Set curTruckSubType_ = selSubType
    With curTruckSubType_
        If Not myMenifest_.getTripControl Is Nothing Then
            If myMenifest_.getTripControl.isTrailer = True Then
                txtTruckType.Text = .SubTypeDesc & " (พ่วง), " & .TypeDesc
            Else
                txtTruckType.Text = .SubTypeDesc & ", " & .TypeDesc
            End If
        End If
        'If isAssigned Then
            myMenifest_.TruckSubTypeID = .id
            myMenifest_.TruckSubTypeName = .SubTypeDesc
            myMenifest_.TruckTypeID = .TruckTypeID
            myMenifest_.TruckTypeName = .TypeDesc
        'End If
        ShowCarRegisID
    End With
End Sub

Public Sub SetDataProvince(selProvince As IMIProvince, Optional isAssigned As Boolean = True)
On Error Resume Next
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
On Error Resume Next
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

Private Sub ShowCarRegisID()
On Error Resume Next
Dim tmpCar As BWGVenderCar
Dim x%, SubTypeID As String
    If curTruckSubType_ Is Nothing Then
        SubTypeID = ""
    Else
        SubTypeID = curTruckSubType_.id
    End If
    cboCarRegisID.Clear
    If Not curTrans2_ Is Nothing Then
        For x = 1 To curTrans2_.VenderCar.Count
            Set tmpCar = curTrans2_.VenderCar(x)
            If Trim(SubTypeID) = "" Or Trim(tmpCar.SubTypeID) = Trim(SubTypeID) Then
                cboCarRegisID.AddItem tmpCar.CarRegisID
            End If
        Next
    End If
End Sub

Private Sub SumTotalWeight()
On Error Resume Next
Dim x%, useCol%
Dim tmpAmt As Double
    If myMenifest_.ChargeWType = "A" Then
        useCol = 7
    Else
        useCol = 6
    End If
    With grItem.getGridObj
        For x = 1 To .Rows - 1
            If IsNumeric(Trim(.TextMatrix(x, useCol))) Then
                tmpAmt = tmpAmt + CDbl(Trim(.TextMatrix(x, useCol)))
            End If
        Next
    End With
    txtRealWeight.Text = tmpAmt
    myMenifest_.ActualQty = tmpAmt
End Sub

Private Sub PaintRow(selGrid As MSFlexGrid, selRow As Long, selColor As Long)
Dim x As Long
Dim OldCol As Long
    OldCol = selGrid.col
    For x = 1 To selGrid.Cols - 1
        selGrid.col = x
        selGrid.CellBackColor = selColor
    Next
    selGrid.col = OldCol
End Sub

Private Sub SumCustCharge()
On Error Resume Next
Dim amt1 As Double, amt2 As Double, amt3 As Double, amt4 As Double
    If IsNumeric(txtTransFee) Then
        amt1 = CDbl(txtTransFee.Text)
    Else
        amt1 = 0
    End If
    
    If IsNumeric(txtWorkerFee.Text) Then
        amt2 = txtWorkerFee.Text
    Else
        amt2 = 0
    End If
    
    If IsNumeric(txtEquipFee.Text) Then
        amt3 = txtEquipFee.Text
    Else
        amt3 = 0
    End If
    
    If IsNumeric(txtOtherFee2.Text) Then
        amt4 = txtOtherFee2.Text
    Else
        amt4 = 0
    End If
    
    txtSumTotal.Text = Format(amt1 + amt2 + amt3 + amt4, "#,##0.00")
End Sub

Private Sub txtWorkerFee_Change()
    SumCustCharge
End Sub

Private Sub ShowWasteDetail()
On Error Resume Next
Dim tmpGrid As MSFlexGrid, tmpJobDet As BWGJobDataDetail
Dim x%, lastAmt%, tmpAddWaste As BWGMenifestWasteAdded, maxTrans As Double
    With myMenifest_
        grItem.ClearAllData
        Set tmpGrid = grItem.getGridObj
        tmpGrid.Rows = .JobDetails.Count + .AddedWastes.Count + 20
        lastAmt = .JobDetails.Count
        For x = 1 To lastAmt
            Set tmpJobDet = .JobDetails(x)
            tmpGrid.TextMatrix(x, 0) = tmpJobDet.id
            tmpGrid.col = 1
            tmpGrid.row = x
            PaintRow tmpGrid, CLng(x), getColorFromError(tmpJobDet.FailType)
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
            ElseIf Trim(tmpJobDet.QuoTreatmentUnitName) = "บาท/กล่อง" Then
                tmpGrid.TextMatrix(x, 8) = "กล่อง"
            ElseIf Trim(tmpJobDet.QuoTreatmentUnitName) = "บาท/ตัน" Then
                tmpGrid.TextMatrix(x, 8) = "ตัน"
            ElseIf Trim(tmpJobDet.QuoTreatmentUnitName) = "บาท/งาน" Then
                tmpGrid.TextMatrix(x, 8) = "งาน"
            ElseIf Trim(tmpJobDet.QuoTreatmentUnitName) = "บาท/กิโลกรัม" Then
                tmpGrid.TextMatrix(x, 8) = "กิโลกรัม"
            ElseIf Trim(tmpJobDet.QuoTreatmentUnitName) = "บาท/ถุง" Then
                tmpGrid.TextMatrix(x, 8) = "ถุง"
            ElseIf Trim(tmpJobDet.QuoTreatmentUnitName) = "บาท/ขวด" Then
                tmpGrid.TextMatrix(x, 8) = "ขวด"
             ElseIf Trim(tmpJobDet.QuoTreatmentUnitName) = "บาท/แล็ก" Then
                tmpGrid.TextMatrix(x, 8) = "แล็ก"
             ElseIf Trim(tmpJobDet.QuoTreatmentUnitName) = "บาท/กระป๋อง" Then
                tmpGrid.TextMatrix(x, 8) = "กระป๋อง"
             ElseIf Trim(tmpJobDet.QuoTreatmentUnitName) = "บาท/หลอด" Then
                tmpGrid.TextMatrix(x, 8) = "หลอด"
             ElseIf Trim(tmpJobDet.QuoTreatmentUnitName) = "บาท/แกลอน" Then
                tmpGrid.TextMatrix(x, 8) = "แกลอน"
             ElseIf Trim(tmpJobDet.QuoTreatmentUnitName) = "บาท/วัน" Then
                tmpGrid.TextMatrix(x, 8) = "วัน"
             ElseIf Trim(tmpJobDet.QuoTreatmentUnitName) = "บาท/เดือน" Then
                tmpGrid.TextMatrix(x, 8) = "เดือน"
             ElseIf Trim(tmpJobDet.QuoTreatmentUnitName) = "บาท/ปี" Then
                tmpGrid.TextMatrix(x, 8) = "ปี"
             ElseIf Trim(tmpJobDet.QuoTreatmentUnitName) = "บาท/ลูก" Then
                tmpGrid.TextMatrix(x, 8) = "ลูก"
             ElseIf Trim(tmpJobDet.QuoTreatmentUnitName) = "บาท/Job" Then
                tmpGrid.TextMatrix(x, 8) = "Job"
             ElseIf Trim(tmpJobDet.QuoTreatmentUnitName) = "บาท/เที่ยว" Then
                If tmpJobDet.WasteType = "DOC" Then
                    tmpGrid.TextMatrix(x, 8) = "เที่ยว"
                Else
                    tmpGrid.TextMatrix(x, 8) = "ตัน"
                End If
                
           Else
                tmpGrid.TextMatrix(x, 8) = "ตัน"
            End If
                tmpGrid.TextMatrix(x, 9) = tmpJobDet.IsNewWaste
                tmpGrid.TextMatrix(x, 10) = tmpJobDet.InWeight
                 tmpGrid.TextMatrix(x, 11) = tmpJobDet.IsNewWasteCR
            Set tmpJobDet = Nothing
        Next
        For x = 1 To .AddedWastes.Count
            Set tmpAddWaste = .AddedWastes(x)
            tmpGrid.TextMatrix(lastAmt + x, 0) = tmpAddWaste.id
            tmpGrid.col = 1
            tmpGrid.row = lastAmt + x
            PaintRow tmpGrid, CLng(x), &HFFC0C0
            tmpGrid.TextMatrix(lastAmt + x, 1) = lastAmt + x
            tmpGrid.TextMatrix(lastAmt + x, 2) = tmpAddWaste.WasteName & " [" & tmpAddWaste.QuotationNo & "]"
            tmpGrid.TextMatrix(lastAmt + x, 3) = tmpAddWaste.WasteEUCode
            tmpGrid.TextMatrix(lastAmt + x, 4) = tmpAddWaste.ContainerCount
            tmpGrid.TextMatrix(lastAmt + x, 5) = tmpAddWaste.ContainerTypeDesc
            tmpGrid.TextMatrix(lastAmt + x, 6) = tmpAddWaste.EstWasteQty
            tmpGrid.TextMatrix(lastAmt + x, 7) = tmpAddWaste.DisposerWeight
            If Trim(tmpAddWaste.TreatmentUnitID) = "บาท/ถัง" Then
                tmpGrid.TextMatrix(lastAmt + x, 8) = "ถัง"
            ElseIf Trim(tmpAddWaste.TreatmentUnitID) = "บาท/ลูกบาศก์เมตร" Then
                tmpGrid.TextMatrix(lastAmt + x, 8) = "ลูกบาศก์เมตร"
            ElseIf Trim(tmpAddWaste.TreatmentUnitID) = "บาท/กล่อง" Then
                tmpGrid.TextMatrix(lastAmt + x, 8) = "กล่อง"
            ElseIf Trim(tmpAddWaste.TreatmentUnitID) = "บาท/ตัน" Then
                tmpGrid.TextMatrix(lastAmt + x, 8) = "ตัน"
            ElseIf Trim(tmpAddWaste.TreatmentUnitID) = "บาท/งาน" Then
                tmpGrid.TextMatrix(lastAmt + x, 8) = "งาน"
            ElseIf Trim(tmpAddWaste.TreatmentUnitID) = "บาท/กิโลกรัม" Then
                tmpGrid.TextMatrix(lastAmt + x, 8) = "กิโลกรัม"
            ElseIf Trim(tmpAddWaste.TreatmentUnitID) = "บาท/ถุง" Then
                tmpGrid.TextMatrix(lastAmt + x, 8) = "ถุง"
            ElseIf Trim(tmpAddWaste.TreatmentUnitID) = "บาท/ขวด" Then
                tmpGrid.TextMatrix(lastAmt + x, 8) = "ขวด"
            ElseIf Trim(tmpAddWaste.TreatmentUnitID) = "บาท/แล็ก" Then
                tmpGrid.TextMatrix(lastAmt + x, 8) = "แล็ก"
            ElseIf Trim(tmpAddWaste.TreatmentUnitID) = "บาท/กระป๋อง" Then
                tmpGrid.TextMatrix(lastAmt + x, 8) = "กระป๋อง"
            ElseIf Trim(tmpAddWaste.TreatmentUnitID) = "บาท/หลอด" Then
                tmpGrid.TextMatrix(lastAmt + x, 8) = "หลอด"
            ElseIf Trim(tmpAddWaste.TreatmentUnitID) = "บาท/แกลอน" Then
                tmpGrid.TextMatrix(lastAmt + x, 8) = "แกลอน"
             ElseIf Trim(tmpAddWaste.TreatmentUnitID) = "บาท/วัน" Then
                tmpGrid.TextMatrix(lastAmt + x, 8) = "วัน"
             ElseIf Trim(tmpAddWaste.TreatmentUnitID) = "บาท/เดือน" Then
                tmpGrid.TextMatrix(lastAmt + x, 8) = "เดือน"
             ElseIf Trim(tmpAddWaste.TreatmentUnitID) = "บาท/ปี" Then
                tmpGrid.TextMatrix(lastAmt + x, 8) = "ปี"
             ElseIf Trim(tmpAddWaste.TreatmentUnitID) = "บาท/ลูก" Then
                tmpGrid.TextMatrix(lastAmt + x, 8) = "ลูก"
             ElseIf Trim(tmpAddWaste.TreatmentUnitID) = "บาท/Job" Then
                tmpGrid.TextMatrix(lastAmt + x, 8) = "Job"
           Else
                tmpGrid.TextMatrix(lastAmt + x, 8) = "ตัน"
            End If
                tmpGrid.TextMatrix(lastAmt + x, 9) = tmpAddWaste.IsNewWaste
                tmpGrid.TextMatrix(lastAmt + x, 10) = tmpAddWaste.InWeight
            Set tmpAddWaste = Nothing
        Next
        
        Set tmpGrid = grCustRate.getGridObj
        grCustRate.ClearAllData
        tmpGrid.Rows = .JobDetails.Count + 20
        For x = 1 To lastAmt
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
        For x = 1 To .AddedWastes.Count
            Set tmpAddWaste = .AddedWastes(x)
            tmpGrid.TextMatrix(lastAmt + x, 0) = tmpAddWaste.id
            tmpGrid.TextMatrix(lastAmt + x, 1) = tmpAddWaste.WasteName & " [" & tmpAddWaste.QuotationNo & "]"
            tmpGrid.TextMatrix(lastAmt + x, 2) = Format(tmpAddWaste.QuoTreatmentRate, "#,##0.00") & " " & tmpAddWaste.TreatmentUnitID
            If tmpAddWaste.isPriceIncTrans Then
                tmpGrid.TextMatrix(lastAmt + x, 3) = "รวมค่าขนส่ง"
            Else
                tmpGrid.TextMatrix(lastAmt + x, 3) = Format(tmpAddWaste.TransportFee, "#,##0.00") & " " & tmpAddWaste.TransferUnitID
                If maxTrans < tmpAddWaste.TransportFee Then maxTrans = tmpAddWaste.TransportFee
            End If
            If tmpAddWaste.TreatmentUnitID = "บาท/เที่ยว" And tmpAddWaste.TreatmentCharge = 0 Then
                tmpAddWaste.TreatmentCharge = tmpAddWaste.QuoTreatmentRate
            End If
            tmpGrid.TextMatrix(lastAmt + x, 4) = tmpAddWaste.TreatmentCharge
            Set tmpAddWaste = Nothing
        Next
    End With
End Sub

Public Sub AddWasteDataItem(selItem As BWGWasteDataCR)
Dim tmpCol As Collection
Dim x%, isNew As Boolean
Dim tmpID As String
Dim tmpDetail As New BWGMenifestWasteAdded
    Set tmpCol = myMenifest_.AddedWastes

    If tmpCol Is Nothing Then Set tmpCol = New Collection
    isNew = True
    tmpID = Trim(selItem.id)
    For x = 1 To tmpCol.Count
        'If tmpCol(x).selitem.ID = tmpID Then
        If tmpCol(x).wastedataID = tmpID Then
            Set tmpDetail = tmpCol(x)
            isNew = False
            Exit For
        End If
    Next

    If isNew Then
        tmpDetail.id = GetGUID
        tmpDetail.menifestID = myMenifest_.id
        tmpDetail.wastedataID = selItem.id
        tmpDetail.WasteEUCode = selItem.EUCode1 & "-" & selItem.EUCode2 & "-" & selItem.EUCode3
        tmpDetail.WasteName = selItem.WasteName
        tmpDetail.WasteNo = selItem.WasteNo
        tmpDetail.WasteType = selItem.WasteType
        tmpDetail.WasteGenNo = selItem.WasteCreateNo
        'tmpDetail.isTrailer = selItem.isTrailer
        tmpCol.Add tmpDetail, "'" & tmpDetail.id & "'"
    End If
    tmpDetail.QuotationNo = selItem.QuotationNo
    tmpDetail.QuotationID = selItem.QuotationID
    tmpDetail.LaborCount = selItem.StaffCount
    tmpDetail.ExtraLaborFee = selItem.StaffCharge
    tmpDetail.TransportFee = selItem.TransportFee
    tmpDetail.TransferUnitID = selItem.TransUnitID
    tmpDetail.QuoTreatmentRate = selItem.TreatmentFee
    tmpDetail.TreatmentUnitID = selItem.TreatmentUnitID
    tmpDetail.isPriceIncTrans = selItem.isPriceIncTrans
    tmpDetail.IsDiscountPriceTrans = selItem.IsDiscountPriceTrans
    tmpDetail.MinWeightPerCar = selItem.MinWeightPerCar
    tmpDetail.CanEvap = selItem.CanEvap
    tmpDetail.EvapDesc = selItem.EvapDesc
    tmpDetail.Precipitation = selItem.Precipitation
    tmpDetail.PrecipitationDesc = selItem.PrecipitationDesc
    tmpDetail.Neutralization = selItem.Neutralization
    tmpDetail.NeutralizationDesc = selItem.NeutralizationDesc
    tmpDetail.CommentWaste = selItem.CommentWaste
    tmpDetail.CommentWasteDesc = selItem.CommentWasteDesc
    isDetailChange_ = True
    myMenifest_.isChanged = True
    Set tmpDetail = Nothing
    
    If myMenifest_.AddedWastes Is Nothing Then
        Set myMenifest_.AddedWastes = tmpCol
        isDetailChange_ = True
        myMenifest_.isChanged = True
    End If
    Set tmpCol = Nothing
End Sub

Public Sub RemoveWasteData(selID As String)
Dim tmpCol As Collection
Dim x%
Dim tmpKey As String
    Set tmpCol = myMenifest_.AddedWastes
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
        myMenifest_.isChanged = True
    End If
End Sub

Public Function getWastDataItems() As Collection
On Error Resume Next
Dim x%
Dim tmpCol As New Collection
Dim tmpDet As BWGMenifestWasteAdded
Dim tmpWaste As BWGWasteDataCR
Dim selCol As Collection, tmpStr As String, tmpArr
    Set selCol = myMenifest_.AddedWastes
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
                tmpWaste.QuotationNo = tmpDet.QuotationNo
                tmpWaste.QuotationID = tmpDet.QuotationID
                tmpWaste.WasteType = tmpDet.WasteType
                tmpWaste.CarTypeID = myMenifest_.TruckTypeID
                tmpWaste.CarTypeName = myMenifest_.TruckTypeName
                tmpWaste.StaffCount = tmpDet.LaborCount
                tmpWaste.StaffCharge = tmpDet.ExtraLaborFee
                'tmpWaste.isTrailer = myMenifest_.tru
                
                tmpCol.Add tmpWaste, "'" & tmpDet.wastedataID & "_" & tmpDet.QuotationID & "'"
                Set tmpWaste = Nothing
            End If
            Set tmpDet = Nothing
        Next
    End With

    Set getWastDataItems = tmpCol
    Set tmpCol = Nothing
End Function
