VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{B4BF0775-2BEE-4C33-B2CC-39B9BC2E4704}#4.0#0"; "DateCtl.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCustomer 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "F023"
   ClientHeight    =   9090
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11025
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9090
   ScaleWidth      =   11025
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport rpt1 
      Left            =   2745
      Top             =   8460
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmdOrderState 
      Caption         =   "บันทึกการระงับออเดอร์"
      Height          =   690
      Left            =   45
      Style           =   1  'Graphical
      TabIndex        =   130
      Top             =   8325
      Width           =   2205
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ยกเลิก"
      Height          =   705
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   8310
      Width           =   1005
   End
   Begin VB.TextBox txtCustCode 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   9060
      TabIndex        =   90
      Text            =   "== AUTO =="
      Top             =   450
      Width           =   1935
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "เก็บข้อมูล"
      Height          =   705
      Left            =   4620
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   8310
      Width           =   1005
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "เปิดข้อมูลเก่า"
      Height          =   705
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   8310
      Width           =   1005
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "สร้างใหม่"
      Height          =   705
      Left            =   6690
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   8310
      Width           =   1005
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "ปิด"
      Height          =   705
      Left            =   9930
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   8310
      Width           =   1035
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "ลบข้อมูล"
      Height          =   705
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   8310
      Visible         =   0   'False
      Width           =   1005
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7440
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   810
      Width           =   10995
      _ExtentX        =   19394
      _ExtentY        =   13123
      _Version        =   393216
      Style           =   1
      Tabs            =   12
      TabsPerRow      =   12
      TabHeight       =   520
      BackColor       =   -2147483638
      TabCaption(0)   =   "ข้อมูล"
      TabPicture(0)   =   "frCustomer.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblเวปไซท์(12)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblเวปไซท์(26)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblเวปไซท์(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblเวปไซท์(1)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblเวปไซท์(2)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblเวปไซท์(3)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblเวปไซท์(4)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblเวปไซท์(5)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lblเวปไซท์(6)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lblเวปไซท์(7)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "lblเวปไซท์(8)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "lblเวปไซท์(9)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "lblเวปไซท์(10)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "lblเวปไซท์(11)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Image1(1)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label1(2)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "lblเวปไซท์(14)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "lblเวปไซท์(20)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "lblเวปไซท์(21)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "lblเวปไซท์(22)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "lblเวปไซท์(23)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "lblเวปไซท์(25)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "lblเวปไซท์(27)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "lblเวปไซท์(24)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Label4"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "lbDate"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "lblเวปไซท์(35)"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "lblเวปไซท์(36)"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Label5(2)"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "lblเวปไซท์(55)"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "lblเวปไซท์(56)"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "lblเวปไซท์(57)"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "lblเวปไซท์(64)"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "cboCustGroup"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "Command1"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "txtGenID"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "txtIndustryType"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "txtCompanyNameTH"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "txtAddr"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "txtTumbol"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "txtAumphur"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "txtProvince"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "cboDistrict"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "txtFactoryID"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "txtContactName"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "txtDepartment"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "txtTelNo"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "txtFaxNo"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "lvWaste"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).Control(49)=   "cboCustType"
      Tab(0).Control(49).Enabled=   0   'False
      Tab(0).Control(50)=   "txtParentCust"
      Tab(0).Control(50).Enabled=   0   'False
      Tab(0).Control(51)=   "txtZipCode"
      Tab(0).Control(51).Enabled=   0   'False
      Tab(0).Control(52)=   "txtWebSite"
      Tab(0).Control(52).Enabled=   0   'False
      Tab(0).Control(53)=   "txtEmail"
      Tab(0).Control(53).Enabled=   0   'False
      Tab(0).Control(54)=   "txtShortName"
      Tab(0).Control(54).Enabled=   0   'False
      Tab(0).Control(55)=   "txtCustRemark"
      Tab(0).Control(55).Enabled=   0   'False
      Tab(0).Control(56)=   "txtSaleName"
      Tab(0).Control(56).Enabled=   0   'False
      Tab(0).Control(57)=   "txtComment"
      Tab(0).Control(57).Enabled=   0   'False
      Tab(0).Control(58)=   "cmdDate"
      Tab(0).Control(58).Enabled=   0   'False
      Tab(0).Control(59)=   "txtSale2"
      Tab(0).Control(59).Enabled=   0   'False
      Tab(0).Control(60)=   "txtCountry"
      Tab(0).Control(60).Enabled=   0   'False
      Tab(0).Control(61)=   "cboDTeam"
      Tab(0).Control(61).Enabled=   0   'False
      Tab(0).Control(62)=   "TxtRefCode"
      Tab(0).Control(62).Enabled=   0   'False
      Tab(0).Control(63)=   "cbDept"
      Tab(0).Control(63).Enabled=   0   'False
      Tab(0).Control(64)=   "cbCRStaff"
      Tab(0).Control(64).Enabled=   0   'False
      Tab(0).Control(65)=   "txtPicOfDay"
      Tab(0).Control(65).Enabled=   0   'False
      Tab(0).ControlCount=   66
      TabCaption(1)   =   "สถานที่เก็บ"
      TabPicture(1)   =   "frCustomer.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Command2"
      Tab(1).Control(1)=   "lvSiteAddress"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "ด้านบัญชี"
      TabPicture(2)   =   "frCustomer.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "CbCredit"
      Tab(2).Control(1)=   "ChkBreak"
      Tab(2).Control(2)=   "txtCreditterm"
      Tab(2).Control(3)=   "txtREAccFax"
      Tab(2).Control(4)=   "txtREAccTel"
      Tab(2).Control(5)=   "txtREAccContact"
      Tab(2).Control(6)=   "Frame3"
      Tab(2).Control(7)=   "Frame2"
      Tab(2).Control(8)=   "txtRemdesc"
      Tab(2).Control(9)=   "Frame1"
      Tab(2).Control(10)=   "Command3"
      Tab(2).Control(11)=   "lvBillAddr"
      Tab(2).Control(12)=   "Label17"
      Tab(2).Control(13)=   "Label16"
      Tab(2).ControlCount=   14
      TabCaption(3)   =   "สัญญา"
      TabPicture(3)   =   "frCustomer.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Command4"
      Tab(3).Control(1)=   "lvContract"
      Tab(3).ControlCount=   2
      TabCaption(4)   =   "เจ้าหน้าที่"
      TabPicture(4)   =   "frCustomer.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Command5"
      Tab(4).Control(1)=   "lvStaff"
      Tab(4).ControlCount=   2
      TabCaption(5)   =   "หมายเหตุ"
      TabPicture(5)   =   "frCustomer.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Command6"
      Tab(5).Control(1)=   "lvCustRem"
      Tab(5).ControlCount=   2
      TabCaption(6)   =   "การขนส่ง"
      TabPicture(6)   =   "frCustomer.frx":00A8
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "ChkGPS"
      Tab(6).Control(1)=   "Frame5"
      Tab(6).Control(2)=   "cmdSaveTR_Remark"
      Tab(6).Control(3)=   "Frame4"
      Tab(6).Control(4)=   "Label3"
      Tab(6).ControlCount=   5
      TabCaption(7)   =   "ข้อตกลง"
      TabPicture(7)   =   "frCustomer.frx":00C4
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "ChkVIPCR"
      Tab(7).Control(1)=   "TxtBeginDate"
      Tab(7).Control(2)=   "ChkISWeigth"
      Tab(7).Control(3)=   "txtRemTimeNo"
      Tab(7).Control(4)=   "chkisCheck"
      Tab(7).Control(5)=   "Command7"
      Tab(7).Control(6)=   "lvCommentQuo"
      Tab(7).Control(7)=   "Label15"
      Tab(7).Control(8)=   "Label14"
      Tab(7).ControlCount=   9
      TabCaption(8)   =   "การใช้ BOX"
      TabPicture(8)   =   "frCustomer.frx":00E0
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "FramBox"
      Tab(8).Control(1)=   "chkisBox"
      Tab(8).ControlCount=   2
      TabCaption(9)   =   "ข้อร้องเรียน"
      TabPicture(9)   =   "frCustomer.frx":00FC
      Tab(9).ControlEnabled=   0   'False
      Tab(9).Control(0)=   "CmdAddProblem"
      Tab(9).Control(1)=   "ChkOvrW"
      Tab(9).Control(2)=   "LvProb"
      Tab(9).ControlCount=   3
      TabCaption(10)  =   "บันทึกการติดต่อ"
      TabPicture(10)  =   "frCustomer.frx":0118
      Tab(10).ControlEnabled=   0   'False
      Tab(10).Control(0)=   "cmbPrintActivity"
      Tab(10).Control(1)=   "CmdAdd"
      Tab(10).Control(2)=   "lvCustActivity"
      Tab(10).ControlCount=   3
      TabCaption(11)  =   "งานประมูล"
      TabPicture(11)  =   "frCustomer.frx":0134
      Tab(11).ControlEnabled=   0   'False
      Tab(11).Control(0)=   "cmAuction"
      Tab(11).Control(1)=   "lvAuction"
      Tab(11).ControlCount=   2
      Begin VB.TextBox txtPicOfDay 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   5300
         TabIndex        =   286
         Top             =   3720
         Width           =   2355
      End
      Begin VB.CommandButton cmAuction 
         Caption         =   "เพิ่มข้อมูลงาน"
         Height          =   345
         Left            =   -74955
         Style           =   1  'Graphical
         TabIndex        =   283
         Top             =   6240
         Width           =   2205
      End
      Begin VB.ComboBox cbCRStaff 
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   257
         Top             =   3720
         Width           =   2715
      End
      Begin VB.CommandButton cmbPrintActivity 
         Caption         =   "พิมพ์รายงาน"
         Height          =   600
         Left            =   -66405
         Style           =   1  'Graphical
         TabIndex        =   255
         Top             =   795
         Width           =   2205
      End
      Begin VB.CommandButton CmdAdd 
         Caption         =   "เพิ่มบันทึกการติดต่อ"
         Height          =   600
         Left            =   -74910
         Style           =   1  'Graphical
         TabIndex        =   253
         Top             =   795
         Width           =   2205
      End
      Begin VB.ComboBox CbCredit 
         Height          =   315
         ItemData        =   "frCustomer.frx":0150
         Left            =   -67665
         List            =   "frCustomer.frx":0152
         Style           =   2  'Dropdown List
         TabIndex        =   251
         Top             =   5970
         Width           =   1050
      End
      Begin VB.ComboBox cbDept 
         Height          =   315
         Left            =   8400
         Style           =   2  'Dropdown List
         TabIndex        =   246
         Top             =   3000
         Width           =   2355
      End
      Begin VB.CheckBox ChkVIPCR 
         Caption         =   "VIP CR"
         Height          =   195
         Left            =   -74505
         TabIndex        =   243
         Top             =   795
         Width           =   1860
      End
      Begin VB.CheckBox ChkBreak 
         Caption         =   "ระงับการขนชั่วคราว"
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   -70365
         TabIndex        =   242
         Top             =   705
         Width           =   2490
      End
      Begin VB.CommandButton CmdAddProblem 
         Caption         =   "เพิ่มข้อร้องเรียน"
         Height          =   600
         Left            =   -66405
         Style           =   1  'Graphical
         TabIndex        =   240
         Top             =   750
         Width           =   2205
      End
      Begin VB.CheckBox ChkOvrW 
         Caption         =   "เคยน้ำหนักเกิน"
         Height          =   195
         Left            =   -74910
         TabIndex        =   239
         Top             =   3990
         Width           =   1815
      End
      Begin VB.CheckBox ChkGPS 
         Caption         =   "ต้องการรายงาน GPS"
         Height          =   330
         Left            =   -74820
         TabIndex        =   238
         Top             =   6450
         Width           =   1995
      End
      Begin VB.TextBox TxtBeginDate 
         Height          =   330
         Left            =   -66585
         TabIndex        =   237
         Top             =   1185
         Width           =   1995
      End
      Begin VB.TextBox TxtRefCode 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   8955
         TabIndex        =   234
         Top             =   1545
         Width           =   1785
      End
      Begin VB.TextBox txtCreditterm 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -65145
         TabIndex        =   233
         Top             =   5925
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Frame FramBox 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   6495
         Left            =   -74280
         TabIndex        =   186
         Top             =   765
         Width           =   8295
         Begin VB.TextBox txtBoutCageW 
            Height          =   285
            Left            =   4575
            TabIndex        =   279
            Top             =   4230
            Width           =   495
         End
         Begin VB.TextBox txtBoutCage 
            Height          =   285
            Left            =   1935
            TabIndex        =   278
            Top             =   4230
            Width           =   495
         End
         Begin VB.TextBox txtBuotCageBack 
            Height          =   285
            Left            =   5415
            TabIndex        =   277
            Top             =   4230
            Width           =   495
         End
         Begin VB.TextBox txtWaterBoutW 
            Height          =   285
            Left            =   4575
            TabIndex        =   273
            Top             =   3870
            Width           =   495
         End
         Begin VB.TextBox txtWaterBuot 
            Height          =   285
            Left            =   1935
            TabIndex        =   272
            Top             =   3870
            Width           =   495
         End
         Begin VB.TextBox txtWaterBoutBack 
            Height          =   285
            Left            =   5415
            TabIndex        =   271
            Top             =   3870
            Width           =   495
         End
         Begin VB.TextBox txtBuotSW 
            Height          =   285
            Left            =   4575
            TabIndex        =   267
            Top             =   3510
            Width           =   495
         End
         Begin VB.TextBox txtBoutS 
            Height          =   285
            Left            =   1935
            TabIndex        =   266
            Top             =   3510
            Width           =   495
         End
         Begin VB.TextBox txtBoutSBack 
            Height          =   285
            Left            =   5415
            TabIndex        =   265
            Top             =   3510
            Width           =   495
         End
         Begin VB.TextBox txtLackW 
            Height          =   285
            Left            =   4575
            TabIndex        =   261
            Top             =   3150
            Width           =   495
         End
         Begin VB.TextBox txtLack 
            Height          =   285
            Left            =   1935
            TabIndex        =   260
            Top             =   3150
            Width           =   495
         End
         Begin VB.TextBox txtLackBack 
            Height          =   285
            Left            =   5415
            TabIndex        =   259
            Top             =   3150
            Width           =   495
         End
         Begin VB.TextBox txtPriceBox 
            Height          =   285
            Left            =   1560
            TabIndex        =   231
            Top             =   750
            Width           =   1215
         End
         Begin VB.CheckBox chkNon 
            Caption         =   "Check1"
            Height          =   195
            Left            =   1680
            TabIndex        =   229
            Top             =   360
            Width           =   195
         End
         Begin VB.CheckBox chkHaz 
            Caption         =   "Check1"
            Height          =   195
            Left            =   960
            TabIndex        =   226
            Top             =   360
            Width           =   195
         End
         Begin VB.TextBox txtTruckBack 
            Height          =   285
            Left            =   5400
            TabIndex        =   218
            Top             =   2790
            Width           =   495
         End
         Begin VB.TextBox txtTwoBack 
            Height          =   285
            Left            =   5400
            TabIndex        =   217
            Top             =   2430
            Width           =   495
         End
         Begin VB.TextBox txtOneBack 
            Height          =   285
            Left            =   5400
            TabIndex        =   216
            Top             =   2070
            Width           =   495
         End
         Begin VB.TextBox txtROLBack 
            Height          =   285
            Left            =   5400
            TabIndex        =   215
            Top             =   1710
            Width           =   495
         End
         Begin VB.TextBox txtLUGBack 
            Height          =   285
            Left            =   5400
            TabIndex        =   214
            Top             =   1350
            Width           =   495
         End
         Begin VB.TextBox txtBoxDesc 
            Height          =   1455
            Left            =   240
            TabIndex        =   212
            Top             =   4845
            Width           =   7455
         End
         Begin VB.TextBox txtTruckBOX 
            Height          =   285
            Left            =   1920
            TabIndex        =   210
            Top             =   2790
            Width           =   495
         End
         Begin VB.TextBox txtLugTruckW 
            Height          =   285
            Left            =   4560
            TabIndex        =   208
            Top             =   2790
            Width           =   495
         End
         Begin VB.TextBox txtBoxTwoW 
            Height          =   285
            Left            =   4560
            TabIndex        =   207
            Top             =   2430
            Width           =   495
         End
         Begin VB.TextBox txtBoxOneW 
            Height          =   285
            Left            =   4560
            TabIndex        =   206
            Top             =   2070
            Width           =   495
         End
         Begin VB.TextBox txtRolBoxW 
            Height          =   285
            Left            =   4560
            TabIndex        =   205
            Top             =   1710
            Width           =   495
         End
         Begin VB.TextBox txtLugBoxW 
            Height          =   285
            Left            =   4560
            TabIndex        =   204
            Top             =   1350
            Width           =   495
         End
         Begin VB.TextBox txtTwoT 
            Height          =   285
            Left            =   1920
            TabIndex        =   199
            Top             =   2430
            Width           =   495
         End
         Begin VB.TextBox txtOneT 
            Height          =   285
            Left            =   1920
            TabIndex        =   196
            Top             =   2070
            Width           =   495
         End
         Begin VB.TextBox txtROLBOX 
            Height          =   285
            Left            =   1920
            TabIndex        =   193
            Top             =   1710
            Width           =   495
         End
         Begin VB.TextBox txtLUGBOX 
            Height          =   285
            Left            =   1920
            TabIndex        =   189
            Top             =   1350
            Width           =   495
         End
         Begin WasteManagment.ctlDate dtLugFromdate 
            Height          =   315
            Left            =   2520
            TabIndex        =   191
            Top             =   1350
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   556
            BackColor       =   -2147483633
         End
         Begin DateCtl.Date dtROLFromdate 
            Height          =   315
            Left            =   2520
            TabIndex        =   194
            Top             =   1710
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   556
            BackColor       =   -2147483633
         End
         Begin DateCtl.Date dtOneTFrom 
            Height          =   315
            Left            =   2520
            TabIndex        =   197
            Top             =   2070
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   556
            BackColor       =   -2147483633
         End
         Begin DateCtl.Date dtTwoTFrom 
            Height          =   315
            Left            =   2520
            TabIndex        =   200
            Top             =   2430
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   556
            BackColor       =   -2147483633
         End
         Begin WasteManagment.ctlDate dtTruckFrom 
            Height          =   315
            Left            =   2520
            TabIndex        =   209
            Top             =   2790
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   556
            BackColor       =   -2147483633
         End
         Begin WasteManagment.ctlDate dtLugBackdate 
            Height          =   315
            Left            =   6000
            TabIndex        =   219
            Top             =   1350
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   556
            BackColor       =   -2147483633
         End
         Begin WasteManagment.ctlDate dtROLBackdate 
            Height          =   315
            Left            =   6000
            TabIndex        =   220
            Top             =   1710
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   556
            BackColor       =   -2147483633
         End
         Begin WasteManagment.ctlDate dtOneTBack 
            Height          =   315
            Left            =   6000
            TabIndex        =   221
            Top             =   2070
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   556
            BackColor       =   -2147483633
         End
         Begin WasteManagment.ctlDate dtTwoTBack 
            Height          =   315
            Left            =   6000
            TabIndex        =   222
            Top             =   2430
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   556
            BackColor       =   -2147483633
         End
         Begin WasteManagment.ctlDate dtTruckBack 
            Height          =   315
            Left            =   6000
            TabIndex        =   223
            Top             =   2790
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   556
            BackColor       =   -2147483633
         End
         Begin WasteManagment.ctlDate dtLackFrom 
            Height          =   315
            Left            =   2535
            TabIndex        =   262
            Top             =   3150
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   556
            BackColor       =   -2147483633
         End
         Begin WasteManagment.ctlDate dtLackBack 
            Height          =   315
            Left            =   5985
            TabIndex        =   264
            Top             =   3150
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   556
            BackColor       =   -2147483633
         End
         Begin WasteManagment.ctlDate dtBoutSFrom 
            Height          =   315
            Left            =   2535
            TabIndex        =   268
            Top             =   3510
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   556
            BackColor       =   -2147483633
         End
         Begin WasteManagment.ctlDate dtBoutSBack 
            Height          =   315
            Left            =   5985
            TabIndex        =   270
            Top             =   3510
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   556
            BackColor       =   -2147483633
         End
         Begin WasteManagment.ctlDate dtWaterBoutFrom 
            Height          =   315
            Left            =   2535
            TabIndex        =   274
            Top             =   3870
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   556
            BackColor       =   -2147483633
         End
         Begin WasteManagment.ctlDate dtWaterBoutBack 
            Height          =   315
            Left            =   5985
            TabIndex        =   276
            Top             =   3870
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   556
            BackColor       =   -2147483633
         End
         Begin WasteManagment.ctlDate dtBoutCageFrom 
            Height          =   315
            Left            =   2535
            TabIndex        =   280
            Top             =   4230
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   556
            BackColor       =   -2147483633
         End
         Begin WasteManagment.ctlDate dtBoutCageBack 
            Height          =   315
            Left            =   5985
            TabIndex        =   282
            Top             =   4230
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   556
            BackColor       =   -2147483633
         End
         Begin VB.Label lblเวปไซท์ 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "กรงเบ้า"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   63
            Left            =   495
            TabIndex        =   281
            Top             =   4230
            Width           =   525
         End
         Begin VB.Label lblเวปไซท์ 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "เบ้าใส่น้ำ"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   62
            Left            =   495
            TabIndex        =   275
            Top             =   3870
            Width           =   630
         End
         Begin VB.Label lblเวปไซท์ 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "เบ้าตัดปาก"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   59
            Left            =   495
            TabIndex        =   269
            Top             =   3510
            Width           =   750
         End
         Begin VB.Label lblเวปไซท์ 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Lack"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   58
            Left            =   495
            TabIndex        =   263
            Top             =   3150
            Width           =   360
         End
         Begin VB.Label lblเวปไซท์ 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ค่าบริการวาง Box"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   54
            Left            =   240
            TabIndex        =   258
            Top             =   750
            Width           =   1200
         End
         Begin VB.Label lblเวปไซท์ 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Non"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   49
            Left            =   1920
            TabIndex        =   230
            Top             =   360
            Width           =   300
         End
         Begin VB.Label lblเวปไซท์ 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ประเภท"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   47
            Left            =   240
            TabIndex        =   228
            Top             =   360
            Width           =   555
         End
         Begin VB.Label lblเวปไซท์ 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Haz"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   46
            Left            =   1200
            TabIndex        =   227
            Top             =   360
            Width           =   285
         End
         Begin VB.Label lblเวปไซท์ 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "บันทึกข้อความ"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   44
            Left            =   240
            TabIndex        =   225
            Top             =   4605
            Width           =   990
         End
         Begin VB.Label lblเวปไซท์ 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "วันที่"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   43
            Left            =   6120
            TabIndex        =   224
            Top             =   1110
            Width           =   315
         End
         Begin VB.Label lblเวปไซท์ 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "เอากลับ"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   61
            Left            =   5400
            TabIndex        =   213
            Top             =   1110
            Width           =   540
         End
         Begin VB.Label lblเวปไซท์ 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ROLL OFF TANK"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   60
            Left            =   480
            TabIndex        =   211
            Top             =   2790
            Width           =   1245
         End
         Begin VB.Label lblเวปไซท์ 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "น้ำหนัก/เที่ยว"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   53
            Left            =   4320
            TabIndex        =   203
            Top             =   1110
            Width           =   975
         End
         Begin VB.Label lblเวปไซท์ 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "BOX 2 T"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   48
            Left            =   480
            TabIndex        =   198
            Top             =   2430
            Width           =   615
         End
         Begin VB.Label lblเวปไซท์ 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "BOX 1 T"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   45
            Left            =   480
            TabIndex        =   195
            Top             =   2070
            Width           =   600
         End
         Begin VB.Label lblเวปไซท์ 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ROLL OFF BOX"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   42
            Left            =   480
            TabIndex        =   192
            Top             =   1710
            Width           =   1140
         End
         Begin VB.Label lblเวปไซท์ 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "เริ่มวางวันที่"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   41
            Left            =   2520
            TabIndex        =   190
            Top             =   1110
            Width           =   825
         End
         Begin VB.Label lblเวปไซท์ 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "จำนวน"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   40
            Left            =   1920
            TabIndex        =   188
            Top             =   1110
            Width           =   495
         End
         Begin VB.Label lblเวปไซท์ 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "LUGGER BOX"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   17
            Left            =   480
            TabIndex        =   187
            Top             =   1350
            Width           =   1050
         End
      End
      Begin VB.CheckBox chkisBox 
         Caption         =   "Check1"
         Height          =   255
         Left            =   -74760
         TabIndex        =   184
         Top             =   690
         Width           =   135
      End
      Begin VB.CheckBox ChkISWeigth 
         Caption         =   "ส่งแยกน้ำหนักฝ่ายการตลาด"
         Height          =   255
         Left            =   -70200
         TabIndex        =   183
         Top             =   1200
         Width           =   2415
      End
      Begin VB.TextBox txtRemTimeNo 
         Height          =   375
         Left            =   -70500
         TabIndex        =   182
         Top             =   6840
         Width           =   5895
      End
      Begin VB.CheckBox chkisCheck 
         Height          =   255
         Left            =   -74520
         TabIndex        =   179
         Top             =   1200
         Width           =   195
      End
      Begin VB.TextBox txtREAccFax 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   -67440
         TabIndex        =   162
         Top             =   4800
         Width           =   2355
      End
      Begin VB.TextBox txtREAccTel 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   -70740
         TabIndex        =   161
         Top             =   4800
         Width           =   2715
      End
      Begin VB.TextBox txtREAccContact 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   -73920
         TabIndex        =   160
         Top             =   4800
         Width           =   2715
      End
      Begin VB.ComboBox cboDTeam 
         Height          =   315
         Left            =   8400
         Style           =   2  'Dropdown List
         TabIndex        =   138
         Top             =   2625
         Width           =   2355
      End
      Begin VB.TextBox txtCountry 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   9420
         TabIndex        =   137
         Top             =   825
         Width           =   1335
      End
      Begin VB.Frame Frame5 
         Height          =   555
         Left            =   -74820
         TabIndex        =   132
         Top             =   1440
         Width           =   10575
         Begin VB.OptionButton Option3 
            Caption         =   "ไม่ระบุ (หรือกำหนดเป็นอย่างอื่น)"
            Height          =   255
            Left            =   7320
            TabIndex        =   135
            Top             =   180
            Width           =   2715
         End
         Begin VB.OptionButton Option2 
            Height          =   195
            Left            =   3900
            TabIndex        =   134
            Top             =   180
            Width           =   195
         End
         Begin VB.OptionButton Option1 
            Height          =   195
            Left            =   180
            TabIndex        =   133
            Top             =   180
            Width           =   195
         End
         Begin VB.Label lblเวปไซท์ 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ปิดน้ำหนักผู้ก่อกำเนิด"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   52
            Left            =   4200
            TabIndex        =   202
            Top             =   185
            Width           =   1515
         End
         Begin VB.Label lblเวปไซท์ 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ปิดน้ำหนักผู้รับกำจัด"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   51
            Left            =   480
            TabIndex        =   201
            Top             =   180
            Width           =   1410
         End
      End
      Begin VB.TextBox txtSale2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   8460
         TabIndex        =   131
         Top             =   3345
         Width           =   2235
      End
      Begin VB.CommandButton Command7 
         Caption         =   "เพิ่มบันทึกข้อตกลง"
         Height          =   375
         Left            =   -74520
         TabIndex        =   128
         Top             =   6840
         Width           =   2535
      End
      Begin VB.CommandButton cmdDate 
         Caption         =   "ระบุวันที่"
         Height          =   285
         Left            =   10200
         TabIndex        =   127
         Top             =   7050
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtComment 
         Height          =   375
         Left            =   6120
         TabIndex        =   124
         Top             =   7005
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.TextBox txtSaleName 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   3345
         Width           =   3255
      End
      Begin VB.CommandButton cmdSaveTR_Remark 
         Caption         =   "บันทึก"
         Height          =   435
         Left            =   -65520
         TabIndex        =   121
         Top             =   6540
         Width           =   1215
      End
      Begin VB.Frame Frame4 
         Height          =   4515
         Left            =   -74820
         TabIndex        =   120
         Top             =   1920
         Width           =   10575
         Begin WasteManagment.ctlGrid ctlGridTR 
            Height          =   4215
            Left            =   120
            TabIndex        =   122
            Top             =   180
            Width           =   10335
            _ExtentX        =   18230
            _ExtentY        =   7435
         End
      End
      Begin VB.CommandButton Command6 
         Caption         =   "เพิ่มข้อมูลหมายเหตุ"
         Height          =   345
         Left            =   -74940
         Style           =   1  'Graphical
         TabIndex        =   116
         Top             =   7050
         Width           =   2205
      End
      Begin VB.Frame Frame3 
         Caption         =   "รายละเอียดการรับเช็ค"
         Height          =   1725
         Left            =   -74880
         TabIndex        =   111
         Top             =   3060
         Width           =   10575
         Begin VB.TextBox txtReTime 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3900
            TabIndex        =   171
            Top             =   960
            Width           =   1035
         End
         Begin VB.TextBox txtReLocateion 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1380
            TabIndex        =   169
            Top             =   960
            Width           =   1995
         End
         Begin VB.ComboBox cboRE_Men 
            Height          =   315
            Left            =   5880
            TabIndex        =   159
            Top             =   900
            Width           =   2355
         End
         Begin VB.CheckBox isRE_CB 
            Caption         =   "รับเช็คที่บริษัทลูกค้า"
            Height          =   195
            Left            =   5100
            TabIndex        =   157
            Top             =   660
            Width           =   1935
         End
         Begin VB.CheckBox isRV_EMS 
            Caption         =   "EMS"
            Height          =   255
            Left            =   4320
            TabIndex        =   156
            Top             =   650
            Width           =   795
         End
         Begin VB.CheckBox isRV_Inbank 
            Caption         =   "รับเช็คผ่านธนาคาร"
            Height          =   315
            Left            =   2580
            TabIndex        =   155
            Top             =   650
            Width           =   1695
         End
         Begin VB.CheckBox isRV_Bank 
            Caption         =   "ลูกค้าโอน"
            Height          =   255
            Left            =   1560
            TabIndex        =   154
            Top             =   650
            Width           =   1035
         End
         Begin VB.OptionButton optCheqType 
            Height          =   255
            Index           =   1
            Left            =   390
            TabIndex        =   49
            Top             =   300
            Width           =   255
         End
         Begin VB.OptionButton optCheqType 
            Height          =   255
            Index           =   2
            Left            =   1620
            TabIndex        =   50
            Top             =   240
            Width           =   255
         End
         Begin VB.TextBox txtCheqDate 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Index           =   1
            Left            =   3120
            TabIndex        =   51
            Top             =   240
            Width           =   315
         End
         Begin VB.TextBox txtCheqDate 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Index           =   2
            Left            =   3570
            TabIndex        =   52
            Top             =   240
            Width           =   315
         End
         Begin VB.TextBox txtCheqDate 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Index           =   3
            Left            =   4020
            TabIndex        =   53
            Top             =   240
            Width           =   315
         End
         Begin VB.TextBox txtCheqDate 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Index           =   4
            Left            =   4470
            TabIndex        =   54
            Top             =   240
            Width           =   315
         End
         Begin VB.TextBox txtCheqDate 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Index           =   5
            Left            =   4920
            TabIndex        =   55
            Top             =   240
            Width           =   315
         End
         Begin VB.OptionButton optCheqType 
            Caption         =   "ทุกวัน"
            Height          =   255
            Index           =   3
            Left            =   5970
            TabIndex        =   56
            Top             =   300
            Width           =   765
         End
         Begin VB.CheckBox chkCheqDay 
            Caption         =   "จ."
            Height          =   195
            Index           =   1
            Left            =   6900
            TabIndex        =   57
            Top             =   210
            Width           =   465
         End
         Begin VB.CheckBox chkCheqDay 
            Caption         =   "อ."
            Height          =   195
            Index           =   2
            Left            =   7440
            TabIndex        =   58
            Top             =   210
            Width           =   465
         End
         Begin VB.CheckBox chkCheqDay 
            Caption         =   "พ."
            Height          =   195
            Index           =   3
            Left            =   7980
            TabIndex        =   59
            Top             =   210
            Width           =   465
         End
         Begin VB.CheckBox chkCheqDay 
            Caption         =   "พฤ."
            Height          =   195
            Index           =   4
            Left            =   6900
            TabIndex        =   60
            Top             =   450
            Width           =   585
         End
         Begin VB.CheckBox chkCheqDay 
            Caption         =   "ศ."
            Height          =   195
            Index           =   5
            Left            =   7980
            TabIndex        =   61
            Top             =   450
            Width           =   465
         End
         Begin VB.ComboBox cboCheqWeekCount 
            Height          =   315
            Left            =   8640
            Style           =   2  'Dropdown List
            TabIndex        =   62
            Top             =   270
            Width           =   1845
         End
         Begin VB.OptionButton optCheqType 
            Height          =   255
            Index           =   4
            Left            =   360
            TabIndex        =   63
            Top             =   1320
            Width           =   285
         End
         Begin VB.TextBox txtCheqDetail 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   1380
            TabIndex        =   64
            Top             =   1260
            Width           =   8925
         End
         Begin VB.Label Label11 
            Caption         =   "เวลา"
            Height          =   255
            Left            =   3480
            TabIndex        =   170
            Top             =   960
            Width           =   435
         End
         Begin VB.Label Label9 
            Caption         =   "สถานที่"
            Height          =   255
            Index           =   1
            Left            =   540
            TabIndex        =   168
            Top             =   960
            Width           =   735
         End
         Begin VB.Label Label8 
            Caption         =   "เจ้าหน้าที่ :"
            Height          =   255
            Left            =   5040
            TabIndex        =   158
            Top             =   960
            Width           =   855
         End
         Begin VB.Label Label7 
            Caption         =   "เฉพาะวันที่"
            Height          =   255
            Index           =   1
            Left            =   1980
            TabIndex        =   148
            Top             =   300
            Width           =   855
         End
         Begin VB.Label Label6 
            Caption         =   "อื่นๆ ระบุ"
            Height          =   315
            Index           =   1
            Left            =   660
            TabIndex        =   147
            Top             =   1260
            Width           =   735
         End
         Begin VB.Label Label5 
            Caption         =   "ทุกวัน"
            Height          =   255
            Index           =   1
            Left            =   720
            TabIndex        =   146
            Top             =   300
            Width           =   735
         End
         Begin VB.Label lblเวปไซท์ 
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
         Begin VB.Label lblเวปไซท์ 
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
         Begin VB.Label lblเวปไซท์ 
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
         Begin VB.Label lblเวปไซท์ 
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
         Height          =   1995
         Left            =   -74880
         TabIndex        =   106
         Top             =   960
         Width           =   10575
         Begin VB.Frame Frame6 
            Height          =   420
            Left            =   4770
            TabIndex        =   247
            Top             =   630
            Width           =   3300
            Begin VB.OptionButton Option6 
               Caption         =   "เจ้าหน้าที่วางบิล"
               Height          =   195
               Left            =   1800
               TabIndex        =   250
               Top             =   135
               Width           =   1455
            End
            Begin VB.OptionButton Option5 
               Caption         =   "EMS"
               Height          =   195
               Left            =   1035
               TabIndex        =   249
               Top             =   135
               Width           =   1005
            End
            Begin VB.OptionButton Option4 
               Caption         =   "ภายใน"
               Height          =   195
               Left            =   90
               TabIndex        =   248
               Top             =   135
               Width           =   1005
            End
         End
         Begin VB.CheckBox ChkOrther 
            Caption         =   "ภายใน"
            Height          =   255
            Left            =   4995
            TabIndex        =   178
            Top             =   1095
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.TextBox txtInv_Fax 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   7140
            TabIndex        =   177
            Top             =   1500
            Width           =   2355
         End
         Begin VB.TextBox txtInv_tel 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   4440
            TabIndex        =   175
            Top             =   1500
            Width           =   2055
         End
         Begin VB.TextBox txtInv_Contact 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1440
            TabIndex        =   173
            Top             =   1500
            Width           =   2595
         End
         Begin VB.TextBox txtInvTime 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3900
            TabIndex        =   167
            Top             =   660
            Width           =   795
         End
         Begin VB.TextBox txtLocation 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1440
            TabIndex        =   165
            Top             =   660
            Width           =   1875
         End
         Begin VB.CheckBox IsInv 
            Caption         =   "EMS"
            Height          =   195
            Left            =   5880
            TabIndex        =   163
            Top             =   1140
            Visible         =   0   'False
            Width           =   675
         End
         Begin VB.ComboBox ComboInvMen 
            Height          =   315
            Left            =   8100
            TabIndex        =   153
            Top             =   720
            Width           =   2415
         End
         Begin VB.CheckBox isInv_Men 
            Caption         =   "เจ้าหน้าที่วางบิล"
            Height          =   195
            Left            =   6660
            TabIndex        =   152
            Top             =   1140
            Visible         =   0   'False
            Width           =   1395
         End
         Begin VB.TextBox txtBillingDet 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   1440
            TabIndex        =   48
            Top             =   1080
            Width           =   8925
         End
         Begin VB.OptionButton optBillType 
            Height          =   255
            Index           =   4
            Left            =   240
            TabIndex        =   47
            Top             =   1080
            Width           =   285
         End
         Begin VB.ComboBox cboBillWeekCount 
            Height          =   315
            Left            =   8640
            Style           =   2  'Dropdown List
            TabIndex        =   46
            Top             =   270
            Width           =   1845
         End
         Begin VB.CheckBox chkBillDay 
            Caption         =   "ศ."
            Height          =   195
            Index           =   5
            Left            =   7980
            TabIndex        =   45
            Top             =   450
            Width           =   465
         End
         Begin VB.CheckBox chkBillDay 
            Caption         =   "พฤ."
            Height          =   195
            Index           =   4
            Left            =   6900
            TabIndex        =   44
            Top             =   450
            Width           =   585
         End
         Begin VB.CheckBox chkBillDay 
            Caption         =   "พ."
            Height          =   195
            Index           =   3
            Left            =   7980
            TabIndex        =   43
            Top             =   210
            Width           =   465
         End
         Begin VB.CheckBox chkBillDay 
            Caption         =   "อ."
            Height          =   195
            Index           =   2
            Left            =   7440
            TabIndex        =   42
            Top             =   210
            Width           =   465
         End
         Begin VB.CheckBox chkBillDay 
            Caption         =   "จ."
            Height          =   195
            Index           =   1
            Left            =   6900
            TabIndex        =   41
            Top             =   210
            Width           =   465
         End
         Begin VB.OptionButton optBillType 
            Caption         =   "ทุกวัน"
            Height          =   255
            Index           =   3
            Left            =   5970
            TabIndex        =   40
            Top             =   300
            Width           =   765
         End
         Begin VB.TextBox txtBillDate 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Index           =   5
            Left            =   4920
            TabIndex        =   39
            Top             =   240
            Width           =   315
         End
         Begin VB.TextBox txtBillDate 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Index           =   4
            Left            =   4470
            TabIndex        =   38
            Top             =   240
            Width           =   315
         End
         Begin VB.TextBox txtBillDate 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Index           =   3
            Left            =   4020
            TabIndex        =   37
            Top             =   240
            Width           =   315
         End
         Begin VB.TextBox txtBillDate 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Index           =   2
            Left            =   3570
            TabIndex        =   36
            Top             =   240
            Width           =   315
         End
         Begin VB.TextBox txtBillDate 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Index           =   1
            Left            =   3120
            TabIndex        =   35
            Top             =   240
            Width           =   315
         End
         Begin VB.OptionButton optBillType 
            Height          =   255
            Index           =   2
            Left            =   1740
            TabIndex        =   34
            Top             =   300
            Width           =   255
         End
         Begin VB.OptionButton optBillType 
            Height          =   255
            Index           =   1
            Left            =   300
            TabIndex        =   33
            Top             =   300
            Width           =   255
         End
         Begin VB.Label Label13 
            Caption         =   "แฟกซ์"
            Height          =   315
            Left            =   6660
            TabIndex        =   176
            Top             =   1500
            Width           =   435
         End
         Begin VB.Label Label12 
            Caption         =   "โทร"
            Height          =   315
            Left            =   4080
            TabIndex        =   174
            Top             =   1500
            Width           =   375
         End
         Begin VB.Label Inv_Contact 
            Caption         =   "ติดต่อ"
            Height          =   255
            Index           =   0
            Left            =   780
            TabIndex        =   172
            Top             =   1560
            Width           =   555
         End
         Begin VB.Label Label10 
            Caption         =   "เวลา"
            Height          =   315
            Left            =   3420
            TabIndex        =   166
            Top             =   660
            Width           =   375
         End
         Begin VB.Label Label9 
            Caption         =   "สถานที่"
            Height          =   255
            Index           =   0
            Left            =   600
            TabIndex        =   164
            Top             =   720
            Width           =   735
         End
         Begin VB.Label Label7 
            Caption         =   "เฉพาะวันที่"
            Height          =   255
            Index           =   0
            Left            =   1980
            TabIndex        =   145
            Top             =   300
            Width           =   855
         End
         Begin VB.Label Label6 
            Caption         =   "อื่นๆ ระบุ"
            Height          =   315
            Index           =   0
            Left            =   660
            TabIndex        =   144
            Top             =   1080
            Width           =   735
         End
         Begin VB.Label Label5 
            Caption         =   "ทุกวัน"
            Height          =   255
            Index           =   0
            Left            =   660
            TabIndex        =   143
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
         Begin VB.Label lblเวปไซท์ 
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
         Begin VB.Label lblเวปไซท์ 
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
         Begin VB.Label lblเวปไซท์ 
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
         Begin VB.Label lblเวปไซท์ 
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
         TabIndex        =   17
         Top             =   2985
         Width           =   5895
      End
      Begin VB.TextBox txtShortName 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   8400
         MaxLength       =   4
         TabIndex        =   14
         Top             =   2265
         Width           =   2355
      End
      Begin VB.TextBox txtEmail 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   5100
         TabIndex        =   16
         Top             =   2625
         Width           =   2355
      End
      Begin VB.TextBox txtWebSite 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1560
         TabIndex        =   15
         Top             =   2625
         Width           =   2655
      End
      Begin VB.TextBox txtZipCode 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   7920
         MaxLength       =   5
         TabIndex        =   6
         Top             =   825
         Width           =   735
      End
      Begin VB.CommandButton Command5 
         Caption         =   "เพิ่มข้อมูลเจ้าหน้าที่"
         Height          =   345
         Left            =   -74940
         Style           =   1  'Graphical
         TabIndex        =   97
         Top             =   7065
         Width           =   2205
      End
      Begin VB.CommandButton Command4 
         Caption         =   "เพิ่มข้อมูลสัญญา"
         Height          =   345
         Left            =   -74940
         Style           =   1  'Graphical
         TabIndex        =   95
         Top             =   7050
         Width           =   2205
      End
      Begin VB.TextBox txtRemdesc 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   -74220
         TabIndex        =   72
         Top             =   6000
         Width           =   5685
      End
      Begin VB.Frame Frame1 
         Caption         =   "เอกสารประกอบที่จำเป็น"
         Height          =   645
         Left            =   -74220
         TabIndex        =   93
         Top             =   5280
         Width           =   8775
         Begin VB.TextBox txtOtherDoc 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   6360
            TabIndex        =   71
            Top             =   240
            Width           =   1755
         End
         Begin VB.CheckBox chkOtherDoc 
            Caption         =   "อื่นๆ"
            Height          =   285
            Left            =   5700
            TabIndex        =   70
            Top             =   240
            Width           =   645
         End
         Begin VB.CheckBox chkQuotation 
            Caption         =   "ใบเสนอราคา"
            Height          =   285
            Left            =   4380
            TabIndex        =   69
            Top             =   240
            Width           =   1215
         End
         Begin VB.CheckBox chkContract 
            Caption         =   "สัญญา"
            Height          =   285
            Left            =   3600
            TabIndex        =   68
            Top             =   240
            Width           =   765
         End
         Begin VB.CheckBox chkBilling 
            Caption         =   "ใบรับวางบิล"
            Height          =   285
            Left            =   2340
            TabIndex        =   67
            Top             =   240
            Width           =   1125
         End
         Begin VB.CheckBox chkInv 
            Caption         =   "ใบกำกับภาษี"
            Height          =   285
            Left            =   960
            TabIndex        =   66
            Top             =   240
            Width           =   1215
         End
         Begin VB.CheckBox chkPO 
            Caption         =   "PO"
            Height          =   285
            Left            =   180
            TabIndex        =   65
            Top             =   240
            Width           =   585
         End
      End
      Begin VB.CommandButton Command3 
         Caption         =   "เพิ่มข้อมูลสถานที่วางบิล"
         Height          =   345
         Left            =   -74940
         Style           =   1  'Graphical
         TabIndex        =   74
         Top             =   7200
         Width           =   2205
      End
      Begin VB.CommandButton Command2 
         Caption         =   "เพิ่มข้อมูลสถานที่เก็บของเสีย"
         Height          =   345
         Left            =   -74940
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   7050
         Width           =   2205
      End
      Begin VB.TextBox txtParentCust 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         Enabled         =   0   'False
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   4395
         TabIndex        =   19
         Top             =   1545
         Width           =   3030
      End
      Begin VB.ComboBox cboCustType 
         Height          =   315
         ItemData        =   "frCustomer.frx":0154
         Left            =   1530
         List            =   "frCustomer.frx":0156
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   1560
         Width           =   1635
      End
      Begin MSComctlLib.ListView lvWaste 
         Height          =   2565
         Left            =   45
         TabIndex        =   20
         Top             =   4395
         Width           =   10725
         _ExtentX        =   18918
         _ExtentY        =   4524
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
            Text            =   "เลขที่ของเสีย"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ชื่อของเสีย"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Waste Code"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Waste Type"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "เลขที่สัญญา"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "วิธีการกำจัด"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "หมายเหตุ"
            Object.Width           =   2117
         EndProperty
      End
      Begin VB.TextBox txtFaxNo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   5100
         TabIndex        =   13
         Top             =   2265
         Width           =   2355
      End
      Begin VB.TextBox txtTelNo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1560
         TabIndex        =   12
         Top             =   2265
         Width           =   2655
      End
      Begin VB.TextBox txtDepartment 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   8400
         TabIndex        =   11
         Top             =   1905
         Width           =   2355
      End
      Begin VB.TextBox txtContactName 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   5100
         TabIndex        =   10
         Top             =   1905
         Width           =   2355
      End
      Begin VB.TextBox txtFactoryID 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1560
         TabIndex        =   9
         Top             =   1905
         Width           =   2655
      End
      Begin VB.ComboBox cboDistrict 
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   3345
         Width           =   1815
      End
      Begin VB.TextBox txtProvince 
         Appearance      =   0  'Flat
         BackColor       =   &H000000C0&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   5520
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   825
         Width           =   1275
      End
      Begin VB.TextBox txtAumphur 
         Appearance      =   0  'Flat
         BackColor       =   &H000000C0&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   3540
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   825
         Width           =   1335
      End
      Begin VB.TextBox txtTumbol 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1560
         TabIndex        =   3
         Top             =   825
         Width           =   1395
      End
      Begin VB.TextBox txtAddr 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   7380
         TabIndex        =   2
         Top             =   465
         Width           =   3375
      End
      Begin VB.TextBox txtCompanyNameTH 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1560
         TabIndex        =   1
         Top             =   465
         Width           =   4335
      End
      Begin MSComctlLib.ListView lvSiteAddress 
         Height          =   5775
         Left            =   -74940
         TabIndex        =   31
         Top             =   1260
         Width           =   10875
         _ExtentX        =   19182
         _ExtentY        =   10186
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
         Height          =   825
         Left            =   -74940
         TabIndex        =   73
         Top             =   6360
         Width           =   10875
         _ExtentX        =   19182
         _ExtentY        =   1455
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
         Height          =   5775
         Left            =   -74940
         TabIndex        =   96
         Top             =   1260
         Width           =   10875
         _ExtentX        =   19182
         _ExtentY        =   10186
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
         Height          =   5775
         Left            =   -74940
         TabIndex        =   98
         Top             =   1260
         Width           =   10875
         _ExtentX        =   19182
         _ExtentY        =   10186
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
            Object.Width           =   3528
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
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Email"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "เรื่องที่ติดต่อ"
            Object.Width           =   4411
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "ที่อยู่"
            Object.Width           =   4410
         EndProperty
      End
      Begin MSComctlLib.ListView lvCustRem 
         Height          =   5775
         Left            =   -74940
         TabIndex        =   117
         Top             =   1260
         Width           =   10875
         _ExtentX        =   19182
         _ExtentY        =   10186
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
      Begin VB.TextBox txtIndustryType 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   1200
         Width           =   9180
      End
      Begin VB.TextBox txtGenID 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   10215
         TabIndex        =   30
         Top             =   3720
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.CommandButton Command1 
         Caption         =   "เพิ่มข้อมูลกากของเสีย"
         Height          =   345
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   7005
         Width           =   1815
      End
      Begin VB.ComboBox cboCustGroup 
         Height          =   315
         ItemData        =   "frCustomer.frx":0158
         Left            =   2910
         List            =   "frCustomer.frx":015A
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   7035
         Width           =   2655
      End
      Begin MSComctlLib.ListView lvCommentQuo 
         Height          =   5175
         Left            =   -74520
         TabIndex        =   129
         Top             =   1560
         Width           =   9915
         _ExtentX        =   17489
         _ExtentY        =   9128
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
            Text            =   "เลขที่ข้อตกลง"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "หมายเหตุ"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "บริษัท"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "ID"
            Object.Width           =   0
         EndProperty
      End
      Begin MSComctlLib.ListView LvProb 
         Height          =   2475
         Left            =   -74910
         TabIndex        =   241
         Top             =   1425
         Width           =   10725
         _ExtentX        =   18918
         _ExtentY        =   4366
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
            Text            =   "ID"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "วันที่"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "เรื่อง"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "เรื่อง"
            Object.Width           =   5337
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "รายละเอียด"
            Object.Width           =   11466
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "วิธีการ"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "ผลลัพท์"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lvCustActivity 
         Height          =   5940
         Left            =   -74910
         TabIndex        =   254
         Top             =   1425
         Width           =   10725
         _ExtentX        =   18918
         _ExtentY        =   10478
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
            Text            =   "ID"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "CustID"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "วันที่"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "เรื่อง"
            Object.Width           =   5337
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "รายละเอียด"
            Object.Width           =   11466
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "ผลลัพธ์"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lvAuction 
         Height          =   5775
         Left            =   -74955
         TabIndex        =   284
         Top             =   450
         Width           =   10875
         _ExtentX        =   19182
         _ExtentY        =   10186
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
            Text            =   "วันที่ประมูล"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "วันที่สิ้นสุดสัญญา"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "เลขที่งานประมูล"
            Object.Width           =   2470
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "หมายเหตุ"
            Object.Width           =   5293
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "มูลค่า"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "ผลการประมูล"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "ประมูลครั้งต่อไป"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label lblเวปไซท์ 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "รูปถ่าย/วัน"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   64
         Left            =   4415
         TabIndex        =   287
         Top             =   3780
         Width           =   765
      End
      Begin VB.Label Label17 
         Caption         =   "VIP CR"
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
         Height          =   255
         Left            =   -74800
         TabIndex        =   285
         Top             =   500
         Width           =   1575
      End
      Begin VB.Label lblเวปไซท์ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ลูกค้าสัมพันธ์"
         ForeColor       =   &H000040C0&
         Height          =   195
         Index           =   57
         Left            =   540
         TabIndex        =   256
         Top             =   3780
         Width           =   900
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "วัน"
         Height          =   195
         Left            =   -66540
         TabIndex        =   252
         Top             =   6015
         Width           =   210
      End
      Begin VB.Label lblเวปไซท์ 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "แผนก"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   56
         Left            =   7695
         TabIndex        =   245
         Top             =   3045
         Width           =   675
      End
      Begin VB.Label lblเวปไซท์ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ประเภทอุตสาหกรรม"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   55
         Left            =   90
         TabIndex        =   244
         Top             =   1245
         Width           =   1410
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "วันที่เริ่มขน"
         Height          =   195
         Left            =   -67395
         TabIndex        =   236
         Top             =   1230
         Width           =   795
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "ลูกค้า Better ME รหัส"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   2
         Left            =   7470
         TabIndex        =   235
         Top             =   1635
         Width           =   1455
      End
      Begin VB.Label lblเวปไซท์ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "เคดิต เทอม"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   50
         Left            =   -68520
         TabIndex        =   232
         Top             =   6000
         Width           =   810
      End
      Begin VB.Label lblเวปไซท์ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ต้องการวาง Box"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   16
         Left            =   -74520
         TabIndex        =   185
         Top             =   465
         Width           =   1095
      End
      Begin VB.Label lblเวปไซท์ 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "หมายเหตุ ใบงาน :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   15
         Left            =   -71820
         TabIndex        =   181
         Top             =   6900
         Width           =   1275
      End
      Begin VB.Label Label14 
         Caption         =   "ลูกค้าอยู่ในช่วงระหว่างปรับใบเสนอราคา ปรับราคาใหม่"
         Height          =   255
         Left            =   -74160
         TabIndex        =   180
         Top             =   1200
         Width           =   4695
      End
      Begin VB.Label lblเวปไซท์ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ชื่อผู้ติดต่อ"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   39
         Left            =   -74700
         TabIndex        =   151
         Top             =   4860
         Width           =   705
      End
      Begin VB.Label lblเวปไซท์ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "แฟกซ์"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   38
         Left            =   -67980
         TabIndex        =   150
         Top             =   4860
         Width           =   450
      End
      Begin VB.Label lblเวปไซท์ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "โทร."
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   37
         Left            =   -71160
         TabIndex        =   149
         Top             =   4860
         Width           =   315
      End
      Begin VB.Label lblเวปไซท์ 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DTeam"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   36
         Left            =   7830
         TabIndex        =   139
         Top             =   2685
         Width           =   525
      End
      Begin VB.Label lblเวปไซท์ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ประเทศ"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   35
         Left            =   8760
         TabIndex        =   136
         Top             =   885
         Width           =   570
      End
      Begin VB.Label lbDate 
         Caption         =   "ระบุวันที่....."
         Height          =   255
         Left            =   9120
         TabIndex        =   126
         Top             =   7080
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "สาเหตุ"
         Height          =   255
         Left            =   5640
         TabIndex        =   125
         Top             =   7080
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label lblเวปไซท์ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "เซลล์ที่ขายจริง"
         ForeColor       =   &H000040C0&
         Height          =   195
         Index           =   24
         Left            =   7320
         TabIndex        =   123
         Top             =   3405
         Width           =   990
      End
      Begin VB.Label Label3 
         Caption         =   "ใส่รายละเอียดการขนส่ง : ข้อมูลที่บันทึกลงไปจะปรากฏที่หน้าปิด Job"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   -74700
         TabIndex        =   119
         Top             =   1140
         Width           =   5715
      End
      Begin VB.Label lblเวปไซท์ 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "หมายเหตุ"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   27
         Left            =   780
         TabIndex        =   105
         Top             =   3045
         Width           =   675
      End
      Begin VB.Label lblเวปไซท์ 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ชื่อย่อ"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   25
         Left            =   7680
         TabIndex        =   103
         Top             =   2325
         Width           =   675
      End
      Begin VB.Label lblเวปไซท์ 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "อีเมลล์"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   23
         Left            =   4320
         TabIndex        =   102
         Top             =   2685
         Width           =   720
      End
      Begin VB.Label lblเวปไซท์ 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "เวปไซท์"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   22
         Left            =   840
         TabIndex        =   101
         Top             =   2685
         Width           =   615
      End
      Begin VB.Label lblเวปไซท์ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "เซลล์"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   21
         Left            =   3480
         TabIndex        =   100
         Top             =   3405
         Width           =   480
      End
      Begin VB.Label lblเวปไซท์ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "รหัสไปรษณีย์"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   20
         Left            =   6900
         TabIndex        =   99
         Top             =   885
         Width           =   975
      End
      Begin VB.Label lblเวปไซท์ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "หมายเหตุ"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   19
         Left            =   -74940
         TabIndex        =   94
         Top             =   6060
         Width           =   675
      End
      Begin VB.Label lblเวปไซท์ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Direct Customer"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   14
         Left            =   3240
         TabIndex        =   92
         Top             =   1635
         Width           =   1155
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
         Height          =   180
         Index           =   2
         Left            =   0
         TabIndex        =   89
         Top             =   4170
         Width           =   10935
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   30
         Index           =   1
         Left            =   0
         Top             =   4110
         Width           =   10965
      End
      Begin VB.Label lblเวปไซท์ 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "แฟกซ์"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   11
         Left            =   4320
         TabIndex        =   87
         Top             =   2325
         Width           =   720
      End
      Begin VB.Label lblเวปไซท์ 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "โทร."
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   10
         Left            =   1140
         TabIndex        =   86
         Top             =   2325
         Width           =   315
      End
      Begin VB.Label lblเวปไซท์ 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "แผนก"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   9
         Left            =   7680
         TabIndex        =   85
         Top             =   1965
         Width           =   675
      End
      Begin VB.Label lblเวปไซท์ 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ชื่อผู้ติดต่อ"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   8
         Left            =   4320
         TabIndex        =   84
         Top             =   1965
         Width           =   720
      End
      Begin VB.Label lblเวปไซท์ 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "เลขทะเบียนโรงงาน"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   7
         Left            =   180
         TabIndex        =   83
         Top             =   1965
         Width           =   1275
      End
      Begin VB.Label lblเวปไซท์ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ประเภทอุตสาหกรรม"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   6
         Left            =   8760
         TabIndex        =   82
         Top             =   4005
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Label lblเวปไซท์ 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "เขตพื้นที่การขาย"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   5
         Left            =   300
         TabIndex        =   81
         Top             =   3405
         Width           =   1155
      End
      Begin VB.Label lblเวปไซท์ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "จังหวัด"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   4
         Left            =   4980
         TabIndex        =   80
         Top             =   885
         Width           =   495
      End
      Begin VB.Label lblเวปไซท์ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "อำเภอ"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   3
         Left            =   3060
         TabIndex        =   79
         Top             =   885
         Width           =   435
      End
      Begin VB.Label lblเวปไซท์ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ตำบล"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   1080
         TabIndex        =   78
         Top             =   885
         Width           =   375
      End
      Begin VB.Label lblเวปไซท์ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "เลขที่, หมู่, ถนน"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   6180
         TabIndex        =   77
         Top             =   525
         Width           =   1095
      End
      Begin VB.Label lblเวปไซท์ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ชื่อบริษัท/โรงงาน"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   270
         TabIndex        =   76
         Top             =   525
         Width           =   1215
      End
      Begin VB.Label lblเวปไซท์ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "กลุ่มลูกค้า"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   26
         Left            =   2130
         TabIndex        =   104
         Top             =   7095
         Width           =   675
      End
      Begin VB.Label lblเวปไซท์ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "เลขที่ผู้กำเนิดของเสีย"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   12
         Left            =   8685
         TabIndex        =   88
         Top             =   3690
         Visible         =   0   'False
         Width           =   1665
      End
   End
   Begin VB.Label lblLastUpdateTR 
      BackColor       =   &H00FFFFFF&
      Caption         =   "อัพเดทข้อมูลขนส่งโดย"
      Height          =   195
      Left            =   4740
      TabIndex        =   142
      Top             =   540
      Width           =   3315
   End
   Begin VB.Label lblLastUpdateAC 
      BackColor       =   &H00FFFFFF&
      Caption         =   "อัพเดทข้อมูลบัญชีโดย"
      Height          =   195
      Left            =   4740
      TabIndex        =   141
      Top             =   300
      Width           =   3315
   End
   Begin VB.Label lblLastUpdateCR 
      BackColor       =   &H00FFFFFF&
      Caption         =   "อัพเดทข้อมูลลูกค้าโดย"
      Height          =   195
      Left            =   4740
      TabIndex        =   140
      Top             =   60
      Width           =   3315
   End
   Begin VB.Label lblCancel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ยกเลิก"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   675
      Left            =   1500
      TabIndex        =   118
      Top             =   0
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label lblเวปไซท์ 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "รหัสลูกค้า"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   13
      Left            =   8310
      TabIndex        =   91
      Top             =   510
      Width           =   660
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
      Height          =   435
      Index           =   1
      Left            =   270
      TabIndex        =   75
      Top             =   60
      Width           =   1005
   End
   Begin VB.Image Image2 
      Height          =   795
      Left            =   0
      Picture         =   "frCustomer.frx":015C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11055
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
Dim colCR_ As Collection
Dim curParentCust_ As BWGCustomer
Dim curSaleStaff_ As BWGCompanyStaff
Dim MyCustomer_ As BWGCustomer
Dim callLocation_ As String
Dim colCustGroup_ As Collection
Dim colBillingInfoStaff_ As Collection
Dim CanAccess As Boolean, CanUpdate As Boolean

Private Sub LoadTR_Remark()
Dim rs As ADODB.Recordset
Dim i As Integer
Dim tSql As String

    tSql = "SELECT * FROM tbCustomer_TR_ChargeWType WHERE CustomerID = '" & MyCustomer_.id & "'"
    Set rs = New ADODB.Recordset
    Set rs = GetRS(tSql, adOpenForwardOnly, adLockReadOnly)
    If rs.RecordCount > 0 Then
        Select Case rs!ChargeWType
            Case "A"
                Option1.Value = True
                Option1.ForeColor = vbBlue
                Option2.ForeColor = vbBlack
                Option3.ForeColor = vbBlack
            Case "B"
                Option2.Value = True
                Option1.ForeColor = vbBlack
                Option2.ForeColor = vbBlue
                Option3.ForeColor = vbBlack
            Case Else
                Option3.Value = True
                Option1.ForeColor = vbBlack
                Option2.ForeColor = vbBlack
                Option3.ForeColor = vbBlue
        End Select
    Else
        Option3.Value = True
    End If

    ctlGridTR.getGridObj.Clear
    With ctlGridTR.getGridObj
        .Cols = 4
        .Rows = 50
        .ColWidth(0) = 0
        .ColWidth(1) = 0
        .ColWidth(2) = 0
        .ColWidth(3) = ctlGridTR.Width - 300
        .TextMatrix(0, 0) = ""
        .TextMatrix(0, 1) = ""
        .TextMatrix(0, 2) = ""
        .TextMatrix(0, 3) = "หมายเหตุ"
    End With
    tSql = "SELECT * FROM tbCustomer_TR_Remark" & vbCrLf
    tSql = tSql & "WHERE CustomerID = '" & MyCustomer_.id & "'" & vbCrLf
    tSql = tSql & "ORDER BY RemarkNo"
    Set rs = New ADODB.Recordset
    Set rs = GetRS(tSql, adOpenForwardOnly, adLockReadOnly)
    If rs.RecordCount > 0 Then
        With ctlGridTR.getGridObj
            For i = 1 To rs.RecordCount
                .TextMatrix(i, 0) = i
                .TextMatrix(i, 1) = rs!CustomerID
                .TextMatrix(i, 2) = rs!RemarkNo
                .TextMatrix(i, 3) = rs!RemarkText
                rs.MoveNext
            Next
        End With
    End If
End Sub

Private Sub AmdAdd_Click()

End Sub

Private Sub CbCredit_Click()
    If CbCredit.ListIndex > 0 Then
        txtCreditterm.Text = CbCredit.Text
        MyCustomer_.Creditterm = txtCreditterm.Text
    End If
End Sub

Private Sub cbCRStaff_Validate(Cancel As Boolean)
Dim tmpCRStaff As BWGCompanyStaff
    With MyCustomer_
        If .CustomerCode = "" Then
            If cbCRStaff.ListIndex > 0 Then
                Set tmpCRStaff = colCR_(cbCRStaff.ListIndex)
                .CRStaffID = tmpCRStaff.id
            Else
                .CRStaffID = ""
            End If
        Else
            If cbCRStaff.ListIndex > 0 Then
                Set tmpCRStaff = colCR_(cbCRStaff.ListIndex)
                .CRStaffID = tmpCRStaff.id
            End If
        End If
    End With

End Sub

Private Sub cbDept_Validate(Cancel As Boolean)
    With MyCustomer_
        If cbDept.ListIndex > 0 Then
            .CustdeptName = cbDept.Text
        End If
    End With
End Sub

Private Sub cboBillWeekCount_Validate(Cancel As Boolean)
    MyCustomer_.BillingInfo.BillWeekCount = cboBillWeekCount.ListIndex
    'If myCustomer_.BillingInfo.isChanged Then myCustomer_.isChanged = True
End Sub

Private Sub cboCheqWeekCount_Validate(Cancel As Boolean)
    MyCustomer_.BillingInfo.CheqWeekCount = cboCheqWeekCount.ListIndex
    'If myCustomer_.BillingInfo.isChanged Then myCustomer_.isChanged = True
End Sub

Private Sub cboCustGroup_Click()
If cboCustGroup.Text = "D - ลูกค้าที่หลุดแล้ว" Then
        Label4.Visible = True
        txtComment.Visible = True
        lbDate.Visible = True
        cmdDate.Visible = True
        cmdDate.Enabled = True
    Else
        If MyCustomer_ Is Nothing Then
            txtComment.Visible = False
            Label4.Visible = False
            lbDate.Visible = False
            cmdDate.Visible = False
            cmdDate.Enabled = False
            Exit Sub
        End If
        If Trim(MyCustomer_.Comment) <> "" Then
            txtComment.Text = "(เคยเป็นลูกค้าหลุด) " & Trim(MyCustomer_.Comment)
            txtComment.Visible = True
            Label4.Visible = True
            lbDate.Visible = True
            cmdDate.Visible = True
            cmdDate.Enabled = False
        Else
            txtComment.Visible = False
            Label4.Visible = False
            lbDate.Visible = False
            cmdDate.Visible = False
            cmdDate.Enabled = False
        End If
    End If
End Sub

Private Sub cboCustGroup_Validate(Cancel As Boolean)
    If cboCustGroup.ListIndex = -1 Then
        MyCustomer_.CustGroupID = ""
        MyCustomer_.CustGroupName = ""
    Else
        MyCustomer_.CustGroupID = colCustGroup_(cboCustGroup.ListIndex + 1).id
        MyCustomer_.CustGroupName = Trim(cboCustGroup.Text)
    End If
End Sub

Public Sub AddCommentDate()
    MyCustomer_.CommentDate = lbDate.Caption
End Sub

Private Sub cboCustType_Click()
On Error Resume Next
    txtParentCust.Enabled = (cboCustType.ListIndex = 2)
    txtParentCust.Enabled = (cboCustType.ListIndex = 4)
    txtParentCust.Enabled = (cboCustType.ListIndex = 6)
        If cboCustType.ListIndex = 3 Then
            txtParentCust.Text = "เบตเตอร์ มี จำกัด"
            txtParentCust.Enabled = True
            txtParentCust.SetFocus
            Else
            txtParentCust.Text = ""
        End If
        If cboCustType.ListIndex = 4 Then
            txtParentCust.Enabled = True
        Else
            txtParentCust.Text = ""
    End If
End Sub

Private Sub cboCustType_Validate(Cancel As Boolean)
    Select Case cboCustType.ListIndex
        Case 0
            MyCustomer_.CustomerType = ""
        Case 1
            MyCustomer_.CustomerType = "C"
        Case 2
            MyCustomer_.CustomerType = "S"
        Case 3
            MyCustomer_.CustomerType = "M"
        Case 4
            MyCustomer_.CustomerType = "MS"
        Case 5
            MyCustomer_.CustomerType = "B"
        Case 6
            MyCustomer_.CustomerType = "E"
    End Select
    MyCustomer_.isChanged = True
End Sub
Private Sub cboDistrict_Validate(Cancel As Boolean)
Dim tmpDistrict As BWGSaleDistrict
    With MyCustomer_
        If .CustomerCode = "" Then
            If cboDistrict.ListIndex > 0 Then
                Set tmpDistrict = colDistrict_(cboDistrict.ListIndex)
                .SaleDistrictID = tmpDistrict.id
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
        Else
            If cboDistrict.ListIndex > 0 Then
                Set tmpDistrict = colDistrict_(cboDistrict.ListIndex)
                .SaleDistrictID = tmpDistrict.id
                .SaleDistrictName = tmpDistrict.DistrictName
                .UnderSaleName = tmpDistrict.SaleStaffName
                .UnderSaleStaffID = tmpDistrict.SaleStaffID
                txtSaleName.Text = tmpDistrict.SaleStaffName
            End If
        End If
    End With
End Sub

Private Sub cboDTeam_Validate(Cancel As Boolean)
    If cboDTeam.ListIndex = -1 Then
        MyCustomer_.DTeam = False
    ElseIf cboDTeam.ListIndex = 0 Then
        MyCustomer_.DTeam = False
    ElseIf cboDTeam.ListIndex = 1 Then
        MyCustomer_.DTeam = True
    End If
End Sub

Private Sub cboRE_Men_Validate(Cancel As Boolean)
    If cboRE_Men.ListIndex = -1 Then
        MyCustomer_.BillingInfo.RE_ManID = ""
    Else
        MyCustomer_.BillingInfo.RE_ManID = colBillingInfoStaff_(cboRE_Men.ListIndex + 1).InFoStaffID
        MyCustomer_.isChanged = True
    End If
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
                MyCustomer_.BillingInfo.BillDay1 = tmpVal
            Case 2
                MyCustomer_.BillingInfo.BillDay2 = tmpVal
            Case 3
                MyCustomer_.BillingInfo.BillDay3 = tmpVal
            Case 4
                MyCustomer_.BillingInfo.BillDay4 = tmpVal
            Case 5
                MyCustomer_.BillingInfo.BillDay5 = tmpVal
        End Select
    Next
    'If myCustomer_.BillingInfo.isChanged Then myCustomer_.isChanged = True
End Sub

Private Sub chkBilling_Validate(Cancel As Boolean)
    MyCustomer_.BillingInfo.NeedBillingNote = chkBilling.Value = 1
    'If myCustomer_.BillingInfo.isChanged Then myCustomer_.isChanged = True
End Sub

Private Sub ChkBreak_Click()
    Dim Sql As String
    If ChkBreak.Value = 1 Then
        Sql = "update tbCustomer SET stoptrn =  '" & 1 & "'" & vbCrLf
        Sql = Sql & " WHERE CustomerID = '" & MyCustomer_.id & "'"
        DBConnExc Sql
    Else
        Sql = "update tbCustomer SET stoptrn =  '" & 0 & "'" & vbCrLf
        Sql = Sql & " WHERE CustomerID = '" & MyCustomer_.id & "'"
        DBConnExc Sql
    End If
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
                MyCustomer_.BillingInfo.CheqDay1 = tmpVal
            Case 2
                MyCustomer_.BillingInfo.CheqDay2 = tmpVal
            Case 3
                MyCustomer_.BillingInfo.CheqDay3 = tmpVal
            Case 4
                MyCustomer_.BillingInfo.CheqDay4 = tmpVal
            Case 5
                MyCustomer_.BillingInfo.CheqDay5 = tmpVal
        End Select
    Next
    'If myCustomer_.BillingInfo.isChanged Then myCustomer_.isChanged = True
End Sub

Private Sub chkContract_Validate(Cancel As Boolean)
    MyCustomer_.BillingInfo.NeedContract = chkContract.Value = 1
    'If myCustomer_.BillingInfo.isChanged Then myCustomer_.isChanged = True
End Sub

Private Sub ChkGPS_Click()
    With ChkGPS
        If .Value = 1 Then
            MyCustomer_.GPSRequest = "Y"
        Else
            MyCustomer_.GPSRequest = "N"
        End If
        MyCustomer_.isChanged = True
    End With
End Sub

Private Sub chkHaz_Validate(Cancel As Boolean)
    If chkHaz.Value = 1 Then
        MyCustomer_.isHaz = "Y"
    Else
        MyCustomer_.isHaz = "N"
    End If
    MyCustomer_.isChanged = True
End Sub

Private Sub chkInv_Validate(Cancel As Boolean)
    MyCustomer_.BillingInfo.NeedInvoice = chkInv.Value = 1
    'If myCustomer_.BillingInfo.isChanged Then myCustomer_.isChanged = True
End Sub

Private Sub chkisBox_Click()
    If chkisBox.Value = 1 Then
        MyCustomer_.isBox = "Y"
        FramBox.Enabled = True
        FramBox.BackColor = &HC0E0FF
    Else
        MyCustomer_.isBox = "N"
        FramBox.Enabled = False
        FramBox.BackColor = &H8000000F
    End If
End Sub

Private Sub chkisCheck_Click()
    If chkisCheck.Value = 1 Then
        MyCustomer_.ischeck = "Y"
        MyCustomer_.isChanged = True
    Else
       MyCustomer_.ischeck = "N"
       MyCustomer_.isChanged = True
    End If
End Sub

Private Sub chkisCheck_Validate(Cancel As Boolean)
    If chkisCheck.Value = 1 Then
        MyCustomer_.ischeck = "Y"
        MyCustomer_.isChanged = True
    Else
       MyCustomer_.ischeck = "N"
       MyCustomer_.isChanged = True
    End If
End Sub



Private Sub ChkISWeigth_Validate(Cancel As Boolean)
    If ChkISWeigth.Value = 1 Then
        MyCustomer_.isWeigth = "Y"
    Else
        MyCustomer_.isWeigth = "N"
    End If
       MyCustomer_.isChanged = True
End Sub

Private Sub chkNon_Validate(Cancel As Boolean)
    If chkNon.Value = 1 Then
        MyCustomer_.isNon = "Y"
    Else
        MyCustomer_.isNon = "N"
    End If
    MyCustomer_.isChanged = True
End Sub

Private Sub ChkOrther_Click()
    IsInv.Value = 0
    isInv_Men.Value = 0
    MyCustomer_.isChanged = True
End Sub

Private Sub ChkOrther_Validate(Cancel As Boolean)
    MyCustomer_.BillingInfo.Inv_State = "O"
    MyCustomer_.isChanged = True
End Sub

Private Sub ChkOvrW_Click()
    With ChkOvrW
        If .Value = 1 Then
            MyCustomer_.OverWeight = "Y"
        Else
            MyCustomer_.OverWeight = "N"
        End If
        MyCustomer_.isChanged = True
    End With
End Sub

Private Sub chkPO_Validate(Cancel As Boolean)
    MyCustomer_.BillingInfo.NeedPO = chkPO.Value = 1
    'If myCustomer_.BillingInfo.isChanged Then myCustomer_.isChanged = True
End Sub

Private Sub chkQuotation_Validate(Cancel As Boolean)
    MyCustomer_.BillingInfo.NeedQuotation = chkQuotation.Value = 1
    'If myCustomer_.BillingInfo.isChanged Then myCustomer_.isChanged = True
End Sub
Private Sub ChkVIPCR_Validate(Cancel As Boolean)
    If ChkVIPCR.Value = 1 Then
        MyCustomer_.isVIPCR = "Y"
        Label17.Visible = True
    Else
        MyCustomer_.isVIPCR = "N"
        Label17.Visible = False
    End If
        MyCustomer_.isChanged = True
End Sub

Sub PrintActivity()
Dim tmpSql As String
    With rpt1
        .LogonInfo(0) = "DSN = " & DSN_NAME & ";UID = " & strUid & ";PWD = " & strPwd
        .ReportFileName = App.Path & "\reports\RptCustActivity.rpt"
        tmpSql = ""
        tmpSql = "SELECT * FROM vw_CustActivity " & vbCrLf
        tmpSql = tmpSql & "WHERE CustomerID='" & MyCustomer_.id & "' "
        tmpSql = tmpSql & vbCrLf & "ORDER BY CreateDate"
'        .Formulas(0) = "FromDate=" & Chr(34) & "ข้อมูลวันที่ " & " " & dtFrom.ValueDMY & " " & " ถึงวันที่ " & " " & dtTo.ValueDMY & Chr(34)
'        .Formulas(0) = "PrintedDate=" & Chr(34) & ข้อมูลวันที่ & FormatYMD_to_DMY(TodayDate, "/", "/") & Chr(34)
        .SQLQuery = tmpSql
        .Destination = 0
        .WindowState = crptMaximized
        .PrintReport
        If Trim(.LastErrorString) <> "" Then MsgBox .LastErrorString
    End With
End Sub

Private Sub cmAuction_Click()
Dim selResult As Integer
    If Trim(txtCompanyNameTH.Text) = "" Then
        MsgBox "กรุณาระบุชื่อบริษัท/โรงงาน", vbExclamation
        SSTab1.Tab = 0
        txtCompanyNameTH.SetFocus
        Exit Sub
    End If
    If MyCustomer_.isChanged Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนสร้างใหม่หรือไม่?", vbQuestion + vbYesNo)
        If selResult = vbYes Then
            CmdSave_Click
            If MyCustomer_.isChanged Then Exit Sub
        Else 'If selResult = vbCancel Then
            Exit Sub
        End If
    End If
    Load frmAuction
    With frmAuction
        .SetCallerForm Me, MyCustomer_
        .Show 1
        ShowAuction
    End With

End Sub

Private Sub cmbPrintActivity_Click()
    PrintActivity
End Sub

Private Sub cmdAdd_Click()
    If MyCustomer_.id <> "" Then
        frmCustActivity.lblCID = MyCustomer_.id
        frmCustActivity.lblID.Caption = ""
        frmCustActivity.Show 1
    Else
        MsgBox "กรุณาค้นหาข้อมูลลูกค้าก่อน", vbCritical + vbOKOnly, "ผิดพลาด"
        Exit Sub
    End If
End Sub

Private Sub CmdAddProblem_Click()
    If txtCustCode.Text = "== AUTO ==" Then
        MsgBox "กรุณาเลือกลูกค้าที่ต้องการก่อนเพิ่มข้อร้องเรียน", vbInformation + vbOKOnly, "ไม่พบข้อมูลลูกค้าที่จะเพิ่มข้อร้องเรียน"
        Exit Sub
    End If
    FrmAddProb.ProbID.Caption = ""
    FrmAddProb.Show 1
End Sub

Private Sub cmdCancel_Click()
Dim tmpStr As String
    If MsgBox("ต้องการยกเลิกข้อมูลลูกค้ารายนี้หรือไม่?", vbQuestion + vbYesNo) = vbYes Then
        tmpStr = MyCustomer_.Cancel
        If Trim(tmpStr) = "" Then
            lblCancel.Caption = "ยกเลิก"
            lblCancel.Visible = True
            ClearScreen
            CreateNewCustomer
            txtCompanyNameTH.SetFocus
            lblCancel.Visible = True
        Else
            MsgBox tmpStr, vbExclamation
        End If
    End If
End Sub

Private Sub cmdDate_Click()
frmCommentDate.Show 1
End Sub

Private Sub cmdDelete_Click()
Dim tmpStr As String
    If MsgBox("ต้องการลบข้อมูลลูกค้านี้หรือไม่?", vbQuestion + vbYesNo) = vbYes Then
        tmpStr = MyCustomer_.Delete
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
Dim selResult As Integer
    If cmdSave.Enabled = False Then MyCustomer_.isChanged = False
    If MyCustomer_.isChanged Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนสร้างใหม่หรือไม่?", vbQuestion + vbYesNoCancel)
        If selResult = vbYes Then
            CmdSave_Click
            If MyCustomer_.isChanged Then Exit Sub
        ElseIf selResult = vbCancel Then
            Exit Sub
        End If
    End If
    ClearScreen
    CreateNewCustomer
    txtCompanyNameTH.SetFocus
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdOrderState_Click()
    If MyCustomer_ Is Nothing Then Exit Sub
    If MyCustomer_.CustomerName = "" Then Exit Sub
    
    Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmCustomerGroup", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    
    Load frmCustomerOrderState
    With frmCustomerOrderState
        .SetDataCustomer MyCustomer_, True
        .Show 1
    End With
End Sub

Private Sub CmdSave_Click()
    Dim tmpStr As String
    
        If Trim(txtCompanyNameTH.Text) = "" Then
        MsgBox "กรุณาระบุชื่อบริษัทหรือโรงงานของลูกค้า", vbExclamation
        txtCompanyNameTH.SetFocus
        Exit Sub
    End If
    If Trim(txtTumbol.Text) = "" Then
        MsgBox "กรุณาระบุตำบลที่ตั้งของลูกค้า", vbExclamation
        txtTumbol.SetFocus
        Exit Sub
    End If
    If Trim(cboCustGroup.Text) = "" Then
        MsgBox "กรุณาระบุกลุ่มลูกค้า", vbExclamation
        cboCustGroup.SetFocus
        Exit Sub
    End If
    If cboDistrict.ListIndex = 0 Then
        MsgBox "กรุณาระบุโซนพื้นที่ ที่ตั้งของลูกค้า", vbExclamation
        cboDistrict.SetFocus
        Exit Sub
    End If
    If cboCustType.ListIndex = 0 Then
        MsgBox "กรุณาระบุประเภทของลูกค้า", vbExclamation
        cboCustType.SetFocus
        Exit Sub
    End If
    If MyCustomer_.CustomerType = "S" And Trim(txtParentCust.Text) = "" Then
        MsgBox "กรุณาระบุ Direct Customer ที่บริษัทนี้เกี่ยวข้องอยู่", vbExclamation
        txtParentCust.SetFocus
        Exit Sub
    End If
    
    If MyCustomer_.BillingInfo.isChanged Then
        tmpStr = CurrentUser.VerifyAccessRole(frmCustomer.Name & "AC", UpdateRole)
        If Trim(tmpStr) <> "" Then
            tmpStr = "ข้อมูลด้านบัญชีถูกแก้ไข" & vbCrLf & tmpStr & " ข้อมูลด้านบัญชีไม่ถูกบันทึก"
            MsgBox tmpStr, vbExclamation
'            myCustomer_.BillingInfo.isChanged = False
        Else
            If Not MyCustomer_.BillingInfo Is Nothing Then MyCustomer_.BillingInfo.Save
            '*** ตรวจสอบว่า ถ้าฝ่ายบัญชีแก้ไขข้อมูลลูกค้า
            If MyCustomer_.isChanged Then
                tmpStr = CurrentUser.VerifyAccessRole(frmCustomer.Name & "CR", UpdateRole)
                If Trim(tmpStr) <> "" Then
                    tmpStr = "ข้อมูลลูกค้าถูกแก้ไข" & vbCrLf & tmpStr
                    MsgBox tmpStr, vbExclamation
                    Exit Sub
                End If
            End If
            
            Exit Sub
        End If
    End If
    
    If MyCustomer_.isChanged Then
        tmpStr = CurrentUser.VerifyAccessRole(frmCustomer.Name & "CR", UpdateRole)
        If Trim(tmpStr) <> "" Then
            tmpStr = "ข้อมูลลูกค้าถูกแก้ไข" & vbCrLf & tmpStr
            MsgBox tmpStr, vbExclamation
            Exit Sub
        End If
    Else
        MsgBox "ไม่มีการเปลี่ยนแปลงข้อมูลลูกค้า ข้อมูลลูกค้าไม่ถูกบันทึก", vbInformation
        MyCustomer_.isChanged = False
        Exit Sub
    End If
    
    If cboCustGroup.Text = "F - ลูกค้าที่หลุดแล้ว" Then
        If txtComment.Text = "" Or MyCustomer_.CommentDate = "" Then
            MsgBox " กรุณาระบุสาเหตุที่ลูกค้าหลุด และระบุวันที่ ในการบันทึกข้อมูลที่ลูกค้าหลุด", vbExclamation
            Exit Sub
        End If
    End If
    If MyCustomer_.isChanged Then
        MyCustomer_.Save
        If Trim(txtCustCode.Text) <> Trim(MyCustomer_.CustomerCode) Then
            ShowSiteAddress
            ShowBillAddress
        End If
        txtCustCode.Text = MyCustomer_.CustomerCode
    End If
        Dim OvW As String
        If ChkOvrW.Value = 1 Then
            OvW = "Y"
        Else
            OvW = "N"
        End If
        Dim Sql As String
    Sql = "update tbCustomer SET OverWeight =  '" & OvW & "'" & vbCrLf
    Sql = Sql & " WHERE CustomerID = '" & MyCustomer_.id & "'"
    DBConnExc Sql
    If MyCustomer_.isChanged = False Then MsgBox "จัดเก็บข้อมูลเรียบร้อยแล้ว", vbInformation
End Sub

Private Sub cmdSaveTR_Remark_Click()
    On Error Resume Next
    Dim rs As ADODB.Recordset
    Dim RSC As ADODB.Recordset
    Dim i As Integer, ChargeWType As String
    Dim tSql As String
    Dim Sql As String
    Dim tmpStr As String
    
    If MyCustomer_ Is Nothing Then Exit Sub
    tmpStr = CurrentUser.VerifyAccessRole(frmCustomer.Name & "TR", UpdateRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
        
    If Option1.Value = True Then
        ChargeWType = "A"
    ElseIf Option2.Value = True Then
        ChargeWType = "B"
    ElseIf Option3.Value = True Then
        ChargeWType = "-"
    Else
        ChargeWType = "-"
    End If
    tSql = "SELECT * FROM tbCustomer_TR_ChargeWType WHERE CustomerID = '" & MyCustomer_.id & "'"
    Set rs = New ADODB.Recordset
    Set rs = GetRS(tSql, adOpenForwardOnly, adLockReadOnly)
    If rs.RecordCount > 0 Then
        tSql = "UPDATE tbCustomer_TR_ChargeWType SET ChargeWType =  '" & ChargeWType & "'" & vbCrLf
        tSql = tSql & "WHERE CustomerID = '" & MyCustomer_.id & "'"
        DBConnExc tSql
    Else
        tSql = "INSERT INTO tbCustomer_TR_ChargeWType(CustomerID, ChargeWType)" & vbCrLf
        tSql = tSql & "VALUES ('" & MyCustomer_.id & "', '" & ChargeWType & "')" & vbCrLf
        DBConnExc tSql
    End If
    rs.Close
    Set rs = Nothing
    MyCustomer_.UpdateTR
    Dim GRQ As String
        If ChkGPS.Value = 1 Then
            GRQ = "Y"
        Else
            GRQ = "N"
        End If
    Sql = "update tbCustomer SET GPSRequest =  '" & GRQ & "'" & vbCrLf
    Sql = Sql & " WHERE CustomerID = '" & MyCustomer_.id & "'"
    DBConnExc Sql
    
    With ctlGridTR.getGridObj
        For i = 1 To .Rows - 1
            If .TextMatrix(i, 3) <> "" Then
                tSql = "SELECT * FROM tbCustomer_TR_Remark" & vbCrLf
                tSql = tSql & "WHERE CustomerID = '" & MyCustomer_.id & "'" & vbCrLf
                tSql = tSql & "AND RemarkNo = '" & .TextMatrix(i, 2) & "'"
                Set rs = New ADODB.Recordset
                Set rs = GetRS(tSql, adOpenForwardOnly, adLockReadOnly)
                If rs.RecordCount > 0 Then
                    tSql = "UPDATE tbCustomer_TR_Remark SET RemarkText =  '" & Replace(Trim(.TextMatrix(i, 3)), "'", "''") & "'" & vbCrLf
                    tSql = tSql & "WHERE CustomerID = '" & MyCustomer_.id & "'" & vbCrLf
                    tSql = tSql & "AND RemarkNo = '" & .TextMatrix(i, 2) & "'"
                    DBConnExc tSql
                Else
                    tSql = "INSERT INTO tbCustomer_TR_Remark(CustomerID, RemarkText)" & vbCrLf
                    tSql = tSql & "VALUES ('" & MyCustomer_.id & "', '" & Replace(Trim(.TextMatrix(i, 3)), "'", "''") & "')" & vbCrLf
                    DBConnExc tSql
                End If
                rs.Close
                Set rs = Nothing
            ElseIf .TextMatrix(i, 2) <> "" Then
                tSql = "DELETE FROM  tbCustomer_TR_Remark" & vbCrLf
                tSql = tSql & "WHERE CustomerID = '" & MyCustomer_.id & "'" & vbCrLf
                tSql = tSql & "AND RemarkNo = '" & .TextMatrix(i, 2) & "'"
                DBConnExc tSql
            End If
        Next
        MsgBox "บันทึกเสร็จแล้ว", vbInformation
        LoadTR_Remark
    End With
End Sub

Private Sub cmdSearch_Click()
On Error Resume Next
Dim selResult As Integer
    
    If MyCustomer_.isChanged Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนทำงานต่อหรือไม่?", vbQuestion + vbYesNoCancel)
        If selResult = vbYes Then
            CmdSave_Click
            If MyCustomer_.isChanged Then Exit Sub
        ElseIf selResult = vbCancel Then
            Exit Sub
        End If
    End If
    callLocation_ = "N" 'Normal
    frmCustomerSearch.Hide
    Load frmCustomerSearch
    With frmCustomerSearch
        .CStatus = "N"
        .SetCallerForm Me
        .Show 1
    End With
End Sub

Private Sub ComboInvMen_Validate(Cancel As Boolean)
  If ComboInvMen.ListIndex = -1 Then
        MyCustomer_.BillingInfo.Inv_ManID = ""
    Else
        MyCustomer_.BillingInfo.Inv_ManID = colBillingInfoStaff_(ComboInvMen.ListIndex + 1).InFoStaffID
        MyCustomer_.isChanged = True
    End If
End Sub

Private Sub Command1_Click()
Dim selResult As Integer
    If Trim(txtCompanyNameTH.Text) = "" Then
        MsgBox "กรุณาระบุชื่อบริษัท/โรงงาน", vbExclamation
        txtCompanyNameTH.SetFocus
        Exit Sub
    End If
    If MyCustomer_.isChanged Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนสร้างใหม่หรือไม่?", vbQuestion + vbYesNo)
        If selResult = vbYes Then
            CmdSave_Click
            If MyCustomer_.isChanged Then Exit Sub
        Else 'If selResult = vbCancel Then
            Exit Sub
        End If
    End If
    Load frmWasteData
    With frmWasteData
        .SetCallerForm Me, MyCustomer_
        .Show 1
        MyCustomer_.RequeryWasteItems
        ShowWasteData
    End With
End Sub

Private Sub Command2_Click()
Dim selResult As Integer
    If Trim(txtCompanyNameTH.Text) = "" Then
        MsgBox "กรุณาระบุชื่อบริษัท/โรงงาน", vbExclamation
        SSTab1.Tab = 0
        txtCompanyNameTH.SetFocus
        Exit Sub
    End If
    If MyCustomer_.isChanged Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนสร้างใหม่หรือไม่?", vbQuestion + vbYesNo)
        If selResult = vbYes Then
            CmdSave_Click
            If MyCustomer_.isChanged Then Exit Sub
        Else 'If selResult = vbCancel Then
            Exit Sub
        End If
    End If
    Load frmCustomerSite
    With frmCustomerSite
        .SetCallerForm Me, MyCustomer_
        .Show 1
        MyCustomer_.RequerySiteAddress
        ShowSiteAddress
    End With
End Sub

Private Sub Command3_Click()
Dim selResult As Integer
    If Trim(txtCompanyNameTH.Text) = "" Then
        MsgBox "กรุณาระบุชื่อบริษัท/โรงงาน", vbExclamation
        SSTab1.Tab = 0
        txtCompanyNameTH.SetFocus
        Exit Sub
    End If
    If MyCustomer_.isChanged Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนสร้างใหม่หรือไม่?", vbQuestion + vbYesNo)
        If selResult = vbYes Then
            CmdSave_Click
            If MyCustomer_.isChanged Then Exit Sub
        Else 'If selResult = vbCancel Then
            Exit Sub
        End If
    End If
    Load frmCustomerBillAddr
    With frmCustomerBillAddr
        .SetCallerForm Me, MyCustomer_
        .Show 1
        MyCustomer_.RequeryBillAddress
        ShowBillAddress
    End With
End Sub

Private Sub Command4_Click()
Dim selResult As Integer
    If Trim(txtCompanyNameTH.Text) = "" Then
        MsgBox "กรุณาระบุชื่อบริษัท/โรงงาน", vbExclamation
        SSTab1.Tab = 0
        txtCompanyNameTH.SetFocus
        Exit Sub
    End If
    If MyCustomer_.isChanged Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนสร้างใหม่หรือไม่?", vbQuestion + vbYesNo)
        If selResult = vbYes Then
            CmdSave_Click
            If MyCustomer_.isChanged Then Exit Sub
        Else 'If selResult = vbCancel Then
            Exit Sub
        End If
    End If
    Load frmContract
    With frmContract
        .SetCallerForm Me, MyCustomer_
        .Show 1
        MyCustomer_.RequeryContracts
        ShowContracts
    End With
End Sub

Private Sub Command5_Click()
Dim selResult As Integer
    If Trim(txtCompanyNameTH.Text) = "" Then
        MsgBox "กรุณาระบุชื่อบริษัท/โรงงาน", vbExclamation
        SSTab1.Tab = 0
        txtCompanyNameTH.SetFocus
        Exit Sub
    End If
    If MyCustomer_.isChanged Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนสร้างใหม่หรือไม่?", vbQuestion + vbYesNo)
        If selResult = vbYes Then
            CmdSave_Click
            If MyCustomer_.isChanged Then Exit Sub
        Else 'If selResult = vbCancel Then
            Exit Sub
        End If
    End If
    Load frmCustomerStaff
    With frmCustomerStaff
        .SetCallerForm Me, MyCustomer_
        .Show 1
        MyCustomer_.RequeryCustStaff
        ShowCustStaff
    End With
End Sub

Private Sub Command6_Click()
Dim selResult As Integer
    If Trim(txtCompanyNameTH.Text) = "" Then
        MsgBox "กรุณาระบุชื่อบริษัท/โรงงาน", vbExclamation
        SSTab1.Tab = 0
        txtCompanyNameTH.SetFocus
        Exit Sub
    End If
    If MyCustomer_.isChanged Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนสร้างใหม่หรือไม่?", vbQuestion + vbYesNo)
        If selResult = vbYes Then
            CmdSave_Click
            If MyCustomer_.isChanged Then Exit Sub
        Else 'If selResult = vbCancel Then
            Exit Sub
        End If
    End If
    Load frmCustOtherRemark
    With frmCustOtherRemark
        .SetCallerForm Me, MyCustomer_
        .Show 1
        MyCustomer_.RequeryOtherRemark
        ShowOtherRemark
    End With
End Sub

Private Sub Command7_Click()
Dim selResult As Integer
    If Trim(txtCompanyNameTH.Text) = "" Then
        MsgBox "กรุณาระบุชื่อบริษัท/โรงงาน", vbExclamation
        SSTab1.Tab = 0
        txtCompanyNameTH.SetFocus
        Exit Sub
    End If
    Load frmCommentQuo
    With frmCommentQuo
        .SetCallerForm Me, MyCustomer_
        .Show 1
        MyCustomer_.RequeryCmmQuo
        ShowCmmQuo
    End With
End Sub

Private Sub ctlGridTR_KeyDown(KeyCode As Integer, Shift As Integer)
Dim tSql As String
    If KeyCode = vbKeyDelete Then
        If MsgBox("ยืนยันการลบข้อมูล", vbInformation + vbOKCancel) = vbCancel Then Exit Sub
        If ctlGridTR.getGridObj.TextMatrix(ctlGridTR.getGridObj.row, 2) <> "" Then
            tSql = "DELETE FROM  tbCustomer_TR_Remark" & vbCrLf
            tSql = tSql & "WHERE CustomerID = '" & MyCustomer_.id & "'" & vbCrLf
            tSql = tSql & "AND RemarkNo = '" & ctlGridTR.getGridObj.TextMatrix(ctlGridTR.getGridObj.row, 2) & "'"
            DBConnExc tSql
            ctlGridTR.getGridObj.RemoveItem ctlGridTR.getGridObj.row
        End If
    End If
End Sub
Private Sub dtBoutCageBack_Validate(Cancel As Boolean)
    MyCustomer_.BoutCageTo = dtBoutCageBack.ValueYMD
    MyCustomer_.isChanged = True
End Sub

Private Sub dtBoutCageFrom_Validate(Cancel As Boolean)
    MyCustomer_.BoutCageFrom = dtBoutCageFrom.ValueYMD
    MyCustomer_.isChanged = True
End Sub

Private Sub dtBoutSBack_Validate(Cancel As Boolean)
    MyCustomer_.BoutSTo = dtBoutSBack.ValueYMD
    MyCustomer_.isChanged = True
End Sub

Private Sub dtBoutSFrom_Validate(Cancel As Boolean)
    MyCustomer_.BoutSFrom = dtBoutSFrom.ValueYMD
    MyCustomer_.isChanged = True
End Sub
Private Sub dtLackBack_Validate(Cancel As Boolean)
    MyCustomer_.LackTo = dtLackBack.ValueYMD
    MyCustomer_.isChanged = True
End Sub
Private Sub dtLackFrom_Validate(Cancel As Boolean)
    MyCustomer_.LackFrom = dtLackFrom.ValueYMD
    MyCustomer_.isChanged = True
End Sub

Private Sub dtLugBackdate_Validate(Cancel As Boolean)
    MyCustomer_.LUGBOXBack = dtLugBackdate.ValueYMD
    MyCustomer_.isChanged = True
End Sub

Private Sub dtLugFromdate_Validate(Cancel As Boolean)
    MyCustomer_.LUGBOXFrom = dtLugFromdate.ValueYMD
            MyCustomer_.isChanged = True
End Sub

Private Sub dtOneTBack_Validate(Cancel As Boolean)
    MyCustomer_.OneTBack = dtOneTBack.ValueYMD
        MyCustomer_.isChanged = True
End Sub

Private Sub dtOneTFrom_Validate(Cancel As Boolean)
    MyCustomer_.OneTFrom = dtOneTFrom.ValueYMD
    MyCustomer_.isChanged = True
End Sub

Private Sub dtROLBackdate_Validate(Cancel As Boolean)
    MyCustomer_.ROLBOXBack = dtROLBackdate.ValueYMD
    MyCustomer_.isChanged = True
End Sub

Private Sub dtROLFromdate_Validate(Cancel As Boolean)
    MyCustomer_.ROLBOXFrom = dtROLFromdate.ValueYMD
    MyCustomer_.isChanged = True
End Sub

Private Sub dtTruckBack_Validate(Cancel As Boolean)
    MyCustomer_.TruckBOXBack = dtTruckBack.ValueYMD
    MyCustomer_.isChanged = True
End Sub

Private Sub dtTruckFrom_Validate(Cancel As Boolean)
    MyCustomer_.TwoTFrom = dtTwoTFrom.ValueYMD
    MyCustomer_.isChanged = True
End Sub

Private Sub dtTwoTBack_Validate(Cancel As Boolean)
    MyCustomer_.TwoTBack = dtTwoTBack.ValueYMD
        MyCustomer_.isChanged = True
End Sub

Private Sub dtTwoTFrom_Validate(Cancel As Boolean)
    MyCustomer_.TwoTFrom = dtTwoTFrom.ValueYMD
    MyCustomer_.isChanged = True
End Sub
Private Sub dtWaterBoutBack_Validate(Cancel As Boolean)
    MyCustomer_.WaterBoutTo = dtWaterBoutBack.ValueYMD
    MyCustomer_.isChanged = True
End Sub

Private Sub dtWaterBoutFrom_Validate(Cancel As Boolean)
    MyCustomer_.WaterBoutFrom = dtWaterBoutFrom.ValueYMD
    MyCustomer_.isChanged = True
End Sub

Private Sub Form_Load()
    CanAccess = False: CanUpdate = False
    SetPermission Me, CanAccess, CanUpdate
    If CanAccess = False Then
        MsgBox "ท่านไม่มีสิทธิใช้ระบบนี้", vbExclamation, "Warning"
        Unload Me
    End If
        If CurrentUser.StaffCode <> "ACC" Then
            ChkBreak.Visible = False
        End If
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

    SSTab1.Tab = 0
    With frmPicture
        cmdSave.Picture = .picSave.Picture
        cmdSearch.Picture = .picSearch.Picture
        cmdClear.Picture = .picNew.Picture
        cmdCancel.Picture = .picDelete.Picture
        cmdClose.Picture = .picExit.Picture
        cmdOrderState.Picture = .picApprove.Picture
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
        .ListIndex = 0
    End With
    With cboDTeam
        .AddItem "Normal"
        .AddItem "DTeam"
    End With
    Set curTumbol_ = Nothing
    Set curTumbol_ = New IMITumbol
    Set curIndustryType_ = Nothing
    Set curIndustryType_ = New BWGIndustryType_New
    Set curParentCust_ = New BWGCustomer
    PopulateDistrict
    PopulateCR
    PopulateCustGroup
    PopulateBillInfo
    CreateNewCustomer
'    InitSaleAP
    Dim Sql As String
    Dim rs As New ADODB.Recordset
    Dim i As Integer
    Sql = "SELECT * FROM tbCreditTerm  " & vbCrLf
    Set rs = GetRS(Sql, adOpenForwardOnly, adLockOptimistic)
    CbCredit.Clear
    If rs.RecordCount > 0 Then
        CbCredit.AddItem "==เลือก=="
            For i = 1 To rs.RecordCount
                CbCredit.AddItem rs!CDTerm
            rs.MoveNext
            Next
            CbCredit.ListIndex = 0
    End If

    With ctlGridTR.getGridObj
        .Cols = 4
        .Rows = 50
        .ColWidth(0) = 0
        .ColWidth(1) = 0
        .ColWidth(2) = 0
        .ColWidth(3) = ctlGridTR.Width - 300
        .TextMatrix(0, 0) = ""
        .TextMatrix(0, 1) = ""
        .TextMatrix(0, 2) = ""
        .TextMatrix(0, 3) = "หมายเหตุ"
    End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim selResult As Integer
    If MyCustomer_ Is Nothing Then Exit Sub
    If CanUpdate = True Then
        If cmdSave.Enabled = False Then Exit Sub
        If MyCustomer_.isChanged Then
            selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนปิดหน้าจอหรือไม่?", vbQuestion + vbYesNoCancel)
            If selResult = vbYes Then
                CmdSave_Click
                If MyCustomer_.isChanged Then Cancel = 1
            ElseIf selResult = vbCancel Then
                Cancel = 1
            End If
        End If
    End If
End Sub
Private Sub IsInv_Click()
  ChkOrther.Value = 0
  isInv_Men.Value = 0
  MyCustomer_.isChanged = True
End Sub

Private Sub isInv_Men_Click()
  ChkOrther.Value = 0
  IsInv.Value = 0
  MyCustomer_.isChanged = True
End Sub

Private Sub isInv_Men_Validate(Cancel As Boolean)
 If isInv_Men.Value = 1 Then
        MyCustomer_.BillingInfo.Inv_State = "M"
    End If
End Sub

Private Sub isRE_CB_Validate(Cancel As Boolean)
    If isRE_CB.Value = 1 Then
        MyCustomer_.BillingInfo.RE_State = "S"
        MyCustomer_.isChanged = True
    End If
End Sub

Private Sub isRV_EMS_Validate(Cancel As Boolean)
    If isRV_EMS.Value = 1 Then
        MyCustomer_.BillingInfo.RE_State = "E"
        MyCustomer_.isChanged = True
    Else
       isRV_Bank.Value = 0
       isRV_Inbank.Value = 0
       isRE_CB.Value = 0
    End If
End Sub

Private Sub isRV_Inbank_Validate(Cancel As Boolean)
    If isRV_Inbank.Value = 1 Then
        MyCustomer_.BillingInfo.RE_State = "B"
        MyCustomer_.isChanged = True
    Else
      isRV_Bank.Value = 0
      isRV_EMS.Value = 0
      isRE_CB.Value = 0
    End If
End Sub
Private Sub lvAuction_DblClick()
    If lvAuction.SelectedItem Is Nothing Then Exit Sub
    Load frmAuction
    With frmAuction
        .SetCallerForm Me, MyCustomer_
'        .SetDataAuction MyCustomer_.Auction(lvAuction.SelectedItem.Key)
        .txtContractNo.Text = lvAuction.SelectedItem.SubItems(2)
        .dtStart.ValueDMY = lvAuction.SelectedItem
        If lvAuction.SelectedItem.SubItems(1) <> "" Then
            .dtEnd.ValueDMY = lvAuction.SelectedItem.SubItems(1)
        End If
        .txtRemDesc.Text = lvAuction.SelectedItem.SubItems(3)
        .txtPrice.Text = lvAuction.SelectedItem.SubItems(4)
        If lvAuction.SelectedItem.SubItems(5) = "ได้งาน" Then
            .IsOK.Value = True
        ElseIf lvAuction.SelectedItem.SubItems(5) = "ไม่ได้งาน" Then
            .IsNoK.Value = True
        End If
        .lblAuctionID.Caption = lvAuction.SelectedItem.Key
        .Show 1
        ShowAuction

        End With
End Sub

Private Sub lvBillAddr_DblClick()
    If lvBillAddr.SelectedItem Is Nothing Then Exit Sub
    Load frmCustomerBillAddr
    With frmCustomerBillAddr
        .SetCallerForm Me, MyCustomer_
        .SetDataBillingAddr MyCustomer_.BillAddresses(lvBillAddr.SelectedItem.Key)
        .Show 1
        MyCustomer_.RequeryBillAddress
        ShowBillAddress
    End With
End Sub

Private Sub lvContract_DblClick()
    If lvContract.SelectedItem Is Nothing Then Exit Sub
    Load frmContract
    With frmContract
        .SetCallerForm Me, MyCustomer_
        .SetDataContract MyCustomer_.Contracts(lvContract.SelectedItem.Key)
        .Show 1
        MyCustomer_.RequeryContracts
        ShowContracts
    End With
End Sub
Private Sub lvCommentQuo_DblClick()
    If lvCommentQuo.SelectedItem Is Nothing Then Exit Sub
    Load frmCommentQuo
    With frmCommentQuo
        .SetCallerForm Me, MyCustomer_
        .SetDataCmmQuotation MyCustomer_.CmmQuo(lvCommentQuo.SelectedItem.Key)
        .Show 1
        MyCustomer_.RequeryCmmQuo
        ShowCmmQuo
    End With
End Sub
Private Sub lvCustActivity_DblClick()
    If lvCustActivity.SelectedItem Is Nothing Then Exit Sub
    Load frmCustActivity
    With frmCustActivity
        .SetData (lvCustActivity.SelectedItem.Key)
        .Show 1
        MyCustomer_.RequeryWasteItems
        ShowWasteData
    End With
End Sub

Private Sub lvCustRem_DblClick()
    If lvCustRem.SelectedItem Is Nothing Then Exit Sub
    Load frmCustOtherRemark
    With frmCustOtherRemark
        .SetCallerForm Me, MyCustomer_
        .SetDataBillingAddr MyCustomer_.OtherRemarks(lvCustRem.SelectedItem.Key)
        .Show 1
        MyCustomer_.RequeryOtherRemark
        ShowOtherRemark
    End With
End Sub


Private Sub LvProb_DblClick()
    If LvProb.SelectedItem Is Nothing Then Exit Sub
    Load FrmAddProb
    With FrmAddProb
        .ProbID.Caption = (LvProb.SelectedItem.Key)
        .dtdate.ValueDMY = Format(LvProb.SelectedItem.SubItems(2), "dd/MM/yyyy")
        .TxtName.Text = LvProb.SelectedItem.SubItems(3)
        .TxtDet.Text = LvProb.SelectedItem.SubItems(4)
        .Show 1
        MyCustomer_.RequeryWasteItems
        ShowWasteData
    End With

End Sub

Private Sub lvSiteAddress_DblClick()
    If lvSiteAddress.SelectedItem Is Nothing Then Exit Sub
    Load frmCustomerSite
    With frmCustomerSite
        .SetCallerForm Me, MyCustomer_
        .SetDataCustomerSite MyCustomer_.SiteAddresses(lvSiteAddress.SelectedItem.Key)
        .Show 1
        MyCustomer_.RequerySiteAddress
        ShowSiteAddress
    End With
End Sub

Private Sub lvStaff_DblClick()
    If lvStaff.SelectedItem Is Nothing Then Exit Sub
    Load frmCustomerStaff
    With frmCustomerStaff
        .SetCallerForm Me, MyCustomer_
        .SetDataCustomerStaff MyCustomer_.CustStaffs(lvStaff.SelectedItem.Key)
        .Show 1
        MyCustomer_.RequeryCustStaff
        ShowCustStaff
    End With
End Sub

Private Sub lvWaste_DblClick()
    If lvWaste.SelectedItem Is Nothing Then Exit Sub
    Load frmWasteData
    With frmWasteData
        .SetCallerForm Me, MyCustomer_
        .SetDataWasteData MyCustomer_.WasteItems(lvWaste.SelectedItem.Key)
        .txtIndustryType.Text = txtIndustryType.Text
        .Show 1
        MyCustomer_.RequeryWasteItems
        ShowWasteData
    End With
End Sub

Private Sub optBillType_Validate(Index As Integer, Cancel As Boolean)
    MyCustomer_.BillingInfo.BillRecType = Index
    'If myCustomer_.BillingInfo.isChanged Then myCustomer_.isChanged = True
End Sub

Private Sub optCheqType_Validate(Index As Integer, Cancel As Boolean)
    MyCustomer_.BillingInfo.CheqRecType = Index
    'If myCustomer_.BillingInfo.isChanged Then myCustomer_.isChanged = True
End Sub

Private Sub Option1_Click()
    Option1.ForeColor = vbBlue
    Option2.ForeColor = vbBlack
    Option3.ForeColor = vbBlack
End Sub

Private Sub Option2_Click()
    Option1.ForeColor = vbBlack
    Option2.ForeColor = vbBlue
    Option3.ForeColor = vbBlack
End Sub

Private Sub Option3_Click()
    Option1.ForeColor = vbBlack
    Option2.ForeColor = vbBlack
    Option3.ForeColor = vbBlue
End Sub

Private Sub isRV_Bank_Validate(Cancel As Boolean)
    If isRV_Bank.Value = 1 Then
       MyCustomer_.BillingInfo.RE_State = "C"
       MyCustomer_.isChanged = True
    Else
       isRV_Inbank.Value = 0
       isRV_EMS.Value = 0
       isRE_CB.Value = 0
    End If
End Sub

Private Sub txtAccContact_Validate(Index As Integer, Cancel As Boolean)
    MyCustomer_.BillingInfo.ContactName = Trim(txtInv_Contact.Text)
    If MyCustomer_.BillingInfo.isChanged Then MyCustomer_.isChanged = True
End Sub

Private Sub txtAccFax_Validate(Index As Integer, Cancel As Boolean)
    MyCustomer_.BillingInfo.TelNo = Trim(txtInv_Fax.Text)
    If MyCustomer_.BillingInfo.isChanged Then MyCustomer_.isChanged = True
End Sub

Private Sub txtAccTel_Validate(Index As Integer, Cancel As Boolean)
    MyCustomer_.BillingInfo.TelNo = Trim(txtInv_tel.Text)
    If MyCustomer_.BillingInfo.isChanged Then MyCustomer_.isChanged = True
End Sub

Private Sub Option4_Validate(Cancel As Boolean)
    If Option4.Value = True Then
        MyCustomer_.BillingInfo.Inv_State = "O"
        MyCustomer_.isChanged = True
    End If
End Sub

Private Sub Option5_Validate(Cancel As Boolean)
    If Option5.Value = True Then
        MyCustomer_.BillingInfo.Inv_State = "S"
        MyCustomer_.isChanged = True
    End If
End Sub

Private Sub Option6_Validate(Cancel As Boolean)
 If Option6.Value = True Then
        MyCustomer_.BillingInfo.Inv_State = "M"
        MyCustomer_.isChanged = True
    End If
End Sub

Private Sub txtAddr_Validate(Cancel As Boolean)
    MyCustomer_.Address = Trim(txtAddr.Text)
End Sub
Private Sub TxtBeginDate_Validate(Cancel As Boolean)
    MyCustomer_.BeginDate = Trim(TxtBeginDate.Text)
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
            MyCustomer_.BillingInfo.BillDay1 = tmpVal
        Case 2
            MyCustomer_.BillingInfo.BillDay2 = tmpVal
        Case 3
            MyCustomer_.BillingInfo.BillDay3 = tmpVal
        Case 4
            MyCustomer_.BillingInfo.BillDay4 = tmpVal
        Case 5
            MyCustomer_.BillingInfo.BillDay5 = tmpVal
    End Select
    'If myCustomer_.BillingInfo.isChanged Then myCustomer_.isChanged = True
End Sub

Private Sub txtBillingDet_Validate(Cancel As Boolean)
    MyCustomer_.BillingInfo.BillingDetail = Trim(txtBillingDet.Text)
    'If myCustomer_.BillingInfo.isChanged Then myCustomer_.isChanged = True
End Sub
Private Sub TxtBoutCage_Validate(Cancel As Boolean)
   If Not IsNumeric(txtBoutCage.Text) Then
        MsgBox "ตัวเลขจำนวนเต็มเท่านั้น"
        txtBoutCage.Text = ""
    Else
        MyCustomer_.BoutCage = txtBoutCage.Text
    End If
        MyCustomer_.isChanged = True
End Sub
Private Sub txtBoutCageW_Validate(Cancel As Boolean)
        MyCustomer_.BoutCageW = txtBoutCageW.Text
        MyCustomer_.isChanged = True
End Sub

Private Sub TxtBoutS_Validate(Cancel As Boolean)
   If Not IsNumeric(txtBoutS.Text) Then
        MsgBox "ตัวเลขจำนวนเต็มเท่านั้น"
        txtBoutS.Text = ""
    Else
        MyCustomer_.BoutS = txtBoutS.Text
    End If
        MyCustomer_.isChanged = True
End Sub
Private Sub txtBoutSBack_Validate(Cancel As Boolean)
        MyCustomer_.BoutSBack = txtBoutSBack.Text
        MyCustomer_.isChanged = True
End Sub

Private Sub txtBoxDesc_Validate(Cancel As Boolean)
    MyCustomer_.BoxDesc = txtBoxDesc.Text
    MyCustomer_.isChanged = True
End Sub

Private Sub txtBoxOneW_Validate(Cancel As Boolean)
'    If Not IsNumeric(txtBoxOneW.Text) Then
'        MsgBox "ตัวเลขจำนวนเต็มเท่านั้น"
'        txtBoxOneW.Text = ""
'    Else
        MyCustomer_.OneTBoxW = txtBoxOneW.Text
'    End If
        MyCustomer_.isChanged = True
End Sub

Private Sub txtBoxTwoW_Validate(Cancel As Boolean)
'    If Not IsNumeric(txtBoxTwoW.Text) Then
'        MsgBox "ตัวเลขจำนวนเต็มเท่านั้น"
'        txtBoxTwoW.Text = ""
'    Else
        MyCustomer_.TwoTBoxW = txtBoxTwoW.Text
'    End If
        MyCustomer_.isChanged = True
End Sub
Private Sub txtBuotCageBack_Validate(Cancel As Boolean)
        MyCustomer_.BoutCageBack = txtBuotCageBack.Text
        MyCustomer_.isChanged = True
End Sub

Private Sub txtBuotSW_Validate(Cancel As Boolean)
        MyCustomer_.BoutSW = txtBuotSW.Text
        MyCustomer_.isChanged = True
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
            MyCustomer_.BillingInfo.CheqDay1 = tmpVal
        Case 2
            MyCustomer_.BillingInfo.CheqDay2 = tmpVal
        Case 3
            MyCustomer_.BillingInfo.CheqDay3 = tmpVal
        Case 4
            MyCustomer_.BillingInfo.CheqDay4 = tmpVal
        Case 5
            MyCustomer_.BillingInfo.CheqDay5 = tmpVal
    End Select
    'If myCustomer_.BillingInfo.isChanged Then myCustomer_.isChanged = True
End Sub
Private Sub IsInv_Validate(Cancel As Boolean)
    If IsInv.Value = 1 Then
        MyCustomer_.BillingInfo.Inv_State = "S"
    End If
End Sub
Private Sub txtCheqDetail_Validate(Cancel As Boolean)
    MyCustomer_.BillingInfo.CheqDetail = Trim(txtCheqDetail.Text)
    'If myCustomer_.BillingInfo.isChanged Then myCustomer_.isChanged = True
End Sub

Private Sub txtComment_Validate(Cancel As Boolean)
    MyCustomer_.Comment = Trim(txtComment.Text)
End Sub
Private Sub txtCompanyNameTH_Validate(Cancel As Boolean)
    MyCustomer_.CustomerName = Trim(txtCompanyNameTH.Text)
End Sub

Private Sub txtContactName_Validate(Cancel As Boolean)
    MyCustomer_.ContactPerson = txtContactName.Text
End Sub

Private Sub txtCountry_Validate(Cancel As Boolean)
    MyCustomer_.Country = Trim(txtCountry.Text)
End Sub

Private Sub txtCreditterm_Validate(Cancel As Boolean)
    MyCustomer_.Creditterm = txtCreditterm.Text
End Sub

Private Sub txtCustRemark_Validate(Cancel As Boolean)
    MyCustomer_.CustRemark = txtCustRemark.Text
    MyCustomer_.isChanged = True
End Sub

Private Sub txtDepartment_Validate(Cancel As Boolean)
    MyCustomer_.DepartmentName = txtDepartment.Text
End Sub

Private Sub txtEmail_Validate(Cancel As Boolean)
    MyCustomer_.EmailAddr = txtEmail.Text
End Sub

Private Sub txtFactoryID_Validate(Cancel As Boolean)
    MyCustomer_.RegisterNo = txtFactoryID.Text
End Sub

Private Sub txtFaxNo_Validate(Cancel As Boolean)
    MyCustomer_.FaxNo = txtFaxNo.Text
End Sub

Private Sub txtGenID_Validate(Cancel As Boolean)
    MyCustomer_.WasteManageID = txtGenID.Text
End Sub

Private Sub txtIndustryType_Validate(Cancel As Boolean)
'    If Trim(txtIndustryType.Text) = "" And Trim(curIndustryType_.TypeName) <> Trim(txtIndustryType.Text) Then
'        Set curIndustryType_ = Nothing
'        Set curIndustryType_ = New BWGIndustryType
'        With MyCustomer_
'            .IndustryTypeID = ""
'            .IndustryTypeName = ""
'        End With
'    Else
'        If Trim(curIndustryType_.TypeName) <> Trim(txtIndustryType.Text) Then
'            Set curIndustryType_ = Nothing
'            Set curIndustryType_ = New BWGIndustryType
'            With MyCustomer_
'                .IndustryTypeID = ""
'                .IndustryTypeName = ""
'            End With
'            Load frmIndustryType
'            With frmIndustryType
'                .SetCallerForm Me
'                .Show 1
'            End With
'            If MyCustomer_.IndustryTypeID = "" Then txtIndustryType.Text = ""
'        End If
'    End If
End Sub

Private Sub txtInv_Contact_Validate(Cancel As Boolean)
    If Trim(txtInv_Contact.Text) = "" Then
        MyCustomer_.BillingInfo.ContactName = ""
    Else
        MyCustomer_.BillingInfo.ContactName = Trim(txtInv_Contact.Text)
        MyCustomer_.isChanged = True
    End If
End Sub

Private Sub txtInv_Fax_Validate(Cancel As Boolean)
    If Trim(txtInv_Fax.Text) = "" Then
        MyCustomer_.BillingInfo.FaxNo = ""
    Else
        MyCustomer_.BillingInfo.FaxNo = Trim(txtInv_Fax.Text)
        MyCustomer_.isChanged = True
    End If
End Sub

Private Sub txtInv_tel_Validate(Cancel As Boolean)
    If Trim(txtInv_tel.Text) = "" Then
        MyCustomer_.BillingInfo.TelNo = ""
    Else
        MyCustomer_.BillingInfo.TelNo = Trim(txtInv_tel.Text)
        MyCustomer_.isChanged = True
    End If
End Sub

Private Sub txtInvTime_Validate(Cancel As Boolean)
    If Trim(txtInvTime.Text) = "" Then
        MyCustomer_.BillingInfo.inv_Time = ""
    Else
        MyCustomer_.BillingInfo.inv_Time = Trim(txtInvTime.Text)
        MyCustomer_.isChanged = True
    End If
End Sub
Private Sub TxtLack_Validate(Cancel As Boolean)
   If Not IsNumeric(txtLack.Text) Then
        MsgBox "ตัวเลขจำนวนเต็มเท่านั้น"
        txtLack.Text = ""
    Else
        MyCustomer_.Lack = txtLack.Text
    End If
        MyCustomer_.isChanged = True
End Sub

Private Sub txtLackBack_Validate(Cancel As Boolean)
        MyCustomer_.LackBack = txtLackBack.Text
        MyCustomer_.isChanged = True
End Sub

Private Sub txtLackW_Validate(Cancel As Boolean)
        MyCustomer_.LackW = txtLackW.Text
        MyCustomer_.isChanged = True
End Sub

Private Sub txtLocation_Validate(Cancel As Boolean)
   If Trim(txtLocation.Text) = "" Then
        MyCustomer_.BillingInfo.inv_Location = ""
   Else
        MyCustomer_.BillingInfo.inv_Location = Trim(txtLocation.Text)
        MyCustomer_.isChanged = True
   End If
End Sub

Private Sub txtLUGBack_Validate(Cancel As Boolean)
'    If Not IsNumeric(txtLUGBack.Text) Then
'        MsgBox "ตัวเลขจำนวนเต็มเท่านั้น"
'        txtLUGBack.Text = ""
'    Else
        MyCustomer_.LUGBOXBackV = txtLUGBack.Text
'    End If
    MyCustomer_.isChanged = True
End Sub

Private Sub txtLUGBOX_Validate(Cancel As Boolean)
    If Not IsNumeric(txtLUGBOX.Text) Then
        MsgBox "ตัวเลขจำนวนเต็มเท่านั้น"
        txtLUGBOX.Text = ""
    Else
        MyCustomer_.LUGBOX = txtLUGBOX.Text
    End If
    MyCustomer_.isChanged = True
End Sub

Private Sub txtLugBoxW_Validate(Cancel As Boolean)
'      If Not IsNumeric(txtLugBoxW.Text) Then
'        MsgBox "ตัวเลขจำนวนเต็มเท่านั้น"
'        txtLugBoxW.Text = ""
'    Else
        MyCustomer_.LUGBoxW = txtLugBoxW.Text
'    End If
        MyCustomer_.isChanged = True
End Sub

Private Sub txtLugTruckW_Validate(Cancel As Boolean)
'    If Not IsNumeric(txtLugTruckW.Text) Then
'        MsgBox "ตัวเลขจำนวนเต็มเท่านั้น"
'        txtLugTruckW.Text = ""
'    Else
        MyCustomer_.TruckBOXW = txtLugTruckW.Text
'    End If
        MyCustomer_.isChanged = True
End Sub

Private Sub txtOneBack_Validate(Cancel As Boolean)
'  If Not IsNumeric(txtOneBack.Text) Then
'        MsgBox "ตัวเลขจำนวนเต็มเท่านั้น"
'        txtOneBack.Text = ""
'    Else
        MyCustomer_.OneTBackv = txtOneBack.Text
'    End If
      MyCustomer_.isChanged = True

End Sub

Private Sub txtOneT_Validate(Cancel As Boolean)
  If Not IsNumeric(txtOneT.Text) Then
        MsgBox "ตัวเลขจำนวนเต็มเท่านั้น"
        txtOneT.Text = ""
    Else
        MyCustomer_.OneT = txtOneT.Text
    End If
      MyCustomer_.isChanged = True
End Sub

Private Sub txtOtherDoc_Validate(Cancel As Boolean)
    If Trim(txtOtherDoc.Text) = "" Then
        chkOtherDoc.Value = 0
        MyCustomer_.BillingInfo.OtherDocDesc = ""
    Else
        chkOtherDoc.Value = 1
        MyCustomer_.BillingInfo.OtherDocDesc = Trim(txtOtherDoc.Text)
        MyCustomer_.isChanged = True
    End If
End Sub

Private Sub txtParentCust_Validate(Cancel As Boolean)
    If Trim(txtParentCust.Text) = "" And Trim(curParentCust_.CustomerName) <> Trim(txtParentCust.Text) Then
        Set curParentCust_ = Nothing
        Set curParentCust_ = New BWGCustomer
        With MyCustomer_
            .ParentCustCode = ""
            .ParentCustID = ""
            .ParentCustName = ""
        End With
    Else
        If Trim(curParentCust_.CustomerName) <> Trim(txtParentCust.Text) Then
            Set curParentCust_ = Nothing
            Set curParentCust_ = New BWGCustomer
            With MyCustomer_
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
            If MyCustomer_.ParentCustID = "" Then txtParentCust.Text = ""
        End If
    End If
End Sub



Private Sub txtPicOfDay_Validate(Cancel As Boolean)
    If IsNumeric(txtPicOfDay.Text) Then
        MyCustomer_.PicOfDay = CInt(txtPicOfDay.Text)
    Else
         MyCustomer_.PicOfDay = 0
    End If
End Sub

Private Sub txtPriceBox_Validate(Cancel As Boolean)
    MyCustomer_.PriceBox = txtPriceBox.Text
End Sub

Private Sub txtREAccContact_Validate(Cancel As Boolean)
    MyCustomer_.BillingInfo.RE_Contact = txtREAccContact.Text
    MyCustomer_.isChanged = True
End Sub

Private Sub txtREAccFax_Validate(Cancel As Boolean)
    MyCustomer_.BillingInfo.RE_FaxNo = txtREAccFax.Text
    MyCustomer_.isChanged = True
End Sub

Private Sub txtREAccTel_Validate(Cancel As Boolean)
    MyCustomer_.BillingInfo.RE_Tel = txtREAccTel.Text
    MyCustomer_.isChanged = True
End Sub
Private Sub TxtRefCode_Validate(Cancel As Boolean)
    MyCustomer_.Refcode = TxtRefCode.Text
End Sub

Private Sub txtReLocateion_Validate(Cancel As Boolean)
    MyCustomer_.BillingInfo.RE_Location = Trim(txtReLocateion.Text)
    MyCustomer_.isChanged = True
End Sub

Private Sub txtRemDesc_Validate(Cancel As Boolean)
    MyCustomer_.BillingInfo.Remdesc = Trim(txtRemDesc.Text)
    MyCustomer_.isChanged = True
End Sub

Private Sub txtRemTimeNo_Validate(Cancel As Boolean)
    MyCustomer_.CommentTimeNo = txtRemTimeNo.Text
    MyCustomer_.isChanged = True
End Sub

Private Sub txtReTime_Validate(Cancel As Boolean)
    MyCustomer_.BillingInfo.RE_Time = txtReTime.Text
    MyCustomer_.isChanged = True
End Sub

Private Sub txtROLBack_Validate(Cancel As Boolean)
'    If Not IsNumeric(txtROLBack.Text) Then
'        MsgBox "ตัวเลขจำนวนเต็มเท่านั้น"
'        txtROLBack.Text = ""
'    Else
        MyCustomer_.ROLBOXBackv = txtROLBack.Text
'    End If
    MyCustomer_.isChanged = True
End Sub

Private Sub txtROLBOX_Validate(Cancel As Boolean)
    If Not IsNumeric(txtROLBOX.Text) Then
        MsgBox "ตัวเลขจำนวนเต็มเท่านั้น"
        txtROLBOX.Text = ""
    Else
        MyCustomer_.ROLBOX = txtROLBOX.Text
    End If
    MyCustomer_.isChanged = True
End Sub

Private Sub txtRolBoxW_Validate(Cancel As Boolean)
'      If Not IsNumeric(txtRolBoxW.Text) Then
'        MsgBox "ตัวเลขจำนวนเต็มเท่านั้น"
'        txtRolBoxW.Text = ""
'    Else
        MyCustomer_.ROLBoxW = txtRolBoxW.Text
'    End If
        MyCustomer_.isChanged = True
End Sub

Private Sub txtSale2_Validate(Cancel As Boolean)
    If Trim(txtSale2.Text) = "" Then Exit Sub
    If curSaleStaff_ Is Nothing Then Set curSaleStaff_ = New BWGCompanyStaff
    If Trim(txtSale2.Text) = "" And Trim(curSaleStaff_.StaffFName & " " & curSaleStaff_.StaffLName) <> Trim(txtSale2.Text) Then
        Set curSaleStaff_ = Nothing
        Set curSaleStaff_ = New BWGCompanyStaff
        With MyCustomer_
            .Sale2ID = ""
            .Sale2Name = ""
        End With
    Else
        If Trim(curSaleStaff_.StaffFName & " " & curSaleStaff_.StaffLName) <> Trim(txtSale2.Text) Then
            Set curSaleStaff_ = Nothing
            Set curSaleStaff_ = New BWGCompanyStaff
            With MyCustomer_
                .Sale2ID = ""
                .Sale2Name = ""
            End With
            Load frmStaffSearch
            With frmStaffSearch
                .getCriteria (txtSale2.Text)
                .SetCallerForm Me
                .Show 1
            End With
            If MyCustomer_.Sale2Name = "" Then txtSale2.Text = ""
        End If
    End If
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
    MyCustomer_.CustomerShortName = txtShortName.Text
End Sub

Private Sub txtTelNo_Validate(Cancel As Boolean)
    MyCustomer_.TelNo = txtTelNo.Text
End Sub

Private Sub txtTruckBack_Validate(Cancel As Boolean)
'   If Not IsNumeric(txtTruckBack.Text) Then
'        MsgBox "ตัวเลขจำนวนเต็มเท่านั้น"
'        txtTruckBack.Text = ""
'    Else
        MyCustomer_.TruckBOXBackv = txtTruckBack.Text
'    End If
        MyCustomer_.isChanged = True
End Sub

Private Sub txtTruckBOX_Validate(Cancel As Boolean)
   If Not IsNumeric(txtTruckBOX.Text) Then
        MsgBox "ตัวเลขจำนวนเต็มเท่านั้น"
        txtTruckBOX.Text = ""
    Else
        MyCustomer_.TruckBOX = txtTruckBOX.Text
    End If
        MyCustomer_.isChanged = True
End Sub

Private Sub txtTumbol_Validate(Cancel As Boolean)
    If Trim(txtTumbol.Text) = "" And Trim(curTumbol_.Name) <> Trim(txtTumbol.Text) Then
        Set curTumbol_ = Nothing
        Set curTumbol_ = New IMITumbol
        txtAumphur.Text = ""
        txtProvince.Text = ""
        With MyCustomer_
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
            With MyCustomer_
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
            If MyCustomer_.TumbolID = "" Then txtTumbol.Text = ""
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
            MyCustomer_.AumphurID = .AumphurID
            MyCustomer_.AumphurName = .AumphurName
            MyCustomer_.TumbolID = .id
            MyCustomer_.TumbolName = .Name
            MyCustomer_.ProvinceID = .ProvinceID
            MyCustomer_.ProvinceName = .ProvinceName
        End If
    End With
End Sub

Private Sub CreateNewCustomer()
    Set MyCustomer_ = Nothing
    Set MyCustomer_ = New BWGCustomer
    With MyCustomer_
        .id = GetGUID
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
        End If
        .CreatedStaffID = CurrentUser.id
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
Public Sub SetProb()
Dim i As Integer
Dim tmpList As ListItem
If txtCustCode.Text = "== AUTO ==" Then
    Exit Sub
End If
LvProb.ListItems.Clear
Dim rs As New ADODB.Recordset
Dim x%, tSql As String
    tSql = "SELECT * FROM tbprob WHERE custid='" & Trim(txtCustCode.Text) & "'  ORDER BY ProbDate"
    Set rs = New ADODB.Recordset
    Set rs = GetRS(tSql, adOpenStatic, adLockOptimistic)
    For x = 1 To rs.RecordCount
                Set tmpList = LvProb.ListItems.Add(, "'" & rs!id & "'")
                tmpList.SubItems(1) = rs!custID
                tmpList.SubItems(2) = rs!ProbDate
                tmpList.SubItems(3) = rs!ProbName
                tmpList.SubItems(4) = rs!ProbDet
                rs.MoveNext
    Next

End Sub
Public Sub SetActivity()
Dim i As Integer
Dim tmpList As ListItem
If txtCustCode.Text = "== AUTO ==" Then
    Exit Sub
End If
lvCustActivity.ListItems.Clear
Dim rs As New ADODB.Recordset
Dim x%, tSql As String
    tSql = "SELECT * FROM tbCustomerActivity WHERE CustomerID='" & MyCustomer_.id & "'  ORDER BY ActivityDate"
    Set rs = New ADODB.Recordset
    Set rs = GetRS(tSql, adOpenStatic, adLockOptimistic)
    For x = 1 To rs.RecordCount
    Dim Subject As String
    Dim SubWaste, SubQT, SubOrder, SubTrackOrder, SubTransport, SubBill, SubOther As String
        If rs!SubWaste = "Y" Then
            SubWaste = " กากฯ ของเสีย,"
        End If
        If rs!SubQT = "Y" Then
            SubQT = " ใบเสนอราคา,"
        End If
        If rs!SubOrder = "Y" Then
            SubOrder = " สั่งงาน Order,"
        End If
        If rs!SubTrackOrder = "Y" Then
            SubTrackOrder = " การติดตาม Order,"
        End If
        If rs!SubTransport = "Y" Then
            SubTransport = " ขนส่ง,"
        End If
        If rs!SubBill = "Y" Then
            SubBill = " วางบิลเก็บเงิน,"
        End If
        If rs!SubOther = "Y" Then
            SubOther = "อื่นๆ"
        End If
    Subject = SubWaste & SubQT & SubOrder & SubTrackOrder & SubTransport & SubBill & SubOther
                Set tmpList = lvCustActivity.ListItems.Add(, "'" & rs!ActivityID & "'")
                tmpList.SubItems(1) = rs!CustomerID
                tmpList.SubItems(2) = Format(rs!ActivityDate, "dd/MM/yyyy")
                tmpList.SubItems(3) = Subject
                tmpList.SubItems(4) = rs!Description
                tmpList.SubItems(5) = rs!ActivityResult
                rs.MoveNext
    Next

End Sub
Private Sub ClearScreen()
    lblCancel.Caption = ""
    lblCancel.Visible = False
    txtCustCode.Text = "== AUTO =="
    txtCompanyNameTH.Text = ""
    txtAddr.Text = ""
    TxtRefCode.Text = ""
    TxtBeginDate.Text = ""
    txtTumbol.Text = ""
    cboCustGroup.ListIndex = -1
    txtAumphur.Text = ""
    txtCustRemark.Text = ""
    txtProvince.Text = ""
    txtShortName.Text = ""
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
    LvProb.ListItems.Clear
    lvCustRem.ListItems.Clear
    Set curTumbol_ = Nothing
    Set curTumbol_ = New IMITumbol
    cboCustType.ListIndex = 0
    txtParentCust.Text = ""
    Set curParentCust_ = Nothing
    Set curParentCust_ = New BWGCustomer
    lvSiteAddress.ListItems.Clear
    txtInv_Contact.Text = ""
    txtInv_tel.Text = ""
    txtInv_Fax.Text = ""
    txtBillingDet.Text = ""
    chkPO.Value = 0
    chkInv.Value = 0
    chkBilling.Value = 0
    chkContract.Value = 0
    chkQuotation.Value = 0
    txtRemDesc.Text = ""
    txtZipCode.Text = ""
    txtSaleName.Text = ""
    txtWebsite.Text = ""
    txtEmail.Text = ""
    txtOtherDoc.Text = ""
    lvContract.ListItems.Clear
    lvCommentQuo.ListItems.Clear
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
    cboDTeam.ListIndex = 0
    chkisCheck.Value = 0
    
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
            MyCustomer_.IndustryTypeID = .TypeID
            MyCustomer_.IndustryTypeName = .TypeDesc
            MyCustomer_.IndustryTypeIDNew = .TypeID
            MyCustomer_.TypeDesc = .TypeDesc
        End If
    End With
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

Private Sub PopulateDistrict()
Dim tmpS As New BWGSearchManager
Dim x%
    Set colDistrict_ = tmpS.SaleDistrictSearch("isActive = 'Y'")
    cboDistrict.Clear
    cboDistrict.AddItem "== กรุณาระบุเขต Sale =="
    For x = 1 To colDistrict_.Count
        cboDistrict.AddItem colDistrict_(x).DistrictName
    Next
    cboDistrict.ListIndex = 0
End Sub

'Private Sub InitSaleAP()
'Dim RS As New ADODB.Recordset
'Dim x%, tSql As String
'    tSql = "SELECT APCode, APSale FROM tbSaleAP_Mapping ORDER BY APCode"
'    Set RS = New ADODB.Recordset
'    Set RS = GetRS(tSql, adOpenStatic, adLockOptimistic)
'    cboSaleAP.Clear
'    cboSaleAP.AddItem ""
'    For x = 1 To RS.RecordCount
'        cboSaleAP.AddItem RS!APSale
'        RS.MoveNext
'    Next
'    cboSaleAP.ListIndex = 0
'End Sub
Public Sub SetDataCustomer(selCust As BWGCustomer, Optional isAssigned As Boolean = True)
On Error GoTo err:
Dim tmpS As New BWGSearchManager
Dim tmpCol As Collection
    If callLocation_ = "N" Then
        Set MyCustomer_ = selCust
        With MyCustomer_
            LoadTR_Remark
            If .Status = "C" Then
                lblCancel.Caption = "ยกเลิก"
                lblCancel.Visible = True
            ElseIf .Status = "Y" Then
                lblCancel.Caption = "ลูกค้าถูกลบ"
                lblCancel.Visible = True
            Else
                lblCancel.Visible = False
            End If
            If .ischeck = "Y" Then
               lblCancel.Caption = "ปรับราคาใหม่"
               lblCancel.Visible = True
               chkisCheck.Value = 1
            End If
            
            lblLastUpdateCR.Caption = "อัพเดทข้อมูลลูกค้าโดย : " & .LastUser & " " & Format(.LastUpdate, "dd/MM/yyyy เวลา HH:mm:ss")
            lblLastUpdateTR.Caption = "อัพเดทข้อมูลขนส่งโดย : " & .LastUserTR & " " & Format(.LastUpdateTR, "dd/MM/yyyy เวลา HH:mm:ss")
            lblLastUpdateAC.Caption = "อัพเดทข้อมูลบัญชีโดย : " & .BillingInfo.LastUser & " " & Format(.BillingInfo.LastUpdate, "dd/MM/yyyy เวลา HH:mm:ss")
        
            txtCustCode.Text = .CustomerCode
            txtCompanyNameTH.Text = .CustomerName
            txtAddr.Text = .Address
            txtZipCode.Text = .ZipCode
            txtEmail.Text = .EmailAddr
            txtWebsite.Text = .Website
            txtCustRemark.Text = .CustRemark
            txtShortName.Text = .CustomerShortName
            txtSaleName.Text = .UnderSaleName
            txtRemTimeNo.Text = .CommentTimeNo
            TxtRefCode.Text = .Refcode
            TxtBeginDate.Text = .BeginDate
            txtPicOfDay.Text = .PicOfDay
            If .StpTrn = "1" Then
            ChkBreak.Value = 1
            Else
            ChkBreak.Value = 0
            End If
            SetTextToCombo .CustGroupName, cboCustGroup
            SetTextToCombo .BillingInfo.Inv_ManName, ComboInvMen
            SetTextToCombo .BillingInfo.RE_ManName, cboRE_Men
            
            Set tmpCol = tmpS.TumbolSearch("TumbolID='" & .TumbolID & "'")
            If tmpCol.Count > 0 Then
                SetDataTumbol tmpCol(1), False
                'txtTumbol.Text = .TumbolName
                'txtAumphur.Text = .AumphurName
                'txtProvince.Text = .ProvinceName
            End If
            SetTextToCombo .SaleDistrictName, cboDistrict
            If .CRStaffID <> "" Then
                SetTextToCombo .CRStaffName, cbCRStaff
            Else
                cbCRStaff.ListIndex = 0
            End If
            If .IndustryTypeIDNew <> "" Then
            Set tmpCol = tmpS.IndustryTypeSearchNew("TypeID='" & .IndustryTypeIDNew & "'")
            If tmpCol.Count > 0 Then
                SetDataIndustryType tmpCol(1), False
                'txtIndustryType.Text = .IndustryTypeName
            End If
            End If
            '*** ค้นหาพนักงานขาย
'            Set tmpCol = Nothing
'            Set tmpCol = tmpS.UserSearch("StaffID='" & .UnderSaleStaffID & "'")
'            If tmpCol.Count > 0 Then
'                SetDataStaff tmpCol(1), False
'            End If
'            Set tmpCol = Nothing
            
            txtSale2.Text = .Sale2Name
    
            '*** จบค้นหาพนักงานขาย
            
            txtFactoryID.Text = .RegisterNo
            txtContactName.Text = .ContactPerson
            txtDepartment.Text = .DepartmentName
            txtTelNo.Text = .TelNo
            txtFaxNo.Text = .FaxNo
            txtGenID.Text = .WasteManageID
            If .CustdeptName <> "" Then
                If .CustdeptName <> "0" Then
                    cbDept.Text = .CustdeptName
                Else
                    cbDept.ListIndex = 0
                End If
            Else
                cbDept.ListIndex = 0
            End If

            If .isVIPCR = "Y" Then
                ChkVIPCR.Value = 1
                Label17.Visible = True
            Else
                ChkVIPCR.Value = 0
                Label17.Visible = False
            End If
            If .GPSRequest = "Y" Then
                ChkGPS.Value = 1
            Else
                ChkGPS.Value = 0
            End If
            If .OverWeight = "Y" Then
                ChkOvrW.Value = 1
            Else
                ChkOvrW.Value = 0
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
            Else
                cboCustType.ListIndex = 0
            End If
            Set tmpCol = tmpS.CustomerSearch("CustomerID='" & .ParentCustID & "'")
            If tmpCol.Count > 0 Then
                Set curParentCust_ = tmpCol(1)
                txtParentCust.Text = curParentCust_.CustomerName
            End If
            'SetTextToCombo .SaleAP, cboSaleAP
            If .DTeam = False Then
                cboDTeam.ListIndex = 0
            Else
                cboDTeam.ListIndex = 1
            End If
        End With
    
        With MyCustomer_.BillingInfo
            txtLocation.Text = .inv_Location
            txtInvTime.Text = .inv_Time
            If Trim(.Inv_State) = "S" Then
                Option5.Value = True
            End If
            If Trim(.Inv_State) = "M" Then
                Option6.Value = True
            End If
             If Trim(.Inv_State) = "E" Then
                Option5.Value = True
            End If
             If Trim(.Inv_State) = "O" Then
                Option4.Value = True
            End If
            
             txtInv_Contact.Text = .ContactName
             txtInv_tel.Text = .TelNo
             txtInv_Fax.Text = .FaxNo
             txtBillingDet.Text = .BillingDetail
             txtInvTime.Text = .inv_Time
             txtReLocateion.Text = .RE_Location
             txtReTime.Text = .RE_Time
             txtREAccContact.Text = .RE_Contact
             txtREAccTel.Text = .RE_Tel
             txtREAccFax.Text = .RE_FaxNo
             txtCheqDetail.Text = .CheqDetail

            If .RE_State = "C" Then
                 isRV_Bank.Value = 1
            Else
                 isRV_Bank.Value = 0
            End If
            If .RE_State = "B" Then
                isRV_Inbank.Value = 1
            Else
                isRV_Inbank.Value = 0
            End If
            
            If .RE_State = "E" Then
                isRV_EMS.Value = 1
            Else
                isRV_EMS.Value = 0
            End If
            If .RE_State = "S" Then
                isRE_CB.Value = 1
            Else
                isRE_CB.Value = 0
            End If
            If MyCustomer_.isWeigth = "Y" Then
                ChkISWeigth.Value = 1
            Else
                ChkISWeigth.Value = 0
            End If

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
            txtRemDesc.Text = .Remdesc
            txtComment.Text = MyCustomer_.Comment
            lbDate.Caption = MyCustomer_.CommentDate
            If MyCustomer_.CustGroupID = "488F2E5C057F46869D1A56185B11D644" Then
                Label4.Visible = True
                txtComment.Visible = True
                lbDate.Visible = True
                cmdDate.Visible = True
                cmdDate.Enabled = True
            Else
                If Trim(MyCustomer_.Comment) <> "" Then
                    txtComment.Text = "(เคยเป็นลูกค้าหลุด) " & Trim(MyCustomer_.Comment)
                    txtComment.Visible = True
                    Label4.Visible = True
                    lbDate.Visible = True
                    cmdDate.Visible = True
                    cmdDate.Enabled = False
                Else
                    txtComment.Visible = False
                    Label4.Visible = False
                    lbDate.Visible = False
                    cmdDate.Visible = False
                    cmdDate.Enabled = False
                End If
            End If
            
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
            txtCheqDate(1).Text = ""
            txtCheqDate(2).Text = ""
            txtCheqDate(3).Text = ""
            txtCheqDate(4).Text = ""
            txtCheqDate(5).Text = ""
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
        ShowAuction
        ShowCmmQuo
        ShowCustStaff
        ShowOtherRemark
        SetProb
        SetActivity
If MyCustomer_.isBox = "Y" Then
        chkisBox.Value = 1
        FramBox.Enabled = True
    txtLUGBOX.Text = MyCustomer_.LUGBOX
    txtROLBOX.Text = MyCustomer_.ROLBOX
    txtOneT.Text = MyCustomer_.OneT
    txtTwoT.Text = MyCustomer_.TwoT
    txtTruckBOX.Text = MyCustomer_.TruckBOX
    txtLugBoxW.Text = MyCustomer_.LUGBoxW
    txtRolBoxW.Text = MyCustomer_.ROLBoxW
    txtBoxOneW.Text = MyCustomer_.OneTBoxW
    txtBoxTwoW.Text = MyCustomer_.TwoTBoxW
    txtLugTruckW.Text = MyCustomer_.TruckBOXW
    txtLUGBack.Text = MyCustomer_.LUGBOXBackV
    txtROLBack.Text = MyCustomer_.ROLBOXBackv
    txtOneBack.Text = MyCustomer_.OneTBackv
    txtTwoBack.Text = MyCustomer_.TwoTBackv
    txtTruckBack.Text = MyCustomer_.TruckBOXBackv
    txtLack.Text = MyCustomer_.Lack
    txtLackBack.Text = MyCustomer_.LackBack
    txtLackW.Text = MyCustomer_.LackW
    txtBoutS.Text = MyCustomer_.BoutS
    txtBoutSBack.Text = MyCustomer_.BoutSBack
    txtBuotSW.Text = MyCustomer_.BoutSW
    txtWaterBuot.Text = MyCustomer_.WaterBout
    txtWaterBoutW.Text = MyCustomer_.WaterBoutW
    txtWaterBoutBack.Text = MyCustomer_.WaterBoutBack
    txtBoutCage.Text = MyCustomer_.BoutCage
    txtBoutCageW.Text = MyCustomer_.BoutCageW
    txtBuotCageBack.Text = MyCustomer_.BoutCageBack
    txtPicOfDay.Text = MyCustomer_.PicOfDay
    dtLugFromdate.ValueYMD = MyCustomer_.LUGBOXFrom
    dtROLFromdate.ValueYMD = MyCustomer_.ROLBOXFrom
    dtOneTFrom.ValueYMD = MyCustomer_.OneTFrom
    dtTwoTFrom.ValueYMD = MyCustomer_.TwoTFrom
    dtTruckFrom.ValueYMD = MyCustomer_.TruckBOXFrom
    
    dtLackFrom.ValueYMD = MyCustomer_.LackFrom
    dtLackBack.ValueYMD = MyCustomer_.LackTo
    dtBoutSFrom.ValueYMD = MyCustomer_.BoutSFrom
    dtBoutSBack.ValueYMD = MyCustomer_.BoutSTo
    dtWaterBoutFrom.ValueYMD = MyCustomer_.WaterBoutFrom
    dtWaterBoutBack.ValueYMD = MyCustomer_.WaterBoutTo
    dtBoutCageFrom.ValueYMD = MyCustomer_.BoutCageFrom
    dtBoutCageBack.ValueYMD = MyCustomer_.BoutCageTo
   
    dtLugBackdate.ValueYMD = MyCustomer_.LUGBOXBack
    dtROLBackdate.ValueYMD = MyCustomer_.ROLBOXBack
    dtOneTBack.ValueYMD = MyCustomer_.OneTBack
    dtTwoTBack.ValueYMD = MyCustomer_.TwoTBack
    dtTruckBack.ValueYMD = MyCustomer_.TruckBOXBack
        
    Else
        chkisBox.Value = 0
    End If
    ElseIf callLocation_ = "P" Then
        Set curParentCust_ = selCust
        With curParentCust_
            txtParentCust.Text = .CustomerName
            If isAssigned Then
                MyCustomer_.ParentCustCode = .CustomerCode
                MyCustomer_.ParentCustID = .id
                MyCustomer_.ParentCustName = .CustomerName
            End If
        End With
    End If
    txtCreditterm.Text = MyCustomer_.Creditterm
    If Trim(txtCreditterm.Text) = "" Then
        CbCredit.ListIndex = 0
    ElseIf Trim(txtCreditterm.Text) = "30" Then
        CbCredit.ListIndex = 1
    ElseIf Trim(txtCreditterm.Text) = "60" Then
        CbCredit.ListIndex = 2
    ElseIf Trim(txtCreditterm.Text) = "90" Then
        CbCredit.ListIndex = 3
    ElseIf Trim(txtCreditterm.Text) = "120" Then
        CbCredit.ListIndex = 4
    End If
    If cboCustGroup.Text = "F - ลูกค้าที่หลุดแล้ว" Then
        txtSaleName.Text = ""
    End If
    MyCustomer_.BillingInfo.isChanged = False
    MyCustomer_.isChanged = False
    Exit Sub

err:
    MsgBox "ล้มเหลวในการเปิดข้อมูล", vbQuestion
End Sub

Public Property Get CallLocation() As String
    CallLocation = callLocation_
End Property
Public Property Let CallLocation(ByVal NewValue As String)
    callLocation_ = Trim(NewValue)
End Property

Private Sub ShowWasteData()
On Error Resume Next
Dim x%
Dim tmpWaste As BWGWasteDataCR
Dim tmpList As ListItem
    lvWaste.ListItems.Clear
    For x = 1 To MyCustomer_.WasteItems.Count
        Set tmpWaste = MyCustomer_.WasteItems(x)
        Set tmpList = lvWaste.ListItems.Add(, "'" & tmpWaste.id & "'", tmpWaste.WasteNo)
        tmpList.SubItems(1) = tmpWaste.WasteName
        tmpList.SubItems(2) = tmpWaste.EUCode1 & "-" & tmpWaste.EUCode2 & "-" & tmpWaste.EUCode3
        tmpList.SubItems(3) = tmpWaste.WasteType
        tmpList.SubItems(4) = ""
        tmpList.SubItems(5) = tmpWaste.TreatmentNo
        tmpList.SubItems(6) = tmpWaste.QuotationPrice
        Set tmpList = Nothing
        Set tmpWaste = Nothing
    Next
End Sub

Private Sub ShowSiteAddress()
Dim x%
Dim tmpSite As BWGCustomerSite
Dim tmpList As ListItem
    lvSiteAddress.ListItems.Clear
    For x = 1 To MyCustomer_.SiteAddresses.Count
        Set tmpSite = MyCustomer_.SiteAddresses(x)
        Set tmpList = lvSiteAddress.ListItems.Add(, "'" & tmpSite.id & "'", tmpSite.SiteName)
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
    For x = 1 To MyCustomer_.BillAddresses.Count
        Set tmpSite = MyCustomer_.BillAddresses(x)
        Set tmpList = lvBillAddr.ListItems.Add(, "'" & tmpSite.id & "'", tmpSite.BillName)
        tmpList.SubItems(1) = tmpSite.Address & " " & tmpSite.TumbolName & " " & tmpSite.AumphurName & " " & tmpSite.ProvinceName & " " & tmpSite.ZipCode
        Set tmpList = Nothing
        Set tmpSite = Nothing
    Next
End Sub
Private Sub ShowAuction()
Dim x%
Dim tSql As String
Dim rs As New ADODB.Recordset
Dim tmpSite As BWGCustomerAuction
Dim tmpList As ListItem
Dim Stat As String
    lvAuction.ListItems.Clear
    tSql = "SELECT * FROM tbCustomerAuction WHERE CustomerID = '" & MyCustomer_.id & "'"
    Set rs = New ADODB.Recordset
    Set rs = GetRS(tSql, adOpenForwardOnly, adLockReadOnly)
    If rs.RecordCount > 0 Then
        For x = 1 To rs.RecordCount
            Set tmpList = lvAuction.ListItems.Add(, "'" & rs!AuctionID & "'", FormatYMD_to_DMY(rs!StartDate, "/", "/"))
            If rs!EndDate <> "" Then
                tmpList.SubItems(1) = FormatYMD_to_DMY(rs!EndDate, "/", "/")
            End If
            If rs!IsOK = "Y" Then
                Stat = "ได้งาน"
            ElseIf rs!IsOK = "N" Then
                Stat = "ไม่ได้งาน"
            Else
                Stat = "ยังไม่ทราบผล"
            End If
            tmpList.SubItems(2) = rs!DocNo
            tmpList.SubItems(3) = rs!Remdesc
            tmpList.SubItems(4) = rs!Price
            tmpList.SubItems(5) = Stat
            tmpList.SubItems(6) = rs!NextDate
            Set tmpList = Nothing
            Set tmpSite = Nothing
            rs.MoveNext
        Next
    End If
End Sub

Private Sub ShowContracts()
Dim x%
Dim tmpSite As BWGContract
Dim tmpList As ListItem
    lvContract.ListItems.Clear
    For x = 1 To MyCustomer_.Contracts.Count
        Set tmpSite = MyCustomer_.Contracts(x)
        Set tmpList = lvContract.ListItems.Add(, "'" & tmpSite.id & "'", FormatYMD_to_DMY(tmpSite.StartDate, "/", "/"))
        tmpList.SubItems(1) = FormatYMD_to_DMY(tmpSite.EndDate, "/", "/")
        tmpList.SubItems(2) = tmpSite.DocNo
        tmpList.SubItems(3) = tmpSite.Remdesc
        Set tmpList = Nothing
        Set tmpSite = Nothing
    Next
End Sub
Private Sub ShowCmmQuo()
Dim x%
Dim tmpSite As BWGCmmQuoHeader
Dim myCmmQuoHeader As BWGCmmQuoHeader
Dim tmpList As ListItem
    lvCommentQuo.ListItems.Clear
    For x = 1 To MyCustomer_.CmmQuo.Count
        Set tmpSite = MyCustomer_.CmmQuo(x)
        Set tmpList = lvCommentQuo.ListItems.Add(, "'" & tmpSite.id & "'", FormatYMD_to_DMY(tmpSite.StartDate, "/", "/"))
        tmpList.SubItems(1) = FormatYMD_to_DMY(tmpSite.EndDate, "/", "/")
        tmpList.SubItems(2) = tmpSite.CmQuoNo
        tmpList.SubItems(3) = tmpSite.Comment
        tmpList.SubItems(4) = tmpSite.CompanyNameTH
        tmpList.SubItems(5) = tmpSite.id
        tmpSite.id = tmpList.SubItems(5)
        Set tmpList = Nothing
        Set tmpSite = Nothing
    Next
    
End Sub
Public Sub ShowCustStaff()
Dim x%
Dim tmpSite As BWGCustomerStaff
Dim tmpList As ListItem
    lvStaff.ListItems.Clear
    For x = 1 To MyCustomer_.CustStaffs.Count
        Set tmpSite = MyCustomer_.CustStaffs(x)
        Set tmpList = lvStaff.ListItems.Add(, "'" & tmpSite.id & "'", tmpSite.FirstName & " " & tmpSite.LastName)
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

Private Sub txtTwoBack_Validate(Cancel As Boolean)
'    If Not IsNumeric(txtTwoBack.Text) Then
'        MsgBox "ตัวเลขจำนวนเต็มเท่านั้น"
'        txtTwoBack.Text = ""
'    Else
        MyCustomer_.TwoTBackv = txtTwoBack.Text
'    End If
    MyCustomer_.isChanged = True
End Sub

Private Sub txtTwoT_Validate(Cancel As Boolean)
    If Not IsNumeric(txtTwoT.Text) Then
        MsgBox "ตัวเลขจำนวนเต็มเท่านั้น"
        txtTwoT.Text = ""
    Else
        MyCustomer_.TwoT = txtTwoT.Text
    End If
    MyCustomer_.isChanged = True
End Sub
Private Sub txtWaterBoutBack_Validate(Cancel As Boolean)
        MyCustomer_.WaterBoutBack = txtWaterBoutBack.Text
        MyCustomer_.isChanged = True
End Sub

Private Sub txtWaterBoutW_Validate(Cancel As Boolean)
        MyCustomer_.WaterBoutW = txtWaterBoutW.Text
        MyCustomer_.isChanged = True
End Sub
Private Sub txtWaterBuot_Validate(Cancel As Boolean)
   If Not IsNumeric(txtWaterBuot.Text) Then
        MsgBox "ตัวเลขจำนวนเต็มเท่านั้น"
        txtWaterBuot.Text = ""
    Else
        MyCustomer_.WaterBout = txtWaterBuot.Text
    End If
        MyCustomer_.isChanged = True
End Sub
Private Sub txtWebSite_Validate(Cancel As Boolean)
    MyCustomer_.Website = txtWebsite.Text
End Sub
Private Sub txtZipCode_Validate(Cancel As Boolean)
    MyCustomer_.ZipCode = Trim(txtZipCode.Text)
End Sub

Public Sub SetDataStaff(selStaff As BWGCompanyStaff, Optional isAssigned As Boolean = True)
    Set curSaleStaff_ = selStaff
    With curSaleStaff_
        txtSale2.Text = .StaffFName & " " & .StaffLName
        If isAssigned Then
            MyCustomer_.Sale2Name = .StaffFName & " " & .StaffLName
            MyCustomer_.Sale2ID = .id
        End If
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
End Sub

Private Sub PopulateBillInfo()
Dim tmpS As New BWGSearchManager
Dim x%
    Set colBillingInfoStaff_ = tmpS.BilInfoStaffSearch
    ComboInvMen.Clear
    cboRE_Men.Clear
    For x = 1 To colBillingInfoStaff_.Count
        ComboInvMen.AddItem colBillingInfoStaff_(x).InFoStaffName
        cboRE_Men.AddItem colBillingInfoStaff_(x).InFoStaffName
    Next
    
End Sub

Private Sub ShowOtherRemark()
Dim x%
Dim tmpSite As BWGCustOrderRemark
Dim tmpList As ListItem
    lvCustRem.ListItems.Clear
    For x = 1 To MyCustomer_.OtherRemarks.Count
        Set tmpSite = MyCustomer_.OtherRemarks(x)
        Set tmpList = lvCustRem.ListItems.Add(, "'" & tmpSite.id & "'", tmpSite.Remdesc)
        tmpList.SubItems(1) = tmpSite.OrdPos
        Set tmpList = Nothing
        Set tmpSite = Nothing
    Next
End Sub
