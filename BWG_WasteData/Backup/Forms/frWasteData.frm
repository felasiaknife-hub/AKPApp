VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmWasteData 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Waste Data"
   ClientHeight    =   10230
   ClientLeft      =   2820
   ClientTop       =   330
   ClientWidth     =   11895
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10230
   ScaleWidth      =   11895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   2
      Left            =   0
      TabIndex        =   300
      Top             =   0
      Width           =   1260
   End
   Begin MSComCtl2.DTPicker DtpStart 
      Height          =   330
      Left            =   3420
      TabIndex        =   72
      Top             =   2655
      Visible         =   0   'False
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   582
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   139526145
      CurrentDate     =   42978
   End
   Begin VB.ComboBox cboWasteCase 
      Height          =   315
      ItemData        =   "frWasteData.frx":0000
      Left            =   405
      List            =   "frWasteData.frx":0002
      TabIndex        =   71
      Text            =   "ปกติ"
      Top             =   2700
      Visible         =   0   'False
      Width           =   2235
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00C0C0FF&
      Caption         =   "ระยะเวลาใบอนุญาติ"
      ForeColor       =   &H00C00000&
      Height          =   540
      Left            =   7545
      TabIndex        =   60
      Top             =   2055
      Width           =   3945
      Begin VB.CheckBox ChkLicense 
         BackColor       =   &H00C0C0FF&
         Caption         =   "ไม่มี"
         Height          =   255
         Left            =   3285
         TabIndex        =   64
         Top             =   180
         Width           =   600
      End
      Begin VB.CommandButton Command5 
         Caption         =   "..."
         Height          =   285
         Left            =   2880
         TabIndex        =   62
         Top             =   180
         Width           =   345
      End
      Begin VB.TextBox txtStartDate 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   480
         TabIndex        =   61
         Top             =   195
         Width           =   2400
      End
      Begin VB.Label Label30 
         BackColor       =   &H00C0C0FF&
         Caption         =   "วันที่"
         Height          =   255
         Left            =   120
         TabIndex        =   63
         Top             =   195
         Width           =   705
      End
   End
   Begin VB.CommandButton cmdRev 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Rev. Waste"
      Height          =   795
      Left            =   2460
      TabIndex        =   58
      Top             =   9390
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0FFC0&
      Caption         =   "ดูข้อมูล"
      Height          =   300
      Left            =   2160
      TabIndex        =   57
      Top             =   1080
      Width           =   800
   End
   Begin VB.ComboBox CboWasteYear 
      Height          =   315
      ItemData        =   "frWasteData.frx":0004
      Left            =   1440
      List            =   "frWasteData.frx":0006
      TabIndex        =   55
      Text            =   "2557"
      Top             =   1080
      Width           =   735
   End
   Begin VB.CommandButton cmdAccepdate 
      BackColor       =   &H00FFFFC0&
      Caption         =   "กำหนดวันที่ตรวจรับ"
      Height          =   280
      Left            =   4320
      MaskColor       =   &H00FFC0FF&
      TabIndex        =   54
      Top             =   2370
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Frame Frame0 
      BackColor       =   &H00FFC0C0&
      Height          =   1200
      Left            =   0
      TabIndex        =   48
      Top             =   -120
      Width           =   11865
      Begin VB.TextBox TxtMECode 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Enabled         =   0   'False
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1620
         Locked          =   -1  'True
         TabIndex        =   76
         TabStop         =   0   'False
         Top             =   900
         Width           =   2130
      End
      Begin VB.ComboBox cboWasteAnalysis 
         BackColor       =   &H0080FFFF&
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   7080
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   540
         Width           =   4695
      End
      Begin VB.TextBox txtWasteNo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000000C0&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1170
         Locked          =   -1  'True
         TabIndex        =   49
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
         Width           =   4695
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "รหัสลูกค้า Better ME"
         Height          =   195
         Index           =   13
         Left            =   135
         TabIndex        =   75
         Top             =   945
         Width           =   1410
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "เลขที่ของเสีย"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   53
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
         TabIndex        =   52
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ชื่อลูกค้า"
         Height          =   195
         Index           =   24
         Left            =   510
         TabIndex        =   51
         Top             =   240
         Width           =   570
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "            บริษัท                          การวิเคราะห์"
         ForeColor       =   &H00000000&
         Height          =   690
         Index           =   32
         Left            =   6045
         TabIndex        =   50
         Top             =   240
         Width           =   1005
      End
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "ไม่มีตัวอย่าง"
      Height          =   195
      Left            =   4320
      TabIndex        =   8
      Top             =   1140
      Value           =   -1  'True
      Width           =   1155
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "มีตัวอย่าง"
      Height          =   195
      Left            =   3240
      TabIndex        =   7
      Top             =   1140
      Width           =   975
   End
   Begin VB.CommandButton cmdApprove 
      BackColor       =   &H00FFFFC0&
      Caption         =   "ตรวจรับโดย Envi"
      Height          =   795
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   9390
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.TextBox txtIndustryType 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   2055
      Width           =   5940
   End
   Begin VB.ComboBox cboLABRes 
      Height          =   315
      Left            =   6600
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   1080
      Width           =   1275
   End
   Begin VB.TextBox txtWasteGenNo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5700
      MaxLength       =   15
      TabIndex        =   15
      Top             =   1755
      Width           =   2115
   End
   Begin VB.TextBox txtFactoryNo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1560
      MaxLength       =   20
      TabIndex        =   14
      Top             =   1755
      Width           =   1965
   End
   Begin VB.TextBox txtContractNo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   9765
      TabIndex        =   13
      Top             =   1410
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H008080FF&
      Caption         =   "ปิด"
      Height          =   795
      Left            =   10260
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   9390
      Width           =   945
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H0000C000&
      Caption         =   "เก็บข้อมูล"
      Height          =   795
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   9390
      Width           =   945
   End
   Begin VB.CommandButton cmdReject 
      BackColor       =   &H00FFFFFF&
      Caption         =   "ตีกลับ"
      Height          =   795
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   9390
      Width           =   825
   End
   Begin VB.CommandButton cmdDel 
      BackColor       =   &H000000FF&
      Caption         =   "ลบข้อมูล"
      Height          =   795
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   9390
      Width           =   945
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H0080FF80&
      Caption         =   "สร้างใหม่"
      Height          =   795
      Left            =   8220
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   9390
      Width           =   945
   End
   Begin VB.ComboBox cboSummary 
      Height          =   315
      ItemData        =   "frWasteData.frx":0008
      Left            =   8820
      List            =   "frWasteData.frx":0015
      Style           =   2  'Dropdown List
      TabIndex        =   25
      Top             =   9000
      Width           =   2385
   End
   Begin VB.TextBox txtEnviStaff 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   5370
      TabIndex        =   24
      Top             =   8985
      Width           =   2685
   End
   Begin VB.TextBox txtMktStaff 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1440
      TabIndex        =   23
      Top             =   8985
      Width           =   2685
   End
   Begin VB.TextBox txtGenerateDesc 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000080&
      Height          =   690
      Left            =   2655
      MultiLine       =   -1  'True
      TabIndex        =   22
      Top             =   3930
      Width           =   9210
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "ปริมาณของเสีย"
      ForeColor       =   &H00C00000&
      Height          =   855
      Left            =   -60
      TabIndex        =   36
      Top             =   3060
      Width           =   11550
      Begin VB.TextBox txtLicenseWeighBWG 
         Height          =   315
         Left            =   8640
         TabIndex        =   68
         Text            =   "0"
         Top             =   480
         Width           =   915
      End
      Begin VB.TextBox txtLicenseWeigh 
         Height          =   315
         Left            =   8640
         TabIndex        =   67
         Text            =   "0"
         Top             =   120
         Width           =   915
      End
      Begin VB.ComboBox cboTimeRate 
         Height          =   315
         ItemData        =   "frWasteData.frx":0046
         Left            =   4170
         List            =   "frWasteData.frx":0048
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   495
         Width           =   795
      End
      Begin VB.ComboBox cboUnitRate 
         Height          =   315
         Left            =   3240
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   495
         Width           =   735
      End
      Begin VB.TextBox txtRate 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1950
         TabIndex        =   19
         Top             =   495
         Width           =   1005
      End
      Begin VB.ComboBox cboQtyUnit 
         Height          =   315
         ItemData        =   "frWasteData.frx":004A
         Left            =   3240
         List            =   "frWasteData.frx":004C
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   135
         Width           =   735
      End
      Begin VB.TextBox txtQty 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1920
         TabIndex        =   17
         Top             =   135
         Width           =   1215
      End
      Begin VB.Label Label37 
         BackColor       =   &H00FFC0FF&
         Caption         =   "ต้น / ปี"
         Height          =   375
         Left            =   9720
         TabIndex        =   70
         Top             =   435
         Width           =   495
      End
      Begin VB.Label Label36 
         BackColor       =   &H00FFC0FF&
         Caption         =   "ตัน / ปี"
         Height          =   255
         Left            =   9720
         TabIndex        =   69
         Top             =   135
         Width           =   615
      End
      Begin VB.Label Label35 
         BackColor       =   &H00FFC0FF&
         Caption         =   "ปริมาณที่ขออนุญาติ กับบริษัท"
         Height          =   255
         Left            =   6420
         TabIndex        =   66
         Top             =   495
         Width           =   1995
      End
      Begin VB.Label Label33 
         BackColor       =   &H00FFC0FF&
         Caption         =   "ปริมาณที่ขออนุญาติ ทั้งหมด"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   6420
         TabIndex        =   65
         Top             =   135
         Width           =   1995
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "/"
         Height          =   195
         Index           =   11
         Left            =   4020
         TabIndex        =   38
         Top             =   585
         Width           =   75
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "จัดเก็บ ณ ปัจจุบัน                   อัตราการเกิดของเสีย"
         ForeColor       =   &H00000000&
         Height          =   645
         Index           =   8
         Left            =   210
         TabIndex        =   37
         Top             =   180
         Width           =   1710
      End
   End
   Begin VB.TextBox txtWasteName 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1560
      MaxLength       =   100
      TabIndex        =   12
      Top             =   1410
      Width           =   6900
   End
   Begin VB.ComboBox cboWasteType 
      Height          =   315
      ItemData        =   "frWasteData.frx":004E
      Left            =   9720
      List            =   "frWasteData.frx":0050
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   1080
      Width           =   1575
   End
   Begin VB.TextBox txt14Digit 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   9630
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   1755
      Width           =   1725
   End
   Begin MSComCtl2.DTPicker DtpEnd 
      Height          =   330
      Left            =   5940
      TabIndex        =   73
      Top             =   2655
      Visible         =   0   'False
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   582
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   139526145
      CurrentDate     =   42978
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4380
      Left            =   0
      TabIndex        =   77
      TabStop         =   0   'False
      Top             =   4560
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   7726
      _Version        =   393216
      Style           =   1
      Tabs            =   11
      Tab             =   6
      TabsPerRow      =   11
      TabHeight       =   520
      WordWrap        =   0   'False
      ShowFocusRect   =   0   'False
      BackColor       =   12648384
      ForeColor       =   12582912
      TabCaption(0)   =   "คุณสมบัติของเสีย"
      TabPicture(0)   =   "frWasteData.frx":0052
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "txtQuoPrice"
      Tab(0).Control(1)=   "txtColor"
      Tab(0).Control(2)=   "chkColor"
      Tab(0).Control(3)=   "txtOdour"
      Tab(0).Control(4)=   "chkOdour"
      Tab(0).Control(5)=   "Frame2"
      Tab(0).Control(6)=   "cboWasteCharDesc"
      Tab(0).Control(7)=   "cboWasteChar"
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "วิธีการจัดเก็บ"
      TabPicture(1)   =   "frWasteData.frx":006E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "CmdDateAnalysis"
      Tab(1).Control(1)=   "chkBottle"
      Tab(1).Control(2)=   "txtBottle"
      Tab(1).Control(3)=   "chkPaperBox"
      Tab(1).Control(4)=   "txtPaperBox"
      Tab(1).Control(5)=   "chkPlasticGallon"
      Tab(1).Control(6)=   "txtPlasticGallon"
      Tab(1).Control(7)=   "chkPlasticDrum"
      Tab(1).Control(8)=   "txtPlasticDrum"
      Tab(1).Control(9)=   "chkPlasticBag"
      Tab(1).Control(10)=   "txtPlasticBag"
      Tab(1).Control(11)=   "cboTreatmentType"
      Tab(1).Control(12)=   "chkYard"
      Tab(1).Control(13)=   "chkTank"
      Tab(1).Control(14)=   "txtOther2"
      Tab(1).Control(15)=   "chkOther2"
      Tab(1).Control(16)=   "txtDrumDesc"
      Tab(1).Control(17)=   "chkDrum"
      Tab(1).Control(18)=   "txtBag"
      Tab(1).Control(19)=   "chkBag"
      Tab(1).Control(20)=   "cboIndoor"
      Tab(1).Control(21)=   "cboStoreType"
      Tab(1).Control(22)=   "lbSameplingDate"
      Tab(1).Control(23)=   "Label26"
      Tab(1).Control(24)=   "Label7(5)"
      Tab(1).Control(25)=   "Label7(4)"
      Tab(1).Control(26)=   "Label7(3)"
      Tab(1).Control(27)=   "Label7(2)"
      Tab(1).Control(28)=   "Label7(1)"
      Tab(1).Control(29)=   "Label7(0)"
      Tab(1).Control(30)=   "Label2(31)"
      Tab(1).Control(31)=   "Label2(19)"
      Tab(1).Control(32)=   "Label2(18)"
      Tab(1).Control(33)=   "Label2(17)"
      Tab(1).ControlCount=   34
      TabCaption(2)   =   "วิธีการกำจัด"
      TabPicture(2)   =   "frWasteData.frx":008A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame4"
      Tab(2).Control(1)=   "Frame3"
      Tab(2).Control(2)=   "lvNewTreatment"
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "ไฟล์แนบ"
      TabPicture(3)   =   "frWasteData.frx":00A6
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
      TabCaption(4)   =   "Tab4"
      TabPicture(4)   =   "frWasteData.frx":00C2
      Tab(4).ControlEnabled=   0   'False
      Tab(4).ControlCount=   0
      TabCaption(5)   =   "ผลการวิเคราะห์ AKP"
      TabPicture(5)   =   "frWasteData.frx":00DE
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Label12"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).Control(1)=   "Label13"
      Tab(5).Control(1).Enabled=   0   'False
      Tab(5).Control(2)=   "Label14"
      Tab(5).Control(2).Enabled=   0   'False
      Tab(5).Control(3)=   "Label16(0)"
      Tab(5).Control(3).Enabled=   0   'False
      Tab(5).Control(4)=   "Label16(1)"
      Tab(5).Control(4).Enabled=   0   'False
      Tab(5).Control(5)=   "Label16(2)"
      Tab(5).Control(5).Enabled=   0   'False
      Tab(5).Control(6)=   "Label16(3)"
      Tab(5).Control(6).Enabled=   0   'False
      Tab(5).Control(7)=   "Label17(0)"
      Tab(5).Control(7).Enabled=   0   'False
      Tab(5).Control(8)=   "Label17(1)"
      Tab(5).Control(8).Enabled=   0   'False
      Tab(5).Control(9)=   "Label18(0)"
      Tab(5).Control(9).Enabled=   0   'False
      Tab(5).Control(10)=   "Label18(1)"
      Tab(5).Control(10).Enabled=   0   'False
      Tab(5).Control(11)=   "Label19"
      Tab(5).Control(11).Enabled=   0   'False
      Tab(5).Control(12)=   "Label20"
      Tab(5).Control(12).Enabled=   0   'False
      Tab(5).Control(13)=   "Label21"
      Tab(5).Control(13).Enabled=   0   'False
      Tab(5).Control(14)=   "Label22"
      Tab(5).Control(14).Enabled=   0   'False
      Tab(5).Control(15)=   "Label23"
      Tab(5).Control(15).Enabled=   0   'False
      Tab(5).Control(16)=   "Label24"
      Tab(5).Control(16).Enabled=   0   'False
      Tab(5).Control(17)=   "Label25"
      Tab(5).Control(17).Enabled=   0   'False
      Tab(5).Control(18)=   "Label28"
      Tab(5).Control(18).Enabled=   0   'False
      Tab(5).Control(19)=   "lbReceivedDate"
      Tab(5).Control(19).Enabled=   0   'False
      Tab(5).Control(20)=   "Label29"
      Tab(5).Control(20).Enabled=   0   'False
      Tab(5).Control(21)=   "lbAnalysisDate"
      Tab(5).Control(21).Enabled=   0   'False
      Tab(5).Control(22)=   "Label27"
      Tab(5).Control(22).Enabled=   0   'False
      Tab(5).Control(23)=   "lbAnalysisNo"
      Tab(5).Control(23).Enabled=   0   'False
      Tab(5).Control(24)=   "Label31"
      Tab(5).Control(24).Enabled=   0   'False
      Tab(5).Control(25)=   "Label39"
      Tab(5).Control(25).Enabled=   0   'False
      Tab(5).Control(26)=   "Label40"
      Tab(5).Control(26).Enabled=   0   'False
      Tab(5).Control(27)=   "Label41"
      Tab(5).Control(27).Enabled=   0   'False
      Tab(5).Control(28)=   "Label42"
      Tab(5).Control(28).Enabled=   0   'False
      Tab(5).Control(29)=   "Label43"
      Tab(5).Control(29).Enabled=   0   'False
      Tab(5).Control(30)=   "Label34"
      Tab(5).Control(30).Enabled=   0   'False
      Tab(5).Control(31)=   "Label9"
      Tab(5).Control(31).Enabled=   0   'False
      Tab(5).Control(32)=   "Label38"
      Tab(5).Control(32).Enabled=   0   'False
      Tab(5).Control(33)=   "Label46"
      Tab(5).Control(33).Enabled=   0   'False
      Tab(5).Control(34)=   "Label63"
      Tab(5).Control(34).Enabled=   0   'False
      Tab(5).Control(35)=   "Label49"
      Tab(5).Control(35).Enabled=   0   'False
      Tab(5).Control(36)=   "txtviscosity"
      Tab(5).Control(36).Enabled=   0   'False
      Tab(5).Control(37)=   "txtBulkdensity"
      Tab(5).Control(37).Enabled=   0   'False
      Tab(5).Control(38)=   "txtHeatingValue"
      Tab(5).Control(38).Enabled=   0   'False
      Tab(5).Control(39)=   "txtmoisturecontent"
      Tab(5).Control(39).Enabled=   0   'False
      Tab(5).Control(40)=   "txtWatercontent"
      Tab(5).Control(40).Enabled=   0   'False
      Tab(5).Control(41)=   "txtAshcontent"
      Tab(5).Control(41).Enabled=   0   'False
      Tab(5).Control(42)=   "txtCL"
      Tab(5).Control(42).Enabled=   0   'False
      Tab(5).Control(43)=   "textF"
      Tab(5).Control(43).Enabled=   0   'False
      Tab(5).Control(44)=   "txtS"
      Tab(5).Control(44).Enabled=   0   'False
      Tab(5).Control(45)=   "txtSludgecontent"
      Tab(5).Control(45).Enabled=   0   'False
      Tab(5).Control(46)=   "txtsolidcontent"
      Tab(5).Control(46).Enabled=   0   'False
      Tab(5).Control(47)=   "txtAs"
      Tab(5).Control(47).Enabled=   0   'False
      Tab(5).Control(48)=   "txtCd"
      Tab(5).Control(48).Enabled=   0   'False
      Tab(5).Control(49)=   "txtCr"
      Tab(5).Control(49).Enabled=   0   'False
      Tab(5).Control(50)=   "txtHg"
      Tab(5).Control(50).Enabled=   0   'False
      Tab(5).Control(51)=   "txtPb"
      Tab(5).Control(51).Enabled=   0   'False
      Tab(5).Control(52)=   "txtAcomment"
      Tab(5).Control(52).Enabled=   0   'False
      Tab(5).Control(53)=   "cboAnalysis"
      Tab(5).Control(53).Enabled=   0   'False
      Tab(5).Control(54)=   "Command2"
      Tab(5).Control(54).Enabled=   0   'False
      Tab(5).Control(55)=   "Command4"
      Tab(5).Control(55).Enabled=   0   'False
      Tab(5).Control(56)=   "txtPhysicalstate"
      Tab(5).Control(56).Enabled=   0   'False
      Tab(5).Control(57)=   "txtCu"
      Tab(5).Control(57).Enabled=   0   'False
      Tab(5).Control(58)=   "txtNi"
      Tab(5).Control(58).Enabled=   0   'False
      Tab(5).Control(59)=   "txtMn"
      Tab(5).Control(59).Enabled=   0   'False
      Tab(5).Control(60)=   "txtZn"
      Tab(5).Control(60).Enabled=   0   'False
      Tab(5).Control(61)=   "txtCut"
      Tab(5).Control(61).Enabled=   0   'False
      Tab(5).Control(62)=   "TxtConcentrate"
      Tab(5).Control(62).Enabled=   0   'False
      Tab(5).Control(63)=   "CbSamplingBy"
      Tab(5).Control(63).Enabled=   0   'False
      Tab(5).Control(64)=   "TxtHeavyFe"
      Tab(5).Control(64).Enabled=   0   'False
      Tab(5).Control(65)=   "Frame5"
      Tab(5).Control(65).Enabled=   0   'False
      Tab(5).Control(66)=   "Frame6"
      Tab(5).Control(66).Enabled=   0   'False
      Tab(5).Control(67)=   "Frame7"
      Tab(5).Control(67).Enabled=   0   'False
      Tab(5).Control(68)=   "TxtAlkalinity"
      Tab(5).Control(68).Enabled=   0   'False
      Tab(5).Control(69)=   "TxtAcidity"
      Tab(5).Control(69).Enabled=   0   'False
      Tab(5).Control(70)=   "TxtSS"
      Tab(5).Control(70).Enabled=   0   'False
      Tab(5).Control(71)=   "txtNCVValue"
      Tab(5).Control(71).Enabled=   0   'False
      Tab(5).ControlCount=   72
      TabCaption(6)   =   "การจัดการ"
      TabPicture(6)   =   "frWasteData.frx":00FA
      Tab(6).ControlEnabled=   -1  'True
      Tab(6).Control(0)=   "ChkCanEvap"
      Tab(6).Control(0).Enabled=   0   'False
      Tab(6).ControlCount=   1
      TabCaption(7)   =   "ประเภทรถ"
      TabPicture(7)   =   "frWasteData.frx":0116
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "ChkisTruckTen"
      Tab(7).Control(0).Enabled=   0   'False
      Tab(7).Control(1)=   "Chkispickup"
      Tab(7).Control(1).Enabled=   0   'False
      Tab(7).Control(2)=   "ChkisLugger"
      Tab(7).Control(2).Enabled=   0   'False
      Tab(7).Control(3)=   "ChkisRollOff"
      Tab(7).Control(3).Enabled=   0   'False
      Tab(7).Control(4)=   "ChkisHuerta"
      Tab(7).Control(4).Enabled=   0   'False
      Tab(7).Control(5)=   "ChkisTrailerW"
      Tab(7).Control(5).Enabled=   0   'False
      Tab(7).Control(6)=   "ChkisTankcar"
      Tab(7).Control(6).Enabled=   0   'False
      Tab(7).Control(7)=   "ChkisFactoryCar"
      Tab(7).Control(7).Enabled=   0   'False
      Tab(7).Control(8)=   "ChkisOrtherCar"
      Tab(7).Control(8).Enabled=   0   'False
      Tab(7).Control(9)=   "TxtFactoryCar"
      Tab(7).Control(9).Enabled=   0   'False
      Tab(7).Control(10)=   "TxtCarOrther"
      Tab(7).Control(10).Enabled=   0   'False
      Tab(7).ControlCount=   11
      TabCaption(8)   =   "หมายเหตุ"
      TabPicture(8)   =   "frWasteData.frx":0132
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "TxtCommentWaste"
      Tab(8).Control(0).Enabled=   0   'False
      Tab(8).ControlCount=   1
      TabCaption(9)   =   "Extraction Test"
      TabPicture(9)   =   "frWasteData.frx":014E
      Tab(9).ControlEnabled=   0   'False
      Tab(9).Control(0)=   "Frame9"
      Tab(9).Control(1)=   "TxtLabcode"
      Tab(9).Control(2)=   "Label44"
      Tab(9).ControlCount=   3
      TabCaption(10)  =   "ผล X-Ray"
      TabPicture(10)  =   "frWasteData.frx":016A
      Tab(10).ControlEnabled=   0   'False
      Tab(10).Control(0)=   "Frame10"
      Tab(10).Control(1)=   "Text1(0)"
      Tab(10).Control(2)=   "Label51"
      Tab(10).ControlCount=   3
      Begin VB.CheckBox ChkCanEvap 
         Caption         =   "Evap โดยตรง"
         Height          =   195
         Left            =   240
         TabIndex        =   310
         Top             =   600
         Width           =   1665
      End
      Begin VB.Frame Frame10 
         Caption         =   "Parameter"
         Height          =   3375
         Left            =   -74880
         TabIndex        =   293
         Top             =   840
         Width           =   11655
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   9
            Left            =   3240
            TabIndex        =   308
            Top             =   1440
            Width           =   8100
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   8
            Left            =   3240
            TabIndex        =   307
            Top             =   1080
            Width           =   8100
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   7
            Left            =   3240
            TabIndex        =   306
            Top             =   720
            Width           =   8100
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   6
            Left            =   3240
            TabIndex        =   305
            Top             =   360
            Width           =   8100
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   5
            Left            =   1200
            TabIndex        =   303
            Top             =   1440
            Width           =   1260
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   4
            Left            =   1200
            TabIndex        =   302
            Top             =   1080
            Width           =   1260
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   3
            Left            =   1200
            TabIndex        =   301
            Top             =   720
            Width           =   1260
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   1
            Left            =   1200
            TabIndex        =   299
            Top             =   360
            Width           =   1260
         End
         Begin VB.Label Label47 
            Caption         =   $"frWasteData.frx":0186
            Height          =   1290
            Index           =   1
            Left            =   2520
            TabIndex        =   304
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label53 
            Caption         =   "Calcium (Ca) :"
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   298
            Top             =   1440
            Width           =   1020
         End
         Begin VB.Label Label53 
            Caption         =   "Iron (Fe) :"
            Height          =   195
            Index           =   2
            Left            =   480
            TabIndex        =   297
            Top             =   1080
            Width           =   780
         End
         Begin VB.Label Label53 
            Caption         =   "Aluminum (Al) :"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   296
            Top             =   720
            Width           =   1140
         End
         Begin VB.Label Label53 
            Caption         =   "Silicon (Si) :"
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   295
            Top             =   360
            Width           =   900
         End
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   0
         Left            =   -73920
         TabIndex        =   291
         Top             =   480
         Width           =   1860
      End
      Begin VB.TextBox txtNCVValue 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   -75000
         TabIndex        =   290
         Top             =   3960
         Width           =   2490
      End
      Begin VB.TextBox TxtSS 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         ForeColor       =   &H00400040&
         Height          =   285
         Left            =   -67695
         TabIndex        =   284
         Top             =   2520
         Width           =   1110
      End
      Begin VB.TextBox TxtAcidity 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         ForeColor       =   &H00400040&
         Height          =   285
         Left            =   -68280
         TabIndex        =   283
         Top             =   2820
         Width           =   1700
      End
      Begin VB.TextBox TxtAlkalinity 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         ForeColor       =   &H00400040&
         Height          =   285
         Left            =   -65280
         TabIndex        =   282
         Top             =   495
         Width           =   1500
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H80000005&
         Caption         =   "ไนไตร์ (Nitrite)"
         ForeColor       =   &H000040C0&
         Height          =   540
         Left            =   -68325
         TabIndex        =   280
         Top             =   765
         Width           =   2000
         Begin VB.CheckBox ChkNt 
            BackColor       =   &H8000000E&
            Caption         =   "เกิดปฏิกิริยา"
            CausesValidation=   0   'False
            Height          =   195
            Left            =   360
            TabIndex        =   281
            Top             =   225
            Width           =   1410
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H80000005&
         Caption         =   "ไซยาไนด์ (Cyanide)"
         ForeColor       =   &H000040C0&
         Height          =   540
         Left            =   -70245
         TabIndex        =   278
         Top             =   765
         Width           =   2000
         Begin VB.CheckBox ChkCn 
            BackColor       =   &H8000000E&
            Caption         =   "เกิดปฏิกิริยา"
            CausesValidation=   0   'False
            Height          =   195
            Left            =   360
            TabIndex        =   279
            Top             =   225
            Width           =   1410
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H80000005&
         Caption         =   "คลอรีน (Free Chlorine)"
         ForeColor       =   &H000040C0&
         Height          =   540
         Left            =   -72165
         TabIndex        =   276
         Top             =   765
         Width           =   2000
         Begin VB.CheckBox ChkCl 
            BackColor       =   &H8000000E&
            Caption         =   "เกิดปฏิกิริยา"
            CausesValidation=   0   'False
            Height          =   195
            Left            =   270
            TabIndex        =   277
            Top             =   225
            Width           =   1410
         End
      End
      Begin VB.TextBox TxtHeavyFe 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         ForeColor       =   &H00400040&
         Height          =   285
         Left            =   -65280
         TabIndex        =   220
         Top             =   2610
         Width           =   1500
      End
      Begin VB.ComboBox CbSamplingBy 
         Height          =   315
         Left            =   -73785
         TabIndex        =   219
         Text            =   "Combo1"
         Top             =   450
         Width           =   2445
      End
      Begin VB.TextBox TxtConcentrate 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         ForeColor       =   &H00400040&
         Height          =   285
         Left            =   -64740
         TabIndex        =   218
         Top             =   2925
         Width           =   960
      End
      Begin VB.Frame Frame9 
         Caption         =   "Parameter"
         Height          =   3255
         Left            =   -74685
         TabIndex        =   192
         Top             =   765
         Width           =   10590
         Begin VB.TextBox TxtRemarkHg 
            Height          =   285
            Left            =   4140
            TabIndex        =   208
            Top             =   2790
            Width           =   5190
         End
         Begin VB.TextBox TxtExtHg 
            Height          =   285
            Left            =   1755
            TabIndex        =   207
            Top             =   2790
            Width           =   1185
         End
         Begin VB.TextBox TxtRemarkZn 
            Height          =   285
            Left            =   4140
            TabIndex        =   206
            Top             =   2430
            Width           =   5190
         End
         Begin VB.TextBox TxtExtZn 
            Height          =   285
            Left            =   1755
            TabIndex        =   205
            Top             =   2430
            Width           =   1185
         End
         Begin VB.TextBox TxtRemarkNi 
            Height          =   285
            Left            =   4140
            TabIndex        =   204
            Top             =   2070
            Width           =   5190
         End
         Begin VB.TextBox TxtExtNi 
            Height          =   285
            Left            =   1755
            TabIndex        =   203
            Top             =   2070
            Width           =   1185
         End
         Begin VB.TextBox TxtRemarkPb 
            Height          =   285
            Left            =   4140
            TabIndex        =   202
            Top             =   1710
            Width           =   5190
         End
         Begin VB.TextBox TxtExtPb 
            Height          =   285
            Left            =   1755
            TabIndex        =   201
            Top             =   1710
            Width           =   1185
         End
         Begin VB.TextBox TxtRemarkCu 
            Height          =   285
            Left            =   4140
            TabIndex        =   200
            Top             =   1350
            Width           =   5190
         End
         Begin VB.TextBox TxtExtCu 
            Height          =   285
            Left            =   1755
            TabIndex        =   199
            Top             =   1350
            Width           =   1185
         End
         Begin VB.TextBox TxtRemarkCr 
            Height          =   285
            Left            =   4140
            TabIndex        =   198
            Top             =   990
            Width           =   5190
         End
         Begin VB.TextBox TxtExtCr 
            Height          =   285
            Left            =   1755
            TabIndex        =   197
            Top             =   990
            Width           =   1185
         End
         Begin VB.TextBox TxtRemarkCd 
            Height          =   285
            Left            =   4140
            TabIndex        =   196
            Top             =   630
            Width           =   5190
         End
         Begin VB.TextBox TxtExtCd 
            Height          =   285
            Left            =   1755
            TabIndex        =   195
            Top             =   630
            Width           =   1185
         End
         Begin VB.TextBox TxtExtAsRemark 
            Height          =   285
            Left            =   4140
            TabIndex        =   194
            Top             =   270
            Width           =   5190
         End
         Begin VB.TextBox TxtExtAs 
            Height          =   285
            Left            =   1755
            TabIndex        =   193
            Top             =   270
            Width           =   1185
         End
         Begin VB.Label Label60 
            Caption         =   "Mercury;Hg :"
            Height          =   240
            Left            =   630
            TabIndex        =   217
            Top             =   2835
            Width           =   960
         End
         Begin VB.Label Label58 
            Caption         =   "Zinc;Zn :"
            Height          =   240
            Left            =   900
            TabIndex        =   216
            Top             =   2475
            Width           =   645
         End
         Begin VB.Label Label56 
            Caption         =   "Nikel;Ni :"
            Height          =   240
            Left            =   900
            TabIndex        =   215
            Top             =   2115
            Width           =   690
         End
         Begin VB.Label Label54 
            Caption         =   "Lead;Pb :"
            Height          =   240
            Left            =   855
            TabIndex        =   214
            Top             =   1755
            Width           =   735
         End
         Begin VB.Label Label52 
            Caption         =   "Copper;Cu :"
            Height          =   240
            Left            =   720
            TabIndex        =   213
            Top             =   1395
            Width           =   870
         End
         Begin VB.Label Label50 
            Caption         =   "Chromiun;Cr :"
            Height          =   240
            Left            =   630
            TabIndex        =   212
            Top             =   1035
            Width           =   1005
         End
         Begin VB.Label Label48 
            Caption         =   "Cadmium;Cd :"
            Height          =   240
            Left            =   585
            TabIndex        =   211
            Top             =   675
            Width           =   1005
         End
         Begin VB.Label Label47 
            Caption         =   $"frWasteData.frx":0271
            Height          =   2850
            Index           =   0
            Left            =   3375
            TabIndex        =   210
            Top             =   270
            Width           =   735
         End
         Begin VB.Label Label45 
            Caption         =   "Arsenic;AS :"
            Height          =   240
            Left            =   720
            TabIndex        =   209
            Top             =   315
            Width           =   870
         End
      End
      Begin VB.TextBox TxtLabcode 
         Height          =   285
         Left            =   -73830
         TabIndex        =   191
         Top             =   405
         Width           =   1860
      End
      Begin VB.TextBox TxtCarOrther 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -69600
         MaxLength       =   15
         TabIndex        =   190
         Top             =   1605
         Width           =   4050
      End
      Begin VB.TextBox TxtFactoryCar 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -69195
         MaxLength       =   15
         TabIndex        =   189
         Top             =   1245
         Width           =   3960
      End
      Begin VB.CheckBox ChkisOrtherCar 
         Caption         =   "อื่นๆ ระบุ :"
         Height          =   195
         Left            =   -70590
         TabIndex        =   188
         Top             =   1650
         Width           =   1815
      End
      Begin VB.CheckBox ChkisFactoryCar 
         Caption         =   "รถโรงงาน ระบุ :"
         Height          =   195
         Left            =   -70590
         TabIndex        =   187
         Top             =   1290
         Width           =   1815
      End
      Begin VB.CheckBox ChkisTankcar 
         Caption         =   "รถแท้งค์"
         Height          =   195
         Left            =   -70590
         TabIndex        =   186
         Top             =   930
         Width           =   1815
      End
      Begin VB.CheckBox ChkisTrailerW 
         Caption         =   "รถเทรลเลอร์"
         Height          =   195
         Left            =   -72885
         TabIndex        =   185
         Top             =   1650
         Width           =   1815
      End
      Begin VB.CheckBox ChkisHuerta 
         Caption         =   "รถเฮียบ"
         Height          =   195
         Left            =   -72885
         TabIndex        =   184
         Top             =   1290
         Width           =   1815
      End
      Begin VB.CheckBox ChkisRollOff 
         Caption         =   "รถ Roll off"
         Height          =   195
         Left            =   -72885
         TabIndex        =   183
         Top             =   885
         Width           =   1815
      End
      Begin VB.CheckBox ChkisLugger 
         Caption         =   "รถ Lugger"
         Height          =   195
         Left            =   -74730
         TabIndex        =   182
         Top             =   1605
         Width           =   1815
      End
      Begin VB.CheckBox Chkispickup 
         Caption         =   "รถกระบะ"
         Height          =   195
         Left            =   -74730
         TabIndex        =   181
         Top             =   1245
         Width           =   1815
      End
      Begin VB.CheckBox ChkisTruckTen 
         Caption         =   "รถสิบล้อ"
         Height          =   195
         Left            =   -74730
         TabIndex        =   180
         Top             =   885
         Width           =   1815
      End
      Begin VB.TextBox TxtCommentWaste 
         Height          =   3480
         Left            =   -74820
         MultiLine       =   -1  'True
         TabIndex        =   179
         Top             =   810
         Width           =   11040
      End
      Begin VB.TextBox txtCut 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         ForeColor       =   &H00400040&
         Height          =   285
         Left            =   -65280
         TabIndex        =   178
         Top             =   2295
         Width           =   1500
      End
      Begin VB.TextBox txtZn 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         ForeColor       =   &H00400040&
         Height          =   285
         Left            =   -65280
         TabIndex        =   177
         Top             =   1995
         Width           =   1500
      End
      Begin VB.TextBox txtMn 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         ForeColor       =   &H00400040&
         Height          =   285
         Left            =   -65280
         TabIndex        =   176
         Top             =   1695
         Width           =   1500
      End
      Begin VB.TextBox txtNi 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         ForeColor       =   &H00400040&
         Height          =   285
         Left            =   -65280
         TabIndex        =   175
         Top             =   1395
         Width           =   1500
      End
      Begin VB.TextBox txtCu 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         ForeColor       =   &H00400040&
         Height          =   285
         Left            =   -65280
         TabIndex        =   174
         Top             =   1095
         Width           =   1500
      End
      Begin VB.TextBox txtQuoPrice 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   405
         Left            =   -73200
         TabIndex        =   173
         Top             =   4620
         Width           =   5425
      End
      Begin VB.TextBox txtColor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         ForeColor       =   &H000000C0&
         Height          =   315
         Left            =   -69540
         TabIndex        =   172
         Top             =   930
         Width           =   1000
      End
      Begin VB.CheckBox chkColor 
         Caption         =   " สี : Color  : "
         Height          =   255
         Left            =   -70680
         TabIndex        =   171
         Top             =   930
         Width           =   1500
      End
      Begin VB.TextBox txtOdour 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         ForeColor       =   &H000000C0&
         Height          =   315
         Left            =   -66660
         TabIndex        =   170
         Top             =   930
         Width           =   1000
      End
      Begin VB.CheckBox chkOdour 
         Caption         =   "กลิ่น : Odour :"
         Height          =   255
         Left            =   -68040
         TabIndex        =   169
         Top             =   930
         Width           =   1500
      End
      Begin VB.TextBox txtPhysicalstate 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         Height          =   375
         Left            =   -74880
         TabIndex        =   168
         Top             =   1770
         Width           =   2415
      End
      Begin VB.CommandButton Command4 
         Caption         =   "เลือก"
         Height          =   300
         Left            =   -64320
         TabIndex        =   167
         Top             =   3900
         Width           =   530
      End
      Begin VB.CommandButton Command2 
         Caption         =   "เลือก"
         Height          =   300
         Left            =   -67080
         TabIndex        =   166
         Top             =   3900
         Width           =   530
      End
      Begin VB.CommandButton CmdDateAnalysis 
         BackColor       =   &H0080C0FF&
         Caption         =   "ระบุวัน"
         Height          =   350
         Left            =   -70560
         MaskColor       =   &H000080FF&
         TabIndex        =   165
         Top             =   3240
         Width           =   1000
      End
      Begin VB.ComboBox cboAnalysis 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         ForeColor       =   &H00400000&
         Height          =   315
         ItemData        =   "frWasteData.frx":035C
         Left            =   -66960
         List            =   "frWasteData.frx":035E
         TabIndex        =   164
         Text            =   "cboAnalysis"
         Top             =   3570
         Width           =   3000
      End
      Begin VB.TextBox txtAcomment 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   -68400
         TabIndex        =   163
         Top             =   3225
         Width           =   4600
      End
      Begin VB.TextBox txtPb 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         ForeColor       =   &H00400040&
         Height          =   285
         Left            =   -65280
         TabIndex        =   162
         Top             =   795
         Width           =   1500
      End
      Begin VB.TextBox txtHg 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         ForeColor       =   &H00400040&
         Height          =   285
         Left            =   -68280
         TabIndex        =   161
         Top             =   2220
         Width           =   1700
      End
      Begin VB.TextBox txtCr 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         ForeColor       =   &H00400040&
         Height          =   285
         Left            =   -68280
         TabIndex        =   160
         Top             =   1920
         Width           =   1700
      End
      Begin VB.TextBox txtCd 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         ForeColor       =   &H00400040&
         Height          =   285
         Left            =   -68280
         TabIndex        =   159
         Top             =   1620
         Width           =   1700
      End
      Begin VB.TextBox txtAs 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         ForeColor       =   &H00400040&
         Height          =   285
         Left            =   -68280
         TabIndex        =   158
         Top             =   1320
         Width           =   1700
      End
      Begin VB.TextBox txtsolidcontent 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         ForeColor       =   &H00000040&
         Height          =   300
         Left            =   -71160
         TabIndex        =   157
         Top             =   3510
         Width           =   1700
      End
      Begin VB.TextBox txtSludgecontent 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         ForeColor       =   &H00000040&
         Height          =   300
         Left            =   -71160
         TabIndex        =   156
         Top             =   3195
         Width           =   1700
      End
      Begin VB.TextBox txtS 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         ForeColor       =   &H00000040&
         Height          =   300
         Left            =   -71160
         TabIndex        =   155
         Top             =   2880
         Width           =   1700
      End
      Begin VB.TextBox textF 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         ForeColor       =   &H00000040&
         Height          =   300
         Left            =   -71160
         TabIndex        =   154
         Top             =   2565
         Width           =   1700
      End
      Begin VB.TextBox txtCL 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         ForeColor       =   &H00000040&
         Height          =   300
         Left            =   -71160
         TabIndex        =   153
         Top             =   2250
         Width           =   1700
      End
      Begin VB.TextBox txtAshcontent 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         ForeColor       =   &H00000040&
         Height          =   300
         Left            =   -71160
         TabIndex        =   152
         Top             =   1935
         Width           =   1700
      End
      Begin VB.TextBox txtWatercontent 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         ForeColor       =   &H00000040&
         Height          =   300
         Left            =   -71160
         TabIndex        =   151
         Top             =   1620
         Width           =   1700
      End
      Begin VB.TextBox txtmoisturecontent 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         ForeColor       =   &H00000040&
         Height          =   300
         Left            =   -71160
         TabIndex        =   150
         Top             =   1305
         Width           =   1700
      End
      Begin VB.TextBox txtHeatingValue 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   -75000
         TabIndex        =   149
         Top             =   3390
         Width           =   2500
      End
      Begin VB.TextBox txtBulkdensity 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         ForeColor       =   &H000000C0&
         Height          =   315
         Left            =   -73500
         TabIndex        =   148
         Top             =   2655
         Width           =   1000
      End
      Begin VB.TextBox txtviscosity 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         ForeColor       =   &H000000C0&
         Height          =   315
         Left            =   -73500
         TabIndex        =   147
         Top             =   2295
         Width           =   1000
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFC0FF&
         Caption         =   "Analysis Type"
         Height          =   1395
         Left            =   -68160
         TabIndex        =   142
         Top             =   750
         Width           =   4095
         Begin VB.TextBox txtAnalysDesc 
            Appearance      =   0  'Flat
            ForeColor       =   &H00FF0000&
            Height          =   645
            Left            =   1200
            Locked          =   -1  'True
            ScrollBars      =   1  'Horizontal
            TabIndex        =   146
            Top             =   660
            Width           =   2715
         End
         Begin VB.OptionButton OptAnalysType3 
            BackColor       =   &H00FFC0FF&
            Caption         =   "Others"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   300
            TabIndex        =   145
            Top             =   720
            Value           =   -1  'True
            Width           =   795
         End
         Begin VB.OptionButton OptAnalysType2 
            BackColor       =   &H00FFC0FF&
            Caption         =   "Finger Print"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   2280
            TabIndex        =   144
            Top             =   360
            Width           =   1275
         End
         Begin VB.OptionButton OptAnalysType1 
            BackColor       =   &H00FFC0FF&
            Caption         =   "Full Analysis"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   300
            TabIndex        =   143
            Top             =   360
            Width           =   1275
         End
      End
      Begin VB.CheckBox chkBottle 
         Caption         =   "ขวดแก้ว : Bottle ขนาด"
         Height          =   255
         Left            =   -74280
         TabIndex        =   141
         Top             =   2385
         Width           =   1995
      End
      Begin VB.TextBox txtBottle 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   -71640
         Locked          =   -1  'True
         TabIndex        =   140
         Top             =   2325
         Width           =   1575
      End
      Begin VB.CheckBox chkPaperBox 
         Caption         =   "กล่องกระดาษ ขนาด"
         Height          =   255
         Left            =   -68820
         TabIndex        =   139
         Top             =   1995
         Width           =   2565
      End
      Begin VB.TextBox txtPaperBox 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   -66120
         Locked          =   -1  'True
         TabIndex        =   138
         Top             =   1965
         Width           =   1575
      End
      Begin VB.CheckBox chkPlasticGallon 
         Caption         =   "แกนลอนพลาสติก : Plastic Gallon ขนาด"
         Height          =   255
         Left            =   -74280
         TabIndex        =   137
         Top             =   2025
         Width           =   3135
      End
      Begin VB.TextBox txtPlasticGallon 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   -71100
         Locked          =   -1  'True
         TabIndex        =   136
         Top             =   1965
         Width           =   1035
      End
      Begin VB.CheckBox chkPlasticDrum 
         Caption         =   "ถังพลาสติก : Plastic Drum ขนาด"
         Height          =   255
         Left            =   -68820
         TabIndex        =   135
         Top             =   1635
         Width           =   2565
      End
      Begin VB.TextBox txtPlasticDrum 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   -66120
         Locked          =   -1  'True
         TabIndex        =   134
         Top             =   1605
         Width           =   1575
      End
      Begin VB.CheckBox chkPlasticBag 
         Caption         =   "ถุงพลาสติก : Plastic Bag ขนาด"
         Height          =   255
         Left            =   -68820
         TabIndex        =   133
         Top             =   1275
         Width           =   2565
      End
      Begin VB.TextBox txtPlasticBag 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   -66120
         Locked          =   -1  'True
         TabIndex        =   132
         Top             =   1245
         Width           =   1575
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFC0FF&
         Caption         =   "เลือกเฉพาะ 075"
         Enabled         =   0   'False
         ForeColor       =   &H00800000&
         Height          =   1275
         Left            =   -68160
         TabIndex        =   128
         Top             =   2190
         Width           =   4095
         Begin VB.OptionButton OptBlend0 
            BackColor       =   &H00FFC0FF&
            Caption         =   "ไม่เลือก"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   300
            TabIndex        =   131
            Top             =   360
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton OptBlend1 
            BackColor       =   &H00FFC0FF&
            Caption         =   "ต้องทำ Waste Blending"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   300
            TabIndex        =   130
            Top             =   660
            Width           =   2235
         End
         Begin VB.OptionButton OptBlend2 
            BackColor       =   &H00FFC0FF&
            Caption         =   "ไม่ต้องทำ Waste Blending"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   300
            TabIndex        =   129
            Top             =   960
            Width           =   2235
         End
      End
      Begin VB.CommandButton cmdAttach 
         Caption         =   "..."
         Height          =   315
         Index           =   0
         Left            =   -67440
         TabIndex        =   127
         Top             =   840
         Width           =   435
      End
      Begin VB.CommandButton cmdCls 
         Caption         =   "ลบไฟล์"
         Height          =   315
         Index           =   0
         Left            =   -65940
         TabIndex        =   126
         Top             =   840
         Width           =   795
      End
      Begin VB.CommandButton cmdOpenFile 
         Caption         =   "เปิดไฟล์"
         Height          =   315
         Index           =   0
         Left            =   -66900
         TabIndex        =   125
         Top             =   840
         Width           =   795
      End
      Begin VB.CommandButton cmdOpenFile 
         Caption         =   "เปิดไฟล์"
         Height          =   315
         Index           =   1
         Left            =   -66900
         TabIndex        =   124
         Top             =   1200
         Width           =   795
      End
      Begin VB.CommandButton cmdOpenFile 
         Caption         =   "เปิดไฟล์"
         Height          =   315
         Index           =   2
         Left            =   -66900
         TabIndex        =   123
         Top             =   1560
         Width           =   795
      End
      Begin VB.CommandButton cmdOpenFile 
         Caption         =   "เปิดไฟล์"
         Height          =   315
         Index           =   3
         Left            =   -66900
         TabIndex        =   122
         Top             =   1920
         Width           =   795
      End
      Begin VB.CommandButton cmdCls 
         Caption         =   "ลบไฟล์"
         Height          =   315
         Index           =   1
         Left            =   -65940
         TabIndex        =   121
         Top             =   1200
         Width           =   795
      End
      Begin VB.CommandButton cmdCls 
         Caption         =   "ลบไฟล์"
         Height          =   315
         Index           =   2
         Left            =   -65940
         TabIndex        =   120
         Top             =   1560
         Width           =   795
      End
      Begin VB.CommandButton cmdCls 
         Caption         =   "ลบไฟล์"
         Height          =   315
         Index           =   3
         Left            =   -65940
         TabIndex        =   119
         Top             =   1920
         Width           =   795
      End
      Begin VB.CommandButton cmdAttach 
         Caption         =   "..."
         Height          =   315
         Index           =   1
         Left            =   -67440
         TabIndex        =   118
         Top             =   1200
         Width           =   435
      End
      Begin VB.CommandButton cmdAttach 
         Caption         =   "..."
         Height          =   315
         Index           =   2
         Left            =   -67440
         TabIndex        =   117
         Top             =   1560
         Width           =   435
      End
      Begin VB.CommandButton cmdAttach 
         Caption         =   "..."
         Height          =   315
         Index           =   3
         Left            =   -67440
         TabIndex        =   116
         Top             =   1920
         Width           =   435
      End
      Begin VB.ComboBox cboTreatmentType 
         BackColor       =   &H00C0FFC0&
         Height          =   315
         ItemData        =   "frWasteData.frx":0360
         Left            =   -66000
         List            =   "frWasteData.frx":036A
         Style           =   2  'Dropdown List
         TabIndex        =   115
         Top             =   825
         Visible         =   0   'False
         Width           =   1995
      End
      Begin VB.CheckBox chkYard 
         Caption         =   "ลานกอง : Stock yard"
         Height          =   255
         Left            =   -68820
         TabIndex        =   114
         Top             =   2385
         Width           =   1845
      End
      Begin VB.CheckBox chkTank 
         Caption         =   "บ่อรวบรวม : Collecting Tank"
         Height          =   255
         Left            =   -74280
         TabIndex        =   113
         Top             =   2745
         Width           =   2355
      End
      Begin VB.TextBox txtOther2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   -68040
         Locked          =   -1  'True
         TabIndex        =   112
         Top             =   2685
         Width           =   3495
      End
      Begin VB.CheckBox chkOther2 
         Caption         =   "อื่นๆ"
         Height          =   255
         Left            =   -68820
         TabIndex        =   111
         Top             =   2745
         Width           =   675
      End
      Begin VB.TextBox txtDrumDesc 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   -71640
         Locked          =   -1  'True
         TabIndex        =   110
         Top             =   1605
         Width           =   1575
      End
      Begin VB.CheckBox chkDrum 
         Caption         =   "ถังเหล็ก : Metal Drum ขนาด"
         Height          =   255
         Left            =   -74280
         TabIndex        =   109
         Top             =   1665
         Width           =   2475
      End
      Begin VB.TextBox txtBag 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   -71640
         Locked          =   -1  'True
         TabIndex        =   108
         Top             =   1245
         Width           =   1575
      End
      Begin VB.CheckBox chkBag 
         Caption         =   "ถุงขนาดใหญ่ : Big Bag ขนาด"
         Height          =   255
         Left            =   -74280
         TabIndex        =   107
         Top             =   1305
         Width           =   2475
      End
      Begin VB.ComboBox cboIndoor 
         BackColor       =   &H00C0E0FF&
         Height          =   315
         Left            =   -70440
         Style           =   2  'Dropdown List
         TabIndex        =   106
         Top             =   825
         Width           =   3015
      End
      Begin VB.ComboBox cboStoreType 
         BackColor       =   &H0080C0FF&
         Height          =   315
         ItemData        =   "frWasteData.frx":0380
         Left            =   -74280
         List            =   "frWasteData.frx":0382
         Style           =   2  'Dropdown List
         TabIndex        =   105
         Top             =   825
         Width           =   3015
      End
      Begin VB.Frame Frame2 
         Caption         =   "คุณสมบัติของเสียอันตราย"
         ForeColor       =   &H00C00000&
         Height          =   2535
         Left            =   -74880
         TabIndex        =   80
         Top             =   1290
         Width           =   10935
         Begin VB.CheckBox chkInflammable 
            Caption         =   "ไวไฟ(Inflammable)"
            Height          =   255
            Left            =   1500
            TabIndex        =   309
            Top             =   300
            Width           =   1755
         End
         Begin VB.TextBox txtpH 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            ForeColor       =   &H000000C0&
            Height          =   315
            Left            =   8580
            TabIndex        =   100
            Top             =   275
            Width           =   1000
         End
         Begin VB.CheckBox chkCorrosive 
            Caption         =   "ค่า pH :"
            Height          =   255
            Left            =   7680
            TabIndex        =   99
            Top             =   300
            Width           =   1000
         End
         Begin VB.TextBox txtFlashPoint 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            ForeColor       =   &H000000C0&
            Height          =   315
            Left            =   6420
            TabIndex        =   98
            Top             =   255
            Width           =   500
         End
         Begin VB.CheckBox chkFlame 
            Caption         =   "  ติดไฟ ( Flammability )"
            Height          =   255
            Left            =   3480
            TabIndex        =   97
            Top             =   300
            Width           =   2000
         End
         Begin VB.TextBox txtCharDesc 
            Appearance      =   0  'Flat
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   2040
            TabIndex        =   96
            Top             =   0
            Visible         =   0   'False
            Width           =   1035
         End
         Begin VB.CheckBox chkHospital 
            Caption         =   "ขยะติดเชื้อ : Hospital Waste"
            Height          =   255
            Left            =   300
            TabIndex        =   95
            Top             =   2100
            Width           =   2295
         End
         Begin VB.TextBox txtOther 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   3780
            TabIndex        =   94
            Top             =   2040
            Width           =   3315
         End
         Begin VB.CheckBox chkOther 
            Caption         =   "อื่นๆ ระบุ"
            Height          =   255
            Left            =   2760
            TabIndex        =   93
            Top             =   2100
            Width           =   975
         End
         Begin VB.TextBox txtPesticide 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   2760
            TabIndex        =   92
            Top             =   1680
            Width           =   4335
         End
         Begin VB.CheckBox chkPesticide 
            Caption         =   "สารกำจัดศัตรูพืช/วัชพืช   ระบุ"
            Height          =   255
            Left            =   300
            TabIndex        =   91
            Top             =   1740
            Width           =   2415
         End
         Begin VB.TextBox txtRadioActive 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   2280
            TabIndex        =   90
            Top             =   1320
            Width           =   4815
         End
         Begin VB.CheckBox chkRadioActive 
            Caption         =   "กัมมันตภาพรังสี   ระบุ"
            Height          =   255
            Left            =   300
            TabIndex        =   89
            Top             =   1380
            Width           =   1815
         End
         Begin VB.ComboBox cboToxic 
            BackColor       =   &H00FFFFC0&
            Height          =   315
            ItemData        =   "frWasteData.frx":0384
            Left            =   2280
            List            =   "frWasteData.frx":038E
            Style           =   2  'Dropdown List
            TabIndex        =   88
            Top             =   960
            Width           =   1875
         End
         Begin VB.TextBox txtToxicDesc 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   4560
            TabIndex        =   87
            Top             =   960
            Width           =   2535
         End
         Begin VB.CheckBox chkToxic 
            Caption         =   "มีสารเป็นพิษต่อร่างกาย"
            Height          =   255
            Left            =   300
            TabIndex        =   86
            Top             =   1020
            Width           =   1995
         End
         Begin VB.TextBox txtCancerDesc 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   5520
            TabIndex        =   85
            Top             =   600
            Width           =   1575
         End
         Begin VB.CheckBox chkCancer 
            Caption         =   "สารก่อมะเร็ง   ระบุ"
            Height          =   255
            Left            =   3480
            TabIndex        =   84
            Top             =   600
            Width           =   1575
         End
         Begin VB.TextBox txtReactDesc 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   1620
            TabIndex        =   83
            Top             =   600
            Width           =   1755
         End
         Begin VB.CheckBox chkReact 
            Caption         =   "ทำปฏิกริยากับ"
            Height          =   255
            Left            =   300
            TabIndex        =   82
            Top             =   660
            Width           =   1335
         End
         Begin VB.CheckBox chkExplode 
            Caption         =   "ระเบิดได้"
            Height          =   255
            Left            =   300
            TabIndex        =   81
            Top             =   300
            Width           =   1095
         End
         Begin WasteDataForm.ctlGrid grKorOrNo 
            Height          =   1635
            Left            =   7200
            TabIndex        =   101
            Top             =   810
            Width           =   3495
            _extentx        =   6165
            _extenty        =   2884
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "C"
            Height          =   195
            Index           =   15
            Left            =   7020
            TabIndex        =   104
            Top             =   360
            Width           =   135
         End
         Begin VB.Label Label15 
            Caption         =   "จุดวาบไฟ :"
            Height          =   255
            Left            =   5640
            TabIndex        =   103
            Top             =   300
            Width           =   1305
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ระบุ"
            Height          =   195
            Index           =   16
            Left            =   4200
            TabIndex        =   102
            Top             =   1020
            Width           =   315
         End
      End
      Begin VB.ComboBox cboWasteCharDesc 
         BackColor       =   &H00C0E0FF&
         Height          =   315
         ItemData        =   "frWasteData.frx":03BB
         Left            =   -72720
         List            =   "frWasteData.frx":03BD
         Style           =   2  'Dropdown List
         TabIndex        =   79
         Top             =   870
         Width           =   1695
      End
      Begin VB.ComboBox cboWasteChar 
         BackColor       =   &H00C0C0FF&
         Height          =   315
         Left            =   -74880
         Style           =   2  'Dropdown List
         TabIndex        =   78
         Top             =   870
         Width           =   1935
      End
      Begin MSComDlg.CommonDialog OpenDialog 
         Left            =   -66600
         Top             =   3420
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DialogTitle     =   "เลือกไฟล์"
      End
      Begin MSComctlLib.ListView lvNewTreatment 
         Height          =   2655
         Left            =   -74760
         TabIndex        =   221
         Top             =   810
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
         BackColor       =   -2147483624
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "วิธีการกำจัดที่ควรใช้"
            Object.Width           =   11298
         EndProperty
      End
      Begin VB.Label Label51 
         Caption         =   "Lab code :"
         Height          =   195
         Left            =   -74760
         TabIndex        =   292
         Top             =   480
         Width           =   780
      End
      Begin VB.Label Label49 
         BackColor       =   &H00FFFFC0&
         Caption         =   "ค่าความร้อน (NCV)"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   -75000
         TabIndex        =   289
         Top             =   3690
         Width           =   2505
      End
      Begin VB.Label Label63 
         Alignment       =   1  'Right Justify
         Caption         =   "SS(Suspended solids) :"
         Height          =   255
         Left            =   -69375
         TabIndex        =   287
         Top             =   2520
         Width           =   1635
      End
      Begin VB.Label Label46 
         Alignment       =   1  'Right Justify
         Caption         =   "Acidity :"
         Height          =   255
         Left            =   -69195
         TabIndex        =   286
         Top             =   2820
         Width           =   855
      End
      Begin VB.Label Label38 
         AutoSize        =   -1  'True
         Caption         =   "Alkalinity :"
         Height          =   195
         Left            =   -66075
         TabIndex        =   285
         Top             =   510
         Width           =   705
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "เหล็ก (Fe) :"
         Height          =   195
         Left            =   -66090
         TabIndex        =   275
         Top             =   2610
         Width           =   780
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         Caption         =   "ความเข้มข้นกรด/ด่าง :"
         Height          =   195
         Left            =   -66360
         TabIndex        =   274
         Top             =   2970
         Width           =   1575
      End
      Begin VB.Label Label44 
         Caption         =   "Lab code :"
         Height          =   195
         Left            =   -74685
         TabIndex        =   273
         Top             =   450
         Width           =   780
      End
      Begin VB.Label Label43 
         Caption         =   "ทองแดง (Cu) :"
         Height          =   255
         Left            =   -66315
         TabIndex        =   272
         Top             =   2340
         Width           =   1005
      End
      Begin VB.Label Label42 
         Caption         =   "สังกะสี (Zn) :"
         Height          =   255
         Left            =   -66180
         TabIndex        =   271
         Top             =   2055
         Width           =   900
      End
      Begin VB.Label Label41 
         Caption         =   "แมงกานีส (Mn) :"
         Height          =   255
         Left            =   -66480
         TabIndex        =   270
         Top             =   1755
         Width           =   1200
      End
      Begin VB.Label Label40 
         Caption         =   "นิเกิล (Ni) :"
         Height          =   255
         Left            =   -66120
         TabIndex        =   269
         Top             =   1455
         Width           =   795
      End
      Begin VB.Label Label39 
         Caption         =   "TDS :"
         Height          =   255
         Left            =   -65775
         TabIndex        =   268
         Top             =   1155
         Width           =   465
      End
      Begin VB.Label Label31 
         Caption         =   "สถานะ (Physical state)"
         Height          =   300
         Left            =   -74880
         TabIndex        =   267
         Top             =   1530
         Width           =   2205
      End
      Begin VB.Label lbAnalysisNo 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "."
         ForeColor       =   &H8000000E&
         Height          =   270
         Left            =   -73800
         TabIndex        =   266
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label27 
         Caption         =   "ผู้เก็บตัวอย่าง  :                          เลขที่วิเคราะห์ :"
         Height          =   645
         Left            =   -74880
         TabIndex        =   265
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label lbAnalysisDate 
         Caption         =   "0000/00/00"
         Height          =   255
         Left            =   -65400
         TabIndex        =   264
         Top             =   3930
         Width           =   975
      End
      Begin VB.Label Label29 
         Caption         =   "วันที่วิเคราะห์ :"
         Height          =   255
         Left            =   -66480
         TabIndex        =   263
         Top             =   3930
         Width           =   1095
      End
      Begin VB.Label lbReceivedDate 
         Caption         =   "0000/00/00"
         Height          =   255
         Left            =   -68160
         TabIndex        =   262
         Top             =   3930
         Width           =   1005
      End
      Begin VB.Label Label28 
         Caption         =   "วันที่รับตัวอย่าง :"
         Height          =   255
         Left            =   -69360
         TabIndex        =   261
         Top             =   3930
         Width           =   1095
      End
      Begin VB.Label lbSameplingDate 
         BackColor       =   &H00C0FFFF&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -72360
         TabIndex        =   260
         Top             =   3285
         Width           =   1605
      End
      Begin VB.Label Label26 
         Caption         =   "วันที่เก็บตัวอย่าง กากฯ  :"
         ForeColor       =   &H00FF00FF&
         Height          =   375
         Left            =   -74280
         TabIndex        =   259
         Top             =   3285
         Width           =   1995
      End
      Begin VB.Label Label25 
         Caption         =   "ขั้นตอนการดำเนินการ วิเคราะห์ :"
         Height          =   255
         Left            =   -69360
         TabIndex        =   258
         Top             =   3615
         Width           =   3000
      End
      Begin VB.Label Label24 
         Caption         =   "หมายเหตุ :"
         Height          =   255
         Left            =   -69240
         TabIndex        =   257
         Top             =   3285
         Width           =   795
      End
      Begin VB.Label Label23 
         Caption         =   "ตะกั่ว (Pb) :"
         Height          =   255
         Left            =   -66150
         TabIndex        =   256
         Top             =   810
         Width           =   795
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         Caption         =   "ปรอท (Hg) :"
         Height          =   255
         Left            =   -69180
         TabIndex        =   255
         Top             =   2220
         Width           =   855
      End
      Begin VB.Label Label21 
         Caption         =   "โครเมียม (Cr) :"
         Height          =   255
         Left            =   -69360
         TabIndex        =   254
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         Caption         =   "แคดเมียม (Cd) :"
         Height          =   255
         Left            =   -69480
         TabIndex        =   253
         Top             =   1620
         Width           =   1155
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         Caption         =   "อะเซนิก (As) :"
         Height          =   255
         Left            =   -69360
         TabIndex        =   252
         Top             =   1320
         Width           =   1005
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         Caption         =   "ปริมาณของแข็ง :"
         Height          =   300
         Index           =   1
         Left            =   -72420
         TabIndex        =   251
         Top             =   3510
         Width           =   1185
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         Caption         =   "ปริมาณสลัดจ์ :"
         Height          =   300
         Index           =   0
         Left            =   -72240
         TabIndex        =   250
         Top             =   3240
         Width           =   1005
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         Caption         =   "ซัลเฟอร์ (S) :"
         Height          =   255
         Index           =   1
         Left            =   -72240
         TabIndex        =   249
         Top             =   2880
         Width           =   1005
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         Caption         =   "ฟลูออรีน (F) :"
         Height          =   255
         Index           =   0
         Left            =   -72240
         TabIndex        =   248
         Top             =   2565
         Width           =   1005
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         Caption         =   "คลอรีน (CL) :"
         Height          =   315
         Index           =   3
         Left            =   -72240
         TabIndex        =   247
         Top             =   2250
         Width           =   1005
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         Caption         =   "ปริมาณเถ้า :"
         Height          =   315
         Index           =   2
         Left            =   -72240
         TabIndex        =   246
         Top             =   1965
         Width           =   1005
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         Caption         =   "ปริมาณน้ำ :"
         Height          =   315
         Index           =   1
         Left            =   -72240
         TabIndex        =   245
         Top             =   1650
         Width           =   1005
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         Caption         =   "ค่าความชื้น :"
         Height          =   315
         Index           =   0
         Left            =   -72240
         TabIndex        =   244
         Top             =   1320
         Width           =   1005
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FFFFC0&
         Caption         =   "ค่าความร้อน (GCV)"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   -75000
         TabIndex        =   243
         Top             =   3135
         Width           =   2505
      End
      Begin VB.Label Label13 
         Caption         =   "ความหนาแน่น   :"
         Height          =   255
         Left            =   -75000
         TabIndex        =   242
         Top             =   2685
         Width           =   1455
      End
      Begin VB.Label Label12 
         Caption         =   "ความหนืด          :"
         Height          =   255
         Left            =   -75000
         TabIndex        =   241
         Top             =   2325
         Width           =   1500
      End
      Begin VB.Label Label7 
         Caption         =   "ลิตร"
         Height          =   195
         Index           =   5
         Left            =   -69900
         TabIndex        =   240
         Top             =   2385
         Width           =   315
      End
      Begin VB.Label Label7 
         Caption         =   "ลิตร"
         Height          =   195
         Index           =   4
         Left            =   -69900
         TabIndex        =   239
         Top             =   2025
         Width           =   315
      End
      Begin VB.Label Label7 
         Caption         =   "ลิตร"
         Height          =   195
         Index           =   3
         Left            =   -64380
         TabIndex        =   238
         Top             =   1665
         Width           =   315
      End
      Begin VB.Label Label7 
         Caption         =   "ลิตร"
         Height          =   195
         Index           =   2
         Left            =   -69900
         TabIndex        =   237
         Top             =   1665
         Width           =   315
      End
      Begin VB.Label Label7 
         Caption         =   "กก."
         Height          =   195
         Index           =   1
         Left            =   -64380
         TabIndex        =   236
         Top             =   1305
         Width           =   315
      End
      Begin VB.Label Label7 
         Caption         =   "กก."
         Height          =   195
         Index           =   0
         Left            =   -69900
         TabIndex        =   235
         Top             =   1305
         Width           =   315
      End
      Begin VB.Label Label5 
         Caption         =   "ไฟล์แนบที่ 1"
         Height          =   195
         Left            =   -74760
         TabIndex        =   234
         Top             =   900
         Width           =   975
      End
      Begin VB.Label lblAttach 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   -73740
         TabIndex        =   233
         Top             =   900
         Width           =   6135
      End
      Begin VB.Label Label6 
         Caption         =   "ไฟล์แนบที่ 2"
         Height          =   195
         Left            =   -74760
         TabIndex        =   232
         Top             =   1260
         Width           =   975
      End
      Begin VB.Label Label8 
         Caption         =   "ไฟล์แนบที่ 3"
         Height          =   195
         Left            =   -74760
         TabIndex        =   231
         Top             =   1620
         Width           =   975
      End
      Begin VB.Label Label10 
         Caption         =   "ไฟล์แนบที่ 4"
         Height          =   195
         Left            =   -74760
         TabIndex        =   230
         Top             =   1980
         Width           =   975
      End
      Begin VB.Label lblAttach 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   -73740
         TabIndex        =   229
         Top             =   1260
         Width           =   6135
      End
      Begin VB.Label lblAttach 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   -73740
         TabIndex        =   228
         Top             =   1620
         Width           =   6135
      End
      Begin VB.Label lblAttach 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   -73740
         TabIndex        =   227
         Top             =   1980
         Width           =   6135
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "สถานที่กำจัดเดิม"
         Height          =   195
         Index           =   20
         Left            =   135
         TabIndex        =   226
         Top             =   -4095
         Width           =   1155
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "รูปแบบการบำบัด"
         Height          =   195
         Index           =   31
         Left            =   -67260
         TabIndex        =   225
         Top             =   885
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
         TabIndex        =   224
         Top             =   1305
         Width           =   375
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "10.2)"
         Height          =   195
         Index           =   18
         Left            =   -70920
         TabIndex        =   223
         Top             =   885
         Width           =   375
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "10.1)"
         Height          =   195
         Index           =   17
         Left            =   -74760
         TabIndex        =   222
         Top             =   885
         Width           =   375
      End
   End
   Begin VB.Label Label55 
      Caption         =   "Lab code :"
      Height          =   195
      Left            =   0
      TabIndex        =   294
      Top             =   0
      Width           =   780
   End
   Begin VB.Label Label32 
      BackColor       =   &H00FFFFC0&
      Caption         =   "ค่าความร้อน (GCV)"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   0
      TabIndex        =   288
      Top             =   0
      Width           =   2505
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "วันที่ตรวจรับ :"
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   34
      Left            =   495
      TabIndex        =   59
      Top             =   2385
      Width           =   975
   End
   Begin VB.Label Label62 
      BackColor       =   &H00C0FFC0&
      Caption         =   "วันที่เริ่ม :                                     วันหมดอายุ : "
      Height          =   195
      Left            =   2655
      TabIndex        =   74
      Top             =   2700
      Visible         =   0   'False
      Width           =   3300
   End
   Begin VB.Label Label11 
      BackColor       =   &H00C0FFC0&
      Caption         =   " Waste   ปี :"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   480
      TabIndex        =   56
      Top             =   1140
      Width           =   855
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
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
      Left            =   3960
      TabIndex        =   47
      Top             =   9750
      Visible         =   0   'False
      Width           =   3075
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
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
      Height          =   375
      Left            =   3960
      TabIndex        =   46
      Top             =   9390
      Visible         =   0   'False
      Width           =   3075
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
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
      Left            =   1560
      TabIndex        =   45
      Top             =   2400
      Visible         =   0   'False
      Width           =   2685
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ประเภทอุตสาหกรรม"
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   30
      Left            =   60
      TabIndex        =   44
      Top             =   2100
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
      TabIndex        =   43
      Top             =   1140
      Width           =   555
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   $"frWasteData.frx":03BF
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   27
      Left            =   135
      TabIndex        =   42
      Top             =   1800
      Width           =   9420
   End
   Begin VB.Line Line1 
      X1              =   30
      X2              =   11400
      Y1              =   9345
      Y2              =   9345
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "เจ้าหน้าที่ Envi                                                                  ผลสรุป"
      Height          =   195
      Index           =   22
      Left            =   4290
      TabIndex        =   41
      Top             =   9015
      Width           =   4455
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "เจ้าหน้าที่การตลาด"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   21
      Left            =   90
      TabIndex        =   40
      Top             =   9015
      Width           =   1290
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "กระบวนการที่ทำให้เกิดของเสีย"
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   12
      Left            =   420
      TabIndex        =   39
      Top             =   4125
      Width           =   2145
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ชื่อกากของเสีย"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   7
      Left            =   360
      TabIndex        =   35
      Top             =   1455
      Width           =   1035
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "ประเภทของเสีย                                              เลขที่สัญญา"
      ForeColor       =   &H00C00000&
      Height          =   600
      Index           =   6
      Left            =   8415
      TabIndex        =   34
      Top             =   1095
      Width           =   1320
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ข้อ"
      Height          =   195
      Index           =   5
      Left            =   3960
      TabIndex        =   33
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
      TabIndex        =   32
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
Dim CKuser_ As CKuser
Dim curEUCode1_ As BWGEUCode1
Dim curEUCode2_ As BWGEUCode2
Dim curEUCode3_ As BWGEUCode3
Dim curMktStaff_ As BWGCompanyStaff
Dim curEnviStaff_ As BWGCompanyStaff
Dim curWasteCase_ As BWGWasteCase
Dim myWasteDataCR_ As BWGWasteDataCR
Dim myWasteDataLab_ As BWGWasteDataLAB
Dim curLabA_ As Collection
Dim curLabB_ As Collection
Dim curTastType_ As Collection

Dim colTreatment_ As Collection
Dim callStaffType_ As String
Dim curIndustryType_ As BWGIndustryType_New
Dim StrGrp As String
Dim isNew As Boolean
Dim isEnv As Boolean, isIT As Boolean
Dim isLab As Boolean
Private Type AttacheFile
    FileName As String
    FullPath As String
End Type
Dim Attaches() As AttacheFile
Private Type AttachesPics
    FileName As String
    FullPath As String
End Type
Dim AttachesPic(2) As AttachesPics
Dim isSave As Boolean
Dim WasteID As String
Dim WasteY As String
Public Accepdate As String
Public ReceivedDate As String
Public AnalysisDate As String
Public SamplingDate As String
Public LicenseStartDate As String
Public LicenseExpDate As String
Dim PicPath As String
Dim PicPath1 As String
Dim DefaultPath1 As String
Dim rs As ADODB.Recordset
Dim tmpCol As Collection
Dim Lbl As String
Private Function AttacheFiles(ByVal Index As Integer, ByRef FileName As String, ByRef FullPath As String) As String
Dim StrFileName As String
Dim StrPath As String, StrFullPath As String
    OpenDialog.Filter = "PDF Files (*.pdf)|*.pdf|Pictures (*.jpg, *.bmp, *.gif, *.png| *.jpg; *.bmp; *.gif; *.png|" & _
    "World Document (*.doc, *.docx)|*.doc; *.docx|Zip (*.rar)|*.rar|Work Sheet (*.xls, *.slsx)|*.xls; *.xlsx|Visio (*.vsd)|*.vsd| all file (*.*)|*.*"
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
Private Function AttachePic(ByVal Index As Integer, ByRef FileName As String, ByRef FullPath As String) As String
Dim StrFileName As String
Dim StrPath As String, StrFullPath As String
    OpenDialog.Filter = "PDF Files (*.pdf)|*.pdf|Pictures (*.jpg, *.bmp, *.gif, *.png| *.jpg; *.bmp; *.gif; *.png|"
    OpenDialog.FileName = ""
    OpenDialog.ShowOpen
    If OpenDialog.FileName = "" Then
        AttachePic = ""
        Exit Function
    End If
    StrFullPath = OpenDialog.FileName
    StrFileName = OpenDialog.FileTitle
    StrPath = Mid(StrFullPath, 1, Len(StrFullPath) - Len(StrFileName))
    
    AttachePic = StrFileName
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
Public Sub setLicenseStartDate()
    txtStartDate.Text = FormatYMD_to_DMY(LicenseStartDate, "/", "/") + " ถึง " + FormatYMD_to_DMY(LicenseExpDate, "/", "/")
    myWasteDataCR_.LicenseStartDate = LicenseStartDate
    myWasteDataCR_.LicenseExpDate = LicenseExpDate
    myWasteDataCR_.isChanged = True
End Sub

Public Sub SetReceivedDate()
    lbReceivedDate.Caption = ReceivedDate
    myWasteDataLab_.ReceivedDate = Trim(lbReceivedDate.Caption)
    myWasteDataLab_.isChanged = True
End Sub
Public Sub SetAnalysisDate()
    lbAnalysisDate.Caption = AnalysisDate
    myWasteDataLab_.AnalysisDate = Trim(lbAnalysisDate.Caption)
    myWasteDataLab_.isChanged = True
End Sub
Public Sub SetSamplingDate()
    lbSameplingDate.Caption = SamplingDate
    myWasteDataLab_.SamplingDate = Trim(lbSameplingDate.Caption)
    myWasteDataLab_.isChanged = True
End Sub

Private Sub cboAnalysis_Click()
 On Error Resume Next
  Select Case cboAnalysis.ListIndex
    Case 0
        myWasteDataLab_.LBstatus = "D" 'ยังไม่ได้ดำเนินการใดๆ
    Case 1
        myWasteDataLab_.LBstatus = "S" ' ส่งวิเคราะห์
    Case 2
        myWasteDataLab_.LBstatus = "R" ' รอวิเคราะห์
    Case 3
        myWasteDataLab_.LBstatus = "A" ' วิเคราะห์ เสร็จเรียบร้อย
    End Select
End Sub

Private Sub cboCompany_Click()
On Error Resume Next
'If cboWasteType.ListIndex = 8 Then Exit Sub
'If cboWasteType.ListIndex = 9 Then Exit Sub
    If cboCompany.ListIndex > 0 Then
        myWasteDataCR_.companyID = CurrentUser.WasteCompany(cboCompany.ListIndex).ID
'        If cboCompany.ListIndex = 2 Then
'            cboWasteType.ListIndex = 5
'        ElseIf cboCompany.ListIndex = 3 Then
'            cboWasteType.ListIndex = 4
'        End If
'        myWasteDataCR_.WasteType = cboWasteType.Text
    Else
        myWasteDataCR_.companyID = ""
'        myWasteDataCR_.WasteType = "N/A"
    End If
    myWasteDataCR_.isChanged = True
End Sub

Private Sub cboCompany_Validate(Cancel As Boolean)
    If cboCompany.ListIndex > 0 Then
        myWasteDataCR_.companyID = CurrentUser.WasteCompany(cboCompany.ListIndex).ID
    Else
        myWasteDataCR_.companyID = ""
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

'Private Sub cboMaxUnit_Validate(Cancel As Boolean)
'    Select Case cboMaxUnit.ListIndex
'        Case 0
'            myWasteDataCR_.MaxUnit = "T" 'ตัน
'        Case 1
'            myWasteDataCR_.MaxUnit = "L" 'ลิตร
'    End Select
'End Sub

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
            myWasteDataCR_.WasteRateTime = "M" 'เดือน
        Case 1
            myWasteDataCR_.WasteRateTime = "Y" 'ปี
    End Select
End Sub

Private Sub cboTimeRate_Validate(Cancel As Boolean)
    Select Case cboTimeRate.ListIndex
        Case 0
            myWasteDataCR_.WasteRateTime = "M" 'เดือน
        Case 1
            myWasteDataCR_.WasteRateTime = "Y" 'ปี
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
            cmdApprove.Visible = True
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


Private Sub cboWasteCase_Click()
 If cboWasteCase.Text = "RDF" Then
        myWasteDataCR_.WasteCaseID = "A1"
        DtpStart.Visible = False
        DtpEnd.Visible = False
        DtpStart.Enabled = False
        DtpEnd.Enabled = False
        Label62.Visible = False
 ElseIf cboWasteCase.Text = "เวสต์น้ำท่วม กรอ." Then
        myWasteDataCR_.WasteCaseID = "10EE3782390242959F296DA722542573"
        DtpStart.Visible = False
        DtpEnd.Visible = False
        DtpStart.Enabled = False
        DtpEnd.Enabled = False
        Label62.Visible = False
 ElseIf cboWasteCase.Text = "เวสต์ทิ้ง" Then
        myWasteDataCR_.companyID = "AF543EA5BC1649259AA7955CE4D3C7CD"
        DtpStart.Visible = False
        DtpEnd.Visible = False
        DtpStart.Enabled = False
        DtpEnd.Enabled = False
         Label62.Visible = False
 ElseIf cboWasteCase.Text = "Haz Cleaning" Then
        myWasteDataCR_.companyID = "SHazcleening101"
        DtpStart.Visible = False
        DtpEnd.Visible = False
        DtpStart.Enabled = False
        DtpEnd.Enabled = False
         Label62.Visible = False
   ElseIf cboWasteCase.Text = "Non ราคาถูก" Then
        myWasteDataCR_.companyID = "SNonLow101"
        DtpStart.Visible = False
        DtpEnd.Visible = False
        DtpStart.Enabled = False
        DtpEnd.Enabled = False
         Label62.Visible = False
   ElseIf cboWasteCase.Text = "ปกติ" Then
    myWasteDataCR_.companyID = ""
        DtpStart.Visible = False
        DtpEnd.Visible = False
         DtpStart.Enabled = False
        DtpEnd.Enabled = False
        Label62.Visible = False
    ElseIf cboWasteCase.Text = "งาน Job" Then
        myWasteDataCR_.WasteCaseID = "1"
        DtpStart.Visible = True
        DtpEnd.Visible = True
        DtpStart.Enabled = True
        DtpEnd.Enabled = True
         Label62.Visible = True
    Else
        DtpStart.Visible = False
        DtpEnd.Visible = False
        DtpStart.Enabled = False
        DtpEnd.Enabled = False
         Label62.Visible = False
End If
    myWasteDataCR_.isChanged = True
End Sub

Private Sub cboWasteCase_Validate(Cancel As Boolean)
 If cboWasteCase.Text = "RDF" Then
        myWasteDataCR_.WasteCaseID = "A1"
        DtpStart.Visible = False
        DtpEnd.Visible = False
         Label62.Visible = False
 ElseIf cboWasteCase.Text = "เวสต์น้ำท่วม กรอ." Then
        myWasteDataCR_.WasteCaseID = "10EE3782390242959F296DA722542573"
        DtpStart.Visible = False
        DtpEnd.Visible = False
         Label62.Visible = False
 ElseIf cboWasteCase.Text = "เวสต์ทิ้ง" Then
        myWasteDataCR_.companyID = "AF543EA5BC1649259AA7955CE4D3C7CD"
        DtpStart.Visible = False
        DtpEnd.Visible = False
         Label62.Visible = False
  ElseIf cboWasteCase.Text = "Haz Cleaning" Then
        myWasteDataCR_.companyID = "SHazcleening101"
        DtpStart.Visible = False
        DtpEnd.Visible = False
         Label62.Visible = False
   ElseIf cboWasteCase.Text = "Non ราคาถูก" Then
        myWasteDataCR_.companyID = "SNonLow101"
        DtpStart.Visible = False
        DtpEnd.Visible = False
         Label62.Visible = False
   ElseIf cboWasteCase.Text = "ปกติ" Then
    myWasteDataCR_.companyID = ""
        DtpStart.Visible = False
        DtpEnd.Visible = False
         Label62.Visible = False
    ElseIf cboWasteCase.Text = "งาน Job" Then
        myWasteDataCR_.WasteCaseID = "1"
        DtpStart.Visible = True
        DtpEnd.Visible = True
         Label62.Visible = True
End If
    myWasteDataCR_.isChanged = True
End Sub

Private Sub cboWasteChar_Click()
On Error Resume Next
Dim CountRow, i As Integer
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
        Case 4
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
'            If Val(txtHeatingValue.Text) >= 500 And Val(txtHeatingValue.Text) < 2800 Then
'                With lvNewTreatment
'                    .ListItems.Clear
'                    PopulateTreatment
'                    .ListItems.item(4).Checked = True
'                    CountRow = myWasteDataLab_.NewTreatments.Count
'                        For i = 1 To CountRow
'                            myWasteDataLab_.NewTreatments.Remove i
'                        Next
'                    myWasteDataLab_.NewTreatments.Add colTreatment_(.ListItems.item(4).Key), .ListItems.item(4).Key
'                End With
'            End If
        Case 4
            myWasteDataLab_.WasteCharacter = "G" 'Gas
    End Select
    myWasteDataLab_.isChanged = True
End Sub

Private Sub cboWasteChar_Validate(Cancel As Boolean)
Dim CountRow, i As Integer

    Select Case cboWasteChar.ListIndex
        Case 0
            myWasteDataLab_.WasteCharacter = "NA"
        Case 1
            myWasteDataLab_.WasteCharacter = "SO" 'Solid
        Case 2
            myWasteDataLab_.WasteCharacter = "SL" 'Sludge
        Case 3
            myWasteDataLab_.WasteCharacter = "LI" 'Liquid
            If Val(txtHeatingValue.Text) >= 500 And Val(txtHeatingValue.Text) < 2800 Then
                With lvNewTreatment
                    .ListItems.Clear
                    PopulateTreatment
                    .ListItems.item(4).Checked = True
                    CountRow = myWasteDataLab_.NewTreatments.Count
                        For i = 1 To CountRow
                            myWasteDataLab_.NewTreatments.Remove i
                        Next
                    myWasteDataLab_.NewTreatments.Add colTreatment_(.ListItems.item(4).Key), .ListItems.item(4).Key
                End With
            End If
        Case 4
            myWasteDataLab_.WasteCharacter = "G" 'Gas
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
        'myWasteDataCR_.WasteType = "N/A"
    Else
        'myWasteDataCR_.WasteType = cboWasteType.Text
        'myWasteDataCR_.companyID = CurrentUser.WasteCompany(cboCompany.ListIndex).ID
    End If
    If cboWasteType.ListIndex = 4 Then
        'cboCompany.ListIndex = 3
        'myWasteDataCR_.companyID = CurrentUser.WasteCompany(cboCompany.ListIndex).ID
    End If
    If cboWasteType.ListIndex = 8 Then
        'cboCompany.ListIndex = 3
        'myWasteDataCR_.companyID = CurrentUser.WasteCompany(cboCompany.ListIndex).ID
    End If
    If cboWasteType.ListIndex >= 1 And cboWasteType.ListIndex <= 3 Then
        'cboCompany.ListIndex = 1
        'myWasteDataCR_.companyID = CurrentUser.WasteCompany(cboCompany.ListIndex).ID
    End If
    If cboWasteType.ListIndex = 6 Then
        'cboCompany.ListIndex = 1
        'myWasteDataCR_.companyID = CurrentUser.WasteCompany(cboCompany.ListIndex).ID
    End If
    If cboWasteType.ListIndex = 7 Then
        'cboCompany.ListIndex = 1
        'myWasteDataCR_.companyID = CurrentUser.WasteCompany(cboCompany.ListIndex).ID
    End If
    If cboWasteType.ListIndex = 5 Then
        'cboCompany.ListIndex = 2
        'myWasteDataCR_.companyID = CurrentUser.WasteCompany(cboCompany.ListIndex).ID
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

Private Sub CboWasteYear_Validate(Cancel As Boolean)
myWasteDataCR_.isChanged = True
End Sub
'Private Sub CboWasteYear_Validate(Cancel As Boolean)
'myWasteDataCR_.isChanged = True
'    isNew = True
'    PopulateCompany
''    datetime
'    Set curEUCode1_ = New BWGEUCode1
'    Set curEUCode2_ = New BWGEUCode2
'    Set curEUCode3_ = New BWGEUCode3
'    Set curMktStaff_ = New BWGCompanyStaff
'    Set curEnviStaff_ = New BWGCompanyStaff
'    PopulateTreatment
'    CreateNewWasteData
'    SetUserGroup
'    ReDim Attaches(4)
'    SSTab1.Tab = 0
'    SSTab1.TabVisible(4) = False
'End Sub

Private Sub CbSamplingBy_Validate(Cancel As Boolean)
 On Error Resume Next
        myWasteDataLab_.SamplingBy = CbSamplingBy.Text
  Select Case CbSamplingBy.ListIndex
    Case 0
        myWasteDataLab_.SamplingByNo = "ว-217-จ-8953"
    Case 1
        myWasteDataLab_.SamplingByNo = "ว-217-จ-8954"
    Case 2
        myWasteDataLab_.SamplingByNo = "ว-217-จ-8955"
    Case 3
        myWasteDataLab_.SamplingByNo = "ว-217-จ-5096"
    Case 4
        myWasteDataLab_.SamplingByNo = "ว-217-จ-6016"
    Case 5
        myWasteDataLab_.SamplingByNo = "ว-217-จ-6017"
    Case 6
        myWasteDataLab_.SamplingByNo = "ว-217-จ-6019"
    Case 7
        myWasteDataLab_.SamplingByNo = "ว-217-จ-7627"
    End Select

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
    myWasteDataLab_.StorageType = "BT"
'    myWasteDataLab_.StorageDesc = Trim(txtBottle.Text)
    myWasteDataLab_.isBottle = chkBottle.Value = 1
End Sub
Private Sub chkBottle_Validate(Cancel As Boolean)
   If chkBottle.Value = 1 Then
        myWasteDataLab_.StorageType = "BT"
   End If
    myWasteDataLab_.isBottle = chkBottle.Value = 1
End Sub
Private Sub chkCancer_Validate(Cancel As Boolean)
    myWasteDataLab_.isCarcinoGenic = chkCancer.Value = 1
End Sub
Private Sub ChkCanEvap_Click()
Dim CountRow, i As Integer
If ChkCanEvap.Value = 1 Then
    myWasteDataCR_.CanEvap = ChkCanEvap.Caption
    myWasteDataLab_.CanEvap = ChkCanEvap.Caption
    Else
     myWasteDataCR_.CanEvap = ""
    myWasteDataLab_.CanEvap = ""
End If
End Sub
Private Sub ChkCanEvap_Validate(Cancel As Boolean)
Dim CountRow As Integer
Dim i As Integer
If ChkCanEvap.Value = 1 Then
    myWasteDataCR_.CanEvap = ChkCanEvap.Caption
    myWasteDataLab_.CanEvap = ChkCanEvap.Caption
    Else
     myWasteDataCR_.CanEvap = ""
    myWasteDataLab_.CanEvap = ""
End If
End Sub
Private Sub ChkCl_Click()
    myWasteDataLab_.isChanged = True
End Sub
Private Sub ChkCn_Click()
    myWasteDataLab_.isChanged = True
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



Private Sub chkInflammable_Validate(Cancel As Boolean)
   myWasteDataLab_.InFlammable = chkInflammable.Value = 1
End Sub

Private Sub ChkisFactoryCar_Validate(Cancel As Boolean)
    myWasteDataCR_.isFactoryCar = ChkisFactoryCar.Value = 1
End Sub
Private Sub ChkisHuerta_Validate(Cancel As Boolean)
    myWasteDataCR_.isHuerta = ChkisHuerta.Value = 1
End Sub
Private Sub ChkisLugger_Validate(Cancel As Boolean)
    myWasteDataCR_.isLugger = ChkisLugger.Value = 1
End Sub
Private Sub ChkisOrtherCar_Validate(Cancel As Boolean)
    myWasteDataCR_.isOrtherCar = ChkisOrtherCar.Value = 1
End Sub
Private Sub Chkispickup_Validate(Cancel As Boolean)
    myWasteDataCR_.isPickup = Chkispickup.Value = 1
End Sub
Private Sub ChkisRollOff_Validate(Cancel As Boolean)
    myWasteDataCR_.isRollOff = ChkisRollOff.Value = 1
End Sub
Private Sub ChkisTankcar_Validate(Cancel As Boolean)
    myWasteDataCR_.isTankcar = ChkisTankcar.Value = 1
End Sub
Private Sub ChkisTrailerW_Validate(Cancel As Boolean)
    myWasteDataCR_.isTrailerW = ChkisTrailerW.Value = 1
End Sub
Private Sub ChkisTruckTen_Validate(Cancel As Boolean)
    myWasteDataCR_.isTruckTen = ChkisTruckTen.Value = 1
End Sub
Private Sub ChkLicense_Validate(Cancel As Boolean)
    If ChkLicense.Value = 1 Then
        myWasteDataCR_.LicenseStartDate = ""
        myWasteDataCR_.LicenseExpDate = ""
        txtStartDate.Text = ""
    End If
End Sub

Private Sub ChkNt_Click()
    myWasteDataLab_.isChanged = True
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
    myWasteDataLab_.StorageType = "PB"
End Sub

Private Sub chkPaperBox_Validate(Cancel As Boolean)
    myWasteDataLab_.isPaperBox = chkPaperBox.Value = 1
    myWasteDataLab_.StorageType = "PB"
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
    myWasteDataLab_.StorageType = "PT"
    myWasteDataLab_.isPlasticBag = chkPlasticBag.Value = 1
End Sub

Private Sub chkPlasticBag_Validate(Cancel As Boolean)
    If chkPlasticBag.Value = 0 Then txtPlasticBag.Text = ""
    myWasteDataLab_.isPlasticBag = chkPlasticBag.Value = 1
    myWasteDataLab_.StorageType = "PT"
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
    myWasteDataLab_.StorageType = "PD"
End Sub

Private Sub chkPlasticDrum_Validate(Cancel As Boolean)
    myWasteDataLab_.isPlasticDrum = chkPlasticDrum.Value = 1
    myWasteDataLab_.StorageType = "PD"
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
        myWasteDataLab_.StorageType = "PG"
    Else
    myWasteDataLab_.isPlasticGallon = ""
    myWasteDataLab_.PlasticGallon = ""
    End If
    myWasteDataLab_.isPlasticGallon = chkPlasticGallon.Value = 1
End Sub

Private Sub chkPlasticGallon_Validate(Cancel As Boolean)
    If chkPlasticGallon.Value = 1 Then
    myWasteDataLab_.StorageType = "PG"
    End If
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
            '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
              If Trim(txtCustomer.Text) = "" Then
                MsgBox "กรุณาระบุลูกค้าก่อนทำการจัดเก็บข้อมูล", vbExclamation
                txtCustomer.SetFocus
                isSave = False
                cmdApprove.Visible = True
                cmdSave.Enabled = False
                Exit Sub
            End If
            If Trim(txtMktStaff.Text) = "" Then
                MsgBox "กรุณาระบุชื่อเจ้าหน้าที่การตลาด", vbExclamation
                txtMktStaff.SetFocus
                isSave = False
                cmdApprove.Visible = True
                cmdSave.Enabled = False
                Exit Sub
            End If
'            If cboCompany.ListIndex = 0 Then
'                MsgBox "กรุณาเลือกบริษัทที่รับกำจัด", vbExclamation, "Warning"
'                isSave = False
'                cmdApprove.Visible = True
'                cmdSave.Enabled = False
'                Exit Sub
'            End If
            If Trim(txtWasteName.Text) = "" Then
                MsgBox "กรุณาระบุชื่อกากของเสีย", vbExclamation
                txtWasteName.SetFocus
                isSave = False
                cmdApprove.Visible = True
                cmdSave.Enabled = False
                Exit Sub
            End If
        
'            If myWasteDataCR_.isQuotationReq = False Then
'                If Trim(txtFactoryNo.Text) = "" Then
'                    MsgBox "กรุณาระบุเลขทะเบียนโรงงาน", vbExclamation
'                    txtFactoryNo.SetFocus
'                    isSave = False
'                    cmdApprove.Visible = True
'                    cmdSave.Enabled = False
'                    Exit Sub
'                End If
'            End If
            If curCustomer_.CustomerType = "L" Or curCustomer_.CustomerType = "P" Then
            Else
'            If Trim(txtWasteGenNo.Text) = "" Then
'                MsgBox "กรุณาระบุเลขที่ผู้กำเนิดของเสีย", vbExclamation
'                txtWasteGenNo.SetFocus
'                isSave = False
'                cmdApprove.Visible = True
'                cmdSave.Enabled = False
'                Exit Sub
'            End If
'            If Trim(txtIndustryType.Text) = "" Then
'                MsgBox "กรุณาระบุประเภทอุตสาหกรรม", vbExclamation
'                txtIndustryType.SetFocus
'                isSave = False
'                cmdApprove.Visible = True
'                cmdSave.Enabled = False
'                Exit Sub
'            End If
            If Trim(txtQty.Text) = "" Or Trim(txtQty.Text) = "0" Then
                MsgBox "กรุณาระบุปริมาณที่จัดเก็บปัจจบัน และมากกว่า 0 ", vbExclamation
                txtQty.SetFocus
                isSave = False
                cmdApprove.Visible = True
                cmdSave.Enabled = False
                Exit Sub
            End If
            If Trim(txtRate.Text) = "" Or Trim(txtRate.Text) = "0" Then
                MsgBox "กรุณาระบุอัตราการเกิดของเสีย และมากกว่า 0", vbExclamation
                txtRate.SetFocus
                isSave = False
                cmdApprove.Visible = True
                cmdSave.Enabled = False
                Exit Sub
            End If
            If Trim(txtGenerateDesc.Text) = "" Then
                MsgBox "กรุณาระบุกระบวนการเกิดของเสีย", vbExclamation
                txtGenerateDesc.SetFocus
                isSave = False
                cmdApprove.Visible = True
                cmdSave.Enabled = False
                Exit Sub
            End If
            If cboWasteChar.ListIndex = 0 Then
                MsgBox "กรุณาระบุคุณสมบัติ", vbExclamation
                cboWasteChar.SetFocus
                isSave = False
                cmdApprove.Visible = True
                cmdSave.Enabled = False
                Exit Sub
            End If
            If cboWasteChar.ListIndex >= 1 And cboWasteChar.ListIndex <= 2 Then
                If cboWasteCharDesc.ListIndex < 0 Then
                    MsgBox "กรุณาระบุรายละเอียดคุณสมบัติ", vbExclamation
                    cboWasteCharDesc.SetFocus
                    isSave = False
                    cmdApprove.Visible = True
                    cmdSave.Enabled = False
                    Exit Sub
                End If
              End If
                If cboStoreType.ListIndex = 0 Then
                    MsgBox "กรุณาระบุวิธีการจัดเก็บ", vbExclamation
                    cboStoreType.SetFocus
                    isSave = False
                    cmdApprove.Visible = True
                    cmdSave.Enabled = False
                    Exit Sub
                End If
                If cboIndoor.ListIndex = 0 Then
                    MsgBox "กรุณาระบุสถานที่ในการจัดเก็บ", vbExclamation
                    cboIndoor.SetFocus
                    isSave = False
                    cmdApprove.Visible = True
                    cmdSave.Enabled = False
                    Exit Sub
                End If
                If Chk10Detail = False Then
                    MsgBox "กรุณาระบุข้อ 10.3", vbExclamation
                    isSave = False
                    cmdApprove.Visible = True
                    cmdSave.Enabled = False
                    Exit Sub
                End If
             If chkBag.Value = 1 Then
                If Trim(txtBag.Text) = "" Or Trim(txtBag.Text) = 0 Then
                    MsgBox "ระบุขนาดจัดเก็บ ของถุง  Big Bag และขนาด มากกว่า 0", vbExclamation
                    cmdApprove.Visible = True
                    cmdSave.Enabled = False
                    Exit Sub
                End If
            End If
            If chkDrum.Value = 1 Then
                If Trim(txtDrumDesc.Text) = "" Or Trim(txtDrumDesc.Text) = 0 Then
                    MsgBox "ระบุขนาดจัดเก็บ Metal Drum และขนาด มากกว่า 0 ", vbExclamation
                    cmdApprove.Visible = True
                    cmdSave.Enabled = False
                    Exit Sub
                End If
            End If
            If chkPlasticGallon.Value = 1 Then
                If Trim(txtPlasticGallon.Text) = "" And Trim(txtPlasticGallon.Text) = 0 Then
                    MsgBox "ระบุขนาดจัดเก็บ แกนลอนพลาสติก : Plastic Gallon  และขนาด มากกว่า 0 ", vbExclamation
                    cmdApprove.Visible = True
                    cmdSave.Enabled = False
                    Exit Sub
                End If
            End If
              If chkBottle.Value = 1 Then
                If Trim(txtBottle.Text) = "" Or Trim(txtBottle.Text) = 0 Then
                    MsgBox "ระบุขนาดจัดเก็บ ใน ขวดแก้ว : Bottle และระบุขนาดจัดเก็บมากกว่า 0", vbExclamation
                    cmdApprove.Visible = True
                    cmdSave.Enabled = False
                    Exit Sub
                End If
            End If
            If chkPlasticBag.Value = 1 Then
                If Trim(txtPlasticBag.Text) = "" Or Trim(txtPlasticBag.Text) = 0 Then
                    MsgBox "ระบุขนาดจัดเก็บ ถุงพลาสติก : Plastic Bag และระบุขนาดจัดเก็บมากกว่า 0 ", vbExclamation
                    cmdApprove.Visible = True
                    cmdSave.Enabled = False
                    Exit Sub
                End If
            End If
            If chkPlasticDrum.Value = 1 Then
                If Trim(txtPlasticDrum.Text) = "" Or Trim(txtPlasticDrum.Text) = 0 Then
                    MsgBox "ระบุขนาดจัดเก็บ ถังพลาสติก : Plastic Drum และระบุขนาดจัดเก็บมากกว่า 0", vbExclamation
                    cmdApprove.Visible = True
                    cmdSave.Enabled = False
                    Exit Sub
                End If
            End If
            If chkPaperBox.Value = 1 Then
                If Trim(txtPaperBox.Text) = "" Or Trim(txtPaperBox.Text) = 0 Then
                    MsgBox "ระบุขนาดจัดเก็บ ของ กล่องกระดาษ และระบุขนาดจัดเก็บมากกว่า 0", vbExclamation
                     cmdApprove.Visible = True
                    cmdSave.Enabled = False
                    Exit Sub
                End If
            End If
            End If
'            If chkOther2.Value = 1 Then
'                If Trim(txtOther2.Text) = "" Then
'                    MsgBox "กรุณา ระบุรายละเอียด อื่น ๆ"
'                    cmdApprove.Visible = True
'                    cmdSave.Enabled = False
'                    Exit Sub
'                End If
'            End If
            If Trim(txtCustomer.Text) = "" Then
                    MsgBox "กรุณาระบุลูกค้าก่อนทำการจัดเก็บข้อมูล", vbExclamation
                    txtCustomer.SetFocus
                    isSave = False
                    cmdApprove.Visible = True
                    cmdSave.Enabled = False
                    Exit Sub
            End If
            If Trim(txtMktStaff.Text) = "" Then
                    MsgBox "กรุณาระบุชื่อเจ้าหน้าที่การตลาด", vbExclamation
                    txtMktStaff.SetFocus
                    isSave = False
                    cmdApprove.Visible = True
                    cmdSave.Enabled = False
                    Exit Sub
            End If
'                If cboCompany.ListIndex = 0 Then
'                    MsgBox "กรุณาเลือกบริษัทที่รับกำจัด", vbExclamation, "Warning"
'                    isSave = False
'                    cmdApprove.Visible = True
'                    cmdSave.Enabled = False
'                    Exit Sub
'                End If
                If Trim(txtWasteName.Text) = "" Then
                    MsgBox "กรุณาระบุชื่อกากของเสีย", vbExclamation
                    txtWasteName.SetFocus
                    isSave = False
                    cmdApprove.Visible = True
                    cmdSave.Enabled = False
                    Exit Sub
                End If
                If isEnv = True Then
                     If OptAnalysType3.Value = True And (txtAnalysDesc.Text = "" Or txtAnalysDesc.Text = "-") Then
                            MsgBox "กรุณาระบุ Analysis Type"
                            txtAnalysDesc.Locked = False
                            cmdApprove.Visible = True
                            cmdSave.Enabled = False
                            Exit Sub
                    End If
'                    If txtEUCode3.Text <> "99" Then
'                        If txtHAType.Text = "HA" Or txtHAType.Text = "HM" Then
'                           If cboWasteType.Text = "Haz" Or cboWasteType.Text = "HBL" Or cboWasteType.Text = "HIC" Or cboWasteType.Text = "NIC" Or cboWasteType.Text = "BBH" Or cboWasteType.Text = "BBL" Then
'                           Else
'                            MsgBox "กรุณาระบุประเภทของเสีย ให้ถูกต้อง ตามเลข 6 หลัก", vbExclamation
'                           Exit Sub
'                           End If
'                        End If
'                        If txtHAType.Text = "" Then
'                            If cboWasteType.Text = "Non" Or cboWasteType.Text = "NBL" Or cboWasteType.Text = "HIC" Or cboWasteType.Text = "NIC" Or cboWasteType.Text = "BBH" Or cboWasteType.Text = "BBL" Then
'                           Else
'                            MsgBox "กรุณาระบุประเภทของเสีย ให้ถูกต้อง ตามเลข 6 หลัก", vbExclamation
'                           Exit Sub
'                           End If
'                        End If
'                    End If
                End If
            '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
            isNew = False
            myWasteDataCR_.isNew = False
            myWasteDataLab_.isNew = False
            isSave = True
            CmdSave_Click
'            myWasteDataCR_.Save
'            If isSave = False Then Exit Sub
            myWasteDataCR_.EnviAccepted
'            Command3_Click
            Label1.Caption = "Envi ตรวจรับแล้วเมื่อ " & FormatYMD_to_DMY(myWasteDataCR_.acceptedDate, "/", "/")
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
        If myWasteDataCR_.isAcceptedByEnvi = True And isEnv = True And (myWasteDataCR_.Status = "A" Or myWasteDataCR_.Status = "P" Or myWasteDataCR_.Status = "IA") Then
            MsgBox myWasteDataCR_.WasteName & " ตรวจรับแล้ว ไม่สามารถส่งให้ตรวจรับได้อีก", vbExclamation, "Warning"
            Exit Sub
        End If
            If Trim(txtCustomer.Text) = "" Then
                MsgBox "กรุณาระบุลูกค้าก่อนทำการจัดเก็บข้อมูล", vbExclamation
                txtCustomer.SetFocus
                isSave = False
                cmdApprove.Visible = False
                cmdSave.Enabled = True
                Exit Sub
            End If
            If Trim(txtMktStaff.Text) = "" Then
                MsgBox "กรุณาระบุชื่อเจ้าหน้าที่การตลาด", vbExclamation
                txtMktStaff.SetFocus
                isSave = False
                cmdApprove.Visible = False
                cmdSave.Enabled = True
                Exit Sub
            End If
'            If cboCompany.ListIndex = 0 Then
'                MsgBox "กรุณาเลือกบริษัทที่รับกำจัด", vbExclamation, "Warning"
'                isSave = False
'                cmdApprove.Visible = False
'                cmdSave.Enabled = True
'                Exit Sub
'            End If
            If Trim(txtWasteName.Text) = "" Then
                MsgBox "กรุณาระบุชื่อกากของเสีย", vbExclamation
                txtWasteName.SetFocus
                isSave = False
                cmdApprove.Visible = False
                cmdSave.Enabled = True
                Exit Sub
            End If
'            If myWasteDataCR_.isQuotationReq = False Then
'                If Trim(txtFactoryNo.Text) = "" Then
'                    MsgBox "กรุณาระบุเลขทะเบียนโรงงาน", vbExclamation
'                    txtFactoryNo.SetFocus
'                    isSave = False
'                    cmdApprove.Visible = False
'                    cmdSave.Enabled = True
'                    Exit Sub
'                End If
'            End If
            If curCustomer_.CustomerType = "L" Or curCustomer_.CustomerType = "P" Then
            Else
            
'            If Trim(txtWasteGenNo.Text) = "" Then
'                MsgBox "กรุณาระบุเลขที่ผู้กำเนิดของเสีย", vbExclamation
'                txtWasteGenNo.SetFocus
'                isSave = False
'                cmdApprove.Visible = False
'                cmdSave.Enabled = True
'                Exit Sub
'            End If
'            If Trim(txtIndustryType.Text) = "" Then
'                MsgBox "กรุณาระบุประเภทอุตสาหกรรม", vbExclamation
'                txtIndustryType.SetFocus
'                isSave = False
'                cmdApprove.Visible = False
'                cmdSave.Enabled = True
'                Exit Sub
'            End If
            If Trim(txtQty.Text) = "" Or Trim(txtQty.Text) = "0" Or txtQty = Null Or Trim(txtQty.Text) = "0.00" Then
                MsgBox "กรุณาระบุปริมาณที่จัดเก็บปัจจบัน และมากกว่า 0 ", vbExclamation
                txtQty.SetFocus
                isSave = False
                cmdApprove.Visible = False
                cmdSave.Enabled = True
                Exit Sub
            End If
            If Trim(txtRate.Text) = "" Or Trim(txtRate.Text) = "0" Then
                MsgBox "กรุณาระบุอัตราการเกิดของเสีย และมากกว่า 0", vbExclamation
                txtRate.SetFocus
                isSave = False
                cmdApprove.Visible = False
                cmdSave.Enabled = True
                Exit Sub
            End If
            If Trim(txtGenerateDesc.Text) = "" Then
                MsgBox "กรุณาระบุกระบวนการเกิดของเสีย", vbExclamation
                txtGenerateDesc.SetFocus
                isSave = False
                cmdApprove.Visible = False
                cmdSave.Enabled = True
                Exit Sub
            End If
            If cboWasteChar.ListIndex = 0 Then
                MsgBox "กรุณาระบุคุณสมบัติ", vbExclamation
                cboWasteChar.SetFocus
                isSave = False
                cmdApprove.Visible = False
                cmdSave.Enabled = True
                Exit Sub
            End If
            If cboWasteChar.ListIndex >= 1 And cboWasteChar.ListIndex <= 2 Then
                If cboWasteCharDesc.ListIndex < 0 Then
                    MsgBox "กรุณาระบุรายละเอียดคุณสมบัติ", vbExclamation
                    cboWasteCharDesc.SetFocus
                    isSave = False
                    cmdApprove.Visible = False
                    cmdSave.Enabled = True
                    Exit Sub
                End If
              End If
                If cboStoreType.ListIndex = 0 Then
                    MsgBox "กรุณาระบุวิธีการจัดเก็บ", vbExclamation
                    cboStoreType.SetFocus
                    isSave = False
                    cmdApprove.Visible = False
                    cmdSave.Enabled = True
                    Exit Sub
                End If
                If cboIndoor.ListIndex = 0 Then
                    MsgBox "กรุณาระบุสถานที่ในการจัดเก็บ", vbExclamation
                    cboIndoor.SetFocus
                    isSave = False
                    cmdApprove.Visible = False
                    cmdSave.Enabled = True
                    Exit Sub
                End If
                If Chk10Detail = False Then
                    MsgBox "กรุณาระบุข้อ 10.3", vbExclamation
                    isSave = False
                    cmdApprove.Visible = False
                    cmdSave.Enabled = True
                    Exit Sub
                End If
             If chkBag.Value = 1 Then
                If Trim(txtBag.Text) = "" Or Trim(txtBag.Text) = 0 Then
                    MsgBox "ระบุขนาดจัดเก็บ ของถุง  Big Bag และขนาด มากกว่า 0", vbExclamation
                    cmdApprove.Visible = False
                    cmdSave.Enabled = True
                    Exit Sub
                End If
            End If
            If chkDrum.Value = 1 Then
                If Trim(txtDrumDesc.Text) = "" Or Trim(txtDrumDesc.Text) = 0 Then
                    MsgBox "ระบุขนาดจัดเก็บ Metal Drum และขนาด มากกว่า 0 ", vbExclamation
                    cmdApprove.Visible = False
                    cmdSave.Enabled = True
                    Exit Sub
                End If
            End If
            If chkPlasticGallon.Value = 1 Then
                If Trim(txtPlasticGallon.Text) = "" And Trim(txtPlasticGallon.Text) = 0 Then
                    MsgBox "ระบุขนาดจัดเก็บ แกนลอนพลาสติก : Plastic Gallon  และขนาด มากกว่า 0 ", vbExclamation
                    cmdApprove.Visible = False
                    cmdSave.Enabled = True
                    Exit Sub
                End If
            End If
              If chkBottle.Value = 1 Then
                If Trim(txtBottle.Text) = "" Or Trim(txtBottle.Text) = 0 Then
                    MsgBox "ระบุขนาดจัดเก็บ ใน ขวดแก้ว : Bottle และระบุขนาดจัดเก็บมากกว่า 0", vbExclamation
                    cmdApprove.Visible = False
                    cmdSave.Enabled = True
                    Exit Sub
                End If
            End If
            If chkPlasticBag.Value = 1 Then
                If Trim(txtPlasticBag.Text) = "" Or Trim(txtPlasticBag.Text) = 0 Then
                    MsgBox "ระบุขนาดจัดเก็บ ถุงพลาสติก : Plastic Bag และระบุขนาดจัดเก็บมากกว่า 0 ", vbExclamation
                    cmdApprove.Visible = False
                    cmdSave.Enabled = True
                    Exit Sub
                End If
            End If
            If chkPlasticDrum.Value = 1 Then
                If Trim(txtPlasticDrum.Text) = "" Or Trim(txtPlasticDrum.Text) = 0 Then
                    MsgBox "ระบุขนาดจัดเก็บ ถังพลาสติก : Plastic Drum และระบุขนาดจัดเก็บมากกว่า 0", vbExclamation
                    cmdApprove.Visible = False
                    cmdSave.Enabled = True
                    Exit Sub
                End If
            End If
            If chkPaperBox.Value = 1 Then
                If Trim(txtPaperBox.Text) = "" Or Trim(txtPaperBox.Text) = 0 Then
                    MsgBox "ระบุขนาดจัดเก็บ ของ กล่องกระดาษ และระบุขนาดจัดเก็บมากกว่า 0", vbExclamation
                    cmdApprove.Visible = False
                    cmdSave.Enabled = True
                    Exit Sub
                End If
            End If
            End If
'            If chkOther2.Value = 1 Then
'                If Trim(txtOther2.Text) = "" Then
'                    MsgBox "กรุณา ระบุรายละเอียด อื่น ๆ"
'                    cmdApprove.Visible = False
'                    cmdSave.Enabled = True
'                    Exit Sub
'                End If
'            End If
            If Trim(txtCustomer.Text) = "" Then
                    MsgBox "กรุณาระบุลูกค้าก่อนทำการจัดเก็บข้อมูล", vbExclamation
                    txtCustomer.SetFocus
                    isSave = False
                    cmdApprove.Visible = False
                    cmdSave.Enabled = True
                    Exit Sub
            End If
            If Trim(txtMktStaff.Text) = "" Then
                    MsgBox "กรุณาระบุชื่อเจ้าหน้าที่การตลาด", vbExclamation
                    txtMktStaff.SetFocus
                    isSave = False
                    cmdApprove.Visible = False
                    cmdSave.Enabled = True
                    Exit Sub
            End If
'                If cboCompany.ListIndex = 0 Then
'                    MsgBox "กรุณาเลือกบริษัทที่รับกำจัด", vbExclamation, "Warning"
'                    isSave = False
'                    cmdApprove.Visible = False
'                    cmdSave.Enabled = True
'                    Exit Sub
'                End If
                If Trim(txtWasteName.Text) = "" Then
                    MsgBox "กรุณาระบุชื่อกากของเสีย", vbExclamation
                    txtWasteName.SetFocus
                    isSave = False
                    cmdApprove.Visible = False
                    cmdSave.Enabled = True
                    Exit Sub
                End If
        If MsgBox("ยืนยันการส่งข้อมูลกากฯ", vbQuestion + vbYesNo) = vbYes Then
'            If cmdApprove.Caption = "ส่งให้สิ่งแวดล้อม" And cboCompany.ListIndex > 1 Then
'                myWasteDataCR_.Status = "IA"
'                myWasteDataCR_.isAcceptedByEnvi = True
'                myWasteDataCR_.EnviAccepted
'                myWasteDataCR_.isAcceptedByEnvi = True
'                isNew = False
'                CmdSave_Click
'                Label1.Caption = "Envi ตรวจรับแล้วเมื่อ " & FormatYMD_to_DMY(myWasteDataCR_.acceptedDate, "/", "/")
'                Label1.Visible = True
'                Label3.Caption = "ตรวจรับแล้ว"
'                Label3.Visible = True
'                cmdReject.Visible = False
'                cmdApprove.Visible = False
'                cmdSave.Enabled = True
'                txtWasteNo.Text = myWasteDataCR_.WasteNo
'                Exit Sub
'            End If
                myWasteDataCR_.Status = "R"
                CmdSave_Click
                If isSave = False Then Exit Sub
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
    
    If SaveAttache(myWasteDataCR_.CustomerID, myWasteDataCR_.ID) = False Then
        MsgBox "การบันทึกไฟล์แนบไม่สมบูรณ์", vbInformation, "ไฟล์แนบ"
    End If

End Sub





'Private Sub CmdAttPic_Click(Index As Integer)
'Dim FileName As String
'Dim FullPath As String
'    If lblAttPic(Index).Caption <> "" Then KillPic Index, myWasteDataCR_.CustomerID, myWasteDataCR_.ID
'    AttachePic Index, FileName, FullPath
'    lblAttPic(Index).Caption = FileName
'    AttachesPic(Index).FileName = FileName
'    AttachesPic(Index).FullPath = FullPath
'
'    If SaveAttPict(myWasteDataCR_.CustomerID, myWasteDataCR_.ID) = False Then
'        MsgBox "การบันทึกไฟล์รูปตกตะกอนไม่สมบูรณ์", vbInformation, "รูปตกตะกอน"
'    End If
'
'    AttachesPic(Index).FileName = ""
'    AttachesPic(Index).FullPath = ""
'
'End Sub

Private Sub CmdClose_Click()
'    Dim FSO As New FileSystemObject
'    If FSO.FolderExists("Z:") = True Then
'         RemoveDrive "Z:"
'    End If
    Unload Me
End Sub

Private Sub cmdCls_Click(Index As Integer)
    KillFile Index, myWasteDataCR_.CustomerID, myWasteDataCR_.ID
End Sub

Private Sub CmdDateAnalysis_Click()
Load frmSamplingDate
frmSamplingDate.Show 1
End Sub



''Private Sub CmdDelPic_Click(Index As Integer)
''    KillPic Index, myWasteDataCR_.CustomerID, myWasteDataCR_.ID
''End Sub





'Private Sub cmdOpen_Click(Index As Integer)
'Dim OpenFile As Double
'Dim oFS As Scripting.FileSystemObject
'Dim FSO As New FileSystemObject
'Dim oNet As Object, MapDrv As Boolean
''    If FSO.FolderExists("Z:") = False Then
''DefaultPath = "S:\AttacheFiles"
'        PicPath = "\\192.168.1.5\PicAtt\"
''        MapDrv = MapDrive(DefaultPath)
''    End If
'    If lblAttPic(Index).Caption <> "" Then
'       ShellExecute 0, vbNullString, AttachesPic(Index).FullPath, vbNullString, vbNullString, vbNormalFocus
'    End If
'
'End Sub

Private Sub cmdReject_Click()
Dim RejectDetail As String
If myWasteDataCR_.Status = Trim("AR") Then
     If MsgBox("ยืนยันการตีกลับ", vbInformation + vbYesNo, "ตีกลับ") = vbNo Then Exit Sub
     RejectDetail = InputBox("ใส่รายละเอียดในการตีกลับ", "ใส่รายละเอียด", "")
     myWasteDataCR_.Reject RejectDetail
     myWasteDataCR_.Status = "RJ"
     If cboWasteCase.Text = "งาน Job" Then
        myWasteDataCR_.SDate = DateValue(DtpStart.Value)
        myWasteDataCR_.EDate = DateValue(DtpEnd.Value)
     End If
     myWasteDataCR_.Save
    Label3.Caption = "ตีกลับ"
    Label3.Visible = True
Else
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
    End If
End Sub

Private Sub cmdRev_Click()
If isEnv = True Then
'        myWasteDataCR_.acceptedDate = TodayDate
If cboWasteCase.Text = "งาน Job" Then
    myWasteDataCR_.SDate = DateValue(DtpStart.Value)
    myWasteDataCR_.EDate = DateValue(DtpEnd.Value)
End If
        myWasteDataCR_.Save
        myWasteDataLab_.Save
            If SaveAttache(myWasteDataCR_.CustomerID, myWasteDataCR_.ID) = False Then
                MsgBox "การบันทึกไฟล์แนบไม่สมบูรณ์", vbInformation, "ไฟล์แนบ"
            End If
            MsgBox " Revised Waste สำเร็จ", vbInformation, "การดำเนินการ"
            cmdSave.Enabled = True
            Command3_Click
    Else
        If MsgBox(" คุณต้องการจะ Rev. Waste รหัส  " & myWasteDataCR_.WasteNo & " ", vbInformation + vbYesNo, "ยืนยัน") = vbNo Then Exit Sub
        If MsgBox(" โปรดตรวจสอบ ชื่อเจ้าหน้าที่การตลาด หากมีการเปลี่ยนแปลง เจ้าหน้าที่การตลาดให้ระบุหลังจาก ทำการ Revised แล้ว  คุณต้องการจะ Rev. Waste รหัส  " & myWasteDataCR_.WasteNo & " ", vbInformation + vbYesNo, "ยืนยันการทำรายการ") = vbNo Then Exit Sub
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
'        If cboCompany.Visible = True Then
'            If cboCompany.ListIndex = 0 Then
'                MsgBox "กรุณาเลือกบริษัทที่รับกำจัด", vbExclamation, "Warning"
'                isSave = False
'                cmdApprove.Visible = False
'            Exit Sub
'        End If
'        End If

        If Trim(txtWasteName.Text) = "" Then
            MsgBox "กรุณาระบุชื่อกากของเสีย", vbExclamation
            txtWasteName.SetFocus
            isSave = False
            cmdApprove.Visible = False
            Exit Sub
        End If
'        If myWasteDataCR_.isQuotationReq = False Then
'            If Trim(txtFactoryNo.Text) = "" Then
'                MsgBox "กรุณาระบุเลขทะเบียนโรงงาน", vbExclamation
'                txtFactoryNo.SetFocus
'                isSave = False
'                cmdApprove.Visible = False
'                Exit Sub
'            End If
'        End If
'        If Trim(txtWasteGenNo.Text) = "" Then
'            MsgBox "กรุณาระบุเลขที่ผู้กำเนิดของเสีย", vbExclamation
'            txtWasteGenNo.SetFocus
'            isSave = False
'            cmdApprove.Visible = False
'            Exit Sub
'        End If
'        If Trim(txtIndustryType.Text) = "" Then
'            MsgBox "กรุณาระบุประเภทอุตสาหกรรม", vbExclamation
'            txtIndustryType.SetFocus
'            isSave = False
'            cmdApprove.Visible = False
'            Exit Sub
'        End If
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
'        myWasteDataCR_.Status = "IA"
        If cboWasteCase.Text = "งาน Job" Then
            myWasteDataCR_.SDate = DateValue(DtpStart.Value)
            myWasteDataCR_.EDate = DateValue(DtpEnd.Value)
        End If
        myWasteDataCR_.Save
        myWasteDataLab_.Save
        myWasteDataCR_.isQuotationReq = True
        myWasteDataCR_.Add
        myWasteDataLab_.Add
        cmdSave.Enabled = False
        MsgBox " บันทึกข้อมูล Waste สำเร็จ", vbInformation, "การดำเนินการ"
    End If
End Sub

Private Sub CmdSave_Click()
    isSave = True
        If cboWasteAnalysis.ListIndex = 0 Then
               ckValus
               If isSave = False Then
                    Exit Sub
               End If
        ElseIf cboWasteAnalysis.ListIndex = 1 And (myWasteDataCR_.Status = Trim("P") Or myWasteDataCR_.Status = Trim("A") Or myWasteDataCR_.Status = Trim("E")) Then
               ckValus
               If isSave = False Then
                    Exit Sub
               End If
        Else
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
                If cboCompany.ListIndex = 0 Or myWasteDataCR_.companyID = "" Or myWasteDataCR_.companyID = "NULL" Then
                    MsgBox "กรุณาเลือกบริษัทก่อน", vbExclamation
                    txtMktStaff.SetFocus
                    isSave = False
                    cmdApprove.Visible = False
                    Exit Sub
                End If
                
    '            If cboCompany.Visible = True Then
    '                If cboCompany.ListIndex = 0 Then
    '                    MsgBox "กรุณาเลือกบริษัทที่รับกำจัด", vbExclamation, "Warning"
    '                    isSave = False
    '                    cmdApprove.Visible = False
    '                    Exit Sub
    '                End If
    '            End If
                If Trim(txtWasteName.Text) = "" Then
                    MsgBox "กรุณาระบุชื่อกากของเสีย", vbExclamation
                    txtWasteName.SetFocus
                    isSave = False
                    cmdApprove.Visible = False
                    Exit Sub
                End If
        End If
        If isEnv = True Then
            If curCustomer_.CustomerType = "L" Or curCustomer_.CustomerType = "P" Or curCustomer_.CustomerType = "R" Then
                myWasteDataCR_.Status = "A"
            Else
                 If OptAnalysType3.Value = True And (txtAnalysDesc.Text = "" Or txtAnalysDesc.Text = "-") Then
                        MsgBox "กรุณาระบุ Analysis Type"
                        txtAnalysDesc.Locked = False
                    Exit Sub
                End If
                If cboSummary.Text = "รับดำเนินการ" Then
                Dim ADT1 As String
                ADT1 = DateValue(Date)
                myWasteDataCR_.WasteApproveStatus = "E" 'Accept
                myWasteDataCR_.Status = "A"
                myWasteDataCR_.isTreatAble = True
                myWasteDataCR_.AcceptStatusDate = ADT1
                
                    If txtEUCode3.Text <> "99" Then
                        If txtHAType.Text = "HA" Or txtHAType.Text = "HM" Then
                           If cboWasteType.Text = "Haz" Or cboWasteType.Text = "HBL" Or cboWasteType.Text = "HIC" Or cboWasteType.Text = "NIC" Or cboWasteType.Text = "BBH" Or cboWasteType.Text = "BBL" Then ' Or cboWasteType.Text = "BWG" Or cboWasteType.Text = "BWC" Or cboWasteType.Text = "E" Or cboWasteType.Text = "BLF" Or cboWasteType.Text = "DOC" Then
                                If IsNull(myWasteDataCR_.AcceptStatusDate) Then
                                    myWasteDataCR_.SaveAccdate
                                End If
''                           Else
''                                MsgBox "กรุณาระบุประเภทของเสีย ให้ถูกต้อง ตามเลข 6 หลัก", vbExclamation
''                                cboWasteType.SetFocus
''                                Exit Sub
                           End If
                        End If
''                        If txtHAType.Text = "" Then
''                            If cboWasteType.Text = "Non" Or cboWasteType.Text = "NBL" Or cboWasteType.Text = "HIC" Or cboWasteType.Text = "NIC" Or cboWasteType.Text = "BBH" Or cboWasteType.Text = "BBL" Then ' Or cboWasteType.Text = "BWG" Or cboWasteType.Text = "BWC" Or cboWasteType.Text = "E" Or cboWasteType.Text = "BLF" Or cboWasteType.Text = "DOC" Then
''                           Else
''                                MsgBox "กรุณาระบุประเภทของเสีย ให้ถูกต้องตาม เลข 6 หลัก", vbExclamation
''                                cboWasteType.SetFocus
''                                Exit Sub
''                           End If
''                        End If
                        End If
                    End If
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
            myWasteDataCR_.acceptedDate = Accepdate
        End If
    
    '    If Trim(myWasteDataCR_.Status) = "P" Or Trim(myWasteDataCR_.Status) = "A" Or Trim(myWasteDataCR_.Status) = "E" Then
    '        myWasteDataCR_.isNew = False
    '        myWasteDataLab_.isNew = False
    '    Else
    '        myWasteDataCR_.isNew = True
    '        myWasteDataLab_.isNew = True
    '    End If
    '    isNew = myWasteDataCR_.isNew
        '**********
        If isNew = True Then
            If myWasteDataCR_.Status = "W" Then
                myWasteDataCR_.Status = "W"
                myWasteDataCR_.isNew = True
            ElseIf myWasteDataCR_.Status = "D" Then
                myWasteDataCR_.Status = "D"  'new waste
                myWasteDataCR_.isNew = True
            ElseIf myWasteDataCR_.Status = "R" Then
                myWasteDataCR_.Status = "R"  'new waste
                myWasteDataCR_.isNew = True
            ElseIf myWasteDataCR_.Status = "J" Then
                myWasteDataCR_.Status = "J"  'new waste
                myWasteDataCR_.isNew = True
            End If
            myWasteDataCR_.isAcceptedByEnvi = False
            myWasteDataCR_.isRequest = False
            myWasteDataCR_.isNew = True
            myWasteDataCR_.Status = "D"
            myWasteDataLab_.isNew = True
        Else
            myWasteDataCR_.isAcceptedByEnvi = True
            myWasteDataCR_.isRequest = True
            myWasteDataLab_.isNew = False
        End If
        '************
        If CboWasteYear.Text = "" Or CboWasteYear.Text = Trim("") Then
            MsgBox "กรุณา ระบุปีของ Waste "
                cmdApprove.Visible = False
                Exit Sub
        End If
        If cboSummary.Text = "รับดำเนินการ" Then
            If cboWasteType.Text = "ยังไม่ระบุ" Then
                MsgBox "กรุณาระบุประเภทกากอุตสาหกรรม", vbInformation, "ประเภท"
                Exit Sub
            End If
        End If
        myWasteDataCR_.WasteYears = CboWasteYear.Text
        myWasteDataCR_.FactoryRegisNo = txtFactoryNo.Text
        myWasteDataCR_.Digit14 = Txt14Digit.Text
        myWasteDataCR_.MarketingStaffID = curCustomer_.UnderSaleStaffID
        myWasteDataCR_.LicenseWeigh = txtLicenseWeigh.Text
        myWasteDataCR_.LicenseWeighBWG = txtLicenseWeighBWG.Text
    '    myWasteDataCR_.companyID = "776EDA5830F744A6B3551348D851BC22"
        myWasteDataCR_.CommentWaste = TxtCommentWaste.Text
        
    '    If cboWasteCase.ListIndex = 1 Then
    '                myWasteDataCR_.WasteCaseID = "10EE3782390242959F296DA722542573"
    '    ElseIf cboWasteCase.ListIndex = 2 Then
    '                myWasteDataCR_.WasteCaseID = "AF543EA5BC1649259AA7955CE4D3C7CD"
    '    Else
    '                myWasteDataCR_.WasteCaseID = ""
    '    End If
        If cboSummary.Text = "รับดำเนินการ" Then
            myWasteDataCR_.WasteApproveStatus = "E" 'Accept
            myWasteDataCR_.Status = "A"
            
            myWasteDataCR_.isTreatAble = True
        End If
       If myWasteDataCR_.isChanged Then
             If curCustomer_.CustomerType = "R" Then
                 myWasteDataCR_.Status = "A"
                 myWasteDataCR_.isTreatAble = True
                 myWasteDataCR_.WasteType = "RDF"
             End If
             If cboWasteCase.Text = "งาน Job" Then
                myWasteDataCR_.SDate = DateValue(DtpStart.Value)
                myWasteDataCR_.EDate = DateValue(DtpEnd.Value)
             End If
            myWasteDataCR_.Save
            If curCustomer_.CustomerType = "R" Then
                 myWasteDataCR_.EnviAccepted
             End If
            If isNew = False Then
                txtWasteNo.Text = myWasteDataCR_.WasteNo
            Else
                txtWasteNo.Text = "NEW No."
                isNew = True
            End If
        End If
        If myWasteDataLab_.isChanged Then
            If ChkCl.Value = 1 Then
                myWasteDataLab_.FreeChlorine = "Y"
                Else
                myWasteDataLab_.FreeChlorine = "N"
            End If
            If ChkCn.Value = 1 Then
                myWasteDataLab_.Cyanide = "Y"
            Else
                myWasteDataLab_.Cyanide = "N"
            End If
            If ChkNt.Value = 1 Then
                myWasteDataLab_.Nitrite = "Y"
            Else
                myWasteDataLab_.Nitrite = "N"
            End If
                        
            'บันทึกผลวิเคราะห์
            If ChkCanEvap.Value = 1 Then
                myWasteDataLab_.CanEvap = ChkCanEvap.Caption
            Else
                myWasteDataLab_.CanEvap = ""
            End If
            myWasteDataLab_.AnalysisNo = "LB-" & myWasteDataCR_.WasteNo
            myWasteDataLab_.LabCode = Trim(TxtLabcode.Text)
            myWasteDataLab_.ExtAs = Trim(TxtExtAs.Text)
            myWasteDataLab_.RemarkAs = Trim(TxtExtAsRemark.Text)
            myWasteDataLab_.ExtCd = Trim(TxtExtCd.Text)
            myWasteDataLab_.RemarkCd = Trim(TxtRemarkCd.Text)
            myWasteDataLab_.ExtCr = Trim(TxtExtCr.Text)
            myWasteDataLab_.RemarkCr = Trim(TxtRemarkCr.Text)
            myWasteDataLab_.ExtCu = Trim(TxtExtCu.Text)
            myWasteDataLab_.RemarkCu = Trim(TxtRemarkCu.Text)
            myWasteDataLab_.ExtPb = Trim(TxtExtPb.Text)
            myWasteDataLab_.RemarkPb = Trim(TxtRemarkPb.Text)
            myWasteDataLab_.ExtNi = Trim(TxtExtNi.Text)
            myWasteDataLab_.RemarkNi = Trim(TxtRemarkNi.Text)
            myWasteDataLab_.ExtZn = Trim(TxtExtZn.Text)
            myWasteDataLab_.RemarkZn = Trim(TxtRemarkZn.Text)
            myWasteDataLab_.ExtHg = Trim(TxtExtHg.Text)
            myWasteDataLab_.RemarkHg = Trim(TxtRemarkHg.Text)
            myWasteDataLab_.SS = Trim(TxtSS.Text)
            myWasteDataLab_.Acidity = Trim(TxtAcidity.Text)
            myWasteDataLab_.Alkalinity = Trim(TxtAlkalinity.Text)
            myWasteDataLab_.Save
            If isEnv = False And isNew = True Then
                cmdApprove.Visible = True
            End If
            If txtFactoryNo.Text = "" Or txtIndustryType.Text = "" Or txtQty.Text = "" Or txtRate.Text = "" Or txtGenerateDesc.Text = "" Or cboStoreType.Text = "" Or cboQtyUnit.Text = "" Then
                cmdApprove.Visible = False
            Else
                cmdApprove.Visible = True
            End If
            If chkBag.Value = 0 And chkDrum.Value = 0 And chkPlasticGallon.Value = 0 And chkBottle.Value = 0 And chkTank.Value = 0 And chkPlasticBag.Value = 0 And chkPlasticDrum.Value = 0 And chkPaperBox.Value = 0 And chkYard.Value = 0 And chkOther2.Value = 0 Then
                cmdApprove.Visible = False
            Else
                cmdApprove.Visible = True
            End If
            If isEnv = False And isNew = True Then
                  If txtFactoryNo.Text = "" Or txtIndustryType.Text = "" Or txtQty.Text = "" Or txtRate.Text = "" Or txtGenerateDesc.Text = "" Or cboStoreType.Text = "" Or cboQtyUnit.Text = "" Then
                      cmdApprove.Visible = False
                  Else
                      cmdApprove.Visible = True
                  End If
              End If
        End If
        If isEnv = False And (myWasteDataCR_.Status = Trim("D") Or myWasteDataCR_.Status = Trim("R") Or myWasteDataCR_.Status = Trim("W")) Then
            cmdApprove.Visible = True
            cmdRev.Visible = False
       Else
            cmdApprove.Visible = False
        End If
        
        'NewProcess
'        If isNew = True Then
'            If Lbl = "NewProcess" Then
'                myWasteDataCR_.isAcceptedByEnvi = True
'                myWasteDataCR_.Status = "P"
'                myWasteDataCR_.WasteApproveStatus = "P"
'                myWasteDataCR_.Save
'                myWasteDataCR_.SaveAccdate
'                myWasteDataLab_.Save
'                myWasteDataCR_.EnviAccepted
'                Lbl = ""
'            End If
'        End If
        
'            Dim a As Integer
'            Dim ServerPic As String
'                    ServerPic = PicPath & curCustomer_.ID & "\" & myWasteDataCR_.ID & "\"
'                For a = 0 To 1
'                    If Check1(a).Value = "1" Then
'                        DBConnExc "update tbPicPath set cansee='Y' where Path='" & ServerPic & AttachesPic(a).FileName & "' "
'                    End If
'                Next
    'งานอื่นที่ไม่ใช่กำจัดกาก

    If Trim(cboWasteType.Text) = "DOC" Then
        myWasteDataCR_.isAcceptedByEnvi = True
        myWasteDataCR_.Status = "A"
        myWasteDataCR_.Save
        myWasteDataCR_.SaveAccdate
        myWasteDataLab_.Save
        myWasteDataCR_.EnviAccepted
    End If
    If myWasteDataCR_.isChanged = False And myWasteDataLab_.isChanged = False Then
        MsgBox "จัดเก็บข้อมูลเรียบร้อยแล้ว", vbExclamation
    End If
End Sub

Private Sub ckValus()
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
'
'            If cboCompany.ListIndex = 0 Then
'                MsgBox "กรุณาเลือกบริษัทที่รับกำจัด", vbExclamation, "Warning"
'                isSave = False
'                cmdApprove.Visible = False
'                Exit Sub
'            End If
            
            If Trim(txtWasteName.Text) = "" Then
                MsgBox "กรุณาระบุชื่อกากของเสีย", vbExclamation
                txtWasteName.SetFocus
                isSave = False
                cmdApprove.Visible = False
                Exit Sub
            End If
            If curCustomer_.CustomerType = "L" Or curCustomer_.CustomerType = "P" Then
            Else
'                If myWasteDataCR_.isQuotationReq = False Then
'                    If Trim(txtFactoryNo.Text) = "" Then
'                        MsgBox "กรุณาระบุเลขทะเบียนโรงงาน", vbExclamation
'                        txtFactoryNo.SetFocus
'                        isSave = False
'                        cmdApprove.Visible = False
'                        Exit Sub
'                    End If
'                End If
'                If Trim(txtWasteGenNo.Text) = "" Then
'                    MsgBox "กรุณาระบุเลขที่ผู้กำเนิดของเสีย", vbExclamation
'                    txtWasteGenNo.SetFocus
'                    isSave = False
'                    cmdApprove.Visible = False
'                    Exit Sub
'                End If
'                If Trim(txtIndustryType.Text) = "" Then
'                    MsgBox "กรุณาระบุประเภทอุตสาหกรรม", vbExclamation
'                    txtIndustryType.SetFocus
'                    isSave = False
'                    cmdApprove.Visible = False
'                    Exit Sub
'                End If
            
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
            End If
'            If chkOther2.Value = 1 Then
'                If Trim(txtOther2.Text) = "" Then
'                    MsgBox "กรุณา ระบุรายละเอียด อื่น ๆ"
'                    cmdApprove.Visible = False
'                    Exit Sub
'                End If
'            End If
End Sub

Private Sub Command1_Click()
    Load frmDateAssign
    With frmDateAssign
        .SetCallerForm Me
        .Show 1
    End With
End Sub

Private Sub Command2_Click()
Load frmReceivedDate
frmReceivedDate.Show 1
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

Private Sub Command4_Click()
Load frmDateAnalysis
frmDateAnalysis.Show 1
End Sub

Private Sub Command5_Click()
  Load frmLicenseStartDate
    With frmLicenseStartDate
        .SetCallerForm Me
        .Show 1
    End With
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 And Me.ActiveControl.Name <> "grKorOrNo" Then SendKeys "{Tab}"
End Sub
Private Sub PopulateCompany()
Dim x%
    cboCompany.Clear
    cboCompany.AddItem "== เลือกบริษัท =="
    For x = 1 To CurrentUser.WasteCompany.Count
        If CurrentUser.WasteCompany(x).isDeleted <> "Y" Then
        cboCompany.AddItem CurrentUser.WasteCompany(x).CompanyNameTH
        End If
        'If CurrentUser.WasteCompany(x).CompanyCode = "AKP" Then cboCompany.AddItem CurrentUser.WasteCompany(x).CompanyNameTH
        'If CurrentUser.WasteCompany(x).CompanyCode = "BWG" Then cboCompany.AddItem CurrentUser.WasteCompany(x).CompanyNameTH
        'If CurrentUser.WasteCompany(x).CompanyCode = "BWC" Then cboCompany.AddItem CurrentUser.WasteCompany(x).CompanyNameTH
    Next
        cboCompany.ListIndex = 0
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0
End Sub
Private Sub Form_Load()
On Error Resume Next
Dim x%
    isNew = True
'    lblAttPic(0).Caption = ""
'    lblAttPic(1).Caption = ""
        AttachesPic(0).FileName = ""
        AttachesPic(1).FileName = ""
        AttachesPic(2).FileName = ""
        AttachesPic(0).FullPath = ""
        AttachesPic(1).FullPath = ""
        AttachesPic(2).FullPath = ""
        
    PicPath = ""
    PicPath1 = ""
    PopulateCompany
    With cboWasteAnalysis
        .AddItem "1.Waste ที่ต้องผ่านการตรวจสอบจากฝ่ายสิ่งแวดล้อม"
        .AddItem "2.Waste ที่ต้องการออกใบเสนอราคาแบบแร่งด่วน"
        .ListIndex = 1
    End With
    With CboWasteYear
        Dim i As Integer
        Dim a As Integer
        a = Year(Now)
        If a < 2557 Then
            a = a + 543
        End If
        For i = 2551 To a
            .AddItem i
        Next i
    End With
    With cboWasteType
        .AddItem "ยังไม่ระบุ"
        For x = 1 To WasteTypes.Count
            .AddItem WasteTypes(x).WasteType
        Next
    End With
    With cboWasteCase
      .AddItem "ปกติ"
      For x = 1 To WasteCases.Count
         .AddItem WasteCases(x).CaseName
      Next
    End With
    With frmPicture
        cmdSave.Picture = .picSave.Picture
        cmdClear.Picture = .picNew.Picture
        cmdReject.Picture = .picDel.Picture
        CmdClose.Picture = .picExit.Picture
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
        .AddItem "ก๊าซ"
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
'    With cboMaxUnit
'        .Clear
'        .AddItem "ตัน"
'        .AddItem "ลิตร"
'    End With
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
    With cboAnalysis
        .Clear
        .AddItem "=เลือก="
        .AddItem "ส่งวิเคราะห์"
        .AddItem "รอวิเคราะห์"
        .AddItem "วิเคราะห์เสร็จเรียบร้อย"
    End With
     With CbSamplingBy
        .Clear
        .AddItem "พิมพ์จันทร์ จินตนาวงศ์"    ' ว-217-จ-8953
        .AddItem "สาวิตรี ลำดับจุด"    'ว-217-จ-8954
        .AddItem "กฤติยาพร มีพันธุ์"  ' ว-217-จ-8955
        .AddItem "สาวิตรี เลี๊ยบประเสริฐ"    ' ว-217-จ-5096
        .AddItem "สุภัคสินี เจริญผล"    'ว-217-จ-6016
        .AddItem "กนกวรรณ สมบัติ"  ' ว-217-จ-6017
        .AddItem "กิตติภา ไชยขันธ์"  ' ว-217-จ-6019
        .AddItem "ศิรินันท์ ทองภู"  ' ว-217-จ-7627
    End With
   datetime
    cboAnalysis.ListIndex = 0
    cboWasteType.ListIndex = 0
    cboWasteChar.ListIndex = 0
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
    
   If CurrentUser.companyID <> "776EDA5830F744A6B3551348D851BC22" Then
     SSTab1.TabVisible(5) = False
     SSTab1.TabVisible(9) = False
     SSTab1.TabVisible(10) = False
   End If

    ReDim Attaches(4)
    SSTab1.Tab = 0
    SSTab1.TabVisible(4) = False
    SSTab1.TabVisible(6) = False
    
'    SSTab1.TabVisible(5) = False
'    If txtFactoryNo.Text = Trim("") Then
'        With curCustomer_
'                txtFactoryNo.Text = .RegisterNo
'        End With
'    End If
    WasteReferent = myWasteDataCR_.ID
End Sub
Public Sub datetime()
    Dim today
    today = Now
    Dim myday
    Dim mymounth
    Dim myyear
Dim acceptedDate As Date
acceptedDate = myyear & mymounth & myday
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Dim FSO As New FileSystemObject
    If FSO.FolderExists("Z:") = True Then
         RemoveDrive "Z:"
    End If
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

Private Sub lblAttPic_Change(Index As Integer)
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

Private Sub Text1_Validate(Index As Integer, Cancel As Boolean)
    If Index = 1 Then
        myWasteDataLab_.Silicon = Trim(Text1(Index).Text)
    ElseIf Index = 3 Then
        myWasteDataLab_.Aluminum = Trim(Text1(Index).Text)
    ElseIf Index = 4 Then
        myWasteDataLab_.Iron = Trim(Text1(Index).Text)
    ElseIf Index = 5 Then
        myWasteDataLab_.Calcium = Trim(Text1(Index).Text)
    ElseIf Index = 6 Then
        myWasteDataLab_.SiliconRemark = Trim(Text1(Index).Text)
    ElseIf Index = 7 Then
        myWasteDataLab_.AluminumRemark = Trim(Text1(Index).Text)
    ElseIf Index = 8 Then
        myWasteDataLab_.IronRemark = Trim(Text1(Index).Text)
    ElseIf Index = 9 Then
        myWasteDataLab_.CalciumRemark = Trim(Text1(Index).Text)
    End If
     myWasteDataLab_.isChanged = True
End Sub

Private Sub textF_Change()
    myWasteDataLab_.f = Trim(textF.Text)
    myWasteDataLab_.isChanged = True
End Sub


Private Sub Txt14Digit_Validate(Cancel As Boolean)
    myWasteDataCR_.Digit14 = Trim(Txt14Digit.Text)
End Sub

Private Sub txtAcomment_Change()
    myWasteDataLab_.CommantLab = Trim(txtAcomment.Text)
    myWasteDataLab_.isChanged = True
End Sub

Private Sub txtAnalysDesc_Validate(Cancel As Boolean)
    myWasteDataLab_.AnalysisDesc = Trim(txtAnalysDesc)
    myWasteDataLab_.isChanged = True
End Sub

Private Sub txtAs_Change()
    myWasteDataLab_.HeavyAs = Trim(txtAs.Text)
    myWasteDataLab_.isChanged = True
End Sub

Private Sub txtAshcontent_Change()
    myWasteDataLab_.AshContent = Trim(txtAshcontent.Text)
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
    If chkBottle.Value = 1 Then myWasteDataLab_.StorageDesc = Trim(txtBottle.Text)
    myWasteDataLab_.Bottle = Trim(txtBottle)
    myWasteDataLab_.isChanged = True
End Sub

Private Sub txtBulkdensity_Change()
    myWasteDataLab_.Bulkdensity = Trim(txtBulkdensity.Text)
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
Private Sub TxtCarOrther_Validate(Cancel As Boolean)
    myWasteDataCR_.CarOrther = Trim(TxtCarOrther.Text)
End Sub

Private Sub txtCd_Change()
    myWasteDataLab_.HeavyCd = Trim(txtCd.Text)
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

Private Sub txtCL_Change()
    myWasteDataLab_.CL = Trim(txtCL.Text)
    myWasteDataLab_.isChanged = True
End Sub

Private Sub txtColor_Validate(Cancel As Boolean)
    myWasteDataLab_.Color = Trim(txtColor)
    myWasteDataLab_.isChanged = True
End Sub

Private Sub TxtCommentWaste_Change()
    myWasteDataCR_.CommentWaste = Trim(TxtCommentWaste.Text)
    myWasteDataCR_.isChanged = True
End Sub

Private Sub TxtConcentrate_Change()
    myWasteDataLab_.Concentrate = Trim(TxtConcentrate.Text)
    myWasteDataLab_.isChanged = True
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

Private Sub txtCr_Change()
    myWasteDataLab_.HeavyCr = Trim(txtCr.Text)
    myWasteDataLab_.isChanged = True
End Sub
Private Sub txtCu_Change()
   myWasteDataLab_.Cu = Trim(txtCu.Text)
    myWasteDataLab_.isChanged = True
End Sub

Private Sub txtCut_Change()
    myWasteDataLab_.Cut = Trim(txtCut.Text)
    myWasteDataLab_.isChanged = True
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
    .MarketingStaffID = curMktStaff_.ID
    .MarketingStaffName = curMktStaff_.StaffFName & " " & curMktStaff_.StaffLName
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
    If isEnv = True Then
            txtLicenseWeigh.Enabled = True
            txtLicenseWeighBWG.Enabled = True
    Else
            txtLicenseWeigh.Enabled = False
            txtLicenseWeighBWG.Enabled = False
    End If
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

Private Sub TxtExtAs_GotFocus()
    TxtExtAs.SelStart = 0
    TxtExtAs.SelLength = Len(TxtExtAs.Text)
End Sub

Private Sub TxtExtAs_Validate(Cancel As Boolean)
    myWasteDataLab_.ExtAs = Trim(TxtExtAs.Text)
    myWasteDataLab_.isChanged = True
End Sub

Private Sub TxtExtAsRemark_GotFocus()
    TxtExtAsRemark.SelStart = 0
    TxtExtAsRemark.SelLength = Len(TxtExtAsRemark.Text)
End Sub

Private Sub TxtExtAsRemark_Validate(Cancel As Boolean)
    myWasteDataLab_.RemarkAs = Trim(TxtExtAsRemark.Text)
    myWasteDataLab_.isChanged = True
End Sub
Private Sub TxtExtCd_GotFocus()
    TxtExtCd.SelStart = 0
    TxtExtCd.SelLength = Len(TxtExtCd.Text)
End Sub

Private Sub TxtExtCd_Validate(Cancel As Boolean)
    myWasteDataLab_.ExtCd = Trim(TxtExtCd.Text)
    myWasteDataLab_.isChanged = True
End Sub
Private Sub TxtExtCr_GotFocus()
    TxtExtCr.SelStart = 0
    TxtExtCr.SelLength = Len(TxtExtCr.Text)
End Sub
Private Sub TxtExtCr_Validate(Cancel As Boolean)
    myWasteDataLab_.ExtCr = Trim(TxtExtCr.Text)
    myWasteDataLab_.isChanged = True
End Sub
Private Sub TxtExtCu_GotFocus()
    TxtExtCu.SelStart = 0
    TxtExtCu.SelLength = Len(TxtExtCu.Text)
End Sub
Private Sub TxtExtCu_Validate(Cancel As Boolean)
    myWasteDataLab_.ExtCu = Trim(TxtExtCu.Text)
    myWasteDataLab_.isChanged = True
End Sub

Private Sub TxtExtHg_GotFocus()
    TxtExtHg.SelStart = 0
    TxtExtHg.SelLength = Len(TxtExtHg.Text)
End Sub

Private Sub TxtExtHg_Validate(Cancel As Boolean)
    myWasteDataLab_.ExtHg = Trim(TxtExtHg.Text)
    myWasteDataLab_.isChanged = True
End Sub

Private Sub TxtExtNi_GotFocus()
    TxtExtNi.SelStart = 0
    TxtExtNi.SelLength = Len(TxtExtNi.Text)
End Sub

Private Sub TxtExtNi_Validate(Cancel As Boolean)
    myWasteDataLab_.ExtNi = Trim(TxtExtNi.Text)
    myWasteDataLab_.isChanged = True
End Sub
Private Sub TxtExtPb_GotFocus()
    TxtExtPb.SelStart = 0
    TxtExtPb.SelLength = Len(TxtExtPb.Text)
End Sub

Private Sub TxtExtPb_Validate(Cancel As Boolean)
    myWasteDataLab_.ExtPb = Trim(TxtExtPb.Text)
    myWasteDataLab_.isChanged = True
End Sub

Private Sub TxtExtZn_GotFocus()
    TxtExtZn.SelStart = 0
    TxtExtZn.SelLength = Len(TxtExtZn.Text)
End Sub

Private Sub TxtExtZn_Validate(Cancel As Boolean)
    myWasteDataLab_.ExtZn = Trim(TxtExtZn.Text)
    myWasteDataLab_.isChanged = True
End Sub

Private Sub TxtFactoryCar_Validate(Cancel As Boolean)
    myWasteDataCR_.FactoryCar = Trim(TxtFactoryCar.Text)
End Sub

'Private Sub txtExpDate_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = 46 Then
'        If MsgBox("ท่านต้องการลบวันที่ใบอนุญาตหมดอายุหรือไม่?", vbQuestion + vbYesNo) = vbYes Then
'            myWasteDataCR_.LicenseExpDate = ""
'            txtExpDate.Text = ""
'        End If
'    End If
'End Sub
Private Sub txtFactoryNo_Validate(Cancel As Boolean)
    myWasteDataCR_.FactoryRegisNo = Trim(txtFactoryNo.Text)
End Sub
'Private Sub txtFlashPoint_KeyPress(KeyAscii As Integer)
'    If IsNumeric(Trim(txtFlashPoint.Text)) Then
'        myWasteDataLab_.FlashPoint = CDbl(Trim(txtFlashPoint.Text))
'    Else
'        myWasteDataLab_.FlashPoint = 0
'    End If
'    myWasteDataLab_.isChanged = True
'End Sub
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

Private Sub txtHeatingValue_Change()
    myWasteDataLab_.HeatingValue = Trim(txtHeatingValue.Text)
    myWasteDataLab_.isChanged = True
End Sub

Private Sub txtHeatingValue_Validate(Cancel As Boolean)
Dim CountRow, i As Integer
myWasteDataLab_.isChanged = True
'    If Val(txtHeatingValue.Text) >= 2800 Then
'        With lvNewTreatment
'            .ListItems.Clear
'            PopulateTreatment
'            .ListItems.item(2).Checked = True
'            CountRow = myWasteDataLab_.NewTreatments.Count
'                For i = 1 To CountRow
'                    myWasteDataLab_.NewTreatments.Remove i
'                Next
'            myWasteDataLab_.NewTreatments.Add colTreatment_(.ListItems.item(2).Key), .ListItems.item(2).Key
'            myWasteDataLab_.isChanged = True
'        End With
'    End If
'    If (Val(txtHeatingValue.Text) >= 500 And Val(txtHeatingValue.Text) < 2800) And cboWasteChar.ListIndex = 3 Then
'        With lvNewTreatment
'            .ListItems.Clear
'            PopulateTreatment
'            .ListItems.item(4).Checked = True
'            CountRow = myWasteDataLab_.NewTreatments.Count
'                For i = 1 To CountRow
'                    myWasteDataLab_.NewTreatments.Remove i
'                Next
'            myWasteDataLab_.NewTreatments.Add colTreatment_(.ListItems.item(4).Key), .ListItems.item(4).Key
'            myWasteDataLab_.isChanged = True
'        End With
'    End If
    
End Sub

Private Sub TxtHeavyFe_Validate(Cancel As Boolean)
    myWasteDataLab_.HeavyFe = Trim(TxtHeavyFe.Text)
    myWasteDataLab_.isChanged = True

End Sub

Private Sub txtHg_Change()
    myWasteDataLab_.HeavyHg = Trim(txtHg.Text)
    myWasteDataLab_.isChanged = True
End Sub

Private Sub txtIndustryType_Validate(Cancel As Boolean)
'    If curIndustryType_ Is Nothing Then Set curIndustryType_ = New BWGIndustryType_New
'    If Trim(txtIndustryType.Text) = "" And Trim(curIndustryType_.TypeDesc) <> Trim(txtIndustryType.Text) Then
'        Set curIndustryType_ = Nothing
'        Set curIndustryType_ = New BWGIndustryType_New
'        With myWasteDataCR_
'            .IndustryTypeID = ""
'            .IndustryTypeName = ""
'            .IndustryTypeIDNew = ""
'            .TypeDesc = ""
'        End With
'    Else
'        If Trim(curIndustryType_.TypeDesc) <> Trim(txtIndustryType.Text) Then
'            Set curIndustryType_ = Nothing
'            Set curIndustryType_ = New BWGIndustryType_New
'            With myWasteDataCR_
'                .IndustryTypeID = ""
'                .IndustryTypeName = ""
'                .IndustryTypeIDNew = ""
'                .TypeDesc = ""
'            End With
'            Load frmIndustryType
'            With frmIndustryType
'                .SetCallerForm Me
'                .Show 1
'            End With
'            If myWasteDataCR_.IndustryTypeIDNew = "" Then txtIndustryType.Text = ""
'        End If
'    End If
    If curCustomer_.IndustryTypeID <> "" Then
        myWasteDataCR_.IndustryTypeIDNew = curCustomer_.IndustryTypeIDNew
    End If
End Sub
Private Sub TxtLabcode_GotFocus()
    TxtLabcode.SelStart = 0
    TxtLabcode.SelLength = Len(TxtLabcode.Text)
End Sub

Private Sub TxtLabcode_Validate(Cancel As Boolean)
    myWasteDataLab_.LabCode = Trim(TxtLabcode.Text)
    myWasteDataLab_.isChanged = True
End Sub

Private Sub txtLicenseWeigh_Validate(Cancel As Boolean)
    If IsNumeric(Trim(txtLicenseWeigh.Text)) Then
        myWasteDataCR_.LicenseWeigh = CDbl(Trim(txtLicenseWeigh.Text))
    Else
        myWasteDataCR_.LicenseWeigh = 0
    End If
     myWasteDataCR_.isChanged = True
End Sub

Private Sub txtLicenseWeighBWG_Validate(Cancel As Boolean)
 If IsNumeric(Trim(txtLicenseWeighBWG.Text)) Then
        myWasteDataCR_.LicenseWeighBWG = CDbl(Trim(txtLicenseWeighBWG.Text))
    Else
        myWasteDataCR_.LicenseWeighBWG = 0
    End If
     myWasteDataCR_.isChanged = True
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

Private Sub txtMn_Change()
    myWasteDataLab_.Mn = Trim(txtMn.Text)
    myWasteDataLab_.isChanged = True
End Sub

Private Sub txtmoisturecontent_Change()
    myWasteDataLab_.MoistureContent = Trim(txtmoisturecontent.Text)
    myWasteDataLab_.isChanged = True
End Sub
'
'Private Sub txtMoud_Validate(Cancel As Boolean)
'    myWasteDataCR_.Moud = txtMoud.Text
'End Sub

Private Sub txtNCVValue_Change()
myWasteDataLab_.NCVValue = Trim(txtNCVValue.Text)
myWasteDataLab_.isChanged = True
End Sub

Private Sub txtNi_Change()
myWasteDataLab_.Ni = Trim(txtNi.Text)
myWasteDataLab_.isChanged = True
End Sub

Private Sub txtOdour_Validate(Cancel As Boolean)
    myWasteDataLab_.Odour = Trim(txtOdour)
    myWasteDataLab_.isChanged = True
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
    myWasteDataLab_.StorageDesc = Trim(txtPaperBox)
    myWasteDataLab_.isChanged = True
End Sub

Private Sub txtPb_Change()
    myWasteDataLab_.HeavyPb = Trim(txtPb.Text)
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

Private Sub txtPhysicalstate_Change()
    myWasteDataLab_.Physicalstate = Trim(txtPhysicalstate.Text)
    myWasteDataLab_.isChanged = True
End Sub

Private Sub txtPlasticBag_Validate(Cancel As Boolean)
    myWasteDataLab_.PlasticBag = Trim(txtPlasticBag)
    myWasteDataLab_.StorageDesc = Trim(txtPlasticBag)
    myWasteDataLab_.isChanged = True
End Sub
Private Sub txtPlasticDrum_Validate(Cancel As Boolean)
    myWasteDataLab_.PlasticDrum = Trim(txtPlasticDrum)
    myWasteDataLab_.StorageDesc = Trim(txtPlasticDrum)
    myWasteDataLab_.isChanged = True
End Sub
Private Sub txtPlasticGallon_Validate(Cancel As Boolean)
    If Trim(txtPlasticGallon.Text) <> "" Then
        myWasteDataLab_.StorageDesc = Trim(txtPlasticGallon)
    End If
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

Private Sub txtQuoPrice_Change()
    myWasteDataCR_.QuoPrice = Trim(txtQuoPrice)
    myWasteDataCR_.isChanged = True
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

Private Sub TxtRemarkCd_GotFocus()
    TxtRemarkCd.SelStart = 0
    TxtRemarkCd.SelLength = Len(TxtRemarkCd.Text)
End Sub

Private Sub TxtRemarkCd_Validate(Cancel As Boolean)
    myWasteDataLab_.RemarkCd = Trim(TxtRemarkCd.Text)
    myWasteDataLab_.isChanged = True
End Sub

Private Sub TxtRemarkCr_GotFocus()
    TxtRemarkCr.SelStart = 0
    TxtRemarkCr.SelLength = Len(TxtRemarkCr.Text)
End Sub

Private Sub TxtRemarkCr_Validate(Cancel As Boolean)
    myWasteDataLab_.RemarkCr = Trim(TxtRemarkCr.Text)
    myWasteDataLab_.isChanged = True
End Sub

Private Sub TxtRemarkCu_GotFocus()
    TxtRemarkCu.SelStart = 0
    TxtRemarkCu.SelLength = Len(TxtRemarkCu.Text)
End Sub

Private Sub TxtRemarkCu_Validate(Cancel As Boolean)
    myWasteDataLab_.RemarkCu = Trim(TxtRemarkCu.Text)
    myWasteDataLab_.isChanged = True
End Sub

Private Sub TxtRemarkHg_GotFocus()
    TxtRemarkHg.SelStart = 0
    TxtRemarkHg.SelLength = Len(TxtRemarkHg.Text)
End Sub

Private Sub TxtRemarkHg_Validate(Cancel As Boolean)
    myWasteDataLab_.RemarkHg = Trim(TxtRemarkHg.Text)
    myWasteDataLab_.isChanged = True
End Sub

Private Sub TxtRemarkNi_GotFocus()
    TxtRemarkNi.SelStart = 0
    TxtRemarkNi.SelLength = Len(TxtRemarkNi.Text)
End Sub

Private Sub TxtRemarkNi_Validate(Cancel As Boolean)
    myWasteDataLab_.RemarkNi = Trim(TxtRemarkNi.Text)
    myWasteDataLab_.isChanged = True
End Sub

Private Sub TxtRemarkPb_GotFocus()
    TxtRemarkPb.SelStart = 0
    TxtRemarkPb.SelLength = Len(TxtRemarkPb.Text)
End Sub

Private Sub TxtRemarkPb_Validate(Cancel As Boolean)
    myWasteDataLab_.RemarkPb = Trim(TxtRemarkPb.Text)
    myWasteDataLab_.isChanged = True
End Sub

Private Sub TxtRemarkZn_GotFocus()
    TxtRemarkZn.SelStart = 0
    TxtRemarkZn.SelLength = Len(TxtRemarkZn.Text)
End Sub

Private Sub TxtRemarkZn_Validate(Cancel As Boolean)
    myWasteDataLab_.RemarkZn = Trim(TxtRemarkZn.Text)
    myWasteDataLab_.isChanged = True
End Sub

Private Sub txtS_Change()
    myWasteDataLab_.s = Trim(txtS.Text)
    myWasteDataLab_.isChanged = True
End Sub

Private Sub txtSludgecontent_Change()
    myWasteDataLab_.SludgeContent = Trim(txtSludgecontent.Text)
    myWasteDataLab_.isChanged = True
End Sub

Private Sub txtsolidcontent_Change()
    myWasteDataLab_.SolidContent = Trim(txtsolidcontent.Text)
    myWasteDataLab_.isChanged = True
End Sub

Private Sub txtToxicDesc_Validate(Cancel As Boolean)
    myWasteDataLab_.ToxicDesc = Trim(txtToxicDesc.Text)
    myWasteDataLab_.isChanged = True
End Sub

Private Sub txtviscosity_Change()
myWasteDataLab_.Viscosity = Trim(txtviscosity.Text)
myWasteDataLab_.isChanged = True
End Sub

Private Sub txtWasteGenNo_Validate(Cancel As Boolean)
    myWasteDataCR_.WasteCreateNo = Trim(txtWasteGenNo.Text)
End Sub
Private Sub txtWasteName_Validate(Cancel As Boolean)
    myWasteDataCR_.WasteName = Trim(txtWasteName.Text)
    
'    Dim RS As ADODB.Recordset
'    Dim tSql As String
'    tSql = "SELECT * FROM tbEUWaste" & vbCrLf
'    tSql = tSql & "WHERE WasteName = '" & Trim(txtWasteName.Text) & "'" & vbCrLf
'    Set RS = New ADODB.Recordset
'    Set RS = GetRS(tSql, adOpenForwardOnly, adLockReadOnly)
'    If Not RS.EOF Then
'    Lbl = "NewProcess"
'        txtEUCode1.Text = Left(RS!FullCode, 2)
'        txtEUCode2.Text = Mid(RS!FullCode, 4, 2)
'        txtEUCode3.Text = Mid(RS!FullCode, 7, 2)
'        txtHAType.Text = RS!HAType
'        myWasteDataCR_.EUCode1 = txtEUCode1.Text
'        myWasteDataCR_.EUCode1ID = RS!Code1ID
'        myWasteDataCR_.EUCode2 = txtEUCode2.Text
'        myWasteDataCR_.EUCode2ID = RS!Code2ID
'        myWasteDataCR_.EUCode3 = txtEUCode3.Text
'        myWasteDataCR_.EUCode3ID = RS!Code3ID
'    End If
End Sub
Private Sub PopulateTreatment()
Dim tmpS As New BWGSearchManager
Dim x%
    Set colTreatment_ = tmpS.TreatmentDataSearch
'    lvOldTreatment.ListItems.Clear
    lvNewTreatment.ListItems.Clear
    For x = 1 To colTreatment_.Count
'        lvOldTreatment.ListItems.Add , "'" & colTreatment_(x).ID & "'", colTreatment_(x).TreatmentDesc
        lvNewTreatment.ListItems.Add , "'" & colTreatment_(x).ID & "'", colTreatment_(x).TreatmentDesc
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
        myWasteDataCR_.MarketingStaffID = .UnderSaleStaffID
        txtMktStaff.Text = .UnderSaleName
    End With
    '******
             With curCustomer_
                 txtFactoryNo.Text = .RegisterNo
                 Txt14Digit.Text = .Digit14
                 txtIndustryType.Text = .TypeDesc
                 txtWasteGenNo.Text = .WasteManageID
                 myWasteDataCR_.IndustryTypeIDNew = .IndustryTypeIDNew
             End With
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
'        DBManager.GetDBConnection.Execute "Update tbWasteDataLab Set WasteDataID='" & tmpGUID & "' Where WasteDataID='" & myWasteDataCR_.ID & "'"
'        If Trim(tmpStr) = "" Then
'            myWasteDataLab_.Delete
'        Else
'            DBManager.GetDBConnection.Execute "Update tbWasteDataLab Set WasteDataID='" & myWasteDataCR_.ID & "' Where WasteDataID='" & tmpGUID & "'"
'            MsgBox tmpStr, vbExclamation
'        End If
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
'    txtMoud.Text = ""
'    txtKor.Text = ""
    cboLABRes.ListIndex = -1
    txtWasteName.Text = ""
    txtQty.Text = ""
    cboQtyUnit.ListIndex = -1
'    cboMaxUnit.ListIndex = -1
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
    txtFactoryNo.Text = ""
    txtWasteGenNo.Text = ""
'    txtExpDate.Text = ""
    txtIndustryType.Text = ""
    TxtLabcode.Text = ""
    TxtExtAs.Text = ""
    TxtExtAsRemark.Text = ""
    TxtExtCd.Text = ""
    TxtRemarkCd.Text = ""
    TxtExtCr.Text = ""
    TxtRemarkCr.Text = ""
    TxtExtCu.Text = ""
    TxtRemarkCu.Text = ""
    TxtExtHg.Text = ""
    TxtRemarkHg.Text = ""
    TxtExtNi.Text = ""
    TxtRemarkNi.Text = ""
    TxtExtPb.Text = ""
    TxtRemarkPb.Text = ""
    TxtExtZn.Text = ""
    TxtRemarkZn.Text = ""
    cboTreatmentType.ListIndex = -1
    Set curIndustryType_ = Nothing
    Set curIndustryType_ = New BWGIndustryType_New
    grKorOrNo.ClearAllData
    grKorOrNo.getGridObj.Rows = 20
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
Dim WType, WType1, WType2
    cmdAccepdate.Visible = False
    Set myWasteDataCR_ = selWaste
    With myWasteDataCR_
        WasteID = .ID
        WType = .WasteType
        If isEnv = True Then
            If curCustomer_.isNew = True Then
                cmdApprove.Visible = False
                cmdReject.Visible = False
                cmdSave.Enabled = False
                cmdDel.Enabled = False
                cmdRev.Visible = False
            Else
                    If Trim(.Status) <> "R" Then
                            cmdApprove.Visible = False
                            cmdReject.Visible = False
                        If Trim(.Status) = "J" Then
                                cmdSave.Enabled = False
                                cmdDel.Enabled = False
                                cmdAccepdate.Visible = False
                                isNew = True
                        ElseIf Trim(.Status) = "W" Then
                                cmdSave.Enabled = False
                                cmdDel.Enabled = False
                                cmdAccepdate.Visible = False
                                isNew = True
                        Else
                                cmdSave.Enabled = True
                                cmdAccepdate.Visible = True
                                isNew = False
                        End If
                    ElseIf Trim(.Status) = "R" Then
                        cmdApprove.Visible = True
                        cmdReject.Visible = True
                        cmdSave.Enabled = False
                        cmdRev.Visible = False
                         cmdDel.Enabled = False
                         cmdAccepdate.Visible = False
                        isNew = True
                End If
            End If
            
            If Trim(.Status) = "AR" Then
            
                    'cmdRev.Visible = True
                    cmdApprove.Visible = False
                    cmdSave.Enabled = False
            ElseIf Trim(.Status) = "IA" Then
                    cmdRev.Visible = False
                    cmdApprove.Visible = False
              ElseIf Trim(.Status) = "RJ" Then
                    cmdApprove.Visible = True
                    cmdRev.Visible = False
            End If
        Else
            cmdAccepdate.Visible = False
            If .isQuotationReq = False Then
                If Trim(.Status) = "D" Or Trim(.Status) = "R" Then
                    cmdApprove.Visible = True
                    cmdRev.Visible = False
                    cmdSave.Enabled = True
                    cmdDel.Enabled = True
                Else
                    cmdApprove.Visible = False
                    cmdSave.Enabled = False
                    cmdDel.Enabled = False
                End If
            Else
                If Trim(.Status) = "J" Then
                    cmdApprove.Visible = True
                    cmdSave.Enabled = True
                    cmdDel.Enabled = True
                Else
                    cmdApprove.Visible = False
                    cmdSave.Enabled = True
                    cmdDel.Enabled = True
                End If
            End If
            If Trim(.Status) = "A" Then
                    'cmdRev.Visible = True
                    cmdApprove.Visible = False
            ElseIf Trim(.Status) = "AR" Then
                cmdApprove.Visible = False
                cmdRev.Visible = False
            End If
            If Trim(.Status) = "A" Then
                cmdApprove.Visible = False
                'cmdRev.Visible = True
            End If
            isNew = .isNew
        End If
        If Trim(.Status) = "P" Or Trim(.Status) = "A" Or Trim(.Status) = "E" Then
            If .acceptedDate = "" Then
                cmdAccepdate.Visible = True
                'cmdRev.Visible = True
                cboWasteAnalysis.Enabled = False
                Else
                cboWasteAnalysis.Enabled = True
            End If
            Label1.Caption = "Envi ตรวจรับแล้วเมื่อ " & FormatYMD_to_DMY(.acceptedDate, "/", "/")
            Label1.Visible = True
            Label3.Caption = "ตรวจรับแล้ว"
            Label3.Visible = True
            If isEnv = True Then
                cmdSave.Enabled = True
                cmdDel.Enabled = True
            Else
                cmdSave.Enabled = False
                cmdDel.Enabled = False
                Frame2.Enabled = False
            End If
        ElseIf Trim(.Status) = "D" Then
            Label1.Caption = ""
            Label1.Visible = True
            If .isQuotationReq = False Then
                Label3.Caption = "ยังไม่ส่งข้อมูล"
            Else
                Label3.Caption = "ขอออกใบเสนอราคาเร่งด่วน"
            End If
            Label3.Visible = True
            Label4.Caption = ""
            Label4.Visible = False
        ElseIf Trim(.Status) = "J" Then
            Label1.Caption = "Envi ตีกลับเมื่อ " & FormatYMD_to_DMY(.acceptedDate, "/", "/")
            Label1.Visible = True
            Label3.Caption = "ตีกลับ"
            Label3.Visible = True
            Label4.Caption = .RejectDesc
            Label4.Visible = True
            cmdSave.Enabled = True
            cmdDel.Enabled = True
        ElseIf Trim(.Status) = "RJ" Then
            Label3.Caption = "ตีกลับ"
            Label3.Visible = True
            Label4.Caption = .RejectDesc
            Label4.Visible = True
            cmdSave.Enabled = True
            cmdDel.Enabled = False
            cmdApprove.Visible = True
        ElseIf Trim(.Status) = "R" And isEnv = True Then
            Label3.Caption = "รอการตรวจรับ"
            cmdApprove.Visible = True
            cmdReject.Visible = True
            Label3.Visible = True
            Label4.Caption = "ส่งให้ Env เมื่อ" & Format(.SaleSendDateAuto, "dd/mm/yyyy")
            Label4.Visible = True
        ElseIf Trim(.Status) = "R" And isEnv = False Then
            Label3.Caption = "ส่งให้ Env เมื่อ"
            Label4.Caption = Format(.SaleSendDateAuto, "dd/mm/yyyy")
            Label3.Visible = True
            Label4.Visible = True
        ElseIf Trim(.Status) = "W" Then
            Label3.Caption = "รออนุมัติใบเสนอราคา"
            Label4.Caption = "ให้ Sales เก็บข้อมูลเพิ่ม"
            Label3.Visible = True
            Label4.Visible = True
            If isEnv = False Then cmdApprove.Visible = True
        ElseIf Trim(.Status) = "AR" Then
            Label3.Caption = "ต้องการ Revised Waste "
            'Label4.Caption = "ให้ Sales เก็บข้อมูลเพิ่ม"
            Label3.Visible = True
            'Label4.Visible = True
            If isEnv = True Then
                cmdApprove.Visible = False
                cmdReject.Visible = True
                'cmdRev.Visible = True
             Else
                cmdSave.Enabled = True
                cmdReject.Visible = False
                cmdRev.Visible = False
            End If
        ElseIf Trim(.Status) = "IA" Then
            Label3.Visible = True
            Label3.Caption = " Waste  ถูก Revised"
            Label4.Visible = False
            cmdRev.Visible = False
            cboSummary.Text = "รับดำเนินการ"
            cboSummary.Enabled = False
            cmdSave.Visible = False
            cmdClear.Visible = False
            cmdDel.Value = False
            SSTab1.TabEnabled(0) = False
            SSTab1.TabEnabled(1) = False
            SSTab1.TabEnabled(2) = False
            SSTab1.TabEnabled(3) = False
            SSTab1.TabEnabled(4) = False
            SSTab1.TabEnabled(5) = False
            SSTab1.TabEnabled(6) = False
            SSTab1.TabEnabled(7) = False
            SSTab1.TabEnabled(8) = False
        End If

        '++++++++++++++++++++++++++++++++++++++++++++++++++++
        txtQuoPrice.Text = .QuoPrice
        txtLicenseWeigh.Text = Trim(.LicenseWeigh)
        txtLicenseWeighBWG.Text = Trim(.LicenseWeighBWG)
        If isEnv = True Then
            txtLicenseWeigh.Enabled = True
            txtLicenseWeighBWG.Enabled = True
        Else
            txtLicenseWeigh.Enabled = False
            txtLicenseWeighBWG.Enabled = False
        End If
        WasteReferent = .ID
        CboWasteYear.Text = .WasteYears
        txtWasteNo.Text = .WasteNo
        txtContractNo.Text = .Contract.docNo
        If isExist(CurrentUser.WorkCompanies, "'" & .companyID & "'") Then
            SetTextToCombo CurrentUser.WorkCompanies("'" & .companyID & "'").CompanyNameTH, cboCompany
        End If
        If .companyID = "" Then
            cboCompany.ListIndex = 0
        Else
          For x = 1 To CurrentUser.WasteCompany.Count
            If CurrentUser.WasteCompany(x).ID = .companyID Then
            cboCompany.ListIndex = x
            End If
          Next
        End If
        If isAkp = True And isBWG = False Then
            cboCompany.Visible = False
        End If
        Set tmpCol = tmpS.EUCode1Search("EUCode1ID='" & .EUCode1ID & "'")
        If tmpCol.Count > 0 Then SetDataEUCode1 tmpCol(1), False
        Set tmpCol = tmpS.EUCode2Search("EUCode2ID='" & .EUCode2ID & "'")
        If tmpCol.Count > 0 Then SetDataEUCode2 tmpCol(1), False
        Set tmpCol = tmpS.EUCode3Search("EUCode3ID='" & .EUCode3ID & "'")
        If tmpCol.Count > 0 Then SetDataEUCode3 tmpCol(1), False
             
        SetTextToCombo .WasteType, cboWasteType
        If cboWasteType.ListIndex = -1 Then cboWasteType.ListIndex = 0
        
        If .WasteCaseID = "A1" Then
            cboWasteCase.Text = "RDF"
        ElseIf .WasteCaseID = "AF543EA5BC1649259AA7955CE4D3C7CD" Then
            cboWasteCase.Text = "เวสต์ทิ้ง"
        ElseIf .WasteCaseID = "10EE3782390242959F296DA722542573" Then
            cboWasteCase.Text = "เวสต์น้ำท่วม กรอ."
        ElseIf .WasteCaseID = "SHazcleening101" Then
            cboWasteCase.Text = "Haz Cleaning"
         ElseIf .WasteCaseID = "SNonLow101" Then
            cboWasteCase.Text = "Non ราคาถูก"
        ElseIf .WasteCaseID = "" Then
            cboWasteCase.Text = "ปกติ"
        ElseIf .WasteCaseID = "1" Then
            cboWasteCase.Text = "งาน Job"
        Else
            cboWasteCase.ListIndex = 0
        End If
        If .SDate <> "" Then
        DtpStart.Value = .SDate
        End If
        If .EDate <> "" Then
        DtpEnd.Value = .EDate
        End If
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
'            txtMoud.Text = .Moud
'            txtKor.Text = .Kor
            txtWasteName.Text = .WasteName
            txtQty.Text = .NormalQty
'            txtMaxQty.Text = .MaxQty
            txtRate.Text = .WasteRate
            If curCustomer_.RegisterNo <> "" Then
                txtFactoryNo.Text = curCustomer_.RegisterNo
            Else
                txtFactoryNo.Text = .FactoryRegisNo
             End If
'            txtFactoryNo.Text = .FactoryRegisNo
'            Txt14Digit.Text = .Digit14
            If Trim(.CanEvap) <> "" Then
                ChkCanEvap.Value = 1
                Else
                ChkCanEvap.Value = 0
            End If
            If curCustomer_.WasteManageID <> "" Then
                txtWasteGenNo.Text = curCustomer_.WasteManageID
            Else
                txtWasteGenNo.Text = .WasteCreateNo
            End If
'            txtWasteGenNo.Text = .WasteCreateNo
            ChkisTruckTen.Value = Abs(CInt(.isTruckTen))
            Chkispickup.Value = Abs(CInt(.isPickup))
            ChkisLugger.Value = Abs(CInt(.isLugger))
            ChkisRollOff.Value = Abs(CInt(.isRollOff))
            ChkisHuerta.Value = Abs(CInt(.isHuerta))
            ChkisTrailerW.Value = Abs(CInt(.isTrailerW))
            ChkisTankcar.Value = Abs(CInt(.isTankcar))
            ChkisFactoryCar.Value = Abs(CInt(.isFactoryCar))
            ChkisOrtherCar.Value = Abs(CInt(.isOrtherCar))
            TxtCarOrther.Text = .CarOrther
            TxtCommentWaste.Text = .CommentWaste
            TxtFactoryCar.Text = .FactoryCar

        If .LicenseStartDate = "" And .LicenseExpDate = "" Then
            txtStartDate.Text = ""
        Else
            txtStartDate.Text = FormatYMD_to_DMY(.LicenseStartDate, "/", "/") + " ถึง " + FormatYMD_to_DMY(.LicenseExpDate, "/", "/")
        End If
        If .NormalUnit = "T" Then
            cboQtyUnit.ListIndex = 0
        ElseIf .NormalUnit = "L" Then
            cboQtyUnit.ListIndex = 1
        End If
'        If .MaxUnit = "T" Then
'            cboMaxUnit.ListIndex = 0
'        ElseIf .MaxUnit = "L" Then
'            cboMaxUnit.ListIndex = 1
'        End If
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
        If .IndustryTypeIDNew <> "" Then
        Set tmpCol = tmpS.IndustryTypeSearch("TypeID='" & .IndustryTypeIDNew & "'")
        If tmpCol.Count > 0 Then SetDataIndustryType tmpCol(1), False
        End If
        callStaffType_ = "ENV"
        Set tmpCol = tmpS.UserSearch("StaffID='" & .EnviStaffID & "'")
        If tmpCol.Count > 0 Then SetDataStaff tmpCol(1), False
        If Trim(.Status) = "P" Then
            cboSummary.ListIndex = 0
        ElseIf Trim(.Status) = "A" Then
            cboSummary.ListIndex = 1
        ElseIf Trim(.Status) = "E" Then
            cboSummary.ListIndex = 2
        End If
        '*************

        If myWasteDataCR_.Status = "A" Then
         'cmdRev.Visible = True
         Else
          cmdRev.Visible = False
          End If
        
'        If isEnv = False And cboSummary.Text = "ยังไม่สรุปผล" Then
'            cmdRev.Visible = False
'        ElseIf isEnv = False And cboSummary.Text = "รับดำเนินการ" And txtWasteNo.Text <> "==NEW==" And Trim(Label3.Caption) <> "Waste  ถูก Revised" Then
'            cmdRev.Visible = True
        If isEnv = True And cboSummary.Text = "ยังไม่สรุปผล" And (Label3.Caption = "รอการตรวจรับ") Then
            cmdApprove.Visible = True
        End If
        '++++++++++++++++++++++++++++++++++++++++++++++++++++

        If .isQuotationReq = True Then
            cboWasteAnalysis.ListIndex = 1
            If Trim(.Status) <> "D" Then
              cboWasteAnalysis.Enabled = False
            End If
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
        If Trim(myWasteDataCR_.Status) = "A" Or Trim(myWasteDataCR_.Status) = "P" Or Trim(myWasteDataCR_.Status) = "E" Or Trim(myWasteDataCR_.Status) = "IA" Or Trim(myWasteDataCR_.Status) = "AR" Then
            myWasteDataCR_.isNew = False
            isNew = False
        Else
            myWasteDataCR_.isNew = True
            isNew = True
        End If
        Set myWasteDataLab_ = Nothing
        Set myWasteDataLab_ = New BWGWasteDataLAB
        myWasteDataLab_.ID = GetGUID
        myWasteDataLab_.wastedataID = myWasteDataCR_.ID
                Set tmpCol = tmpS.WasteDataLABSearch("WasteDataID='" & .ID & "'", isNew, "vw_WasteDataLABSearch2")
        If tmpCol.Count > 0 Then Set myWasteDataLab_ = tmpCol(1)
        .isChanged = False
            
                Dim DD, mm As String
                    If .acceptedDate <> "" Then
                    DD = DateAdd("m", 3, .acceptedDate)
                    mm = DateAdd("m", -3, Date)
                    Else
                        DD = ""
                        mm = ""
                    End If
                    If Trim(.WasteYears) = "2563" Or (mm <= DD And DD <= Date) Or DD = "" Then
                        cmdRev.Visible = False
'                        Else
'                        cmdRev.Visible = True
                    End If
                    If isEnv = False And Left(.WasteNo, 2) = "63" Then
                        cmdRev.Visible = False
                        ElseIf isEnv = False And Left(.WasteNo, 2) < "63" Then
                         'cmdRev.Visible = True
                    End If
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
            Case "G"
                cboWasteChar.ListIndex = 4
        End Select
        chkExplode.Value = Abs(CInt(.isExplode))
        chkInflammable.Value = Abs(CInt(.InFlammable))
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
            Case "BA" 'ถุงขนาดใหญ่
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
            Case "PG"
                chkPlasticGallon.Value = 1
                txtPlasticGallon.Text = .StorageDesc
            Case "BT"
                chkBottle.Value = 1
                txtBottle.Text = .StorageDesc
             Case "PT"
                chkPlasticBag.Value = 1
                txtPlasticBag.Text = .StorageDesc
            Case "PD"
                chkPlasticDrum.Value = 1
                txtPlasticDrum.Text = .StorageDesc
            Case "PB"
                chkPaperBox.Value = 1
                txtPaperBox.Text = .StorageDesc
        End Select
    
        .isNew = isNew
'        txtOldStorePlace.Text = .OldEliminator
'        For x = 1 To .OldTreatments.Count
'            lvOldTreatment.ListItems("'" & .OldTreatments(x).ID & "'").Checked = isExist(lvOldTreatment.ListItems, "'" & .OldTreatments(x).ID & "'")
'        Next
        For x = 1 To .NewTreatments.Count
            lvNewTreatment.ListItems("'" & .NewTreatments(x).ID & "'").Checked = isExist(lvNewTreatment.ListItems, "'" & .NewTreatments(x).ID & "'")
        Next
        '***********For Analysis **********************
        txtviscosity.Text = Trim(.Viscosity)
        txtBulkdensity.Text = Trim(.Bulkdensity)
        lbAnalysisNo.Caption = "LB-" & Trim(myWasteDataCR_.WasteNo)
        ' FreeChlorine********************************
        If .FreeChlorine = "Y" Then
            ChkCl.Value = 1
        Else
            ChkCl.Value = 0
        End If

        ' Cyanide************************************
        If .Cyanide = "Y" Then
            ChkCn.Value = 1
        Else
            ChkCn.Value = 0
        End If
        ' Nitrite***************************************
        If .Nitrite = "Y" Then
            ChkNt.Value = 1
        Else
            ChkNt.Value = 0
        End If
        '//*******************
        txtPhysicalstate.Text = Trim(.Physicalstate)
        txtHeatingValue.Text = Trim(.HeatingValue)
        txtNCVValue.Text = Trim(.NCVValue)
        txtmoisturecontent.Text = Trim(.MoistureContent)
        txtWatercontent.Text = Trim(.WaterContent)
        txtAshcontent.Text = Trim(.AshContent)
        txtCL.Text = Trim(.CL)
        textF.Text = Trim(.f)
        txtS.Text = Trim(.s)
        txtSludgecontent.Text = Trim(.SludgeContent)
        txtsolidcontent.Text = Trim(.SolidContent)
        txtAs.Text = Trim(.HeavyAs)
        txtCd.Text = Trim(.HeavyCd)
        txtCr.Text = Trim(.HeavyCr)
        txtHg.Text = Trim(.HeavyHg)
        txtPb.Text = Trim(.HeavyPb)
        TxtConcentrate.Text = Trim(.Concentrate)
        TxtSS.Text = Trim(.SS)
        TxtAcidity.Text = Trim(.Acidity)
        TxtAlkalinity.Text = Trim(.Alkalinity)
        TxtHeavyFe.Text = Trim(.HeavyFe)
        txtAcomment.Text = Trim(.CommantLab)
        txtCu.Text = Trim(.Cu)
        txtNi.Text = Trim(.Ni)
        txtMn.Text = Trim(.Mn)
        txtZn.Text = Trim(.Zn)
        txtCut.Text = Trim(.Cut)
        lbReceivedDate.Caption = Trim(.ReceivedDate)
        lbAnalysisDate.Caption = Trim(.AnalysisDate)
        lbSameplingDate.Caption = Trim(.SamplingDate)
        TxtLabcode.Text = .LabCode
        TxtExtAs.Text = .ExtAs
        TxtExtAsRemark.Text = .RemarkAs
        TxtExtCd.Text = .ExtCd
        TxtRemarkCd.Text = .RemarkCd
        TxtExtCr.Text = .ExtCr
        TxtRemarkCr.Text = .RemarkCr
        TxtExtCu.Text = .ExtCu
        TxtRemarkCu.Text = .RemarkCu
        TxtExtPb.Text = .ExtPb
        TxtRemarkPb.Text = .RemarkPb
        TxtExtNi.Text = .ExtNi
        TxtRemarkNi.Text = .RemarkNi
        TxtExtZn.Text = .ExtZn
        TxtRemarkZn.Text = .RemarkZn
        TxtExtHg.Text = .ExtHg
        TxtRemarkHg.Text = .RemarkHg
        Text1(1).Text = .Silicon
        Text1(3).Text = .Aluminum
        Text1(4).Text = .Iron
        Text1(5).Text = .Calcium
        Text1(6).Text = .SiliconRemark
        Text1(7).Text = .AluminumRemark
        Text1(8).Text = .IronRemark
        Text1(9).Text = .CalciumRemark
        
        
        If Trim(.LBstatus) = "D" Then
            cboAnalysis.ListIndex = 0
        ElseIf Trim(.LBstatus) = "S" Then
            cboAnalysis.ListIndex = 1
        ElseIf Trim(.LBstatus) = "R" Then
            cboAnalysis.ListIndex = 2
        ElseIf Trim(.LBstatus) = "A" Then
            cboAnalysis.ListIndex = 3
        Else
            cboAnalysis.ListIndex = 0
        End If
                
        If Trim(.SamplingBy) <> "" Then
            If .SamplingBy = "พิมพ์จันทร์ จินตนาวงศ์" Then
                CbSamplingBy.ListIndex = 0
            ElseIf .SamplingBy = "สาวิตรี ลำดับจุด" Then
                CbSamplingBy.ListIndex = 1
            ElseIf .SamplingBy = "กฤติยาพร มีพันธุ์" Then
                CbSamplingBy.ListIndex = 2
            ElseIf .SamplingBy = "สาวิตรี เลี๊ยบประเสริฐ" Then
                CbSamplingBy.ListIndex = 3
            ElseIf .SamplingBy = "สุภัคสินี เจริญผล" Then
                CbSamplingBy.ListIndex = 4
            ElseIf .SamplingBy = "กนกวรรณ สมบัติ" Then
                CbSamplingBy.ListIndex = 5
            ElseIf .SamplingBy = "กิตติภา ไชยขันธ์" Then
                CbSamplingBy.ListIndex = 6
            ElseIf .SamplingBy = "ศิรินันท์ ทองภู" Then
                CbSamplingBy.ListIndex = 7
            End If
        End If
        
        txtAcomment.Text = Trim(.CommantLab)
        '********** for new Varliable **************
        chkHospital = IIf(.isHospital = True, 1, 0)
        txtOdour = .Odour
        chkOdour = IIf(.isOdour = True, 1, 0)
        txtColor = .Color
        chkColor = IIf(.isColor = True, 1, 0)
        
'        'แกนลอนพลาสติก
'        chkPlasticGallon = IIf(.StorageType = "PG", 1, 0)
'        txtPlasticGallon = .StorageDesc
'        If .StorageType <> "PG" Then txtPlasticGallon.Text = ""
'        'ขวดแก้ว
'        chkBottle = IIf(.StorageType = "BT", 1, 0)
'        txtBottle = .StorageDesc
'        If .StorageType <> "BT" Then txtBottle.Text = ""
'        'ถุงพลาสติก
'        chkPlasticBag = IIf(.StorageType = "PT", 1, 0)
'        txtPlasticBag = .StorageDesc
'        If .StorageType <> "PT" Then txtPlasticBag.Text = ""
'        'ถังพลาสติก
'         chkPlasticDrum = IIf(.StorageType = "PD", 1, 0)
'         txtPlasticDrum = .StorageDesc
'         If .StorageType <> "PD" Then txtPlasticDrum.Text = ""
'         'กล่องกระดาษ
'         chkPaperBox = IIf(.StorageType = "PB", 1, 0)
'         txtPaperBox = .StorageDesc
'         If .StorageType <> "PB" Then txtPaperBox.Text = ""
        If .AnalysisType = 1 Then
            OptAnalysType1.Value = True
        ElseIf .AnalysisType = 2 Then
            OptAnalysType2.Value = True
        ElseIf .AnalysisType = 3 Then
            OptAnalysType3.Value = True
            txtAnalysDesc.Text = .AnalysisDesc
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
'    Debug.Print lvNewTreatment.ListItems(1).Height
End Sub
Public Sub SetDataContract(selContract As BWGContract, Optional isAssigned As Boolean = True)
    txtContractNo.Text = selContract.docNo
    If isAssigned Then
        Set myWasteDataCR_.Contract = selContract
    End If
End Sub
Public Sub SetDate(selDate As String, StDate As String)
    myWasteDataCR_.LicenseExpDate = selDate
'    txtExpDate.Text = FormatYMD_to_DMY(selDate, "/", "/")
End Sub
Public Sub SetDataIndustryType(selIndustry As BWGIndustryType_New, Optional isAssigned As Boolean = True)
    Set curIndustryType_ = selIndustry
    With curIndustryType_
        txtIndustryType.Text = .TypeDesc
        If isAssigned Then
            myWasteDataCR_.IndustryTypeID = .TypeID
            myWasteDataCR_.IndustryTypeName = .TypeDesc
            myWasteDataCR_.IndustryTypeIDNew = .TypeID
            myWasteDataCR_.TypeDesc = .TypeDesc
        End If
    End With
End Sub
Private Sub SetUserGroup()
Dim cVal As Boolean
Dim UserGrp As Collection
Dim Group As New AWSUserGroup
Dim i As Integer
    isEnv = False: isIT = False: isLab = False
    Set UserGrp = CurrentUser.getUserGroups
    For i = 1 To UserGrp.Count
        Set Group = UserGrp(i)
        If Group.Name = "Administrator" Then
            StrGrp = Group.Name
            isIT = True: isEnv = False
            Exit For
        ElseIf LCase(Group.Name) = "environment" Then
            StrGrp = Group.Name
            isEnv = True
            Exit For
        ElseIf LCase(Group.Name) = "Lab" Or LCase(Group.Name) = "Senior LAB" Then
            StrGrp = Group.Name
            isLab = True
        End If
    Next
    If isEnv = True Then
        SSTab1.TabVisible(9) = True
        SSTab1.TabEnabled(2) = True
        cboLABRes.Enabled = True
        cmdApprove.Caption = "ตรวจรับโดย Envi"
        cmdApprove.Visible = True
        cmdReject.Visible = True
        lvNewTreatment.Enabled = True
        txtEnviStaff.Enabled = True
        cboSummary.Enabled = True
        cmdRev.Visible = False
        cmdRev.Caption = "Rev. Waste"
        cmdRev.Left = cmdApprove.Left
        If isAkp = True Then
            SSTab1.TabEnabled(5) = True
        Else
            SSTab1.TabEnabled(5) = False
        End If
    Else
        SSTab1.TabVisible(9) = False
        cboLABRes.Enabled = False
        cboWasteType.Enabled = True
        cmdApprove.Caption = "ส่งให้สิ่งแวดล้อม"
        cmdReject.Visible = False
        lvNewTreatment.Enabled = False
        txtEnviStaff.Enabled = False
        cboSummary.Enabled = False
'        cboWasteType.Enabled = False
        'cmdRev.Visible = True
        cmdRev.Caption = "ต้องการ Rev. Waste "
        cmdRev.Left = cmdApprove.Left
'        ptChlorineY.Enabled = False
'        ptChlorineN.Enabled = False
'        ptCyanideY.Enabled = False
'        ptCyanideN.Enabled = False
'        ptNitriteY.Enabled = False
'        ptNitriteN.Enabled = False
        ChkCl.Enabled = False
        ChkCn.Enabled = False
        ChkNt.Enabled = False
        chkFlame.Enabled = False
        txtFlashPoint.Enabled = False
        txtviscosity.Enabled = False
        txtBulkdensity.Enabled = False
        chkCorrosive.Enabled = False
        txtpH.Enabled = False
        txtmoisturecontent.Enabled = False
        txtWatercontent.Enabled = False
        txtAshcontent.Enabled = False
        txtPhysicalstate.Enabled = False
        txtCL.Enabled = False
        textF.Enabled = False
        txtS.Enabled = False
        txtSludgecontent.Enabled = False
        txtsolidcontent.Enabled = False
        txtHeatingValue.Enabled = False
        txtAs.Enabled = False
        txtCd.Enabled = False
        txtCr.Enabled = False
        txtHg.Enabled = False
        txtPb.Enabled = False
        txtAcomment.Enabled = False
        cboAnalysis.Enabled = False
        Command2.Enabled = False
        Command4.Enabled = False
    End If
    cmdApprove.Picture = frmPicture.picApprove.Picture
End Sub
'Private Function SaveAttPict(ByVal CustomerID As String, ByVal wastedataID As String) As Boolean
'On Error GoTo ErrD
'Dim i As Integer
'Dim ServerPic As String
'Dim ServerPic1 As String
'Dim SourcePic As String
'Dim SourcePic1 As String
'Dim objNetwork
'Dim oNet As Object, MapDrv As Boolean
'Dim oFS As Scripting.FileSystemObject
'Dim Drivesys As New FileSystemObject
'    SaveAttPict = True
'
'        PicPath = "\\192.168.1.5\PicAtt\"
'        PicPath1 = "\\172.25.1.103\PicAtt\"
'        ServerPic = PicPath & CustomerID & "\"
'        ServerPic1 = PicPath1 & CustomerID & "\"
'         '======= Check Path ===================
'        If Dir(ServerPic, vbDirectory) = "" Then
'              MkDir ServerPic
'        End If
'        If Dir(ServerPic1, vbDirectory) = "" Then
'              MkDir ServerPic1
'        End If
'        ServerPic = ServerPic & wastedataID & "\"
'        ServerPic1 = ServerPic1 & wastedataID & "\"
'        If Dir(ServerPic, vbDirectory) = "" Then
'              MkDir ServerPic
'        End If
'        If Dir(ServerPic1, vbDirectory) = "" Then
'              MkDir ServerPic1
'        End If
'        Set oFS = New Scripting.FileSystemObject
'        For i = 0 To 1
'            If AttachesPic(i).FileName <> "" Then
'                oFS.CopyFile AttachesPic(i).FullPath, ServerPic & AttachesPic(i).FileName, True  ', "Administrator", "bwgadmin"
'                oFS.CopyFile AttachesPic(i).FullPath, ServerPic1 & AttachesPic(i).FileName, True  ', "Administrator", "bwgadmin"
'                Dim CanSee As String
'                If Check1(i).Value = 1 Then
'                    CanSee = "Y"
'                Else
'                    CanSee = "N"
'                End If
'                If lblAttPic(i).Caption <> "" Then
'                    Dim tSql As String
'                    Dim RS As ADODB.Recordset
'                    tSql = "SELECT * FROM tbPicPath" & vbCrLf
'                    tSql = tSql & " WHERE path = '" & ServerPic & AttachesPic(i).FileName & "'"
'                    Set RS = New ADODB.Recordset
'                    Set RS = GetRS(tSql, adOpenForwardOnly, adLockReadOnly)
'                        If RS.RecordCount <= 0 Then
'                            DBConnExc "Insert into tbPicPath (wastedataID,CustomerID,Path,Cansee) Values ('" & wastedataID & "','" & CustomerID & "','" & ServerPic & AttachesPic(i).FileName & "','" & CanSee & "' )"
'                        Else
'                            DBConnExc "update tbPicPath set Path='" & ServerPic & AttachesPic(i).FileName & "', cansee='" & CanSee & "' where path='" & ServerPic & AttachesPic(i).FileName & " ','" & CanSee & "' "
'                        End If
'                End If
'            End If
'        Next
'        AttachesPic(0).FileName = ""
'        AttachesPic(1).FileName = ""
'        AttachesPic(2).FileName = ""
'    Exit Function
'ErrD:
'    MsgBox Err.Description
'    SaveAttPict = False
'    Resume Next
'End Function
Private Function SaveAttache(ByVal CustomerID As String, ByVal wastedataID As String) As Boolean
On Error GoTo ErrD
Dim i As Integer
Dim ServerPath As String
Dim ServerPath1 As String
Dim SourceFile As String
Dim SourceFile1 As String
Dim objNetwork
Dim oNet As Object, MapDrv As Boolean
Dim oFS As Scripting.FileSystemObject
Dim Drivesys As New FileSystemObject
    SaveAttache = True
'    If Drivesys.FolderExists("Z:") = False Then
'        DefaultPath = "\AttacheFiles"
'        MapDrv = MapDrive(DefaultPath)
'        If MapDrv = False Then SaveAttache = False: Exit Function
'        ServerPath = DefaultPath & "\" & CustomerID & "\"
'        If Dir(ServerPath, vbDirectory) = "" Then
'              MkDir ServerPath
'        End If
'        ServerPath = ServerPath & wastedataID & "\"
'        If Dir(ServerPath, vbDirectory) = "" Then
'              MkDir ServerPath
'        End If
'        Set oFS = New Scripting.FileSystemObject
'        For i = 0 To 3
'            If Attaches(i).FileName <> "" Then
'                oFS.CopyFile Attaches(i).FullPath, ServerPath & Attaches(i).FileName ', True ', "Administrator", "bwgadmin"
'    '            FileCopy Attaches(i).FullPath, ServerPath & Attaches(i).FileName
'            End If
'        Next
'        RemoveDrive "Z:"
'    Else
'        DefaultPath = "S:\AttacheFiles\"
        DefaultPath = "\\192.168.1.2\AttacheFiles\"
        DefaultPath1 = "\\172.25.1.103\AttacheFiles\"
        ServerPath = DefaultPath & CustomerID & "\"
        ServerPath1 = DefaultPath1 & CustomerID & "\"
         '======= Check Path ===================
        If Dir(ServerPath, vbDirectory) = "" Then
              MkDir ServerPath
        End If
        If Dir(ServerPath1, vbDirectory) = "" Then
              MkDir ServerPath1
        End If
        ServerPath = ServerPath & wastedataID & "\"
        ServerPath1 = ServerPath1 & wastedataID & "\"
        If Dir(ServerPath, vbDirectory) = "" Then
              MkDir ServerPath
        End If
        If Dir(ServerPath1, vbDirectory) = "" Then
              MkDir ServerPath1
        End If
        Set oFS = New Scripting.FileSystemObject
        For i = 0 To 3
            If Attaches(i).FileName <> "" Then
                'oFS.CopyFile Attaches(i).FullPath, ServerPath & Attaches(i).FileName, "Administrator", "akpadmin123;'"
                'oFS.CopyFile Attaches(i).FullPath, ServerPath1 & Attaches(i).FileName, "Administrator", "akpadmin123;'"
                'oFS.CopyFile Attaches(i).FullPath, ServerPath & Attaches(i).FileName, "akpdomian\Administrator", "akpadminpassword"
                'oFS.CopyFile Attaches(i).FullPath, ServerPath1 & Attaches(i).FileName, "akpdomian\Administrator", "akpadminpassword"
                oFS.CopyFile Attaches(i).FullPath, ServerPath & Attaches(i).FileName
                oFS.CopyFile Attaches(i).FullPath, ServerPath1 & Attaches(i).FileName
                'FileCopy Attaches(i).FullPath, ServerPath & Attaches(i).FileName
            End If
        Next
'    End If
    Exit Function
ErrD:
    MsgBox Err.Description
    SaveAttache = False
    Resume Next
End Function

Private Function MapDrive(ByVal MapPath As String) As Boolean
On Error GoTo ErrD
Dim objNetwork
Dim strDriveLetter, strRemotePath, strDrv
Dim i, oFS, strUserName
Dim bForce, bUpdateProfile
Dim strUNCPrinter, strPrnDrv
    MapDrive = True
    strDriveLetter = "Z:"
    If Trim("MapPath") <> "" Then
        strRemotePath = IIf(Right(MapPath, 1) = "\", Mid(MapPath, 1, Len(MapPath) - 1), MapPath)
    Else
        strRemotePath = "\\192.168.1.2\AttacheFiles"
    End If
    Set objNetwork = CreateObject("WScript.Network")
    Set oFS = CreateObject("Scripting.FileSystemObject")
    If oFS.DriveExists(strDriveLetter) = True Then
        bForce = "True"
        bUpdateProfile = "True"
        Err.Number = vbEmpty
        objNetwork.RemoveNetworkDrive strDriveLetter, bForce, bUpdateProfile
    End If
    objNetwork.MapNetWorkDrive strDriveLetter, strRemotePath, True, "Administrator", "bwgadmin"
    Exit Function
ErrD:
        'MsgBox Err.Description
        MapDrive = False
End Function

Private Function RemoveDrive(ByVal strDrv As String) As Boolean
On Error GoTo ErrD
Dim objNetwork, oFS As Scripting.FileSystemObject
Dim bForce, bUpdateProfile
    Set oFS = New Scripting.FileSystemObject
    If oFS.DriveExists(strDrv) = True Then
        RemoveDrive = True
        bForce = "True"
        bUpdateProfile = "True"
        Err.Number = vbEmpty
        Set objNetwork = CreateObject("WScript.Network")
        objNetwork.RemoveNetworkDrive strDrv, bForce, bUpdateProfile
    End If
    Exit Function
ErrD:
    RemoveDrive = False
End Function
Private Sub cmdOpenFile_Click(Index As Integer)
Dim OpenFile As Double
Dim oFS As Scripting.FileSystemObject
Dim FSO As New FileSystemObject
Dim oNet As Object, MapDrv As Boolean
'    If FSO.FolderExists("Z:") = False Then
'DefaultPath = "S:\AttacheFiles"
        DefaultPath = "\\192.168.1.2\attachefiles\"
'        MapDrv = MapDrive(DefaultPath)
'    End If
    If lblAttach(Index).Caption <> "" Then
       ShellExecute 0, vbNullString, Attaches(Index).FullPath, vbNullString, vbNullString, vbNormalFocus
    End If
End Sub
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
'Private Function KillPic(ByVal Index As Integer, ByVal CustomerID As String, ByVal wastedataID As String) As Boolean
'On Error GoTo ErrD
'Dim FileName, FileName1 As String
'    KillPic = True
'    FileName = PicPath & CustomerID & "\" & wastedataID & "\" & lblAttPic(Index).Caption
'    FileName1 = PicPath1 & CustomerID & "\" & wastedataID & "\" & lblAttPic(Index).Caption
'    DBConnExc "delete  from tbPicPath where path='" & FileName & "' "
'    If Dir(FileName, vbNormal) <> "" Then
'        Kill FileName
'        Kill FileName1
'        lblAttPic(Index).Caption = ""
'    End If
'    Exit Function
'ErrD:
'    KillPic = False
'End Function
Private Sub LoadAttacheFile(ByVal CustomerID As String, ByVal wastedataID As String)
On Error GoTo ErrD
Dim FSO As New FileSystemObject
Dim oNet As Object, MapDrv As Boolean
Dim Filex As File
Dim FileObject As Files
Dim FolderObject As Folder
Dim ServerFilePath As String
Dim ServerPicPath As String
Dim i As Integer, j As Integer, x As Integer, Files As String
Dim Extension As String
'    If FSO.FolderExists("Z:\") = False Then
        DefaultPath = "\\192.168.1.2\attachefiles\"
        PicPath = "\\192.168.1.5\PicAtt\"
'        MapDrv = MapDrive(DefaultPath)
'        DefaultPath = "Z:\"
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
                    cmdAttach(x).Enabled = False
                    x = x + 1
                End If
                If x > 3 Then Exit For
            Next Filex
        End If
'        RemoveDrive "Z:"
'        ServerPicPath = PicPath & CustomerID & "\" & wastedataID
'        If FSO.FolderExists(ServerPicPath) = True Then
'            Set FolderObject = FSO.GetFolder(ServerPicPath)
'            Set FileObject = FolderObject.Files
'            i = FileObject.Count
'            x = 0
'            'ReDim AttachesPic(2)
'            For Each Filex In FileObject
'                For j = Len(Filex.Name) To 1 Step -1
'                    Extension = Right(Filex.Name, Len(Filex.Name) - j)
'                    If Mid(Filex.Name, j, 1) = "." Then Exit For
'                Next
'                If (LCase(Extension) <> "db") And (LCase(Extension) <> "dll") And (LCase(Extension) <> "exe") And (LCase(Extension) <> "com") And (LCase(Extension) <> "bat") Then
'                    AttachesPic(x).FileName = Filex.Name
'                    AttachesPic(x).FullPath = ServerPicPath & "\" & Filex.Name
'                    lblAttPic(x).Caption = Filex.Name
'                    CmdAttPic(x).Enabled = False
'                    Dim tSql As String
'                    Dim RS As ADODB.Recordset
'                    tSql = "SELECT * FROM tbPicPath" & vbCrLf
'                    tSql = tSql & " WHERE path = '" & AttachesPic(x).FullPath & "'"
'                    Set RS = New ADODB.Recordset
'                    Set RS = GetRS(tSql, adOpenForwardOnly, adLockReadOnly)
'                        If RS.RecordCount > 0 Then
'                            If RS!CanSee = "Y" Then
'                                Check1(x).Value = 1
'                            Else
'                                Check1(x).Value = 0
'                            End If
'                        End If
'                    x = x + 1
'                End If
'                If x > 1 Then Exit For
'            Next Filex
'        End If
'        Exit Sub
'    End If
'    DefaultPath = "Z:\AttacheFiles\"
'    ServerFilePath = DefaultPath & CustomerID & "\" & wastedataID
'    If FSO.FolderExists(ServerFilePath) = True Then
'        Set FolderObject = FSO.GetFolder(ServerFilePath)
'        Set FileObject = FolderObject.Files
'        i = FileObject.Count
'        x = 0
'        ReDim Attaches(4)
'        For Each Filex In FileObject
'            For j = Len(Filex.Name) To 1 Step -1
'                Extension = Right(Filex.Name, Len(Filex.Name) - j)
'                If Mid(Filex.Name, j, 1) = "." Then Exit For
'            Next
'            If (LCase(Extension) <> "db") And (LCase(Extension) <> "dll") And (LCase(Extension) <> "exe") And (LCase(Extension) <> "com") And (LCase(Extension) <> "bat") Then
'                Attaches(x).FileName = Filex.Name
'                Attaches(x).FullPath = ServerFilePath & "\" & Filex.Name
'                lblAttach(x).Caption = Filex.Name
'                x = x + 1
'            End If
'            If x > 3 Then Exit For
'        Next Filex
'    End If
    Exit Sub
ErrD:
    MsgBox Err.Description
End Sub

Private Sub txtWatercontent_Change()
    myWasteDataLab_.WaterContent = Trim(txtWatercontent.Text)
    myWasteDataLab_.isChanged = True
End Sub

Private Sub txtZn_Change()
    myWasteDataLab_.Zn = Trim(txtZn.Text)
    myWasteDataLab_.isChanged = True
End Sub
