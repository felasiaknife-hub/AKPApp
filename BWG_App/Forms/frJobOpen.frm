VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmJobOpen 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "F042"
   ClientHeight    =   10245
   ClientLeft      =   60
   ClientTop       =   2580
   ClientWidth     =   16995
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   10245
   ScaleWidth      =   16995
   Begin VB.CheckBox ChkCovidWaste 
      BackColor       =   &H000000FF&
      Caption         =   "ขยะโควิด"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1800
      TabIndex        =   112
      Top             =   9690
      Width           =   1230
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00C0C0FF&
      Caption         =   "ภาชนะ"
      Height          =   2805
      Left            =   12015
      TabIndex        =   83
      Top             =   3825
      Width           =   5000
      Begin VB.TextBox txtRack 
         Height          =   285
         Left            =   3450
         TabIndex        =   119
         Top             =   2430
         Width           =   510
      End
      Begin VB.TextBox CovidBin200 
         Height          =   285
         Left            =   3450
         TabIndex        =   118
         Top             =   2115
         Width           =   510
      End
      Begin VB.TextBox CovidBin 
         Height          =   285
         Left            =   3450
         TabIndex        =   110
         Top             =   1800
         Width           =   510
      End
      Begin VB.TextBox TxtBoutCage 
         Height          =   285
         Left            =   3450
         TabIndex        =   109
         Top             =   1485
         Width           =   510
      End
      Begin VB.TextBox TxtPaperBox 
         Height          =   285
         Left            =   3450
         TabIndex        =   107
         Top             =   1170
         Width           =   510
      End
      Begin VB.TextBox TxtBagBack 
         Height          =   285
         Left            =   3450
         TabIndex        =   105
         Top             =   855
         Width           =   510
      End
      Begin VB.TextBox TxtBagTypeP 
         Height          =   285
         Left            =   3450
         TabIndex        =   103
         Top             =   540
         Width           =   510
      End
      Begin VB.TextBox TxtTankTypeS 
         Height          =   285
         Left            =   3450
         TabIndex        =   101
         Top             =   225
         Width           =   510
      End
      Begin VB.TextBox TxtTankTypeL 
         Height          =   285
         Left            =   1550
         TabIndex        =   99
         Top             =   2430
         Width           =   510
      End
      Begin VB.TextBox TxtPlastic 
         Height          =   285
         Left            =   1550
         TabIndex        =   97
         Top             =   2115
         Width           =   510
      End
      Begin VB.TextBox txtPlate 
         Height          =   285
         Left            =   1550
         TabIndex        =   95
         Top             =   1800
         Width           =   510
      End
      Begin VB.TextBox TxtLackGas 
         Height          =   285
         Left            =   1550
         TabIndex        =   93
         Top             =   1485
         Width           =   510
      End
      Begin VB.TextBox TxtLackNew 
         Height          =   285
         Left            =   1550
         TabIndex        =   91
         Top             =   1170
         Width           =   510
      End
      Begin VB.TextBox TxtLack 
         Height          =   285
         Left            =   1550
         TabIndex        =   89
         Top             =   855
         Width           =   510
      End
      Begin VB.TextBox TxtBoutS 
         Height          =   285
         Left            =   1550
         TabIndex        =   87
         Top             =   540
         Width           =   510
      End
      Begin VB.TextBox TxtContainer 
         Height          =   285
         Left            =   1550
         TabIndex        =   85
         Top             =   225
         Width           =   510
      End
      Begin VB.Label Label25 
         BackColor       =   &H00C0C0FF&
         Caption         =   "แร็ค"
         Height          =   255
         Left            =   2200
         TabIndex        =   120
         Top             =   2475
         Width           =   1095
      End
      Begin VB.Label Label24 
         BackColor       =   &H00C0C0FF&
         Caption         =   "ถังขยะโควิด 200"
         Height          =   195
         Left            =   2200
         TabIndex        =   117
         Top             =   2115
         Width           =   1150
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         Caption         =   "ถังขยะโควิด 150"
         Height          =   195
         Left            =   2200
         TabIndex        =   111
         Top             =   1845
         Width           =   1125
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         Caption         =   "กรงเบ้าท์"
         Height          =   195
         Left            =   2200
         TabIndex        =   108
         Top             =   1530
         Width           =   630
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         Caption         =   "กล่องกระดาษ"
         Height          =   195
         Left            =   2200
         TabIndex        =   106
         Top             =   1215
         Width           =   930
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         Caption         =   "ถุงดำ"
         Height          =   195
         Left            =   2200
         TabIndex        =   104
         Top             =   900
         Width           =   360
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         Caption         =   "BigBag"
         Height          =   195
         Left            =   2200
         TabIndex        =   102
         Top             =   585
         Width           =   510
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         Caption         =   "ถัง200 ลิตรปากจู๋"
         Height          =   195
         Left            =   2200
         TabIndex        =   100
         Top             =   270
         Width           =   1155
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         Caption         =   "ถัง200 ลิตรปากกว้าง"
         Height          =   195
         Left            =   90
         TabIndex        =   98
         Top             =   2475
         Width           =   1410
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         Caption         =   "พลาสติกแล็ป"
         Height          =   195
         Left            =   90
         TabIndex        =   96
         Top             =   2160
         Width           =   900
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         Caption         =   "พาเลท"
         Height          =   195
         Left            =   90
         TabIndex        =   94
         Top             =   1845
         Width           =   465
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         Caption         =   "แลคใส่ถังแก๊ส"
         Height          =   195
         Left            =   90
         TabIndex        =   92
         Top             =   1530
         Width           =   990
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         Caption         =   "แลครุ่นใหม่"
         Height          =   195
         Left            =   90
         TabIndex        =   90
         Top             =   1215
         Width           =   825
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         Caption         =   "แลครุ่นเก่า"
         Height          =   195
         Left            =   90
         TabIndex        =   88
         Top             =   900
         Width           =   780
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         Caption         =   "เบ้าท์ตัดปาก"
         Height          =   195
         Left            =   90
         TabIndex        =   86
         Top             =   585
         Width           =   855
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         Caption         =   "เบ้าท์ใส่น้ำเสีย"
         Height          =   195
         Left            =   90
         TabIndex        =   84
         Top             =   270
         Width           =   1005
      End
   End
   Begin VB.ComboBox CbRemark 
      ForeColor       =   &H00000080&
      Height          =   315
      ItemData        =   "frJobOpen.frx":0000
      Left            =   1800
      List            =   "frJobOpen.frx":0002
      TabIndex        =   66
      Text            =   "Combo1"
      Top             =   9330
      Width           =   3525
   End
   Begin VB.CheckBox Chk 
      BackColor       =   &H0000FF00&
      Caption         =   "ตรวจสอบข้อมูลการสั่งงานแล้ว"
      Height          =   285
      Left            =   12780
      TabIndex        =   58
      Top             =   405
      Width           =   2400
   End
   Begin VB.Frame Frame5 
      BorderStyle     =   0  'None
      Height          =   1035
      Left            =   0
      TabIndex        =   45
      Top             =   9225
      Width           =   1515
      Begin VB.CommandButton cmdAddItem 
         Caption         =   "เลือกกากของเสีย"
         Height          =   795
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   120
         Width           =   1365
      End
   End
   Begin VB.Frame Frame4 
      BorderStyle     =   0  'None
      Height          =   1035
      Left            =   11000
      TabIndex        =   44
      Top             =   9225
      Width           =   6015
      Begin VB.CommandButton cmdSave 
         Caption         =   "เก็บข้อมูล"
         Height          =   795
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   120
         Width           =   945
      End
      Begin VB.CommandButton cmdOpen 
         Caption         =   "เปิดข้อมูล"
         Height          =   795
         Left            =   1020
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   120
         Width           =   945
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "สร้างใหม่"
         Height          =   795
         Left            =   2010
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   120
         Width           =   945
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "ปิด"
         Height          =   795
         Left            =   4980
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   120
         Width           =   945
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "ยกเลิก"
         Height          =   795
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   120
         Width           =   945
      End
      Begin VB.CommandButton cmdDone 
         Caption         =   "ปิด Job"
         Enabled         =   0   'False
         Height          =   795
         Left            =   3990
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   120
         Width           =   945
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H8000000E&
      Caption         =   "รถที่ใช้ในการขน"
      Height          =   2115
      Left            =   30
      TabIndex        =   38
      Top             =   7055
      Width           =   17000
      Begin WasteManagment.ctlGrid grTimeTable 
         Height          =   1515
         Left            =   13500
         TabIndex        =   12
         Top             =   540
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   2672
      End
      Begin VSFlex7LCtl.VSFlexGrid grCar 
         Height          =   1680
         Left            =   90
         TabIndex        =   46
         ToolTipText     =   "ถ้าต้องการแก้ไขข้อความ ให้เลือกเซลล์แล้วกด F2 เพื่อแก้ไข"
         Top             =   270
         Width           =   13380
         _cx             =   23601
         _cy             =   2963
         _ConvInfo       =   1
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   8421504
         BackColorAlternate=   -2147483643
         GridColor       =   12632256
         GridColorFixed  =   0
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   0
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   12
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "ตารางการเดินรถ"
         Height          =   285
         Left            =   13020
         TabIndex        =   39
         Top             =   285
         Width           =   2295
      End
      Begin VB.Shape Shape1 
         Height          =   1845
         Left            =   60
         Top             =   210
         Width           =   15990
      End
   End
   Begin VB.ComboBox cboCompany 
      Height          =   315
      Left            =   7830
      Style           =   2  'Dropdown List
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   360
      Width           =   3525
   End
   Begin Crystal.CrystalReport rpt1 
      Left            =   3825
      Top             =   -45
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H8000000E&
      Caption         =   "รายการกากของเสีย"
      Height          =   2085
      Left            =   45
      TabIndex        =   23
      Top             =   5000
      Width           =   11940
      Begin WasteManagment.ctlGrid grWasteData 
         Height          =   1800
         Left            =   45
         TabIndex        =   11
         Top             =   225
         Width           =   11835
         _ExtentX        =   20876
         _ExtentY        =   3175
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000E&
      Caption         =   "รายละเอียดใบเสนอราคา"
      Height          =   4235
      Left            =   45
      TabIndex        =   25
      Top             =   765
      Width           =   17000
      Begin VB.TextBox txtCostTrans 
         Height          =   315
         Left            =   4560
         TabIndex        =   127
         Top             =   2520
         Width           =   2175
      End
      Begin VB.ComboBox cboTruckType 
         Height          =   315
         Left            =   1485
         TabIndex        =   125
         Text            =   "Combo1"
         Top             =   2520
         Width           =   1935
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H8000000E&
         Caption         =   "ค่าคอม BME"
         Height          =   475
         Left            =   11760
         TabIndex        =   121
         Top             =   140
         Width           =   2295
         Begin VB.OptionButton rdoComBME 
            BackColor       =   &H8000000E&
            Caption         =   "ไม่คิด"
            Height          =   255
            Index           =   1
            Left            =   960
            TabIndex        =   123
            Top             =   200
            Width           =   1215
         End
         Begin VB.OptionButton rdoComBME 
            BackColor       =   &H8000000E&
            Caption         =   "คิด"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   122
            Top             =   200
            Value           =   -1  'True
            Width           =   600
         End
      End
      Begin VB.ComboBox cbProject 
         Height          =   315
         Left            =   9810
         TabIndex        =   114
         Top             =   2160
         Width           =   2085
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00C0FFC0&
         Caption         =   "ถ่ายรูป"
         Height          =   2175
         Left            =   11970
         TabIndex        =   72
         Top             =   855
         Width           =   5000
         Begin VB.CheckBox chkPhotoS 
            BackColor       =   &H00C0FFC0&
            Caption         =   "ถ่ายรูปพิเศษ"
            Height          =   240
            Left            =   1485
            TabIndex        =   116
            Top             =   315
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.CheckBox chkS 
            BackColor       =   &H00C0FFC0&
            Caption         =   "ถ่ายรูปพิเศษ"
            Height          =   240
            Left            =   2970
            TabIndex        =   115
            Top             =   1170
            Visible         =   0   'False
            Width           =   1545
         End
         Begin VB.CheckBox ChkNewCustJobBI 
            BackColor       =   &H00C0FFC0&
            Caption         =   "งานลูกค้าใหม่ BI"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   90
            TabIndex        =   82
            Top             =   1845
            Width           =   1725
         End
         Begin VB.CheckBox ChkW 
            BackColor       =   &H00C0FFC0&
            Caption         =   "ถ่ายรูปชั่ง + ลงของ"
            Height          =   240
            Left            =   2970
            TabIndex        =   81
            Top             =   630
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox ChkB 
            BackColor       =   &H00C0FFC0&
            Caption         =   "ถ่ายรูปทำลาย"
            Height          =   240
            Left            =   2970
            TabIndex        =   80
            Top             =   900
            Visible         =   0   'False
            Width           =   1545
         End
         Begin VB.CheckBox ChkNewCustJob 
            BackColor       =   &H00C0FFC0&
            Caption         =   "งานลูกค้าใหม่"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   90
            TabIndex        =   79
            Top             =   270
            Width           =   1725
         End
         Begin VB.CheckBox ChkC 
            BackColor       =   &H00C0FFC0&
            Caption         =   "เก็บภาชนะไว้คืนลูกค้า"
            Height          =   240
            Left            =   90
            TabIndex        =   78
            Top             =   1530
            Width           =   1905
         End
         Begin VB.CheckBox ChkCus 
            BackColor       =   &H00C0FFC0&
            Caption         =   "ลูกค้าตามดูทำลาย"
            Height          =   240
            Left            =   90
            TabIndex        =   77
            Top             =   1215
            Width           =   1725
         End
         Begin VB.CheckBox ChkCloseWeight 
            BackColor       =   &H00C0FFC0&
            Caption         =   "ปิดน้ำหนักลูกค้า"
            Height          =   240
            Left            =   11475
            TabIndex        =   76
            Top             =   0
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.ComboBox CboC 
            Height          =   315
            ItemData        =   "frJobOpen.frx":0004
            Left            =   2025
            List            =   "frJobOpen.frx":0006
            TabIndex        =   75
            Top             =   1485
            Width           =   1500
         End
         Begin VB.CheckBox ChkEm 
            BackColor       =   &H00C0FFC0&
            Caption         =   "อบรมพนักงาน"
            Height          =   240
            Left            =   90
            TabIndex        =   74
            Top             =   900
            Width           =   1815
         End
         Begin VB.CheckBox ChkCarCheck 
            BackColor       =   &H00C0FFC0&
            Caption         =   "ตรวจสภาพรถ"
            Height          =   240
            Left            =   90
            TabIndex        =   73
            Top             =   585
            Width           =   1455
         End
      End
      Begin VB.ComboBox cboConfirm 
         Height          =   315
         Left            =   9360
         TabIndex        =   67
         Text            =   "Combo1"
         Top             =   1785
         Width           =   2265
      End
      Begin VB.ComboBox cboCallWay 
         Height          =   315
         ItemData        =   "frJobOpen.frx":0008
         Left            =   7380
         List            =   "frJobOpen.frx":001B
         TabIndex        =   64
         Top             =   2160
         Width           =   1680
      End
      Begin VB.TextBox txtCallName 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   4680
         TabIndex        =   63
         Top             =   2160
         Width           =   2055
      End
      Begin VB.ComboBox cboCaller 
         Height          =   315
         ItemData        =   "frJobOpen.frx":0054
         Left            =   3465
         List            =   "frJobOpen.frx":0064
         TabIndex        =   62
         Top             =   2160
         Width           =   1230
      End
      Begin VB.ComboBox CbOrdCompany 
         Height          =   315
         ItemData        =   "frJobOpen.frx":008D
         Left            =   1485
         List            =   "frJobOpen.frx":008F
         TabIndex        =   60
         Text            =   "Combo1"
         Top             =   2160
         Width           =   1455
      End
      Begin VB.TextBox TxtRefME 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   14240
         TabIndex        =   55
         Top             =   540
         Width           =   2235
      End
      Begin VB.ComboBox CboType 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "frJobOpen.frx":0091
         Left            =   7095
         List            =   "frJobOpen.frx":00AD
         TabIndex        =   53
         Text            =   "--"
         Top             =   1785
         Width           =   1440
      End
      Begin VB.TextBox txtPONo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4785
         TabIndex        =   49
         Top             =   1800
         Width           =   1320
      End
      Begin VB.ComboBox cboOrderType 
         Height          =   315
         ItemData        =   "frJobOpen.frx":0106
         Left            =   1470
         List            =   "frJobOpen.frx":0108
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1800
         Width           =   2130
      End
      Begin VB.ComboBox cboWasteGenNo 
         Height          =   315
         Left            =   8730
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1080
         Width           =   2265
      End
      Begin VB.ComboBox cboCustSite 
         Height          =   315
         Left            =   1470
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1080
         Width           =   5715
      End
      Begin VB.TextBox txtRem 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1485
         TabIndex        =   22
         Top             =   1440
         Width           =   5475
      End
      Begin VB.TextBox txtSaleStaff 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   9825
         TabIndex        =   6
         Top             =   660
         Width           =   1515
      End
      Begin VB.TextBox txtM 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   11760
         MaxLength       =   2
         TabIndex        =   21
         Top             =   720
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.TextBox txtH 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   11310
         MaxLength       =   2
         TabIndex        =   20
         Top             =   720
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.TextBox txtIssuedDate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C00000&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   3690
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   300
         Width           =   1425
      End
      Begin VB.TextBox txtIssuedStaff 
         Appearance      =   0  'Flat
         BackColor       =   &H00C00000&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   8190
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   1440
         Width           =   2115
      End
      Begin VB.TextBox txtCustomer 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1470
         TabIndex        =   4
         Top             =   690
         Width           =   4125
      End
      Begin VB.TextBox txtJobOpenNo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C00000&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1470
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Text            =   "== AUTO =="
         Top             =   300
         Width           =   1755
      End
      Begin WasteManagment.ctlDate dtWorkDate 
         Height          =   315
         Left            =   6285
         TabIndex        =   3
         Top             =   300
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   556
         BackColor       =   -2147483633
      End
      Begin MSComctlLib.ListView lvResult 
         Height          =   1035
         Left            =   60
         TabIndex        =   10
         Top             =   3040
         Width           =   6960
         _ExtentX        =   12277
         _ExtentY        =   1826
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
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "หมายเหตุทั่วไปของลูกค้า"
            Object.Width           =   20814
         EndProperty
      End
      Begin WasteManagment.ctlDate dtFromWork 
         Height          =   315
         Left            =   6810
         TabIndex        =   52
         Top             =   720
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         BackColor       =   -2147483633
      End
      Begin VSFlex7LCtl.VSFlexGrid GrCon 
         Height          =   2645
         Left            =   -720
         TabIndex        =   57
         ToolTipText     =   "สำหรับแสดงข้อมูลเท่านั้น ไม่สามารถแก้ไขได้"
         Top             =   3010
         Visible         =   0   'False
         Width           =   3435
         _cx             =   6059
         _cy             =   4665
         _ConvInfo       =   1
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   8421504
         BackColorAlternate=   -2147483643
         GridColor       =   12632256
         GridColorFixed  =   0
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   0
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   12
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "ราคาต้นทุน"
         Height          =   195
         Left            =   3600
         TabIndex        =   126
         Top             =   2540
         Width           =   800
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         Caption         =   "ประเภทรถ"
         Height          =   195
         Left            =   600
         TabIndex        =   124
         Top             =   2540
         Width           =   815
      End
      Begin VB.Label Label23 
         BackColor       =   &H8000000E&
         Caption         =   "โปรเจค"
         Height          =   195
         Left            =   9225
         TabIndex        =   113
         Top             =   2205
         Width           =   600
      End
      Begin VB.Label Obc2 
         Caption         =   "0"
         Height          =   195
         Left            =   8235
         TabIndex        =   71
         Top             =   3150
         Visible         =   0   'False
         Width           =   1410
      End
      Begin VB.Label Obc1 
         Caption         =   "0"
         Height          =   195
         Left            =   8325
         TabIndex        =   70
         Top             =   2790
         Visible         =   0   'False
         Width           =   1545
      End
      Begin VB.Label Label8 
         BackColor       =   &H8000000E&
         Caption         =   "Confirm"
         Height          =   195
         Left            =   8775
         TabIndex        =   69
         Top             =   1845
         Width           =   600
      End
      Begin VB.Label Label6 
         BackColor       =   &H8000000E&
         Caption         =   "ช่องทาง"
         Height          =   195
         Left            =   6795
         TabIndex        =   65
         Top             =   2205
         Width           =   600
      End
      Begin VB.Label Label5 
         BackColor       =   &H8000000E&
         Caption         =   "ผู้แจ้ง"
         Height          =   195
         Left            =   3060
         TabIndex        =   61
         Top             =   2205
         Width           =   465
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000E&
         Caption         =   "ต้นทางสั่งงาน"
         Height          =   240
         Left            =   405
         TabIndex        =   59
         Top             =   2205
         Width           =   1095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
         Caption         =   "อ้างอิงใบสั่งงาน Better ME"
         Height          =   375
         Index           =   14
         Left            =   14240
         TabIndex        =   56
         Top             =   315
         Width           =   2250
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ประเภทงาน"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   11
         Left            =   6255
         TabIndex        =   54
         Top             =   1830
         Width           =   825
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "วันที่ต้องการ"
         Height          =   195
         Index           =   2
         Left            =   5880
         TabIndex        =   51
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "สั่งงานตาม PO."
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   9
         Left            =   3705
         TabIndex        =   48
         Top             =   1845
         Width           =   1035
      End
      Begin VB.Label lblDTeam 
         Alignment       =   2  'Center
         BackColor       =   &H00FF00FF&
         Caption         =   "ลูกค้า DTeam"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   375
         Left            =   15040
         TabIndex        =   47
         Top             =   0
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ประเภทของ Order"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   7
         Left            =   165
         TabIndex        =   43
         Top             =   1890
         Width           =   1245
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Manifest ที่ออกได้/ที่ออกแล้ว"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   6
         Left            =   8310
         TabIndex        =   42
         Top             =   330
         Width           =   1995
         WordWrap        =   -1  'True
      End
      Begin VB.Label lbMenifestCount 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   10470
         TabIndex        =   41
         Top             =   360
         Width           =   1035
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00FF0000&
         Height          =   375
         Left            =   10410
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "เลขทะเบียนโรงงาน"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   5
         Left            =   7290
         TabIndex        =   40
         Top             =   1080
         Width           =   1350
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "หมายเหตุ"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   750
         TabIndex        =   37
         Top             =   1500
         Width           =   675
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ผู้แจ้ง Order"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   13
         Left            =   8865
         TabIndex        =   36
         Top             =   750
         Width           =   825
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "สถานที่เก็บขน"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   12
         Left            =   420
         TabIndex        =   35
         Top             =   1140
         Width           =   990
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   10
         Left            =   11790
         TabIndex        =   34
         Top             =   780
         Visible         =   0   'False
         Width           =   180
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "เวลา"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   3
         Left            =   10650
         TabIndex        =   33
         Top             =   780
         Visible         =   0   'False
         Width           =   390
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "วันที่ปฎิบัติงาน"
         Height          =   195
         Index           =   0
         Left            =   5205
         TabIndex        =   32
         Top             =   360
         Width           =   1005
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ผู้ออกใบสั่งงาน"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   7110
         TabIndex        =   29
         Top             =   1500
         Width           =   1035
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ลูกค้า"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   8
         Left            =   1020
         TabIndex        =   28
         Top             =   750
         Width           =   375
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "เลขที่ใบสั่งงาน"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   390
         TabIndex        =   27
         Top             =   360
         Width           =   990
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "วันที่"
         Height          =   195
         Index           =   8
         Left            =   3330
         TabIndex        =   26
         Top             =   360
         Width           =   315
      End
   End
   Begin VB.Label Label28 
      Caption         =   "Label28"
      Height          =   285
      Left            =   15240
      TabIndex        =   128
      Top             =   405
      Width           =   1695
   End
   Begin VB.Label Label7 
      BackColor       =   &H8000000E&
      Caption         =   "ช่องทาง"
      Height          =   195
      Left            =   9270
      TabIndex        =   68
      Top             =   2610
      Width           =   600
   End
   Begin VB.Label LbNumber 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   645
      Left            =   6360
      TabIndex        =   50
      Top             =   30
      Width           =   855
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
      Left            =   4680
      TabIndex        =   31
      Top             =   30
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "รายการรออนุมัติ, หรืออนุมัติแล้วเมื่อวันที่..... โดย...."
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Index           =   4
      Left            =   7830
      TabIndex        =   30
      Top             =   0
      Visible         =   0   'False
      Width           =   3555
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   4
      Index           =   0
      X1              =   90
      X2              =   7560
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "เปิดใบสั่งงานเก็บขนกากของเสีย"
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
      TabIndex        =   24
      Top             =   90
      Width           =   2730
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      Index           =   1
      X1              =   180
      X2              =   7650
      Y1              =   555
      Y2              =   555
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   13105
      Y1              =   8640
      Y2              =   8640
   End
   Begin VB.Image Image2 
      Height          =   705
      Left            =   0
      Picture         =   "frJobOpen.frx":010A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15720
   End
End
Attribute VB_Name = "frmJobOpen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim myJobOpen_ As BWGJobData
Dim curCustomer_ As BWGCustomer
Dim isDetailChange_ As Boolean
Dim curCar_ As BWGJobDataCarUsed
Dim Index As Integer
Dim colSearchResult_ As Collection

Dim TypeOfBox As String
Dim BoxAKP As String
Const TypeOfCar = "#1;พร้อม|#2;ไม่พร้อม"
Const TypeOfPhotography = "#1;ถ่ายรูปทำลาย|#2;ไม่ถ่าย|#3;ถ่ายรูปพิเศษ"
Const TypeTank = "#1;ปากกว้าง|#2;ปากแคบ"
Const TypeBag = "#1;ปากป่อง|#2;เปิดปาก"
Dim AKPBoxCode As String
Dim AKPBoxCode2 As String
Dim Frame2H As Double
Dim Frame3H As Double
Public gVal As String
Dim colDistrict_ As Collection

Private Sub LoadBoxType(consql As String)
    Dim rs As New ADODB.Recordset
    Dim tSql As String, i As Integer
    tSql = "SELECT     TOP (100) PERCENT BoxType, BoxTypeDesc, SUM(BoxStock) AS BoxStock" & vbCrLf
    tSql = tSql & "From dbo.vw_BoxQty" & vbCrLf
    tSql = tSql & "WHERE (BoxStatus = 1)" & vbCrLf
      tSql = tSql & consql & vbCrLf
    tSql = tSql & "GROUP BY BoxType, BoxTypeDesc"
    Set rs = GetRS(tSql, adOpenForwardOnly, adLockOptimistic, DBConStock)
    i = 0
    Do While Not rs.EOF
        i = i + 1
        If i = 1 Then
            TypeOfBox = "#" & rs!BoxType & "*" & i & ";" & rs!BoxTypeDesc & vbTab & rs!BoxType & vbTab & rs!BoxStock
        Else
            TypeOfBox = TypeOfBox & "|#" & rs!BoxType & ";" & rs!BoxTypeDesc & vbTab & rs!BoxType & vbTab & rs!BoxStock
        End If
        rs.MoveNext
    Loop
    Set rs = Nothing
End Sub
Private Sub LoadBoxAKP(consql As String)
    Dim rs As New ADODB.Recordset
    Dim tSql As String, i As Integer
    tSql = "SELECT * From dbo.tbAKPBox" & vbCrLf
    tSql = tSql & "order BY Boxcode "
    Set rs = GetRS(tSql, adOpenForwardOnly, adLockOptimistic, DBConString)
    i = 0
    Do While Not rs.EOF
        i = i + 1
        If i = 1 Then
            BoxAKP = "#" & i & ";" & rs!boxname
        Else
            BoxAKP = BoxAKP & "|" & "#" & i & ";" & rs!boxname
        End If
        rs.MoveNext
    Loop
    Set rs = Nothing
End Sub

Private Function BoxTypeCount(ByVal BoxType As String) As Integer
    Dim rs As New ADODB.Recordset
    Dim tSql As String, i As Integer
    Dim Reserve As Integer, BoxCount As Integer
    BoxTypeCount = 0
    tSql = "SELECT SUM(BoxStock) AS BoxStock" & vbCrLf
    tSql = tSql & "FROM dbo.vw_BoxQty" & vbCrLf
    tSql = tSql & "WHERE (BoxStatus = 1) AND BoxType = '" & BoxType & "'" & vbCrLf
    tSql = tSql & "GROUP BY BoxType"
    Set rs = GetRS(tSql, adOpenForwardOnly, adLockOptimistic, DBConStock)
    If Not rs.EOF Then
        BoxCount = rs!BoxStock
    Else
        BoxCount = 0
    End If
    Set rs = Nothing
    
    tSql = "SELECT SUM(BoxReservNumber) AS Reserve" & vbCrLf
    tSql = tSql & "From BWG_BOX.dbo.Box_BoxReserve" & vbCrLf
    tSql = tSql & "WHERE BoxType = '" & BoxType & "'" & vbCrLf
    tSql = tSql & "GROUP BY BoxType"
    Set rs = GetRS(tSql, adOpenForwardOnly, adLockOptimistic, DBConStock)
    If Not rs.EOF Then
        Reserve = rs!Reserve
    Else
        Reserve = 0
    End If
    
    BoxTypeCount = BoxCount - Reserve
End Function

Private Sub CboC_Validate(Cancel As Boolean)
Dim tmpDet As BWGJobDataCarUsed
Dim tmpKey As String, BoxCheck As Integer
Dim tmpVal As String, OldValue As String
    tmpKey = Trim(grCar.TextMatrix(1, 0))
    If isExist(myJobOpen_.Cars, "'" & tmpKey & "'") = False Then
        Set tmpDet = New BWGJobDataCarUsed
        tmpDet.id = GetGUID
        tmpDet.JobID = myJobOpen_.id
        grCar.TextMatrix(1, 0) = tmpDet.id
    Else
        Set tmpDet = myJobOpen_.Cars("'" & tmpKey & "'")
    End If
    tmpVal = Trim(grCar.TextMatrix(1, 0))
    tmpDet.RowID = 1
    With tmpDet
        If ChkC.Value = 1 Then
            .ConName = Trim(CboC.Text)
        Else
            .ConName = ""
        End If
                myJobOpen_.isChanged = True
    End With
End Sub
Private Sub cboCaller_Validate(Cancel As Boolean)
    If cboCaller.ListIndex <> 0 Then
        myJobOpen_.CallOrd = cboCaller.Text
    End If
End Sub
Private Sub cboCallWay_Validate(Cancel As Boolean)
    If cboCallWay.ListIndex <> 0 Then
        myJobOpen_.CallerWay = cboCallWay.Text
    End If
End Sub

Private Sub cboCompany_Click()
If cboCompany.ListIndex > 1 Then
    Label4.Caption = "ต้นทางสั่งงาน"
Else
    Label4.Caption = "ปลายทางกำจัด"
End If
End Sub

Private Sub cboCompany_Validate(Cancel As Boolean)
Dim tmpCol As Collection
    Set tmpCol = myJobOpen_.JobDetails
    If Not tmpCol Is Nothing Then
        If tmpCol.Count > 0 Then
            MsgBox "ไม่สามารถเปลี่ยนบริษัทได้ เพราะมีการเลือกเวสต์ดาต้าแล้ว", vbExclamation, "Waring"
            cboCompany.ListIndex = Index
            Exit Sub
        End If
    End If
    
    If cboCompany.ListIndex >= 0 Then
        If cboCompany.ListCount > 1 Then
            If cboCompany.ListIndex = 0 Then Exit Sub
            myJobOpen_.CompanyID = CurrentUser.WorkCompanies(cboCompany.ListIndex).id
        Else
            myJobOpen_.CompanyID = CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).id
        End If
        Index = cboCompany.ListIndex
    Else
        myJobOpen_.CompanyID = ""
    End If
End Sub



Private Sub cboConfirm_LostFocus()
  ConfirmValidate
End Sub

Private Sub cboConfirm_Validate(Cancel As Boolean)
    ConfirmValidate
End Sub
Private Sub ConfirmValidate()
    Dim tSql, Sql, sqlStr As String
    Dim rs As New ADODB.Recordset
    Dim RSDet As New ADODB.Recordset
    Dim i As Integer
    Dim OrderCode As String
    Dim QuotationDetID As String
    Dim QuotationID As String
    OrderCode = ""
    OrderCode = Trim(Replace(Mid(cboConfirm.Text, 1, 12), "#", ""))
    tSql = "SELECT * FROM Service_OrderHD " & vbCrLf
    tSql = tSql & "where OrderNo='" & OrderCode & "' "
    Set rs = GetRS(tSql, adOpenForwardOnly, adLockOptimistic)
    If rs.RecordCount > 0 Then
        myJobOpen_.ServiceOrderID = rs!OrderID
        myJobOpen_.ServicePhoto = IIf(IsNull(rs!TakePhoto), "N", rs!TakePhoto)
        'myJobOpen_.ServicePhoto = rs!TakePhoto
        myJobOpen_.isChanged = True
    End If

End Sub
Private Sub cboCustSite_Validate(Cancel As Boolean)
    If curCustomer_ Is Nothing Then Set curCustomer_ = New BWGCustomer
    If cboCustSite.ListIndex > -1 Then
        myJobOpen_.CustomerSiteID = curCustomer_.SiteAddresses(cboCustSite.ListIndex + 1).id
        myJobOpen_.CustomerSiteName = cboCustSite.Text
    Else
        myJobOpen_.CustomerSiteID = ""
        myJobOpen_.CustomerSiteName = ""
    End If
End Sub

Private Sub cboOrderType_Validate(Cancel As Boolean)
    If cboOrderType.ListIndex = 0 Then
        myJobOpen_.isJob = "N"
        myJobOpen_.isCase = ""
    ElseIf cboOrderType.ListIndex = 1 Then
        myJobOpen_.isJob = "Y"
        myJobOpen_.isCase = ""
    ElseIf cboOrderType.ListIndex = 2 Then
        myJobOpen_.isJob = "D"
'        myJobOpen_.isCase = WasteCases(cboOrderType.ListIndex - 1).WasteCaseID
    ElseIf cboOrderType.ListIndex > 2 Then
        myJobOpen_.isJob = "N"
        myJobOpen_.isCase = WasteCases(cboOrderType.ListIndex - 1).WasteCaseID
    End If
    myJobOpen_.isChanged = True
End Sub

Private Sub CbOrdCompany_Validate(Cancel As Boolean)
    Dim tSql As String
    Dim rs As New ADODB.Recordset
    Dim i As Integer
    tSql = "SELECT * FROM vw_SaleDistrictSearch" & vbCrLf
    tSql = tSql & "where Zonedesc='" & CbOrdCompany.Text & "' and SaleStaffID = '" & curCustomer_.UnderSaleStaffID & "' "
    Set rs = GetRS(tSql, adOpenForwardOnly, adLockOptimistic)
    If rs.RecordCount > 0 Then
        If CbOrdCompany.ListIndex <> 0 Then
            myJobOpen_.SaleDistrictIDOrder = rs!SaleDistrictID
        End If
    End If
End Sub



Private Sub cboTruckType_Validate(Cancel As Boolean)
        If cboTruckType.ListIndex = 0 Then
            myJobOpen_.TruckType = "Y"
        Else
            myJobOpen_.TruckType = "N"
        End If
    myJobOpen_.isChanged = True
End Sub

Private Sub CboType_Validate(Cancel As Boolean)
    If cboType.ListIndex = 0 Then
        myJobOpen_.OrderType = "N"
    ElseIf cboType.ListIndex = 1 Then
        myJobOpen_.OrderType = "B"
    ElseIf cboType.ListIndex = 2 Then
        myJobOpen_.OrderType = "K"
    ElseIf cboType.ListIndex = 3 Then
        myJobOpen_.OrderType = "P"
    ElseIf cboType.ListIndex = 4 Then
        myJobOpen_.OrderType = "U"
    ElseIf cboType.ListIndex = 5 Then
        myJobOpen_.OrderType = "S"
    ElseIf cboType.ListIndex = 6 Then
        myJobOpen_.OrderType = "W"
    ElseIf cboType.ListIndex = 7 Then
        myJobOpen_.OrderType = "D"
    Else
        myJobOpen_.OrderType = "N"
    End If
    
End Sub

Private Sub cboWasteGenNo_Validate(Cancel As Boolean)
    If cboWasteGenNo.ListIndex = 0 Then
        myJobOpen_.WasteCreateNo = ""
    Else
        myJobOpen_.WasteCreateNo = cboWasteGenNo.Text
    End If
End Sub

Private Sub cbProject_Validate(Cancel As Boolean)
Dim ProID
Dim id As String
    If cbProject.Text <> "" Then
        ProID = Split(cbProject.Text, "#", -1, vbTextCompare)
        id = ProID(1)
    End If
    myJobOpen_.ProjectID = id
End Sub

Private Sub CbRemark_Validate(Cancel As Boolean)
    If CbRemark.ListIndex > 0 Then
        myJobOpen_.CancelRemark = CbRemark.Text
    End If

End Sub
Private Sub Chk_Click()
    If Chk.Value = 1 Then
        If myJobOpen_.IsChk <> "Y" Then
            If myJobOpen_.JobNo <> "" Then
                myJobOpen_.IsChk = "Y"
                myJobOpen_.UserChk = Trim(CurrentUser.StaffFName & " " & CurrentUser.StaffLName)
                myJobOpen_.SaveUserChk
                Chk.Caption = "ตรวจสอบ โดย " & Trim(CurrentUser.StaffFName & " " & CurrentUser.StaffLName)
                Chk.Enabled = False
            Else
                MsgBox "ไม่สามารถบันทึกเช็คได้กรุณาบันทึก ๋ Job ก่อน", vbExclamation
            End If
        End If
    Else
        myJobOpen_.IsChk = "N"
    End If
   
End Sub

Private Sub ChkB_Click()
Dim tmpDet As BWGJobDataCarUsed
Dim tmpKey As String, BoxCheck As Integer
Dim tmpVal As String, OldValue As String
    If grCar.TextMatrix(1, 0) = "" Then Exit Sub
    tmpKey = Trim(grCar.TextMatrix(1, 0))
    
    If isExist(myJobOpen_.Cars, "'" & tmpKey & "'") = False Then
        Set tmpDet = New BWGJobDataCarUsed
        tmpDet.id = GetGUID
        tmpDet.JobID = myJobOpen_.id
        
        grCar.TextMatrix(1, 0) = tmpDet.id
        myJobOpen_.Cars.Add tmpDet, "'" & tmpDet.id & "'"
    Else
        Set tmpDet = myJobOpen_.Cars("'" & tmpKey & "'")
    End If
    
    tmpVal = Trim(grCar.TextMatrix(1, 0))
    tmpDet.RowID = 1
    With tmpDet
        If ChkB.Value = 1 Then
            .PhotoOnBelt = "Y"
        Else
            .PhotoOnBelt = "N"
        End If
                myJobOpen_.isChanged = True
    End With

End Sub

Private Sub ChkC_Click()
Dim tmpDet As BWGJobDataCarUsed
Dim tmpKey As String, BoxCheck As Integer
Dim tmpVal As String, OldValue As String
    tmpKey = Trim(grCar.TextMatrix(1, 0))
    If isExist(myJobOpen_.Cars, "'" & tmpKey & "'") = False Then
        Set tmpDet = New BWGJobDataCarUsed
        tmpDet.id = GetGUID
        tmpDet.JobID = myJobOpen_.id
        
        grCar.TextMatrix(1, 0) = tmpDet.id
    Else
        Set tmpDet = myJobOpen_.Cars("'" & tmpKey & "'")
    End If
    tmpVal = Trim(grCar.TextMatrix(1, 0))
    tmpDet.RowID = 1
    With tmpDet
        If ChkC.Value = 1 Then
            .CustCon = "Y"
            CboC.Enabled = True
            If gVal = "1" Then
                .ConName = CboC.Text
                Else
                .ConName = .ConName
            End If
       Else
            .CustCon = "N"
            .ConName = ""
            CboC.Enabled = False
        End If
        If Trim(CboC.Text) = "" Then
            .CustCon = "N"
            .ConName = ""
        End If
                myJobOpen_.isChanged = True
    End With
        
End Sub

Private Sub ChkCarCheck_Click()
Dim tmpDet As BWGJobDataCarUsed
Dim tmpKey As String, BoxCheck As Integer
Dim tmpVal As String, OldValue As String
    If grCar.TextMatrix(1, 0) = "" Then Exit Sub
    tmpKey = Trim(grCar.TextMatrix(1, 0))
    If isExist(myJobOpen_.Cars, "'" & tmpKey & "'") = False Then
        Set tmpDet = New BWGJobDataCarUsed
        tmpDet.id = GetGUID
        tmpDet.JobID = myJobOpen_.id
    Else
        Set tmpDet = myJobOpen_.Cars("'" & tmpKey & "'")
    End If
    tmpVal = Trim(grCar.TextMatrix(1, 0))
    tmpDet.RowID = 1
    With tmpDet
        If ChkCarCheck.Value = 1 Then
            .CarCheck = "Y"
        Else
            .CarCheck = "N"
        End If
                myJobOpen_.isChanged = True
    End With


End Sub

Private Sub ChkCloseWeight_Click()
Dim tmpDet As BWGJobDataCarUsed
Dim tmpKey As String, BoxCheck As Integer
Dim tmpVal As String, OldValue As String
    If grCar.TextMatrix(1, 0) = "" Then Exit Sub
    tmpKey = Trim(grCar.TextMatrix(1, 0))
    If isExist(myJobOpen_.Cars, "'" & tmpKey & "'") = False Then
        Set tmpDet = New BWGJobDataCarUsed
        tmpDet.id = GetGUID
        tmpDet.JobID = myJobOpen_.id
        
        grCar.TextMatrix(1, 0) = tmpDet.id
        myJobOpen_.Cars.Add tmpDet, "'" & tmpDet.id & "'"
    Else
        Set tmpDet = myJobOpen_.Cars("'" & tmpKey & "'")
    End If
    tmpVal = Trim(grCar.TextMatrix(1, 0))
    tmpDet.RowID = 1
    With tmpDet
        If ChkCloseWeight.Value = 1 Then
            .CloseWeight = "Y"
        Else
            .CloseWeight = "N"
        End If
                myJobOpen_.isChanged = True
    End With

End Sub

Private Sub ChkCovidWaste_Click()
    If ChkCovidWaste.Value = 1 Then
        myJobOpen_.CovidWaste = "Y"
        Else
        myJobOpen_.CovidWaste = "N"
    End If
End Sub


Private Sub ChkCus_Click()
Dim tmpDet As BWGJobDataCarUsed
Dim tmpKey As String, BoxCheck As Integer
Dim tmpVal As String, OldValue As String
    tmpKey = Trim(grCar.TextMatrix(1, 0))
    If isExist(myJobOpen_.Cars, "'" & tmpKey & "'") = False Then
        Set tmpDet = New BWGJobDataCarUsed
        tmpDet.id = GetGUID
        tmpDet.JobID = myJobOpen_.id
    Else
        Set tmpDet = myJobOpen_.Cars("'" & tmpKey & "'")
    End If
    tmpVal = Trim(grCar.TextMatrix(1, 0))
    tmpDet.RowID = 1
    With tmpDet
        If ChkCus.Value = 1 Then
            .CustPhoto = "Y"
        Else
            .CustPhoto = "N"
        End If
                myJobOpen_.isChanged = True
    End With

End Sub

Private Sub ChkEm_Click()
Dim tmpDet As BWGJobDataCarUsed
Dim tmpKey As String, BoxCheck As Integer
Dim tmpVal As String, OldValue As String
    tmpKey = Trim(grCar.TextMatrix(1, 0))
    If isExist(myJobOpen_.Cars, "'" & tmpKey & "'") = False Then
        Set tmpDet = New BWGJobDataCarUsed
        tmpDet.id = GetGUID
        tmpDet.JobID = myJobOpen_.id
    Else
        Set tmpDet = myJobOpen_.Cars("'" & tmpKey & "'")
    End If
    tmpVal = Trim(grCar.TextMatrix(1, 0))
    tmpDet.RowID = 1
    With tmpDet
        If ChkEm.Value = 1 Then
            .EmpTrain = "Y"
        Else
            .EmpTrain = "N"
        End If
                myJobOpen_.isChanged = True
    End With

End Sub

Private Sub ChkNewCustJob_Click()
Dim tmpDet As BWGJobDataCarUsed
Dim tmpKey As String, BoxCheck As Integer
Dim tmpVal As String, OldValue As String
    tmpKey = Trim(grCar.TextMatrix(1, 0))
    If isExist(myJobOpen_.Cars, "'" & tmpKey & "'") = False Then
        Set tmpDet = New BWGJobDataCarUsed
        tmpDet.id = GetGUID
        tmpDet.JobID = myJobOpen_.id
    Else
        Set tmpDet = myJobOpen_.Cars("'" & tmpKey & "'")
    End If
    tmpVal = Trim(grCar.TextMatrix(1, 0))
    tmpDet.RowID = 1
    With tmpDet
        If ChkNewCustJob.Value = 1 Then
            .NewCustJob = "Y"
        Else
            .NewCustJob = "N"
        End If
                myJobOpen_.isChanged = True
    End With

End Sub

Private Sub ChkNewCustJobBI_Click()
Dim tmpDet As BWGJobDataCarUsed
Dim tmpKey As String, BoxCheck As Integer
Dim tmpVal As String, OldValue As String
    tmpKey = Trim(grCar.TextMatrix(1, 0))
    If isExist(myJobOpen_.Cars, "'" & tmpKey & "'") = False Then
        Set tmpDet = New BWGJobDataCarUsed
        tmpDet.id = GetGUID
        tmpDet.JobID = myJobOpen_.id
        grCar.TextMatrix(1, 0) = tmpDet.id
    Else
        Set tmpDet = myJobOpen_.Cars("'" & tmpKey & "'")
    End If
    tmpVal = Trim(grCar.TextMatrix(1, 0))
    tmpDet.RowID = 1
    With tmpDet
        If ChkNewCustJobBI.Value = 1 Then
            .NewCustJobBI = "Y"
        Else
            .NewCustJobBI = "N"
        End If
                myJobOpen_.isChanged = True
    End With
End Sub


Private Sub chkPhotoS_Validate(Cancel As Boolean)
    If chkPhotoS.Value = 1 Then
        myJobOpen_.PhotoS = "Y"
    Else
        myJobOpen_.PhotoS = "N"
    End If
    myJobOpen_.isChanged = True
End Sub

Private Sub chkS_Click()
Dim tmpDet As BWGJobDataCarUsed
Dim tmpKey As String, BoxCheck As Integer
Dim tmpVal As String, OldValue As String
    If grCar.TextMatrix(1, 0) = "" Then Exit Sub
    tmpKey = Trim(grCar.TextMatrix(1, 0))
    If isExist(myJobOpen_.Cars, "'" & tmpKey & "'") = False Then
        Set tmpDet = New BWGJobDataCarUsed
        tmpDet.id = GetGUID
        tmpDet.JobID = myJobOpen_.id
        
        grCar.TextMatrix(1, 0) = tmpDet.id
        myJobOpen_.Cars.Add tmpDet, "'" & tmpDet.id & "'"
    Else
        Set tmpDet = myJobOpen_.Cars("'" & tmpKey & "'")
    End If
    tmpVal = Trim(grCar.TextMatrix(1, 0))
    tmpDet.RowID = 1
    With tmpDet
        If chkS.Value = 1 Then
            .PhotoS = "Y"
        Else
            .PhotoS = "N"
        End If
                myJobOpen_.isChanged = True
    End With
End Sub

Private Sub ChkW_Click()
Dim tmpDet As BWGJobDataCarUsed
Dim tmpKey As String, BoxCheck As Integer
Dim tmpVal As String, OldValue As String
    If grCar.TextMatrix(1, 0) = "" Then Exit Sub
    tmpKey = Trim(grCar.TextMatrix(1, 0))
    If isExist(myJobOpen_.Cars, "'" & tmpKey & "'") = False Then
        Set tmpDet = New BWGJobDataCarUsed
        tmpDet.id = GetGUID
        tmpDet.JobID = myJobOpen_.id
        
        grCar.TextMatrix(1, 0) = tmpDet.id
        myJobOpen_.Cars.Add tmpDet, "'" & tmpDet.id & "'"
    Else
        Set tmpDet = myJobOpen_.Cars("'" & tmpKey & "'")
    End If
    tmpVal = Trim(grCar.TextMatrix(1, 0))
    tmpDet.RowID = 1
    With tmpDet
        If ChkW.Value = 1 Then
            .PhotoOnWeight = "Y"
        Else
            .PhotoOnWeight = "N"
        End If
                myJobOpen_.isChanged = True
    End With
End Sub

Public Sub cmdAddItem_Click()
Dim Sql As String
Dim RSDet As New ADODB.Recordset
Dim i As Integer
    If Trim(txtCustomer.Text) = "" Then
        MsgBox "กรุณาเลือกลูกค้าก่อนทำการเพิ่มรายการกากของเสีย", vbExclamation
        txtCustomer.SetFocus
        Exit Sub
    End If
    If cboCompany.ListCount > 1 Then
                If cboCompany.ListIndex < 0 Then
            MsgBox "กรุณาเลือกบริษัทที่ทำการเปิด Job ก่อนการเพิ่มกากของเสีย", vbExclamation
            cboCompany.SetFocus
            Exit Sub
        End If
    End If
    If Trim(curCustomer_.WasteSK2Exp) <> "" Then
    Dim DD As Date
    DD = FormatYMD_to_DMY(curCustomer_.WasteSK2Exp)
'    If DD < Date Then
'        MsgBox "ข้อมูล สก.2 ของลูกค้านี้ หมดอายุแล้ว กรุณาต่ออายุ สก.2 ก่อนทำการออกใบเสนอราคา", vbExclamation + vbOKOnly, "ติดต่อแผนก สวล."
'        Exit Sub
'    End If
    End If
    isDetailChange_ = False
    frmJobWasteDataAdd.Label3.Caption = Me.Caption
    frmJobWasteDataAdd.Label4.Caption = Trim(curCustomer_.WasteSK2Exp)
    
    If myJobOpen_.ServiceOrderID <> "" Then
        Sql = "SELECT QuotationDetID FROM Service_OrderDT " & vbCrLf
        Sql = Sql & "where OrderID='" & myJobOpen_.ServiceOrderID & "' "
        Set RSDet = GetRS(Sql, adOpenForwardOnly, adLockOptimistic)
            If RSDet.RecordCount > 0 Then
            With frmJobWasteDataAdd.GrOrderID
            .Clear
            .Rows = RSDet.RecordCount + 1
                For i = 1 To RSDet.RecordCount
                    .TextMatrix(i, 0) = RSDet!QuotationDetID
                RSDet.MoveNext
                Next
            End With
            End If
    End If
    
    Load frmJobWasteDataAdd
    With frmJobWasteDataAdd
            .WorkDate = dtWorkDate.ValueDMY
        If curCustomer_.CustomerType = "R" Then

        
            .Status = " OR (Status IN ('A', 'IA','P'))"
              If cboWasteGenNo.ListIndex = 0 Then
                .SetCallerForm Me, curCustomer_, "", CurrentUser.WorkCompanies(cboCompany.ListIndex).id, "", ""
            Else
                .SetCallerForm Me, curCustomer_, "", "776EDA5830F744A6B3551348D851BC22", cboWasteGenNo.Text, ""
            End If
        Else
            .Status = " AND (Status IN ('A', 'IA'))"
            If cboOrderType.ListIndex = 0 Then
                .isJob = "N"
            ElseIf cboOrderType.ListIndex = 1 Then
                .isJob = "Y"
            ElseIf cboOrderType.ListIndex = 2 Then
                .isJob = "D"
            End If
        If cboWasteGenNo.ListIndex = 0 Then
            .SetCallerForm Me, curCustomer_, "", "776EDA5830F744A6B3551348D851BC22", "", ""
        Else
            .SetCallerForm Me, curCustomer_, "", "776EDA5830F744A6B3551348D851BC22", cboWasteGenNo.Text, ""
        End If
        End If
        If myJobOpen_.ProjectID <> "" Then
            .ProjectID = myJobOpen_.ProjectID
        End If
        .Show 1
        If isDetailChange_ Then
            ShowWasteData
        End If
    End With
End Sub

Private Sub cmdCancel_Click()
    Dim result
    result = MsgBox("ต้องการยกเลิก Job นี้ใช่หรือไม่", vbQuestion + vbYesNo)
    If result = vbYes Then
        If CbRemark.ListIndex = 0 Then
            MsgBox "หากต้องการยกเลิก Order กรุณาระบุสาเหตุของการยกเลิกก่อน", vbExclamation
            Exit Sub
        End If
        myJobOpen_.CancelStaff = CurrentUser.StaffFName & " " & CurrentUser.StaffLName
        myJobOpen_.cancelTime = CStr(datetime.Now)
        myJobOpen_.CancelJob
        If myJobOpen_.isCancel Then
            lbCancel.Visible = True
        Else
            lbCancel.Visible = False
        End If
        SetEnabledScreen
    End If
End Sub

Private Sub cmdClear_Click()
Dim selResult As Integer
    If myJobOpen_.isChanged And myJobOpen_.isCancel = False And myJobOpen_.MenifestCountIssued = 0 And myJobOpen_.isCompleted = False Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนสร้างใหม่หรือไม่?", vbQuestion + vbYesNoCancel)
        If selResult = vbYes Then
            CmdSave_Click
            If myJobOpen_.isChanged Then Exit Sub
        ElseIf selResult = vbCancel Then
            Exit Sub
        End If
    End If
    ClearScreen
    CreateNewJobOpen
    dtWorkDate.SetFocus
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub CmdSave_Click()
Dim x%, y%, isCarInUse As Boolean, OrderBreak As String
Dim TypeCode As String
    If curCustomer_ Is Nothing = False Then
        TypeCode = Left(curCustomer_.CustomerCode, 2)
        If TypeCode <> "MS" And TypeCode <> "MI" Then
            TypeCode = Left(curCustomer_.CustomerCode, 1)
        End If
            If curCustomer_.CustomerType <> TypeCode Then
                MsgBox "ข้อมูลฐานลูกค้ากับ Code ของลูกค้าไม่ตรงกัน กรุณาแก้ไขให้ถูกต้อง" & curCustomer_.CustomerCode & "ฐานลูกค้า " & curCustomer_.CustomerType
                Exit Sub
            End If
            If (TypeCode = "C" Or TypeCode = "S") And curCustomer_.ZoneName = "BME_AP" Then
                MsgBox "ข้อมูลฐานลูกค้ากับ เขตการขายไม่สอดคล้องกัน กรุณาตรวจสอบ" & curCustomer_.CustomerCode & "ฐานลูกค้า " & curCustomer_.CustomerType
                Exit Sub
            End If
            If (TypeCode = "M" Or TypeCode = "MS") And curCustomer_.ZoneName = "AKP_Zone" Then
                MsgBox "ข้อมูลฐานลูกค้ากับ เขตการขายไม่สอดคล้องกัน กรุณาตรวจสอบ" & curCustomer_.CustomerCode & "ฐานลูกค้า " & curCustomer_.CustomerType
                Exit Sub
            End If
    End If
    
    'เช็คข้อมูล Service Order
    If cboConfirm.ListCount > 0 Then
        If cboConfirm.ListIndex < 0 Then
            MsgBox "ลูกค้ามีการสั่งออเดอร์ผ่านระบบ AKP Customer แต่ท่านยังไม่ได้เลือก Confirm", vbExclamation + vbOKOnly, "แจ้งเตือนการ Confirm Order ที่มีการสั่งผ่าน AKPCustomer"
        End If
    End If
     If txtCostTrans.Text = "" And Not IsNumeric(txtCostTrans.Text) Then
        txtCostTrans.Text = "0"
     End If
        If cboTruckType.ListIndex = 1 Then
            If CDbl(txtCostTrans.Text) <= 0 Then
                MsgBox "กรุณากรอกต้นทุนของรถเช่า", vbExclamation + vbOKOnly, "เลือกรถเช่า"
                Exit Sub
            End If
        End If
'If LbNumber.Caption <> "" Then
'    MsgBox "ใบงานได้จัดลำดับไปแล้ว  ไม่สามารถแก้ไขได้  หาแก้ไขต้องยกเลิกแล้วลงข้อมูลใหม่", vbExclamation
'    Exit Sub
'End If

If CbOrdCompany.ListIndex = 0 Then
            MsgBox "กรุณาระบุ" & Label4.Caption, vbExclamation
            CbOrdCompany.SetFocus
    Exit Sub
End If

    If cboCompany.ListCount > 1 Then
        If cboCompany.ListIndex < 0 Then
            MsgBox "กรุณาระบุบริษัทที่ทำการเปิด Job", vbExclamation
            cboCompany.SetFocus
            Exit Sub
        End If
    End If
    If Trim(txtCustomer.Text) = "" Then
        MsgBox "กรุณาใส่ชื่อลูกค้า", vbExclamation
        txtCustomer.SetFocus
        Exit Sub
    End If
    If Trim(cboCustSite.Text) = "" Then
        MsgBox "กรุณาระบุสถานที่เก็บขน", vbExclamation
        cboCustSite.SetFocus
        Exit Sub
    End If
    If myJobOpen_.JobDetails Is Nothing Then
        MsgBox "กรุณาระบุรายการกากของเสียอย่างน้อย 1 รายการ", vbExclamation
        Exit Sub
    End If
    If myJobOpen_.JobDetails.Count = 0 Then
        MsgBox "กรุณาระบุรายการกากของเสียอย่างน้อย 1 รายการ", vbExclamation
        Exit Sub
    End If
    If myJobOpen_.Cars Is Nothing Then
        MsgBox "กรุณาระบุรถที่ต้องใช้ในการขนอย่างน้อย 1 รายการ", vbExclamation
        Exit Sub
    End If
    If myJobOpen_.Cars.Count = 0 Then
        MsgBox "กรุณาระบุรถที่ต้องใช้ในการขนอย่างน้อย 1 รายการ", vbExclamation
        Exit Sub
    End If
    ' ตรวจสอบการระงับออเดอร์
    OrderBreak = CheckforOrderBreak
    If OrderBreak <> "" Then
        MsgBox OrderBreak, vbExclamation
        Exit Sub
    End If
    
    For x = 1 To myJobOpen_.JobDetails.Count
        If Trim(myJobOpen_.JobDetails(x).TSDFID) = "" Then
            MsgBox "กรุณาระบุปลายทางบ่อกำจัดให้ครบถ้วน", vbExclamation
            grWasteData.SetFocus
            Exit Sub
        End If
    Next
    For x = 1 To myJobOpen_.Cars.Count
'       จุ๊เอาออกเพราะรถพวงคือ1เที่ยว
'        If myJobOpen_.Cars(x).isTrailer And myJobOpen_.Cars(x).TruckCount Mod 2 <> 0 Then
'            MsgBox "กรุณาระบุจำนวนเที่ยวรถสำหรับรถพ่วงเป็นเลขคู่", vbExclamation
'            Exit Sub
        If myJobOpen_.Cars(x).GoSameTime = "" Then
            MsgBox "กรุณาระบุว่ารถไปพร้อมกันหรือไม่ในแต่ละเที่ยวให้ครบถ้วน", vbExclamation
            Exit Sub
        End If
        If myJobOpen_.Cars(x).TimeTables.Count <> myJobOpen_.Cars(x).TruckCount Then
            MsgBox "กรุณาระบุตารางการเดินรถของรถคันที่ " & myJobOpen_.Cars(x).carNo & " ให้ครบถ้วน", vbExclamation
            Exit Sub
        Else
            For y = 1 To myJobOpen_.Cars(x).TimeTables.Count
                If Trim(myJobOpen_.Cars(x).TimeTables(y).StartTime) = "" Then
                    MsgBox "กรุณาระบุตารางการเดินรถของรถคันที่ " & myJobOpen_.Cars(x).carNo & " ให้ครบถ้วน", vbExclamation
                    Exit Sub
                End If
            Next
        End If
    Next
    
    With grWasteData.getGridObj
        For x = 1 To myJobOpen_.Cars.Count
            isCarInUse = False
            For y = 1 To .Rows - 1
                If IsNumeric(Trim(.TextMatrix(y, 8))) Then
                    If myJobOpen_.Cars(x).carNo = CInt(Trim(.TextMatrix(y, 8))) Then
                        isCarInUse = True
                        Exit For
                    End If
                End If
            Next y
            If isCarInUse = False Then
                MsgBox "รถคันที่ " & myJobOpen_.Cars(x).carNo & " ไม่ได้ถูกใช้งานแล้ว กรุณาลบออกจากรายการรถที่ใช้ขน", vbExclamation
                grCar.SetFocus
        Exit Sub
            End If
        Next x
    End With
    If curCustomer_.CustomerType = "R" And Trim(txtPONo.Text) = "" Then
        MsgBox "กรุณาระบุเลข PO ", vbExclamation
        Exit Sub
    End If
    
    If Not (curCar_ Is Nothing) Then
    With curCar_
        If ChkNewCustJob.Value = 1 Then
            .NewCustJob = "Y"
        Else
            .NewCustJob = "N"
        End If
        If ChkCarCheck.Value = 1 Then
            .CarCheck = "Y"
        Else
            .CarCheck = "N"
        End If
        If ChkEm.Value = 1 Then
            .EmpTrain = "Y"
        Else
            .EmpTrain = "N"
        End If
        If ChkCus.Value = 1 Then
            .CustPhoto = "Y"
        Else
            .CustPhoto = "N"
        End If
        If ChkC.Value = 1 Then
            .CustCon = "Y"
            .ConName = CboC.Text
       Else
            .CustCon = "N"
            .ConName = ""
        End If
        If Trim(CboC.Text) = "" Then
            .CustCon = "N"
            .ConName = ""
        End If
        If ChkNewCustJobBI.Value = 1 Then
            .NewCustJobBI = "Y"
        Else
            .NewCustJobBI = "N"
        End If
        ' --------ภาชนะ
        If TxtContainer.Text <> "" Then
            .Container = Trim(TxtContainer.Text)
        End If
        If txtBoutS.Text <> "" Then
            .BoutS = Trim(txtBoutS.Text)
        End If
        If txtLack.Text <> "" Then
            .Lack = Trim(txtLack.Text)
        End If
        If TxtLackNew.Text <> "" Then
            .LackNew = Trim(TxtLackNew.Text)
        End If
        If TxtLackGas.Text <> "" Then
            .LackGas = Trim(TxtLackGas.Text)
        End If
        If txtPlate.Text <> "" Then
            .Plate = Trim(txtPlate.Text)
        End If
        If TxtPlastic.Text <> "" Then
            .Plastic = Trim(TxtPlastic.Text)
        End If
        If TxtTankTypeL.Text <> "" Then
            .TankType = "L"  'ถังปากกว้าง
            .Tankcount = Trim(TxtTankTypeL.Text)
            .TankPrice = .Tankcount * 150
            .TankPriceV = .Tankcount * 100
        End If
        If Trim(TxtTankTypeS.Text) <> "" Then
            .TankTypeS = "S"   'ถังปากจู๋
            .TankTypeSCount = Trim(TxtTankTypeS.Text)
            .TankPriceS = .TankTypeSCount * 150
            .TankPriceSV = .TankTypeSCount * 100
        End If
        If Trim(TxtBagTypeP.Text) <> "" Then
            .BagType = "P"   'ถุงเปิดปาก
            .BagCount = Trim(TxtBagTypeP.Text)
            .BagPrice = .BagCount * 120
            .BagPriceV = .BagCount * 100
        End If
        If Trim(TxtBagBack.Text) <> "" Then
            .BagBack = Trim(TxtBagBack.Text)
        End If
        If Trim(txtPaperBox.Text) <> "" Then
            .PaperBox = Trim(txtPaperBox.Text)
        End If
        If Trim(txtBoutCage.Text) <> "" Then
            .BoutCage = Trim(txtBoutCage.Text)
        End If
        If Trim(CovidBin.Text) <> "" Then
            .CovidBin = Trim(CovidBin.Text)
        End If
         If Trim(CovidBin200.Text) <> "" Then
            .CovidBin200 = Trim(CovidBin200.Text)
        End If
          If Trim(txtRack.Text) <> "" Then
            .Rack = Trim(txtRack.Text)
        End If
    End With
    End If
   
    If myJobOpen_.isChanged And myJobOpen_.isCancel = False And myJobOpen_.MenifestCountIssued = 0 And myJobOpen_.isCompleted = False Then
        If Trim(myJobOpen_.UnderStaffID) = "" Or myJobOpen_.UnderStaffID <> curCustomer_.UnderSaleStaffID Then
            myJobOpen_.UnderStaffID = curCustomer_.UnderSaleStaffID
            myJobOpen_.UnderStaffName = curCustomer_.UnderSaleName
        End If
        If myJobOpen_.IssuedStaffID <> "" Then
            If myJobOpen_.IssuedStaffID <> CurrentUser.id Then
                If MsgBox("ชื่อผู้บันทึก ไม่ใช่คนเดียวกับชื่อผู้สร้าง Order. กรุณายืนยันการบันทึกข้อมูล", vbExclamation + vbYesNo, "ยืนยันการบันทึก") = vbYes Then
                    myJobOpen_.Save (AKPBoxCode), (AKPBoxCode2)
                End If
            Else
                myJobOpen_.Save (AKPBoxCode), (AKPBoxCode2)
            End If
        Else
            myJobOpen_.Save (AKPBoxCode), (AKPBoxCode2)
        End If
        txtJobOpenNo.Text = myJobOpen_.JobNo
        Dim SubTypeID As String
    End If
    
    Dim sqlVisit As String
    Dim RSVisit As New ADODB.Recordset
        sqlVisit = "SELECT * FROM tbCustomerVisitHeader WHERE CustomerID='" & myJobOpen_.CustomerID & "' and VisitDate='" & myJobOpen_.WorkDate & "' "
        Set RSVisit = GetRS(sqlVisit, adOpenForwardOnly, adLockOptimistic)
            If ChkCus.Value = 1 Then
            Dim VisID, Tp, VT, Fp As String
                Tp = "ลูกค้าติดตามดูทำลาย"
                VT = "10.00"
                Fp = "30"
                VisID = GetGUID
                    If RSVisit.RecordCount > 0 Then
                         DBConn.Execute "UPDATE tbCustomerVisitHeader SET VisitDate = '" & Format(myJobOpen_.WorkDate, "yyyy/MM/dd") & "' WHERE VisitID='" & RSVisit!VisitID & "'"
                    Else
                        DBConnExc "Insert into tbCustomerVisitHeader(VisitID,Companyname,NameDisplay, CustomerID, VisitDate,VisitType,VisitTime,FullPoint) Values ('" & VisID & "','" & Trim(txtCustomer.Text) & "','" & Trim(txtCustomer.Text) & "','" & myJobOpen_.CustomerID & "','" & myJobOpen_.WorkDate & "','" & Tp & "','" & VT & "','" & Fp & "')"
                    End If
            End If
        
    If myJobOpen_.isChanged = False Then
            MsgBox "จัดเก็บข้อมูลเรียบร้อยแล้ว", vbExclamation
            dtWorkDate.Enabled = False
    End If
End Sub

Private Sub cmdOpen_Click()
Dim selResult As Integer
Dim x%
Dim tmpKey As String
Dim tmpVal As String
Dim tmpDet As BWGJobDataDetail
Dim tmpCol As Collection
    If myJobOpen_.isChanged And myJobOpen_.isCancel = False And myJobOpen_.MenifestCountIssued = 0 And myJobOpen_.isCompleted = False Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนทำงานต่อหรือไม่?", vbQuestion + vbYesNoCancel)
        If selResult = vbYes Then
            CmdSave_Click
            If myJobOpen_.isChanged Then Exit Sub
        ElseIf selResult = vbCancel Then
            Exit Sub
        End If
    End If
    Load frmJobOpenSearch
    With frmJobOpenSearch
        .SetCallerForm Me
        .Show 1
    End With
End Sub

Private Sub cmdDone_Click()
Dim result
    If myJobOpen_.isChanged And myJobOpen_.isCancel = False And cmdSave.Enabled Then
        result = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนทำงานต่อหรือไม่?", vbQuestion + vbYesNo)
        If result = vbYes Then
            CmdSave_Click
            If myJobOpen_.isChanged Then Exit Sub
        Else 'If result = vbCancel Then
            Exit Sub
        End If
    End If
    result = MsgBox("ต้องการปิด Job นี้ใช่หรือไม่", vbQuestion + vbYesNo)
    If result = vbYes Then
        myJobOpen_.CloseJob
        SetEnabledScreen
    End If
End Sub

Private Sub CovidBin_Validate(Cancel As Boolean)
Dim tmpDet As BWGJobDataCarUsed
Dim tmpKey As String, BoxCheck As Integer
Dim tmpVal As String, OldValue As String
    tmpKey = Trim(grCar.TextMatrix(1, 0))
    If isExist(myJobOpen_.Cars, "'" & tmpKey & "'") = False Then
        Set tmpDet = New BWGJobDataCarUsed
        tmpDet.JobID = myJobOpen_.id
    Else
        Set tmpDet = myJobOpen_.Cars("'" & tmpKey & "'")
    End If
    tmpDet.RowID = 1
    With tmpDet
    If Trim(CovidBin.Text) <> "" Then
        .CovidBin = Trim(CovidBin.Text)
        myJobOpen_.isChanged = True
    End If
    End With
End Sub


Private Sub CovidBin200_Validate(Cancel As Boolean)
    Dim tmpDet As BWGJobDataCarUsed
    Dim tmpKey As String, BoxCheck As Integer
    Dim tmpVal As String, OldValue As String
        tmpKey = Trim(grCar.TextMatrix(1, 0))
        If isExist(myJobOpen_.Cars, "'" & tmpKey & "'") = False Then
            Set tmpDet = New BWGJobDataCarUsed
            tmpDet.JobID = myJobOpen_.id
        Else
            Set tmpDet = myJobOpen_.Cars("'" & tmpKey & "'")
        End If
        tmpDet.RowID = 1
        With tmpDet
        If Trim(CovidBin200.Text) <> "" Then
            .CovidBin200 = Trim(CovidBin200.Text)
            myJobOpen_.isChanged = True
        End If
        End With
End Sub

Private Sub dtFromWork_Validate(Cancel As Boolean)
    myJobOpen_.FromWork = dtFromWork.ValueYMD
End Sub

Private Sub dtWorkDate_Validate(Cancel As Boolean)
    myJobOpen_.WorkDate = dtWorkDate.ValueYMD
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn And Me.ActiveControl.Name <> "grWasteData" And Me.ActiveControl.Name <> "grCar" And Me.ActiveControl.Name <> "grTimeTable" Then SendKeys "{Tab}"
'    If Shift = vbCtrlMask And KeyCode = vbKeyF12 Then
            cmdSave.Enabled = True
            Frame6.Enabled = True
            Frame7.Enabled = True
            grCar.Editable = flexEDKbdMouse
            grTimeTable.ColEnabled(0) = True
            grTimeTable.ColEnabled(1) = True
            myJobOpen_.isChanged = True
'    End If
End Sub
Sub JobProject()
    Dim Sql As String
    Dim rs As New ADODB.Recordset
    Dim i As Integer
    Dim ProjectList As Collection
    Sql = "SELECT * FROM vw_jobProject where CustomerID='" & curCustomer_.id & "' and ProjectStatus='Q' " & vbCrLf
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
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Dim x As Integer
    CanAccess = False: CanUpdate = False
    SetPermission Me, CanAccess, CanUpdate
    If CanAccess = False Then
        MsgBox "ท่านไม่มีสิทธิใช้ระบบนี้", vbExclamation, "Warning"
        Unload Me
    End If
    With frmPicture
        cmdAddItem.Picture = .picSearch.Picture
        cmdSave.Picture = .picSave.Picture
        cmdOpen.Picture = .picOpen.Picture
        cmdClear.Picture = .picNew.Picture
        cmdCancel.Picture = .picDelete.Picture
        cmdDone.Picture = .picDone.Picture
        cmdClose.Picture = .picExit.Picture
    End With
    With cboOrderType
        .AddItem "ออร์เดอร์แบบปกติ (Routine)"
        .AddItem "ออร์เดอร์แบบ Jobs"
        .AddItem "ออร์เดอร์แบบ DOC"
        For x = 1 To WasteCases.Count
            .AddItem WasteCases(x).WasteCase
        Next
        .ListIndex = 0
    End With
    With cboTruckType
        .AddItem "รถบริษัท"
        .AddItem "รถเช่า"
        .ListIndex = 0
    End With
    
        With CboC
        .AddItem "เบาว์"
        .AddItem "ถัง 200 ลิตร"
        .AddItem "กรงเบาว์"
        .AddItem "ถังแก๊ส"
        .AddItem "Box แดง"
        .AddItem "พาเลท"
        .AddItem "กล่องพลาสติก"
        .ListIndex = 0
        .Enabled = False
    End With
    cboCaller.ListIndex = 0
    cboCallWay.ListIndex = 0

    PopulateUnit
    PopulateCompany cboCompany, cboIndex
    PopulateCancelRemark
    With CbRemark
        CbRemark.AddItem "== หากยกเลิก กรุณาเลือก =="
        For x = 1 To colCancelRemark_.Count
                            .AddItem colCancelRemark_(x).Remdesc
                        Next
                        .ListIndex = 0
    End With

    LoadBoxType ""
    LoadBoxAKP ""
    dtWorkDate.ValueYMD = TodayDate
    txtH.Text = Format(Hour(Now), "00")
    txtM.Text = Format(Minute(Now), "00")
    txtIssuedDate.Text = FormatYMD_to_DMY(TodayDate, "/", "/")
    txtIssuedStaff.Text = CurrentUser.StaffFName & " " & CurrentUser.StaffLName
    Call SetGrid
    CreateNewJobOpen
     SetGrCon
    
    'SetFrom Height & Width
    Frame2H = Frame2.Height / Me.Height * 100
    Frame3H = Frame3.Height / Me.Height * 90
    
    AKPBoxCode = ""
    AKPBoxCode2 = ""
    
    End Sub
    Public Sub PopulateDistrictFrontEnd(StaffID As String)
        Dim tmpS As New BWGSearchManager
        Dim x%
        Set colDistrict_ = tmpS.SaleDistrictSearch("isActive = 'Y'  And  ZoneActive = 'Y'  and SaleStaffID = '" & StaffID & "' ")
'        Set colDistrict_ = tmpS.SaleDistrictSearch("isActive = 'Y'  And  ZoneActive = 'Y' ")
        CbOrdCompany.Clear
        CbOrdCompany.AddItem "== กรุณาเลือก =="
        For x = 1 To colDistrict_.Count
        Dim item As String
        
          CbOrdCompany.AddItem colDistrict_(x).Zonedesc
        Next
        CbOrdCompany.ListIndex = 0
End Sub
    
    Private Sub PopulateDistrict()
        Dim tmpS As New BWGSearchManager
        Dim x%
        Set colDistrict_ = tmpS.SaleDistrictSearch("isActive = 'Y'  And  ZoneActive = 'Y'  and SaleStaffID = '" & myJobOpen_.UnderStaffID & "' ")
'        Set colDistrict_ = tmpS.SaleDistrictSearch("isActive = 'Y'  And  ZoneActive = 'Y' ")
        CbOrdCompany.Clear
        CbOrdCompany.AddItem "== กรุณาเลือก =="
        For x = 1 To colDistrict_.Count
        Dim item As String
        
          CbOrdCompany.AddItem colDistrict_(x).Zonedesc
        Next
        CbOrdCompany.ListIndex = 0
End Sub
Sub setConfirmOrderFrontEnd(OrderID As String)
    Dim Sql As String
    Dim rs As New ADODB.Recordset
    Dim i As Integer
    Dim OrderList As Collection
    If OrderID <> "" Then
    Sql = "SELECT orderno+ ' # ' +workdate as orderno,orderid FROM Service_OrderHD where orderid='" & OrderID & "' " & vbCrLf
    Set rs = GetRS(Sql, adOpenForwardOnly, adLockOptimistic)
    cboConfirm.Clear
    If rs.RecordCount > 0 Then
            For i = 1 To rs.RecordCount
                cboConfirm.AddItem rs!OrderNo
                Set OrderList = Nothing
            rs.MoveNext
            Next
            cboConfirm.ListIndex = 0
            ConfirmValidate
    End If
    Else
    Sql = "SELECT orderno+ ' # ' +workdate as orderno,orderid FROM Service_OrderHD where CustomerID='" & myJobOpen_.CustomerID & "' and (ConfirmOrder<>'Y' or ConfirmOrder is null) and isCancel<>'Y' " & vbCrLf
    Sql = Sql & "order by orderid"
    Set rs = GetRS(Sql, adOpenForwardOnly, adLockOptimistic)
    cboConfirm.Clear
    If rs.RecordCount > 0 Then
            For i = 1 To rs.RecordCount
                cboConfirm.AddItem rs!OrderNo
                Set OrderList = Nothing
            rs.MoveNext
            Next
    End If
    End If
End Sub
Sub setConfirmOrder()
    Dim Sql As String
    Dim rs As New ADODB.Recordset
    Dim i As Integer
    Dim OrderList As Collection
    Sql = "SELECT orderno+ ' # ' +workdate as orderno,orderid FROM Service_OrderHD where CustomerID='" & myJobOpen_.CustomerID & "' and (ConfirmOrder<>'Y' or ConfirmOrder is null) and isCancel<>'Y' " & vbCrLf
    Sql = Sql & "order by orderid"
    Set rs = GetRS(Sql, adOpenForwardOnly, adLockOptimistic)
    cboConfirm.Clear
    If rs.RecordCount > 0 Then
            For i = 1 To rs.RecordCount
                cboConfirm.AddItem rs!OrderNo
                Set OrderList = Nothing
            rs.MoveNext
            Next
    End If
End Sub
    Sub SetGrCon()
    Dim tSql As String
    Dim rs As New ADODB.Recordset
    Dim i As Integer
    tSql = "SELECT * FROM TbContainer" & vbCrLf
    tSql = tSql & "order by ConID"
    Set rs = GetRS(tSql, adOpenForwardOnly, adLockOptimistic)
    If rs.RecordCount > 0 Then
        With GrCon
            For i = 1 To rs.RecordCount
                .TextMatrix(i, 0) = rs!ConID
                .TextMatrix(i, 1) = rs!ConDes
                .TextMatrix(i, 2) = rs!ConNum
            rs.MoveNext
            Next
        End With
    End If
    
End Sub

Private Sub SetGrid()
Dim x%
    With grWasteData.getGridObj
        .Rows = 20
        .Cols = 14
        grWasteData.ColInputType(4) = ComboBoxEnum
        grWasteData.ColInputType(8) = ComboBoxEnum
        grWasteData.ColInputType(12) = ComboBoxEnum
        grWasteData.ColInputType(13) = ComboBoxEnum
        .TextMatrix(0, 1) = "Waste No."
        .TextMatrix(0, 2) = "ชื่อของเสีย"
        .TextMatrix(0, 3) = "ปริมาณของเสีย"
        .TextMatrix(0, 4) = "หน่วย"
        .TextMatrix(0, 5) = "บรรจุภัณฑ์"
        .TextMatrix(0, 6) = "จำนวนบรรจุภัณฑ์"
        .TextMatrix(0, 7) = "ปลายทางบ่อกำจัด"
        .TextMatrix(0, 8) = "รถคันที่"
        .TextMatrix(0, 9) = "หมายเหตุ"
        .TextMatrix(0, 10) = "ราคา"
        .TextMatrix(0, 11) = "หน่วย"
        .TextMatrix(0, 12) = "Waste ใหม่"
        .TextMatrix(0, 13) = "Waste ใหม่ (CR)่"
        .ColWidth(0) = 0
        .ColWidth(1) = 0
        .ColWidth(2) = 2700
        .ColWidth(3) = 1185
        .ColWidth(4) = 1065
        .ColWidth(5) = 0
        .ColWidth(6) = 0
        .ColWidth(7) = 1545
        .ColWidth(8) = 800
        .ColWidth(9) = 0
        .ColWidth(10) = 1045
        .ColWidth(11) = 1045
        .ColWidth(12) = 1000
        .ColWidth(13) = 1245
        For x = 1 To .Cols - 1
            .FixedAlignment(x) = 4
        Next
    End With
    With grCar
        .Rows = 20
        .Cols = 21
        .TextMatrix(0, 1) = "รถคันที่"
        .TextMatrix(0, 2) = "ประเภทรถ"
        .TextMatrix(0, 3) = "จำนวนเที่ยว"
        .TextMatrix(0, 4) = "คนงาน/คัน"
        .TextMatrix(0, 5) = "ไปพร้อมกัน?"
        .TextMatrix(0, 6) = "ถ่ายรูป?"
        .TextMatrix(0, 7) = "Box"
        .TextMatrix(0, 8) = "จำนวน"
        .TextMatrix(0, 9) = "Box อัคคี"
        .TextMatrix(0, 10) = "Box อัคคี2"
        .TextMatrix(0, 11) = "ปากกว้าง"
        .TextMatrix(0, 12) = "ปากจู๋"
        .TextMatrix(0, 13) = "เอากลับ"
        .TextMatrix(0, 14) = "ถุงบิกแบ็ค"
        .TextMatrix(0, 15) = "เบ้า"
        .TextMatrix(0, 16) = "ถุงดำ"
        .TextMatrix(0, 17) = "แล็ค"
        .TextMatrix(0, 18) = "กล่องกระดาษ"
        .TextMatrix(0, 19) = "เบ้าตัดปาก"
        .TextMatrix(0, 20) = "หมายเหตุ"
        .ColWidth(0) = 0
        .ColWidth(1) = 650
        .ColWidth(2) = 2500
        .ColWidth(3) = 1000
        .ColWidth(4) = 1000
        .ColWidth(5) = 1000
        .ColWidth(6) = 1400
        .ColWidth(7) = 700
        .ColWidth(8) = 800
        .ColWidth(9) = 950
        .ColWidth(10) = 950
        .ColWidth(11) = 0
        .ColWidth(12) = 0
        .ColWidth(13) = 0
        .ColWidth(14) = 0
        .ColWidth(15) = 0
        .ColWidth(16) = 0
        .ColWidth(17) = 0
        .ColWidth(18) = 0
        .ColWidth(19) = 0
        .ColWidth(20) = 8000
        For x = 1 To .Cols - 1
            .FixedAlignment(x) = 4
        Next
        '*** grCar ในรูปแบบ Grid Class
'        .ColInputType(5) = ComboBoxEnum
'        .ColInputType(6) = ComboBoxEnum

        '*** New VsFlexGrid
        .ColComboList(5) = ""
        .ColComboList(5) = TypeOfCar
        .ColComboList(6) = TypeOfPhotography
        .ColComboList(7) = TypeOfBox
        .ColComboList(9) = BoxAKP
        .ColComboList(10) = BoxAKP
'        .ColComboList(9) = TypeTank
'        .ColComboList(11) = TypeBag
        .Editable = flexEDKbd
    End With
    With grTimeTable.getGridObj
        .Rows = 2
        .Cols = 3
        .TextMatrix(0, 1) = "เที่ยวที่"
        .TextMatrix(0, 2) = "เวลา"
        .ColWidth(0) = 0
        .ColWidth(1) = 1050
        .ColWidth(2) = 1185
        For x = 1 To .Cols - 1
            .FixedAlignment(x) = 4
        Next
    End With
    
    With GrCon
        .Rows = 9
        .Cols = 3
        .TextMatrix(0, 1) = "บรรจุภัณฑ์"
        .TextMatrix(0, 2) = "จำนวน"
        .ColWidth(0) = 0
        .ColWidth(1) = 2250
        .ColWidth(2) = 1200
        For x = 1 To .Cols - 1
            .FixedAlignment(x) = 4
        Next
    End With

End Sub

Private Sub CreateNewJobOpen()
    Set myJobOpen_ = Nothing
    Set myJobOpen_ = New BWGJobData
    With myJobOpen_
        .id = GetGUID
        If cboCompany.ListCount > 1 Then
            If cboCompany.ListIndex > 0 Then
                .CompanyID = CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).id
            End If
        Else
            If cboCompany.ListIndex = 0 Then
                .CompanyID = CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).id
            End If
        End If
        .IssuedDate = Format(GetSystemDateTimeThai.SystemDate, "yyyy/MM/dd")
        .IssuedTime = GetSystemDateTimeThai.SystemTime
        .IssuedStaffID = CurrentUser.id
        .IssuedStaffName = CurrentUser.StaffFName & " " & CurrentUser.StaffLName
        .WorkDate = GetSystemDateTimeThai.SystemDate
        .WorkTime = Trim(txtH.Text) & ":" & Trim(txtM.Text)
        .isChanged = False
    End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim selResult As Integer
    If myJobOpen_ Is Nothing Then Exit Sub
    If myJobOpen_.isChanged And myJobOpen_.isCancel = False And myJobOpen_.MenifestCountIssued = 0 And myJobOpen_.isCompleted = False Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนปิดหน้าจอหรือไม่?", vbQuestion + vbYesNoCancel)
        If selResult = vbYes Then
            CmdSave_Click
            If myJobOpen_.isChanged Then Cancel = 1
        ElseIf selResult = vbCancel Then
            Cancel = 1
        End If
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.Width < 13995 Then Me.Width = 13995
'    If Me.Height < 10050 Then Me.Height = 10050
'    If Me.Height > 10050 Then Me.Height = 10050
    
    If Screen.Height / 15 = 600 Then Me.Height = Screen.Height
'    Me.StartUpPosition = StartUpPositionConstants.vbStartUpScreen
    Image2.Width = Me.Width
    Line2.X2 = Me.Width
    
    Frame1.Width = Me.Width - 150
'    Frame2.Width = Me.Width - 300
    Frame3.Width = Me.Width - 150
    Frame4.Left = Me.Width - Frame4.Width - 150
    Frame5.Width = Me.Width - Frame4.Width - 150
    
    Frame2.Height = Frame2H * Me.Height / 100
    Frame3.Height = Frame3H * Me.Height / 100
    Frame3.Top = Frame2.Height + Frame2.Top + 50
    Frame4.Top = Me.Height - (Frame4.Height * 1.5)
    Frame5.Top = Me.Height - (Frame5.Height * 1.5)
    
    grWasteData.Height = Frame2.Height - 300
    grCar.Height = Frame3.Height - 250
    grTimeTable.Height = Frame3.Height - 600
    
    lvResult.Width = Frame1.Width - 4800
    grWasteData.Width = Frame2.Width - 150
    Shape1.Width = Frame3.Width - 150
    Shape1.Left = Frame3.Left
    Dim OverAllWidth As Long
    OverAllWidth = 9555 + 2535
    grCar.Width = 9555 / OverAllWidth * Shape1.Width
    grTimeTable.Width = 2535 / OverAllWidth * Shape1.Width - 50
    grTimeTable.Left = grCar.Width + 50
    Label3.Width = grTimeTable.Width
    Label3.Left = grTimeTable.Left
    GrCon.Left = lvResult.Width + 1000
End Sub

Private Sub grCar_AfterEdit(ByVal row As Long, ByVal col As Long)
Dim tmpDet As BWGJobDataCarUsed
Dim tmpKey As String, BoxCheck As Integer
Dim tmpVal As String, OldValue As String
    If grCar.TextMatrix(row, 0) = "" Then Exit Sub
    tmpKey = Trim(grCar.TextMatrix(row, 0))
    If isExist(myJobOpen_.Cars, "'" & tmpKey & "'") = False Then
        Set tmpDet = New BWGJobDataCarUsed
        tmpDet.id = GetGUID
        tmpDet.JobID = myJobOpen_.id
        grCar.TextMatrix(row, 0) = tmpDet.id
        myJobOpen_.Cars.Add tmpDet, "'" & tmpDet.id & "'"
    Else
        Set tmpDet = myJobOpen_.Cars("'" & tmpKey & "'")
    End If
    tmpVal = Trim(grCar.TextMatrix(row, col))
    tmpDet.RowID = row
    With tmpDet
        Select Case col
            Case 3
                If IsNumeric(grCar.TextMatrix(row, col)) Then
                    If CDbl(grCar.TextMatrix(row, col)) < 0 Then
                        MsgBox "ระบุจำนวนที่มากกว่า 0", vbInformation
                        grCar.TextMatrix(row, col) = "0"
                        Exit Sub
                    End If
                End If
                TruckAvailable.TruckAvailable = 0
                TruckAvailable.TruckTotal = 0
                Dim TruckSubTypeID2 As String
                TruckSubTypeID2 = ""
                If .TruckSubTypeID = "33D46FC64F3D450897D1F72986EAE700" Then TruckSubTypeID2 = "F0D9A37CFC034EB89BF0F0DBA15D9523"
                If .TruckSubTypeID = "F0D9A37CFC034EB89BF0F0DBA15D9523" Then TruckSubTypeID2 = "33D46FC64F3D450897D1F72986EAE700"
                If .TruckSubTypeID = "2EFEDFB372C448E3B93C786DA8613F6A" Then TruckSubTypeID2 = "CD50C04FD207477BBA66216557868208"
                If .TruckSubTypeID = "CD50C04FD207477BBA66216557868208" Then TruckSubTypeID2 = "2EFEDFB372C448E3B93C786DA8613F6A"
                
'                If .TruckSubTypeID = "058A0646AD8D401BB7884A63C7F1F92E" Or .TruckSubTypeID = "2B89CD8FD51346EF888BD7B062170A36" Or .TruckSubTypeID = "33D46FC64F3D450897D1F72986EAE700" Or .TruckSubTypeID = "F0D9A37CFC034EB89BF0F0DBA15D9523" Then
'                    LoadBoxType "AND BoxType = 'ROL'"
'                ElseIf .TruckSubTypeID = "2EFEDFB372C448E3B93C786DA8613F6A" Or .TruckSubTypeID = "CD50C04FD207477BBA66216557868208" Then
'                    LoadBoxType "AND BoxType = 'LUG'"
'                Else
'                    LoadBoxType "AND BoxType = 'NO'"
'                End If
                If myJobOpen_.CompanyID = "776EDA5830F744A6B3551348D851BC22" Then
                Else
                    TruckAvailableCount myJobOpen_.WorkDate, .TruckSubTypeID, TruckSubTypeID2
                    If TruckAvailable.TruckUsed > TruckAvailable.TruckTotal Then
                        MsgBox "จำนวนรถ """ & .TruckSubTypeName & """ ไม่พอ ท่านอาจไม่ได้รถ" & vbCrLf & "กรุณาตรวจสอบจำนวนรถของวันนี้กับฝ่ายขนส่งอีกครั้ง, จำนวนรถที่ลงแล้ว(" & TruckAvailable.TruckUsed & ")", vbInformation
'                        grCar.TextMatrix(Row, Col) = "0"
'                        Set tmpDet = Nothing
'                        Exit Sub
                    End If
                 End If
'                If TruckAvailable.TruckTotal > 0 Then
'                    If TruckAvailable.TruckAvailable < tmpVal Then
'                        MsgBox "จำนวนรถ """ & .TruckSubTypeName & """ ไม่พอ ไม่สามารถลงออร์เดอร์ได้" & vbCrLf & "กรุณาตรวจสอบจำนวนรถของวันนี้กับฝ่ายขนส่งอีกครั้ง, จำนวนรถคงเหลือ (" & TruckAvailable.TruckAvailable & ")", vbInformation
''                        grCar.TextMatrix(Row, Col) = "0"
''                        Set tmpDet = Nothing
''                        Exit Sub
'                    End If
'                End If
                OldValue = Trim("" & .TruckCount)
                If IsNumeric(tmpVal) Then
                    .TruckCount = tmpVal
                Else
                    .TruckCount = 0
                End If
                If Trim(OldValue) <> Trim("" & .TruckCount) Then
                    .ClearTimeTable
                End If
                
            Case 4
                If IsNumeric(tmpVal) Then
                    .WorkerPerCar = tmpVal
                Else
                    .WorkerPerCar = 0
                End If
            Case 5
                OldValue = .GoSameTime
                If grCar.TextMatrix(row, col) = "1" Then
                    .GoSameTime = "Y"
                ElseIf grCar.TextMatrix(row, col) = "2" Then
                    .GoSameTime = "N"
                End If
                If Trim(OldValue) <> Trim(.GoSameTime) Then
                    If myJobOpen_.isCancel Or myJobOpen_.isCompleted = False Then
                        .ClearTimeTable
                        grTimeTable.ClearAllData
                    Else
                        MsgBox "ใบงานนี้ไม่สามารถแก้ไขได้"
                    End If
                End If
            Case 6
                If grCar.TextMatrix(row, col) = "1" Then
                    .TakePhoto = "Y"
                ElseIf grCar.TextMatrix(row, col) = "2" Then
                    .TakePhoto = "N"
                ElseIf grCar.TextMatrix(row, col) = "3" Then
                    .TakePhoto = "S"
'                ElseIf grCar.TextMatrix(row, col) = "4" Then
'                    .TakePhoto = "N"
                End If
            Case 7
            'โค้ดเดิม
                .BoxType = grCar.TextMatrix(row, col)
                If grCar.TextMatrix(row, 2) = "รถ Lugger Box, รถ Lugger Box" Or grCar.TextMatrix(row, 2) = "รถ Lugger Box (พ่วง), รถ Lugger Box" Then
                    If grCar.TextMatrix(row, 7) <> "LUG" Then
                        MsgBox "ท่านต้องระบุประเภท Box ให้ตรงกับประเภทรถ", vbExclamation
                        grCar.TextMatrix(row, 7) = ""
                    End If
                ElseIf grCar.TextMatrix(row, 2) = "รถ Roll off Truck - พ่วง (พ่วง), รถ Roll off Truck พ่วง" Or grCar.TextMatrix(row, 2) = "รถ Roll off Truck - เดี่ยว, รถ Roll off Truck- เดี่ยว" Then
                   If grCar.TextMatrix(row, 7) <> "ROL" Then
                        MsgBox "ท่านต้องระบุประเภท Box ให้ตรงกับประเภทรถ", vbExclamation
                        grCar.TextMatrix(row, 7) = ""
                    End If
                End If

            Case 8
                If Trim(grCar.TextMatrix(row, 7)) <> "" Then
                    BoxCheck = CInt(grCar.TextMatrix(row, 8))
                    If BoxCheck > BoxTypeCount(grCar.TextMatrix(row, 7)) Then
                        .BoxCount = 0
                        grCar.TextMatrix(row, 8) = ""
                        MsgBox "จำนวนที่ระบุมากกว่าจำนวนที่สามารถจัดบ็อกได้", vbInformation
                    Else
                        .BoxCount = BoxCheck
                    End If
                Else
                    grCar.TextMatrix(row, 8) = ""
                End If
                
            Case 9
                AKPBoxCode = grCar.TextMatrix(row, col)
            Case 10
                 AKPBoxCode2 = grCar.TextMatrix(row, col)
           Case 11
              .TankType = "L"  'ถังปากกว้าง
              .Tankcount = Trim(tmpVal)
              .TankPrice = .Tankcount * 150
              .TankPriceV = .Tankcount * 100
              
                    Dim tSql As String
                    Dim rs As New ADODB.Recordset
                    Dim Num As Integer
                    Dim Sql As String
                    Dim RS1 As New ADODB.Recordset
                    tSql = "SELECT * FROM TbContainerTake " & vbCrLf
                    tSql = tSql & "where jobid='" & .JobID & "' and conid='1' "
                    Set rs = GetRS(tSql, adOpenForwardOnly, adLockOptimistic)
                    
                    Sql = "SELECT * FROM TbContainer " & vbCrLf
                    Sql = Sql & "where conid='1' "
                    Set RS1 = GetRS(Sql, adOpenForwardOnly, adLockOptimistic)
                    Num = RS1!ConNum
                    
                        If rs.RecordCount > 0 Then
                        Dim OldN As Integer
                            OldN = rs!TakeNum
                            If OldN > Trim(tmpVal) Then
                                OldN = CDbl(OldN) - CDbl(Trim(tmpVal))
                                DBConnExc "Update TbContainer Set ConNum = '" & CDbl(Num) + CDbl(OldN) & "' WHERE conid='1' "
                            Else
                                OldN = CDbl(Trim(tmpVal)) - CDbl(OldN)
                                DBConnExc "Update TbContainer Set ConNum = '" & CDbl(Num) - CDbl(OldN) & "' WHERE conid='1' "
                            End If
                            DBConnExc "Update TbContainerTake Set TakeNum = '" & Trim(tmpVal) & "',LastUpdate='" & FormatYMD_to_DMY(TodayDate, "/", "/") & "' ,LastUser='" & Trim(CurrentUser.StaffFName & " " & CurrentUser.StaffLName) & "' WHERE JobID='" & .JobID & "' and conid='1' "
                        Else
                            DBConnExc "Insert into TbContainerTake(JobID, ConID, SaleID,SaleName,TakeNum,TakeDate) Values ('" & .JobID & "','1','" & myJobOpen_.UnderStaffID & "','" & txtSaleStaff.Text & "','" & Trim(tmpVal) & "','" & FormatYMD_to_DMY(TodayDate, "/", "/") & "')"
                            DBConnExc "Update TbContainer Set ConNum = '" & CDbl(Num) - CDbl(Trim(tmpVal)) & "' WHERE conid='1' "
                       End If
                    SetGrCon
                    
            Case 12
               .TankTypeS = "S"   'ถังปากจู๋
                .TankTypeSCount = Trim(tmpVal)
                .TankPriceS = .TankTypeSCount * 150
                .TankPriceSV = .TankTypeSCount * 100
                
                    tSql = "SELECT * FROM TbContainerTake " & vbCrLf
                    tSql = tSql & "where jobid='" & .JobID & "' and conid='2' "
                    Set rs = GetRS(tSql, adOpenForwardOnly, adLockOptimistic)
                    
                    Sql = "SELECT * FROM TbContainer " & vbCrLf
                    Sql = Sql & "where conid='2' "
                    Set RS1 = GetRS(Sql, adOpenForwardOnly, adLockOptimistic)
                    Num = RS1!ConNum
                    
                        If rs.RecordCount > 0 Then
                            OldN = rs!TakeNum
                            If OldN > Trim(tmpVal) Then
                                OldN = CDbl(OldN) - CDbl(Trim(tmpVal))
                                DBConnExc "Update TbContainer Set ConNum = '" & CDbl(Num) + CDbl(OldN) & "' WHERE conid='2' "
                            Else
                                OldN = CDbl(Trim(tmpVal)) - CDbl(OldN)
                                DBConnExc "Update TbContainer Set ConNum = '" & CDbl(Num) - CDbl(OldN) & "' WHERE conid='2' "
                            End If
                            DBConnExc "Update TbContainerTake Set TakeNum = '" & Trim(tmpVal) & "',LastUpdate='" & FormatYMD_to_DMY(TodayDate, "/", "/") & "' ,LastUser='" & Trim(CurrentUser.StaffFName & " " & CurrentUser.StaffLName) & "' WHERE JobID='" & .JobID & "' and conid='2' "
                        Else
                            DBConnExc "Insert into TbContainerTake(JobID, ConID, SaleID,SaleName,TakeNum,TakeDate) Values ('" & .JobID & "','2','" & myJobOpen_.UnderStaffID & "','" & txtSaleStaff.Text & "','" & Trim(tmpVal) & "','" & FormatYMD_to_DMY(TodayDate, "/", "/") & "')"
                            DBConnExc "Update TbContainer Set ConNum = '" & CDbl(Num) - CDbl(Trim(tmpVal)) & "' WHERE conid='2' "
                       End If
                    SetGrCon

            Case 13
                .TankTypeSB = Trim(tmpVal)
            Case 14
                  .BagType = "P"   'ถุงเปิดปาก
                  .BagCount = Trim(tmpVal)
                  .BagPrice = .BagCount * 120
                  .BagPriceV = .BagCount * 100
                    
                    tSql = "SELECT * FROM TbContainerTake " & vbCrLf
                    tSql = tSql & "where jobid='" & .JobID & "' and conid='4' "
                    Set rs = GetRS(tSql, adOpenForwardOnly, adLockOptimistic)
                    
                    Sql = "SELECT * FROM TbContainer " & vbCrLf
                    Sql = Sql & "where conid='4' "
                    Set RS1 = GetRS(Sql, adOpenForwardOnly, adLockOptimistic)
                    Num = RS1!ConNum
                    
                        If rs.RecordCount > 0 Then
                            OldN = rs!TakeNum
                            If OldN > Trim(tmpVal) Then
                                OldN = CDbl(OldN) - CDbl(Trim(tmpVal))
                                DBConnExc "Update TbContainer Set ConNum = '" & CDbl(Num) + CDbl(OldN) & "' WHERE conid='4' "
                            Else
                                OldN = CDbl(Trim(tmpVal)) - CDbl(OldN)
                                DBConnExc "Update TbContainer Set ConNum = '" & CDbl(Num) - CDbl(OldN) & "' WHERE conid='4' "
                            End If
                            DBConnExc "Update TbContainerTake Set TakeNum = '" & Trim(tmpVal) & "',LastUpdate='" & FormatYMD_to_DMY(TodayDate, "/", "/") & "' ,LastUser='" & Trim(CurrentUser.StaffFName & " " & CurrentUser.StaffLName) & "' WHERE JobID='" & .JobID & "' and conid='4' "
                        Else
                            DBConnExc "Insert into TbContainerTake(JobID, ConID, SaleID,SaleName,TakeNum,TakeDate) Values ('" & .JobID & "','4','" & myJobOpen_.UnderStaffID & "','" & txtSaleStaff.Text & "','" & Trim(tmpVal) & "','" & FormatYMD_to_DMY(TodayDate, "/", "/") & "')"
                            DBConnExc "Update TbContainer Set ConNum = '" & CDbl(Num) - CDbl(Trim(tmpVal)) & "' WHERE conid='4' "
                       End If
                    SetGrCon

            Case 15  'เบ้า
                  .Container = Trim(tmpVal)
                  
                    tSql = "SELECT * FROM TbContainerTake " & vbCrLf
                    tSql = tSql & "where jobid='" & .JobID & "' and conid='5' "
                    Set rs = GetRS(tSql, adOpenForwardOnly, adLockOptimistic)
                    
                    Sql = "SELECT * FROM TbContainer " & vbCrLf
                    Sql = Sql & "where conid='5' "
                    Set RS1 = GetRS(Sql, adOpenForwardOnly, adLockOptimistic)
                    Num = RS1!ConNum
                    
                        If rs.RecordCount > 0 Then
                            OldN = rs!TakeNum
                            If OldN > Trim(tmpVal) Then
                                OldN = CDbl(OldN) - CDbl(Trim(tmpVal))
                                DBConnExc "Update TbContainer Set ConNum = '" & CDbl(Num) + CDbl(OldN) & "' WHERE conid='5' "
                            Else
                                OldN = CDbl(Trim(tmpVal)) - CDbl(OldN)
                                DBConnExc "Update TbContainer Set ConNum = '" & CDbl(Num) - CDbl(OldN) & "' WHERE conid='5' "
                            End If
                            DBConnExc "Update TbContainerTake Set TakeNum = '" & Trim(tmpVal) & "',LastUpdate='" & FormatYMD_to_DMY(TodayDate, "/", "/") & "' ,LastUser='" & Trim(CurrentUser.StaffFName & " " & CurrentUser.StaffLName) & "' WHERE JobID='" & .JobID & "' and conid='5' "
                        Else
                            DBConnExc "Insert into TbContainerTake(JobID, ConID, SaleID,SaleName,TakeNum,TakeDate) Values ('" & .JobID & "','5','" & myJobOpen_.UnderStaffID & "','" & txtSaleStaff.Text & "','" & Trim(tmpVal) & "','" & FormatYMD_to_DMY(TodayDate, "/", "/") & "')"
                            DBConnExc "Update TbContainer Set ConNum = '" & CDbl(Num) - CDbl(Trim(tmpVal)) & "' WHERE conid='5' "
                       End If
                    SetGrCon

            Case 16    'ถุงดำ
                  .BagBack = Trim(tmpVal)
                    tSql = "SELECT * FROM TbContainerTake " & vbCrLf
                    tSql = tSql & "where jobid='" & .JobID & "' and conid='6' "
                    Set rs = GetRS(tSql, adOpenForwardOnly, adLockOptimistic)
                    
                    Sql = "SELECT * FROM TbContainer " & vbCrLf
                    Sql = Sql & "where conid='6' "
                    Set RS1 = GetRS(Sql, adOpenForwardOnly, adLockOptimistic)
                    Num = RS1!ConNum
                    
                        If rs.RecordCount > 0 Then
                            OldN = rs!TakeNum
                            If OldN > Trim(tmpVal) Then
                                OldN = CDbl(OldN) - CDbl(Trim(tmpVal))
                                DBConnExc "Update TbContainer Set ConNum = '" & CDbl(Num) + CDbl(OldN) & "' WHERE conid='6' "
                            Else
                                OldN = CDbl(Trim(tmpVal)) - CDbl(OldN)
                                DBConnExc "Update TbContainer Set ConNum = '" & CDbl(Num) - CDbl(OldN) & "' WHERE conid='6' "
                            End If
                            DBConnExc "Update TbContainerTake Set TakeNum = '" & Trim(tmpVal) & "',LastUpdate='" & FormatYMD_to_DMY(TodayDate, "/", "/") & "' ,LastUser='" & Trim(CurrentUser.StaffFName & " " & CurrentUser.StaffLName) & "' WHERE JobID='" & .JobID & "' and conid='6' "
                        Else
                            DBConnExc "Insert into TbContainerTake(JobID, ConID, SaleID,SaleName,TakeNum,TakeDate) Values ('" & .JobID & "','6','" & myJobOpen_.UnderStaffID & "','" & txtSaleStaff.Text & "','" & Trim(tmpVal) & "','" & FormatYMD_to_DMY(TodayDate, "/", "/") & "')"
                            DBConnExc "Update TbContainer Set ConNum = '" & CDbl(Num) - CDbl(Trim(tmpVal)) & "' WHERE conid='6' "
                       End If
                    SetGrCon

            Case 17             'แล็ค
                    .Lack = Trim(tmpVal)
                    tSql = "SELECT * FROM TbContainerTake " & vbCrLf
                    tSql = tSql & "where jobid='" & .JobID & "' and conid='3' "
                    Set rs = GetRS(tSql, adOpenForwardOnly, adLockOptimistic)
                    
                    Sql = "SELECT * FROM TbContainer " & vbCrLf
                    Sql = Sql & "where conid='3' "
                    Set RS1 = GetRS(Sql, adOpenForwardOnly, adLockOptimistic)
                    Num = RS1!ConNum
                    
                        If rs.RecordCount > 0 Then
                            OldN = rs!TakeNum
                            If OldN > Trim(tmpVal) Then
                                OldN = CDbl(OldN) - CDbl(Trim(tmpVal))
                                DBConnExc "Update TbContainer Set ConNum = '" & CDbl(Num) + CDbl(OldN) & "' WHERE conid='3' "
                            Else
                                OldN = CDbl(Trim(tmpVal)) - CDbl(OldN)
                                DBConnExc "Update TbContainer Set ConNum = '" & CDbl(Num) - CDbl(OldN) & "' WHERE conid='3' "
                            End If
                            DBConnExc "Update TbContainerTake Set TakeNum = '" & Trim(tmpVal) & "',LastUpdate='" & FormatYMD_to_DMY(TodayDate, "/", "/") & "' ,LastUser='" & Trim(CurrentUser.StaffFName & " " & CurrentUser.StaffLName) & "' WHERE JobID='" & .JobID & "' and conid='3' "
                        Else
                            DBConnExc "Insert into TbContainerTake(JobID, ConID, SaleID,SaleName,TakeNum,TakeDate) Values ('" & .JobID & "','3','" & myJobOpen_.UnderStaffID & "','" & txtSaleStaff.Text & "','" & Trim(tmpVal) & "','" & FormatYMD_to_DMY(TodayDate, "/", "/") & "')"
                            DBConnExc "Update TbContainer Set ConNum = '" & CDbl(Num) - CDbl(Trim(tmpVal)) & "' WHERE conid='3' "
                       End If
                    SetGrCon
            
            Case 18         'กล่องกระดาษ
                    .PaperBox = Trim(tmpVal)
                    tSql = "SELECT * FROM TbContainerTake " & vbCrLf
                    tSql = tSql & "where jobid='" & .JobID & "' and conid='7' "
                    Set rs = GetRS(tSql, adOpenForwardOnly, adLockOptimistic)
                    
                    Sql = "SELECT * FROM TbContainer " & vbCrLf
                    Sql = Sql & "where conid='7' "
                    Set RS1 = GetRS(Sql, adOpenForwardOnly, adLockOptimistic)
                    Num = RS1!ConNum
                    
                        If rs.RecordCount > 0 Then
                            OldN = rs!TakeNum
                            If OldN > Trim(tmpVal) Then
                                OldN = CDbl(OldN) - CDbl(Trim(tmpVal))
                                DBConnExc "Update TbContainer Set ConNum = '" & CDbl(Num) + CDbl(OldN) & "' WHERE conid='7' "
                            Else
                                OldN = CDbl(Trim(tmpVal)) - CDbl(OldN)
                                DBConnExc "Update TbContainer Set ConNum = '" & CDbl(Num) - CDbl(OldN) & "' WHERE conid='7' "
                            End If
                            DBConnExc "Update TbContainerTake Set TakeNum = '" & Trim(tmpVal) & "',LastUpdate='" & FormatYMD_to_DMY(TodayDate, "/", "/") & "' ,LastUser='" & Trim(CurrentUser.StaffFName & " " & CurrentUser.StaffLName) & "' WHERE JobID='" & .JobID & "' and conid='7' "
                        Else
                            DBConnExc "Insert into TbContainerTake(JobID, ConID, SaleID,SaleName,TakeNum,TakeDate) Values ('" & .JobID & "','7','" & myJobOpen_.UnderStaffID & "','" & txtSaleStaff.Text & "','" & Trim(tmpVal) & "','" & FormatYMD_to_DMY(TodayDate, "/", "/") & "')"
                            DBConnExc "Update TbContainer Set ConNum = '" & CDbl(Num) - CDbl(Trim(tmpVal)) & "' WHERE conid='7' "
                       End If
                    SetGrCon
            
            Case 19             'เบ้าตัดปาก
                    .BoutS = Trim(tmpVal)
                    tSql = "SELECT * FROM TbContainerTake " & vbCrLf
                    tSql = tSql & "where jobid='" & .JobID & "' and conid='8' "
                    Set rs = GetRS(tSql, adOpenForwardOnly, adLockOptimistic)
                    
                    Sql = "SELECT * FROM TbContainer " & vbCrLf
                    Sql = Sql & "where conid='8' "
                    Set RS1 = GetRS(Sql, adOpenForwardOnly, adLockOptimistic)
                    Num = RS1!ConNum
                    
                        If rs.RecordCount > 0 Then
                            OldN = rs!TakeNum
                            If OldN > Trim(tmpVal) Then
                                OldN = CDbl(OldN) - CDbl(Trim(tmpVal))
                                DBConnExc "Update TbContainer Set ConNum = '" & CDbl(Num) + CDbl(OldN) & "' WHERE conid='8' "
                            Else
                                OldN = CDbl(Trim(tmpVal)) - CDbl(OldN)
                                DBConnExc "Update TbContainer Set ConNum = '" & CDbl(Num) - CDbl(OldN) & "' WHERE conid='8' "
                            End If
                            DBConnExc "Update TbContainerTake Set TakeNum = '" & Trim(tmpVal) & "',LastUpdate='" & FormatYMD_to_DMY(TodayDate, "/", "/") & "' ,LastUser='" & Trim(CurrentUser.StaffFName & " " & CurrentUser.StaffLName) & "' WHERE JobID='" & .JobID & "' and conid='8' "
                        Else
                            DBConnExc "Insert into TbContainerTake(JobID, ConID, SaleID,SaleName,TakeNum,TakeDate) Values ('" & .JobID & "','8','" & myJobOpen_.UnderStaffID & "','" & txtSaleStaff.Text & "','" & Trim(tmpVal) & "','" & FormatYMD_to_DMY(TodayDate, "/", "/") & "')"
                            DBConnExc "Update TbContainer Set ConNum = '" & CDbl(Num) - CDbl(Trim(tmpVal)) & "' WHERE conid='8' "
                       End If
                    SetGrCon

            Case 20
                .Remark = Trim(tmpVal)
        End Select
        myJobOpen_.isChanged = True
    End With
    Set tmpDet = Nothing
End Sub

Private Sub grCar_BeforeEdit(ByVal row As Long, ByVal col As Long, Cancel As Boolean)
Dim tmpCar As BWGJobDataCarUsed
    If (col = 3 Or col = 4) And grCar.TextMatrix(row, col) = "" Then Exit Sub
    If (col = 3 Or col = 4) And IsNumeric(grCar.TextMatrix(row, col)) = False Then
        MsgBox "กรุณากรอกข้อมูลเป็นตัวเลขเท่านั้น", vbExclamation
        grCar.TextMatrix(row, col) = ""
        Cancel = True
        Exit Sub
    ElseIf col = 3 Then
        Set tmpCar = myJobOpen_.Cars("'" & Trim(grCar.TextMatrix(row, 0)) & "'")
        If tmpCar.isTrailer And CLng(grCar.TextMatrix(row, col)) Mod 2 <> 0 Then
            MsgBox "ท่านเลือกรถพ่วงต้องกรอกข้อมูลจำนวนเที่ยวเป็นเลขคู่เท่านั้น", vbExclamation
            grCar.TextMatrix(row, col) = ""
            Cancel = True
            Exit Sub
        End If
    End If
   
End Sub

Private Sub grCar_Click()
    Select Case grCar.col
        Case 11
            grCar.Editable = flexEDKbdMouse
        Case Else
            grCar.Editable = flexEDKbd
    End Select
End Sub

Private Sub grCar_DblClick()
Dim x%, tmpKey$, tmpCarTypeID$
Dim isTrailer As Boolean, carKey$, tmpQuotationID As String
Dim carNo As String
Dim tmpS As BWGSearchManager
Dim colSearchResult_ As Collection
tmpQuotationID = ""
    With grCar
            
                If txtJobOpenNo.Text <> "== AUTO ==" Then
                    MsgBox "ไม่สามารถเปลี่ยนประเภทรถได้ ถ้าต้องการเปลี่ยนต้องลง Order ใหม่ อย่ามั่ว +!", vbInformation
                    Exit Sub
                End If
            If Trim(.TextMatrix(.RowSel - 1, 2)) <> "" And .col = 2 And Trim(.TextMatrix(.RowSel, 0)) <> "" Then
            
                carNo = .TextMatrix(.RowSel, 1)
                carKey = Trim(.TextMatrix(.RowSel, 0))
                With grWasteData.getGridObj
                    For x = 1 To .Rows - 1
                        If IsNumeric(.TextMatrix(x, 8)) Then
                            If carNo = CInt(.TextMatrix(x, 8)) Then
                                tmpKey = Trim(.TextMatrix(x, 0))
                                Exit For
                            End If
                        End If
                    Next
                End With
                If isExist(myJobOpen_.JobDetails, "'" & tmpKey & "'") Then
                    tmpCarTypeID = Trim(myJobOpen_.JobDetails("'" & tmpKey & "'").QuoTruckTypeID)
                    'tmpQuotationID = Trim(myJobOpen_.JobDetails("'" & tmpKey & "'").QuotationID)
                    If Trim(myJobOpen_.Cars("'" & carKey & "'").TruckSubTypeID) = "" Then
                        isTrailer = Trim(myJobOpen_.JobDetails("'" & tmpKey & "'").isTrailer)
                    Else
                        isTrailer = Trim(myJobOpen_.Cars("'" & carKey & "'").isTrailer)
                    End If
                Else
                    tmpCarTypeID = ""
                End If
                For x = 1 To myJobOpen_.JobDetails.Count
                    If tmpQuotationID = "" Then
                    tmpQuotationID = "'" & myJobOpen_.JobDetails(x).QuotationID & "'"
                    Else
                     tmpQuotationID = tmpQuotationID & ",'" & myJobOpen_.JobDetails(x).QuotationID & "'"
                    End If
                Next
                Set colSearchResult_ = Nothing
                Set colSearchResult_ = New Collection
                Set tmpS = New BWGSearchManager
                Set colSearchResult_ = tmpS.QuotationTrucSubTypekRateSearch("QuotationID in (" & tmpQuotationID & ")")
                Set tmpS = Nothing
                If colSearchResult_.Count > 0 Then
                  Load frmCarQuotationSearch
                    With frmCarQuotationSearch
                        .SetCallerForm Me, tmpQuotationID, carNo
                        .Show 1
                    End With
                Else
                   Load frmCarTypeSearch
                    With frmCarTypeSearch
                        .SetCallerForm Me, tmpCarTypeID, , isTrailer
                        .Show 1
                    End With
                End If
'                Load frmCarTypeSearch
'                With frmCarTypeSearch
'                    .SetCallerForm Me, tmpCarTypeID, , isTrailer
'                    .Show 1
'                End With
            End If
    End With
End Sub

Private Sub grCar_KeyDown(KeyCode As Integer, Shift As Integer)
Dim x%, carInUsed As Boolean, isTrailer As Boolean
Dim carNo As String, tmpKey As String, tmpCarTypeID As String, carKey As String, tmpQuotationID As String
Dim tmpS As BWGSearchManager
Dim colSearchResult_ As Collection
    With grCar
        If Trim(.TextMatrix(.row, 1)) = "" Or IsNumeric(.TextMatrix(.row, 1)) = False Then
            Exit Sub
        End If
        'If Trim(.TextMatrix(.row - 1, 2)) <> "" And Trim(.TextMatrix(.row, 2)) = "" And .col = 2 And Trim(.TextMatrix(.row, 0)) <> "" Then
        carNo = .TextMatrix(.row, 1)
        carKey = Trim(.TextMatrix(.row, 0))
        If KeyCode = 46 Then
            With grWasteData.getGridObj
                For x = 1 To .Rows - 1
                    If IsNumeric(.TextMatrix(x, 8)) Then
                        If carNo = CInt(.TextMatrix(x, 8)) Then
                            MsgBox "ไม่สามารถลบรถคันนี้ออกจากการเปิดงานได้เพราะมีการเรียกใช้งานอยู่", vbExclamation
                            Exit Sub
                        End If
                    End If
                Next
            End With
            If MsgBox("ท่านต้องการลบรถที่ใช้ขนกากของเสียคันนี้ออกจากการเปิดงานครั้งนี้หรือไม่?", vbQuestion + vbYesNo) = vbYes Then
                tmpKey = "'" & Trim(.TextMatrix(.row, 0)) & "'"
                myJobOpen_.Cars.Remove tmpKey
                .RemoveItem .row
            End If
        Else
            If Trim(.TextMatrix(.row - 1, 2)) <> "" And .col = 2 And Trim(.TextMatrix(.row, 0)) <> "" Then
                With grWasteData.getGridObj
                    For x = 1 To .Rows - 1
                        If IsNumeric(.TextMatrix(x, 8)) Then
                            If carNo = CInt(.TextMatrix(x, 8)) Then
                                tmpKey = Trim(.TextMatrix(x, 0))
                                Exit For
                            End If
                        End If
                    Next
                End With
                If isExist(myJobOpen_.JobDetails, "'" & tmpKey & "'") Then
                    tmpCarTypeID = Trim(myJobOpen_.JobDetails("'" & tmpKey & "'").QuoTruckTypeID)
                    tmpQuotationID = Trim(myJobOpen_.JobDetails("'" & tmpKey & "'").QuotationID)
                    If Trim(myJobOpen_.Cars("'" & carKey & "'").TruckSubTypeID) = "" Then
                        isTrailer = Trim(myJobOpen_.JobDetails("'" & tmpKey & "'").isTrailer)
                    Else
                        isTrailer = Trim(myJobOpen_.Cars("'" & carKey & "'").isTrailer)
                    End If
                Else
                    tmpCarTypeID = ""
                End If
               Set colSearchResult_ = Nothing
                Set colSearchResult_ = New Collection
                Set tmpS = New BWGSearchManager
                Set colSearchResult_ = tmpS.QuotationTruckRateSearch("QuotationID='" & tmpQuotationID & "'")
                Set tmpS = Nothing
                If colSearchResult_.Count > 0 Then
                  Load frmCarQuotationSearch
                    With frmCarQuotationSearch
                        .SetCallerForm Me, tmpQuotationID, carNo
                        .Show 1
                    End With
                Else
                 MsgBox "กรุณาเลือกใบเสนอราคาแบบใหม่", vbExclamation
'                   Load frmCarTypeSearch
'                    With frmCarTypeSearch
'                        .SetCallerForm Me, tmpCarTypeID, , isTrailer
'                        .Show 1
'                    End With
                End If
             
            End If
        End If
    End With
End Sub

Private Sub grCar_RowColChange()
    On Error Resume Next
    Dim tmpKey As String
    Dim tmpCar As BWGJobDataCarUsed
    Dim x%, ColEnable As Boolean
    tmpKey = Trim(grCar.TextMatrix(grCar.RowSel, 0))
   grTimeTable.ClearAllData
    grTimeTable.getGridObj.Rows = 2
            
    If isExist(myJobOpen_.Cars, "'" & Trim(tmpKey) & "'") Then
        Set curCar_ = myJobOpen_.Cars("'" & Trim(tmpKey) & "'")
    Else
        Set curCar_ = Nothing
    End If
    
    If Trim(grCar.TextMatrix(grCar.RowSel, 0)) = "" Then
' ** ปิดการแก้ไขคอลัมน์
'        grCar.ColEnabled(grCar.ColSel) = False
        ColEnable = False
        Label3.Caption = "ตารางเดินรถ"
    Else
        If isExist(myJobOpen_.Cars, "'" & tmpKey & "'") Then
            Set tmpCar = myJobOpen_.Cars("'" & tmpKey & "'")
            Label3.Caption = "ตารางเดินรถ [คันที่ " & tmpCar.carNo & "]"
            If tmpCar.TimeTables.Count = 0 Then
                grTimeTable.getGridObj.Rows = 2
            Else
                grTimeTable.getGridObj.Rows = tmpCar.TimeTables.Count
            End If
            PopulateTimeTalbe tmpCar
            For x = 1 To tmpCar.TimeTables.Count
                grTimeTable.getGridObj.TextMatrix(x, 0) = tmpCar.TimeTables(x).JobDataCarID & "_" & tmpCar.TimeTables(x).TripNo
                grTimeTable.getGridObj.TextMatrix(x, 1) = tmpCar.TimeTables(x).TripNo
                grTimeTable.getGridObj.TextMatrix(x, 2) = tmpCar.TimeTables(x).StartTime
            Next
            Set tmpCar = Nothing
        End If
        If grCar.ColSel > 2 Then
            'grCar.ColEnabled(Col) = True And cmdSave.Enabled
            ColEnable = True And cmdSave.Enabled
        Else
            'grCar.ColEnabled(Col) = False
            ColEnable = False
        End If
    End If
End Sub

Private Sub grCar_Validate(Cancel As Boolean)
'If grCar.TextMatrix(Row, 2) = "รถ Lugger Box, รถ Lugger Box" Or grCar.TextMatrix(Row, 2) = "รถ Lugger Box (พ่วง), รถ Lugger Box" Then
'             If grCar.TextMatrix(Row, 7) <> "LUG" Then
'                 MsgBox "ท่านต้องระบุประเภท Box ให้ตรงกับประเภทรถ", vbExclamation
'                grCar.TextMatrix(Row, 7) = ""
'             End If
'         End If
End Sub

Private Sub grTimeTable_RolColChanged(row As Integer, col As Integer)
    If col = 1 Then
        grTimeTable.ColEnabled(col) = False
    ElseIf col = 2 Then
        grTimeTable.ColEnabled(col) = True And cmdSave.Enabled
    End If
End Sub

Private Sub grTimeTable_UpdateValue(row As Integer, col As Integer)
Dim tmpDet As BWGJobDataTimeTable
Dim tmpKey As String
Dim tmpVal As String, x%
    If curCar_ Is Nothing Then Exit Sub
    tmpVal = Trim(grTimeTable.getGridObj.TextMatrix(row, col))
    Select Case col
        Case 2
            If curCar_.GoSameTime = "Y" Then
                With grTimeTable.getGridObj
                    For x = 1 To .Rows - 1
                        tmpKey = Trim(grTimeTable.getGridObj.TextMatrix(x, 0))
                        If isExist(curCar_.TimeTables, "'" & tmpKey & "'") = False Then
                            Set tmpDet = New BWGJobDataTimeTable
                            tmpDet.JobDataCarID = curCar_.id
                            tmpDet.TripNo = grTimeTable.getGridObj.TextMatrix(x, 1)
                            tmpDet.WorkDate = myJobOpen_.WorkDate
                            tmpDet.WorkerCount = curCar_.WorkerPerCar
                            grTimeTable.getGridObj.TextMatrix(x, 0) = tmpDet.JobDataCarID & "_" & Trim("" & tmpDet.TripNo)
                            curCar_.TimeTables.Add tmpDet, "'" & tmpDet.JobDataCarID & "_" & Trim("" & tmpDet.TripNo) & "'"
                        Else
                            Set tmpDet = curCar_.TimeTables("'" & tmpKey & "'")
                        End If
                        tmpDet.StartTime = tmpVal
                        grTimeTable.getGridObj.TextMatrix(x, 2) = Trim(tmpVal)
                        Set tmpDet = Nothing
                    Next
                End With
            Else
                tmpKey = Trim(grTimeTable.getGridObj.TextMatrix(row, 0))
                If isExist(curCar_.TimeTables, "'" & tmpKey & "'") = False Then
                    Set tmpDet = New BWGJobDataTimeTable
                    tmpDet.JobDataCarID = curCar_.id
                    tmpDet.TripNo = grTimeTable.getGridObj.TextMatrix(row, 1)
                    tmpDet.WorkDate = myJobOpen_.WorkDate
                    grTimeTable.getGridObj.TextMatrix(row, 0) = tmpDet.JobDataCarID & "_" & Trim("" & tmpDet.TripNo)
                    curCar_.TimeTables.Add tmpDet, "'" & tmpDet.JobDataCarID & "_" & Trim("" & tmpDet.TripNo) & "'"
                Else
                    Set tmpDet = curCar_.TimeTables("'" & tmpKey & "'")
                End If
                tmpDet.StartTime = tmpVal
            End If
    End Select
    myJobOpen_.isChanged = True
    Set tmpDet = Nothing
End Sub

Private Sub grWasteData_BeforeUpdateValue(NewValue As String, row As Integer, col As Integer, Cancel As Boolean)
    If (col = 3 Or col = 6) And IsNumeric(NewValue) = False Then
        MsgBox "กรุณาใส่ข้อมูลเป็นตัวเลขเท่านั้น", vbExclamation
        Cancel = True
    End If
End Sub

Private Sub grWasteData_DblClick(row As Integer, col As Integer)
    With grWasteData.getGridObj
        If col = 7 And Trim(.TextMatrix(.row, 0)) <> "" Then
            Load frmTSDFSearch
            With frmTSDFSearch
                .getCriteria "*"
                .SetCallerForm Me
                .Show 1
            End With
        End If
    End With
End Sub

Private Sub grWasteData_KeyDown(KeyCode As Integer, Shift As Integer)
    With grWasteData.getGridObj
        If .col = 7 And Trim(.TextMatrix(.row, 0)) <> "" And KeyCode <> 13 Then
            Load frmTSDFSearch
            With frmTSDFSearch
                .getCriteria "*"
                .SetCallerForm Me
                .Show 1
            End With
        End If
    End With
End Sub

Private Sub grWasteData_RolColChanged(row As Integer, col As Integer)
Dim x%, maxCarNo%
    With grWasteData.getGridObj
        If Trim(.TextMatrix(row, 0)) = "" Then
            grWasteData.ColEnabled(col) = False
        Else
            If col = 1 Or col = 2 Or col = 7 Then
                grWasteData.ColEnabled(col) = False
            Else
                grWasteData.ColEnabled(col) = True And cmdSave.Enabled
                If col = 4 Then
                    With grWasteData.getColObject(4)
                        .Clear
                        For x = 1 To colUnit_.Count
                            .AddItem colUnit_(x).UnitName
                        Next
                    End With
                ElseIf col = 8 Then
                    maxCarNo = 1
                    For x = 1 To .Rows - 1
                        If IsNumeric(.TextMatrix(x, 8)) Then
                            If maxCarNo < CInt(.TextMatrix(x, 8)) Then
                                maxCarNo = CInt(.TextMatrix(x, 8))
                            End If
                        End If
                    Next
                    With grWasteData.getColObject(8)
                        .Clear
                        For x = 1 To maxCarNo
                            .AddItem x
                        Next
                    End With
                ElseIf col = 12 Then
                    With grWasteData.getColObject(12)
                        .Clear
                        .AddItem "New"
                        grWasteData.ColEnabled(col) = True
                    End With
                     ElseIf col = 13 Then
                    With grWasteData.getColObject(13)
                        .Clear
                        .AddItem "New"
                        grWasteData.ColEnabled(col) = True
                    End With
                End If
            End If
        End If
    End With
End Sub

Private Sub grWasteData_UpdateValue(row As Integer, col As Integer)
Dim tmpDet As BWGJobDataDetail
Dim tmpKey As String
Dim tmpVal As String
Dim x%, tmpCar As BWGJobDataCarUsed, hasThisCar As Boolean
    tmpKey = Trim(grWasteData.getGridObj.TextMatrix(row, 0))
    If isExist(myJobOpen_.JobDetails, "'" & tmpKey & "'") = False Then
        Set tmpDet = New BWGJobDataDetail
        tmpDet.id = GetGUID
        tmpDet.JobID = myJobOpen_.id
        grWasteData.getGridObj.TextMatrix(row, 0) = tmpDet.id
        myJobOpen_.JobDetails.Add tmpDet, "'" & tmpDet.id & "'"
    Else
        Set tmpDet = myJobOpen_.JobDetails("'" & tmpKey & "'")
    End If
    tmpVal = Trim(grWasteData.getGridObj.TextMatrix(row, col))
    tmpDet.RowID = row
    With tmpDet
        Select Case col
            Case 3
                If IsNumeric(tmpVal) Then
                    .EstWasteQty = tmpVal
                Else
                    .EstWasteQty = 0
                End If
            Case 4
                If grWasteData.getGridObj.TextMatrix(row, 4) = "" Then
                    .WasteQtyUnitID = ""
                    .WasteQtyUnitName = ""
                Else
                    .WasteQtyUnitID = colUnit_(grWasteData.getColObject(4).ListIndex + 1).id
                    .WasteQtyUnitName = grWasteData.getColObject(4).Text
                End If
            Case 5
                .ContainerTypeDesc = Trim(tmpVal)
            Case 6
                If IsNumeric(tmpVal) Then
                    .ContainerCount = tmpVal
                Else
                    .ContainerCount = 0
                End If
            Case 7
                .DestinationDesc = Trim(tmpVal)
            Case 8
                If IsNumeric(tmpVal) Then
                    .carNo = tmpVal
                    
                End If
                If .carNo > 0 Then
                SetCar .carNo, .isTrailer
'                    hasThisCar = False
'                    For x = 1 To myJobOpen_.Cars.Count
'                        If .carno = myJobOpen_.Cars(x).carno Then
'                            hasThisCar = True
'                            Exit For
'                        End If
'                    Next
'                    If hasThisCar = False Then
'                        Set tmpCar = New BWGJobDataCarUsed
'                        tmpCar.id = GetGUID
'                        tmpCar.carno = .carno
'                        'จุ๊เพิ่ม
'                        tmpCar.TakePhoto = myJobOpen_.ServicePhoto
'                        tmpCar.JobID = myJobOpen_.id
'                        tmpCar.isTrailer = .isTrailer
'                        myJobOpen_.Cars.Add tmpCar, "'" & tmpCar.id & "'"
'                        grCar.TextMatrix(myJobOpen_.Cars.Count, 0) = tmpCar.id
'                        grCar.TextMatrix(myJobOpen_.Cars.Count, 1) = tmpCar.carno
'                        If myJobOpen_.ServicePhoto = "Y" Then
'                          grCar.TextMatrix(myJobOpen_.Cars.Count, 6) = 1
'                        ElseIf myJobOpen_.ServicePhoto = "N" Then
'                          grCar.TextMatrix(myJobOpen_.Cars.Count, 6) = 2
'                        ElseIf myJobOpen_.ServicePhoto = "S" Then
'                          grCar.TextMatrix(myJobOpen_.Cars.Count, 6) = 3
'                        End If
                        
                      
                    'End If
                End If
            Case 9
                .Remdesc = Trim(tmpVal)
            Case 12
                .IsNewWaste = Trim(tmpVal)
            Case 13
                .IsNewWasteCR = Trim(tmpVal)
        End Select
        myJobOpen_.isChanged = True
    End With
    Set tmpDet = Nothing
End Sub

Private Sub rdoComBME_Validate(Index As Integer, Cancel As Boolean)
        If rdoComBME(0).Value = True Then
            myJobOpen_.ComBme = "Y"
        Else
            myJobOpen_.ComBme = "N"
        End If
End Sub

Private Sub TxtBagBack_Validate(Cancel As Boolean)
Dim tmpDet As BWGJobDataCarUsed
Dim tmpKey As String, BoxCheck As Integer
Dim tmpVal As String, OldValue As String
    tmpKey = Trim(grCar.TextMatrix(1, 0))
    If isExist(myJobOpen_.Cars, "'" & tmpKey & "'") = False Then
        Set tmpDet = New BWGJobDataCarUsed
        tmpDet.JobID = myJobOpen_.id
    Else
        Set tmpDet = myJobOpen_.Cars("'" & tmpKey & "'")
    End If
    tmpDet.RowID = 1
    With tmpDet
        If Trim(TxtBagBack.Text) <> "" Then
            .BagBack = Trim(TxtBagBack.Text)
            myJobOpen_.isChanged = True
        End If
    End With

End Sub

Private Sub TxtBagTypeP_Validate(Cancel As Boolean)
Dim tmpDet As BWGJobDataCarUsed
Dim tmpKey As String, BoxCheck As Integer
Dim tmpVal As String, OldValue As String
    tmpKey = Trim(grCar.TextMatrix(1, 0))
    If isExist(myJobOpen_.Cars, "'" & tmpKey & "'") = False Then
        Set tmpDet = New BWGJobDataCarUsed
        tmpDet.JobID = myJobOpen_.id
    Else
        Set tmpDet = myJobOpen_.Cars("'" & tmpKey & "'")
    End If
    tmpDet.RowID = 1
    With tmpDet
        If Trim(TxtBagTypeP.Text) <> "" Then
            .BagType = "P"   'ถุงเปิดปาก
            .BagCount = Trim(TxtBagTypeP.Text)
            .BagPrice = .BagCount * 120
            .BagPriceV = .BagCount * 100
            myJobOpen_.isChanged = True
        End If
    End With

End Sub

Private Sub TxtBoutCage_Validate(Cancel As Boolean)
Dim tmpDet As BWGJobDataCarUsed
Dim tmpKey As String, BoxCheck As Integer
Dim tmpVal As String, OldValue As String
    tmpKey = Trim(grCar.TextMatrix(1, 0))
    If isExist(myJobOpen_.Cars, "'" & tmpKey & "'") = False Then
        Set tmpDet = New BWGJobDataCarUsed
        tmpDet.JobID = myJobOpen_.id
    Else
        Set tmpDet = myJobOpen_.Cars("'" & tmpKey & "'")
    End If
    tmpDet.RowID = 1
    With tmpDet
    If Trim(txtBoutCage.Text) <> "" Then
        .BoutCage = Trim(txtBoutCage.Text)
        myJobOpen_.isChanged = True
    End If
    End With

End Sub

Private Sub TxtBoutS_Validate(Cancel As Boolean)
Dim tmpDet As BWGJobDataCarUsed
Dim tmpKey As String, BoxCheck As Integer
Dim tmpVal As String, OldValue As String
    tmpKey = Trim(grCar.TextMatrix(1, 0))
    If isExist(myJobOpen_.Cars, "'" & tmpKey & "'") = False Then
        Set tmpDet = New BWGJobDataCarUsed
        tmpDet.JobID = myJobOpen_.id
    Else
        Set tmpDet = myJobOpen_.Cars("'" & tmpKey & "'")
    End If
    tmpDet.RowID = 1
    With tmpDet
    If Trim(txtBoutS.Text) <> "" Then
        .BoutS = Trim(txtBoutS.Text)
        myJobOpen_.isChanged = True
    End If
    End With
End Sub

Private Sub txtCallName_Validate(Cancel As Boolean)
    If Trim(txtCallName.Text) <> "" Then
        myJobOpen_.CallerName = txtCallName.Text
    End If
End Sub
Private Sub TxtContainer_Validate(Cancel As Boolean)
Dim tmpDet As BWGJobDataCarUsed
Dim tmpKey As String, BoxCheck As Integer
Dim tmpVal As String, OldValue As String
    tmpKey = Trim(grCar.TextMatrix(1, 0))
    If isExist(myJobOpen_.Cars, "'" & tmpKey & "'") = False Then
        Set tmpDet = New BWGJobDataCarUsed
        tmpDet.JobID = myJobOpen_.id
    Else
        Set tmpDet = myJobOpen_.Cars("'" & tmpKey & "'")
    End If
    tmpDet.RowID = 1
    With tmpDet
    If Trim(TxtContainer.Text) <> "" Then
        .Container = Trim(TxtContainer.Text)
        myJobOpen_.isChanged = True
    End If
    End With
End Sub
Public Sub SetTxtCustomer()
Dim tmpCustID As String
    If curCustomer_ Is Nothing Then Set curCustomer_ = New BWGCustomer
    If Trim(txtCustomer.Text) = "" And Trim(curCustomer_.CustomerName) <> Trim(txtCustomer.Text) Then
        Set curCustomer_ = Nothing
        Set curCustomer_ = New BWGCustomer
        With myJobOpen_
            .CustomerID = ""
            .CustomerName = ""
        End With
        lblDTeam.Visible = False
    Else
        tmpCustID = myJobOpen_.CustomerID
        If Trim(curCustomer_.CustomerName) <> Trim(txtCustomer.Text) Then
            Set curCustomer_ = Nothing
            Set curCustomer_ = New BWGCustomer
            With myJobOpen_
                .CustomerID = ""
                .CustomerName = ""
            End With
            lblDTeam.Visible = False
            Load frmCustomerSearch
            With frmCustomerSearch
                .CStatus = "N"
                .getCriteria (txtCustomer.Text)
                .GetCriteriaOptonal " AND CustGroupID <> '488F2E5C057F46869D1A56185B11D644'"
                .SetCallerForm Me
                .Show 1
            End With
            If curCustomer_.id <> tmpCustID Then
                grWasteData.ClearAllData
                ClearNotSaveWasteAdd
            End If
            If myJobOpen_.CustomerID = "" Then txtCustomer.Text = ""
            lblDTeam.Visible = curCustomer_.DTeam
        End If
    End If
End Sub


Private Sub txtCostTrans_Validate(Cancel As Boolean)
    myJobOpen_.CostTrans = CDbl(txtCostTrans)
End Sub

Private Sub txtCustomer_Validate(Cancel As Boolean)
            SetTxtCustomer
            PopulateDistrict
            setConfirmOrder
            JobProject
End Sub
Public Sub SetDataCustomer(selCustomer As BWGCustomer, Optional isAssigned As Boolean = True)
On Error GoTo ErrD
Dim x%, tmpDet As BWGCustomerSite
Dim tmpRec As New ADODB.Recordset
Dim sqlStr As String
Dim SqlC As String
Dim TmpC As New ADODB.Recordset
    Set curCustomer_ = Nothing
    Set curCustomer_ = selCustomer
    If curCustomer_ Is Nothing Then Exit Sub
    cboCustSite.Clear
    With curCustomer_
        SqlC = "Select * from tbcustomer WHERE CustomerID='" & .id & "'  and StopTrn='1' "
        Set TmpC = GetRS(SqlC, adOpenForwardOnly, adLockReadOnly)
        If TmpC.RecordCount > 0 Then
            MsgBox "ลูกค้ารายนี้ถูกระงับการขนชั่วคราว กรุณาติดต่อแผนกบัญชีเพื่อสอบถามข้อมูล", vbExclamation, "Waring"
            Exit Sub
        End If
        lblDTeam.Visible = .DTeam
        For x = 1 To .SiteAddresses.Count
            Set tmpDet = .SiteAddresses(x)
            cboCustSite.AddItem Trim(tmpDet.Address & " " & tmpDet.TumbolName & " " & tmpDet.AumphurName & " " & tmpDet.ProvinceName)
            Set tmpDet = Nothing
        Next
        If cboCustSite.ListCount = 1 Then
            cboCustSite.ListIndex = 0
        End If
        sqlStr = "Select FactoryRegisNo from tbWasteDataCR WHERE CustomerID='" & .id & "' Group By FactoryRegisNo"
        Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
        If tmpRec Is Nothing Then Exit Sub
        cboWasteGenNo.Clear
        cboWasteGenNo.AddItem "ไม่ระบุเลขทะเบียนโรงงาน"
        
        lvResult.ListItems.Clear
        For x = 1 To .OtherRemarks.Count
            lvResult.ListItems.Add , "'" & .OtherRemarks(x).id & "'", .OtherRemarks(x).Remdesc
        Next
        
        Do Until tmpRec.EOF
            If Trim("" & tmpRec("FactoryRegisNo")) <> "" Then cboWasteGenNo.AddItem Trim("" & tmpRec("FactoryRegisNo"))
            tmpRec.MoveNext
        Loop
        cboWasteGenNo.ListIndex = 0
        tmpRec.Close
        Set tmpRec = Nothing
        txtCustomer.Text = .CustomerName
        
        If isAssigned Then
            myJobOpen_.CustomerID = .id
            myJobOpen_.CustomerName = .CustomerName
            myJobOpen_.UnderStaffID = .UnderSaleStaffID
            myJobOpen_.UnderStaffName = .UnderSaleName
            If cboCustSite.ListCount = 1 Then
                myJobOpen_.CustomerSiteID = .SiteAddresses(cboCustSite.ListIndex + 1).id
                myJobOpen_.CustomerSiteName = cboCustSite.Text
            End If
        End If
    End With
    Exit Sub
ErrD:
    GetMsg "", vbCritical
    Exit Sub
End Sub

Private Sub txtH_Validate(Cancel As Boolean)
    If IsNumeric(txtH.Text) = False Then
        txtH.Text = Hour(Now)
    End If
    myJobOpen_.WorkTime = Format(Trim(txtH.Text), "00") & ":" & Format(Trim(txtM.Text), "00")
End Sub
Private Sub TxtLack_Validate(Cancel As Boolean)
Dim tmpDet As BWGJobDataCarUsed
Dim tmpKey As String, BoxCheck As Integer
Dim tmpVal As String, OldValue As String
    tmpKey = Trim(grCar.TextMatrix(1, 0))
    If isExist(myJobOpen_.Cars, "'" & tmpKey & "'") = False Then
        Set tmpDet = New BWGJobDataCarUsed
        tmpDet.JobID = myJobOpen_.id
    Else
        Set tmpDet = myJobOpen_.Cars("'" & tmpKey & "'")
    End If
    tmpDet.RowID = 1
    With tmpDet
    If Trim(txtLack.Text) <> "" Then
        .Lack = Trim(txtLack.Text)
        myJobOpen_.isChanged = True
    End If
    End With
End Sub
Private Sub TxtLackGas_Validate(Cancel As Boolean)
Dim tmpDet As BWGJobDataCarUsed
Dim tmpKey As String, BoxCheck As Integer
Dim tmpVal As String, OldValue As String
    tmpKey = Trim(grCar.TextMatrix(1, 0))
    If isExist(myJobOpen_.Cars, "'" & tmpKey & "'") = False Then
        Set tmpDet = New BWGJobDataCarUsed
        tmpDet.JobID = myJobOpen_.id
    Else
        Set tmpDet = myJobOpen_.Cars("'" & tmpKey & "'")
    End If
    tmpDet.RowID = 1
    With tmpDet
    If Trim(TxtLackGas.Text) <> "" Then
        .LackGas = Trim(TxtLackGas.Text)
        myJobOpen_.isChanged = True
    End If
    End With

End Sub

Private Sub TxtLackNew_Validate(Cancel As Boolean)
Dim tmpDet As BWGJobDataCarUsed
Dim tmpKey As String, BoxCheck As Integer
Dim tmpVal As String, OldValue As String
    tmpKey = Trim(grCar.TextMatrix(1, 0))
    If isExist(myJobOpen_.Cars, "'" & tmpKey & "'") = False Then
        Set tmpDet = New BWGJobDataCarUsed
        tmpDet.JobID = myJobOpen_.id
    Else
        Set tmpDet = myJobOpen_.Cars("'" & tmpKey & "'")
    End If
    tmpDet.RowID = 1
    With tmpDet
    If Trim(TxtLackNew.Text) <> "" Then
        .LackNew = Trim(TxtLackNew.Text)
        myJobOpen_.isChanged = True
    End If
    End With

End Sub

Private Sub txtM_Validate(Cancel As Boolean)
    If IsNumeric(txtM.Text) = False Then
        txtM.Text = Minute(Now)
    End If
    myJobOpen_.WorkTime = Format(Trim(txtH.Text), "00") & ":" & Format(Trim(txtM.Text), "00")
End Sub
Private Sub txtPaperBox_Validate(Cancel As Boolean)
Dim tmpDet As BWGJobDataCarUsed
Dim tmpKey As String, BoxCheck As Integer
Dim tmpVal As String, OldValue As String
    tmpKey = Trim(grCar.TextMatrix(1, 0))
    If isExist(myJobOpen_.Cars, "'" & tmpKey & "'") = False Then
        Set tmpDet = New BWGJobDataCarUsed
        tmpDet.JobID = myJobOpen_.id
    Else
        Set tmpDet = myJobOpen_.Cars("'" & tmpKey & "'")
    End If
    tmpDet.RowID = 1
    With tmpDet
    If Trim(txtPaperBox.Text) <> "" Then
        .PaperBox = Trim(txtPaperBox.Text)
        myJobOpen_.isChanged = True
    End If
    End With

End Sub

Private Sub TxtPlastic_Validate(Cancel As Boolean)
Dim tmpDet As BWGJobDataCarUsed
Dim tmpKey As String, BoxCheck As Integer
Dim tmpVal As String, OldValue As String
    tmpKey = Trim(grCar.TextMatrix(1, 0))
    If isExist(myJobOpen_.Cars, "'" & tmpKey & "'") = False Then
        Set tmpDet = New BWGJobDataCarUsed
        tmpDet.JobID = myJobOpen_.id
    Else
        Set tmpDet = myJobOpen_.Cars("'" & tmpKey & "'")
    End If
    tmpDet.RowID = 1
    With tmpDet
    If Trim(TxtPlastic.Text) <> "" Then
        .Plastic = Trim(TxtPlastic.Text)
        myJobOpen_.isChanged = True
    End If
    End With

End Sub

Private Sub txtPlate_Validate(Cancel As Boolean)
Dim tmpDet As BWGJobDataCarUsed
Dim tmpKey As String, BoxCheck As Integer
Dim tmpVal As String, OldValue As String
    tmpKey = Trim(grCar.TextMatrix(1, 0))
    If isExist(myJobOpen_.Cars, "'" & tmpKey & "'") = False Then
        Set tmpDet = New BWGJobDataCarUsed
        tmpDet.JobID = myJobOpen_.id
    Else
        Set tmpDet = myJobOpen_.Cars("'" & tmpKey & "'")
    End If
    tmpDet.RowID = 1
    With tmpDet
    If Trim(txtPlate.Text) <> "" Then
        .Plate = Trim(txtPlate.Text)
        myJobOpen_.isChanged = True
    End If
    End With

End Sub

Private Sub txtPONo_Validate(Cancel As Boolean)
    myJobOpen_.PONo = Trim(txtPONo.Text)
End Sub



Private Sub txtRack_Validate(Cancel As Boolean)
  Dim tmpDet As BWGJobDataCarUsed
    Dim tmpKey As String, BoxCheck As Integer
    Dim tmpVal As String, OldValue As String
        tmpKey = Trim(grCar.TextMatrix(1, 0))
        If isExist(myJobOpen_.Cars, "'" & tmpKey & "'") = False Then
            Set tmpDet = New BWGJobDataCarUsed
            tmpDet.JobID = myJobOpen_.id
        Else
            Set tmpDet = myJobOpen_.Cars("'" & tmpKey & "'")
        End If
        tmpDet.RowID = 1
        With tmpDet
        If Trim(txtRack.Text) <> "" Then
            .Rack = Trim(txtRack.Text)
            myJobOpen_.isChanged = True
        End If
        End With
End Sub

Private Sub TxtRefME_Validate(Cancel As Boolean)
    myJobOpen_.RefME = Trim(TxtRefME.Text)
End Sub

Private Sub txtRem_Validate(Cancel As Boolean)
    myJobOpen_.Remdesc = Trim(txtRem.Text)
End Sub

Private Sub txtSaleStaff_Validate(Cancel As Boolean)
    myJobOpen_.SaleStaffName = txtSaleStaff.Text
End Sub

Private Sub SetEnabledScreen()
Dim JobCancel As Boolean
    JobCancel = myJobOpen_.isCancel
    cmdAddItem.Enabled = Not JobCancel And myJobOpen_.MenifestCountIssued = 0 And Not myJobOpen_.isCompleted
    cmdSave.Enabled = Not JobCancel And myJobOpen_.MenifestCountIssued = 0 And Not myJobOpen_.isCompleted
    Frame6.Enabled = Not JobCancel And myJobOpen_.MenifestCountIssued = 0 And Not myJobOpen_.isCompleted
    Frame7.Enabled = Not JobCancel And myJobOpen_.MenifestCountIssued = 0 And Not myJobOpen_.isCompleted
    cmdCancel.Enabled = Not JobCancel And myJobOpen_.MenifestCountIssued = 0 And Not myJobOpen_.isCompleted
    cmdDone.Enabled = Not JobCancel And myJobOpen_.MenifestCountIssued = 0 And Not myJobOpen_.isCompleted
    lbCancel.Visible = JobCancel Or myJobOpen_.isCompleted
    If JobCancel = True Then
        lbCancel.Caption = "ยกเลิก"
    ElseIf myJobOpen_.isCompleted Then
        lbCancel.Caption = "ปิด Job"
    End If
    If JobCancel Or myJobOpen_.isCompleted Then
        grCar.Editable = flexEDNone
    Else
        grCar.Editable = flexEDKbd
    End If
End Sub

Private Sub ClearScreen()
    If cboCompany.ListCount > 0 Then
        cboCompany.ListIndex = 0
    End If
    lbMenifestCount.Caption = ""
    lvResult.ListItems.Clear
    cmdSave.Enabled = True
    Frame6.Enabled = True
    Frame7.Enabled = True
    cmdCancel.Enabled = True
    cmdAddItem.Enabled = True
    
    txtJobOpenNo.Text = "== AUTO =="
    txtIssuedDate.Text = FormatYMD_to_DMY(TodayDate, "/", "/")
    dtWorkDate.ValueYMD = TodayDate
    txtH.Text = Format(Hour(Now), "00")
    txtM.Text = Format(Minute(Now), "00")
    cboWasteGenNo.ListIndex = -1
    cboType.ListIndex = 0
    dtFromWork.ValueYMD = TodayDate
    txtCustomer.Text = ""
    Set curCustomer_ = Nothing
    cboCustSite.Clear
    txtIssuedStaff.Text = CurrentUser.StaffFName & " " & CurrentUser.StaffLName
    txtSaleStaff.Text = ""
    txtRem.Text = ""
    
    grWasteData.ClearAllData
    grWasteData.getGridObj.Rows = 20
    
    grCar.Clear
    grCar.Rows = 20
    grCar.Editable = flexEDKbd
    
    grTimeTable.ClearAllData
    grTimeTable.getGridObj.Rows = 2
    Call SetGrid
    lblDTeam.Visible = False
End Sub

Private Sub ClearNotSaveWasteAdd()
Dim tmpCol As Collection
    Set tmpCol = myJobOpen_.JobDetails
    Set tmpCol = Nothing
    Set myJobOpen_.JobDetails = tmpCol
End Sub

Public Sub SetDataJobOpen(selJob As BWGJobData)
Dim tmpStr, i As Integer
Dim tmpS As New BWGSearchManager
Dim tmpCol As Collection
     Set myJobOpen_ = selJob
    With myJobOpen_
        txtJobOpenNo.Text = .JobNo
        txtIssuedDate.Text = FormatYMD_to_DMY(.IssuedDate, "/", "/")
        dtWorkDate.ValueYMD = .WorkDate
        dtFromWork.ValueYMD = .FromWork
        If .CallOrd <> "" Then
            cboCaller.Text = .CallOrd
        Else
            cboCaller.ListIndex = 0
        End If
        If .CallerWay <> "" Then
            cboCallWay.Text = .CallerWay
        Else
            cboCallWay.ListIndex = 0
        End If
        Dim sqlPro As String
        Dim tmpPro As ADODB.Recordset
        Set tmpPro = New ADODB.Recordset
        sqlPro = " select * from tbJobProject  Where ProjectID = '" & .ProjectID & "' "
        Set tmpPro = GetRS(sqlPro, adOpenForwardOnly, adLockReadOnly)
        If tmpPro.EOF = False Then
            cbProject.Text = tmpPro!ProjectName & " # " & tmpPro!ProjectID
        End If

        PopulateDistrict

        If .SaleDistrictIDOrder <> "" Then
            Dim tSql As String
            Dim rs As New ADODB.Recordset
                tSql = "SELECT * FROM vw_SaleDistrictSearch" & vbCrLf
                tSql = tSql & "where SaleDistrictID = '" & .SaleDistrictIDOrder & "' "
                Set rs = GetRS(tSql, adOpenForwardOnly, adLockOptimistic)
                    If rs.RecordCount > 0 Then
                        CbOrdCompany.Text = rs!Zonedesc
                    End If
        End If
    If .CancelRemark <> "" Then
        CbRemark.Text = .CancelRemark
    End If
    
        txtCallName.Text = .CallerName
        
        If .isJob = "N" Then
            cboOrderType.ListIndex = 0
        ElseIf .isJob = "Y" Then
            cboOrderType.ListIndex = 1
        ElseIf .isJob = "D" Then
            cboOrderType.ListIndex = 2
        End If
        
        If .CovidWaste = "Y" Then
            ChkCovidWaste.Value = 1
            Else
            ChkCovidWaste.Value = 0
        End If
        
        If Trim("" & .isCase) <> "" Then
            For i = 1 To WasteCases.Count
                If .isCase = WasteCases(i).WasteCaseID Then
                    cboOrderType.Text = WasteCases(i).WasteCase
                    Exit For
                End If
            Next i
        End If
        
        If .PhotoS = "Y" Then
            chkPhotoS.Value = 1
        Else
            chkPhotoS.Value = 0
        End If
        If .ComBme = "Y" Then
            rdoComBME(0).Value = True
        Else
            rdoComBME(1).Value = True
        End If
        If .TruckType = "N" Then
            cboTruckType.ListIndex = 1
        Else
            cboTruckType.ListIndex = 0
        End If
        If IsNumeric(.CostTrans) Then
            txtCostTrans.Text = .CostTrans
        Else
            txtCostTrans.Text = "0"
        End If
        SetTextToCombo CurrentUser.WorkCompanies("'" & .CompanyID & "'").CompanyNameTH, cboCompany
        Index = cboCompany.ListIndex
        If InStr(1, .WorkTime, ":") Then
            tmpStr = Split(.WorkTime, ":")
            txtH.Text = Format(Trim(tmpStr(0)), "00")
            txtM.Text = Format(Trim(tmpStr(1)), "00")
        End If
        Set tmpCol = tmpS.CustomerSearch("CustomerID='" & .CustomerID & "'")
        If tmpCol.Count > 0 Then SetDataCustomer tmpCol(1), False
        SetTextToCombo .CustomerSiteName, cboCustSite
        lbMenifestCount.Caption = .MenifestCountAll & "/" & .MenifestCountIssued
        If .MenifestCountIssued > 0 Then
            dtWorkDate.Enabled = False
        Else
            dtWorkDate.Enabled = True
        End If
        cmdSave.Enabled = Not .MenifestCountIssued > 0
        Frame6.Enabled = Not .MenifestCountIssued > 0
        Frame7.Enabled = Not .MenifestCountIssued > 0
        cmdCancel.Enabled = Not .MenifestCountIssued > 0
        cmdAddItem.Enabled = Not .MenifestCountIssued > 0
        txtIssuedStaff.Text = .IssuedStaffName
        txtPONo.Text = .PONo
        If Trim(.IsChk) <> "Y" Then
            Chk.Value = 0
        Else
            Chk.Value = 1
            Chk.Enabled = False
            Chk.Caption = "ตรวจสอบ โดย " & .UserChk
        End If
        TxtRefME.Text = .RefME
        LbNumber.Caption = .Numrow
        If .OrderType = "N" Then
            cboType.ListIndex = 0
        ElseIf .OrderType = "B" Then
            cboType.ListIndex = 1
        ElseIf .OrderType = "K" Then
            cboType.ListIndex = 2
        ElseIf .OrderType = "P" Then
            cboType.ListIndex = 3
        ElseIf .OrderType = "U" Then
            cboType.ListIndex = 4
        ElseIf .OrderType = "S" Then
            cboType.ListIndex = 5
        ElseIf .OrderType = "W" Then
            cboType.ListIndex = 6
        ElseIf .OrderType = "D" Then
            cboType.ListIndex = 7
        Else
            cboType.ListIndex = 0
        End If
             
        If Trim(.WasteCreateNo) = "" Then
            cboWasteGenNo.ListIndex = 0
        Else
            SetTextToCombo .WasteCreateNo, cboWasteGenNo
        End If
        
        If .ServiceOrderID <> "" Then
            Dim SQLService As String
            Dim RSService As New ADODB.Recordset
            Dim j As Integer
            SQLService = "SELECT orderno+ ' # ' +workdate as orderno,orderid FROM Service_OrderHD where orderid='" & .ServiceOrderID & "' " & vbCrLf
            SQLService = SQLService & " order by orderid"
            Set RSService = GetRS(SQLService, adOpenForwardOnly, adLockOptimistic)
            cboConfirm.Clear
            If RSService.RecordCount > 0 Then
                cboConfirm.Text = RSService!OrderNo
            End If
        Else
            setConfirmOrder
        End If
        
        txtSaleStaff.Text = .SaleStaffName
        txtRem.Text = .Remdesc
        ShowWasteData
        ShowCar
    End With
    SetEnabledScreen
    txtCustomer.Enabled = False
End Sub

Private Sub ShowWasteData()
Dim x%, i%, tmpDet As BWGJobDataDetail
Dim hasThisCar    As Boolean
    With grWasteData.getGridObj
        grWasteData.ClearAllData
        .Rows = myJobOpen_.JobDetails.Count + 20
        For x = 1 To myJobOpen_.JobDetails.Count
            Set tmpDet = myJobOpen_.JobDetails(x)
            .TextMatrix(x, 0) = tmpDet.id
            .TextMatrix(x, 1) = tmpDet.WasteNo
            .TextMatrix(x, 2) = tmpDet.WasteName & " [" & tmpDet.WasteType & "]"
            .TextMatrix(x, 3) = tmpDet.EstWasteQty
            .TextMatrix(x, 4) = tmpDet.WasteQtyUnitName
            .TextMatrix(x, 5) = tmpDet.ContainerTypeDesc
            .TextMatrix(x, 6) = tmpDet.ContainerCount
            .TextMatrix(x, 7) = tmpDet.TSDFRegisNo
            .TextMatrix(x, 8) = tmpDet.carNo
            .TextMatrix(x, 9) = tmpDet.Remdesc
            .TextMatrix(x, 10) = tmpDet.JobTreatmentRate
            .TextMatrix(x, 11) = tmpDet.JobTreatmentUnitID
            .TextMatrix(x, 12) = tmpDet.IsNewWaste
            .TextMatrix(x, 13) = tmpDet.IsNewWasteCR
            SetCar tmpDet.carNo, False
            Set tmpDet = Nothing
        
        Next
    End With
End Sub
Private Sub SetCar(carNo As String, isTrailer As Boolean)
Dim i%, hasThisCar As Boolean
Dim tmpCar As BWGJobDataCarUsed
         hasThisCar = False
                    For i = 1 To myJobOpen_.Cars.Count
                        If carNo = myJobOpen_.Cars(i).carNo Then
                            hasThisCar = True
                            Exit For
                        End If
                    Next
                    If hasThisCar = False Then
                        Set tmpCar = New BWGJobDataCarUsed
                        tmpCar.id = GetGUID
                        tmpCar.carNo = 1
                        'จุ๊เพิ่ม
                        tmpCar.TakePhoto = myJobOpen_.ServicePhoto
                        tmpCar.JobID = myJobOpen_.id
                        tmpCar.isTrailer = isTrailer
                        myJobOpen_.Cars.Add tmpCar, "'" & tmpCar.id & "'"
                        grCar.TextMatrix(myJobOpen_.Cars.Count, 0) = tmpCar.id
                        grCar.TextMatrix(myJobOpen_.Cars.Count, 1) = tmpCar.carNo
                        If myJobOpen_.ServicePhoto = "Y" Then
                          grCar.TextMatrix(myJobOpen_.Cars.Count, 6) = 1
                        ElseIf myJobOpen_.ServicePhoto = "N" Then
                          grCar.TextMatrix(myJobOpen_.Cars.Count, 6) = 2
                        ElseIf myJobOpen_.ServicePhoto = "S" Then
                          grCar.TextMatrix(myJobOpen_.Cars.Count, 6) = 3
                        End If
                        End If
                        
End Sub

Private Sub ShowCar()
Dim x%, tmpDet As BWGJobDataCarUsed
    With grCar

'        grCar.Clear flexClearData, flexClearText
        .Clear
        .Rows = myJobOpen_.Cars.Count + 20

        For x = 1 To myJobOpen_.Cars.Count
            Set tmpDet = myJobOpen_.Cars(x)
            .TextMatrix(x, 0) = tmpDet.id
            .TextMatrix(x, 1) = tmpDet.carNo
            If tmpDet.isTrailer Then
                .TextMatrix(x, 2) = tmpDet.TruckTypeName & " , " & tmpDet.TruckSubTypeName
            Else
                .TextMatrix(x, 2) = tmpDet.TruckTypeName & ", " & tmpDet.TruckSubTypeName
            End If
            .TextMatrix(x, 3) = tmpDet.TruckCount
            .TextMatrix(x, 4) = tmpDet.WorkerPerCar
            If tmpDet.GoSameTime = "Y" Then
                .TextMatrix(x, 5) = "พร้อม"
            ElseIf tmpDet.GoSameTime = "N" Then
                .TextMatrix(x, 5) = "ไม่พร้อม"
            End If
            If tmpDet.TakePhoto = "Y" Then
                .TextMatrix(x, 6) = "ถ่ายรูปทำลาย"
            ElseIf tmpDet.TakePhoto = "N" Then
                .TextMatrix(x, 6) = "ไม่ถ่าย"
            'ElseIf tmpDet.TakePhoto = "W" Then
                '.TextMatrix(x, 6) = "ถ่ายรูปชั่ง + ลงของ"
            ElseIf tmpDet.TakePhoto = "S" Then
                .TextMatrix(x, 6) = "ถ่ายรูปพิเศษ"
'            ElseIf tmpDet.TakePhoto = "B" Then
'                .TextMatrix(x, 6) = "ถ่ายรูปขึ้นสายพาน"
'            ElseIf tmpDet.TakePhoto = "F" Then
'                .TextMatrix(x, 6) = "ถ่ายรูป Boxfeed"
            End If
            .TextMatrix(x, 7) = tmpDet.BoxType
            .TextMatrix(x, 8) = tmpDet.BoxCount
            .TextMatrix(x, 9) = tmpDet.AKPBoxCode
            .TextMatrix(x, 10) = tmpDet.AKPBoxCode2
            Obc1.Caption = tmpDet.AKPBoxCode
            Obc2.Caption = tmpDet.AKPBoxCode2
            .TextMatrix(x, 11) = tmpDet.Tankcount
            .TextMatrix(x, 12) = tmpDet.TankTypeSCount
            .TextMatrix(x, 13) = tmpDet.TankTypeSB
            .TextMatrix(x, 14) = tmpDet.BagCount
            .TextMatrix(x, 15) = tmpDet.Container
            .TextMatrix(x, 16) = tmpDet.BagBack
            .TextMatrix(x, 17) = tmpDet.Lack
            .TextMatrix(x, 18) = tmpDet.PaperBox
            .TextMatrix(x, 19) = tmpDet.BoutS
            .TextMatrix(x, 20) = tmpDet.Remark
            
            TxtContainer.Text = tmpDet.Container
            txtBoutS.Text = tmpDet.BoutS
            txtLack.Text = tmpDet.Lack
            TxtLackNew.Text = tmpDet.LackNew
            TxtLackGas.Text = tmpDet.LackGas
            txtPlate.Text = tmpDet.Plate
            TxtPlastic.Text = tmpDet.Plastic
            TxtTankTypeL.Text = tmpDet.Tankcount
            TxtTankTypeS.Text = tmpDet.TankTypeSCount
            TxtBagTypeP.Text = tmpDet.BagCount
            TxtBagBack.Text = tmpDet.BagBack
            txtPaperBox.Text = tmpDet.PaperBox
            txtBoutCage.Text = tmpDet.BoutCage
            CovidBin.Text = tmpDet.CovidBin
            CovidBin200.Text = tmpDet.CovidBin200
            
            If tmpDet.PhotoOnWeight = "Y" Then
                ChkW.Value = 1
            Else
                ChkW.Value = 0
            End If
            If tmpDet.PhotoOnBelt = "Y" Then
                ChkB.Value = 1
            Else
                ChkB.Value = 0
            End If
            If tmpDet.NewCustJob = "Y" Then
                ChkNewCustJob.Value = 1
            Else
                ChkNewCustJob.Value = 0
            End If
            If tmpDet.NewCustJobBI = "Y" Then
                ChkNewCustJobBI.Value = 1
            Else
                ChkNewCustJobBI.Value = 0
            End If
            If tmpDet.CustPhoto = "Y" Then
                ChkCus.Value = 1
            Else
                ChkCus.Value = 0
            End If
            If tmpDet.EmpTrain = "Y" Then
                ChkEm.Value = 1
            Else
                ChkEm.Value = 0
            End If
            If tmpDet.CarCheck = "Y" Then
                ChkCarCheck.Value = 1
            Else
                ChkCarCheck.Value = 0
            End If
'            If tmpDet.CustCon = "Y" Then
'                ChkC.Value = 1
'            Else
'                ChkC.Value = 0
'            End If
            If tmpDet.CloseWeight = "Y" Then
                ChkCloseWeight.Value = 1
            Else
                ChkCloseWeight.Value = 0
            End If
            If Trim(tmpDet.ConName) <> "" Then
                CboC.Text = tmpDet.ConName
                ChkC.Value = 1
            Else
                CboC.Text = ""
                ChkC.Value = 0
            End If
            
            Set tmpDet = Nothing
        Next
    End With
End Sub

Public Sub AddWasteDataItem(selItem As BWGWasteDataCR)
Dim tmpCol As Collection
Dim x%, isNew As Boolean
Dim tmpID As String
Dim tmpDetail As New BWGJobDataDetail
Dim tmpCar As BWGJobDataCarUsed

    Set tmpCol = myJobOpen_.JobDetails

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
        tmpDetail.JobID = myJobOpen_.id
        tmpDetail.wastedataID = selItem.id
        tmpDetail.WasteEUCode = selItem.EUCode1 & "-" & selItem.EUCode2 & "-" & selItem.EUCode3
        tmpDetail.WasteName = selItem.WasteName
        tmpDetail.WasteNo = selItem.WasteNo
        tmpDetail.carNo = 1
        tmpDetail.WasteType = selItem.WasteType
        tmpDetail.WasteQuotationNo = selItem.QuotationNo
        tmpDetail.isTrailer = selItem.isTrailer
        tmpDetail.TSDFID = selItem.TSDFID
        tmpDetail.TSDFRegisNo = selItem.TSDFNo
        tmpDetail.IsDiscountPriceTrans = selItem.IsDiscountPriceTrans
        
'                tmpCar.ID = GetGUID
'                tmpCar.carNo = tmpDetail.carNo
'                tmpCar.JobID = myJobOpen_.ID
'                tmpCar.isTrailer = tmpDetail.isTrailer
'                myJobOpen_.Cars.Add tmpCar, "'" & tmpCar.ID & "'"
'                grCar.TextMatrix(myJobOpen_.Cars.Count, 0) = tmpCar.ID
'                grCar.TextMatrix(myJobOpen_.Cars.Count, 1) = tmpCar.carNo

        tmpCol.Add tmpDetail, "'" & tmpDetail.id & "'"
    End If
    tmpDetail.WasteQuotationNo = selItem.QuotationNo
    tmpDetail.QuotationID = selItem.QuotationID
    tmpDetail.QuoTruckTypeID = selItem.CarTypeID
    tmpDetail.QuoTruckTypeName = selItem.CarTypeName
    tmpDetail.QuoLaborCount = selItem.StaffCount
    tmpDetail.QuoLaborCharge = selItem.StaffCharge
    tmpDetail.QuoTransportFee = selItem.TransportFee
    tmpDetail.QuoTransUnitID = selItem.TransUnitID
    tmpDetail.QuoTransUnitName = selItem.TransUnitName
    tmpDetail.JobTreatmentRate = selItem.TreatmentFee
    tmpDetail.JobTreatmentUnitID = selItem.TreatmentUnitID
    tmpDetail.QuoTreatmentUnitName = selItem.TreatmentUnitName
    tmpDetail.isPriceIncTrans = selItem.isPriceIncTrans
    tmpDetail.IsDiscountPriceTrans = selItem.IsDiscountPriceTrans
    tmpDetail.QuoMinWeightPerCar = selItem.MinWeightPerCar
    tmpDetail.CanEvap = selItem.CanEvap
    tmpDetail.EvapDesc = selItem.EvapDesc
    tmpDetail.Precipitation = selItem.Precipitation
    tmpDetail.PrecipitationDesc = selItem.PrecipitationDesc
    tmpDetail.Neutralization = selItem.Neutralization
    tmpDetail.NeutralizationDesc = selItem.NeutralizationDesc
    tmpDetail.CommentWaste = selItem.CommentWaste
    tmpDetail.CommentWasteDesc = selItem.CommentWasteDesc
'            Set tmpCar = New BWGJobDataCarUsed
'    tmpDetail.carNo = "1"
'                Set tmpCar = New BWGJobDataCarUsed
'                tmpCar.ID = GetGUID
'                tmpCar.carNo = tmpDetail.carNo
'                tmpCar.JobID = myJobOpen_.ID
'                tmpCar.isTrailer = tmpDetail.isTrailer
'                myJobOpen_.Cars.Add tmpCar, "'" & tmpCar.ID & "'"
'                grCar.TextMatrix(myJobOpen_.Cars.Count, 0) = tmpCar.ID
'                grCar.TextMatrix(myJobOpen_.Cars.Count, 1) = tmpCar.carNo

    tmpDetail.NonMinRate = selItem.NonMinRate
    tmpDetail.HazMinRate = selItem.HazMinRate
    tmpDetail.NBLMinRate = selItem.NBLMinRate
    tmpDetail.HBLMinRate = selItem.HBLMinRate
    
    isDetailChange_ = True
    myJobOpen_.isChanged = True
    Set tmpDetail = Nothing
    
    If myJobOpen_.JobDetails Is Nothing Then
        Set myJobOpen_.JobDetails = tmpCol
        isDetailChange_ = True
        myJobOpen_.isChanged = True
    End If
    Set tmpCol = Nothing
End Sub

Public Sub RemoveWasteData(selID As String)
Dim tmpCol As Collection
Dim x%
Dim tmpKey As String
    Set tmpCol = myJobOpen_.JobDetails
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
        myJobOpen_.isChanged = True
    End If
End Sub

Public Function getWastDataItems() As Collection
Dim x%
Dim tmpCol As New Collection
Dim tmpDet As BWGJobDataDetail
Dim tmpWaste As BWGWasteDataCR
Dim selCol As Collection, tmpStr As String, tmpArr
    Set selCol = myJobOpen_.JobDetails
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
                tmpWaste.QuotationNo = tmpDet.WasteQuotationNo
                tmpWaste.QuotationID = tmpDet.QuotationID
                tmpWaste.TreatmentFee = tmpDet.QuoTreatmentRate
                tmpWaste.TreatmentUnitName = tmpDet.QuoTreatmentUnitName
                tmpWaste.CarTypeID = tmpDet.QuoTruckTypeID
                tmpWaste.CarTypeName = tmpDet.QuoTruckTypeName
                tmpWaste.StaffCount = tmpDet.QuoLaborCount
                tmpWaste.StaffCharge = tmpDet.QuoLaborCharge
                tmpWaste.isTrailer = tmpDet.isTrailer
                tmpCol.Add tmpWaste, "'" & tmpDet.wastedataID & "_" & tmpDet.QuotationID & "'"
                Set tmpWaste = Nothing
            End If
            Set tmpDet = Nothing
        Next
    End With
    Set getWastDataItems = tmpCol
    Set tmpCol = Nothing
End Function

Public Sub SetDataTruckSubType(selSubType As BWGTruckSubType, isTrailer As Boolean)
'Dim tmpCar As New BWGJobDataCarUsed
Dim x%
'    With grcar
'        For x = 1 To .Rows - 1
'            If Trim(.TextMatrix(x, 0)) = "" Then Exit For
'        Next
'    End With
    If curCar_ Is Nothing Then
        If isExist(myJobOpen_.Cars, "'" & Trim(grCar.TextMatrix(grCar.row, 0)) & "'") Then
            Set curCar_ = myJobOpen_.Cars("'" & Trim(grCar.TextMatrix(grCar.row, 0)) & "'")
        End If
    End If
    With curCar_
        '.ID = GetGUID
        '.JobID = myJobOpen_.ID
        '.RowID = x
        .isTrailer = isTrailer
        .TruckSubTypeID = selSubType.id
        .TruckSubTypeName = selSubType.SubTypeDesc
        .TruckTypeName = selSubType.TypeDesc
        
    End With
    'myJobOpen_.Cars.Add tmpCar, "'" & tmpCar.ID & "'"
    With grCar
        For x = 1 To .Rows - 1
            If Trim(.TextMatrix(x, 0)) = Trim(curCar_.id) Then
                If curCar_.isTrailer Then
                    .TextMatrix(x, 2) = curCar_.TruckTypeName & " , " & curCar_.TruckSubTypeName
                Else
                    .TextMatrix(x, 2) = curCar_.TruckTypeName & ", " & curCar_.TruckSubTypeName
                End If
                Exit For
            End If
        Next
    End With
    
    myJobOpen_.isChanged = True
End Sub
Public Sub SetDataTruckSubTypeRate(selSubType As BWGQuotationTruckRate, carNo As String)
'Dim tmpCar As New BWGJobDataCarUsed
Dim x%
'    With grcar
'        For x = 1 To .Rows - 1
'            If Trim(.TextMatrix(x, 0)) = "" Then Exit For
'        Next
'    End With
    If curCar_ Is Nothing Then
        If isExist(myJobOpen_.Cars, "'" & Trim(grCar.TextMatrix(grCar.row, 0)) & "'") Then
            Set curCar_ = myJobOpen_.Cars("'" & Trim(grCar.TextMatrix(grCar.row, 0)) & "'")
        End If
    End If
    With curCar_
        '.ID = GetGUID
        '.JobID = myJobOpen_.ID
        '.RowID = x
        '.isTrailer = isTrailer
        .TruckSubTypeID = selSubType.id
        .TruckSubTypeName = selSubType.SubTypeDesc
        .TruckTypeName = selSubType.TypeDesc
        .TruckRateID = selSubType.TruckRateID
    End With
    For x = 1 To myJobOpen_.JobDetails.Count
        If myJobOpen_.JobDetails(x).carNo = carNo Then
            myJobOpen_.JobDetails(x).QuoTransportFee = selSubType.Rate
            myJobOpen_.JobDetails(x).QuoTruckTypeID = selSubType.TruckTypeID
        End If
    Next x
    'myJobOpen_.Cars.Add tmpCar, "'" & tmpCar.ID & "'"
    With grCar
        For x = 1 To .Rows - 1
            If Trim(.TextMatrix(x, 0)) = Trim(curCar_.id) Then
                If curCar_.isTrailer Then
                    .TextMatrix(x, 2) = curCar_.TruckTypeName & ", " & curCar_.TruckSubTypeName
                Else
                    .TextMatrix(x, 2) = curCar_.TruckTypeName & ", " & curCar_.TruckSubTypeName
                End If
                Exit For
            End If
        Next
    End With
    
    myJobOpen_.isChanged = True
End Sub

Private Sub PopulateTimeTalbe(selCar As BWGJobDataCarUsed)
Dim x%
    With grTimeTable.getGridObj
'        If selCar.GoSameTime = "Y" Then
'            .Rows = 2
'            .TextMatrix(1, 1) = 1
'        ElseIf selCar.GoSameTime = "N" Then
            .Rows = selCar.TruckCount + 1
            For x = 1 To .Rows - 1
                .TextMatrix(x, 1) = x
            Next
'        End If
    End With
End Sub

Public Sub SetDataTSDF(selTsdf As BWGTSDF)
Dim tmpID As String
Dim tmpDet As BWGJobDataDetail
    With grWasteData.getGridObj
        tmpID = Trim(.TextMatrix(.row, 0))
        If isExist(myJobOpen_.JobDetails, "'" & tmpID & "'") Then
            Set tmpDet = myJobOpen_.JobDetails("'" & tmpID & "'")
            tmpDet.TSDFID = selTsdf.id
            tmpDet.TSDFName = selTsdf.CompanyName
            tmpDet.TSDFRegisNo = selTsdf.TSDFRegisNo
            Set tmpDet = Nothing
            .TextMatrix(.row, 7) = selTsdf.TSDFRegisNo
        End If
    End With
End Sub

Private Function CheckforOrderBreak() As String
    Dim tSql As String, CurrentDate As String
    Dim rs As New ADODB.Recordset
    CheckforOrderBreak = ""
    If myJobOpen_ Is Nothing Then CheckforOrderBreak = "ไม่ได้ระบุลูกค้า"
    tSql = "SELECT CONVERT(varchar(6), GETDATE(), 103) + CONVERT(varchar(4), YEAR(GETDATE()) + 543) AS CurrentDate"
    Set rs = GetRS(tSql, adOpenForwardOnly, adLockOptimistic)
    CurrentDate = Format(rs!CurrentDate, "yyyy/MM/dd")
    
    tSql = "SELECT * FROM tbCustomerOrderBreak" & vbCrLf
    tSql = tSql & "WHERE CustomerID = '" & myJobOpen_.CustomerID & "' AND OrderStatus = 2" & vbCrLf
    tSql = tSql & "AND OrderStatusActive = 1" '"AND OrderDateStart <= '" & CurrentDate & "'"
    Set rs = GetRS(tSql, adOpenForwardOnly, adLockOptimistic)
    If Not rs.EOF Then
        CheckforOrderBreak = "ลูกค้ารายนี้ถูกระงับออเดอร์ ตั้งแต่วันที่ " & Format(rs!OrderDateStart, "dd/MM/yyyy") & " เป็นต้นไป ติดต่อฝ่ายที่รับผิดชอบเพื่อขอข้อมูลเพิ่มเติม"
    End If
    Set rs = Nothing
End Function

Private Sub TxtTankTypeL_Validate(Cancel As Boolean)
Dim tmpDet As BWGJobDataCarUsed
Dim tmpKey As String, BoxCheck As Integer
Dim tmpVal As String, OldValue As String
    tmpKey = Trim(grCar.TextMatrix(1, 0))
    If isExist(myJobOpen_.Cars, "'" & tmpKey & "'") = False Then
        Set tmpDet = New BWGJobDataCarUsed
        tmpDet.JobID = myJobOpen_.id
    Else
        Set tmpDet = myJobOpen_.Cars("'" & tmpKey & "'")
    End If
    tmpDet.RowID = 1
    With tmpDet
        If Trim(TxtTankTypeL.Text) <> "" Then
            .TankType = "L"  'ถังปากกว้าง
            .Tankcount = Trim(TxtTankTypeL.Text)
            .TankPrice = .Tankcount * 150
            .TankPriceV = .Tankcount * 100
        myJobOpen_.isChanged = True
        End If
    End With

End Sub

Private Sub TxtTankTypeS_Validate(Cancel As Boolean)
Dim tmpDet As BWGJobDataCarUsed
Dim tmpKey As String, BoxCheck As Integer
Dim tmpVal As String, OldValue As String
    tmpKey = Trim(grCar.TextMatrix(1, 0))
    If isExist(myJobOpen_.Cars, "'" & tmpKey & "'") = False Then
        Set tmpDet = New BWGJobDataCarUsed
        tmpDet.JobID = myJobOpen_.id
    Else
        Set tmpDet = myJobOpen_.Cars("'" & tmpKey & "'")
    End If
    tmpDet.RowID = 1
    With tmpDet
        If Trim(TxtTankTypeS.Text) <> "" Then
        .TankTypeS = "S"   'ถังปากจู๋
        .TankTypeSCount = Trim(TxtTankTypeS.Text)
        .TankPriceS = .TankTypeSCount * 150
        .TankPriceSV = .TankTypeSCount * 100
        myJobOpen_.isChanged = True
        End If
    End With

End Sub
