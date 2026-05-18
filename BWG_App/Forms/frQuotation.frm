VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmQuotation_Old 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "F111"
   ClientHeight    =   9240
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11415
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9240
   ScaleWidth      =   11415
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtRem 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   7260
      TabIndex        =   24
      Top             =   7980
      Width           =   2085
   End
   Begin VB.CommandButton cmdCloseBill 
      Caption         =   "หยุดใช้งานใบเสนอราคา"
      Height          =   375
      Left            =   9420
      TabIndex        =   34
      Top             =   7950
      Width           =   1965
   End
   Begin VB.TextBox txtPosition 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   4320
      TabIndex        =   23
      Top             =   7980
      Width           =   2085
   End
   Begin VB.ComboBox cboApprove 
      Height          =   315
      Left            =   990
      TabIndex        =   22
      Top             =   7980
      Width           =   2505
   End
   Begin VB.CommandButton Command1 
      Caption         =   "เลือกหมายเหตุ"
      Height          =   795
      Left            =   3960
      TabIndex        =   33
      Top             =   8400
      Width           =   1305
   End
   Begin VB.ComboBox cboCompany 
      Height          =   315
      Left            =   7830
      Style           =   2  'Dropdown List
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   300
      Width           =   3525
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H8000000E&
      Height          =   2400
      Left            =   30
      TabIndex        =   66
      Top             =   5550
      Width           =   11385
      Begin WasteManagment.ctlGrid grRemark 
         Height          =   2235
         Left            =   30
         TabIndex        =   21
         Top             =   120
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   3942
      End
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
      Height          =   795
      Left            =   1050
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   8400
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
      TabIndex        =   58
      Top             =   7500
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
      TabIndex        =   57
      Top             =   7500
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
      TabIndex        =   56
      Top             =   7500
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
      TabIndex        =   55
      Top             =   7500
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000E&
      Caption         =   "รายละเอียดใบเสนอราคา"
      Height          =   2295
      Left            =   30
      TabIndex        =   39
      Top             =   690
      Width           =   11385
      Begin VB.TextBox txtHBLRate 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   9750
         TabIndex        =   19
         Top             =   1830
         Width           =   645
      End
      Begin VB.TextBox txtNBLRate 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   7140
         TabIndex        =   18
         Top             =   1830
         Width           =   645
      End
      Begin VB.TextBox txtHazRate 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   4470
         TabIndex        =   17
         Top             =   1830
         Width           =   645
      End
      Begin VB.TextBox txtNonRate 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1740
         TabIndex        =   16
         Top             =   1830
         Width           =   645
      End
      Begin VB.ComboBox cboIsTrailer 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frQuotation.frx":0000
         Left            =   3360
         List            =   "frQuotation.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   1440
         Width           =   1215
      End
      Begin VB.ComboBox cboTransUnit 
         Height          =   315
         ItemData        =   "frQuotation.frx":0020
         Left            =   6600
         List            =   "frQuotation.frx":0022
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   1440
         Width           =   1245
      End
      Begin VB.TextBox txtTransRate 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   5520
         TabIndex        =   13
         Top             =   1440
         Width           =   1035
      End
      Begin VB.TextBox txtCarType 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1740
         TabIndex        =   11
         Top             =   1440
         Width           =   1605
      End
      Begin VB.TextBox txtExtraCharge 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   9750
         TabIndex        =   15
         Top             =   1440
         Width           =   645
      End
      Begin VB.CheckBox chkIncVAT 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "ราคารวม VAT แล้ว"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   9330
         TabIndex        =   6
         Top             =   690
         Width           =   1665
      End
      Begin VB.TextBox txtMinQty 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   9750
         TabIndex        =   10
         Top             =   1050
         Width           =   645
      End
      Begin VB.TextBox txtLaborCount 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   5520
         TabIndex        =   8
         Top             =   1050
         Width           =   435
      End
      Begin VB.CheckBox chkIncTrans 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "ราคารวมค่าขนส่งแล้ว"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   6750
         TabIndex        =   9
         Top             =   1110
         Width           =   1815
      End
      Begin WasteManagment.ctlDate dtIssuedDate 
         Height          =   315
         Left            =   4320
         TabIndex        =   2
         Top             =   270
         Width           =   1785
         _ExtentX        =   3149
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
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1050
         Width           =   2595
      End
      Begin VB.ComboBox cboCustStaff 
         Height          =   315
         Left            =   6870
         TabIndex        =   5
         Top             =   660
         Width           =   2355
      End
      Begin VB.ComboBox cboDiscountType 
         Height          =   315
         Left            =   10560
         Style           =   2  'Dropdown List
         TabIndex        =   36
         Top             =   2190
         Visible         =   0   'False
         Width           =   870
      End
      Begin VB.TextBox txtDiscount 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   9810
         TabIndex        =   35
         Top             =   2190
         Visible         =   0   'False
         Width           =   705
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
         Top             =   660
         Width           =   2595
      End
      Begin VB.TextBox txtCustomer 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   6870
         TabIndex        =   3
         Top             =   270
         Width           =   4125
      End
      Begin VB.TextBox txtQuoNo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C00000&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1740
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Text            =   "== AUTO =="
         Top             =   270
         Width           =   1755
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ราคาขั้นต่ำ HBL"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   24
         Left            =   8580
         TabIndex        =   80
         Top             =   1890
         Width           =   1200
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "บาท/ตัน"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   23
         Left            =   10470
         TabIndex        =   79
         Top             =   1890
         Width           =   585
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ราคาขั้นต่ำ NBL"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   22
         Left            =   5970
         TabIndex        =   78
         Top             =   1890
         Width           =   1200
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "บาท/ตัน"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   21
         Left            =   7860
         TabIndex        =   77
         Top             =   1890
         Width           =   585
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ราคาขั้นต่ำ Haz"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   20
         Left            =   3300
         TabIndex        =   76
         Top             =   1890
         Width           =   1200
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "บาท/ตัน"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   19
         Left            =   5190
         TabIndex        =   75
         Top             =   1890
         Width           =   585
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ราคาขั้นต่ำ Non"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   18
         Left            =   570
         TabIndex        =   74
         Top             =   1890
         Width           =   1200
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "บาท/ตัน"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   17
         Left            =   2460
         TabIndex        =   73
         Top             =   1890
         Width           =   585
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ค่าขนส่ง"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   14
         Left            =   4890
         TabIndex        =   70
         Top             =   1500
         Width           =   780
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ประเภทรถ"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   13
         Left            =   930
         TabIndex        =   69
         Top             =   1500
         Width           =   750
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ค่าแรงคนงานเพิ่ม"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   10
         Left            =   8400
         TabIndex        =   68
         Top             =   1500
         Width           =   1260
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "บาท/คน"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   3
         Left            =   10470
         TabIndex        =   67
         Top             =   1530
         Width           =   585
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ตัน/คัน"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   9
         Left            =   10470
         TabIndex        =   65
         Top             =   1140
         Width           =   570
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ปริมาณขนขั้นต่ำ"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   7
         Left            =   8580
         TabIndex        =   64
         Top             =   1110
         Width           =   1200
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "คน/คัน"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   6
         Left            =   6030
         TabIndex        =   63
         Top             =   1110
         Width           =   570
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "รวมคนงาน"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   5
         Left            =   4680
         TabIndex        =   62
         Top             =   1110
         Width           =   780
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ผู้อนุมัติใบเสนอราคา"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   61
         Top             =   1110
         Width           =   1440
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ส่วนลด"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   12
         Left            =   9270
         TabIndex        =   46
         Top             =   2250
         Visible         =   0   'False
         Width           =   555
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ผู้รับใบเสนอราคา"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   11
         Left            =   5580
         TabIndex        =   45
         Top             =   720
         Width           =   1215
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ผู้ออกใบเสนอราคา"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   360
         TabIndex        =   43
         Top             =   720
         Width           =   1305
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ลูกค้า"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   8
         Left            =   6420
         TabIndex        =   42
         Top             =   330
         Width           =   375
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "เลขที่ใบเสนอราคา"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   390
         TabIndex        =   41
         Top             =   330
         Width           =   1275
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "วันที่ออก"
         Height          =   195
         Index           =   8
         Left            =   3645
         TabIndex        =   40
         Top             =   330
         Width           =   600
      End
   End
   Begin VB.CommandButton cmdAddItem 
      Caption         =   "เลือกกาก"
      Height          =   795
      Left            =   2610
      TabIndex        =   32
      Top             =   8400
      Width           =   1305
   End
   Begin VB.CommandButton cmdPrintPO 
      Caption         =   "พิมพ์"
      Height          =   795
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   8400
      Width           =   945
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ยกเลิก"
      Height          =   795
      Left            =   9450
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   8400
      Width           =   945
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "ปิด"
      Height          =   795
      Left            =   10440
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   8400
      Width           =   945
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "สร้างใหม่"
      Height          =   795
      Left            =   8460
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   8400
      Width           =   945
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "เปิดข้อมูล"
      Height          =   795
      Left            =   7470
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   8400
      Width           =   945
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "เก็บข้อมูล"
      Height          =   795
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   8400
      Width           =   945
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H8000000E&
      Caption         =   "รายการการให้บริการการขนกากของเสีย"
      Height          =   2520
      Left            =   30
      TabIndex        =   37
      Top             =   3030
      Width           =   11385
      Begin WasteManagment.ctlGrid grItem 
         Height          =   2205
         Left            =   30
         TabIndex        =   20
         Top             =   240
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   3889
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "หมายเหตุ"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   25
      Left            =   6510
      TabIndex        =   81
      Top             =   8040
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
      TabIndex        =   72
      Top             =   8040
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
      TabIndex        =   71
      Top             =   8040
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
      TabIndex        =   60
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
      TabIndex        =   59
      Top             =   7560
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
      TabIndex        =   54
      Top             =   7560
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "บาท"
      Height          =   195
      Index           =   5
      Left            =   11430
      TabIndex        =   53
      Top             =   6300
      Width           =   1395
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VAT"
      Height          =   195
      Index           =   6
      Left            =   5730
      TabIndex        =   52
      Top             =   7560
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
      TabIndex        =   51
      Top             =   7560
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
      TabIndex        =   50
      Top             =   7560
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
      TabIndex        =   49
      Top             =   7560
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
      TabIndex        =   48
      Top             =   7560
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
      TabIndex        =   47
      Top             =   7560
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
      Left            =   7830
      TabIndex        =   44
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
      TabIndex        =   38
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
      X2              =   13105
      Y1              =   8340
      Y2              =   8340
   End
   Begin VB.Image Image2 
      Height          =   645
      Left            =   0
      Picture         =   "frQuotation.frx":0024
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12615
   End
End
Attribute VB_Name = "frmQuotation_Old"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim myQuotation_ As BWGQuotationHeader
Dim curCustomer_ As BWGCustomer
Dim isDetailChange_ As Boolean
Dim curCarType_ As BWGTruckType
Dim curRole_ As AWSRoleType
Dim colPendingJob_ As Collection

Private Sub cboApprove_Click()
    If Trim(cboApprove.Text) = SystemConfig.SignName1 Then
        txtPosition.Text = SystemConfig.SignPosition1
    ElseIf Trim(cboApprove.Text) = SystemConfig.SignName2 Then
        txtPosition.Text = SystemConfig.SignPosition2
    ElseIf Trim(cboApprove.Text) = SystemConfig.SignName3 Then
        txtPosition.Text = SystemConfig.SignPosition3
    End If
End Sub

Private Sub cboApprove_Validate(Cancel As Boolean)
    myQuotation_.SignName = cboApprove.Text
End Sub

Private Sub cboCompany_Validate(Cancel As Boolean)
    If cboCompany.ListIndex >= 0 Then
        myQuotation_.companyID = CurrentUser.WorkCompanies(cboCompany.ListIndex + 1).ID
    Else
        myQuotation_.companyID = ""
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

Private Sub cboTransUnit_Validate(Cancel As Boolean)
    If cboTransUnit.ListIndex > -1 Then
        myQuotation_.TransferUnitID = cboTransUnit.Text
        myQuotation_.TransferUnitName = cboTransUnit.Text
    Else
        myQuotation_.TransferUnitID = ""
        myQuotation_.TransferUnitName = ""
    End If
End Sub

Private Sub chkIncTrans_Click()
    If chkIncTrans.Value = 1 Then
        txtTransRate.Enabled = False
        cboTransUnit.Enabled = False
    Else
        txtTransRate.Enabled = True
        cboTransUnit.Enabled = True
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
    tmpStr = CurrentUser.VerifyAccessRole(Me.Name, ApproveRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    
    If Trim(txtCustomer.Text) = "" Then
        MsgBox "กรุณาใส่ชื่อลูกค้า", vbExclamation
        Exit Sub
    End If
    If myQuotation_.QuotationDetails Is Nothing Then
        MsgBox "กรุณาระบุรายการสินค้าที่ต้องการสั่งซื้อ", vbExclamation
        Exit Sub
    End If
    If myQuotation_.QuotationDetails.Count = 0 Then
        MsgBox "กรุณาระบุรายการสินค้าที่ต้องการสั่งซื้อ", vbExclamation
        Exit Sub
    End If

    If myQuotation_.isChanged And cmdSave.Enabled Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนสร้างใหม่หรือไม่?", vbQuestion + vbYesNo)
        If selResult = vbYes Then
            cmdSave_Click
            If myQuotation_.isChanged Then Exit Sub
        Else 'If selResult = vbCancel Then
            Exit Sub
        End If
    End If

    result = MsgBox("ต้องการอนุมัติใบเสนอราคานี้ใช่หรือไม่", vbQuestion + vbYesNo)
    If result = vbYes Then
        myQuotation_.IsApproved = True
        myQuotation_.ApprovedDate = TodayDate
        myQuotation_.ApprovedStaffID = CurrentUser.ID
        myQuotation_.ApprovedStaffName = CurrentUser.StaffFName & " " & CurrentUser.StaffLName
        myQuotation_.Save
        Label2(4).Caption = "อนุมัติแล้วเมื่อวันที่ " & FormatYMD_to_DMY(myQuotation_.ApprovedDate, "/", "/") & " โดย " & myQuotation_.ApprovedStaffName
        txtApprovedStaff.Text = myQuotation_.ApprovedStaffName
        SetEnabledScreen
    
        If myQuotation_.IsApproved Then
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
            cmdSave_Click
            If myQuotation_.isChanged Then Exit Sub
        ElseIf selResult = vbCancel Then
            Exit Sub
        End If
    End If
    ClearScreen
    CreateNewQuotation
    txtCustomer.SetFocus
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
        If CheckPendingJob Then
            If MsgBox("มี Job ที่ยังไม่ได้ออก Manifest ใช้งานใบเสนอราคานี้อยู่ ท่านต้องการย้าย Job เหล่านี้ไปใช้งานใบเสนอราคาอื่นหรือไม่", vbQuestion + vbYesNo) = vbYes Then
                Load frmQuotationTransfer2
                With frmQuotationTransfer2
                    .SetDataQuotation myQuotation_
                    .SetPendingJob colPendingJob_, txtQuoNo.Text
                    .Show 1
                End With
                If CheckPendingJob Then
                    MsgBox "ยังมี Job ที่ยังไม่ได้ออก Manifest ใช้งานใบเสนอราคานี้อยู่ ไม่สามารถหยุดใช้งานใบเสนอราคานี้ได้", vbExclamation
                    Exit Sub
                End If
            Else
                MsgBox "ไม่สามารถทำการหยุดใช้งานใบเสนอราคาได้", vbExclamation
                Exit Sub
            End If
        End If
        
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

Private Sub cmdPrintPO_Click()
Dim tmpSql As String
Dim selResult
Dim tmpStr As String
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
            cmdSave_Click
            If myQuotation_.isChanged Then Exit Sub
        Else 'If selResult = vbCancel Then
            Exit Sub
        End If
    End If

    With rpt1
        .LogonInfo(0) = "DSN = " & DSN_NAME & ";UID = " & strUid & ";PWD = " & strPwd
        .ReportFileName = App.Path & "\reports\Quotation.rpt"
        .Formulas(0) = "ShowPODate=" & Chr(34) & dtIssuedDate.ValueDMY & Chr(34)
        .Formulas(1) = "CustName=" & Chr(34) & txtCustomer.Text & Chr(34)
        .Formulas(2) = "Addr1=" & Chr(34) & curCustomer_.Address & " " & curCustomer_.TumbolName & " " & curCustomer_.AumphurName & " " & curCustomer_.ProvinceName & Chr(34)
        .Formulas(3) = "TelNo=" & Chr(34) & curCustomer_.TelNo & Chr(34)
        .Formulas(4) = "FaxNo=" & Chr(34) & curCustomer_.FaxNo & Chr(34)
        If chkIncTrans.Value = 0 Then
            .Formulas(5) = "IsIncTransFee=" & Chr(34) & "ราคายังไม่รวมค่าขนส่ง" & Chr(34)
        Else
            .Formulas(5) = "IsIncTransFee=" & Chr(34) & "ราคารวมค่าขนส่งแล้ว" & Chr(34)
        End If
        .Formulas(6) = "AttentionTo=" & Chr(34) & cboCustStaff.Text & Chr(34)
        .Formulas(7) = "QuotationNo=" & Chr(34) & txtQuoNo.Text & Chr(34)
        .Formulas(8) = "ApprovalName=" & Chr(34) & "(  " & cboApprove.Text & "  )" & Chr(34)
        .Formulas(9) = "ApprovePosition=" & Chr(34) & txtPosition.Text & Chr(34)
        .Formulas(10) = "SaleName=" & Chr(34) & curCustomer_.UnderSaleName & " โทร. " & curCustomer_.UnderSaleTelNo & Chr(34)
        tmpSql = ""
        tmpSql = "Select * from Com_vwQuotationDetailSearch " & vbCrLf
        tmpSql = tmpSql & "Where QuotationID='" & myQuotation_.ID & "' " & vbCrLf
        tmpSql = tmpSql & "Order By RowPos"
        .SQLQuery = tmpSql
        .Destination = crptToWindow
        .PrintReport
        If Trim(.LastErrorString) <> "" Then MsgBox .LastErrorString
    End With
End Sub

Private Sub cmdSave_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole(Me.Name, curRole_)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
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
        End If
        If Trim(cboTransUnit.Text) = "" Then
            MsgBox "กรุณาระบุหน่วยในการคิดค่าขนส่ง", vbExclamation
            cboTransUnit.SetFocus
            Exit Sub
        End If
    End If
    If myQuotation_.isChanged And cmdSave.Enabled Then
        myQuotation_.Save
        txtQuoNo.Text = myQuotation_.QuotationNo
        If myQuotation_.IsApproved Then
            Label2(4).Caption = "อนุมัติแล้วเมื่อวันที่ " & FormatYMD_to_DMY(myQuotation_.ApprovedDate, "/", "/") & " โดย " & myQuotation_.ApprovedStaffName
        Else
            Label2(4).Caption = "รายการรออนุมัติ"
        End If
    End If
    If myQuotation_.isChanged = False Then
        MsgBox "จัดเก็บข้อมูลเรียบร้อยแล้ว", vbExclamation
    End If
End Sub

Private Sub cmdSearch_Click()
Dim selResult As Integer
Dim x%
Dim tmpKey As String
Dim tmpVal As String
Dim tmpDet As BWGQuotationDetail
Dim tmpCol As Collection
    If myQuotation_.isChanged And cmdSave.Enabled Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนทำงานต่อหรือไม่?", vbQuestion + vbYesNoCancel)
        If selResult = vbYes Then
            cmdSave_Click
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

Private Sub dtIssuedDate_Validate(Cancel As Boolean)
    myQuotation_.IssuedDate = dtIssuedDate.ValueYMD
End Sub

Private Sub Form_Load()
    CanAccess = False: CanUpdate = False
    SetPermission Me, CanAccess, CanUpdate
    If CanAccess = False Then
        MsgBox "ท่านไม่มีสิทธิใช้ระบบนี้", vbExclamation, "Warning"
        Unload Me
    End If
    
    PopulateCompany
    With cboTransUnit
        .AddItem "บาท/เที่ยว"
        .AddItem "บาท/ตัน"
    End With
    cboIsTrailer.ListIndex = 0
    With frmPicture
        cmdSave.Picture = .picSave.Picture
        cmdSearch.Picture = .picSearch.Picture
        cmdClear.Picture = .picNew.Picture
        cmdCancel.Picture = .picDelete.Picture
        cmdClose.Picture = .picExit.Picture
        cmdPrintPO.Picture = .picPrint
    End With
    With cboApprove
        .AddItem SystemConfig.SignName1
        If Trim(SystemConfig.SignName2) <> "" Then .AddItem SystemConfig.SignName2
        If Trim(SystemConfig.SignName3) <> "" Then .AddItem SystemConfig.SignName3
    End With
    PopulateUnit
    txtIssuedStaff.Text = CurrentUser.StaffFName & " " & CurrentUser.StaffLName
    Label2(4).Caption = "==สถานะการอนุมัติ=="
    cboDiscountType.AddItem "%"
    cboDiscountType.AddItem "บาท"
    cboDiscountType.ListIndex = 0
    grItem.ColInputType(3) = ComboBoxEnum
    grItem.ColInputType(5) = ComboBoxEnum
    grItem.ColInputType(7) = ComboBoxEnum
    grItem.GridMoveType 1
    With grItem.getGridObj
        .Cols = 7
        .ColWidth(0) = 0
        .ColWidth(1) = 4425
        .ColWidth(2) = 1215
        .ColWidth(3) = 1065
        .ColWidth(4) = 1200
        .ColWidth(5) = 1065
        .ColWidth(6) = 2070
        .Rows = 20
        .FixedAlignment(1) = 4
        .FixedAlignment(2) = 4
        .FixedAlignment(3) = 4
        .FixedAlignment(4) = 4
        .FixedAlignment(5) = 4
        .FixedAlignment(6) = 4

        .ColAlignment(1) = 1
        .ColAlignment(2) = 7
        .ColAlignment(3) = 1
        .ColAlignment(4) = 7
        .ColAlignment(5) = 1
        .ColAlignment(6) = 1

        .TextMatrix(0, 1) = "ชื่อกากของเสีย"
        .TextMatrix(0, 2) = "ปริมาณ"
        .TextMatrix(0, 3) = "หน่วย"
        .TextMatrix(0, 4) = "อัตราค่าบำบัด"
        .TextMatrix(0, 5) = "หน่วย"
        .TextMatrix(0, 6) = "หมายเหตุ"
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
    CreateNewQuotation
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
            cmdSave_Click
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

Private Sub grItem_BeforeUpdateValue(NewValue As String, Row As Integer, Col As Integer, Cancel As Boolean)
    If Col = 2 Or Col = 4 Then
        If IsNumeric(NewValue) = False Then
            MsgBox "ข้อมูลไม่ถูกต้อง ต้องเป็นตัวเลขเท่านั้น", vbExclamation
            Cancel = True
        End If
    End If
End Sub

Private Sub grItem_Click(Row As Integer, Col As Integer)
'    MsgBox grItem.getGridObj.ColWidth(col)
End Sub

Private Sub grItem_KeyDown(KeyCode As Integer, Shift As Integer)
Dim selResult As Integer
Dim x%, tmpKey$, itemCount%
Dim selCol As Collection
    With grItem.getGridObj
        If KeyCode = 46 And Trim(.TextMatrix(.Row, 0)) <> "" Then
            selResult = MsgBox("ต้องการลบรายการกากของเสียนี้ออกจากใบเสนอราคาหรือไม่?", vbQuestion + vbYesNo)
            If selResult = vbYes Then
                tmpKey = "'" & .TextMatrix(.Row, 0) & "'"
                .RemoveItem .Row
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

Private Sub grItem_RolColChanged(Row As Integer, Col As Integer)
Dim x%
Dim tmpCbo As ComboBox
Dim tmpKey As String
    tmpKey = "'" & grItem.getGridObj.TextMatrix(Row, 0) & "'"
    If Trim(tmpKey) = "''" Then
        grItem.ColEnabled(Col) = False
    End If
    If isExist(myQuotation_.QuotationDetails, tmpKey) Then
        If Col = 1 Or Col = 2 Or Col = 3 Or Col = 4 Or Col = 5 Or Col = 6 Then
            grItem.ColEnabled(Col) = True
        Else
            grItem.ColEnabled(Col) = False
        End If
    End If
    If (Col = 3) And Not colUnit_ Is Nothing Then
        Set tmpCbo = grItem.getColObject(Col)
        tmpCbo.Clear
        For x = 1 To colUnit_.Count
            tmpCbo.AddItem colUnit_(x).UnitName
        Next
        Set tmpCbo = Nothing
    ElseIf Col = 5 Then
        Set tmpCbo = grItem.getColObject(Col)
        tmpCbo.Clear
        With tmpCbo
            .AddItem "บาท/เที่ยว"
            .AddItem "บาท/ตัน"
            .AddItem "บาท/ถัง"
            .AddItem "บาท/ลบ.ม."
        End With
    End If
End Sub

Private Sub grItem_UpdateValue(Row As Integer, Col As Integer)
Dim tmpDet As BWGQuotationDetail
Dim tmpKey As String
Dim tmpVal As String
    tmpKey = Trim(grItem.getGridObj.TextMatrix(Row, 0))
    Set tmpDet = myQuotation_.QuotationDetails("'" & tmpKey & "'")
    tmpVal = Trim(grItem.getGridObj.TextMatrix(Row, Col))
    With tmpDet
        Select Case Col
            Case 1
                .BillDesc = Trim(tmpVal)
            Case 2
                If IsNumeric(Trim(tmpVal)) Then
                    .YearQty = CDbl(Trim(tmpVal))
                Else
                    .YearQty = 0
                End If
            Case 3
                If Trim(tmpVal) <> "" Then
                    .WasteQtyUnitID = colUnit_(grItem.getColObject(3).ListIndex + 1).ID
                    .WasteQtyUnitName = Trim(tmpVal)
                Else
                    .WasteQtyUnitID = ""
                    .WasteQtyUnitName = ""
                End If
            Case 4
                If IsNumeric(Trim(tmpVal)) Then
                    .TreatmentRate = CDbl(Trim(tmpVal))
                Else
                    .TreatmentRate = 0
                End If
            Case 5
                If Trim(tmpVal) <> "" Then
                    .TreatmentUnitID = Trim(tmpVal)     'colUnit_(grItem.getColObject(7).ListIndex + 1).ID
                    .TreatmentUnitName = Trim(tmpVal)
                Else
                    .TreatmentUnitID = ""
                    .TreatmentUnitName = ""
                End If
            Case 6
                .RemDesc = Trim(tmpVal)
        End Select
        myQuotation_.isChanged = True
    End With
    Set tmpDet = Nothing
End Sub

Private Sub grRemark_KeyDown(KeyCode As Integer, Shift As Integer)
Dim selResult As Integer
Dim x%, tmpKey$, itemCount%
Dim selCol As Collection
    With grRemark.getGridObj
        If KeyCode = 46 And Trim(.TextMatrix(.Row, 0)) <> "" Then
            selResult = MsgBox("ต้องการลบรายการหมายเหตุนี้ออกจากใบเสนอราคาหรือไม่?", vbQuestion + vbYesNo)
            If selResult = vbYes Then
                tmpKey = "'" & .TextMatrix(.Row, 0) & "'"
                .RemoveItem .Row
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

Private Sub grRemark_UpdateValue(Row As Integer, Col As Integer)
Dim tmpDet As BWGQuotationRem
Dim tmpKey As String
Dim tmpVal As String
    tmpKey = Trim(grRemark.getGridObj.TextMatrix(Row, 0))
    If isExist(myQuotation_.RemarkDetails, "'" & tmpKey & "'") = False Then
        Set tmpDet = New BWGQuotationRem
        tmpDet.ID = GetGUID
        tmpDet.QuotationID = myQuotation_.ID
        tmpDet.ItemNo = Row
        grRemark.getGridObj.TextMatrix(Row, 0) = tmpDet.ID
        myQuotation_.RemarkDetails.Add tmpDet, "'" & tmpDet.ID & "'"
    Else
        Set tmpDet = myQuotation_.RemarkDetails("'" & tmpKey & "'")
    End If
    tmpVal = Trim(grRemark.getGridObj.TextMatrix(Row, Col))
    With tmpDet
        Select Case Col
            Case 1
                .RemDesc = Trim(tmpVal)
                .ItemNo = Row
        End Select
        myQuotation_.isChanged = True
    End With
    Set tmpDet = Nothing
End Sub

Private Sub txtCarType_Validate(Cancel As Boolean)
    If curCarType_ Is Nothing Then Set curCarType_ = New BWGTruckType
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
                .getCriteria (txtCustomer.Text)
                .SetCallerForm Me
                .Show 1
            End With
            If myQuotation_.CustomerID = "" Then txtCustomer.Text = ""
        End If
    End If
End Sub

Private Sub CreateNewQuotation()
Set myQuotation_ = Nothing
Set myQuotation_ = New BWGQuotationHeader
    curRole_ = AddNewRole
    With myQuotation_
        .ID = GetGUID
        If cboCompany.ListIndex >= 0 Then
            .companyID = CurrentUser.WorkCompanies(cboCompany.ListIndex + 1).ID
        End If
        If Not curCustomer_ Is Nothing Then
            .CustomerID = curCustomer_.ID
            .CustomerName = curCustomer_.CustomerName
        End If
        .IssuedDate = dtIssuedDate.ValueYMD
        .IsApproved = False
        .isCancel = False
        .IssuedStaffID = CurrentUser.ID
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
    txtIssuedStaff.Text = CurrentUser.StaffFName & " " & CurrentUser.StaffLName
    cboCustStaff.Text = ""
    chkIncVAT.Value = 0
    txtExtraCharge.Text = ""
    txtCarType.Text = ""
    Set curCarType_ = Nothing
    Set curCarType_ = New BWGTruckType
    txtTransRate.Text = ""
    cboTransUnit.ListIndex = -1
    cboApprove.Text = ""
    txtRem.Text = ""
    txtPosition.Text = ""
    cboIsTrailer.ListIndex = 0
    cboIsTrailer.Enabled = False
    If cboCompany.ListCount > 0 Then
        cboCompany.ListIndex = 0
    End If
    txtApprovedStaff.Text = ""
    txtLaborCount.Text = ""
    chkIncTrans.Value = 0
    txtMinQty.Text = ""
    txtNonRate.Text = ""
    txtHazRate.Text = ""
    txtNBLRate.Text = ""
    txtHBLRate.Text = ""
    lbCancel.Visible = False
    Label2(4).Caption = "==สถานะการอนุมัติ=="
    grItem.ClearAllData
    grItem.getGridObj.Rows = 20
    grRemark.ClearAllData
    grRemark.getGridObj.Rows = 20
    Set curCustomer_ = Nothing
    Set curCustomer_ = New BWGCustomer
End Sub

Public Sub SetDataQuotation(selItem As BWGQuotationHeader)
Dim tmpS As New BWGSearchManager
Dim tmpCol As Collection
    curRole_ = UpdateRole
    Set myQuotation_ = Nothing
    Set myQuotation_ = selItem
    If myQuotation_ Is Nothing Then Exit Sub
    With myQuotation_
        txtQuoNo.Text = .QuotationNo
        dtIssuedDate.ValueYMD = .IssuedDate
        Set tmpCol = tmpS.CustomerSearch("CustomerID='" & .CustomerID & "'")
        If tmpCol.Count > 0 Then SetDataCustomer tmpCol(1), False
        txtIssuedStaff.Text = .IssuedStaffName
        cboCustStaff.Text = .AttentionTo
        If .IsPriceIncVAT Then
            chkIncVAT.Value = 1
        Else
            chkIncVAT.Value = 0
        End If
        If isExist(CurrentUser.WorkCompanies, "'" & .companyID & "'") Then
            SetTextToCombo CurrentUser.WorkCompanies("'" & .companyID & "'").CompanyNameTH, cboCompany
        End If
        txtRem.Text = .RemDesc
        txtExtraCharge.Text = .ExtraLaborFee
        txtApprovedStaff.Text = .ApprovedStaffName
        If .StaffCountPerCar > 0 Then txtLaborCount.Text = .StaffCountPerCar
        If .IsPriceIncTransport Then
            chkIncTrans.Value = 1
        Else
            chkIncTrans.Value = 0
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
        If .isTrailer Then
            cboIsTrailer.ListIndex = 1
        Else
            cboIsTrailer.ListIndex = 0
        End If
        If .IsApproved = True Then
            Label2(4).Caption = "อนุมัติแล้วเมื่อวันที่ " & FormatYMD_to_DMY(myQuotation_.ApprovedDate, "/", "/") & " โดย " & myQuotation_.ApprovedStaffName
        Else
            Label2(4).Caption = "รายการรออนุมัติ"
        End If
        If .isClosed Then
            lbCancel.Caption = "Closed"
            lbCancel.Visible = True
        ElseIf .isCancel Then
            lbCancel.Caption = "ยกเลิก"
            lbCancel.Visible = True
        ElseIf .IsApproved Then
            lbCancel.Caption = "Approved"
            lbCancel.Visible = True
        Else
            lbCancel.Visible = False
        End If
        cboApprove.Text = .SignName
        txtPosition.Text = .SignPosition
        SetEnabledScreen
        
    End With
    ShowQuoDetails
    ShowQuoRemarks
    myQuotation_.isChanged = False
    Set tmpCol = Nothing
    Set tmpS = Nothing
End Sub

Public Sub AddWasteDataItem(selItem As BWGWasteDataCR)
Dim tmpCol As Collection
Dim x%, isNew As Boolean
Dim tmpID As String
Dim tmpDetail As New BWGQuotationDetail
    Set tmpCol = myQuotation_.QuotationDetails

    If tmpCol Is Nothing Then Set tmpCol = New Collection
    isNew = True
    tmpID = Trim(selItem.ID)
    For x = 1 To tmpCol.Count
        'If tmpCol(x).selitem.ID = tmpID Then
        If tmpCol(x).wastedataID = tmpID Then
            isNew = False
            Exit For
        End If
    Next

    If isNew Then
        tmpDetail.ID = GetGUID
        tmpDetail.QuotationID = myQuotation_.ID
        tmpDetail.wastedataID = selItem.ID
        tmpDetail.WasteEUCode = selItem.EUCode1 & "-" & selItem.EUCode2 & "-" & selItem.EUCode3
        tmpDetail.WasteName = selItem.WasteName
        tmpDetail.WasteNo = selItem.WasteNo
        tmpDetail.BillDesc = selItem.WasteName
        tmpDetail.WasteType = selItem.WasteType
        tmpCol.Add tmpDetail, "'" & tmpDetail.ID & "'"
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
Dim sqlStr As String
    Set curCustomer_ = Nothing
    Set curCustomer_ = selCustomer
    If curCustomer_ Is Nothing Then Exit Sub
    cboCustStaff.Clear
    With curCustomer_
        sqlStr = "Select * from tbCustomerStaff Where CustomerID='" & .ID & "'"
        Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
        If tmpRec Is Nothing Then Exit Sub
        If Trim(curCustomer_.ContactPerson) <> "" Then cboCustStaff.AddItem curCustomer_.ContactPerson
        Do Until tmpRec.EOF
            cboCustStaff.AddItem Trim("" & tmpRec("FName") & " " & tmpRec("LName"))
            tmpRec.MoveNext
        Loop
        If cboCustStaff.ListCount > 0 Then
            cboCustStaff.ListIndex = 0
            myQuotation_.AttentionTo = Trim(cboCustStaff.Text)
        End If
        tmpRec.Close
        Set tmpRec = Nothing
        txtCustomer.Text = .CustomerName
        If isAssigned Then
            myQuotation_.CustomerID = .ID
            myQuotation_.CustomerName = .CustomerName
        End If
    End With
    Exit Sub
ErrD:
    GetMsg "", vbCritical
    Exit Sub
End Sub

Private Sub ShowQuoDetails()
Dim x%, Y%
Dim tmpCol As Collection
Dim tmpDet As BWGQuotationDetail
Dim tmpKey As String, rowCount As Long
Dim forColor As String
    With grItem.getGridObj
        grItem.ClearAllData
                
        Set tmpCol = myQuotation_.QuotationDetails
        rowCount = 0
        forColor = vbBlack
        If .Rows - 1 <= tmpCol.Count Then .Rows = tmpCol.Count + 20
        For x = 1 To tmpCol.Count
            Set tmpDet = tmpCol(x)
            .TextMatrix(x + rowCount, 0) = tmpDet.ID
            .Row = x + rowCount
            .Col = 1
            .CellForeColor = forColor
            'If Trim(tmpDet.WasteType) <> "" Then
            '    .TextMatrix(x + rowCount, 1) = tmpDet.BillDesc & " [" & tmpDet.WasteType & "]"
            'Else
                .TextMatrix(x + rowCount, 1) = tmpDet.BillDesc
            'End If
            .TextMatrix(x + rowCount, 2) = tmpDet.YearQty
            .TextMatrix(x + rowCount, 3) = tmpDet.WasteQtyUnitName
            .TextMatrix(x + rowCount, 4) = tmpDet.TreatmentRate
            .TextMatrix(x + rowCount, 5) = tmpDet.TreatmentUnitName
            .TextMatrix(x + rowCount, 6) = tmpDet.RemDesc
            tmpDet.RowPos = x
        Next
        .Row = 1
        .Col = 1
    End With
    Set tmpCol = Nothing
End Sub

Public Function getWastDataItems() As Collection
Dim x%
Dim tmpCol As New Collection
Dim tmpDet As BWGQuotationDetail
Dim tmpWaste As BWGWasteDataCR
Dim selCol As Collection, tmpStr As String, tmpArr
    Set selCol = myQuotation_.QuotationDetails
    With selCol
        For x = 1 To .Count
            Set tmpDet = .Item(x)
            If Trim(tmpDet.wastedataID) <> "" Then
                Set tmpWaste = New BWGWasteDataCR
                tmpWaste.ID = tmpDet.wastedataID
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

Private Sub SetEnabledScreen()
Dim eFlag As Boolean

Dim UserGrp As Collection
Dim Group As New AWSUserGroup
Dim StrGrp As String
Dim i As Integer
    eFlag = Not (myQuotation_.isCancel Or myQuotation_.isClosed Or myQuotation_.IsApproved)
    
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
            tmpKey = "'" & tmpCol(x).ID & "'"
            Exit For
        End If
    Next
    If isExist(tmpCol, tmpKey) Then
        tmpCol.Remove tmpKey
        isDetailChange_ = True
        myQuotation_.isChanged = True
    End If
End Sub

Private Sub PopulateCompany()
Dim x%
    cboCompany.Clear
    For x = 1 To CurrentUser.WorkCompanies.Count
        cboCompany.AddItem CurrentUser.WorkCompanies(x).CompanyNameTH
    Next
    If cboCompany.ListCount >= 0 Then
        cboCompany.ListIndex = 0
    End If
End Sub

Private Sub ShowQuoRemarks()
Dim x%, Y%
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
            .TextMatrix(x + rowCount, 0) = tmpDet.ID
            .TextMatrix(x + rowCount, 1) = tmpDet.RemDesc
            tmpDet.ItemNo = x
        Next
        .Row = 1
        .Col = 1
        If .Rows < 3 Then .Rows = 20
    End With
    Set tmpCol = Nothing
End Sub

Private Sub txtExtraCharge_Validate(Cancel As Boolean)
    If IsNumeric(txtExtraCharge.Text) Then
        myQuotation_.ExtraLaborFee = txtExtraCharge.Text
    Else
        myQuotation_.ExtraLaborFee = 0
    End If
End Sub

Private Sub txtHazRate_Validate(Cancel As Boolean)
    If IsNumeric(txtHazRate.Text) Then
        myQuotation_.HazMinRate = txtHazRate.Text
    Else
        myQuotation_.HazMinRate = 0
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
    If IsNumeric(txtLaborCount.Text) Then
        myQuotation_.StaffCountPerCar = txtLaborCount.Text
    Else
        myQuotation_.StaffCountPerCar = 0
    End If
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
        .ID = GetGUID
        .QuotationID = myQuotation_.ID
        .ItemNo = myQuotation_.RemarkDetails.Count + 1
        .RemDesc = selRem.RemDesc
        myQuotation_.RemarkDetails.Add tmpDet, "'" & .ID & "'"
    End With
    Set tmpDet = Nothing
    ShowQuoRemarks
End Sub

Public Sub SetDataTruckType(selTruck As BWGTruckType, Optional isAssigned As Boolean = True)
    Set curCarType_ = selTruck
    With curCarType_
        txtCarType.Text = .TypeDesc
        cboIsTrailer.Enabled = .CanBeTrailer
        If cboIsTrailer.Enabled = False Then cboIsTrailer.ListIndex = 0
        If isAssigned Then
            myQuotation_.TruckTypeID = .ID
            myQuotation_.TruckTypeName = .TypeDesc
            myQuotation_.isTrailer = (cboIsTrailer.ListIndex = 1)
        End If
    End With
End Sub

Private Sub txtNBLRate_Validate(Cancel As Boolean)
    If IsNumeric(txtNBLRate.Text) Then
        myQuotation_.NBLMinRate = txtNBLRate.Text
    Else
        myQuotation_.NBLMinRate = 0
    End If
End Sub

Private Sub txtNonRate_Validate(Cancel As Boolean)
    If IsNumeric(txtNonRate.Text) Then
        myQuotation_.NonMinRate = txtNonRate.Text
    Else
        myQuotation_.NonMinRate = 0
    End If
End Sub

Private Sub txtPosition_Validate(Cancel As Boolean)
    myQuotation_.SignPosition = txtPosition.Text
End Sub

Private Sub txtRem_Validate(Cancel As Boolean)
    myQuotation_.RemDesc = txtRem.Text
End Sub

Private Sub txtTransRate_Validate(Cancel As Boolean)
    If IsNumeric(Trim(txtTransRate.Text)) Then
        myQuotation_.TransferRate = Trim(txtTransRate.Text)
    Else
        myQuotation_.TransferRate = 0
    End If
End Sub

Private Function CheckPendingJob() As Boolean
On Error GoTo ErrD
Dim tmpRec As New ADODB.Recordset
Dim tmpJobDet As BWGJobDataDetail
Dim sqlStr As String
    sqlStr = "Select * from vw_PendingJobOrder Where QuotationID='" & myQuotation_.ID & "'"
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
            tmpJobDet.ID = tmpRec("JobDetID")
            tmpJobDet.JobNo = Trim("" & tmpRec("JobNo"))
            tmpJobDet.JobID = Trim("" & tmpRec("JobID"))
            tmpJobDet.wastedataID = tmpRec("WasteDataID")
            tmpJobDet.WasteNo = tmpRec("WasteNo")
            tmpJobDet.WasteName = tmpRec("WasteName")
            tmpJobDet.QuoTruckTypeID = tmpRec("TruckTypeID")
            'tmpJobDet.QuoTruckTypeName = tmpRec("TruckTypeDesc")
            colPendingJob_.Add tmpJobDet, "'" & tmpJobDet.ID & "'"
            Set tmpJobDet = Nothing
            tmpRec.MoveNext
        Loop
    End If
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function
