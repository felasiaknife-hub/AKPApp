VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmTransportVender 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "F128"
   ClientHeight    =   8760
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11385
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8760
   ScaleWidth      =   11385
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtVenderCode 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   9390
      TabIndex        =   32
      Text            =   "== AUTO =="
      Top             =   480
      Width           =   1935
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "เก็บข้อมูล"
      Height          =   795
      Left            =   6090
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   7860
      Width           =   1005
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "เปิดข้อมูลเก่า"
      Height          =   795
      Left            =   7140
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   7860
      Width           =   1005
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "สร้างใหม่"
      Height          =   795
      Left            =   8190
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   7860
      Width           =   1005
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "ปิด"
      Height          =   795
      Left            =   10290
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   7860
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ลบข้อมูล"
      Height          =   795
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   7860
      Width           =   1005
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6975
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   795
      Width           =   11355
      _ExtentX        =   20029
      _ExtentY        =   12303
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "ข้อมูลผู้ประกอบการขนส่งร่วม"
      TabPicture(0)   =   "frTransportVender.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "รายละเอียดพนักงาน"
      TabPicture(1)   =   "frTransportVender.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).Control(1)=   "grDriver"
      Tab(1).Control(2)=   "Label2(9)"
      Tab(1).ControlCount=   3
      Begin VB.Frame Frame3 
         Caption         =   "เอกสารที่นำมาประกอบการขึ้นทะเบียน"
         Height          =   3255
         Left            =   -74880
         TabIndex        =   44
         Top             =   3240
         Width           =   11085
         Begin VB.CheckBox chkDocTrans7 
            Appearance      =   0  'Flat
            Caption         =   "รูปถ่ายตัวอย่างรถขนส่ง ทั้ง 4 ด้าน"
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   360
            TabIndex        =   25
            Top             =   2520
            Width           =   2655
         End
         Begin VB.CheckBox chkDocTrans6 
            Appearance      =   0  'Flat
            Caption         =   "รูปถ่ายเจ้าของรถ 1 ใบ / พนักงานขับรถ 2 ใบ (ขนาด 1 นิ้ว)"
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   360
            TabIndex        =   24
            Top             =   2160
            Width           =   4575
         End
         Begin VB.CheckBox chkDocTrans5 
            Appearance      =   0  'Flat
            Caption         =   "สำเนาใบอนุญาติขับขี่ของพนักงานขับรถ"
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   360
            TabIndex        =   23
            Top             =   1800
            Width           =   3135
         End
         Begin VB.CheckBox chkDocTrans4 
            Appearance      =   0  'Flat
            Caption         =   "สำเนาทะเบียนรถ (ทุกคัน)"
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   360
            TabIndex        =   22
            Top             =   1440
            Width           =   2175
         End
         Begin VB.CheckBox chkDocTrans3 
            Appearance      =   0  'Flat
            Caption         =   "สำเนาใบอนุญาติประกอบการขนส่ง"
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   360
            TabIndex        =   21
            Top             =   1080
            Width           =   2895
         End
         Begin VB.CheckBox chkDocTrans2 
            Appearance      =   0  'Flat
            Caption         =   "สำเนาทะเบียนบ้าน"
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   360
            TabIndex        =   20
            Top             =   720
            Width           =   2895
         End
         Begin VB.CheckBox chkDocTrans1 
            Appearance      =   0  'Flat
            Caption         =   "สำเนาบัตรประจำตัวประชาชน"
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   360
            TabIndex        =   19
            Top             =   360
            Width           =   2895
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "รายละเอียดรถขนส่ง"
         Height          =   3900
         Left            =   120
         TabIndex        =   40
         Top             =   2985
         Width           =   11055
         Begin VB.TextBox txtCarSum 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   9330
            Locked          =   -1  'True
            TabIndex        =   17
            Top             =   3495
            Width           =   1215
         End
         Begin WasteManagment.ctlGrid grCarDetail 
            Height          =   3195
            Left            =   120
            TabIndex        =   16
            Top             =   270
            Width           =   10785
            _ExtentX        =   19024
            _ExtentY        =   5636
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "คัน"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   7
            Left            =   10650
            TabIndex        =   42
            Top             =   3555
            Width           =   225
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "จำนวนรถที่เข้าร่วม"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   6
            Left            =   7890
            TabIndex        =   41
            Top             =   3555
            Width           =   1320
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "เจ้าของรถ"
         Height          =   2580
         Left            =   120
         TabIndex        =   33
         Top             =   360
         Width           =   11055
         Begin VB.TextBox TxtVDProvince 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   1110
            TabIndex        =   64
            Top             =   2115
            Width           =   2340
         End
         Begin VB.TextBox TxtVDDateExp 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   7185
            TabIndex        =   62
            Top             =   1755
            Width           =   2340
         End
         Begin VB.TextBox TxtVDNumber 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   3780
            TabIndex        =   60
            Top             =   1755
            Width           =   2340
         End
         Begin VB.ComboBox CBVDType 
            Height          =   315
            ItemData        =   "frTransportVender.frx":0038
            Left            =   1095
            List            =   "frTransportVender.frx":0042
            Style           =   2  'Dropdown List
            TabIndex        =   57
            Top             =   1740
            Width           =   1635
         End
         Begin VB.ComboBox cboVAT 
            Height          =   315
            ItemData        =   "frTransportVender.frx":0058
            Left            =   9750
            List            =   "frTransportVender.frx":0062
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   1350
            Width           =   1185
         End
         Begin VB.TextBox txtWorkerCharge 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   8430
            TabIndex        =   14
            Top             =   1350
            Width           =   525
         End
         Begin VB.TextBox txtAddr 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   1110
            TabIndex        =   4
            Top             =   570
            Width           =   2415
         End
         Begin VB.TextBox txtZipCode 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   1110
            TabIndex        =   8
            Top             =   960
            Width           =   1185
         End
         Begin VB.ComboBox cboWTax 
            Height          =   315
            ItemData        =   "frTransportVender.frx":0079
            Left            =   6300
            List            =   "frTransportVender.frx":007B
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   1350
            Width           =   645
         End
         Begin VB.TextBox txtTaxID 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   9060
            TabIndex        =   3
            Top             =   180
            Width           =   1875
         End
         Begin VB.TextBox txtTransLicenseNo 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   9090
            TabIndex        =   11
            Top             =   960
            Width           =   1845
         End
         Begin VB.TextBox txtFax 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   5700
            TabIndex        =   10
            Top             =   960
            Width           =   1845
         End
         Begin VB.TextBox txtTel 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   3150
            TabIndex        =   9
            Top             =   960
            Width           =   1845
         End
         Begin VB.TextBox txtTumbol 
            Appearance      =   0  'Flat
            BackColor       =   &H00FF8080&
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   4140
            TabIndex        =   5
            Top             =   570
            Width           =   1815
         End
         Begin VB.TextBox txtAumphur 
            Appearance      =   0  'Flat
            BackColor       =   &H000000C0&
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   6570
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   570
            Width           =   1845
         End
         Begin VB.TextBox txtProvince 
            Appearance      =   0  'Flat
            BackColor       =   &H000000C0&
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   9090
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   570
            Width           =   1845
         End
         Begin VB.TextBox txtOccup 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   1110
            TabIndex        =   12
            Top             =   1350
            Width           =   3465
         End
         Begin VB.TextBox txtAge 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   6750
            TabIndex        =   2
            Top             =   180
            Width           =   825
         End
         Begin VB.TextBox txtName 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   1110
            TabIndex        =   1
            Top             =   180
            Width           =   5145
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "จังหวัด"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   24
            Left            =   495
            TabIndex        =   63
            Top             =   2175
            Width           =   465
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "วันหมดอายุ"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   23
            Left            =   6255
            TabIndex        =   61
            Top             =   1815
            Width           =   795
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "เลขที่อนุญาต"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   22
            Left            =   2850
            TabIndex        =   59
            Top             =   1815
            Width           =   885
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ประเภท"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   21
            Left            =   435
            TabIndex        =   58
            Top             =   1800
            Width           =   555
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ค่าแรงคนงาน"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   19
            Left            =   7410
            TabIndex        =   56
            Top             =   1410
            Width           =   960
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "บาท/คน"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   18
            Left            =   9030
            TabIndex        =   55
            Top             =   1410
            Width           =   585
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "เลขที่,หมู่,ถนน"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   17
            Left            =   60
            TabIndex        =   54
            Top             =   630
            Width           =   1005
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "รหัสไปรษณีย์"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   16
            Left            =   90
            TabIndex        =   53
            Top             =   1020
            Width           =   930
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "%"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   14
            Left            =   6990
            TabIndex        =   52
            Top             =   1410
            Width           =   120
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "อัตราภาษีหัก ณ. ที่จ่าย"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   15
            Left            =   4710
            TabIndex        =   51
            Top             =   1410
            Width           =   1515
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "เลขที่ผู้เสียภาษี"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   13
            Left            =   7950
            TabIndex        =   50
            Top             =   240
            Width           =   1035
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "เลขประจำตัวผู้ขนส่ง"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   8
            Left            =   7650
            TabIndex        =   49
            Top             =   1020
            Width           =   1365
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ตำบล"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   12
            Left            =   3660
            TabIndex        =   47
            Top             =   630
            Width           =   375
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "อำเภอ"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   11
            Left            =   6090
            TabIndex        =   46
            Top             =   630
            Width           =   435
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "จังหวัด"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   10
            Left            =   8580
            TabIndex        =   45
            Top             =   630
            Width           =   465
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "อาชีพ"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   20
            Left            =   630
            TabIndex        =   39
            Top             =   1410
            Width           =   390
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "โทรสาร"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   4
            Left            =   5100
            TabIndex        =   38
            Top             =   1020
            Width           =   525
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "โทรศัพท์"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   5
            Left            =   2460
            TabIndex        =   37
            Top             =   1020
            Width           =   615
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ปี"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   2
            Left            =   7710
            TabIndex        =   36
            Top             =   240
            Width           =   105
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "อายุ"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   0
            Left            =   6390
            TabIndex        =   35
            Top             =   240
            Width           =   270
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ชื่อผู้ติดต่อ"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   3
            Left            =   360
            TabIndex        =   34
            Top             =   240
            Width           =   705
         End
      End
      Begin WasteManagment.ctlGrid grDriver 
         Height          =   2355
         Left            =   -74880
         TabIndex        =   18
         Top             =   780
         Width           =   11085
         _ExtentX        =   19553
         _ExtentY        =   4154
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ข้อมูลพนักงาน :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   9
         Left            =   -74880
         TabIndex        =   43
         Top             =   480
         Width           =   1080
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "รหัส Vendor"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   1
      Left            =   8430
      TabIndex        =   48
      Top             =   540
      Width           =   870
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
      Caption         =   "ข้อมูลผู้ประกอบการขนส่งร่วม (Vender)"
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
      TabIndex        =   31
      Top             =   60
      Width           =   3435
   End
   Begin VB.Image Image2 
      Height          =   795
      Left            =   0
      Picture         =   "frTransportVender.frx":007D
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11655
   End
End
Attribute VB_Name = "frmTransportVender"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim coltruckType_ As Collection
Dim colTruckSubType_ As Collection
Dim colSelTruckSubType_ As Collection
Dim myTransportVender_ As BWGTransportVendor
Dim curTruckTypeID_ As BWGTruckSubType
Dim curTumbol_ As IMITumbol

Private Sub SetMSFlexGrid()
   
    grCarDetail.ColInputType(2) = ComboBoxEnum
    grCarDetail.ColInputType(3) = ComboBoxEnum
    grCarDetail.ColInputType(4) = ComboBoxEnum
    grCarDetail.ColInputType(5) = ComboBoxEnum
    grCarDetail.ColInputType(8) = ComboBoxEnum
    
    With grCarDetail.getGridObj
        .Cols = 9
        .ColWidth(0) = 0
        .ColWidth(1) = 1140
        .ColWidth(2) = 1605
        .ColWidth(3) = 1560
        .ColWidth(4) = 1845
        .ColWidth(5) = 1170
        .ColWidth(6) = 1170
        .ColWidth(7) = 2040
        .ColWidth(8) = 1170
        .Rows = 12
        .FixedAlignment(1) = 4
        .FixedAlignment(2) = 4
        .FixedAlignment(3) = 4
        .FixedAlignment(4) = 4
        .FixedAlignment(5) = 4
        .FixedAlignment(6) = 4
        .FixedAlignment(7) = 4
         .FixedAlignment(8) = 4
        
        .ColAlignment(1) = 1
        .ColAlignment(2) = 1
        .ColAlignment(3) = 1
        .ColAlignment(4) = 1
        .ColAlignment(5) = 1
        .ColAlignment(6) = 1
        .ColAlignment(7) = 1
        .ColAlignment(8) = 1
        
        .TextMatrix(0, 1) = "ทะเบียนรถ"
        .TextMatrix(0, 2) = "กลุ่มหลักประเภทรถ"
        .TextMatrix(0, 3) = "กลุ่มย่อยประเภทรถ"
        .TextMatrix(0, 4) = "ผู้ขับขี่"
        .TextMatrix(0, 5) = "ประเภทรับจ้าง"
        .TextMatrix(0, 6) = "หัว/หาง รถ"
        .TextMatrix(0, 7) = "หมายเหตุ"
        .TextMatrix(0, 8) = "สถานะขาย"
        
    End With
        grDriver.ColInputType(8) = ComboBoxEnum
    With grDriver.getGridObj
        .Cols = 9
        .ColWidth(0) = 0
        .ColWidth(1) = 3000
        .ColWidth(2) = 2000
        .ColWidth(3) = 0
        .ColWidth(4) = 2000
        .ColWidth(5) = 1500
        .ColWidth(6) = 1000
        .ColWidth(7) = 2500
        .ColWidth(8) = 2000
        .Rows = 10
        .FixedAlignment(1) = 4
        .FixedAlignment(2) = 4
        .FixedAlignment(3) = 4
        .FixedAlignment(4) = 4
        .FixedAlignment(5) = 4
        .FixedAlignment(6) = 4
        .FixedAlignment(7) = 4
        .FixedAlignment(8) = 4
        .ColAlignment(1) = 1
        .ColAlignment(2) = 1
        .ColAlignment(3) = 1
        .ColAlignment(4) = 1
        .ColAlignment(5) = 1
        .ColAlignment(6) = 1
        .ColAlignment(7) = 1
        .ColAlignment(8) = 1
        .TextMatrix(0, 1) = "ชื่อ - สกุล"
        .TextMatrix(0, 2) = "เบอร์โทรฯ"
        .TextMatrix(0, 3) = "ใบขับขี่ประเภท"
        .TextMatrix(0, 4) = "เลขที่ใบขับขี่"
        .TextMatrix(0, 5) = "ชนิด"
        .TextMatrix(0, 6) = "วันหมดอายุ"
        .TextMatrix(0, 7) = "เลขบัตรประชาชน"
        .TextMatrix(0, 8) = "สถานะ"
        
    End With
End Sub
Private Sub cboVAT_Validate(Cancel As Boolean)
    myTransportVender_.HasVAT = cboVAT.ListIndex
End Sub
Private Sub cboWTax_Validate(Cancel As Boolean)
    If IsNumeric(cboWTax.Text) Then
        myTransportVender_.WTaxRate = cboWTax.Text
    Else
        myTransportVender_.WTaxRate = 0
    End If
End Sub
Private Sub CBVDType_Validate(Cancel As Boolean)
    myTransportVender_.VDType = CBVDType.ListIndex
End Sub

Private Sub chkDocTrans1_Validate(Cancel As Boolean)
    myTransportVender_.DocTrans1 = (chkDocTrans1.Value = 1)
End Sub

Private Sub chkDocTrans2_Validate(Cancel As Boolean)
    myTransportVender_.DocTrans2 = (chkDocTrans2.Value = 1)
End Sub

Private Sub chkDocTrans3_Validate(Cancel As Boolean)
    myTransportVender_.DocTrans3 = (chkDocTrans3.Value = 1)
End Sub

Private Sub chkDocTrans4_Validate(Cancel As Boolean)
    myTransportVender_.DocTrans4 = (chkDocTrans4.Value = 1)
End Sub

Private Sub chkDocTrans5_Validate(Cancel As Boolean)
    myTransportVender_.DocTrans5 = (chkDocTrans5.Value = 1)
End Sub

Private Sub chkDocTrans6_Validate(Cancel As Boolean)
    myTransportVender_.DocTrans6 = (chkDocTrans6.Value = 1)
End Sub

Private Sub chkDocTrans7_Validate(Cancel As Boolean)
    myTransportVender_.DocTrans7 = (chkDocTrans7.Value = 1)
End Sub

Private Sub cmdCancel_Click()
Dim tmpStr As String
    If MsgBox("ต้องการลบข้อมูลนี้หรือไม่?", vbQuestion + vbYesNo) = vbYes Then
        tmpStr = myTransportVender_.Delete
        If Trim(tmpStr) = "" Then
            ClearScreen
            TxtName.SetFocus
        Else
            MsgBox tmpStr, vbExclamation
        End If
    End If

End Sub

Private Sub cmdClear_Click()
Dim selResult As Integer
    If myTransportVender_.isChanged Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนสร้างใหม่หรือไม่?", vbQuestion + vbYesNoCancel)
        If selResult = vbYes Then
            CmdSave_Click
            If myTransportVender_.isChanged Then Exit Sub
        ElseIf selResult = vbCancel Then
            Exit Sub
        End If
    End If
    ClearScreen
    CreateNewTransportVender
    PopulateTruckSubType

    TxtName.SetFocus

End Sub
Private Sub ClearScreen()
    txtVenderCode.Text = "== AUTO =="
    TxtName.Text = ""
    txtAge.Text = ""
    txtAddr.Text = ""
    cboVAT.ListIndex = 0
    txtZipCode.Text = ""
    txtTumbol.Text = ""
    txtAumphur.Text = ""
    txtProvince.Text = ""
    TxtTel.Text = ""
    txtFax.Text = ""
    txtTaxID.Text = ""
    txtTransLicenseNo.Text = ""
    txtOccup.Text = ""
    cboWTax.ListIndex = 0
    txtWorkerCharge.Text = ""
    TxtVDNumber.Text = ""
    TxtVDDateExp.Text = ""
    TxtVDProvince.Text = ""
    CBVDType.ListIndex = 0
    Set curTumbol_ = Nothing
    Set curTumbol_ = New IMITumbol
    Set myTransportVender_ = Nothing
    Set myTransportVender_ = New BWGTransportVendor
    Set curTruckTypeID_ = Nothing
    Set curTruckTypeID_ = New BWGTruckSubType

    grCarDetail.ClearAllData
    grCarDetail.getGridObj.Rows = 12
    grDriver.ClearAllData
    grDriver.getGridObj.Rows = 10
    chkDocTrans1.Value = 0
    chkDocTrans2.Value = 0
    chkDocTrans3.Value = 0
    chkDocTrans4.Value = 0
    chkDocTrans5.Value = 0
    chkDocTrans6.Value = 0
    chkDocTrans7.Value = 0
End Sub

Private Sub cmdClose_Click()
    Unload frmTransportVenderSearch
    Unload Me
End Sub

Private Sub CmdSave_Click()
    If Trim(TxtName.Text) = "" Then
        MsgBox "กรุณาระบุชื่อเจ้าของรถ", vbExclamation
        TxtName.SetFocus
        Exit Sub
    End If
    If Trim(txtTumbol.Text) = "" Then
        MsgBox "กรุณาระบุตำบลที่ตั้งของเจ้าของรถ", vbExclamation
        txtTumbol.SetFocus
        Exit Sub
    End If
    If Trim(txtWorkerCharge.Text) <> "" And IsNumeric(txtWorkerCharge.Text) = False Then
        MsgBox "กรุณาใส่ข้อมูลค่าแรงคนงานเป็นตัวเลขเท่านั้น", vbExclamation
        txtWorkerCharge.SetFocus
        Exit Sub
    End If
    If myTransportVender_.isChanged Then
        myTransportVender_.Save
        txtVenderCode.Text = myTransportVender_.VendorCode
        MsgBox "จัดเก็บข้อมูลเรียบร้อยแล้ว", vbExclamation
    End If
End Sub

Private Sub cmdSearch_Click()
Dim selResult As Integer
    
    If myTransportVender_.isChanged Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนทำงานต่อหรือไม่?", vbQuestion + vbYesNoCancel)
        If selResult = vbYes Then
            CmdSave_Click
            If myTransportVender_.isChanged Then Exit Sub
        ElseIf selResult = vbCancel Then
            Exit Sub
        End If
    End If
    
    Load frmTransportVenderSearch
    With frmTransportVenderSearch
        .SetCallerForm Me
        .VendorType "VC"
        .Show 1
    End With
    PopulateTruckSubType
End Sub

Private Sub Form_Load()
    CanAccess = False: CanUpdate = False
    SetPermission Me, CanAccess, CanUpdate
    If CanAccess = False Then
        MsgBox "ท่านไม่มีสิทธิใช้ระบบนี้", vbExclamation, "Warning"
        Unload Me
    End If
    
    Call SetMSFlexGrid
    cboVAT.ListIndex = 0
    With frmPicture
        cmdSave.Picture = .picSave.Picture
        cmdSearch.Picture = .picSearch.Picture
        cmdClear.Picture = .picNew.Picture
        cmdCancel.Picture = .picDelete.Picture
        cmdClose.Picture = .picExit.Picture
    End With
    With cboWTax
        .AddItem "0"
        .AddItem "1"
        .AddItem "3"
        .ListIndex = 0
    End With
    Set curTumbol_ = Nothing
    Set curTumbol_ = New IMITumbol
    Set myTransportVender_ = Nothing
    Set myTransportVender_ = New BWGTransportVendor
    Set curTruckTypeID_ = Nothing
    Set curTruckTypeID_ = New BWGTruckSubType
    
    SSTab1.Tab = 0
    CreateNewTransportVender
    PopulateTruckType
    PopulateTruckSubType
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 And Me.ActiveControl.Name <> "grCarDetail" And Me.ActiveControl.Name <> "grDriver" Then SendKeys "{Tab}"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0
End Sub
Private Sub CreateNewTransportVender()
Set myTransportVender_ = Nothing
Set myTransportVender_ = New BWGTransportVendor
    With myTransportVender_
        .ID = GetGUID
        .VendorType = "VC"
        .CreatedDate = TodayDate
        .HasVAT = cboVAT.ListIndex
        .isDeleted = False
        .WTaxRate = 0
        .isChanged = False
    End With
    SetEnabledScreen
End Sub
Private Sub SetEnabledScreen()
'Dim eFlag As Boolean
'    eFlag = Not myTransportVender_.isChanged
'    cmdApproved.Enabled = eFlag
'    cmdAddItem.Enabled = eFlag
'    cmdSave.Enabled = eFlag
'    cmdCancel.Enabled = eFlag
'    Command1.Enabled = eFlag
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload frmTransportVenderSearch
End Sub

Private Sub grCarDetail_Click(row As Integer, col As Integer)
'    MsgBox grCarDetail.getGridObj.ColWidth(col)
End Sub

Private Sub grCarDetail_KeyDown(KeyCode As Integer, Shift As Integer)
Dim selResult As Integer
Dim x%, tmpKey$, itemCount%
Dim selCol As Collection
    With grCarDetail.getGridObj
        If KeyCode = 46 And Trim(.TextMatrix(.row, 0)) <> "" Then
            selResult = MsgBox("ต้องการลบรายการรถขนส่งนี้หรือไม่?", vbQuestion + vbYesNo)
            If selResult = vbYes Then
                tmpKey = "'" & .TextMatrix(.row, 0) & "'"
                .RemoveItem .row
                Set selCol = myTransportVender_.VenderCar
                selCol.Remove tmpKey
                myTransportVender_.isChanged = True
            End If
        End If
    End With
    txtCarSum.Text = myTransportVender_.VenderCar.Count
End Sub

Private Sub grCarDetail_RolColChanged(row As Integer, col As Integer)
Dim x%
Dim tmpS As BWGSearchManager
Dim tmpCbo As ComboBox
Dim tmpKey As String
Dim tmpDet As BWGVenderCar
    
    With grCarDetail.getGridObj
        If Trim(.TextMatrix(row - 1, 1)) = "" And col > 1 Then
            grCarDetail.ColEnabled(col) = False
            Exit Sub
        Else
            If Trim(.TextMatrix(row, 1)) = "" And col > 1 Then
                grCarDetail.ColEnabled(col) = False
                Exit Sub
            Else
                grCarDetail.ColEnabled(col) = True
            End If
        End If
    End With
    
    tmpKey = "'" & grCarDetail.getGridObj.TextMatrix(row, 0) & "'"
    If isExist(myTransportVender_.VenderCar, tmpKey) Then
        Set tmpDet = myTransportVender_.VenderCar(tmpKey)
        If (col = 2) And Not coltruckType_ Is Nothing Then
            Set tmpCbo = grCarDetail.getColObject(2)
            tmpCbo.Clear
            For x = 1 To coltruckType_.Count
                tmpCbo.AddItem coltruckType_(x).TypeDesc
            Next
            Set tmpCbo = Nothing
        ElseIf (col = 3) And Not colTruckSubType_ Is Nothing Then
            Set tmpCbo = grCarDetail.getColObject(3)
            tmpCbo.Clear
            If Trim(tmpDet.TruckTypeID) = "" Then
                Set colSelTruckSubType_ = Nothing
                Set colSelTruckSubType_ = colTruckSubType_
                For x = 1 To colTruckSubType_.Count
                    tmpCbo.AddItem colTruckSubType_(x).SubTypeDesc
                Next
                Set tmpCbo = Nothing
            Else
                Set colSelTruckSubType_ = Nothing
                Set colSelTruckSubType_ = New Collection
                For x = 1 To colTruckSubType_.Count
                    If Trim(tmpDet.TruckTypeID) = Trim(colTruckSubType_(x).TruckTypeID) Then
                        colSelTruckSubType_.Add colTruckSubType_(x), "'" & colTruckSubType_(x).ID & "'"
                        tmpCbo.AddItem colTruckSubType_(x).SubTypeDesc
                    End If
                Next
            End If
            Set tmpCbo = Nothing
        ElseIf (col = 4) Then
            Set tmpCbo = grCarDetail.getColObject(4)
            tmpCbo.Clear
            For x = 1 To myTransportVender_.SupplierContact.Count
                tmpCbo.AddItem myTransportVender_.SupplierContact(x).ContactName
            Next
            Set tmpCbo = Nothing
        ElseIf col = 5 Then
            Set tmpCbo = grCarDetail.getColObject(5)
            tmpCbo.Clear
            tmpCbo.AddItem "ป้ายดำ"
            tmpCbo.AddItem "ป้ายเหลือง"
            tmpCbo.AddItem "ไม่มีข้อมูล"
            Set tmpCbo = Nothing
        ElseIf col = 6 Then
            If Trim(tmpDet.TruckTypeID) <> "" Then
                If tmpDet.CanBeTrailer Then
                    grCarDetail.ColEnabled(col) = True
                    grCarDetail.ColInputType(col) = ComboBoxEnum
                    Set tmpCbo = grCarDetail.getColObject(6)
                    tmpCbo.Clear
                    tmpCbo.AddItem "หัวรถ"
                    tmpCbo.AddItem "หางรถ"
                    Set tmpCbo = Nothing
           
                Else
                    grCarDetail.ColEnabled(col) = False
                End If
            Else
                grCarDetail.ColEnabled(col) = False
            End If
         ElseIf col = 8 Then
            Set tmpCbo = grCarDetail.getColObject(8)
            tmpCbo.Clear
            tmpCbo.AddItem "ใช้งาน"
            tmpCbo.AddItem "ขาย"
            tmpCbo.AddItem "ไม่ใช้งาน"
            Set tmpCbo = Nothing
        End If
    End If
End Sub

Private Sub PopulateTruckType()
Dim tmpS As New BWGSearchManager
    Set coltruckType_ = tmpS.TruckTypeSearch
    Set tmpS = Nothing
End Sub

Private Sub PopulateTruckSubType()
Dim tmpS As New BWGSearchManager
    Set colTruckSubType_ = tmpS.TruckSubTypeSearch
    Set tmpS = Nothing
End Sub

Private Sub grCarDetail_UpdateValue(row As Integer, col As Integer)
On Error GoTo ErrD
Dim tmpDet As BWGVenderCar
Dim tmpKey As String
Dim tmpVal As String
Dim tmpRow%

    tmpKey = Trim(grCarDetail.getGridObj.TextMatrix(row, 0))
    If isExist(myTransportVender_.VenderCar, "'" & tmpKey & "'") = False Then
        Set tmpDet = New BWGVenderCar
        tmpDet.ID = GetGUID
        tmpDet.VendorID = myTransportVender_.ID
        tmpDet.SubTypeID = row
        grCarDetail.getGridObj.TextMatrix(row, 0) = tmpDet.ID
        myTransportVender_.VenderCar.Add tmpDet, "'" & tmpDet.ID & "'"
    Else
        Set tmpDet = myTransportVender_.VenderCar("'" & tmpKey & "'")
    End If
    tmpVal = Trim(grCarDetail.getGridObj.TextMatrix(row, col))
    
    With tmpDet
        Select Case col
            Case 1
                .CarRegisID = Trim(tmpVal)
            Case 2
                
                If Trim(tmpVal) <> "" Then
                    .TruckTypeID = coltruckType_(grCarDetail.getColObject(2).ListIndex + 1).ID
                    .TruckTypeDesc = Trim(tmpVal)
                    .CanBeTrailer = coltruckType_(grCarDetail.getColObject(2).ListIndex + 1).CanBeTrailer
                    If .CanBeTrailer = False Then
                        grCarDetail.getGridObj.TextMatrix(row, 6) = ""
                        .headOrTrail = ""
                    End If
                Else
                    .TruckTypeID = ""
                    .TruckTypeDesc = ""
                End If
            Case 3
                If Trim(tmpVal) <> "" Then
                    .SubTypeID = colSelTruckSubType_(grCarDetail.getColObject(3).ListIndex + 1).ID
                    .SubTypeDesc = Trim(tmpVal)
                    If Trim(.TruckTypeID) = "" Then
                        .TruckTypeID = colSelTruckSubType_(grCarDetail.getColObject(3).ListIndex + 1).TruckTypeID
                        .TruckTypeDesc = colSelTruckSubType_(grCarDetail.getColObject(3).ListIndex + 1).TypeDesc
                        grCarDetail.getGridObj.TextMatrix(row, 2) = .TruckTypeDesc
                    End If
                Else
                    .SubTypeID = ""
                    .SubTypeDesc = ""
                End If
            Case 4
                If Trim(tmpVal) <> "" Then
                    .ContactID = myTransportVender_.SupplierContact(grCarDetail.getColObject(4).ListIndex + 1).ID
                    .ContactName = Trim(tmpVal)
                Else
                    .ContactID = ""
                    .ContactName = ""
                End If
            Case 5
                If grCarDetail.getColObject(5).ListIndex = 0 Then
                    .LicenseType = "B"
                ElseIf grCarDetail.getColObject(5).ListIndex = 1 Then
                    .LicenseType = "Y"
                ElseIf grCarDetail.getColObject(5).ListIndex = 2 Then
                    .LicenseType = ""
                End If
            Case 6
                If grCarDetail.getColObject(6).ListIndex = 0 Then
                    .headOrTrail = "H"
                ElseIf grCarDetail.getColObject(6).ListIndex = 1 Then
                    .headOrTrail = "T"
                End If
            Case 7
                .Remark = Trim(tmpVal)
            Case 8
                If grCarDetail.getColObject(8).ListIndex = 0 Then
                    .IsSell = "U"
                ElseIf grCarDetail.getColObject(8).ListIndex = 1 Then
                    .IsSell = "Y"
                ElseIf grCarDetail.getColObject(8).ListIndex = 2 Then
                    .IsSell = "N"
                End If
        End Select
            .TruckCount = 1
            
        myTransportVender_.isChanged = True
    End With
    txtCarSum.Text = myTransportVender_.VenderCar.Count
    Set tmpDet = Nothing
    Exit Sub
ErrD:
    MsgBox err.Description
    Exit Sub
End Sub

Private Sub grDriver_KeyDown(KeyCode As Integer, Shift As Integer)
Dim selResult As Integer
Dim x%, tmpKey$, itemCount%
Dim selCol As Collection
    With grDriver.getGridObj
        If KeyCode = 46 And Trim(.TextMatrix(.row, 0)) <> "" Then
            For x = 1 To myTransportVender_.VenderCar.Count
                If myTransportVender_.VenderCar(x).ContactID = Trim(.TextMatrix(.row, 0)) Then
                    MsgBox "ไม่สามารถลบพนักงานคนนี้ได้เนื่องจากถูกกำหนดให้เป็นพนักงานขับรถ", vbExclamation
                    Exit Sub
                End If
            Next
        
            selResult = MsgBox("ต้องการลบข้อมูลพนักงานนี้หรือไม่?", vbQuestion + vbYesNo)
            If selResult = vbYes Then
                tmpKey = "'" & .TextMatrix(.row, 0) & "'"
                .RemoveItem .row
                Set selCol = myTransportVender_.SupplierContact
                selCol.Remove tmpKey
                itemCount = myTransportVender_.SupplierContact.Count
                myTransportVender_.isChanged = True
            End If
        End If

    End With
End Sub

Private Sub grDriver_RolColChanged(row As Integer, col As Integer)
Dim tmpCbo As ComboBox
        If col = 8 Then
            Set tmpCbo = grDriver.getColObject(8)
            tmpCbo.Clear
            tmpCbo.AddItem "อยู่"
            tmpCbo.AddItem "ลาออก"
            Set tmpCbo = Nothing
        Else
            grDriver.ColEnabled(col) = False
        End If
        
End Sub

Private Sub grDriver_UpdateValue(row As Integer, col As Integer)
Dim tmpDet As BWGSupplierContact
Dim tmpKey As String
Dim tmpVal As String
    tmpKey = Trim(grDriver.getGridObj.TextMatrix(row, 0))
    If isExist(myTransportVender_.SupplierContact, "'" & tmpKey & "'") = False Then
        Set tmpDet = New BWGSupplierContact
        tmpDet.ID = GetGUID
        tmpDet.VendorID = myTransportVender_.ID
        'tmpDet.r = row
        grDriver.getGridObj.TextMatrix(row, 0) = tmpDet.ID
        myTransportVender_.SupplierContact.Add tmpDet, "'" & tmpDet.ID & "'"
    Else
        Set tmpDet = myTransportVender_.SupplierContact("'" & tmpKey & "'")
    End If
    tmpVal = Trim(grDriver.getGridObj.TextMatrix(row, col))
    With tmpDet
        Select Case col
            Case 1
                .ContactName = Trim(tmpVal)
            Case 2
                .TelNo = Trim(tmpVal)
            Case 3
                .CarType = Trim(tmpVal)
            Case 4
                .DVNumber = Trim(tmpVal)
            Case 5
                .DVType = Trim(tmpVal)
            Case 6
                .DVDateExp = Trim(tmpVal)
            Case 7
                .DVPNumber = Trim(tmpVal)
            Case 8
            If tmpVal = "อยู่" Then
                .Stat = "Y"
                Else
                .Stat = "N"
            End If
        End Select
        myTransportVender_.isChanged = True
    End With
    Set tmpDet = Nothing
End Sub

Private Sub txtAddr_Validate(Cancel As Boolean)
    myTransportVender_.Addr = Trim(txtAddr.Text)
End Sub

Private Sub txtAge_Validate(Cancel As Boolean)
    If IsNumeric(txtAge.Text) Then
        myTransportVender_.Age = txtAge.Text
    Else
        myTransportVender_.Age = 0
    End If
End Sub

Private Sub txtFax_Validate(Cancel As Boolean)
    myTransportVender_.FaxNo = txtFax.Text
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
    myTransportVender_.NameTH = TxtName.Text
End Sub

Private Sub txtOccup_Validate(Cancel As Boolean)
    myTransportVender_.Occupancy = txtOccup.Text
End Sub

Private Sub txtTaxID_Validate(Cancel As Boolean)
    myTransportVender_.TaxID = txtTaxID.Text
End Sub

Private Sub txtTel_Validate(Cancel As Boolean)
    myTransportVender_.TelNo = TxtTel.Text
End Sub

Private Sub txtTransLicenseNo_Validate(Cancel As Boolean)
    myTransportVender_.TransLicenseNo = txtTransLicenseNo.Text
End Sub

Private Sub txtTumbol_Validate(Cancel As Boolean)
    If Trim(txtTumbol.Text) = "" And Trim(curTumbol_.Name) <> Trim(txtTumbol.Text) Then
        Set curTumbol_ = Nothing
        Set curTumbol_ = New IMITumbol
        txtAumphur.Text = ""
        txtProvince.Text = ""
        With myTransportVender_
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
            With myTransportVender_
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
            If myTransportVender_.TumbolID = "" Then txtTumbol.Text = ""
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
            myTransportVender_.AumphurID = .AumphurID
            myTransportVender_.AumphurName = .AumphurName
            myTransportVender_.TumbolID = .ID
            myTransportVender_.TumbolName = .Name
            myTransportVender_.ProvinceID = .ProvinceID
            myTransportVender_.ProvinceName = .ProvinceName
        End If
    End With
End Sub

Public Sub SetDataTransportVender(selCust As BWGTransportVendor, Optional isAssigned As Boolean = True)
Dim tmpS As New BWGSearchManager
Dim tmpCol As Collection
        
        Set myTransportVender_ = selCust
        With myTransportVender_
            TxtName.Text = .NameTH
            txtAge.Text = .Age
            TxtTel.Text = .TelNo
            txtFax.Text = .FaxNo
            txtAddr.Text = .Addr
            txtZipCode.Text = .ZipCode
            cboVAT.ListIndex = .HasVAT
            txtOccup.Text = .Occupancy
            txtVenderCode.Text = .VendorCode
            txtTaxID.Text = .TaxID
            txtTransLicenseNo.Text = .TransLicenseNo
            If Trim(.VDType) <> "" Then CBVDType.ListIndex = .VDType
            TxtVDNumber.Text = .VDNumber
            TxtVDDateExp.Text = .VDDateExp
            TxtVDProvince.Text = .VDProvince
            SetTextToCombo .WTaxRate, cboWTax
            If .WorkerCharge > 0 Then txtWorkerCharge.Text = .WorkerCharge
            Set tmpCol = tmpS.TumbolSearch("TumbolID='" & .TumbolID & "'")
            If tmpCol.Count > 0 Then
                SetDataTumbol tmpCol(1), False
                txtTumbol.Text = .TumbolName
                txtAumphur.Text = .AumphurName
                txtProvince.Text = .ProvinceName
            End If
            
            If .DocTrans1 Then
                chkDocTrans1.Value = 1
            Else
                chkDocTrans1.Value = 0
            End If
            If .DocTrans2 Then
                chkDocTrans2.Value = 1
            Else
                chkDocTrans2.Value = 0
            End If
            If .DocTrans3 Then
                chkDocTrans3.Value = 1
            Else
                chkDocTrans3.Value = 0
            End If
            If .DocTrans4 Then
                chkDocTrans4.Value = 1
            Else
                chkDocTrans4.Value = 0
            End If
            If .DocTrans5 Then
                chkDocTrans5.Value = 1
            Else
                chkDocTrans5.Value = 0
            End If
            If .DocTrans6 Then
                chkDocTrans6.Value = 1
            Else
                chkDocTrans6.Value = 0
            End If
            If .DocTrans7 Then
                chkDocTrans7.Value = 1
            Else
                chkDocTrans7.Value = 0
            End If
        End With
        ShowVenderCar
        ShowSupplierContact
End Sub

Private Sub ShowVenderCar()
Dim x%, Y%
Dim tmpCol As Collection
Dim tmpDet As BWGVenderCar
Dim tmpKey As String, rowCount As Long
Dim forColor As String
    With grCarDetail.getGridObj
        grCarDetail.ClearAllData
                
        Set tmpCol = myTransportVender_.VenderCar
        rowCount = 0
        forColor = vbBlack
        If .Rows - 1 <= tmpCol.Count Then .Rows = tmpCol.Count + 20
        For x = 1 To tmpCol.Count
            Set tmpDet = tmpCol(x)
            .TextMatrix(x + rowCount, 0) = tmpDet.ID
            .row = x + rowCount
            .col = 1
            .CellForeColor = forColor
            .TextMatrix(x + rowCount, 1) = tmpDet.CarRegisID
            .TextMatrix(x + rowCount, 2) = tmpDet.TruckTypeDesc
            .TextMatrix(x + rowCount, 3) = tmpDet.SubTypeDesc
            .TextMatrix(x + rowCount, 4) = tmpDet.ContactName
            If tmpDet.LicenseType = "B" Then
                .TextMatrix(x + rowCount, 5) = "ป้ายดำ"
            ElseIf tmpDet.LicenseType = "Y" Then
                .TextMatrix(x + rowCount, 5) = "ป้ายเหลือง"
            ElseIf tmpDet.LicenseType = "" Then
                .TextMatrix(x + rowCount, 5) = "ไม่มีข้อมูล"
            End If
            If tmpDet.headOrTrail = "H" Then
                .TextMatrix(x + rowCount, 6) = "หัวรถ"
            ElseIf tmpDet.headOrTrail = "T" Then
                .TextMatrix(x + rowCount, 6) = "หางรถ"
            Else
                .TextMatrix(x + rowCount, 6) = ""
            End If
            .TextMatrix(x + rowCount, 7) = tmpDet.Remark
               If tmpDet.IsSell = "U" Then
                .TextMatrix(x + rowCount, 8) = "ใช้งาน"
            ElseIf tmpDet.IsSell = "Y" Then
                .TextMatrix(x + rowCount, 8) = "ขาย"
            ElseIf tmpDet.IsSell = "N" Then
                .TextMatrix(x + rowCount, 8) = "ไม่ใช้งาน"
            Else
               .TextMatrix(x + rowCount, 8) = ""
                End If
            txtCarSum.Text = .row
        Next
        .row = 1
        .col = 1
    End With
    Set tmpCol = Nothing
End Sub

Private Sub ShowSupplierContact()
Dim x%, Y%
Dim tmpCol As Collection
Dim tmpDet As BWGSupplierContact
Dim tmpKey As String, rowCount As Long
Dim forColor As String
    With grDriver.getGridObj
        grDriver.ClearAllData
                
        Set tmpCol = myTransportVender_.SupplierContact
        rowCount = 0
        forColor = vbBlack
        If .Rows - 1 <= tmpCol.Count Then .Rows = tmpCol.Count + 20
        For x = 1 To tmpCol.Count
            Set tmpDet = tmpCol(x)
            .TextMatrix(x + rowCount, 0) = tmpDet.ID
            .row = x + rowCount
            .col = 1
            .CellForeColor = forColor
            .TextMatrix(x + rowCount, 1) = tmpDet.ContactName
            .TextMatrix(x + rowCount, 2) = tmpDet.TelNo
            .TextMatrix(x + rowCount, 3) = tmpDet.CarType
            .TextMatrix(x + rowCount, 4) = tmpDet.DVNumber
            .TextMatrix(x + rowCount, 5) = tmpDet.DVType
            .TextMatrix(x + rowCount, 6) = tmpDet.DVDateExp
            .TextMatrix(x + rowCount, 7) = tmpDet.DVPNumber
            If tmpDet.Stat = "N" Then
            .TextMatrix(x + rowCount, 8) = "ลาออก"
            Else
            .TextMatrix(x + rowCount, 8) = "อยู่"
            End If
        Next
        .row = 1
        .col = 1
    End With
    Set tmpCol = Nothing
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim selResult As Integer
    If myTransportVender_ Is Nothing Then Exit Sub
    If myTransportVender_.isChanged Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนปิดหน้าจอหรือไม่?", vbQuestion + vbYesNoCancel)
        If selResult = vbYes Then
            CmdSave_Click
            If myTransportVender_.isChanged Then Cancel = 1
        ElseIf selResult = vbCancel Then
            Cancel = 1
        End If
    End If
End Sub
Private Sub TxtVDDateExp_Validate(Cancel As Boolean)
    myTransportVender_.VDDateExp = Trim(TxtVDDateExp)
End Sub
Private Sub TxtVDNumber_Validate(Cancel As Boolean)
        myTransportVender_.VDNumber = Trim(TxtVDNumber.Text)
End Sub
Private Sub TxtVDProvince_Validate(Cancel As Boolean)
    myTransportVender_.VDProvince = Trim(TxtVDProvince.Text)
End Sub
Private Sub txtWorkerCharge_Validate(Cancel As Boolean)
    If IsNumeric(txtWorkerCharge.Text) Then
        myTransportVender_.WorkerCharge = txtWorkerCharge.Text
    Else
        myTransportVender_.WorkerCharge = 0
    End If
End Sub

Private Sub txtZipCode_Validate(Cancel As Boolean)
    myTransportVender_.ZipCode = Trim(txtZipCode.Text)
End Sub
