VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmInvoice 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " "
   ClientHeight    =   9105
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11775
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9105
   ScaleWidth      =   11775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdInv 
      Caption         =   "ออก INV"
      Height          =   795
      Left            =   1575
      Style           =   1  'Graphical
      TabIndex        =   65
      Top             =   8280
      Width           =   945
   End
   Begin VB.CheckBox ChkSC 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "ส่งให้ลูกค้าได้"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   5490
      TabIndex        =   64
      Top             =   1260
      Width           =   1230
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1815
      Left            =   3120
      TabIndex        =   58
      Top             =   3420
      Width           =   5355
      Begin VB.Label lblCancelDetail 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1515
         Left            =   60
         TabIndex        =   59
         Top             =   180
         Width           =   5235
      End
   End
   Begin VB.TextBox txtHeaderText 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1200
      TabIndex        =   56
      Text            =   "ค่าบริการกำจัดและบำบัด"
      Top             =   3060
      Width           =   4275
   End
   Begin VB.CheckBox CheckVAT 
      BackColor       =   &H00FFFFFF&
      Caption         =   "ไม่คิดภาษีมูลค่าเพิ่ม"
      Height          =   195
      Left            =   6840
      TabIndex        =   55
      Top             =   2400
      Width           =   2115
   End
   Begin VB.TextBox txtBillName 
      Appearance      =   0  'Flat
      BackColor       =   &H00C00000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   8760
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1620
      Width           =   2895
   End
   Begin VB.TextBox txtCheqDesc 
      Appearance      =   0  'Flat
      BackColor       =   &H00C00000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1200
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2340
      Width           =   5535
   End
   Begin VB.TextBox txtBillDesc 
      Appearance      =   0  'Flat
      BackColor       =   &H00C00000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1200
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1980
      Width           =   10455
   End
   Begin VB.ComboBox cboWTax 
      Height          =   315
      ItemData        =   "frCollectAR.frx":0000
      Left            =   10680
      List            =   "frCollectAR.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   2340
      Width           =   735
   End
   Begin VB.ComboBox cboSiteName 
      Height          =   315
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   1620
      Width           =   7515
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "เลือกใบ Menifest"
      Height          =   795
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   8280
      Width           =   1455
   End
   Begin VB.CommandButton cmdReceive 
      Caption         =   "บันทึกรับเงินจากลูกค้า"
      Height          =   795
      Left            =   2565
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   8280
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.CheckBox chkAccepted 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "ลูกค้ารับใบวางบิลแล้ว"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   6840
      TabIndex        =   4
      Top             =   1260
      Width           =   1815
   End
   Begin VB.TextBox txtCustomer 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1200
      TabIndex        =   3
      Top             =   1260
      Width           =   4005
   End
   Begin WasteManagment.ctlDate dtBilldate 
      Height          =   345
      Left            =   3435
      TabIndex        =   1
      Top             =   870
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   609
      BackColor       =   -2147483633
   End
   Begin VB.TextBox txtRem 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6840
      MaxLength       =   80
      TabIndex        =   12
      Top             =   2700
      Width           =   4785
   End
   Begin VB.TextBox txtRecBillStaff 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1200
      TabIndex        =   11
      Top             =   2700
      Width           =   4275
   End
   Begin Crystal.CrystalReport rpt1 
      Left            =   2520
      Top             =   -30
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin VB.ComboBox cboPrinter 
      Height          =   315
      Left            =   8190
      Style           =   2  'Dropdown List
      TabIndex        =   37
      Top             =   30
      Width           =   3525
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "เก็บข้อมูล"
      Height          =   795
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   8280
      Width           =   1035
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "เปิดข้อมูลเก่า"
      Height          =   795
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   8280
      Width           =   1095
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "สร้างใบใหม่"
      Height          =   795
      Left            =   7380
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   8280
      Width           =   1095
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "ปิด"
      Height          =   795
      Left            =   10680
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   8280
      Width           =   1035
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "พิมพ์"
      Height          =   795
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   8280
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ยกเลิก"
      Height          =   795
      Left            =   9600
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   8280
      Width           =   1035
   End
   Begin VB.ComboBox cboCompany 
      Height          =   315
      Left            =   6840
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   870
      Width           =   4785
   End
   Begin VB.TextBox txtBillNo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Text            =   "== AUTO =="
      Top             =   900
      Width           =   1740
   End
   Begin WasteManagment.ctlDate dtDueDate 
      Height          =   345
      Left            =   9840
      TabIndex        =   5
      Top             =   1230
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   609
      BackColor       =   -2147483633
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "ค่ากำจัดจากใบ Menifest ที่เรียกเก็บเงินในงวดนี้"
      ForeColor       =   &H80000008&
      Height          =   2475
      Index           =   0
      Left            =   30
      TabIndex        =   27
      Top             =   3420
      Width           =   11715
      Begin WasteManagment.ctlGrid grMenifest 
         Height          =   2205
         Left            =   60
         TabIndex        =   13
         Top             =   240
         Width           =   11595
         _ExtentX        =   20452
         _ExtentY        =   3889
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "รายการอื่นๆ ที่เรียกเก็บเพิ่มเติม"
      ForeColor       =   &H80000008&
      Height          =   1815
      Index           =   2
      Left            =   30
      TabIndex        =   44
      Top             =   5970
      Width           =   6165
      Begin WasteManagment.ctlGrid grItem 
         Height          =   1485
         Left            =   60
         TabIndex        =   14
         Top             =   240
         Width           =   6045
         _ExtentX        =   10663
         _ExtentY        =   2619
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "ค่าขนส่ง"
      ForeColor       =   &H80000008&
      Height          =   1815
      Index           =   1
      Left            =   6180
      TabIndex        =   53
      Top             =   5970
      Width           =   5565
      Begin MSComctlLib.ListView lvTrip 
         Height          =   1515
         Left            =   60
         TabIndex        =   54
         Top             =   240
         Width           =   5445
         _ExtentX        =   9604
         _ExtentY        =   2672
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "เลขที่ใบคุม"
            Object.Width           =   3360
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ทะเบียนรถ"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "ยอดเงิน"
            Object.Width           =   2646
         EndProperty
      End
   End
   Begin WasteManagment.ctlDate ctlDatepast 
      Height          =   345
      Left            =   9720
      TabIndex        =   63
      Top             =   3120
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   609
      BackColor       =   -2147483633
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "วันที่เข้าไปวางบิล"
      Height          =   195
      Index           =   16
      Left            =   8520
      TabIndex        =   62
      Top             =   3120
      Width           =   1185
   End
   Begin VB.Label lblLastUpdateRE 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00C0E0FF&
      Caption         =   "แก้ไขล่าสุดโดย :"
      Height          =   195
      Left            =   8220
      TabIndex        =   61
      Top             =   600
      Width           =   3480
   End
   Begin VB.Label lblLastUpdate 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Caption         =   "แก้ไขล่าสุดโดย :"
      Height          =   195
      Left            =   8220
      TabIndex        =   60
      Top             =   420
      Width           =   3480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "หัวบิล"
      Height          =   195
      Index           =   14
      Left            =   420
      TabIndex        =   57
      Top             =   3120
      Width           =   390
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "การเก็บเช็ค"
      Height          =   195
      Index           =   13
      Left            =   300
      TabIndex        =   52
      Top             =   2400
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "การวางบิล"
      Height          =   195
      Index           =   12
      Left            =   420
      TabIndex        =   51
      Top             =   2040
      Width           =   735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TAX"
      Height          =   195
      Index           =   11
      Left            =   3690
      TabIndex        =   50
      Top             =   7890
      Width           =   315
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "บาท"
      Height          =   195
      Index           =   10
      Left            =   5670
      TabIndex        =   49
      Top             =   7890
      Width           =   285
   End
   Begin VB.Label lbWTax 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   4230
      TabIndex        =   48
      Top             =   7890
      Width           =   1245
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ภาษีหัก ณ. ที่จ่าย"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   15
      Left            =   9420
      TabIndex        =   47
      Top             =   2400
      Width           =   1155
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   14
      Left            =   11520
      TabIndex        =   46
      Top             =   2400
      Width           =   135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "สถานที่วางบิล"
      Height          =   195
      Index           =   9
      Left            =   180
      TabIndex        =   45
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ลูกหนี้"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   8
      Left            =   660
      TabIndex        =   43
      Top             =   1320
      Width           =   435
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "หมายเหตุ"
      Height          =   195
      Index           =   7
      Left            =   6090
      TabIndex        =   42
      Top             =   2760
      Width           =   675
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "กำหนดรับเงิน"
      Height          =   195
      Index           =   18
      Left            =   8790
      TabIndex        =   41
      Top             =   1290
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ผู้รับวางบิล"
      Height          =   195
      Index           =   15
      Left            =   420
      TabIndex        =   40
      Top             =   2760
      Width           =   735
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
      Left            =   4140
      TabIndex        =   39
      Top             =   0
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "เลือกเครื่องพิมพ์"
      Height          =   195
      Index           =   8
      Left            =   6960
      TabIndex        =   38
      Top             =   90
      Width           =   1140
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   11880
      Y1              =   8220
      Y2              =   8220
   End
   Begin VB.Label lbVat 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   6570
      TabIndex        =   36
      Top             =   7890
      Width           =   1485
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "บาท"
      Height          =   195
      Index           =   21
      Left            =   8160
      TabIndex        =   35
      Top             =   7890
      Width           =   285
   End
   Begin VB.Label lbTotalPrice 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   1080
      TabIndex        =   34
      Top             =   7890
      Width           =   1725
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "บาท"
      Height          =   195
      Index           =   20
      Left            =   2910
      TabIndex        =   33
      Top             =   7890
      Width           =   285
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "มูลค่าบริการ"
      Height          =   195
      Index           =   19
      Left            =   60
      TabIndex        =   32
      Top             =   7890
      Width           =   840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VAT"
      Height          =   195
      Index           =   6
      Left            =   6060
      TabIndex        =   31
      Top             =   7890
      Width           =   345
   End
   Begin VB.Label lbNetPrice 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   9420
      TabIndex        =   30
      Top             =   7890
      Width           =   1875
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "บาท"
      Height          =   195
      Index           =   5
      Left            =   11430
      TabIndex        =   29
      Top             =   7890
      Width           =   285
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "รวมทั้งสิ้น"
      Height          =   195
      Index           =   4
      Left            =   8550
      TabIndex        =   28
      Top             =   7890
      Width           =   675
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ใบวางบิล"
      BeginProperty Font 
         Name            =   "Angsana New"
         Size            =   18
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   180
      TabIndex        =   26
      Top             =   90
      Width           =   915
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      Index           =   1
      X1              =   240
      X2              =   7050
      Y1              =   630
      Y2              =   630
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   4
      Index           =   0
      X1              =   150
      X2              =   6960
      Y1              =   540
      Y2              =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "เรียกเก็บในนาม"
      Height          =   195
      Index           =   1
      Left            =   5580
      TabIndex        =   25
      Top             =   930
      Width           =   1170
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "วันที่"
      Height          =   195
      Index           =   2
      Left            =   3015
      TabIndex        =   24
      Top             =   960
      Width           =   315
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "เลขที่ใบวางบิล"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   23
      Top             =   960
      Width           =   1035
   End
   Begin VB.Image Image2 
      Height          =   795
      Left            =   0
      Picture         =   "frCollectAR.frx":0004
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12030
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFFF80&
      Height          =   315
      Index           =   3
      Left            =   9300
      Top             =   7830
      Width           =   2055
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFFF80&
      Height          =   315
      Index           =   2
      Left            =   6450
      Top             =   7830
      Width           =   1665
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFFF80&
      Height          =   315
      Index           =   1
      Left            =   960
      Top             =   7830
      Width           =   1905
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFFF80&
      Height          =   315
      Index           =   0
      Left            =   4110
      Top             =   7830
      Width           =   1425
   End
End
Attribute VB_Name = "frmInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim curCustomer_ As BWGCustomer
Dim myBillHeader_ As BWGBillingHeader
Dim isDetailChange_ As Boolean
Dim CanAccess As Boolean, CanUpdate As Boolean
Dim pageCount As Integer, CurPage As Integer, MaxPrintPage As Integer
Dim PageContinue As Boolean
Dim ResumePrint As Boolean
Dim X1 As Integer, X2 As Double
            
Private Sub cboCompany_Validate(Cancel As Boolean)
    If cboCompany.ListIndex >= 0 Then
        If cboCompany.ListCount > 1 Then
            If cboCompany.ListIndex = 0 Then Exit Sub
            myBillHeader_.CompanyID = CurrentUser.WorkCompanies(cboCompany.ListIndex).ID
        Else
            myBillHeader_.CompanyID = CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).ID
        End If
        myBillHeader_.CompanyName = cboCompany.Text
    Else
        myBillHeader_.CompanyID = ""
        myBillHeader_.CompanyName = ""
    End If
    If myBillHeader_.CompanyID = "3F764F1713974948A42CC0FAF926C730" Then
        txtHeaderText.Text = "ค่าบริการ"
    Else
        txtHeaderText.Text = "ค่าบริการกำจัดและบำบัด"
    End If
End Sub

Private Sub cboSiteName_Click()
    If cboSiteName.ListIndex > -1 Then
        txtBillName.Text = curCustomer_.BillAddresses(cboSiteName.ListIndex + 1).BillName
    Else
        txtBillName.Text = ""
    End If
End Sub

Private Sub cboSiteName_Validate(Cancel As Boolean)
    If cboSiteName.ListIndex = -1 Then
        myBillHeader_.BillAddr = ""
        myBillHeader_.BillAddrID = ""
    Else
        myBillHeader_.BillAddr = cboSiteName.Text
        myBillHeader_.BillAddrID = curCustomer_.BillAddresses(cboSiteName.ListIndex + 1).ID
    End If
End Sub

Private Sub cboWTax_Click()
    If Me.Visible Then
        If Me.ActiveControl.Name = "cboWTax" Then
            myBillHeader_.WTaxRate = Replace(cboWTax.Text, "%", "")
            SumMoneyAmt
        End If
    End If
End Sub

Private Sub CheckVAT_Click()
    If CheckVAT.Value = 1 Then
        lbNetPrice = Format(lbNetPrice - lbVat, "#,#00.00")
        lbVat.Caption = "0.00"
    Else
        lbVat.Caption = Format(CDbl(lbTotalPrice) * SystemConfig.VATRate / 100, "#,#00.00")
        lbNetPrice = Format(CDbl(lbNetPrice) + CDbl(lbVat), "#,#00.00")
    End If
    myBillHeader_.IncVAT = CheckVAT.Value
End Sub

Private Sub chkAccepted_Validate(Cancel As Boolean)
    If chkAccepted.Value = 1 Then
        myBillHeader_.isCustAccepted = True
    Else
        myBillHeader_.isCustAccepted = False
    End If
End Sub

Private Sub ChkSC_Click()
    If ChkSC.Value = 1 Then
        myBillHeader_.isServiceCustomer = True
    Else
        myBillHeader_.isServiceCustomer = False
    End If

End Sub

Private Sub cmdCancel_Click()
Dim tmpStr As String, CancelDetail As String
    If MsgBox("ท่านต้องการยกเลิกใบวางบิลนี้หรือไม่?", vbQuestion + vbYesNo) = vbYes Then
        CancelDetail = Trim(frmCancelDetail.CancelText(myBillHeader_.CancelDetail))
        If Trim(Replace(CancelDetail, "*", "")) = "" Then
            MsgBox "กรุณาระบุเหตุผลของการยกเลิก", vbExclamation + vbOKOnly, "Warning"
            Exit Sub
        End If
        myBillHeader_.CancelDetail = CancelDetail
        tmpStr = myBillHeader_.CancelBill(Trim(CurrentUser.StaffFName & " " & CurrentUser.StaffLName), CancelDetail)
        
        If Trim(tmpStr) = "" Then
            SetCancelScreen
        Else
            MsgBox tmpStr
        End If
    End If
End Sub

Private Sub cmdClear_Click()
Dim selResult As Integer
    If myBillHeader_.isChanged And cmdSave.Enabled Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนสร้างใหม่หรือไม่?", vbQuestion + vbYesNoCancel)
        If selResult = vbYes Then
            CmdSave_Click
            If myBillHeader_.isChanged Then Exit Sub
        ElseIf selResult = vbCancel Then
            Exit Sub
        End If
    End If
    ClearScreen
    CreateNewBillHeader
    txtCustomer.SetFocus
End Sub

Private Sub cmdClose_Click()
    If frmInvoiceSearch.Appearance Then
        Unload frmInvoiceSearch
    End If
    Unload Me
End Sub

Public Sub SetDataCustomer(selCust As BWGCustomer, Optional isAssigned As Boolean = True)
Dim tmpRec As New ADODB.Recordset
Dim x%
    Set curCustomer_ = selCust
    cboSiteName.Clear
    With curCustomer_
        For x = 1 To .BillAddresses.Count
            cboSiteName.AddItem .BillAddresses(x).Address & " " & .BillAddresses(x).TumbolName & " " & .BillAddresses(x).AumphurName & " " & .BillAddresses(x).ProvinceName & " " & .BillAddresses(x).ZipCode
        Next
        
        txtCustomer.Text = .CustomerName
        txtBillDesc.Text = .BillingInfo.getBillingDesc
        txtCheqDesc.Text = .BillingInfo.getCheqDesc
        If isAssigned Then
            myBillHeader_.CustomerID = .ID
            myBillHeader_.CustomerName = .CustomerName
            If cboSiteName.ListCount = 1 Then
                cboSiteName.ListIndex = 0
                myBillHeader_.BillAddrID = curCustomer_.BillAddresses(1).ID
                myBillHeader_.BillAddr = cboSiteName.Text
            End If
        End If
    End With
End Sub
Private Sub CmdInv_Click()
Dim tmpRec As New ADODB.Recordset
    If myBillHeader_.IsInv <> "Y" Then
        myBillHeader_.InvDate = Trim(dtBillDate.ValueYMD)
        DBConnExc "update tbBillingHeader set invdate= '" & Trim(dtBillDate.ValueYMD) & "' where BillingHeaderID='" & myBillHeader_.ID & "'"
        myBillHeader_.isChanged = True
        DBConnExc "Insert into tbRunBillingNoNew (BillingHeaderID,CompanyID,custtype) Values ('" & Trim(myBillHeader_.ID) & "','" & Trim(myBillHeader_.CompanyID) & "','" & Trim(curCustomer_.CustomerType) & "')"
            sqlStr = "SELECT BillingNo FROM tbBillingHeader WHERE BillingHeaderID='" & myBillHeader_.ID & "' "
            Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
                If tmpRec.RecordCount > 0 Then
                    txtBillNo.Text = tmpRec!BillingNo
                    Set tmpRec = Nothing
                End If
    End If
End Sub

Private Sub cmdOpen_Click()
On Error Resume Next
Dim selResult As Integer
    If myBillHeader_.isChanged And cmdSave.Enabled Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนทำงานต่อหรือไม่?", vbQuestion + vbYesNoCancel)
        If selResult = vbYes Then
            CmdSave_Click
            If myBillHeader_.isChanged Then Exit Sub
        ElseIf selResult = vbCancel Then
            Exit Sub
        End If
    End If
    Load frmInvoiceSearch
    With frmInvoiceSearch
        .SetCallerForm Me
        If Label1(3).Caption = "Pre INV" Then
            .AumphurSearch(7).Caption = "ค้นหา PreInv"
        End If
        .Show 1
    End With
End Sub

Private Sub cmdPrint_Click()
Dim tmpGrid As MSFlexGrid, curTripID$, selResult%, rowPerPage%, x%, detailCount%, curRow%, curRecNo%
Dim curTimeTableNo$, y%, curCode$
Dim tmpTrip As BWGJobDataTimeTable, tmpSite As BWGCustomerBillAddr, tmpList As ListItem
    If Trim(txtCustomer.Text) = "" Then
        MsgBox "กรุณาระบุลูกค้า", vbExclamation
        txtCustomer.SetFocus
        Exit Sub
    End If
    If myBillHeader_.isChanged And cmdSave.Enabled Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนทำงานต่อหรือไม่?", vbQuestion + vbYesNoCancel)
        If selResult = vbYes Then
            CmdSave_Click
            If myBillHeader_.isChanged Then Exit Sub
        ElseIf selResult = vbCancel Then
            Exit Sub
        End If
    End If
    Dim comp As String
    If cboCompany.ListIndex = 0 Then
        comp = "AKP"
    Else
        comp = CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).CompanyCode
    End If
    
    If cboCompany.ListIndex >= 0 Then
        curCode = comp
    Else
        curCode = ""
    End If
    
    rowPerPage = 26
    With grMenifest.getGridObj
        For x = 1 To .Rows - 1
            If Trim(.TextMatrix(x, 1)) <> "" Then
                detailCount = detailCount + 1
            End If
        Next
    End With
    With grItem.getGridObj
        For x = 1 To .Rows - 1
            If Trim(.TextMatrix(x, 1)) <> "" Then
                detailCount = detailCount + 1
            End If
        Next
    End With
    For x = 1 To lvTrip.ListItems.Count
        If lvTrip.ListItems(x).Checked Then
            detailCount = detailCount + 1
        End If
    Next
    curRecNo = 0
    pageCount = 1
    CurPage = 1
    X1 = X2
    With rpt1
        .LogonInfo(0) = "DSN = " & DSN_NAME & ";UID = " & strUid & ";PWD = " & strPwd
        If myBillHeader_.BillingNo <> "" Then
        .ReportFileName = App.Path & "\reports\CustInvoice_" & Trim(curCode) & ".rpt"
        Else
         .ReportFileName = App.Path & "\reports\CustInvoice_PRE.rpt"
        End If
        .PrinterName = Printer.DeviceName
        .PrinterDriver = Printer.DriverName
        .PrinterPort = Printer.Port
        .Destination = crptToWindow
        curRow = 1
        ClearInvReport rpt1
        .Formulas(20 + (curRow - 1) * 4) = "Col1_" & curRow & "=" & Chr(34) & Trim(txtHeaderText.Text) & Chr(34)
        .Formulas(21 + (curRow - 1) * 4) = "Col2_" & curRow & "=" & Chr(34) & "" & Chr(34)
        .Formulas(22 + (curRow - 1) * 4) = "Col3_" & curRow & "=" & Chr(34) & "" & Chr(34)
        .Formulas(23 + (curRow - 1) * 4) = "Col4_" & curRow & "=" & Chr(34) & "" & Chr(34)
        curRow = curRow + 1
        PrintNow rpt1, curRow, rowPerPage, curRecNo, detailCount
        Set tmpGrid = grMenifest.getGridObj
        For x = 1 To tmpGrid.Rows - 1
            If Trim(tmpGrid.TextMatrix(x, 1)) <> "" And isExist(myBillHeader_.TripControls, "'" & Trim(tmpGrid.TextMatrix(x, 0)) & "'") Then
                'If Trim(curTripID) <> Trim(tmpGrid.TextMatrix(x, 0)) Then
                If Trim(curTimeTableNo) <> Trim(myBillHeader_.TripControls("'" & tmpGrid.TextMatrix(x, 0) & "'").TimeTableNo) Then
                    Set tmpList = lvTrip.FindItem(curTimeTableNo)
                    If Not tmpList Is Nothing Then
                        If tmpList.Checked Then
                            For y = 1 To myBillHeader_.TripControls.Count
                                If Trim(curTimeTableNo) = Trim(myBillHeader_.TripControls(y).TimeTableNo) Then 'And myBillHeader_.TripControls(y).isMain Then
                                    Set tmpTrip = myBillHeader_.TripControls(y)
                                End If
                            Next
                            If tmpTrip.TripEquipFee + tmpTrip.TripOtherFee + tmpTrip.TripTranFee > 0 Then
                                '.Formulas(20 + (curRow - 1) * 4) = "Col1_" & curRow & "=" & Chr(34) & "ค่าขนส่งรถ เลขทะเบียน " & tmpTrip.CarRegisNo & Chr(34)
                                .Formulas(20 + (curRow - 1) * 4) = "Col1_" & curRow & "=" & Chr(34) & "ค่าขนส่งรถ เลขทะเบียน " & tmpList.SubItems(1) & Chr(34)
                                .Formulas(21 + (curRow - 1) * 4) = "Col2_" & curRow & "=" & Chr(34) & "1 เที่ยว" & Chr(34)
                                .Formulas(22 + (curRow - 1) * 4) = "Col3_" & curRow & "=" & Chr(34) & Format(tmpTrip.TripEquipFee + tmpTrip.TripOtherFee + tmpTrip.TripTranFee, "#,##0.00") & " บาท/เที่ยว" & Chr(34)
                                .Formulas(23 + (curRow - 1) * 4) = "Col4_" & curRow & "=" & Chr(34) & Format(tmpTrip.TripEquipFee + tmpTrip.TripOtherFee + tmpTrip.TripTranFee, "#,##0.00") & Chr(34)
                                curRow = curRow + 1
                                PrintNow rpt1, curRow, rowPerPage, curRecNo, detailCount
                            End If
                            If tmpTrip.TripMoreWorker > 0 Then
                                .Formulas(20 + (curRow - 1) * 4) = "Col1_" & curRow & "=" & Chr(34) & "ค่าแรงคนงาน " & tmpTrip.CarRegisNo & Chr(34)
                                .Formulas(21 + (curRow - 1) * 4) = "Col2_" & curRow & "=" & Chr(34) & tmpTrip.TripMoreWorker & " คน" & Chr(34)
                                .Formulas(22 + (curRow - 1) * 4) = "Col3_" & curRow & "=" & Chr(34) & Format(tmpTrip.TripMoreWorkerFee / tmpTrip.TripMoreWorker, "#,##0.00") & Chr(34)
                                .Formulas(23 + (curRow - 1) * 4) = "Col4_" & curRow & "=" & Chr(34) & Format(tmpTrip.TripMoreWorkerFee, "#,##0.00") & Chr(34)
                                'tmpTotalAmt = tmpTotalAmt + Format(tmpTrip.TripMoreWorkerFee, "0.00") * 1
                                curRow = curRow + 1
                                PrintNow rpt1, curRow, rowPerPage, curRecNo, detailCount
                            End If
                            curRecNo = curRecNo + 1
                            PrintNow rpt1, curRow, rowPerPage, curRecNo, detailCount
                        End If
                    End If
                    Set tmpList = Nothing
                    curTimeTableNo = Trim(myBillHeader_.TripControls("'" & tmpGrid.TextMatrix(x, 0) & "'").TimeTableNo)
                    curTripID = Trim(tmpGrid.TextMatrix(x, 0))
                End If
                .Formulas(20 + (curRow - 1) * 4) = "Col1_" & curRow & "=" & Chr(34) & tmpGrid.TextMatrix(x, 1) & Chr(34)
                .Formulas(21 + (curRow - 1) * 4) = "Col2_" & curRow & "=" & Chr(34) & Trim(tmpGrid.TextMatrix(x, 2) & " " & tmpGrid.TextMatrix(x, 3)) & Chr(34)
                .Formulas(22 + (curRow - 1) * 4) = "Col3_" & curRow & "=" & Chr(34) & Trim(tmpGrid.TextMatrix(x, 4) & " " & tmpGrid.TextMatrix(x, 5)) & Chr(34)
                .Formulas(23 + (curRow - 1) * 4) = "Col4_" & curRow & "=" & Chr(34) & Trim(tmpGrid.TextMatrix(x, 6)) & Chr(34)
                curRow = curRow + 1
                curRecNo = curRecNo + 1
                PrintNow rpt1, curRow, rowPerPage, curRecNo, detailCount
            End If
            If PageContinue = True Then
                PageContinue = False
                Exit Sub
            End If
        Next x
        If isExist(myBillHeader_.TripControls, "'" & curTripID & "'") Then
            'If lvTrip.ListItems("'" & curTripID & "'").Checked Then
            Set tmpList = lvTrip.FindItem(myBillHeader_.TripControls("'" & curTripID & "'").TimeTableNo)
            If tmpList.Checked Then
                For y = 1 To myBillHeader_.TripControls.Count
                    If Trim(curTimeTableNo) = Trim(myBillHeader_.TripControls(y).TimeTableNo) Then 'And myBillHeader_.TripControls(y).isMain Then
                        Set tmpTrip = myBillHeader_.TripControls(y)
                    End If
                Next
                'Set tmpTrip = myBillHeader_.TripControls(tmpList.Key)
                If Not tmpTrip Is Nothing Then
                    If tmpTrip.TripEquipFee + tmpTrip.TripOtherFee + tmpTrip.TripTranFee > 0 Then
                        .Formulas(20 + (curRow - 1) * 4) = "Col1_" & curRow & "=" & Chr(34) & "ค่าขนส่งรถ เลขทะเบียน " & tmpList.SubItems(1) & Chr(34)
                        .Formulas(21 + (curRow - 1) * 4) = "Col2_" & curRow & "=" & Chr(34) & "1 เที่ยว" & Chr(34)
                        .Formulas(22 + (curRow - 1) * 4) = "Col3_" & curRow & "=" & Chr(34) & Format(tmpTrip.TripEquipFee + tmpTrip.TripOtherFee + tmpTrip.TripTranFee, "#,##0.00") & " บาท/เที่ยว" & Chr(34)
                        .Formulas(23 + (curRow - 1) * 4) = "Col4_" & curRow & "=" & Chr(34) & Format(tmpTrip.TripEquipFee + tmpTrip.TripOtherFee + tmpTrip.TripTranFee, "#,##0.00") & Chr(34)
                        'tmpTotalAmt = tmpTotalAmt + Format(tmpTrip.TripEquipFee + tmpTrip.TripOtherFee + tmpTrip.TripTranFee, "0.00") * 1
                        curRow = curRow + 1
                        PrintNow rpt1, curRow, rowPerPage, curRecNo, detailCount
                    End If
                    If tmpTrip.TripMoreWorker > 0 Then
                        .Formulas(20 + (curRow - 1) * 4) = "Col1_" & curRow & "=" & Chr(34) & "ค่าแรงคนงาน " & tmpTrip.CarRegisNo & Chr(34)
                        .Formulas(21 + (curRow - 1) * 4) = "Col2_" & curRow & "=" & Chr(34) & tmpTrip.TripMoreWorker & " คน" & Chr(34)
                        .Formulas(22 + (curRow - 1) * 4) = "Col3_" & curRow & "=" & Chr(34) & Format(tmpTrip.TripMoreWorkerFee / tmpTrip.TripMoreWorker, "#,##0.00") & Chr(34)
                        .Formulas(23 + (curRow - 1) * 4) = "Col4_" & curRow & "=" & Chr(34) & Format(tmpTrip.TripMoreWorkerFee, "#,##0.00") & Chr(34)
                        'tmpTotalAmt = tmpTotalAmt + Format(tmpTrip.TripMoreWorkerFee, "0.00") * 1
                        curRow = curRow + 1
                        PrintNow rpt1, curRow, rowPerPage, curRecNo, detailCount
                    End If
                End If
                curRecNo = curRecNo + 1
                PrintNow rpt1, curRow, rowPerPage, curRecNo, detailCount
            End If
            Set tmpList = Nothing
        End If
        
        Set tmpGrid = grItem.getGridObj
        For x = 1 To tmpGrid.Rows - 1
            If Trim(tmpGrid.TextMatrix(x, 1)) <> "" Then
                .Formulas(20 + (curRow - 1) * 4) = "Col1_" & curRow & "=" & Chr(34) & tmpGrid.TextMatrix(x, 1) & Chr(34)
                .Formulas(21 + (curRow - 1) * 4) = "Col2_" & curRow & "=" & Chr(34) & "" & Chr(34)
                .Formulas(22 + (curRow - 1) * 4) = "Col3_" & curRow & "=" & Chr(34) & "" & Chr(34)
                .Formulas(23 + (curRow - 1) * 4) = "Col4_" & curRow & "=" & Chr(34) & "" & Chr(34)
                If IsNumeric(tmpGrid.TextMatrix(x, 2)) Then
                    If CDbl(tmpGrid.TextMatrix(x, 2)) <> 0 Then
                        .Formulas(21 + (curRow - 1) * 4) = "Col2_" & curRow & "=" & Chr(34) & "1" & Chr(34)
                        .Formulas(22 + (curRow - 1) * 4) = "Col3_" & curRow & "=" & Chr(34) & Format(tmpGrid.TextMatrix(x, 2), "#,##0.00") & Chr(34)
                        .Formulas(23 + (curRow - 1) * 4) = "Col4_" & curRow & "=" & Chr(34) & Format(tmpGrid.TextMatrix(x, 2), "#,##0.00") & Chr(34)
                    End If
                End If
                'tmpTotalAmt = tmpTotalAmt + tmpDet.TotalAmtToPay
                curRow = curRow + 1
                curRecNo = curRecNo + 1
                PrintNow rpt1, curRow, rowPerPage, curRecNo, detailCount
            End If
        Next
            If Trim(myBillHeader_.BillRemark) <> "" Then
                    curRow = curRow + 1
                    PrintNow rpt1, curRow, rowPerPage, curRecNo, detailCount
                    .Formulas(20 + (curRow - 1) * 4) = "Col1_" & curRow & "=" & Chr(34) & myBillHeader_.BillRemark & Chr(34)
                    .Formulas(21 + (curRow - 1) * 4) = "Col2_" & curRow & "=" & Chr(34) & "" & Chr(34)
                    .Formulas(22 + (curRow - 1) * 4) = "Col3_" & curRow & "=" & Chr(34) & "" & Chr(34)
                    .Formulas(23 + (curRow - 1) * 4) = "Col4_" & curRow & "=" & Chr(34) & "" & Chr(34)
                    curRow = curRow + 1
                    PrintNow rpt1, curRow, rowPerPage, curRecNo, detailCount
            End If
            If (curRow >= rowPerPage Or curRecNo >= detailCount) And curRecNo <> -999 Then
                .Formulas(0) = "InvNo=" & Chr(34) & Trim(txtBillNo.Text) & Chr(34)
                .Formulas(1) = "InvDate=" & Chr(34) & Trim(dtBillDate.ValueDMY) & Chr(34)
                
                If cboSiteName.ListIndex > -1 Then
                    Set tmpSite = curCustomer_.BillAddresses(cboSiteName.ListIndex + 1)
                    .Formulas(2) = "CustName=" & Chr(34) & curCustomer_.CustomerCode & Chr(34) & Trim(tmpSite.BillName) & Chr(34)
                    If InStr(1, tmpSite.ProvinceName, "กรุงเทพ") > 0 Then
                        .Formulas(3) = "CustAddr=" & Chr(34) & tmpSite.Address & " แขวง " & tmpSite.TumbolName & " เขต " & tmpSite.AumphurName & " " & tmpSite.ProvinceName & " " & tmpSite.ZipCode & Chr(34)
                    Else
                        .Formulas(3) = "CustAddr=" & Chr(34) & tmpSite.Address & " ต. " & tmpSite.TumbolName & " อ. " & tmpSite.AumphurName & " จ. " & tmpSite.ProvinceName & " " & tmpSite.ZipCode & Chr(34)
                    End If
                    
''                     .Formulas(99) = "CustTax=" & Chr(34) & "เลขประจำตัวผู้เสียภาษีลูกค้า " & tmpSite.TaxNum & Chr(34)
                     
                    Set tmpSite = Nothing
                    
                End If
                'If curRecNo >= detailCount Then
                    'tmpAmt1 = Format(tmpTotalAmt, "0.00")
                    'tmpAmt2 = Format(tmpAmt1 * SystemConfig.VATRate / 100, "0.00")
                    .Formulas(4) = "TotalAmt=" & Chr(34) & lbTotalPrice.Caption & Chr(34)
                    .Formulas(5) = "VATAmt=" & Chr(34) & lbVat.Caption & Chr(34)
                    .Formulas(6) = "GrandTotalAmt=" & Chr(34) & lbNetPrice.Caption & Chr(34)
                    .Formulas(7) = "AMTWord=" & Chr(34) & Trim(AmountToString(Format2Decimal(lbNetPrice.Caption))) & Chr(34)
                    .Formulas(11) = "WTaxAMT=" & Chr(34) & Trim(Me.lbWTax.Caption) & Chr(34)
                        
                    '.Formulas(9) = "DueMonth=" & Chr(34) & Trim(dtDueDate.MonthValue) & Chr(34)
                    '.Formulas(10) = "DueYear=" & Chr(34) & Trim(dtDueDate.YearValue) & Chr(34)
                'End If
                .PrintReport
                If Trim(.LastErrorString) <> "" Then MsgBox .LastErrorString
                ClearInvReport rpt1
                curRow = 1
            End If
    End With
    PageContinue = False
    CurPage = 1
    pageCount = 1
    X1 = 0
    X2 = 0
End Sub

Private Sub CmdSave_Click()
Dim x%
    If Trim(txtCustomer.Text) = "" Then
        MsgBox "กรุณาเลือกลูกหนี้ที่ต้องการวางบิล", vbExclamation
        txtCustomer.SetFocus
        Exit Sub
    End If
    If cboCompany.ListIndex < 0 Then
        MsgBox "กรุณาระบุบริษัทที่ออกใบวางบิล", vbExclamation
        cboCompany.SetFocus
        Exit Sub
    End If
    If myBillHeader_.BillingDetails.Count = 0 Then
        MsgBox "กรุณาระบุรายละเอียดที่จะทำการวางบิล", vbExclamation
        Exit Sub
    End If
    If myBillHeader_.isChanged And cmdSave.Enabled Then
        With myBillHeader_
            If IsNumeric(lbNetPrice.Caption) Then
                .TotalBillAmt = lbNetPrice.Caption
            Else
                .TotalBillAmt = 0
            End If
            If IsNumeric(lbTotalPrice.Caption) Then
                .AmtB4VAT = lbTotalPrice.Caption
            Else
                .AmtB4VAT = 0
            End If
            If IsNumeric(lbVat.Caption) Then
                .VATAmt = lbVat.Caption
            Else
                .VATAmt = 0
            End If
            If IsNumeric(cboWTax.Text) Then
                .WTaxRate = cboWTax.Text
            Else
                .WTaxRate = 0
            End If
            
            .PreDate = Format(TodayDate, "yyyy/MM/dd")
            .PreStaff = CurrentUser.StaffFName & " " & CurrentUser.StaffLName
            
            If .isCancel = True Then MsgBox "เอกสารถูกยกเลิก ไม่สามารถบันทึกได้", vbCritical: Exit Sub
            Screen.MousePointer = 11
            frmMain.Stb1.Panels(1).Text = App.Title
            frmMain.Stb1.Panels(2).Text = ""
            .Save
            If .BillingNo = "" Then
                txtBillNo.Text = myBillHeader_.PreInvNo
            Else
                txtBillNo.Text = myBillHeader_.BillingNo
            End If
        End With
    End If
   ' ** กรณีการบันทึกค่าขนส่ง
'For X = 1 To lvTrip.ListItems.Count
'        If lvTrip.ListItems(X).Checked And IsNumeric(lvTrip.ListItems(X).SubItems(2)) Then
''            tmpAmt = tmpAmt + CDbl(lvTrip.ListItems(X).SubItems(2))
'        End If
'    Next
       
    If myBillHeader_.isChanged = False And cmdSave.Enabled Then
        MsgBox "ข้อมูลไม่ได้เปลี่ยนแปลง ไม่บันทึกข้อมูล", vbInformation, "Info"
    ElseIf myBillHeader_.isSave = True Then
        frmMain.Stb1.Panels(1).Text = "บันทึกสำเร็จ"
        frmMain.Stb1.Panels(2).Text = ""
        MsgBox "จัดเก็บข้อมูลเรียบร้อยแล้ว", vbInformation
        myBillHeader_.isChanged = False
        SetCancelScreen
    Else
        frmMain.Stb1.Panels(1).Text = App.Title
        frmMain.Stb1.Panels(2).Text = err.Number
        MsgBox "เกิดข้อผิดพลาด ไม่สามารถบันทึกข้อมูลได้", vbCritical, "Error"
        Screen.MousePointer = 0
    End If
End Sub

Private Sub cmdReceive_Click()
    Load frmCheqPayment
    With frmCheqPayment
        .SetCallerForm Me
        .Show 1
    End With
End Sub

Private Sub cmdSelect_Click()
    If Trim(txtCustomer.Text) = "" Then
        MsgBox "กรุณาเลือกลูกหนี้ก่อนทำการเลือกใบ Menifest", vbExclamation
        txtCustomer.SetFocus
        Exit Sub
    End If
    If cboCompany.Text = "==เลือกทุกบริษัท==" And cboCompany.ListIndex = 0 Then
        MsgBox "กรุณาเลือกบริษัทที่จะทำการออกใบวางบิลก่อน", vbExclamation
        Exit Sub
    End If
    isDetailChange_ = False
    Load frmMenifestAdd
    With frmMenifestAdd
        .SetCallerForm Me, curCustomer_, myBillHeader_, CurrentUser.getUserCompany(cboCompany.ListIndex + cboIndex).ID
        .Show 1
        If isDetailChange_ Then
            ShowBillingDetail
            ShowTripControl
            SumMoneyAmt
        End If
    End With
End Sub


Private Sub ctlDatepast_Validate(Cancel As Boolean)
myBillHeader_.DatePast = ctlDatepast.ValueYMD
End Sub

Private Sub dtBilldate_Validate(Cancel As Boolean)
    myBillHeader_.IssuedDate = dtBillDate.ValueYMD
End Sub

Private Sub dtDueDate_Validate(Cancel As Boolean)
    myBillHeader_.PayDueDate = dtDueDate.ValueYMD
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 And Me.ActiveControl.Name <> "grItem" Then SendKeys "{Tab}"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
Dim x%, selX As Integer
    CanAccess = False: CanUpdate = False
    SetPermission Me, CanAccess, CanUpdate
    If CanAccess = False Then
        MsgBox "ท่านไม่มีสิทธิใช้ระบบนี้", vbExclamation, "Warning"
        Unload Me
    End If

    selX = -1
    For x = 0 To Printers.Count - 1
        cboPrinter.AddItem Printers(x).DeviceName
        If Printers(x).DeviceName = Printer.DeviceName Then
            selX = x
        End If
    Next
    cboPrinter.ListIndex = selX
    
    With grMenifest.getGridObj
        .Cols = 8
        .Rows = 50
        .TextMatrix(0, 1) = "รายการ"
        .TextMatrix(0, 2) = "จำนวน"
        .TextMatrix(0, 3) = "หน่วย"
        .TextMatrix(0, 4) = "อัตราค่าบริการ"
        .TextMatrix(0, 5) = "หน่วย"
        .TextMatrix(0, 6) = "รวมเป็นเงิน"
        .ColWidth(0) = 0
        .ColWidth(1) = 5625
        .ColWidth(2) = 975
        .ColWidth(3) = 915
        .ColWidth(4) = 1275
        .ColWidth(5) = 1155
        .ColWidth(6) = 1395
        .ColWidth(7) = 0
    End With
    
    PopulateCompany cboCompany, cboIndex
    With cboWTax
        .AddItem "0"
        .AddItem "1"
        .AddItem "3"
        .ListIndex = 0
    End With
    With frmPicture
        cmdSave.Picture = .picSave.Picture
        cmdOpen.Picture = .picSearch.Picture
        cmdClear.Picture = .picNew.Picture
        cmdCancel.Picture = .picDelete.Picture
        cmdClose.Picture = .picExit.Picture
        cmdPrint.Picture = .picPrint
    End With
    With grItem.getGridObj
        .Cols = 3
        .Rows = 50
        .TextMatrix(0, 1) = "รายการ"
        .TextMatrix(0, 2) = "จำนวนเงิน"
        .ColWidth(0) = 0
        .ColWidth(1) = 4425
        .ColWidth(2) = 1320
    End With
    CreateNewBillHeader
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim selResult As Integer
    If myBillHeader_ Is Nothing Then Exit Sub
    If myBillHeader_.isChanged And cmdSave.Enabled Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนปิดหน้าจอหรือไม่?", vbQuestion + vbYesNoCancel)
        If selResult = vbYes Then
            CmdSave_Click
            If myBillHeader_.isChanged Then Cancel = 1
        ElseIf selResult = vbCancel Then
            Cancel = 1
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMain.Stb1.Panels(1).Text = App.Title
    frmMain.Stb1.Panels(2).Text = ""
End Sub

Private Sub grItem_BeforeUpdateValue(NewValue As String, row As Integer, col As Integer, Cancel As Boolean)
    If col = 2 And IsNumeric(NewValue) = False Then
        MsgBox "กรุณาใส่ข้อมูลเป็นตัวเลขเท่านั้น", vbExclamation
        Cancel = True
    End If
End Sub

'Private Sub grItem_Click(row As Integer, col As Integer)
''    MsgBox grItem.getGridObj.ColWidth(col)
'End Sub

Private Sub grItem_KeyDown(KeyCode As Integer, Shift As Integer)
Dim tmpID As String
    tmpID = Trim(grItem.getGridObj.TextMatrix(grItem.getGridObj.row, 0))
    If KeyCode = 46 And Trim(tmpID) <> "" Then
        If MsgBox("ท่านต้องการลบรายการนี้ออกจากใบแจ้งหนี้หรือไม่?", vbQuestion + vbYesNo) = vbYes Then
            grItem.getGridObj.RemoveItem grItem.getGridObj.row
            myBillHeader_.BillingDetails.Remove "'" & tmpID & "'"
            myBillHeader_.isChanged = True
            SumMoneyAmt
        End If
    End If
End Sub

Private Sub grItem_RolColChanged(row As Integer, col As Integer)
    With grItem.getGridObj
        If Trim(.TextMatrix(row - 1, 1)) = "" Then
            grItem.ColEnabled(col) = False
        Else
            If col = 2 And Trim(.TextMatrix(row, 1)) = "" Then
                grItem.ColEnabled(col) = False
            Else
                grItem.ColEnabled(col) = True
            End If
        End If
    End With
End Sub

Private Sub grItem_UpdateValue(row As Integer, col As Integer)
Dim tmpDet As BWGBillingDetail
Dim tmpID As String, tmpVal As String
    tmpID = Trim(grItem.getGridObj.TextMatrix(row, 0))
    tmpVal = Trim(grItem.getGridObj.TextMatrix(row, col))
    If isExist(myBillHeader_.BillingDetails, "'" & tmpID & "'") Then
        Set tmpDet = myBillHeader_.BillingDetails("'" & tmpID & "'")
    Else
        Set tmpDet = New BWGBillingDetail
        With tmpDet
            .ID = GetGUID
            .BillHeaderID = myBillHeader_.ID
            .isCancel = False
            .VATRate = SystemConfig.VATRate
        End With
        myBillHeader_.BillingDetails.Add tmpDet, "'" & tmpDet.ID & "'"
        tmpDet.RowID = myBillHeader_.BillingDetails.Count
        grItem.getGridObj.TextMatrix(row, 0) = tmpDet.ID
    End If
    If col = 1 Then
        tmpDet.BillDesc = tmpVal
    ElseIf col = 2 Then
        If IsNumeric(tmpVal) Then
            tmpDet.AmtB4VAT = tmpVal
            tmpDet.TotalAmtToPay = tmpVal
            tmpDet.VATAmt = Format(tmpDet.AmtB4VAT * (1 + tmpDet.VATRate / 100), "0.00")
        Else
            tmpDet.AmtB4VAT = 0
            tmpDet.TotalAmtToPay = 0
            tmpDet.VATAmt = 0
        End If
        SumMoneyAmt
    End If
    myBillHeader_.isChanged = True
End Sub

Private Sub grMenifest_Click(row As Integer, col As Integer)
    'MsgBox grMenifest.getGridObj.ColWidth(col)
End Sub

Private Sub grMenifest_DblClick(row As Integer, col As Integer)
Dim tmpS As New BWGSearchManager, tmpCol As Collection
    With grMenifest.getGridObj
        If Trim(.TextMatrix(row, 7)) <> "" Then
            Screen.MousePointer = 11
            Set tmpCol = tmpS.MenifestSearch("MenifestID='" & .TextMatrix(row, 7) & "'")
            If tmpCol.Count > 0 And isFormLoaded("frmMenifest") = False Then
                Load frmMenifest
                frmMenifest.SetDataMenifest tmpCol(1)
                Screen.MousePointer = 0
                frmMenifest.Show 1
            End If
            Screen.MousePointer = 0
        End If
    End With
    Set tmpCol = Nothing
    Set tmpS = Nothing
End Sub

Private Sub grMenifest_RolColChanged(row As Integer, col As Integer)
    grMenifest.ColEnabled(col) = False
End Sub

Private Sub lvTrip_ItemCheck(ByVal item As MSComctlLib.ListItem)
Dim tmpTrip As BWGJobDataTimeTable, x%, tmpTimeNo As String
    If item.Checked And IsNumeric(Trim(item.SubItems(2))) = False Then
        MsgBox "ใบคุมนี้ทำการเรียกเก็บค่าขนส่งกับลูกค้าแล้วไม่สามารถเรียกเก็บซ้ำได้", vbExclamation
        item.Checked = False
        Exit Sub
    End If
    tmpTimeNo = Trim(item.Text)
    For x = 1 To myBillHeader_.TripControls.Count
        Set tmpTrip = myBillHeader_.TripControls(x)
        If Trim(tmpTrip.TimeTableNo) = Trim(tmpTimeNo) Then
            If item.Checked Then
                tmpTrip.CustBillID = myBillHeader_.ID
            Else
                tmpTrip.CustBillID = ""
            End If
        End If
    Next
    myBillHeader_.isChanged = True
    SumMoneyAmt
End Sub

Private Sub txtCustomer_Validate(Cancel As Boolean)
Dim selStr As String
    If curCustomer_ Is Nothing Then Set curCustomer_ = New BWGCustomer
    selStr = Trim(curCustomer_.CustomerName)
    If Trim(txtCustomer.Text) = "" And Trim(selStr) <> Trim(txtCustomer.Text) Then
        Set curCustomer_ = Nothing
        Set curCustomer_ = New BWGCustomer
        With myBillHeader_
            .CustomerID = ""
            .CustomerName = ""
        End With
    Else
        If Trim(selStr) <> Trim(txtCustomer.Text) Then
            Set curCustomer_ = Nothing
            Set curCustomer_ = New BWGCustomer
            With myBillHeader_
                .CustomerID = ""
                .CustomerName = ""
            End With
            Load frmCustomerSearch
            With frmCustomerSearch
                .ShowSubContact = True
                .CStatus = "N"
                .getCriteria (txtCustomer.Text)
                .GetCriteriaOptonal ""
'                .GetCriteriaOptonal " AND CustGroupID <> '488F2E5C057F46869D1A56185B11D644'"
                .SetCallerForm Me
                .Show 1
            End With
            If myBillHeader_.CustomerID = "" Then txtCustomer.Text = ""
        End If
    End If
End Sub

Private Sub SumMoneyAmt()
Dim x%, tmpAmt As Double
    myBillHeader_.WTaxRate = Replace(cboWTax.Text, "%", "")
    With grMenifest.getGridObj
        For x = 1 To .Rows - 1
            If IsNumeric(.TextMatrix(x, 6)) Then tmpAmt = tmpAmt + CDbl(.TextMatrix(x, 6))
        Next
    End With
    With grItem.getGridObj
        For x = 1 To .Rows - 1
            If IsNumeric(.TextMatrix(x, 2)) Then tmpAmt = tmpAmt + CDbl(.TextMatrix(x, 2))
        Next
    End With
    For x = 1 To lvTrip.ListItems.Count
        If lvTrip.ListItems(x).Checked And IsNumeric(lvTrip.ListItems(x).SubItems(2)) Then
            tmpAmt = tmpAmt + CDbl(lvTrip.ListItems(x).SubItems(2))
        End If
    Next
    lbTotalPrice.Caption = Format(tmpAmt, "#,##0.00")
    lbWTax.Caption = Format(CDbl(lbTotalPrice.Caption) * (cboWTax.Text) / 100, "#,##0.00")
    If CheckVAT.Value = 1 Then
        lbVat.Caption = "0.00"
    Else
        lbVat.Caption = Format(CDbl(lbTotalPrice.Caption) * (SystemConfig.VATRate) / 100, "#,##0.00")
    End If
    lbNetPrice.Caption = Format(CDbl(lbTotalPrice.Caption) - CDbl(lbWTax.Caption) + CDbl(lbVat.Caption), "#,##0.00")
End Sub

Private Sub ShowBillingDetail()
On Error GoTo ErrD
Dim x%, tmpList As ListItem, y%, z%
Dim tmpTrip As BWGJobDataTimeTable
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim selMenifestID As String, tmpJobDet As BWGJobDataDetail, tmpWaste As BWGWasteDataCR, curRow%, tmpS As New BWGSearchManager
Dim tmpCol As Collection, ColWasteWeight As Collection
Dim curTreatUnitID$, tmpCol2 As Collection, tmpItem As BWGMenifest, curMfID As String, maxType As String, minType As String, tmpTotalAmt As Double
Dim tmpTranUnit$, nonGroupWeight As Double, maxNonTreatRate As Double, maxNonUnit$, curGroupWeight As Double, tmpWeight As Double
Dim tmpWasteType$, curTripID As String, colTrip As Collection, colMinWeightCheck As New Collection, curTimeTableNo$, X1%, tmpTimeTableNo$
Dim curUnitTimeTable$, selMID$, curMenifest$, ExtraCharge$

Dim Dup As Boolean, i As Integer
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim NonOnly As Boolean, isMinWeight As Boolean
Dim TWeight As Double
    selMenifestID = ""
    For x = 1 To myBillHeader_.BillingDetails.Count
        If Trim(myBillHeader_.BillingDetails(x).menifestID) <> "" Then
            selMenifestID = selMenifestID & "'" & myBillHeader_.BillingDetails(x).menifestID & "',"
        End If
    Next
    
    If InStr(1, selMenifestID, ",") > 0 Then
        selMenifestID = Left(selMenifestID, Len(selMenifestID) - 1)
    End If
    grMenifest.ClearAllData
    If Trim(selMenifestID) = "" Then Exit Sub
    curRow = 1
    '''''''''''''เริ่มการหาขั้นต่ำ
    frmMain.Stb1.Panels(1).Text = "Calculating Invoice..."
    frmMain.Stb1.Panels(1).Text = "Inv No " & myBillHeader_.BillingNo & " Detail : " & myBillHeader_.BillingDetails.Count
    If myBillHeader_.BillingDetails.Count > 0 Then
        Set colTrip = tmpS.JobDetailWasteInTripSearch("MenifestID in (" & selMenifestID & ")", "WorkDate,TimeTableNo,DocNo,WasteType,TreatmentUnitID")
    End If
    Set ColWasteWeight = New Collection

    
    For x = 1 To colTrip.Count '******** วบเช็ค Waste วนตามใบคุม, ใบเมนิเฟส, ประเภทเวสต์
        
        Set tmpJobDet = colTrip(x)
        If curTimeTableNo <> tmpJobDet.TimeTableNo Then
            If Trim(curTimeTableNo) <> "" Then
                '*** เพิ่มข้อมูลลงใน colMinWeightCheck ในกรณีที่มีใบคุมหลายใบ
                Call AddMinWeight(colMinWeightCheck, ColWasteWeight, curTimeTableNo)
            End If
            curTimeTableNo = tmpJobDet.TimeTableNo
            curMenifest = tmpJobDet.menifestID
        End If
        
        '***** Calculate min rate with waste type ********************
        If LCase(tmpJobDet.WasteType) = "non" Then
            If isExist(ColWasteWeight, "'NON'") Then
                Set tmpWaste = ColWasteWeight("'NON'")
                If tmpJobDet.ChargeWeightType = "A" Then
                    tmpWaste.MaxQty = tmpWaste.MaxQty + tmpJobDet.DisposerWeight
                Else
                    tmpWaste.MaxQty = tmpWaste.MaxQty + tmpJobDet.EstWasteQty
                End If
                If tmpWaste.MinWeightPerCar = 0 Then tmpWaste.MinWeightPerCar = tmpJobDet.QuoMinWeightPerCar
                If tmpWaste.NonMinRate = 0 Then tmpWaste.NonMinRate = tmpJobDet.NonMinRate

                If tmpWaste.MaxWeightPerCar = 0 Then tmpWaste.MaxWeightPerCar = tmpJobDet.QuoMaxWeightPerCar
                If tmpWaste.MaxRate = 0 Then tmpWaste.MaxRate = tmpJobDet.QuoMaxRate
                If tmpWaste.MaxRateType = "" Then tmpWaste.MaxRateType = tmpJobDet.QuoMaxRateType
                Set tmpWaste = Nothing
            Else
                Set tmpWaste = New BWGWasteDataCR
                tmpWaste.ID = "NON"
                tmpWaste.WasteName = "NON"
                tmpWaste.WasteType = "Non"
                tmpWaste.menifestID = tmpJobDet.menifestID
                tmpWaste.TreatmentUnitID = tmpJobDet.JobTreatmentUnitID
                tmpWaste.MinWeightPerCar = tmpJobDet.QuoMinWeightPerCar
                tmpWaste.NonMinRate = tmpJobDet.NonMinRate
                tmpWaste.MaxWeightPerCar = tmpJobDet.QuoMaxWeightPerCar
                tmpWaste.MaxRate = tmpJobDet.QuoMaxRate
                tmpWaste.MaxRateType = tmpJobDet.QuoMaxRateType
                If tmpJobDet.ChargeWeightType = "A" Then
                    tmpWaste.MaxQty = tmpJobDet.DisposerWeight
                Else
                    tmpWaste.MaxQty = tmpJobDet.EstWasteQty
                End If
                ColWasteWeight.Add tmpWaste, "'NON'"
                Set tmpWaste = Nothing
            End If
        ElseIf LCase(tmpJobDet.WasteType) = "nbl" Then
            If isExist(ColWasteWeight, "'NBL'") Then
                Set tmpWaste = ColWasteWeight("'NBL'")
                If tmpJobDet.ChargeWeightType = "A" Then
                    tmpWaste.MaxQty = tmpWaste.MaxQty + tmpJobDet.DisposerWeight
                Else
                    tmpWaste.MaxQty = tmpWaste.MaxQty + tmpJobDet.EstWasteQty
                End If
                If tmpWaste.MinWeightPerCar = 0 Then tmpWaste.MinWeightPerCar = tmpJobDet.QuoMinWeightPerCar
                If tmpWaste.NBLMinRate = 0 Then tmpWaste.NBLMinRate = tmpJobDet.NBLMinRate

                If tmpWaste.MaxWeightPerCar = 0 Then tmpWaste.MaxWeightPerCar = tmpJobDet.QuoMaxWeightPerCar
                If tmpWaste.MaxRate = 0 Then tmpWaste.MaxRate = tmpJobDet.QuoMaxRate
                If tmpWaste.MaxRateType = "" Then tmpWaste.MaxRateType = tmpJobDet.QuoMaxRateType
                Set tmpWaste = Nothing
            Else
                Set tmpWaste = New BWGWasteDataCR
                tmpWaste.ID = "NBL"
                tmpWaste.WasteName = "NBL"
                tmpWaste.WasteType = "NBL"
                tmpWaste.menifestID = tmpJobDet.menifestID
                tmpWaste.TreatmentUnitID = tmpJobDet.JobTreatmentUnitID
                tmpWaste.MinWeightPerCar = tmpJobDet.QuoMinWeightPerCar
                tmpWaste.NBLMinRate = tmpJobDet.NBLMinRate
                tmpWaste.MaxWeightPerCar = tmpJobDet.QuoMaxWeightPerCar
                tmpWaste.MaxRate = tmpJobDet.QuoMaxRate
                tmpWaste.MaxRateType = tmpJobDet.QuoMaxRateType
                If tmpJobDet.ChargeWeightType = "A" Then
                    tmpWaste.MaxQty = tmpJobDet.DisposerWeight
                Else
                    tmpWaste.MaxQty = tmpJobDet.EstWasteQty
                End If
                ColWasteWeight.Add tmpWaste, "'NBL'"
                Set tmpWaste = Nothing
            End If
        Else
            If isExist(ColWasteWeight, "'" & tmpJobDet.wastedataID & "'") Then
                Set tmpWaste = ColWasteWeight("'" & tmpJobDet.wastedataID & "'")
                If tmpJobDet.ChargeWeightType = "A" Then
                    tmpWaste.MaxQty = tmpWaste.MaxQty + tmpJobDet.DisposerWeight
                Else
                    tmpWaste.MaxQty = tmpWaste.MaxQty + tmpJobDet.EstWasteQty
                End If
                If tmpWaste.MinWeightPerCar = 0 Then tmpWaste.MinWeightPerCar = tmpJobDet.QuoMinWeightPerCar
                If tmpWaste.MaxWeightPerCar = 0 Then tmpWaste.MaxWeightPerCar = tmpJobDet.QuoMaxWeightPerCar
                If tmpWaste.MaxRate = 0 Then tmpWaste.MaxRate = tmpJobDet.QuoMaxRate
                If tmpWaste.MaxRateType = "" Then tmpWaste.MaxRateType = tmpJobDet.QuoMaxRateType
                Set tmpWaste = Nothing
            Else
                Set tmpWaste = New BWGWasteDataCR
                tmpWaste.ID = tmpJobDet.wastedataID
                tmpWaste.WasteName = tmpJobDet.WasteName
                tmpWaste.WasteType = tmpJobDet.WasteType
                tmpWaste.menifestID = tmpJobDet.menifestID
                tmpWaste.TreatmentUnitID = tmpJobDet.JobTreatmentUnitID
                tmpWaste.MaxWeightPerCar = tmpJobDet.QuoMaxWeightPerCar
                tmpWaste.MaxRate = tmpJobDet.QuoMaxRate
                tmpWaste.MaxRateType = tmpJobDet.QuoMaxRateType
                tmpWaste.MinWeightPerCar = tmpJobDet.QuoMinWeightPerCar
                If LCase(tmpJobDet.WasteType) = "haz" Then
                    If tmpJobDet.HazMinRate > 0 Then
                        tmpWaste.HazMinRate = tmpJobDet.HazMinRate
                    Else
                        tmpWaste.HazMinRate = tmpJobDet.JobTreatmentRate
                    End If
                ElseIf LCase(tmpJobDet.WasteType) = "hbl" Then
                    If tmpJobDet.HBLMinRate > 0 Then
                        tmpWaste.HBLMinRate = tmpJobDet.HBLMinRate
                    Else
                        tmpWaste.HBLMinRate = tmpJobDet.JobTreatmentRate
                    End If
                ElseIf LCase(tmpJobDet.WasteType) = "nbl" Then
                    If tmpJobDet.NBLMinRate > 0 Then
                        tmpWaste.NBLMinRate = tmpJobDet.NBLMinRate
                    Else
                        tmpWaste.NBLMinRate = tmpJobDet.JobTreatmentRate
                    End If
                Else
                    If tmpJobDet.HazMinRate > 0 Then
                        tmpWaste.HazMinRate = tmpJobDet.HazMinRate
                    Else
                        tmpWaste.HazMinRate = tmpJobDet.JobTreatmentRate
                    End If
                End If
                If tmpJobDet.ChargeWeightType = "A" Then
                    tmpWaste.MaxQty = tmpJobDet.DisposerWeight
                Else
                    tmpWaste.MaxQty = tmpJobDet.EstWasteQty
                End If
                ColWasteWeight.Add tmpWaste, "'" & tmpWaste.ID & "'"
                Set tmpWaste = Nothing
            End If
        End If
        tmpTimeTableNo = tmpJobDet.TimeTableNo
        selMID = tmpJobDet.menifestID
        Set tmpJobDet = Nothing
    Next '******** วนเช็คน้ำหนัก Waste วนตามใบคุม, ใบเมนิเฟส, ประเภทเวสต์
    
    '*** เพิ่มข้อมูลลงใน colMinWeightCheck ในกรณีที่มีใบคุมใบเดียว หรือ ใบคุมใบสุดท้าย
    Call AddMinWeight(colMinWeightCheck, ColWasteWeight, tmpTimeTableNo)

'************* End Calculate min rate with waste type ***********************

    frmMain.Stb1.Panels(2).Text = ""
    With grMenifest.getGridObj
        .Rows = 3
        curTreatUnitID = ""
        'Set tmpCol2 = myTripControl_.GetMenifests("T")  'Or "A"
        Set tmpCol2 = tmpS.MenifestSearch("MenifestID in (" & selMenifestID & ") And (isCanceled<>'Y' Or isCanceled is null)")
'        Set tmpCol2 = myTripControl_.GetMenifests("A")
        
        frmMain.Stb1.Panels(2).Text = "Get Manifest Count " & tmpCol2.Count
        
        For x = 1 To tmpCol2.Count
            Set tmpItem = tmpCol2(x)
            'If Trim(curTripID) <> tmpItem.JobDataCarID & "_" & tmpItem.tripNo Then
            
            If Trim(curTripID) <> tmpItem.TimeTableNo & "_" & tmpItem.TripNo Then
                If Trim(curTripID) = "" Then
                    nonGroupWeight = 0
                    curGroupWeight = 0
                ElseIf tmpItem.TimeTableNo <> Replace(Trim(Split(curTripID, "_")(0)), "'", "") Then
                    nonGroupWeight = 0
                    curGroupWeight = 0
                End If
                curTreatUnitID = ""
                'curTripID = tmpItem.JobDataCarID & "_" & tmpItem.tripNo
                curTripID = tmpItem.TimeTableNo & "_" & tmpItem.TripNo
            End If
            If Trim(tmpItem.ID) <> Trim(curMfID) Then
                curMfID = Trim(tmpItem.ID)
                If maxType < LCase(tmpItem.NonOrHz) Then maxType = LCase(tmpItem.NonOrHz)
                If minType > LCase(tmpItem.NonOrHz) Then minType = LCase(tmpItem.NonOrHz)
                .TextMatrix(curRow, 7) = tmpItem.ID
                .TextMatrix(curRow, 0) = tmpItem.JobDataCarID & "_" & tmpItem.TripNo
                If tmpItem.RefME <> "" Then
                    .TextMatrix(curRow, 1) = "วันที่ " & FormatYMD_to_DMY(tmpItem.WorkDate, "/", "/") & " เลขที่ " & tmpItem.DocNo & " อ้างอิง " & tmpItem.RefME
                Else
                    .TextMatrix(curRow, 1) = "วันที่ " & FormatYMD_to_DMY(tmpItem.WorkDate, "/", "/") & " เลขที่ " & tmpItem.DocNo
                End If
                .TextMatrix(curRow, 2) = ""
                .TextMatrix(curRow, 3) = ""
                .TextMatrix(curRow, 4) = ""
                .TextMatrix(curRow, 5) = ""
                .TextMatrix(curRow, 6) = ""
                curRow = curRow + 1
                If .Rows - 1 = curRow Then .Rows = .Rows + 5
                If Trim(tmpItem.InOtherDesc) <> "" Then
                    .TextMatrix(curRow, 7) = tmpItem.ID
                    .TextMatrix(curRow, 0) = tmpItem.JobDataCarID & "_" & tmpItem.TripNo
                    .TextMatrix(curRow, 1) = tmpItem.InOtherDesc
                    .TextMatrix(curRow, 2) = "1"
                    .TextMatrix(curRow, 3) = ""
                    .TextMatrix(curRow, 4) = tmpItem.InOtherFee
                    .TextMatrix(curRow, 5) = ""
                    .TextMatrix(curRow, 6) = tmpItem.InOtherFee
                    tmpTotalAmt = tmpTotalAmt + tmpItem.InOtherFee
                    curRow = curRow + 1
                    If .Rows - 1 = curRow Then .Rows = .Rows + 5
                End If
            End If
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Set tmpCol = Nothing
            Set tmpCol = New Collection
            For z = 1 To colTrip.Count
                If Trim(colTrip(z).menifestID) = Trim(tmpItem.ID) Then
                    tmpCol.Add colTrip(z), "'" & colTrip(z).ID & "'"
                End If
            Next
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            'Set tmpCol = tmpS.JobDetailInTripSearch("MenifestID='" & tmpItem.ID & "'", "WasteType,TreatmentUnitID")
            If tmpCol.Count <= 0 Then
                frmMain.Stb1.Panels(1).Text = "Menifest detail can't not load, system will be load data again."
                frmMain.Stb1.Panels(2).Text = "Menifest detail: " & Trim("" & tmpCol.Count)
                ShowBillingDetail
            End If
            For y = 1 To tmpCol.Count
                Set tmpJobDet = tmpCol(y)
                If curUnitTimeTable <> tmpJobDet.TimeTableNo Then If curUnitTimeTable <> "" Then maxNonTreatRate = 0: curGroupWeight = 0: nonGroupWeight = 0
                tmpTranUnit = tmpJobDet.QuoTransUnitID
'                If curTreatUnitID = "" Then curTreatUnitID = Trim(tmpJobDet.JobTreatmentUnitID)
                
                If Trim(tmpJobDet.FailType) = "" Then
                    .TextMatrix(curRow, 7) = tmpItem.ID
                    .TextMatrix(curRow, 0) = tmpItem.JobDataCarID & "_" & tmpItem.TripNo
                    .TextMatrix(curRow, 1) = tmpJobDet.WasteName
                    If UCase(tmpJobDet.WasteType) = "NON" Or UCase(tmpJobDet.WasteType) = "NBL" Then
                        If tmpJobDet.ChargeWeightType = "A" Then
                            nonGroupWeight = nonGroupWeight + tmpJobDet.DisposerWeight
                        Else
                            nonGroupWeight = nonGroupWeight + tmpJobDet.EstWasteQty
                        End If
                        If maxNonTreatRate <= tmpJobDet.JobTreatmentRate Then
                            maxNonTreatRate = tmpJobDet.JobTreatmentRate
                            maxNonUnit = tmpJobDet.JobTreatmentUnitID
                        End If
                    End If
                    If tmpJobDet.JobTreatmentUnitID = "บาท/เที่ยว" Or tmpJobDet.JobTreatmentUnitID = "บาท/Job" Then
                        If Trim(curTreatUnitID) = "" And curUnitTimeTable <> tmpJobDet.TimeTableNo Then
                            .TextMatrix(curRow - 1, 2) = "1"
                            If tmpJobDet.JobTreatmentUnitID = "บาท/เที่ยว" Then
                            .TextMatrix(curRow - 1, 3) = "เที่ยว"
                            Else
                             .TextMatrix(curRow - 1, 3) = "Job"
                            End If
                            .TextMatrix(curRow - 1, 4) = Format(tmpJobDet.JobTreatmentRate, "#,##0.00")
                            .TextMatrix(curRow - 1, 5) = tmpJobDet.JobTreatmentUnitID
                            .TextMatrix(curRow - 1, 6) = Format(tmpJobDet.JobTreatmentRate, "#,##0.00")
                            tmpTotalAmt = tmpTotalAmt + tmpJobDet.JobTreatmentRate
                            curUnitTimeTable = tmpJobDet.TimeTableNo
                        End If
                        If UCase(tmpItem.NonOrHz) = "NON" Or UCase(tmpItem.NonOrHz) = "NBL" Then
                            .TextMatrix(curRow, 2) = ""
                            .TextMatrix(curRow, 3) = ""
                            .TextMatrix(curRow, 4) = ""
                            .TextMatrix(curRow, 5) = ""
                            .TextMatrix(curRow, 6) = ""
                        Else
                            If tmpJobDet.ChargeWeightType = "A" Then
                                .TextMatrix(curRow, 2) = Format(tmpJobDet.DisposerWeight, "#,##0.000")
                            Else
                                .TextMatrix(curRow, 2) = Format(tmpJobDet.EstWasteQty, "#,##0.000")
                            End If
                            .TextMatrix(curRow, 3) = "ตัน"
                            .TextMatrix(curRow, 4) = ""
                            .TextMatrix(curRow, 5) = ""
                            .TextMatrix(curRow, 6) = ""
                        End If
                        If tmpJobDet.ChargeWeightType = "A" Then
                            curGroupWeight = curGroupWeight + tmpJobDet.DisposerWeight
                        Else
                            curGroupWeight = curGroupWeight + tmpJobDet.EstWasteQty
                        End If
                        curTreatUnitID = Trim(tmpJobDet.JobTreatmentUnitID)
                    Else
                        If Trim(curTreatUnitID) <> "" Then
                            .TextMatrix(curRow, 7) = tmpItem.ID
                            .TextMatrix(curRow, 0) = tmpItem.JobDataCarID & "_" & tmpItem.TripNo
                            .TextMatrix(curRow, 1) = "น้ำหนักรวม"
                            .TextMatrix(curRow, 2) = Format(tmpJobDet.MenifestActQty, "#,##0.000")
                            .TextMatrix(curRow, 3) = "ตัน"
                            .TextMatrix(curRow, 4) = ""
                            .TextMatrix(curRow, 5) = ""
                            .TextMatrix(curRow, 6) = ""
                            curRow = curRow + 1
                            If .Rows - 1 = curRow Then .Rows = .Rows + 5
                        End If
                        curTreatUnitID = ""
                        curUnitTimeTable = tmpJobDet.TimeTableNo
                        curGroupWeight = 0
                        If UCase(tmpJobDet.WasteType) <> "NON" And UCase(tmpJobDet.WasteType) <> "NBL" Then
                            If tmpJobDet.ChargeWeightType = "A" Then
                                If InStr(1, tmpWaste.TreatmentUnitID, "กิโลกรัม") Then
                                .TextMatrix(curRow, 2) = Format(tmpJobDet.DisposerWeight, "#,##0")
                                Else
                                .TextMatrix(curRow, 2) = Format(tmpJobDet.DisposerWeight, "#,##0.000")
                                End If
                                .TextMatrix(curRow, 3) = Replace(tmpJobDet.JobTreatmentUnitID, "บาท/", "")
                                tmpWeight = tmpJobDet.DisposerWeight
                            Else
                                 If InStr(1, tmpWaste.TreatmentUnitID, "กิโลกรัม") Then
                                .TextMatrix(curRow, 2) = Format(tmpJobDet.DisposerWeight, "#,##0")
                                Else
                                .TextMatrix(curRow, 2) = Format(tmpJobDet.DisposerWeight, "#,##0.000")
                                End If
                                .TextMatrix(curRow, 3) = Replace(tmpJobDet.JobTreatmentUnitID, "บาท/", "")
                                tmpWeight = tmpJobDet.EstWasteQty
                            End If
                            If colMinWeightCheck("'" & tmpJobDet.TimeTableNo & "'").isUnderWeight And LCase(Trim(colMinWeightCheck("'" & tmpJobDet.TimeTableNo & "'").MaxWasteID)) = LCase(Trim(tmpJobDet.wastedataID)) Then
                                If LCase(tmpJobDet.WasteType) = "haz" Then
                                    If tmpJobDet.HazMinRate > 0 Then
'                                        .TextMatrix(curRow, 4) = Format(tmpJobDet.HazMinRate, "#,##0.00")  'คูณราคาขั้นต่ำ
                                         .TextMatrix(curRow, 4) = Format(tmpJobDet.JobTreatmentRate, "#,##0.00")
                                    Else
                                        .TextMatrix(curRow, 4) = Format(tmpJobDet.JobTreatmentRate, "#,##0.00")
                                    End If
                                ElseIf LCase(tmpJobDet.WasteType) = "hbl" Then
                                    If tmpJobDet.HBLMinRate > 0 Then
'                                        .TextMatrix(curRow, 4) = Format(tmpJobDet.HBLMinRate, "#,##0.00") ' คูณราคาขั้นต่ำ
                                         .TextMatrix(curRow, 4) = Format(tmpJobDet.JobTreatmentRate, "#,##0.00")
                                    Else
                                        .TextMatrix(curRow, 4) = Format(tmpJobDet.JobTreatmentRate, "#,##0.00")
                                    End If
                                Else
                                    If tmpJobDet.HazMinRate > 0 Then
'                                        .TextMatrix(curRow, 4) = Format(tmpJobDet.HazMinRate, "#,##0.00")
                                         .TextMatrix(curRow, 4) = Format(tmpJobDet.JobTreatmentRate, "#,##0.00")
                                    Else
                                        .TextMatrix(curRow, 4) = Format(tmpJobDet.JobTreatmentRate, "#,##0.00")
                                    End If
                                End If
                                .row = curRow
                                .col = 4
                                .CellForeColor = vbRed
                            Else
                                .TextMatrix(curRow, 4) = Format(tmpJobDet.JobTreatmentRate, "#,##0.00")
                            End If
                            .TextMatrix(curRow, 5) = tmpJobDet.JobTreatmentUnitID
                            .TextMatrix(curRow, 6) = Format(tmpWeight * CDbl(.TextMatrix(curRow, 4)), "#,##0.00")
                            tmpTotalAmt = tmpTotalAmt + CDbl(.TextMatrix(curRow, 6))
                        End If
                    End If
                    curRow = curRow + 1
                    If .Rows - 1 = curRow Then .Rows = .Rows + 5
                End If
                tmpWasteType = tmpJobDet.WasteType
                tmpTimeTableNo = tmpJobDet.TimeTableNo
                Set tmpJobDet = Nothing
            Next y
            
            Select Case UCase(tmpWasteType)
                Case "NON", "NBL"
                    If InStr(1, maxNonUnit, "ตัน") > 0 Then
                        If colMinWeightCheck("'" & tmpTimeTableNo & "'").isUnderWeight And LCase(Trim(colMinWeightCheck("'" & tmpTimeTableNo & "'").MaxWasteType)) = LCase(Trim(tmpWasteType)) Then
                            .TextMatrix(curRow, 7) = tmpItem.ID
                            .TextMatrix(curRow, 0) = tmpItem.JobDataCarID & "_" & tmpItem.TripNo
                            .TextMatrix(curRow, 1) = "น้ำหนักรวม"
                            .TextMatrix(curRow, 2) = Format(nonGroupWeight, "#,##0.000")
                            .TextMatrix(curRow, 3) = "ตัน"
                            '=================================
                            .TextMatrix(curRow, 4) = Format(maxNonTreatRate, "#,##0.00")
                            .TextMatrix(curRow, 5) = maxNonUnit
                            .TextMatrix(curRow, 6) = Format(nonGroupWeight * maxNonTreatRate, "#,##0.00")
                            tmpTotalAmt = tmpTotalAmt + Format(nonGroupWeight * maxNonTreatRate, "0.00") * 1
                            '=================================
                            curRow = curRow + 1
                            If .Rows - 1 = curRow Then .Rows = .Rows + 5
                            If Trim(colMinWeightCheck("'" & tmpTimeTableNo & "'").menifestID) = Trim(tmpItem.ID) Then
                                Set tmpWaste = colMinWeightCheck("'" & tmpTimeTableNo & "'").ColWasteWeight("'" & colMinWeightCheck("'" & tmpTimeTableNo & "'").MaxWasteID & "'")
                                .TextMatrix(curRow, 7) = tmpItem.ID
                                .TextMatrix(curRow, 0) = tmpItem.JobDataCarID & "_" & tmpItem.TripNo
                                ExtraCharge = "น้ำหนักไม่ถึง " & tmpWaste.MinWeightPerCar & " ตัน คิดน้ำหนักเพิ่ม"
                                .TextMatrix(curRow, 1) = ExtraCharge
        '                        .TextMatrix(curRow, 2) = Format(tmpWaste.MinWeightPerCar, "#,##0.000")
                                '===============================
                                .TextMatrix(curRow, 2) = Format(tmpWaste.MinWeightPerCar - colMinWeightCheck("'" & tmpTimeTableNo & "'").TotalWeight, "#,##0.000")
                                TWeight = CDbl(.TextMatrix(curRow, 2))
                                '===============================
                                .TextMatrix(curRow, 3) = "ตัน"
                                If LCase(tmpWaste.WasteType) = "non" Then
                                    If tmpWaste.NonMinRate > 0 Then
                                        .TextMatrix(curRow, 4) = Format(tmpWaste.NonMinRate, "#,##0.00")
                                    Else
                                        .TextMatrix(curRow, 4) = Format(tmpWaste.TreatmentFee, "#,##0.00")
                                    End If
                                ElseIf LCase(tmpWaste.WasteType) = "nbl" Then
                                    If tmpWaste.NBLMinRate > 0 Then
                                        .TextMatrix(curRow, 4) = Format(tmpWaste.NBLMinRate, "#,##0.00")
                                    Else
                                        .TextMatrix(curRow, 4) = Format(tmpWaste.TreatmentFee, "#,##0.00")
                                    End If
                                End If
                                .TextMatrix(curRow, 5) = "บาท/ตัน"
        '                        .TextMatrix(curRow, 6) = Format(CDbl(.TextMatrix(curRow, 4)) * CDbl(.TextMatrix(curRow, 2)), "#,##0.00")
                                '===============
                                .TextMatrix(curRow, 6) = Format(CDbl(.TextMatrix(curRow, 4)) * TWeight, "#,##0.00")
                                '===============
                                tmpTotalAmt = tmpTotalAmt + CDbl(.TextMatrix(curRow, 6))
                                curRow = curRow + 1
                                If .Rows - 1 = curRow Then .Rows = .Rows + 5
                            End If
                        ElseIf colMinWeightCheck("'" & tmpTimeTableNo & "'").isOverWeight And LCase(Trim(colMinWeightCheck("'" & tmpTimeTableNo & "'").MaxWasteType)) = LCase(Trim(tmpWasteType)) Then
                            .TextMatrix(curRow, 7) = tmpItem.ID
                            .TextMatrix(curRow, 0) = tmpItem.JobDataCarID & "_" & tmpItem.TripNo
                            .TextMatrix(curRow, 1) = "น้ำหนักรวม"
                            .TextMatrix(curRow, 2) = Format(nonGroupWeight, "#,##0.000")
                            .TextMatrix(curRow, 3) = "ตัน"
                            '=================================
                            .TextMatrix(curRow, 4) = Format(maxNonTreatRate, "#,##0.00")
                            .TextMatrix(curRow, 5) = maxNonUnit
                            .TextMatrix(curRow, 6) = Format(nonGroupWeight * maxNonTreatRate, "#,##0.00")
                            tmpTotalAmt = tmpTotalAmt + Format(nonGroupWeight * maxNonTreatRate, "0.00") * 1
                            '=================================
                            curRow = curRow + 1
                            If .Rows - 1 = curRow Then .Rows = .Rows + 5
                            If Trim(colMinWeightCheck("'" & tmpTimeTableNo & "'").menifestID) = Trim(tmpItem.ID) Then
                                Set tmpWaste = colMinWeightCheck("'" & tmpTimeTableNo & "'").ColWasteWeight("'" & colMinWeightCheck("'" & tmpTimeTableNo & "'").MaxWasteID & "'")
                                .TextMatrix(curRow, 7) = tmpItem.ID
                                .TextMatrix(curRow, 0) = tmpItem.JobDataCarID & "_" & tmpItem.TripNo
                                .TextMatrix(curRow, 1) = "น้ำหนักเกินที่กำหนด " & tmpWaste.MaxWeightPerCar & " ตัน คิดน้ำหนักส่วนเกิน"
                                .TextMatrix(curRow, 2) = Format(colMinWeightCheck("'" & tmpTimeTableNo & "'").TotalWeight - tmpWaste.MaxWeightPerCar, "#,##0.000")
                                .TextMatrix(curRow, 3) = "ตัน"
                                .TextMatrix(curRow, 4) = Format(tmpWaste.MaxRate, "#,##0.00")
                                .TextMatrix(curRow, 5) = tmpWaste.MaxRateType
                                If InStr(1, tmpWaste.MaxRateType, "เที่ยว") > 0 Then
                                    .TextMatrix(curRow, 6) = Format(tmpWaste.MaxRate, "#,##0.00")
                                Else
                                    .TextMatrix(curRow, 6) = Format(CDbl(.TextMatrix(curRow, 4)) * CDbl(.TextMatrix(curRow, 2)), "#,##0.00")
                                End If
                                tmpTotalAmt = tmpTotalAmt + CDbl(.TextMatrix(curRow, 6))
                                curRow = curRow + 1
                                If .Rows - 1 = curRow Then .Rows = .Rows + 5
                            End If
                        Else
                            .TextMatrix(curRow, 7) = tmpItem.ID
                            .TextMatrix(curRow, 0) = tmpItem.JobDataCarID & "_" & tmpItem.TripNo
                            .TextMatrix(curRow, 1) = "น้ำหนักรวม"
                            .TextMatrix(curRow, 2) = Format(nonGroupWeight, "#,##0.000")
                            .TextMatrix(curRow, 3) = "ตัน"
                            .TextMatrix(curRow, 4) = Format(maxNonTreatRate, "#,##0.00")
                            .TextMatrix(curRow, 5) = maxNonUnit
                            .TextMatrix(curRow, 6) = Format(nonGroupWeight * maxNonTreatRate, "#,##0.00")
                            tmpTotalAmt = tmpTotalAmt + Format(nonGroupWeight * maxNonTreatRate, "0.00") * 1
                            curRow = curRow + 1
                            If .Rows - 1 = curRow Then .Rows = .Rows + 5
                        End If
                        nonGroupWeight = 0
                        maxNonTreatRate = 0
                        ElseIf InStr(1, maxNonUnit, "กิโลกรัม") > 0 Then
                        If colMinWeightCheck("'" & tmpTimeTableNo & "'").isUnderWeight And LCase(Trim(colMinWeightCheck("'" & tmpTimeTableNo & "'").MaxWasteType)) = LCase(Trim(tmpWasteType)) Then
                            .TextMatrix(curRow, 7) = tmpItem.ID
                            .TextMatrix(curRow, 0) = tmpItem.JobDataCarID & "_" & tmpItem.TripNo
                            .TextMatrix(curRow, 1) = "น้ำหนักรวม"
                            .TextMatrix(curRow, 2) = Format(nonGroupWeight, "#,##0")
                            .TextMatrix(curRow, 3) = "กิโลกรัม"
                            '=================================
                            .TextMatrix(curRow, 4) = Format(maxNonTreatRate, "#,##0.00")
                            .TextMatrix(curRow, 5) = maxNonUnit
                            .TextMatrix(curRow, 6) = Format(nonGroupWeight * maxNonTreatRate, "#,##0.00")
                            tmpTotalAmt = tmpTotalAmt + Format(nonGroupWeight * maxNonTreatRate, "0.00") * 1
                            '=================================
                            curRow = curRow + 1
                            If .Rows - 1 = curRow Then .Rows = .Rows + 5
                            If Trim(colMinWeightCheck("'" & tmpTimeTableNo & "'").menifestID) = Trim(tmpItem.ID) Then
                                Set tmpWaste = colMinWeightCheck("'" & tmpTimeTableNo & "'").ColWasteWeight("'" & colMinWeightCheck("'" & tmpTimeTableNo & "'").MaxWasteID & "'")
                                .TextMatrix(curRow, 7) = tmpItem.ID
                                .TextMatrix(curRow, 0) = tmpItem.JobDataCarID & "_" & tmpItem.TripNo
                                ExtraCharge = "น้ำหนักไม่ถึง " & tmpWaste.MinWeightPerCar & " กิโลกรัม คิดน้ำหนักเพิ่ม"
                                .TextMatrix(curRow, 1) = ExtraCharge
        '                        .TextMatrix(curRow, 2) = Format(tmpWaste.MinWeightPerCar, "#,##0.000")
                                '===============================
                                .TextMatrix(curRow, 2) = Format((tmpWaste.MinWeightPerCar - colMinWeightCheck("'" & tmpTimeTableNo & "'").TotalWeight) * 1000, "#,##0")
                                TWeight = CDbl(.TextMatrix(curRow, 2))
                                '===============================
                                .TextMatrix(curRow, 3) = "กิโลกรัม"
                                If LCase(tmpWaste.WasteType) = "non" Then
                                    If tmpWaste.NonMinRate > 0 Then
                                        .TextMatrix(curRow, 4) = Format(tmpWaste.NonMinRate, "#,##0.00")
                                    Else
                                        .TextMatrix(curRow, 4) = Format(tmpWaste.TreatmentFee, "#,##0.00")
                                    End If
                                ElseIf LCase(tmpWaste.WasteType) = "nbl" Then
                                    If tmpWaste.NBLMinRate > 0 Then
                                        .TextMatrix(curRow, 4) = Format(tmpWaste.NBLMinRate, "#,##0.00")
                                    Else
                                        .TextMatrix(curRow, 4) = Format(tmpWaste.TreatmentFee, "#,##0.00")
                                    End If
                                End If
                                .TextMatrix(curRow, 5) = "บาท/ตัน"
        '                        .TextMatrix(curRow, 6) = Format(CDbl(.TextMatrix(curRow, 4)) * CDbl(.TextMatrix(curRow, 2)), "#,##0.00")
                                '===============
                                .TextMatrix(curRow, 6) = Format(CDbl(.TextMatrix(curRow, 4)) * TWeight, "#,##0.00")
                                '===============
                                tmpTotalAmt = tmpTotalAmt + CDbl(.TextMatrix(curRow, 6))
                                curRow = curRow + 1
                                If .Rows - 1 = curRow Then .Rows = .Rows + 5
                            End If
                        ElseIf colMinWeightCheck("'" & tmpTimeTableNo & "'").isOverWeight And LCase(Trim(colMinWeightCheck("'" & tmpTimeTableNo & "'").MaxWasteType)) = LCase(Trim(tmpWasteType)) Then
                            .TextMatrix(curRow, 7) = tmpItem.ID
                            .TextMatrix(curRow, 0) = tmpItem.JobDataCarID & "_" & tmpItem.TripNo
                            .TextMatrix(curRow, 1) = "น้ำหนักรวม"
                            .TextMatrix(curRow, 2) = Format(nonGroupWeight, "#,##0")
                            .TextMatrix(curRow, 3) = "กิโลกรัม"
                            '=================================
                            .TextMatrix(curRow, 4) = Format(maxNonTreatRate, "#,##0.00")
                            .TextMatrix(curRow, 5) = maxNonUnit
                            .TextMatrix(curRow, 6) = Format(nonGroupWeight * maxNonTreatRate, "#,##0.00")
                            tmpTotalAmt = tmpTotalAmt + Format(nonGroupWeight * maxNonTreatRate, "0.00") * 1
                            '=================================
                            curRow = curRow + 1
                            If .Rows - 1 = curRow Then .Rows = .Rows + 5
                            If Trim(colMinWeightCheck("'" & tmpTimeTableNo & "'").menifestID) = Trim(tmpItem.ID) Then
                                Set tmpWaste = colMinWeightCheck("'" & tmpTimeTableNo & "'").ColWasteWeight("'" & colMinWeightCheck("'" & tmpTimeTableNo & "'").MaxWasteID & "'")
                                .TextMatrix(curRow, 7) = tmpItem.ID
                                .TextMatrix(curRow, 0) = tmpItem.JobDataCarID & "_" & tmpItem.TripNo
                                .TextMatrix(curRow, 1) = "น้ำหนักเกินที่กำหนด " & tmpWaste.MaxWeightPerCar & " กิโลกรัม คิดน้ำหนักส่วนเกิน"
                                .TextMatrix(curRow, 2) = Format((colMinWeightCheck("'" & tmpTimeTableNo & "'").TotalWeight - tmpWaste.MaxWeightPerCar) * 1000, "#,##0")
                                .TextMatrix(curRow, 3) = "กิโลกรัม"
                                .TextMatrix(curRow, 4) = Format(tmpWaste.MaxRate, "#,##0.00")
                                .TextMatrix(curRow, 5) = tmpWaste.MaxRateType
                                If InStr(1, tmpWaste.MaxRateType, "เที่ยว") > 0 Then
                                    .TextMatrix(curRow, 6) = Format(tmpWaste.MaxRate, "#,##0.00")
                                Else
                                    .TextMatrix(curRow, 6) = Format(CDbl(.TextMatrix(curRow, 4)) * CDbl(.TextMatrix(curRow, 2)), "#,##0.00")
                                End If
                                tmpTotalAmt = tmpTotalAmt + CDbl(.TextMatrix(curRow, 6))
                                curRow = curRow + 1
                                If .Rows - 1 = curRow Then .Rows = .Rows + 5
                            End If
                        Else
                            .TextMatrix(curRow, 7) = tmpItem.ID
                            .TextMatrix(curRow, 0) = tmpItem.JobDataCarID & "_" & tmpItem.TripNo
                            .TextMatrix(curRow, 1) = "น้ำหนักรวม"
                            .TextMatrix(curRow, 2) = Format(nonGroupWeight, "#,##0")
                            .TextMatrix(curRow, 3) = "กิโลกรัม"
                            .TextMatrix(curRow, 4) = Format(maxNonTreatRate, "#,##0.00")
                            .TextMatrix(curRow, 5) = maxNonUnit
                            .TextMatrix(curRow, 6) = Format(nonGroupWeight * maxNonTreatRate, "#,##0.00")
                            tmpTotalAmt = tmpTotalAmt + Format(nonGroupWeight * maxNonTreatRate, "0.00") * 1
                            curRow = curRow + 1
                            If .Rows - 1 = curRow Then .Rows = .Rows + 5
                        End If
                        nonGroupWeight = 0
                        maxNonTreatRate = 0
                    Else
                        GoTo RatePerTrip
                    End If
                    
                Case "HAZ", "HBL", "HIC", "HBI", "BBH", "BBL", "NIC", "NBI"
                    If InStr(1, Trim(curTreatUnitID), "เที่ยว") = 0 Then
                        If Trim(colMinWeightCheck("'" & tmpTimeTableNo & "'").MaxWasteID) <> "" Then
                            Set tmpWaste = colMinWeightCheck("'" & tmpTimeTableNo & "'").ColWasteWeight("'" & colMinWeightCheck("'" & tmpTimeTableNo & "'").MaxWasteID & "'")
                            If InStr(1, tmpWaste.TreatmentUnitID, "ตัน") > 0 Then
                                If LCase(Trim(tmpWaste.WasteType)) = LCase(Trim(tmpWasteType)) And Trim(colMinWeightCheck("'" & tmpTimeTableNo & "'").menifestID) = Trim(tmpItem.ID) Then
                                    If colMinWeightCheck("'" & tmpTimeTableNo & "'").isUnderWeight Then
                                        .TextMatrix(curRow, 7) = tmpItem.ID
                                        .TextMatrix(curRow, 0) = tmpItem.JobDataCarID & "_" & tmpItem.TripNo
                                        If Trim(.TextMatrix(curRow, 1)) <> "" Then
                                            curRow = curRow + 1
                                            If .Rows - 1 = curRow Then .Rows = .Rows + 5
                                        End If
                                        .TextMatrix(curRow, 1) = "น้ำหนักไม่ถึง " & tmpWaste.MinWeightPerCar & " ตัน คิดน้ำหนักเพิ่มอีก"
                                        .TextMatrix(curRow, 2) = Format(tmpWaste.MinWeightPerCar - colMinWeightCheck("'" & tmpTimeTableNo & "'").TotalWeight, "#,##0.000")
                                        .TextMatrix(curRow, 3) = "ตัน"
                                        If LCase(tmpWaste.WasteType) = "haz" Then
                                            If tmpWaste.HazMinRate > 0 Then
                                                .TextMatrix(curRow, 4) = Format(tmpWaste.HazMinRate, "#,##0.00")
                                            Else
                                                .TextMatrix(curRow, 4) = Format(tmpWaste.TreatmentFee, "#,##0.00")
                                            End If
                                        ElseIf LCase(tmpWaste.WasteType) = "hbl" Then
                                            If tmpWaste.HBLMinRate > 0 Then
                                                .TextMatrix(curRow, 4) = Format(tmpWaste.HBLMinRate, "#,##0.00")
                                            Else
                                                .TextMatrix(curRow, 4) = Format(tmpWaste.TreatmentFee, "#,##0.00")
                                            End If
                                        Else
                                            If tmpWaste.HazMinRate > 0 Then
                                                .TextMatrix(curRow, 4) = Format(tmpWaste.HazMinRate, "#,##0.00")
                                            Else
                                                .TextMatrix(curRow, 4) = Format(tmpWaste.TreatmentFee, "#,##0.00")
                                            End If
                                        End If
                                        .TextMatrix(curRow, 5) = "บาท/ตัน"
                                        .TextMatrix(curRow, 6) = Format(CDbl(.TextMatrix(curRow, 4)) * CDbl(.TextMatrix(curRow, 2)), "#,##0.00")
                                        tmpTotalAmt = tmpTotalAmt + CDbl(.TextMatrix(curRow, 6))
                                        curRow = curRow + 1
                                        If .Rows - 1 = curRow Then .Rows = .Rows + 5
                                    ElseIf colMinWeightCheck("'" & tmpTimeTableNo & "'").isOverWeight Then
                                        .TextMatrix(curRow, 7) = tmpItem.ID
                                        .TextMatrix(curRow, 0) = tmpItem.JobDataCarID & "_" & tmpItem.TripNo
                                        If Trim(.TextMatrix(curRow, 1)) <> "" Then
                                            curRow = curRow + 1
                                            If .Rows - 1 = curRow Then .Rows = .Rows + 5
                                        End If
                                        .TextMatrix(curRow, 1) = "น้ำหนักเกินที่กำหนด " & tmpWaste.MaxWeightPerCar & " ตัน คิดน้ำหนักส่วนเกิน"
                                        .TextMatrix(curRow, 2) = Format(colMinWeightCheck("'" & tmpTimeTableNo & "'").TotalWeight - tmpWaste.MaxWeightPerCar, "#,##0.000")
                                        .TextMatrix(curRow, 3) = "ตัน"
                                        .TextMatrix(curRow, 4) = Format(tmpWaste.MaxRate, "#,##0.00")
                                        .TextMatrix(curRow, 5) = tmpWaste.MaxRateType
                                        If InStr(1, tmpWaste.MaxRateType, "เที่ยว") > 0 Then
                                            .TextMatrix(curRow, 6) = Format(tmpWaste.MaxRate, "#,##0.00")
                                        Else
                                            .TextMatrix(curRow, 6) = Format(CDbl(.TextMatrix(curRow, 4)) * CDbl(.TextMatrix(curRow, 2)), "#,##0.00")
                                        End If
                                        tmpTotalAmt = tmpTotalAmt + CDbl(.TextMatrix(curRow, 6))
                                        curRow = curRow + 1
                                        If .Rows - 1 = curRow Then .Rows = .Rows + 5
                                    End If
                                End If
                            ElseIf InStr(1, tmpWaste.TreatmentUnitID, "กิโลกรัม") Then
                                If LCase(Trim(tmpWaste.WasteType)) = LCase(Trim(tmpWasteType)) And Trim(colMinWeightCheck("'" & tmpTimeTableNo & "'").menifestID) = Trim(tmpItem.ID) Then
                                    If colMinWeightCheck("'" & tmpTimeTableNo & "'").isUnderWeight Then
                                        .TextMatrix(curRow, 7) = tmpItem.ID
                                        .TextMatrix(curRow, 0) = tmpItem.JobDataCarID & "_" & tmpItem.TripNo
                                        If Trim(.TextMatrix(curRow, 1)) <> "" Then
                                            curRow = curRow + 1
                                            If .Rows - 1 = curRow Then .Rows = .Rows + 5
                                        End If
                                        .TextMatrix(curRow, 1) = "น้ำหนักไม่ถึง " & tmpWaste.MinWeightPerCar & " กิโลกรัม คิดน้ำหนักเพิ่มอีก"
                                        .TextMatrix(curRow, 2) = Format((tmpWaste.MinWeightPerCar - colMinWeightCheck("'" & tmpTimeTableNo & "'").TotalWeight) * 1000, "#,##0")
                                        .TextMatrix(curRow, 3) = "กิโลกรัม"
                                        If LCase(tmpWaste.WasteType) = "haz" Then
                                            If tmpWaste.HazMinRate > 0 Then
                                                .TextMatrix(curRow, 4) = Format(tmpWaste.HazMinRate, "#,##0.00")
                                            Else
                                                .TextMatrix(curRow, 4) = Format(tmpWaste.TreatmentFee, "#,##0.00")
                                            End If
                                        ElseIf LCase(tmpWaste.WasteType) = "hbl" Then
                                            If tmpWaste.HBLMinRate > 0 Then
                                                .TextMatrix(curRow, 4) = Format(tmpWaste.HBLMinRate, "#,##0.00")
                                            Else
                                                .TextMatrix(curRow, 4) = Format(tmpWaste.TreatmentFee, "#,##0.00")
                                            End If
                                        Else
                                            If tmpWaste.HazMinRate > 0 Then
                                                .TextMatrix(curRow, 4) = Format(tmpWaste.HazMinRate, "#,##0.00")
                                            Else
                                                .TextMatrix(curRow, 4) = Format(tmpWaste.TreatmentFee, "#,##0.00")
                                            End If
                                        End If
                                        .TextMatrix(curRow, 5) = "บาท/กิโลกรัม"
                                        .TextMatrix(curRow, 6) = Format(CDbl(.TextMatrix(curRow, 4)) * CDbl(.TextMatrix(curRow, 2)), "#,##0.00")
                                        tmpTotalAmt = tmpTotalAmt + CDbl(.TextMatrix(curRow, 6))
                                        curRow = curRow + 1
                                        If .Rows - 1 = curRow Then .Rows = .Rows + 5
                                    ElseIf colMinWeightCheck("'" & tmpTimeTableNo & "'").isOverWeight Then
                                        .TextMatrix(curRow, 7) = tmpItem.ID
                                        .TextMatrix(curRow, 0) = tmpItem.JobDataCarID & "_" & tmpItem.TripNo
                                        If Trim(.TextMatrix(curRow, 1)) <> "" Then
                                            curRow = curRow + 1
                                            If .Rows - 1 = curRow Then .Rows = .Rows + 5
                                        End If
                                        .TextMatrix(curRow, 1) = "น้ำหนักเกินที่กำหนด " & tmpWaste.MaxWeightPerCar & "กิโลกรัม  คิดน้ำหนักส่วนเกิน"
                                        .TextMatrix(curRow, 2) = Format((colMinWeightCheck("'" & tmpTimeTableNo & "'").TotalWeight - tmpWaste.MaxWeightPerCar) * 1000, "#,##0")
                                        .TextMatrix(curRow, 3) = "กิโลกรัม"
                                        .TextMatrix(curRow, 4) = Format(tmpWaste.MaxRate, "#,##0.00")
                                        .TextMatrix(curRow, 5) = tmpWaste.MaxRateType
                                        If InStr(1, tmpWaste.MaxRateType, "เที่ยว") > 0 Then
                                            .TextMatrix(curRow, 6) = Format(tmpWaste.MaxRate, "#,##0.00")
                                        Else
                                            .TextMatrix(curRow, 6) = Format(CDbl(.TextMatrix(curRow, 4)) * CDbl(.TextMatrix(curRow, 2)), "#,##0.00")
                                        End If
                                        tmpTotalAmt = tmpTotalAmt + CDbl(.TextMatrix(curRow, 6))
                                        curRow = curRow + 1
                                        If .Rows - 1 = curRow Then .Rows = .Rows + 5
                                    End If
                                End If
                            End If
                            Set tmpWaste = Nothing
                        Else
                            GoTo RatePerTrip
                        End If
                    Else
                        GoTo RatePerTrip
                    End If
                    
                Case Else
RatePerTrip:
                    If Trim(curTreatUnitID) <> "" Then
                     If InStr(1, tmpWaste.TreatmentUnitID, "กิโลกรัม") Then
                      If colMinWeightCheck("'" & tmpTimeTableNo & "'").isOverWeight And LCase(Trim(colMinWeightCheck("'" & tmpTimeTableNo & "'").MaxWasteType)) = LCase(Trim(tmpWasteType)) Then
                             Set tmpWaste = colMinWeightCheck("'" & tmpTimeTableNo & "'").ColWasteWeight("'" & colMinWeightCheck("'" & tmpTimeTableNo & "'").MaxWasteID & "'")
                             .TextMatrix(curRow, 0) = tmpItem.JobDataCarID & "_" & tmpItem.TripNo
                             .TextMatrix(curRow, 1) = "น้ำหนักรวม"
                             .TextMatrix(curRow, 2) = Format(curGroupWeight, "#,##0")
                             .TextMatrix(curRow, 3) = "กิโลกรัม"
                             .TextMatrix(curRow, 4) = ""
                             .TextMatrix(curRow, 5) = ""
                             .TextMatrix(curRow, 6) = ""
                             .TextMatrix(curRow, 7) = tmpItem.ID
                             If Trim(.TextMatrix(curRow, 1)) <> "" Then
                                 curRow = curRow + 1
                                 If .Rows - 1 = curRow Then .Rows = .Rows + 5
                             End If
                             If Trim(.TextMatrix(curRow, 1)) <> "" Then
                                 curRow = curRow + 1
                                 If .Rows - 1 = curRow Then .Rows = .Rows + 5
                             End If
                             If tmpWaste.menifestID = tmpItem.ID Then
                                 .TextMatrix(curRow, 0) = tmpItem.JobDataCarID & "_" & tmpItem.TripNo
                                 .TextMatrix(curRow, 1) = "น้ำหนักเกินที่กำหนด " & tmpWaste.MaxWeightPerCar & " กิโลกรัม คิดน้ำหนักส่วนเกิน"
                                 .TextMatrix(curRow, 2) = Format((colMinWeightCheck("'" & tmpTimeTableNo & "'").TotalWeight - tmpWaste.MaxWeightPerCar) * 1000, "#,##0")
                                 .TextMatrix(curRow, 3) = "กิโลกรัม"
                                 .TextMatrix(curRow, 4) = Format(tmpWaste.MaxRate, "#,##0.00")
                                 .TextMatrix(curRow, 5) = tmpWaste.MaxRateType
                                 If InStr(1, tmpWaste.MaxRateType, "เที่ยว") > 0 Then
                                     .TextMatrix(curRow, 6) = Format(tmpWaste.MaxRate, "#,##0.00")
                                 Else
                                     .TextMatrix(curRow, 6) = Format(CDbl(.TextMatrix(curRow, 4)) * CDbl(.TextMatrix(curRow, 2)), "#,##0.00")
                                 End If
                                 .TextMatrix(curRow, 7) = tmpItem.ID
                                 tmpTotalAmt = tmpTotalAmt + CDbl(.TextMatrix(curRow, 6))
                             End If
                             curRow = curRow + 1
                             If .Rows - 1 = curRow Then .Rows = .Rows + 5
                         Else
                             .TextMatrix(curRow, 0) = tmpItem.JobDataCarID & "_" & tmpItem.TripNo
                             .TextMatrix(curRow, 1) = "น้ำหนักรวม"
                             .TextMatrix(curRow, 2) = Format(curGroupWeight, "#,##0")
                             .TextMatrix(curRow, 3) = "กิโลกรัม"
                             .TextMatrix(curRow, 4) = ""
                             .TextMatrix(curRow, 5) = ""
                             .TextMatrix(curRow, 6) = ""
                             .TextMatrix(curRow, 7) = tmpItem.ID
                             curRow = curRow + 1
                             If .Rows - 1 = curRow Then .Rows = .Rows + 5
                             curGroupWeight = 0
                         End If
                     Else
                         If colMinWeightCheck("'" & tmpTimeTableNo & "'").isOverWeight And LCase(Trim(colMinWeightCheck("'" & tmpTimeTableNo & "'").MaxWasteType)) = LCase(Trim(tmpWasteType)) Then
                             Set tmpWaste = colMinWeightCheck("'" & tmpTimeTableNo & "'").ColWasteWeight("'" & colMinWeightCheck("'" & tmpTimeTableNo & "'").MaxWasteID & "'")
                             .TextMatrix(curRow, 0) = tmpItem.JobDataCarID & "_" & tmpItem.TripNo
                             .TextMatrix(curRow, 1) = "น้ำหนักรวม"
                             .TextMatrix(curRow, 2) = Format(curGroupWeight, "#,##0")
                             .TextMatrix(curRow, 3) = "กิโลกรัม"
                             .TextMatrix(curRow, 4) = ""
                             .TextMatrix(curRow, 5) = ""
                             .TextMatrix(curRow, 6) = ""
                             .TextMatrix(curRow, 7) = tmpItem.ID
                             If Trim(.TextMatrix(curRow, 1)) <> "" Then
                                 curRow = curRow + 1
                                 If .Rows - 1 = curRow Then .Rows = .Rows + 5
                             End If
                             If Trim(.TextMatrix(curRow, 1)) <> "" Then
                                 curRow = curRow + 1
                                 If .Rows - 1 = curRow Then .Rows = .Rows + 5
                             End If
                             If tmpWaste.menifestID = tmpItem.ID Then
                                 .TextMatrix(curRow, 0) = tmpItem.JobDataCarID & "_" & tmpItem.TripNo
                                 .TextMatrix(curRow, 1) = "น้ำหนักเกินที่กำหนด " & tmpWaste.MaxWeightPerCar & " ตัน คิดน้ำหนักส่วนเกิน"
                                 .TextMatrix(curRow, 2) = Format((colMinWeightCheck("'" & tmpTimeTableNo & "'").TotalWeight - tmpWaste.MaxWeightPerCar) * 1000, "#,##0")
                                 .TextMatrix(curRow, 3) = "กิโลกรัม"
                                 .TextMatrix(curRow, 4) = Format(tmpWaste.MaxRate, "#,##0.00")
                                 .TextMatrix(curRow, 5) = tmpWaste.MaxRateType
                                 If InStr(1, tmpWaste.MaxRateType, "เที่ยว") > 0 Then
                                     .TextMatrix(curRow, 6) = Format(tmpWaste.MaxRate, "#,##0.00")
                                 Else
                                     .TextMatrix(curRow, 6) = Format(CDbl(.TextMatrix(curRow, 4)) * CDbl(.TextMatrix(curRow, 2)), "#,##0.00")
                                 End If
                                 .TextMatrix(curRow, 7) = tmpItem.ID
                                 tmpTotalAmt = tmpTotalAmt + CDbl(.TextMatrix(curRow, 6))
                             End If
                             curRow = curRow + 1
                             If .Rows - 1 = curRow Then .Rows = .Rows + 5
                         Else
                             .TextMatrix(curRow, 0) = tmpItem.JobDataCarID & "_" & tmpItem.TripNo
                             .TextMatrix(curRow, 1) = "น้ำหนักรวม"
                             .TextMatrix(curRow, 2) = Format(curGroupWeight, "#,##0")
                             .TextMatrix(curRow, 3) = "กิโลกรัม"
                             .TextMatrix(curRow, 4) = ""
                             .TextMatrix(curRow, 5) = ""
                             .TextMatrix(curRow, 6) = ""
                             .TextMatrix(curRow, 7) = tmpItem.ID
                             curRow = curRow + 1
                             If .Rows - 1 = curRow Then .Rows = .Rows + 5
                             curGroupWeight = 0
                         End If
                     End If
                    End If
            End Select
            Set tmpItem = Nothing
        Next x
    End With
    Exit Sub
ErrD:
    Resume Next
'    MsgBox Err.Description & vbCrLf & "เกิดข้อผิดพลาดกรุณาลองใหม่ หรือติดต่อผู้พัฒนาระบบ", vbCritical, "Error"
'    Exit Sub
End Sub

Private Sub CreateNewBillHeader()
    Set myBillHeader_ = New BWGBillingHeader
    With myBillHeader_
        .ID = GetGUID
        .IssuedDate = dtBillDate.ValueYMD
        .PayDueDate = dtDueDate.ValueYMD
        'myBillHeader_.va = SystemConfig.VATRate
        If cboCompany.ListCount > 1 Then
            If cboCompany.ListIndex > 0 Then
                .CompanyID = CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).ID
                .CompanyName = cboCompany.Text
            End If
        Else
            If cboCompany.ListIndex >= 0 Then
                .CompanyID = CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).ID
                .CompanyName = cboCompany.Text
            End If
        End If
        .isChanged = False
        SetCancelScreen
    End With
End Sub

Public Sub RemoveMenifest(selID As String)
Dim tmpCol As Collection
Dim x%, tmpCol2 As Collection, tmpS As New BWGSearchManager
Dim tmpKey As String
    Set tmpCol = myBillHeader_.BillingDetails
    If tmpCol Is Nothing Then Exit Sub
    For x = 1 To tmpCol.Count
        If tmpCol(x).menifestID = selID Then
            tmpKey = "'" & tmpCol(x).ID & "'"
            Set tmpCol2 = tmpS.MenifestSearch("MenifestID='" & Trim(selID) & "'")
            If tmpCol2.Count > 0 Then
                If isExist(myBillHeader_.TripControls, "'" & tmpCol2(1).JobDataCarID & "_" & tmpCol2(1).TripNo & "'") Then
                    myBillHeader_.TripControls.Remove "'" & tmpCol2(1).JobDataCarID & "_" & tmpCol2(1).TripNo & "'"
                End If
            End If
            Set tmpCol2 = Nothing
            If tmpCol(x).menifestID <> "" Then
'                DBConnExc "UPDATE tbMenifestHeader SET isInv = 'N', UserInv = '' WHERE MenifestID = '" & tmpCol(x).MenifestID & "'"
            End If
            Exit For
        End If
    Next
    If isExist(tmpCol, tmpKey) Then
        tmpCol.Remove tmpKey
        isDetailChange_ = True
        myBillHeader_.isChanged = True
    End If
    Set tmpS = Nothing
End Sub

Public Sub AddMenifest(selItem As BWGMenifest)
Dim tmpCol As Collection
Dim x%, isNew As Boolean
Dim tmpID As String, tmpTrip As BWGJobDataTimeTable
Dim tmpDetail As New BWGBillingDetail
    Set tmpCol = myBillHeader_.BillingDetails
    'Set tmpTrip = selItem.getTripControl
    
    Set tmpTrip = New BWGJobDataTimeTable
    With selItem
        tmpTrip.JobDataCarID = .JobDataCarID
        tmpTrip.TripNo = .TripNo
        tmpTrip.isMain = .isMain
        tmpTrip.TimeTableNo = .TimeTableNo
        tmpTrip.CarRegisNo = .CarRegisID
        tmpTrip.CustBillID = ""
        tmpTrip.TripEquipFee = .TripEquipFee
        tmpTrip.TripMoreWorker = .TripMoreWorker
        tmpTrip.TripMoreWorkerFee = .TripMoreWorkerFee
        tmpTrip.TripOtherFee = .TripOtherFee
        tmpTrip.TripTranFee = .TripTranFee
        tmpTrip.TripAllFee = .TripEquipFee + .TripMoreWorkerFee + .TripOtherFee + .TripTranFee
    End With
    If isExist(myBillHeader_.TripControls, "'" & tmpTrip.JobDataCarID & "_" & tmpTrip.TripNo & "'") = False Then
        'tmpTrip.CustBillID = myBillHeader_.ID
        myBillHeader_.TripControls.Add tmpTrip, "'" & tmpTrip.JobDataCarID & "_" & tmpTrip.TripNo & "'"
    End If

    If tmpCol Is Nothing Then Set tmpCol = New Collection
    isNew = True
    tmpID = Trim(selItem.ID)
    For x = 1 To tmpCol.Count
        If tmpCol(x).menifestID = tmpID Then
            Set tmpDetail = tmpCol(x)
            isNew = False
            Exit For
        End If
    Next

    If isNew Then
        tmpDetail.ID = GetGUID
        tmpCol.Add tmpDetail, "'" & tmpDetail.ID & "'"
    End If
    tmpDetail.AmtB4VAT = selItem.InCollectFee + selItem.InEquipFee + selItem.InOtherFee + selItem.IntransFee + selItem.InWorkerFee
    tmpDetail.BillDesc = "ค่ากำจัดกาก " & selItem.MenifestNo
    tmpDetail.BillHeaderID = myBillHeader_.ID
    tmpDetail.DiscountAmt = 0
    tmpDetail.DiscountType = ""
    tmpDetail.MenifestAddr = cboSiteName.Text
    tmpDetail.MenifestAmt = selItem.InCollectFee + selItem.InEquipFee + selItem.InOtherFee + selItem.IntransFee + selItem.InWorkerFee
    tmpDetail.MenifestDate = selItem.WorkDate
    tmpDetail.MenifestDocNo = selItem.DocNo
    tmpDetail.menifestID = selItem.ID
    tmpDetail.MenifestNo = selItem.MenifestNo
    tmpDetail.WasteType = selItem.NonOrHz
    tmpDetail.RowID = tmpCol.Count
    tmpDetail.TotalAmtToPay = selItem.InCollectFee + selItem.InEquipFee + selItem.InOtherFee + selItem.IntransFee + selItem.InWorkerFee
    tmpDetail.VATAmt = Format(tmpDetail.AmtB4VAT * (1 + SystemConfig.VATRate / 100), "0.00")
    tmpDetail.VATRate = SystemConfig.VATRate
    isDetailChange_ = True
    myBillHeader_.isChanged = True
    Set tmpDetail = Nothing

    If myBillHeader_.BillingDetails Is Nothing Then
        Set myBillHeader_.BillingDetails = tmpCol
        isDetailChange_ = True
        myBillHeader_.isChanged = True
    End If
    Set tmpCol = Nothing
End Sub

Public Function GetMenifests() As Collection
Dim x%, tmpS As New BWGSearchManager
Dim tmpCol As New Collection
Dim tmpDet As BWGBillingDetail
Dim selCol As Collection, tmpStr As String
    Set selCol = myBillHeader_.BillingDetails
    With selCol
        tmpStr = "'A'"
        For x = 1 To .Count
            Set tmpDet = .item(x)
            If Trim(tmpDet.menifestID) <> "" Then
                tmpStr = tmpStr & ",'" & tmpDet.menifestID & "'"
            End If
            Set tmpDet = Nothing
        Next
        Set tmpCol = tmpS.MenifestSearch("MenifestID in (" & tmpStr & ")")
    End With

    Set GetMenifests = tmpCol
    Set tmpCol = Nothing
End Function

Private Sub txtRecBillStaff_Validate(Cancel As Boolean)
    myBillHeader_.BillRecStaff = Trim(txtRecBillStaff.Text)
End Sub

Private Sub txtRem_Validate(Cancel As Boolean)
    myBillHeader_.BillRemark = Trim(txtRem.Text)
End Sub

Private Sub ClearScreen()
    txtBillNo.Text = "== AUTO =="
    dtBillDate.ValueYMD = TodayDate
    grMenifest.ClearAllData
    grMenifest.getGridObj.Rows = 50
    txtCustomer.Text = ""
    Set curCustomer_ = Nothing
    Set curCustomer_ = New BWGCustomer
    chkAccepted.Value = 0
    dtDueDate.ValueYMD = TodayDate
    cboSiteName.Clear
    cboWTax.ListIndex = 0
    txtBillDesc.Text = ""
    txtBillName.Text = ""
    txtCheqDesc.Text = ""
    txtRecBillStaff.Text = ""
    txtRem.Text = ""
    lvTrip.ListItems.Clear
    grItem.ClearAllData
    grItem.getGridObj.Rows = 50
    lbTotalPrice.Caption = "0.00"
    lbWTax.Caption = "0.00"
    lbVat.Caption = "0.00"
    lbNetPrice.Caption = "0.00"
    
    CurPage = 1
    pageCount = 1
    X1 = 0
    X2 = 0
    
    lblLastUpdate.Caption = ""
    lblLastUpdateRE.Caption = ""
End Sub

Private Sub SetCancelScreen()
Dim cVal As Boolean
Dim UserGrp As Collection
Dim Group As New AWSUserGroup
Dim StrGrp As String
Dim i As Integer
    cVal = (myBillHeader_.isCancel)
    lbCancel.Visible = cVal
    Frame2.Visible = cVal
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
        cmdSave.Enabled = True
        cmdSelect.Enabled = True
        cmdReceive.Enabled = True
    Else
        cmdSave.Enabled = Not cVal
        cmdSelect.Enabled = Not cVal
        cmdReceive.Enabled = Not cVal
        cmdCancel.Enabled = Not cVal
    End If
    
    If myBillHeader_.isCancel Then
        lbCancel.Caption = "ยกเลิก"
        Frame2.Visible = True
        Frame2.BackColor = vbRed
        lblCancelDetail.BackColor = vbRed
        lblCancelDetail.Caption = "ยกเลิกโดย : " & myBillHeader_.CancelStaff & vbCrLf & myBillHeader_.CancelDetail
    ElseIf myBillHeader_.isCustAccepted Then
        lbCancel.Caption = "Accepted"
        lbCancel.Visible = True
    End If
    Screen.MousePointer = 0
    
    CurPage = 1
    pageCount = 1
    X1 = 0
    X2 = 0
End Sub

Public Sub SetDataBillHeader(selBill As BWGBillingHeader)
Dim tmpS As New BWGSearchManager
Dim tmpCol As Collection
Dim x%
    Set myBillHeader_ = selBill
    With myBillHeader_
        If .BillingNo <> "" Then
            txtBillNo.Text = .BillingNo
        Else
            txtBillNo.Text = .PreInvNo
        End If
            dtBillDate.ValueYMD = .IssuedDate
        SetTextToCombo .CompanyName, cboCompany
        Call cboCompany_Validate(False)
        Set tmpCol = tmpS.CustomerSearch("CustomerID='" & .CustomerID & "'")
        If tmpCol.Count > 1 Then
            SetDataCustomer tmpCol(1), False
        ElseIf tmpCol.Count = 1 Then
            SetDataCustomer tmpCol(1), True
        End If
        If .isCustAccepted Then
            chkAccepted.Value = 1
        Else
            chkAccepted.Value = 0
        End If
        If .isServiceCustomer Then
            ChkSC.Value = 1
        Else
            ChkSC.Value = 0
        End If
        dtDueDate.ValueYMD = .PayDueDate
        ctlDatepast.ValueYMD = .DatePast
        SetTextToCombo .BillAddr, cboSiteName
        SetTextToCombo .WTaxRate, cboWTax
        txtRecBillStaff.Text = .BillRecStaff
        txtRem.Text = .BillRemark
        CheckVAT.Value = .IncVAT
        
        lblLastUpdate.Caption = "อัพเดทล่าสุดโดย : " & .LastUser & " " & Format(.LastUpdate, "dd/MM/yyyy เวลา HH:mm:ss")
        lblLastUpdateRE.Caption = "กำหนดรับวางบิลโดย : " & .LastUserRE & " " & Format(.LastUpdateRE, "dd/MM/yyyy เวลา HH:mm:ss")
        
        ShowBillingDetail
        ShowOtherBillDet
        ShowTripControl
        SumMoneyAmt
    End With
    SetCancelScreen
End Sub

Private Sub ShowOtherBillDet()
Dim x%, rowCount%
Dim tmpDet As BWGBillingDetail
    grItem.ClearAllData
    With grItem.getGridObj
        .Rows = myBillHeader_.BillingDetails.Count + 50
        rowCount = 1
        For x = 1 To myBillHeader_.BillingDetails.Count
            Set tmpDet = myBillHeader_.BillingDetails(x)
            If Trim(tmpDet.menifestID) = "" Then
                .TextMatrix(rowCount, 0) = tmpDet.ID
                .TextMatrix(rowCount, 1) = tmpDet.BillDesc
                .TextMatrix(rowCount, 2) = tmpDet.AmtB4VAT
                rowCount = rowCount + 1
            End If
            Set tmpDet = Nothing
        Next
    End With
End Sub

Private Sub ShowTripControl()
Dim tmpList As ListItem
Dim x%, tmpTrip As BWGJobDataTimeTable, y%
    lvTrip.ListItems.Clear
    For x = 1 To myBillHeader_.TripControls.Count
        Set tmpTrip = myBillHeader_.TripControls(x)
        If lvTrip.FindItem(tmpTrip.TimeTableNo) Is Nothing Then
            Set tmpList = lvTrip.ListItems.Add(, "'" & tmpTrip.JobDataCarID & "_" & tmpTrip.TripNo & "'", tmpTrip.TimeTableNo)
            tmpList.SubItems(1) = tmpTrip.CarRegisNo
            If Trim(tmpTrip.CustBillID) <> "" And Trim(tmpTrip.CustBillID) <> Trim(myBillHeader_.ID) Then
                tmpList.SubItems(2) = Format(tmpTrip.TripAllFee, "#,##0.00") & " (Billed)"
                tmpList.ForeColor = vbRed
                tmpList.ListSubItems(1).ForeColor = vbRed
                tmpList.ListSubItems(2).ForeColor = vbRed
            Else
                tmpList.SubItems(2) = Format(tmpTrip.TripEquipFee + tmpTrip.TripMoreWorkerFee + tmpTrip.TripOtherFee + tmpTrip.TripTranFee, "#,##0.00")
                tmpList.ForeColor = vbBlack
                tmpList.ListSubItems(1).ForeColor = vbBlack
                tmpList.ListSubItems(2).ForeColor = vbBlack
            End If
            tmpList.Checked = Trim(tmpTrip.CustBillID) = Trim(myBillHeader_.ID)
        Else
            Set tmpList = lvTrip.FindItem(tmpTrip.TimeTableNo)
            If tmpTrip.isMain Then
                tmpList.SubItems(1) = tmpTrip.CarRegisNo & ", " & tmpList.SubItems(1)
            Else
                tmpList.SubItems(1) = tmpList.SubItems(1) & ", " & tmpTrip.CarRegisNo
            End If
        End If
        Set tmpList = Nothing
        Set tmpTrip = Nothing
    Next
End Sub

Public Property Get RefID() As String
    RefID = myBillHeader_.ID
End Property

Private Sub PrintNow(selRpt As CrystalReport, ByRef curRow%, rowPerPage%, ByRef curRecNo%, detailCount%)
Dim tmpSite As BWGCustomerBillAddr
Dim tmpAmt1 As Double, tmpAmt2 As Double
    MaxPrintPage = 20
    If (curRow >= rowPerPage Or curRecNo >= detailCount) And curRecNo <> -999 Then
        With selRpt
            .Formulas(0) = "InvNo=" & Chr(34) & Trim(txtBillNo.Text) & Chr(34)
            .Formulas(1) = "InvDate=" & Chr(34) & Trim(dtBillDate.ValueDMY) & Chr(34)
            '.Formulas(2) = "CustName=" & Chr(34) & Trim(txtCustomer.Text) & Chr(34)
            If cboSiteName.ListIndex > -1 Then
                Set tmpSite = curCustomer_.BillAddresses(cboSiteName.ListIndex + 1)
                .Formulas(2) = "CustName=" & Chr(34) & "(" & curCustomer_.CustomerCode & ")" & " " & Trim(tmpSite.BillName) & Chr(34)
                If InStr(1, tmpSite.ProvinceName, "กรุงเทพ") > 0 Then
                    .Formulas(3) = "CustAddr=" & Chr(34) & tmpSite.Address & " แขวง " & tmpSite.TumbolName & " เขต " & tmpSite.AumphurName & " " & tmpSite.ProvinceName & " " & tmpSite.ZipCode & "     เลขประจำตัวผู้เสียภาษีลูกค้า " & tmpSite.TaxNum & Chr(34)
                Else
                    .Formulas(3) = "CustAddr=" & Chr(34) & tmpSite.Address & " ต. " & tmpSite.TumbolName & " อ. " & tmpSite.AumphurName & " จ. " & tmpSite.ProvinceName & " " & tmpSite.ZipCode & "     เลขประจำตัวผู้เสียภาษีลูกค้า " & tmpSite.TaxNum & Chr(34)
                End If
                Set tmpSite = Nothing
            End If
            If curRecNo >= detailCount Then
                
                If Trim(myBillHeader_.BillRemark) <> "" Then
                        curRow = curRow + 1
                        .Formulas(20 + (curRow - 1) * 4) = "Col1_" & curRow & "=" & Chr(34) & myBillHeader_.BillRemark & Chr(34)
                        .Formulas(21 + (curRow - 1) * 4) = "Col2_" & curRow & "=" & Chr(34) & "" & Chr(34)
                        .Formulas(22 + (curRow - 1) * 4) = "Col3_" & curRow & "=" & Chr(34) & "" & Chr(34)
                        .Formulas(23 + (curRow - 1) * 4) = "Col4_" & curRow & "=" & Chr(34) & "" & Chr(34)
                End If
                
                'tmpAmt1 = Format(totalAmt, "0.00")
                'tmpAmt2 = Format(tmpAmt1 * SystemConfig.VATRate / 100, "0.00")
                .Formulas(4) = "TotalAmt=" & Chr(34) & Format(lbTotalPrice, "#,##0.00") & Chr(34)
                .Formulas(5) = "VATAmt=" & Chr(34) & Format(lbVat, "#,##0.00") & Chr(34)
                .Formulas(6) = "GrandTotalAmt=" & Chr(34) & Format(lbNetPrice, "#,##0.00") & Chr(34)
                .Formulas(7) = "AMTWord=" & Chr(34) & Trim(AmountToString(Format2Decimal(lbNetPrice))) & Chr(34)
                .Formulas(11) = "WTaxAMT=" & Chr(34) & Format(lbWTax, "#,##0.00") & Chr(34)
                '.Formulas(9) = "DueMonth=" & Chr(34) & Trim(dtDueDate.MonthValue) & Chr(34)
                '.Formulas(10) = "DueYear=" & Chr(34) & Trim(dtDueDate.YearValue) & Chr(34)
            End If
'                 Set tmpSite = curCustomer_.BillAddresses(cboSiteName.ListIndex + 1)
'                .Formulas(99) = "ZCustTax=" & Chr(34) & "เลขประจำตัวผู้เสียภาษีลูกค้า " & tmpSite.TaxNum & Chr(34)
'                Set tmpSite = Nothing
            X2 = Int(CurPage / MaxPrintPage)
            pageCount = X1 * MaxPrintPage
            If CurPage > pageCount Then
                .WindowTitle = "Page " & CurPage
                .PrintReport
                PageContinue = False
            Else
                CurPage = CurPage + 1
                If curRow >= rowPerPage Then curRow = 1
                If curRecNo >= detailCount Then curRecNo = -999
                Exit Sub
            End If
            If CurPage = (pageCount * X1) + MaxPrintPage Then
                MsgBox "ใบวางบิลที่ต้องการพิมพ์มีมากกว่า 20 หน้า กรุณาพิมพ์ 20 หน้าแรกก่อน แล้วปิดหน้าที่พิมพ์เสร็จแล้ว" & _
                                    vbCrLf & "หลังจากนั้นจึงคลิกปุ่ม พิมพ์ อีกครั้งเพื่อพิมพ์ส่วนที่เหลือ", vbInformation
                PageContinue = True
                ClearInvReport rpt1
                If curRow >= rowPerPage Then curRow = 1
                If curRecNo >= detailCount Then curRecNo = -999
                Exit Sub
            End If
            CurPage = CurPage + 1
            
            
            If Trim(.LastErrorString) <> "" Then MsgBox .LastErrorString
            .LastErrorString = ""
            ClearInvReport rpt1
            If curRow >= rowPerPage Then curRow = 1
            If curRecNo >= detailCount Then curRecNo = -999
        End With
    End If
End Sub
