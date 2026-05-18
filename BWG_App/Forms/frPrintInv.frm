VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmReceipt 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "F087"
   ClientHeight    =   8865
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11070
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   11070
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox CBReStaff 
      Height          =   315
      ItemData        =   "frPrintInv.frx":0000
      Left            =   7350
      List            =   "frPrintInv.frx":0002
      TabIndex        =   56
      Text            =   "CBReStaff"
      Top             =   2040
      Width           =   3645
   End
   Begin VB.ComboBox CBType 
      Height          =   315
      ItemData        =   "frPrintInv.frx":0004
      Left            =   5535
      List            =   "frPrintInv.frx":0017
      TabIndex        =   55
      Top             =   2040
      Width           =   1785
   End
   Begin VB.TextBox txtHeaderText 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1035
      TabIndex        =   52
      Text            =   "ค่าบริการกำจัดและบำบัด"
      Top             =   1980
      Width           =   3690
   End
   Begin VB.CheckBox CheckTax 
      BackColor       =   &H00FFFFFF&
      Caption         =   "คิดยอด Tax ของใบลดหนี้ ตามใบ INV"
      Height          =   195
      Left            =   4140
      TabIndex        =   51
      Top             =   5820
      Width           =   2955
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "ไม่คิดภาษีมูลค่าเพิ่ม"
      Height          =   195
      Left            =   9180
      TabIndex        =   49
      Top             =   5820
      Width           =   1755
   End
   Begin VB.TextBox txtPaidCash 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   1380
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   7620
      Width           =   1635
   End
   Begin VB.CommandButton cmdBillAdd 
      Caption         =   "เพิ่มใบวางบิล"
      Height          =   795
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   8040
      Width           =   1275
   End
   Begin VB.CommandButton cmdPrint2 
      Caption         =   "พิมพ์ใบสำคัญรับ"
      Height          =   795
      Left            =   2700
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   8040
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.CommandButton cmdPrintInv 
      Caption         =   "พิมพ์ใบเสร็จ"
      Height          =   795
      Left            =   1380
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   8040
      Width           =   1275
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ยกเลิก"
      Height          =   795
      Left            =   8940
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   8040
      Width           =   1005
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "ปิด"
      Height          =   795
      Left            =   9990
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   8040
      Width           =   1035
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "สร้างใหม่"
      Height          =   795
      Left            =   7890
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   8040
      Width           =   1005
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "เปิดข้อมูลเก่า"
      Height          =   795
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   8040
      Width           =   1005
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "เก็บข้อมูล"
      Height          =   795
      Left            =   5790
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   8040
      Width           =   1005
   End
   Begin VB.TextBox txtRem 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   6435
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1230
      Width           =   4530
   End
   Begin VB.ComboBox cboBillAddr 
      Height          =   315
      Left            =   1035
      Style           =   2  'Dropdown List
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1620
      Width           =   9960
   End
   Begin VB.TextBox txtCustomer 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1035
      TabIndex        =   3
      Top             =   1230
      Width           =   4215
   End
   Begin VB.ComboBox cboCompany 
      Height          =   315
      ItemData        =   "frPrintInv.frx":0065
      Left            =   6765
      List            =   "frPrintInv.frx":0067
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   840
      Width           =   4230
   End
   Begin VB.TextBox txtInvNo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1035
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Text            =   "== AUTO =="
      Top             =   840
      Width           =   1635
   End
   Begin Crystal.CrystalReport rpt1 
      Left            =   4200
      Top             =   3630
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "เลือกเครื่องพิมพ์"
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   4620
      TabIndex        =   17
      Top             =   8040
      Visible         =   0   'False
      Width           =   3787
      Begin VB.ComboBox cboPrinter 
         Height          =   315
         Left            =   180
         Style           =   2  'Dropdown List
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   270
         Width           =   3525
      End
   End
   Begin WasteManagment.ctlGrid grBill 
      Height          =   2955
      Left            =   0
      TabIndex        =   6
      Top             =   2400
      Width           =   10995
      _ExtentX        =   19394
      _ExtentY        =   5212
   End
   Begin WasteManagment.ctlDate dtInvDate 
      Height          =   315
      Left            =   3405
      TabIndex        =   1
      Top             =   840
      Width           =   1875
      _ExtentX        =   3307
      _ExtentY        =   556
      BackColor       =   -2147483633
   End
   Begin WasteManagment.ctlGrid grCheq 
      Height          =   855
      Left            =   60
      TabIndex        =   7
      Top             =   6720
      Width           =   7155
      _ExtentX        =   12621
      _ExtentY        =   1508
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "ประเภท :"
      Height          =   195
      Left            =   4815
      TabIndex        =   54
      Top             =   2040
      Width           =   735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "หัวบิล"
      Height          =   195
      Index           =   14
      Left            =   555
      TabIndex        =   53
      Top             =   2040
      Width           =   465
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "ยอดเงินที่แสดงเป็น ยอดรวมก่อน VAT"
      ForeColor       =   &H000000FF&
      Height          =   435
      Left            =   7260
      TabIndex        =   50
      Top             =   5760
      Width           =   1815
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "รวมยอดภาษีหัก ณ ที่จ่าย"
      Height          =   195
      Index           =   13
      Left            =   7380
      TabIndex        =   48
      Top             =   6600
      Width           =   1695
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "บาท"
      Height          =   195
      Index           =   12
      Left            =   10680
      TabIndex        =   47
      Top             =   6600
      Width           =   315
   End
   Begin VB.Label lbWTax 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   9300
      TabIndex        =   46
      Top             =   6600
      Width           =   1275
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "รวมทั้งสิ้น"
      Height          =   195
      Index           =   9
      Left            =   8400
      TabIndex        =   45
      Top             =   7680
      Width           =   675
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "บาท"
      Height          =   195
      Index           =   8
      Left            =   10680
      TabIndex        =   44
      Top             =   7680
      Width           =   315
   End
   Begin VB.Label lbNetTotal 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   9300
      TabIndex        =   43
      Top             =   7680
      Width           =   1275
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ยอด VAT"
      Height          =   195
      Index           =   7
      Left            =   8400
      TabIndex        =   42
      Top             =   7320
      Width           =   735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "บาท"
      Height          =   195
      Index           =   6
      Left            =   10680
      TabIndex        =   41
      Top             =   7320
      Width           =   315
   End
   Begin VB.Label lbVAT 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   9300
      TabIndex        =   40
      Top             =   7320
      Width           =   1275
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "บาท"
      Height          =   195
      Index           =   5
      Left            =   3060
      TabIndex        =   39
      Top             =   7680
      Width           =   315
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ชำระด้วยเงินสด"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   4
      Left            =   180
      TabIndex        =   38
      Top             =   7680
      Width           =   1155
   End
   Begin VB.Label Label3 
      Caption         =   "(กรุณาระบุวันที่ในเช็คเป็น วว/ดด/ปปปป และปีใช้ปี พ.ศ. เท่านั้น)"
      ForeColor       =   &H000000FF&
      Height          =   225
      Index           =   1
      Left            =   1980
      TabIndex        =   35
      Top             =   6450
      Visible         =   0   'False
      Width           =   5115
   End
   Begin VB.Line Line2 
      X1              =   30
      X2              =   11070
      Y1              =   7980
      Y2              =   7980
   End
   Begin VB.Label lbTotalCheq 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   5580
      TabIndex        =   34
      Top             =   7680
      Width           =   1275
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "บาท"
      Height          =   195
      Index           =   4
      Left            =   7020
      TabIndex        =   33
      Top             =   7680
      Width           =   315
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "รวมยอดเช็คสั่งจ่าย"
      Height          =   195
      Index           =   3
      Left            =   4020
      TabIndex        =   32
      Top             =   7680
      Width           =   1335
   End
   Begin VB.Label lbTotalPay 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   9300
      TabIndex        =   31
      Top             =   6960
      Width           =   1275
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "บาท"
      Height          =   195
      Index           =   2
      Left            =   10680
      TabIndex        =   30
      Top             =   6960
      Width           =   315
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "รวมยอดที่ลูกค้าจ่ายจริง"
      Height          =   195
      Index           =   0
      Left            =   7500
      TabIndex        =   29
      Top             =   6960
      Width           =   1635
   End
   Begin VB.Label lbTotalBill 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   9300
      TabIndex        =   28
      Top             =   6240
      Width           =   1275
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "บาท"
      Height          =   195
      Index           =   10
      Left            =   10680
      TabIndex        =   27
      Top             =   6240
      Width           =   315
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "รวมยอดเรียกเก็บ"
      Height          =   195
      Index           =   11
      Left            =   7860
      TabIndex        =   26
      Top             =   6240
      Width           =   1275
   End
   Begin VB.Label Label3 
      Caption         =   " รายการเช็คที่รับชำระเงิน"
      Height          =   225
      Index           =   0
      Left            =   60
      TabIndex        =   25
      Top             =   6450
      Visible         =   0   'False
      Width           =   7125
   End
   Begin VB.Shape Shape1 
      Height          =   885
      Index           =   1
      Left            =   30
      Top             =   6690
      Visible         =   0   'False
      Width           =   7185
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
      Left            =   4440
      TabIndex        =   24
      Top             =   60
      Visible         =   0   'False
      Width           =   2925
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "หมายเหตุ"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   3
      Left            =   5685
      TabIndex        =   23
      Top             =   1290
      Width           =   765
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ที่อยู่"
      Height          =   195
      Index           =   2
      Left            =   630
      TabIndex        =   22
      Top             =   1680
      Width           =   390
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ลูกค้า"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   55
      Left            =   555
      TabIndex        =   21
      Top             =   1320
      Width           =   465
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ออกในนามบริษัท"
      Height          =   195
      Index           =   0
      Left            =   5445
      TabIndex        =   20
      Top             =   900
      Width           =   1305
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ถึงวันที่"
      Height          =   195
      Index           =   1
      Left            =   2805
      TabIndex        =   19
      Top             =   900
      Width           =   585
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "เลขที่ใบเสร็จ"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   68
      Left            =   45
      TabIndex        =   18
      Top             =   900
      Width           =   990
   End
   Begin VB.Shape Shape2 
      Height          =   3285
      Left            =   0
      Top             =   2370
      Width           =   11025
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   4
      Index           =   0
      X1              =   120
      X2              =   6030
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      Index           =   1
      X1              =   210
      X2              =   6180
      Y1              =   570
      Y2              =   570
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ใบเสร็จรับเงิน"
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
      TabIndex        =   16
      Top             =   60
      Width           =   1185
   End
   Begin VB.Image Image2 
      Height          =   735
      Index           =   1
      Left            =   0
      Picture         =   "frPrintInv.frx":0069
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11280
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFFF80&
      Height          =   315
      Index           =   0
      Left            =   9180
      Top             =   6180
      Width           =   1455
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFFF80&
      Height          =   315
      Index           =   2
      Left            =   9180
      Top             =   6900
      Width           =   1455
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFFF80&
      Height          =   315
      Index           =   3
      Left            =   5460
      Top             =   7620
      Width           =   1455
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFFF80&
      Height          =   315
      Index           =   4
      Left            =   9180
      Top             =   7260
      Width           =   1455
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFFF80&
      Height          =   315
      Index           =   5
      Left            =   9180
      Top             =   7620
      Width           =   1455
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFFF80&
      Height          =   315
      Index           =   6
      Left            =   9180
      Top             =   6540
      Width           =   1455
   End
End
Attribute VB_Name = "frmReceipt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim myPrintInv_ As BWGPrintInvHeader
Dim colCompany_ As Collection
Dim curCustomer_ As BWGCustomer
Dim isDetailChange_  As Boolean

Private Sub SetMSFlexGrid()
   
    With grCheq.getGridObj
        .Cols = 5
        .ColWidth(0) = 0
        .ColWidth(1) = 1845
        .ColWidth(2) = 1890
        .ColWidth(3) = 1350
        .ColWidth(4) = 1515
        .Rows = 50
        .FixedAlignment(1) = 4
        .FixedAlignment(2) = 4
        .FixedAlignment(3) = 4
        .FixedAlignment(4) = 4
        
        .ColAlignment(1) = 1
        .ColAlignment(2) = 1
        .ColAlignment(3) = 7
        .ColAlignment(4) = 4
    
        .TextMatrix(0, 1) = "เลขที่เช็ค"
        .TextMatrix(0, 2) = "ธนาคาร"
        .TextMatrix(0, 3) = "จำนวนเงิน"
        .TextMatrix(0, 4) = "วันที่ขึ้นเงิน"
    End With
    
    With grBill.getGridObj
        .Cols = 6
        .ColWidth(0) = 0
        .ColWidth(1) = 1455
        .ColWidth(2) = 5175
        .ColWidth(3) = 1410
        .ColWidth(4) = 1335
        .ColWidth(5) = 1335
        .Rows = 100
        .FixedAlignment(1) = 4
        .FixedAlignment(2) = 4
        .FixedAlignment(3) = 4
        .FixedAlignment(4) = 4
        .FixedAlignment(5) = 4
        
        .ColAlignment(1) = 1
        .ColAlignment(2) = 1
        .ColAlignment(3) = 7
        .ColAlignment(4) = 7
        .ColAlignment(5) = 7
    
        .TextMatrix(0, 1) = "ใบวางบิลที่เลือก"
        .TextMatrix(0, 2) = "คำอธิบายในใบกำกับ"
        .TextMatrix(0, 3) = "ยอดจากใบวางบิล"
        .TextMatrix(0, 4) = "ภาษีหัก ณ ที่จ่าย"
        .TextMatrix(0, 5) = "ยอดรับจริงงวดนี้"
    End With
End Sub

Private Sub cboBillAddr_GotFocus()
    cboBillAddr.Locked = Not CmdSave.Enabled
End Sub

Private Sub cboBillAddr_Validate(Cancel As Boolean)
Dim tmpBillAddr As BWGCustomerBillAddr
    With myPrintInv_
        If cboBillAddr.ListIndex = -1 Then
            .BillAddr = ""
            .BillAddrID = ""
            .BillAumphur = ""
            .BillProvince = ""
            .BillTumbol = ""
            .BillZipCode = ""
        Else
            Set tmpBillAddr = curCustomer_.BillAddresses(cboBillAddr.ListIndex + 1)
            .BillAddr = tmpBillAddr.Address
            .BillAddrID = tmpBillAddr.ID
            .BillAumphur = tmpBillAddr.AumphurName
            .BillProvince = tmpBillAddr.ProvinceName
            .BillTumbol = tmpBillAddr.TumbolName
            .BillZipCode = tmpBillAddr.ZipCode
        End If
    End With
End Sub

Private Sub cboCompany_GotFocus()
    cboCompany.Locked = Not CmdSave.Enabled
End Sub

Private Sub cboCompany_Validate(Cancel As Boolean)
    If cboCompany.ListIndex >= 0 Then
        If cboCompany.ListCount > 1 Then
            If cboCompany.ListIndex = 0 Then Exit Sub
            myPrintInv_.InvCompanyID = CurrentUser.WorkCompanies(cboCompany.ListIndex).ID
        Else
            myPrintInv_.InvCompanyID = CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).ID
        End If
    Else
        myPrintInv_.InvCompanyID = ""
    End If
    
    If myPrintInv_.InvCompanyID = "3F764F1713974948A42CC0FAF926C730" Then
        txtHeaderText.Text = "ค่าบริการ"
    Else
        txtHeaderText.Text = "ค่าบริการกำจัดและบำบัด"
    End If
End Sub

Private Sub CBReStaff_Validate(Cancel As Boolean)
    myPrintInv_.BankName = Replace(Left(CBReStaff.Text, 3), " ", "")
    myPrintInv_.BankNo = Replace(Mid(CBReStaff.Text, 6, 14), " ", "")
End Sub

Private Sub CBType_Click()
    If CBType.ListIndex = "1" Then
        Dim str As String
            If Trim(cboCompany.Text) = "บริษัท อัคคีปราการ จำกัด (มหาชน)" Then
                str = "AccountName like '%อัคคีปราการ%'"
                ElseIf Trim(cboCompany.Text) = "เบตเตอร์ เวิลด์ กรีน จำกัด (มหาชน)" Then
                str = "AccountName like '%เบตเตอร์ เวิลด์ กรีน%'"
                Else
                str = "AccountName like '%เบตเตอร์ เวสท์ แคร์%'"
            End If
            PopulateBankAccInfo CBReStaff, str
    Else
        PopulateStaffInfo CBReStaff
    End If

End Sub
Private Sub Check1_Click()
    If Check1.Value = 1 Then
        lbNetTotal = Format(CDbl(lbNetTotal) - CDbl(lbVat), "#,#00.00")
        lbVat.Caption = "0.00"
    Else
        lbVat.Caption = Format(CDbl(lbTotalPay) * SystemConfig.VATRate / 100)
        lbNetTotal = Format(CDbl(lbNetTotal) + CDbl(lbVat), "#,#00.00")
    End If
    myPrintInv_.isChanged = True
    myPrintInv_.IncVAT = Check1.Value
End Sub

Private Sub CheckTax_Click()
    If CheckTax.Value = 1 Then
        myPrintInv_.AdjTax = "Y"
    Else
        myPrintInv_.AdjTax = "N"
    End If
    SumMoney
End Sub

Private Sub cmdBillAdd_Click()
    If Trim(txtCustomer.Text) = "" Then
        MsgBox "กรุณาระบุชื่อลูกค้า", vbExclamation
        txtCustomer.SetFocus
        Exit Sub
    End If
    isDetailChange_ = False
    Load frmReceiptInvAdd
    With frmReceiptInvAdd
        .SetCallerForm Me, curCustomer_, myPrintInv_.ID
        .Show 1
        If isDetailChange_ Then
            ShowInvDetails
        End If
    End With
End Sub

Private Sub cmdCancel_Click()
Dim tmpStr As String
    If MsgBox("ต้องการยกเลิกใบกำกับ/ใบเสร็จ นี้หรือไม่?", vbQuestion + vbYesNo) = vbYes Then
        tmpStr = myPrintInv_.CancelBill
        If Trim(tmpStr) = "" Then
            SetEnabledScreen
        Else
            MsgBox tmpStr
        End If
    End If
End Sub

Private Sub cmdClear_Click()
Dim selResult As Integer
    If myPrintInv_.isChanged And CmdSave.Enabled Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนสร้างใหม่หรือไม่?", vbQuestion + vbYesNoCancel)
        If selResult = vbYes Then
            CmdSave_Click
            If myPrintInv_.isChanged Then Exit Sub
        ElseIf selResult = vbCancel Then
            Exit Sub
        End If
    End If
    
    ClearScreen
    CreateNewInv
    dtInvDate.SetFocus
End Sub

Private Sub ClearScreen()
    txtInvNo.Text = "== AUTO =="
    dtInvDate.ValueYMD = TodayDate
    txtCustomer.Text = ""
    Set curCustomer_ = Nothing
    txtRem.Text = ""
    txtPaidCash.Text = ""
    cboBillAddr.Clear
    grBill.ClearAllData
    grBill.getGridObj.Rows = 100
    grCheq.ClearAllData
    grCheq.getGridObj.Rows = 50
    lbTotalBill.Caption = "0.00"
    lbWTax.Caption = "0.00"
    lbTotalPay.Caption = "0.00"
    lbTotalCheq.Caption = "0.00"
    Check1.Value = 0
    CheckTax.Value = 0
End Sub

Private Sub cmdClose_Click()
    Unload frmReceiptInvAdd
    Unload frmReceiptSearch
    Unload Me
End Sub

Private Sub cmdPrintInv_Click()
Dim detailCount%, x%, tmpDet As BWGPrintInvDetail, tmpDet2 As BWGPrintRVDetail_Adjust
Dim curMfID$, tmpWeight As Double
Dim rowPerPage%, curRow%, tmpCol As Collection, tmpJobDet As BWGJobDataDetail
Dim tmpTrip As BWGJobDataTimeTable, curRecNo%, Y%, tmpWasteAdd As BWGMenifestWasteAdded
Dim tmpSite As BWGCustomerBillAddr, tmpCode$
Dim selResult As Integer
    
    If myPrintInv_.isChanged And CmdSave.Enabled Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนทำงานต่อหรือไม่?", vbQuestion + vbYesNoCancel)
        If selResult = vbYes Then
            CmdSave_Click
            If myPrintInv_.isChanged Then Exit Sub
        ElseIf selResult = vbCancel Then
            Exit Sub
        End If
    End If
    
    rowPerPage = 28
    detailCount = myPrintInv_.InvDetails.Count + myPrintInv_.InvDetails2.Count
    tmpCode = CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).CompanyCode
    With rpt1
        .LogonInfo(0) = "DSN = " & DSN_NAME & ";UID = " & strUid & ";PWD = " & strPwd
        .ReportFileName = App.Path & "\reports\CustReceipt_" & tmpCode & ".rpt"
'        .PrinterName = Printer.DeviceName
'        .PrinterDriver = Printer.DriverName
'        .PrinterPort = Printer.Port
        .Destination = crptToWindow
        
        curRow = 1
        ClearReceiptReport rpt1
        .Formulas(20 + (curRow - 1) * 4) = "Col1_" & curRow & "=" & Chr(34) & Trim(txtHeaderText.Text) & Chr(34)
        .Formulas(21 + (curRow - 1) * 4) = "Col2_" & curRow & "=" & Chr(34) & "" & Chr(34)
        .Formulas(22 + (curRow - 1) * 4) = "Col3_" & curRow & "=" & Chr(34) & "" & Chr(34)
        .Formulas(23 + (curRow - 1) * 4) = "Col4_" & curRow & "=" & Chr(34) & "" & Chr(34)
        curRow = curRow + 1
        Do Until curRecNo >= detailCount
            For x = 1 To myPrintInv_.InvDetails.Count
                Set tmpDet = myPrintInv_.InvDetails(x)
                .Formulas(20 + (curRow - 1) * 4) = "Col1_" & curRow & "=" & Chr(34) & "Invoice เลขที่ " & tmpDet.BillingNo & Chr(34)
                .Formulas(21 + (curRow - 1) * 4) = "Col2_" & curRow & "=" & Chr(34) & "" & Chr(34)
                .Formulas(22 + (curRow - 1) * 4) = "Col3_" & curRow & "=" & tmpDet.RealRecAmt
                .Formulas(23 + (curRow - 1) * 4) = "Col4_" & curRow & "=" & tmpDet.RealRecAmt
                curRow = curRow + 1
                curRecNo = curRecNo + 1
                Set tmpDet = Nothing
            Next x
            For x = 1 To myPrintInv_.InvDetails2.Count
                Set tmpDet2 = myPrintInv_.InvDetails2(x)
                '.Formulas(20 + (curRow - 1) * 4) = "Col1_" & curRow & "=" & Chr(34) & "ใบลดหนี้ เลขที่ " & tmpDet2.BillingNo & Chr(34)
                .Formulas(20 + (curRow - 1) * 4) = "Col1_" & curRow & "=" & Chr(34) & tmpDet2.InvBillDesc & Chr(34)
                .Formulas(21 + (curRow - 1) * 4) = "Col2_" & curRow & "=" & Chr(34) & "" & Chr(34)
                .Formulas(22 + (curRow - 1) * 4) = "Col3_" & curRow & "=" & tmpDet2.RealRecAmt
                .Formulas(23 + (curRow - 1) * 4) = "Col4_" & curRow & "=" & tmpDet2.RealRecAmt
                curRow = curRow + 1
                curRecNo = curRecNo + 1
                Set tmpDet2 = Nothing
            Next x
            
            
            If curRow >= rowPerPage Or curRecNo >= detailCount Then
                .Formulas(20 + (curRow - 1) * 4) = "Col1_" & curRow & "=" & Chr(34) & Trim(txtRem.Text) & Chr(34)
                
                .Formulas(0) = "InvNo=" & Chr(34) & Trim(txtInvNo.Text) & Chr(34)
                .Formulas(1) = "InvDate=" & Chr(34) & Trim(dtInvDate.ValueDMY) & Chr(34)
                '.Formulas(2) = "CustName=" & Chr(34) & Trim(txtCustomer.Text) & Chr(34)
                '.Formulas(3) = "CustAddr=" & Chr(34) & Trim(cboBillAddr.Text) & Chr(34)
                If cboBillAddr.ListIndex > -1 Then
                    Set tmpSite = curCustomer_.BillAddresses(cboBillAddr.ListIndex + 1)
                    .Formulas(2) = "CustName=" & Chr(34) & Trim(tmpSite.BillName) & Chr(34)
                    
                    If InStr(1, tmpSite.ProvinceName, "กรุงเทพ") > 0 Then
                        .Formulas(3) = "CustAddr=" & Chr(34) & tmpSite.Address & " แขวง " & tmpSite.TumbolName & " เขต " & tmpSite.AumphurName & " " & tmpSite.ProvinceName & " " & tmpSite.ZipCode & Chr(34)
                    Else
                        If InStr(1, tmpSite.AumphurName, "กิ่ง อ") > 0 Then
                            .Formulas(3) = "CustAddr=" & Chr(34) & tmpSite.Address & " ต. " & tmpSite.TumbolName & " " & tmpSite.AumphurName & " จ. " & tmpSite.ProvinceName & " " & tmpSite.ZipCode & Chr(34)
                        Else
                            .Formulas(3) = "CustAddr=" & Chr(34) & tmpSite.Address & " ต. " & tmpSite.TumbolName & " อ. " & tmpSite.AumphurName & " จ. " & tmpSite.ProvinceName & " " & tmpSite.ZipCode & Chr(34)
                        End If
                    End If
                        .Formulas(99) = "CustTax=" & Chr(34) & "เลขประจำตัวผู้เสียภาษีลูกค้า " & tmpSite.TaxNum & Chr(34)
                    Set tmpSite = Nothing
                        .Formulas(9) = "BankName=" & Chr(34) & myPrintInv_.BankName & Chr(34)
                        .Formulas(10) = "CheqNo=" & Chr(34) & myPrintInv_.BankNo & Chr(34)
                End If
                If IsNumeric(lbTotalBill.Caption) Then
                    .Formulas(4) = "TotalAmt=" & Format(Trim(lbTotalPay.Caption), "###.00")
                    If Check1.Value = 0 Then
                        .Formulas(5) = "VATAmt=" & Format(CDbl(lbTotalPay.Caption) * SystemConfig.VATRate / 100, "###.00")
                    Else
                        .Formulas(5) = "VATAmt=0.00"
                    End If
                    
                    If Check1.Value = 0 Then
                        If IsNumeric(Trim(lbWTax.Caption)) Then
                            .Formulas(6) = "GrandTotalAmt=" & CDbl(Format(Trim(lbTotalPay.Caption), "###.00")) - CDbl(Format(Trim(lbWTax.Caption), "###.00")) + CDbl(Format(CDbl(lbTotalPay.Caption) * SystemConfig.VATRate / 100, "###.00"))
                            .Formulas(7) = "AMTWord=" & Chr(34) & Trim(AmountToString(CDbl(Format(Trim(lbTotalPay.Caption), "###.00")) - CDbl(Format(Trim(lbWTax.Caption), "###.00")) + CDbl(Format(CDbl(lbTotalPay.Caption) * SystemConfig.VATRate / 100, "###.00")))) & Chr(34)
                        Else
                            .Formulas(6) = "GrandTotalAmt=" & CDbl(Format(Trim(lbTotalPay.Caption), "###.00")) + CDbl(Format(CDbl(lbTotalPay.Caption) * SystemConfig.VATRate / 100, "###.00"))
                            .Formulas(7) = "AMTWord=" & Chr(34) & Trim(AmountToString(CDbl(Format(Trim(lbTotalPay.Caption), "###.00")) + CDbl(Format(CDbl(lbTotalPay.Caption) * SystemConfig.VATRate / 100, "###.00")))) & Chr(34)
                        End If
                    Else
                        If IsNumeric(Trim(lbWTax.Caption)) Then
                            .Formulas(6) = "GrandTotalAmt=" & CDbl(Format(Trim(lbTotalPay.Caption), "###.00")) - CDbl(Format(Trim(lbWTax.Caption), "###.00"))
                            .Formulas(7) = "AMTWord=" & Chr(34) & Trim(AmountToString(CDbl(Format(Trim(lbTotalPay.Caption), "###.00")) - CDbl(Format(Trim(lbWTax.Caption), "###.00")))) & Chr(34)
                        Else
                            .Formulas(6) = "GrandTotalAmt=" & CDbl(Format(Trim(lbTotalBill.Caption), "###.00"))
                            .Formulas(7) = "AMTWord=" & Chr(34) & Trim(AmountToString(CDbl(Format(Trim(lbTotalPay.Caption), "###.00")))) & Chr(34)
                        End If
                    End If
                    
                    .Formulas(8) = "WTaxAmt=" & Chr(34) & Trim(lbWTax.Caption) & Chr(34)
                    Dim Tot
                    Tot = Replace(.Formulas(6), "GrandTotalAmt=", "")
                    If CBType.ListIndex = "1" Then
                            .Formulas(11) = "INVDate2=" & Chr(34) & Trim(dtInvDate.ValueDMY) & Chr(34)
                            .Formulas(12) = "TotAmt=" & Tot
                        Else
                            .Formulas(9) = "BankName=" & ""
                            .Formulas(10) = "CheqNo=" & ""
                            .Formulas(11) = "INVDate2=" & ""
                            .Formulas(12) = "TotAmt=" & ""
                    End If

                End If
                '.Formulas(8) = "DueDay=" & Chr(34) & Trim(dtDueDate.DayValue) & Chr(34)
                '.Formulas(9) = "DueMonth=" & Chr(34) & Trim(dtDueDate.MonthValue) & Chr(34)
                '.Formulas(10) = "DueYear=" & Chr(34) & Trim(dtDueDate.YearValue) & Chr(34)
                .PrintReport
                If Trim(.LastErrorString) <> "" Then MsgBox .LastErrorString
                ClearReceiptReport rpt1
                curRow = 1
            End If
        Loop
    End With
    
    If myPrintInv_.isPrintAfterChanged = False Then
        DBConnExc "Update tbPrintInvHeader Set isPrintAfterChanged='Y' WHERE PrintInvID='" & myPrintInv_.ID & "'"
        myPrintInv_.isPrintAfterChanged = True
        myPrintInv_.isChanged = False
        SetEnabledScreen
    End If
End Sub

Private Sub CmdSave_Click()
Dim x%
Dim tmpDet As BWGLABTemplateDetail

    If cboCompany.ListIndex < 0 Then
        MsgBox "กรุณาระบุบริษัทที่เป็นผู้ออกใบกำกับ/ใบเสร็จ", vbExclamation
        cboCompany.SetFocus
        Exit Sub
    End If
    If Trim(txtCustomer.Text) = "" Then
        MsgBox "กรุณาระบุลูกค้าที่จะทำการเก็บเงิน", vbExclamation
        txtCustomer.SetFocus
        Exit Sub
    End If
    If cboBillAddr.ListIndex = -1 Then
        MsgBox "กรุณาระบุที่อยู่ในการออกใบกำกับ/ใบเสร็จ", vbExclamation
        cboBillAddr.SetFocus
        Exit Sub
    End If
    If (myPrintInv_.InvDetails.Count = 0) And (myPrintInv_.InvDetails2.Count = 0) Then
        MsgBox "กรุณาเลือกใบวางบิลที่ต้องการออกใบกำกับ/ใบเสร็จ", vbExclamation
        Exit Sub
    End If
    
    myPrintInv_.IncVAT = Check1.Value
    myPrintInv_.NetValue = lbTotalPay
    If myPrintInv_.isChanged And CmdSave.Enabled Then
        myPrintInv_.Save
        txtInvNo.Text = myPrintInv_.InvNo
    End If
    If myPrintInv_.isChanged = False Then MsgBox "จัดเก็บข้อมูลเรียบร้อยแล้ว", vbExclamation
    
End Sub

Private Sub cmdSearch_Click()
Dim selResult As Integer
    
    If myPrintInv_.isChanged And CmdSave.Enabled Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนทำงานต่อหรือไม่?", vbQuestion + vbYesNoCancel)
        If selResult = vbYes Then
            CmdSave_Click
            If myPrintInv_.isChanged Then Exit Sub
        ElseIf selResult = vbCancel Then
            Exit Sub
        Else
            Set myPrintInv_.InvDetails = Nothing
            Set myPrintInv_.InvDetails2 = Nothing
            myPrintInv_.isChanged = False
            cmdClear_Click
        End If
    End If
    Load frmReceiptSearch
    With frmReceiptSearch
        .SetCallerForm Me
        .Show 1
    End With
End Sub

Private Sub dtInvDate_Validate(Cancel As Boolean)
    myPrintInv_.InvDate = Trim(dtInvDate.ValueYMD)
End Sub

Private Sub Form_Load()
Dim x%, selX As Integer
Dim str As String
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
    CBType.ListIndex = 0
    PopulateCompany cboCompany, cboIndex
    PopulateStaffInfo CBReStaff
    Call SetMSFlexGrid
    With frmPicture
        CmdSave.Picture = .picSave.Picture
        cmdSearch.Picture = .picSearch.Picture
        cmdClear.Picture = .picNew.Picture
        cmdCancel.Picture = .picDelete.Picture
        cmdClose.Picture = .picExit.Picture
    End With
     CreateNewInv
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 And Me.ActiveControl.Name <> "grBill" And Me.ActiveControl.Name <> "grCheq" Then SendKeys "{Tab}"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub CreateNewInv()
Set myPrintInv_ = Nothing
Set myPrintInv_ = New BWGPrintInvHeader
    With myPrintInv_
        .ID = GetGUID
        .InvDate = dtInvDate.ValueYMD
        If cboCompany.ListCount > 1 Then
            If cboCompany.ListIndex > 0 Then
                .InvCompanyID = CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).ID
            End If
        Else
            If cboCompany.ListIndex >= 0 Then
                .InvCompanyID = CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).ID
            End If
        End If
        If Not curCustomer_ Is Nothing Then
            .CustomerID = curCustomer_.ID
            .CustomerName = curCustomer_.CustomerName
        End If
        .isChanged = False
    End With
    SetEnabledScreen
End Sub

Private Sub SetEnabledScreen()
Dim eFlag As Boolean
Dim UserGrp As Collection
Dim Group As New AWSUserGroup
Dim StrGrp As String
Dim i As Integer

    eFlag = Not (myPrintInv_.isCancel)
'    eFlag = Not (myPrintInv_.isCancel Or myPrintInv_.isPrintAfterChanged)
    
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
    lbCancel.Visible = False
    If StrGrp = "Administrator" Then
        CmdSave.Enabled = True
        cmdCancel.Enabled = True
        cmdBillAdd.Enabled = True
        dtInvDate.Enabled = True
    Else
        CmdSave.Enabled = eFlag
        cmdCancel.Enabled = eFlag
        cmdBillAdd.Enabled = eFlag
        lbCancel.Visible = Not eFlag
        dtInvDate.Enabled = eFlag
    End If
    If myPrintInv_.isCancel Then
        lbCancel.Caption = "ยกเลิก"
        lbCancel.Visible = True
'        cmdSave.Enabled = False
'        cmdCancel.Enabled = False
'        cmdBillAdd.Enabled = False
'        dtInvDate.Enabled = False
    ElseIf myPrintInv_.isPrintAfterChanged Then
        lbCancel.Caption = "Printed"
        lbCancel.Visible = True
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim selResult As Integer
    If myPrintInv_ Is Nothing Then Exit Sub
    If myPrintInv_.isChanged And CmdSave.Enabled Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนปิดหน้าจอหรือไม่?", vbQuestion + vbYesNoCancel)
        If selResult = vbYes Then
            CmdSave_Click
            If myPrintInv_.isChanged Then Cancel = 1
        ElseIf selResult = vbCancel Then
            Cancel = 1
        End If
    End If
End Sub



Private Sub grBill_BeforeUpdateValue(NewValue As String, row As Integer, col As Integer, Cancel As Boolean)
    If col = 5 And IsNumeric(NewValue) = False Then
        MsgBox "กรุณาระบุยอดรับจริงเป็นตัวเลขเท่านั้น", vbExclamation
        Cancel = True
    End If
End Sub

Private Sub grBill_Click(row As Integer, col As Integer)
'    MsgBox grBill.getGridObj.ColWidth(col)
End Sub

Private Sub grBill_RolColChanged(row As Integer, col As Integer)
    If CmdSave.Enabled = False Then
        grBill.ColEnabled(col) = False
    Else
        If col = 1 Or col = 3 Or col = 4 Then
            grBill.ColEnabled(col) = False
        ElseIf (col = 2 Or col = 5) And Trim(grBill.getGridObj.TextMatrix(row, 0)) <> "" Then
            grBill.ColEnabled(col) = True
        End If
    End If
End Sub

Private Sub grBill_UpdateValue(row As Integer, col As Integer)
Dim tmpDet As BWGPrintInvDetail
Dim tmpKey As String, tmpVal As String
    With grBill.getGridObj
        tmpKey = "'" & Trim(.TextMatrix(row, 0)) & "'"
        tmpVal = Trim(.TextMatrix(row, col))
        If isExist(myPrintInv_.InvDetails, tmpKey) And (col = 2 Or col = 5) Then
            Set tmpDet = myPrintInv_.InvDetails(tmpKey)
            Select Case col
                Case 2
                    tmpDet.InvBillDesc = Trim(tmpVal)
                Case 5
                    If IsNumeric(tmpVal) Then
                        grBill.getGridObj.TextMatrix(row, col) = Format(grBill.getGridObj.TextMatrix(row, col), "#,##0.00")
                        tmpDet.RealRecAmt = CDbl(tmpVal)
                    Else
                        tmpDet.RealRecAmt = 0
                    End If
                    SumMoney
            End Select
            myPrintInv_.isChanged = True
            Set tmpDet = Nothing
        End If
    End With
End Sub

Private Sub grCheq_BeforeUpdateValue(NewValue As String, row As Integer, col As Integer, Cancel As Boolean)
    If col = 3 And IsNumeric(NewValue) = False Then
        MsgBox "กรุณาระบุจำนวนเงินในเช็คเป็นตัวเลข", vbExclamation
        Cancel = True
    End If
End Sub

Private Sub grCheq_Click(row As Integer, col As Integer)
'    MsgBox grCheq.getGridObj.ColWidth(col)
End Sub

Private Sub grCheq_RolColChanged(row As Integer, col As Integer)
    If CmdSave.Enabled Then
        If col = 1 And Trim(grCheq.getGridObj.TextMatrix(row - 1, col)) <> "" Then
            grCheq.ColEnabled(col) = True
        ElseIf col > 1 Then
            If Trim(grCheq.getGridObj.TextMatrix(row, 0)) <> "" Then
                grCheq.ColEnabled(col) = True
            Else
                grCheq.ColEnabled(col) = False
            End If
        End If
    Else
        grCheq.ColEnabled(col) = False
    End If
End Sub

Private Sub grCheq_UpdateValue(row As Integer, col As Integer)
Dim tmpDet As BWGInvCheqPayment
Dim tmpKey As String, tmpVal As String
    tmpKey = "'" & Trim(grCheq.getGridObj.TextMatrix(row, 0)) & "'"
    tmpVal = Trim(grCheq.getGridObj.TextMatrix(row, col))
    If isExist(myPrintInv_.CheqPayments, tmpKey) Then
        Set tmpDet = myPrintInv_.CheqPayments(tmpKey)
    Else
        Set tmpDet = New BWGInvCheqPayment
        tmpDet.ID = GetGUID
        tmpDet.PrintInvID = myPrintInv_.ID
        grCheq.getGridObj.TextMatrix(row, 0) = tmpDet.ID
        myPrintInv_.CheqPayments.Add tmpDet, "'" & tmpDet.ID & "'"
    End If
    
    Select Case col
        Case 1
            tmpDet.CheqNo = Trim(tmpVal)
        Case 2
            tmpDet.BankName = Trim(tmpVal)
        Case 3
            If IsNumeric(tmpVal) Then
                tmpDet.CheqAmt = CDbl(tmpVal)
            Else
                tmpDet.CheqAmt = 0
            End If
            SumMoney
        Case 4
            tmpDet.cheqDate = tmpVal
    End Select
    myPrintInv_.isChanged = True
    Set tmpDet = Nothing
End Sub

Private Sub lbTotalPay_Change()
    If IsNumeric(lbTotalPay.Caption) Then
        lbVat.Caption = Format(CDbl(lbTotalPay.Caption) * SystemConfig.VATRate / 100, "#,##0.00")
        lbNetTotal.Caption = Format(CDbl(lbTotalPay.Caption) + CDbl(lbVat.Caption) - CDbl(lbWTax.Caption), "#,##0.00")
    Else
        lbVat.Caption = "0.00"
        lbWTax.Caption = "0.00"
        lbNetTotal.Caption = lbTotalPay.Caption
    End If
End Sub

Private Sub txtCustomer_GotFocus()
    txtCustomer.Locked = Not CmdSave.Enabled
End Sub

Private Sub txtCustomer_Validate(Cancel As Boolean)
    If curCustomer_ Is Nothing Then Set curCustomer_ = New BWGCustomer
    If Trim(txtCustomer.Text) = "" And Trim(curCustomer_.CustomerName) <> Trim(txtCustomer.Text) Then
        Set curCustomer_ = Nothing
        Set curCustomer_ = New BWGCustomer
        With myPrintInv_
            .CustomerID = ""
            .CustomerName = ""
        End With
    Else
        If Trim(curCustomer_.CustomerName) <> Trim(txtCustomer.Text) Then
            Set curCustomer_ = Nothing
            Set curCustomer_ = New BWGCustomer
            With myPrintInv_
                .CustomerID = ""
                .CustomerName = ""
            End With
            Load frmCustomerSearch
            With frmCustomerSearch
                .CStatus = "N"
                .getCriteria (txtCustomer.Text)
                .GetCriteriaOptonal ""
'                .GetCriteriaOptonal " AND CustGroupID <> '488F2E5C057F46869D1A56185B11D644'"
                .SetCallerForm Me
                .Show 1
            End With
            If myPrintInv_.CustomerID = "" Then txtCustomer.Text = ""
        End If
    End If
End Sub

Public Sub SetDataCustomer(selCustomer As BWGCustomer, Optional isAssigned As Boolean = True)
Dim tmpRec As New ADODB.Recordset
Dim x%
Dim tmpWaste As BWGWasteDataCR
    Set curCustomer_ = Nothing
    Set curCustomer_ = selCustomer
    If curCustomer_ Is Nothing Then Exit Sub
    cboBillAddr.Clear
    
    With curCustomer_
        For x = 1 To .BillAddresses.Count
            cboBillAddr.AddItem .BillAddresses(x).Address & " " & .BillAddresses(x).TumbolName & " " & .BillAddresses(x).AumphurName & " " & .BillAddresses(x).ProvinceName
        Next
        txtCustomer.Text = .CustomerName
        If isAssigned Then
            myPrintInv_.CustomerID = .ID
            myPrintInv_.CustomerName = .CustomerName
            If cboBillAddr.ListCount = 1 Then
                cboBillAddr.ListIndex = 0
                myPrintInv_.BillAddr = .BillAddresses(1).Address
                myPrintInv_.BillAddrID = .BillAddresses(1).ID
                myPrintInv_.BillAumphur = .BillAddresses(1).AumphurName
                myPrintInv_.BillProvince = .BillAddresses(1).ProvinceName
                myPrintInv_.BillTumbol = .BillAddresses(1).TumbolName
                myPrintInv_.BillZipCode = .BillAddresses(1).ZipCode
            End If
        End If
    End With
End Sub

Private Sub txtPaidCash_Validate(Cancel As Boolean)
    If IsNumeric(txtPaidCash.Text) Then
        myPrintInv_.PaidCashAmt = txtPaidCash.Text
    Else
        myPrintInv_.PaidCashAmt = 0
    End If
End Sub

Private Sub txtRem_GotFocus()
    txtRem.Locked = Not CmdSave.Enabled
End Sub

Private Sub txtRem_Validate(Cancel As Boolean)
    myPrintInv_.InvRemark = Trim(txtRem.Text)
End Sub

Private Sub ShowInvDetails()
Dim x%, Y%
Dim tmpDet As BWGPrintInvDetail
Dim tmpDet2 As BWGPrintRVDetail_Adjust
Dim tmpTax As Integer
    grBill.ClearAllData
    With grBill.getGridObj
        .Rows = myPrintInv_.InvDetails.Count + myPrintInv_.InvDetails2.Count + 50
        For x = 1 To myPrintInv_.InvDetails.Count
            Set tmpDet = myPrintInv_.InvDetails(x)
            .TextMatrix(x, 0) = tmpDet.ID
            .TextMatrix(x, 1) = tmpDet.BillingNo
            .TextMatrix(x, 2) = tmpDet.InvBillDesc
            .TextMatrix(x, 3) = Format(tmpDet.BillingTotalAmt, "#,##0.00")
            .TextMatrix(x, 4) = Format(tmpDet.BillingTotalAmt * tmpDet.WTaxRate / 100, "#,##0.00")
            .TextMatrix(x, 5) = Format(tmpDet.RealRecAmt, "#,##0.00")
            If tmpDet.WTaxRate > 0 Then
                If tmpTax = 0 Then tmpTax = tmpDet.WTaxRate
            End If
            Set tmpDet = Nothing
        Next
        For Y = 1 To myPrintInv_.InvDetails2.Count
            Set tmpDet2 = myPrintInv_.InvDetails2(Y)
            .TextMatrix(x + Y, 0) = tmpDet2.ID
            .TextMatrix(x + Y, 1) = tmpDet2.BillingNo
            .TextMatrix(x + Y, 2) = tmpDet2.InvBillDesc
            If tmpDet2.AdjType = "A" Then
                .TextMatrix(x + Y, 3) = Format(tmpDet2.BillingTotalAmt, "#,##0.00")
                If CheckTax.Value = 1 Then
                    .TextMatrix(x + Y, 4) = Format(tmpDet2.BillingTotalAmt * tmpTax / 100, "#,##0.00")
                Else
                    .TextMatrix(x + Y, 4) = Format(tmpDet2.BillingTotalAmt * tmpDet2.WTaxRate / 100, "#,##0.00")
                End If
            Else
                .TextMatrix(x + Y, 3) = Format(tmpDet2.RealRecAmt, "#,##0.00")
                .TextMatrix(x + Y, 4) = Format(tmpDet2.RealRecAmt * tmpDet2.WTaxRate / 100, "#,##0.00")
            End If
            .TextMatrix(x + Y, 5) = Format(tmpDet2.RealRecAmt, "#,##0.00")
            Set tmpDet = Nothing
        Next
    End With
    SumMoney
End Sub

Public Sub RemoveBillData(selID As String)
Dim tmpCol As Collection
Dim x%
Dim tmpKey As String
    Set tmpCol = myPrintInv_.InvDetails
    If tmpCol Is Nothing Then Exit Sub
    For x = 1 To tmpCol.Count
        If tmpCol(x).BillingHeaderID = selID Then
            tmpKey = "'" & tmpCol(x).ID & "'"
            Exit For
        End If
    Next
    If isExist(tmpCol, tmpKey) Then
        tmpCol.Remove tmpKey
        isDetailChange_ = True
        myPrintInv_.isChanged = True
    End If
    
    Set tmpCol = myPrintInv_.InvDetails2
    If tmpCol Is Nothing Then Exit Sub
    For x = 1 To tmpCol.Count
        If tmpCol(x).AdjustID = selID Then
            tmpKey = "'" & tmpCol(x).ID & "'"
            Exit For
        End If
    Next
    If isExist(tmpCol, tmpKey) Then
        tmpCol.Remove tmpKey
        isDetailChange_ = True
        myPrintInv_.isChanged = True
    End If
End Sub

Public Sub AddBillItem(selItem As BWGBillingHeader)
Dim tmpCol As Collection
Dim x%, isNew As Boolean
Dim tmpID As String
Dim tmpDetail As New BWGPrintInvDetail
    Set tmpCol = myPrintInv_.InvDetails

    If tmpCol Is Nothing Then Set tmpCol = New Collection
    isNew = True
    tmpID = Trim(selItem.ID)
    For x = 1 To tmpCol.Count
        'If tmpCol(x).selitem.ID = tmpID Then
        If tmpCol(x).BillingHeaderID = tmpID Then
            isNew = False
            Exit For
        End If
    Next

    If isNew Then
        tmpDetail.ID = GetGUID
        tmpDetail.BillingDate = selItem.IssuedDate
        tmpDetail.BillingHeaderID = selItem.ID
        tmpDetail.BillingNo = selItem.BillingNo
        tmpDetail.BillingPaidAmt = selItem.PaidAmt
        tmpDetail.BillingTotalAmt = selItem.AmtB4VAT
        tmpDetail.InvBillDesc = "ชำระเงินตามใบวางบิลเลขที่ " & selItem.BillingNo
        tmpDetail.PrintInvID = myPrintInv_.ID
        tmpDetail.RealRecAmt = selItem.AmtB4VAT - selItem.PaidAmt
        tmpDetail.RowID = tmpCol.Count + 1
        tmpDetail.WTaxRate = selItem.WTaxRate
        tmpCol.Add tmpDetail, "'" & tmpDetail.ID & "'"
        isDetailChange_ = True
        myPrintInv_.isChanged = True
        Set tmpDetail = Nothing
    End If
    If myPrintInv_.InvDetails Is Nothing Then
        Set myPrintInv_.InvDetails = tmpCol
        isDetailChange_ = True
        myPrintInv_.isChanged = True
    End If
    Set tmpCol = Nothing
End Sub

Public Sub AddBillItem2(selItem As BWGAdjDeptHeader)
Dim tmpCol As Collection
Dim x%, isNew As Boolean
Dim tmpID As String
Dim tmpDetail As New BWGPrintRVDetail_Adjust
    Set tmpCol = myPrintInv_.InvDetails2

    If tmpCol Is Nothing Then Set tmpCol = New Collection
    isNew = True
    tmpID = Trim(selItem.ID)
    For x = 1 To tmpCol.Count
        If tmpCol(x).AdjustID = tmpID Then
            isNew = False
            Exit For
        End If
    Next

    If isNew Then
        tmpDetail.ID = GetGUID
        tmpDetail.BillingDate = selItem.IssuedDate
        tmpDetail.AdjustID = selItem.ID
        tmpDetail.BillingNo = selItem.AdjustBillNo
        tmpDetail.BillingPaidAmt = 0
        tmpDetail.AdjType = selItem.AdjustType
        If tmpDetail.AdjType = "A" Then
            tmpDetail.BillingTotalAmt = selItem.OldBillAmt - selItem.CorrectBillAmt
            tmpDetail.InvBillDesc = "ลดหนี้ ตามใบลดหนี้/ไม่ใช่ใบกำกับภาษี เลขที่ " & selItem.AdjustBillNo
            tmpDetail.RealRecAmt = selItem.OldBillAmt - selItem.CorrectBillAmt
        Else
            tmpDetail.BillingTotalAmt = selItem.CorrectBillAmt - selItem.OldBillAmt
            tmpDetail.InvBillDesc = "เพิ่มหนี้ ตามใบเพิ่มหนี้/ไม่ใช่ใบกำกับภาษี เลขที่ " & selItem.AdjustBillNo
            tmpDetail.RealRecAmt = selItem.CorrectBillAmt - selItem.OldBillAmt
        End If
        tmpDetail.PrintInvID = myPrintInv_.ID
        tmpDetail.RowID = tmpCol.Count + 1
        tmpDetail.WTaxRate = 0
        tmpCol.Add tmpDetail, "'" & tmpDetail.ID & "'"
        isDetailChange_ = True
        myPrintInv_.isChanged = True
        Set tmpDetail = Nothing
    End If
    If myPrintInv_.InvDetails2 Is Nothing Then
        Set myPrintInv_.InvDetails2 = tmpCol
        isDetailChange_ = True
        myPrintInv_.isChanged = True
    End If
    Set tmpCol = Nothing
End Sub

Public Function getBillItems() As Collection
Dim x%
Dim tmpCol As New Collection
Dim tmpDet As BWGPrintInvDetail
Dim tmpDet2 As BWGPrintRVDetail_Adjust
Dim tmpBill As BWGBillingHeader
Dim tmpBill2 As BWGAdjDeptHeader
Dim selCol As Collection, tmpStr As String, tmpArr
    Set selCol = myPrintInv_.InvDetails
    With selCol
        For x = 1 To .Count
            Set tmpDet = .item(x)
            If Trim(tmpDet.BillingHeaderID) <> "" Then
                Set tmpBill = New BWGBillingHeader
                tmpBill.IssuedDate = tmpDet.BillingDate
                tmpBill.ID = tmpDet.BillingHeaderID
                tmpBill.BillingNo = tmpDet.BillingNo
                tmpBill.PaidAmt = tmpDet.BillingPaidAmt
                tmpBill.WTaxRate = tmpDet.WTaxRate
                tmpBill.AmtB4VAT = tmpDet.BillingTotalAmt
                tmpCol.Add tmpBill, "'" & tmpDet.BillingHeaderID & "'"
                Set tmpBill = Nothing
            End If
            Set tmpDet = Nothing
        Next
    End With
    
    Set selCol = myPrintInv_.InvDetails2
    For x = 1 To selCol.Count
        Set tmpDet2 = selCol.item(x)
        If Trim(tmpDet2.AdjustID) <> "" Then
            Set tmpBill2 = New BWGAdjDeptHeader
            tmpBill2.IssuedDate = tmpDet2.BillingDate
            tmpBill2.ID = tmpDet2.AdjustID
            tmpBill2.AdjustType = tmpDet2.AdjType
            tmpBill2.AdjustBillNo = tmpDet2.BillingNo
            tmpBill2.CorrectBillAmt = tmpDet2.BillingTotalAmt
            tmpCol.Add tmpBill2, "'" & tmpDet2.AdjustID & "'"
            
            Set tmpBill = Nothing
        End If
        Set tmpDet = Nothing
    Next

    Set getBillItems = tmpCol
    Set tmpCol = Nothing
End Function

Private Sub SumMoney()
Dim x%, Y%, totalAmt As Double, realPayAmt As Double, CheqAmt As Double
Dim tmpDet As BWGPrintInvDetail, tmpCheq As BWGInvCheqPayment
Dim tmpDet2 As BWGPrintRVDetail_Adjust
Dim WTaxAmt As Double, AdjTax As Double
Dim tmpTax As Integer
    For x = 1 To myPrintInv_.InvDetails.Count
        Set tmpDet = myPrintInv_.InvDetails(x)
        totalAmt = totalAmt + tmpDet.BillingTotalAmt
        realPayAmt = realPayAmt + tmpDet.RealRecAmt
        WTaxAmt = WTaxAmt + CDbl(Format(tmpDet.BillingTotalAmt * tmpDet.WTaxRate / 100, "0.00"))
        If tmpDet.WTaxRate > 0 Then
            If tmpTax = 0 Then tmpTax = tmpDet.WTaxRate
        End If
    Next
    
    For Y = 1 To myPrintInv_.InvDetails2.Count
        Set tmpDet2 = myPrintInv_.InvDetails2(Y)
        If tmpDet2.AdjType = "A" Then
            totalAmt = totalAmt - tmpDet2.BillingTotalAmt
            realPayAmt = realPayAmt - tmpDet2.RealRecAmt
            'WTaxAmt = WTaxAmt - CDbl(Format(tmpDet2.BillingTotalAmt * 0 / 100, "0.00"))
        Else
            If tmpDet2.AdjType = "A" Then
                totalAmt = totalAmt + tmpDet2.BillingTotalAmt
                realPayAmt = realPayAmt + tmpDet2.RealRecAmt
            Else
                totalAmt = totalAmt + tmpDet2.RealRecAmt
                realPayAmt = realPayAmt + tmpDet2.RealRecAmt
            End If
        End If
        If tmpDet2.AdjType = "A" Then
            If CheckTax.Value = 1 Then
                WTaxAmt = WTaxAmt - CDbl(Format(tmpDet2.RealRecAmt * tmpTax / 100, "0.00"))
                AdjTax = tmpDet2.RealRecAmt * tmpTax / 100
            Else
                WTaxAmt = WTaxAmt - CDbl(Format(tmpDet2.RealRecAmt * tmpDet2.WTaxRate / 100, "0.00"))
                AdjTax = tmpDet2.RealRecAmt * tmpDet2.WTaxRate / 100
            End If
        Else
            WTaxAmt = WTaxAmt - CDbl(Format(tmpDet2.RealRecAmt * tmpDet2.WTaxRate / 100, "0.00"))
            AdjTax = tmpDet2.RealRecAmt * tmpDet2.WTaxRate / 100
        End If
        grBill.getGridObj.TextMatrix(x + Y, 4) = Format(AdjTax, "#,##0.00")
    Next
    
    
    For x = 1 To myPrintInv_.CheqPayments.Count
        Set tmpCheq = myPrintInv_.CheqPayments(x)
        CheqAmt = CheqAmt + tmpCheq.CheqAmt
    Next
    lbWTax.Caption = Format(WTaxAmt, "#,##0.00")
    lbTotalBill.Caption = Format(totalAmt, "#,##0.00")
    If realPayAmt <> lbTotalPay Then
        lbTotalPay.Caption = Format(realPayAmt, "#,##0.00")
        myPrintInv_.NetValue = lbTotalPay
    End If
    lbTotalCheq.Caption = Format(CheqAmt, "#,##0.00")
End Sub

Private Sub ShowCheqDetail()
Dim x%
Dim tmpDet As BWGInvCheqPayment
    grCheq.ClearAllData
    With grCheq.getGridObj
        .Rows = myPrintInv_.CheqPayments.Count + 50
        For x = 1 To myPrintInv_.CheqPayments.Count
            Set tmpDet = myPrintInv_.CheqPayments(x)
            .TextMatrix(x, 0) = tmpDet.ID
            .TextMatrix(x, 1) = tmpDet.CheqNo
            .TextMatrix(x, 2) = tmpDet.BankName
            .TextMatrix(x, 3) = Format(tmpDet.CheqAmt, "#,##0.00")
            .TextMatrix(x, 4) = FormatYMD_to_DMY(tmpDet.cheqDate, "/", "/")
            Set tmpDet = Nothing
        Next
    End With
    SumMoney
End Sub

Public Sub SetDataPrintInv(selInv As BWGPrintInvHeader)
Dim tmpS As New BWGSearchManager
Dim tmpCol As Collection
    Set myPrintInv_ = Nothing
    Set myPrintInv_ = selInv
    With myPrintInv_
        txtInvNo.Text = .InvNo
        dtInvDate.ValueYMD = .InvDate
        If isExist(CurrentUser.WorkCompanies, "'" & .InvCompanyID & "'") Then
            SetTextToCombo CurrentUser.WorkCompanies("'" & .InvCompanyID & "'").CompanyNameTH, cboCompany
            cboCompany_Validate False
        End If
        Set tmpCol = tmpS.CustomerSearch("CustomerID='" & .CustomerID & "'")
        If tmpCol.Count > 0 Then SetDataCustomer tmpCol(1), False
        Set tmpCol = Nothing
        Set tmpS = Nothing
        txtRem.Text = .InvRemark
        txtPaidCash.Text = .PaidCashAmt
        'SetTextToCombo .BillAddr & " " & .BillTumbol & " " & .BillAumphur & " " & .BillProvince & " " & .BillZipCode, cboBillAddr
        SetTextToCombo .BillAddr & " " & .BillTumbol & " " & .BillAumphur & " " & .BillProvince, cboBillAddr
        If .AdjTax = "Y" Then
            CheckTax.Value = 1
        Else
            CheckTax.Value = 0
        End If
        lbTotalPay = .NetValue
    End With
    Check1.Value = myPrintInv_.IncVAT
    ShowInvDetails
    ShowCheqDetail
    
    If Check1.Value = 1 Then
        lbNetTotal = Format(CDbl(lbNetTotal) - CDbl(lbVat), "#,#00.00")
        lbVat.Caption = "0.00"
    Else
        lbVat.Caption = Format(CDbl(lbTotalPay) * SystemConfig.VATRate / 100)
        lbNetTotal = Format(CDbl(lbNetTotal), "#,##0.00")
    End If
    
    SetEnabledScreen
End Sub
