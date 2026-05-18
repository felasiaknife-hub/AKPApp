VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmMain 
   BackColor       =   &H00FFFFFF&
   Caption         =   "AKP Waste Management System"
   ClientHeight    =   7260
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   14460
   Icon            =   "frMain.frx":0000
   LinkTopic       =   "Form1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   4200
      Top             =   5880
   End
   Begin MSComctlLib.ProgressBar pgb1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   1
      Top             =   6540
      Visible         =   0   'False
      Width           =   14460
      _ExtentX        =   25506
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
   End
   Begin MSComctlLib.StatusBar Stb1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   0
      Top             =   6855
      Width           =   14460
      _ExtentX        =   25506
      _ExtentY        =   714
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   7
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2646
            MinWidth        =   2646
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   3519
            MinWidth        =   3528
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            AutoSize        =   2
            Bevel           =   2
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   4057
            MinWidth        =   4057
            Text            =   "Regional: Thai ""dd/MM/yyyy"""
            TextSave        =   "Regional: Thai ""dd/MM/yyyy"""
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            Object.Width           =   882
            MinWidth        =   882
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            Object.Width           =   882
            MinWidth        =   882
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   882
            MinWidth        =   882
            TextSave        =   "INS"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0FF&
      Height          =   5715
      Left            =   0
      ScaleHeight     =   5655
      ScaleWidth      =   14400
      TabIndex        =   2
      Top             =   0
      Width           =   14460
      Begin VB.Image Image1 
         Height          =   4995
         Left            =   1080
         Stretch         =   -1  'True
         Top             =   0
         Width           =   5235
      End
   End
   Begin VB.Menu mnu_MainData 
      Caption         =   "Master"
      Begin VB.Menu mnu_CompanyData 
         Caption         =   "ข้อมูลบริษัท"
      End
      Begin VB.Menu Fin092DIW 
         Caption         =   "กำหนดข้อมูล FIN09_2 กรมฯ"
      End
      Begin VB.Menu mnu_UserGroup 
         Caption         =   "สิทธิการใช้งานระบบ"
      End
      Begin VB.Menu mnu_User 
         Caption         =   "ข้อมูลผู้ใช้งานระบบ"
      End
      Begin VB.Menu mnu_CHangePwd 
         Caption         =   "เปลี่ยนรหัสผ่าน"
      End
      Begin VB.Menu Sep1003 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Unit 
         Caption         =   "หน่วยทั่วไป"
      End
      Begin VB.Menu mnu_LABUnit 
         Caption         =   "หน่วยสำหรับห้อง LAB"
      End
      Begin VB.Menu Sep1004 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Province 
         Caption         =   "รายชื่อจังหวัด/อำเภอ/ตำบล"
      End
      Begin VB.Menu mnu_IndustryType 
         Caption         =   "ประเภทอุตสาหกรรม"
      End
      Begin VB.Menu mnu_Zone 
         Caption         =   "เขตพื้นที่ตั้งของลูกค้า"
      End
      Begin VB.Menu Sep1005 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_SaleZone 
         Caption         =   "Zone"
      End
      Begin VB.Menu mnu_SaleDistrict 
         Caption         =   "เขตการขาย"
      End
      Begin VB.Menu Sep1006 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_TruckType 
         Caption         =   "ประเภทกลุ่มหลักรถบรรทุก"
      End
      Begin VB.Menu mnu_SubTruckType 
         Caption         =   "ประเภทกลุ่มย่อยรถบรรทุก"
      End
      Begin VB.Menu Sep1007 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_EUCodeData 
         Caption         =   "EU Code"
      End
      Begin VB.Menu mnu_TreatmentData 
         Caption         =   "วิธีการบำบัด"
      End
      Begin VB.Menu Sep1008 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Alert 
         Caption         =   "ระบบแจ้งเตือน"
      End
      Begin VB.Menu mnu_InternalMail 
         Caption         =   "Internal Email"
      End
      Begin VB.Menu Sep1009 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Logout 
         Caption         =   "Login"
      End
      Begin VB.Menu mnu_DecryptPwd 
         Caption         =   "Decrypt Password"
      End
   End
   Begin VB.Menu mnu_AccountPayable 
      Caption         =   "AP"
      Begin VB.Menu mnu_AccountPayableVendor 
         Caption         =   "เจ้าหนี้ (Vendor)"
      End
      Begin VB.Menu mnu_AccountPayableGerneral 
         Caption         =   "เจ้าหนี้ (เจ้าหนี้การค้าทั่วไป)"
      End
      Begin VB.Menu mnu_AccountPayableLAB 
         Caption         =   "เจ้าหนี้ (ห้องแล็บ)"
      End
   End
   Begin VB.Menu mnu_CRJobs 
      Caption         =   "CR"
      Begin VB.Menu mnu_CustGroup 
         Caption         =   "กลุ่มลูกค้า"
      End
      Begin VB.Menu mnu_Customer 
         Caption         =   "ข้อมูลลูกค้า"
      End
      Begin VB.Menu mnuCustBase 
         Caption         =   "เปลี่ยนฐานลูกค้า"
      End
      Begin VB.Menu Sep2001 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_RemDesc 
         Caption         =   "หมายเหตุในใบเสนอราคา"
      End
      Begin VB.Menu mnu_Quotation 
         Caption         =   "ใบเสนอราคา"
      End
      Begin VB.Menu mnu_QuotationDiscount 
         Caption         =   "ส่วนลดใบเสนอราคา"
      End
      Begin VB.Menu mnu_QuotationList 
         Caption         =   "Quotation List"
      End
      Begin VB.Menu mnu_QuotationListHeader 
         Caption         =   "Quotation List Header"
      End
      Begin VB.Menu Sep2002 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Schedule 
         Caption         =   "Calendar Order"
      End
      Begin VB.Menu mnu_OrderOpen 
         Caption         =   "Sales Order"
      End
      Begin VB.Menu Mnu_JobQuotationDet 
         Caption         =   "ตรวจสอบข้อมูล Job"
      End
      Begin VB.Menu mnu_ImportDataFix 
         Caption         =   "แก้ไข Import Data"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuJobProject 
         Caption         =   "บันทึกข้อมูลงานโปรเจค"
      End
      Begin VB.Menu mnu_2003 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_MenifestCheckDoc 
         Caption         =   "ตรวจเอกสารใบเมนิเฟส"
      End
      Begin VB.Menu mnu_FrontEnd 
         Caption         =   "ข้อมูลการสั่งงาน Front end"
      End
      Begin VB.Menu MnuCarDaily 
         Caption         =   "ข้อมูลการจองรถ"
      End
   End
   Begin VB.Menu mnu_Envi 
      Caption         =   "ENV"
      Begin VB.Menu mnu_WasteQue 
         Caption         =   "รายการของเสียที่รอการตรวจสอบ"
      End
      Begin VB.Menu mnu_Sample 
         Caption         =   "บันทึกการส่งตัวอย่างกากเพื่อการวิเคราะห์"
      End
   End
   Begin VB.Menu mnu_LAB 
      Caption         =   "LAB"
      Begin VB.Menu mnu_LABVendor 
         Caption         =   "ข้อมูล Vendor ห้อง LAB ภายนอก"
      End
      Begin VB.Menu mnu_LabType 
         Caption         =   "ประเภทการทดสอบ"
      End
      Begin VB.Menu mnu_Parameters 
         Caption         =   "พารามิเตอร์ในการทดสอบ"
      End
      Begin VB.Menu mnu_TestMethod 
         Caption         =   "วิธีการทดสอบ"
      End
      Begin VB.Menu mnu_TestForm 
         Caption         =   "สร้างแบบฟอร์มการทดสอบ"
      End
      Begin VB.Menu Sep4001 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_IntSampleType 
         Caption         =   "ประเภทตัวอย่างของเสียภายใน"
      End
      Begin VB.Menu mnu_SamplingPoint 
         Caption         =   "จุดเก็บตัวอย่างของเสีย"
      End
      Begin VB.Menu mnu_IntSample 
         Caption         =   "บันทึกส่งตัวอย่างของเสีย Mornitoring"
      End
      Begin VB.Menu Sep4002 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_WasteStable 
         Caption         =   "ส่งตัวอย่างกากของเสียปรับเสถียร"
      End
      Begin VB.Menu Sep4003 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_TestQue 
         Caption         =   "รายการตัวอย่างของเสียรอการทดสอบ"
      End
      Begin VB.Menu Sep4004 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_TestData 
         Caption         =   "บันทึกผลการทดสอบกากของเสีย"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_InternalTest 
         Caption         =   "บันทึกผลการทดสอบภายใน"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnu_Trans 
      Caption         =   "TRN"
      Begin VB.Menu mnu_CarVendor 
         Caption         =   "ข้อมูล Vendor รถ"
      End
      Begin VB.Menu mnu_TSDFCompany 
         Caption         =   "ข้อมูลบริษัท TSDF"
      End
      Begin VB.Menu mnu_OrderList 
         Caption         =   "Download ข้อมูล Manifest ที่รอดำเนินการ"
         Visible         =   0   'False
      End
      Begin VB.Menu Sep5001 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_TranVendorRate 
         Caption         =   "อัตราค่าขนส่งที่จ่ายให้ Vendor (ตาราง)"
      End
      Begin VB.Menu mnu_TranCustomerRate 
         Caption         =   "อัตราค่าขนส่งที่เก็บกับลูกค้า (ตาราง)"
      End
      Begin VB.Menu Sep5002 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_JobListForCar 
         Caption         =   "รายการรอจัดรถ"
      End
      Begin VB.Menu mnu_TripControlCarBoxAssign 
         Caption         =   "บันทึกใบคุมสำหรับวางบ็อก"
      End
      Begin VB.Menu mnu_TripControlCarBoxStatus 
         Caption         =   "รายการพิมพ์ใบคุมสำหรับวางบ็อก"
      End
      Begin VB.Menu Sep5003 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_MenifestQue 
         Caption         =   "รายการสถานะใบ Manifest"
      End
      Begin VB.Menu mnu_MenifestClose 
         Caption         =   "รายการสำหรับปิดใบ Manifest"
      End
      Begin VB.Menu mnu_MenifestClosedToApprove 
         Caption         =   "รายการสำหรับอนุมัติใบเมนิเฟสที่ปิดงานแล้ว"
      End
      Begin VB.Menu Sep5004 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_ManifestBuild 
         Caption         =   "ข้อมูลใบ Manifest และใบคุม"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_QuickOrder 
         Caption         =   "เปิด Manifest แบบเร่งด่วน"
      End
      Begin VB.Menu Sep5005 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_FuelCost 
         Caption         =   "บันทึกการซื้อน้ำมันเชื้อเพลิงของ Vendor"
      End
      Begin VB.Menu mnu_MoreCost 
         Caption         =   "บันทึกค่าปรับอื่นๆ ของ Vendor ขนส่ง"
      End
      Begin VB.Menu Sep5006 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_CarStatus 
         Caption         =   "สถานะและคิวงานของรถ"
      End
      Begin VB.Menu Sep5007 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_BoxManagment 
         Caption         =   "การจัดการบ็อก"
         Begin VB.Menu mnu_BoxType 
            Caption         =   "BoxType"
         End
         Begin VB.Menu mnu_BoxOwners 
            Caption         =   "Box Owners"
         End
         Begin VB.Menu mnu_BoxNo 
            Caption         =   "Box No."
         End
      End
      Begin VB.Menu mnu_TripControlBox 
         Caption         =   "คิวงานรอจัดบ็อก"
      End
      Begin VB.Menu mnu_TruckManage 
         Caption         =   "Truck Management"
      End
      Begin VB.Menu mmu_CalOil 
         Caption         =   "คำนวนค่าขนส่ง"
      End
   End
   Begin VB.Menu mnu_Finance 
      Caption         =   "ACC"
      Begin VB.Menu mnu_TransDocCheck 
         Caption         =   "ใบเมนิเฟส"
      End
      Begin VB.Menu mnu_CloseJob 
         Caption         =   "รายการ Manifest ที่ปิดงานแล้ว"
      End
      Begin VB.Menu mnu_ListInv 
         Caption         =   "รายการ Invoice"
      End
      Begin VB.Menu Sep6001 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_PreInv 
         Caption         =   "Pre INV"
      End
      Begin VB.Menu mnu_BillingNote 
         Caption         =   "Invoice (ลูกหนี้)"
      End
      Begin VB.Menu mnu_ListInvoiceCustRecieve 
         Caption         =   "บันทึกลูกค้ารับวางบิล"
      End
      Begin VB.Menu Sep6002 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_TaxInvoice 
         Caption         =   "ออกใบกำกับภาษี/ใบเสร็จ"
      End
      Begin VB.Menu mnu_AddDept 
         Caption         =   "ออกใบเพิ่ม/ลด หนี้"
      End
      Begin VB.Menu Sep6003 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_DeptorList 
         Caption         =   "รายการเจ้าหนี้รอจ่าย"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_PayDeptor 
         Caption         =   "ชำระเงินเจ้าหนี้ (Vendor)"
      End
      Begin VB.Menu mnu_PaymentVoucher 
         Caption         =   "ชำระเงินเจ้าหนี้การค้าทั่วไป"
      End
      Begin VB.Menu Sep6004 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_MnfBill 
         Caption         =   "รายการใบเมนิเฟสที่ยังไม่ได้เก็บเงิน"
      End
      Begin VB.Menu mnu_Acc_PreInvoice 
         Caption         =   "Pre Invoice"
      End
      Begin VB.Menu Mnu_AccIvcBill 
         Caption         =   "ใบรับวางบิล"
      End
      Begin VB.Menu mnu_FnRatio 
         Caption         =   "Financial Ratios"
      End
   End
   Begin VB.Menu MnuMainPR 
      Caption         =   "PR"
      Begin VB.Menu MnuCustIn 
         Caption         =   "Customer Calendar"
      End
   End
   Begin VB.Menu mnuMainOP 
      Caption         =   "OP"
      Begin VB.Menu mnuOPCost 
         Caption         =   "ต้นทุนเผา"
      End
   End
   Begin VB.Menu mnu_CommissionSYstem 
      Caption         =   "Commission"
      Begin VB.Menu mnu_SaleTarget 
         Caption         =   "เป้าการขายรายเดือน"
      End
      Begin VB.Menu mnu_SaleTarget_List 
         Caption         =   "เป้าการขาย (LIST)"
      End
      Begin VB.Menu mnu_SalePaidRate 
         Caption         =   "กำหนดอัตราผลตอบแทนพนักงานขาย"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_PaidSupRate 
         Caption         =   "กำหนดอัตราผลตอบแทน Supervisor"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_PaidManagerRate 
         Caption         =   "กำหนดอัตราผลตอบแทนผู้จัดการ"
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_ComDeduct 
         Caption         =   "กำหนดส่วนลดยอดขายที่ราคาต่ำกว่าเกณฑ์"
         Visible         =   0   'False
      End
      Begin VB.Menu Sep7001 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_ExternalCom5 
         Caption         =   "กำหนดอัตราค่าคอมนอก"
      End
      Begin VB.Menu mnu_ComConfig 
         Caption         =   "ค่าคงที่อื่นๆ ในการคิดต่าคอมมิชชั่น"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnu_Rpt 
      Caption         =   "Report"
      Begin VB.Menu mnu_CustomerReport 
         Caption         =   "รายงานลูกค้า"
         Begin VB.Menu mnu_CustomerReport01 
            Caption         =   "รายชื่อลูกค้า"
         End
         Begin VB.Menu mnu_CustomerReport02 
            Caption         =   "รายชื่อลูกค้า DTeam"
         End
         Begin VB.Menu mnu_AllCustomerPrice 
            Caption         =   "รายงานลูกค้าใช้บริการ"
         End
      End
      Begin VB.Menu mnu_QUO_CR 
         Caption         =   "ใบเสนอราคาและงาน CR"
         Begin VB.Menu mnu_QUO001 
            Caption         =   "เปรียบเทียบจำนวนใบเสนอราคา (QUO01)"
         End
         Begin VB.Menu mnu_RptQUO02 
            Caption         =   "ใบเสนอราคาที่ออกไปในช่วงเวลาที่กำหนดจำแนกตามลูกค้า (QUO02)"
         End
         Begin VB.Menu mnu_RptQuo03 
            Caption         =   "รายชื่อลูกค้าที่เปิดใหม่จำแนกตามลูกค้า (QUO03)"
         End
         Begin VB.Menu mnu_RptQuo04 
            Caption         =   "สัญญาการใช้บริการ (QUO04)"
         End
         Begin VB.Menu mnu_Auction 
            Caption         =   "งานประมูล"
         End
         Begin VB.Menu mnu_RptCR01 
            Caption         =   "จำนวนใบเสนอราคาและจำนวนเที่ยวจำแนกตาม CR (CR01)"
         End
         Begin VB.Menu mnu_RptCR02 
            Caption         =   "*สรุปผลการดำเนินงานเจ้าหน้าที่การตลาด (CR02)"
         End
         Begin VB.Menu mnu_CR03 
            Caption         =   "รายงานแจ้งการกำจัดและการบำบัดกาก (CR03)"
         End
         Begin VB.Menu CR05 
            Caption         =   "รายงานเปรียบเทียบใบเสนอราคา (CR05)"
         End
         Begin VB.Menu CRBigbag 
            Caption         =   "รายงาน Bigbag"
         End
         Begin VB.Menu MnuRptCovidbin 
            Caption         =   "รายงานถังขยะโควิด"
         End
         Begin VB.Menu MnuRptQuotation 
            Caption         =   "รายงานใบเสนอราคาปลายทางอื่น"
         End
      End
      Begin VB.Menu mnu_RptOrderMain 
         Caption         =   "Order/ยอดขาย"
         Begin VB.Menu mnu_ORD01 
            Caption         =   "จำนวนเที่ยวที่เปิด Order การขนส่งรายเดือน (ORD01)"
         End
         Begin VB.Menu mnu_RptORD02 
            Caption         =   "ยอดขายเทียบเป้าหมาย รายเขต ทั้งจำนวนเที่ยว และมูลค่า (ORD02)"
         End
         Begin VB.Menu mnu_RptOrd04 
            Caption         =   "รายงานบริษัทที่ไม่มีการขน (ORD04)"
         End
         Begin VB.Menu mnu_RptOrd05 
            Caption         =   "จำนวนเที่ยวที่เปิด Order การขนส่งรายปี (ORD05)"
         End
         Begin VB.Menu ORD06 
            Caption         =   "รายงานยกเลิก Order"
         End
      End
      Begin VB.Menu mnu_RptLABENVI 
         Caption         =   "Lab / สิ่งแวดล้อม"
         Begin VB.Menu mnu_RptLAB01 
            Caption         =   "ผลการวิเคราะห์ LAB พร้อมทั้งค่าใช้จ่าย (LAB01)"
         End
         Begin VB.Menu mnu_RptLAB03 
            Caption         =   "จำนวน waste ที่มีการส่งเพื่อวิเคราะห์ (LAB03)"
         End
         Begin VB.Menu mnu_RptENV01 
            Caption         =   "รายงานการตรวจรับ Waste (ENV01)"
         End
         Begin VB.Menu mnu_RptEnv02 
            Caption         =   "รายงานปริมาณกากอุตสาหกรรมที่รับกำจัด (ENV02)"
         End
         Begin VB.Menu mnu_RptENV03 
            Caption         =   "จำนวน Waste ที่มีการส่งเพื่อวิเคราะห์ (ENV03)"
         End
         Begin VB.Menu mnu_RptENV04 
            Caption         =   "รายงานวันใบอนุญาตหมดอายุของ Waste (ENV04)"
         End
         Begin VB.Menu rptEnv05 
            Caption         =   "แบบแสดงรหัสของชนิดและประเภทสิ่งปฏิกูล (ENV05)"
         End
      End
      Begin VB.Menu mnu_TransRpt 
         Caption         =   "ระบบขนส่ง"
         Begin VB.Menu mnu_RptTRN01 
            Caption         =   "รายงานการขนกากตามประเภทรถ (TRN01)"
         End
         Begin VB.Menu mnu_RptTRN02 
            Caption         =   "ใบกำกับการขนส่งที่ปิดงานแล้ว (TRN02)"
         End
         Begin VB.Menu mnu_RptTRN03 
            Caption         =   "รายงานตารางการขนส่งรายวัน (TRN03)"
         End
         Begin VB.Menu mnu_RptTRN03A 
            Caption         =   "รายงานตารางการขนส่งรายวัน (TRN03-A)"
         End
         Begin VB.Menu mnu_RptTRN04 
            Caption         =   "ค่าขนส่ง, ค่าแรงคนงาน by vendor (TRN04)"
         End
         Begin VB.Menu mnu_RptTRN05 
            Caption         =   "Order ที่มีการยืนยันการบริการขนกาก (TRN05)"
         End
         Begin VB.Menu mnu_RptTRN06 
            Caption         =   "รายงานการขนส่งประจำวันแยกตามประเภทของเสีย (TRN06)"
         End
         Begin VB.Menu mnu_RptTRN07 
            Caption         =   "รายงานสรุปสถานะใบเมนิเฟส (TRN07)"
         End
         Begin VB.Menu mnu_RptTRN08 
            Caption         =   "รายงานสรุปใบเมนิเฟสที่ไม่มีการขนเวสต์ (TRN08)"
         End
         Begin VB.Menu mnu_RptTRN09 
            Caption         =   "รายงานสรุปการขนส่ง (TRN09)"
         End
         Begin VB.Menu Sep8004 
            Caption         =   "-"
         End
         Begin VB.Menu mnu_RptTRN10 
            Caption         =   "รายงานการขนกากแยกตามประเภทเวสต์ (TRN10)"
         End
         Begin VB.Menu mnu_RptTRN10A 
            Caption         =   "รายงานระยะทางการขนกาก(TRN10A)"
         End
         Begin VB.Menu mnu_RptTRN10A1 
            Caption         =   "รายงานระยะทางการขนกากของรถแต่ละคัน(TRN10A-1)"
         End
         Begin VB.Menu mnu_RptTRN11 
            Caption         =   "รายงานการขนกากที่มีการจ่ายค่าคอมฯ นอก (TRN11)"
         End
         Begin VB.Menu Sep8005 
            Caption         =   "-"
         End
         Begin VB.Menu mnu_RptTRN12 
            Caption         =   "รายงานการขนกากแยกตามประเภทรถ (TRN12)"
         End
         Begin VB.Menu mnu_RptTRN13 
            Caption         =   "รายงานการขนกากแยกตามประเภทรถ (TRN13)"
         End
         Begin VB.Menu Mu_Trn14 
            Caption         =   "รายงานการขนกากที่ไม่ได้ของ(TRN14)"
         End
         Begin VB.Menu Mu_Trn15 
            Caption         =   "รายงานสรุปจ่ายเบี้ยเลี้ยงพนักงานขับ(TRN15)"
         End
         Begin VB.Menu Mu_Trn16 
            Caption         =   "รายงานสรุปเบี้ยเลี้ยง(TRN16)"
         End
         Begin VB.Menu TRN17 
            Caption         =   "รายงานการขนใบงานที่ส่งเจ้าหน้าที่แยกน้ำหนัก(TRN17)"
         End
         Begin VB.Menu mnu_RptBoxTransaction 
            Caption         =   "รายงานการติดตามบ็อก"
            Visible         =   0   'False
         End
         Begin VB.Menu MnuTRN10EX 
            Caption         =   "รายงานคุณวันชัย"
            Visible         =   0   'False
         End
      End
      Begin VB.Menu mnu_RptAcc 
         Caption         =   "บัญชี / การเงิน"
         Begin VB.Menu mnu_RPTFIN01 
            Caption         =   "สรุปใบวางบิล (FIN01)"
         End
         Begin VB.Menu mnu_RptFIN02 
            Caption         =   "@ สรุปยอดใบสำคัญรับ เพื่อประมาณการยอดรายรับ (FIN02)"
            Visible         =   0   'False
         End
         Begin VB.Menu mnu_RptFIN03 
            Caption         =   "สรุปยอดใบสำคัญจ่าย โดยคำนวณภาษีหัก ณ. ที่จ่าย (FIN03)"
         End
         Begin VB.Menu mnu_RptFIN04 
            Caption         =   "สรุปการนัดรับเช็ค (FIN04)"
         End
         Begin VB.Menu mnu_RptFIN05 
            Caption         =   "ภาษีขาย (FIN05)"
         End
         Begin VB.Menu mnu_RptFIN06 
            Caption         =   "*จำนวนเที่ยวการขนและจำนวนเงิน (FIN06)"
            Visible         =   0   'False
         End
         Begin VB.Menu mnu_RptFIN07 
            Caption         =   "อายุลูกหนี้ (FIN07)"
         End
         Begin VB.Menu mnu_RptFin08 
            Caption         =   "อายุเจ้าหนี้ (FIN08)"
         End
         Begin VB.Menu mnu_RptFIn09 
            Caption         =   "การขนแยกตาม Manifest และสถานะการตรวจสอบเอกสาร (FIN09)"
         End
         Begin VB.Menu mnu_Rpt9_2 
            Caption         =   "การขนแยกตาม Manifest (FIN09_2)"
         End
         Begin VB.Menu FiN10 
            Caption         =   "ใบงานยังไม่ได้วางบิล(ไม่ได้ของ) (FIN10)"
         End
         Begin VB.Menu FIN11 
            Caption         =   "รายงานเนนิเฟสที่วางบิลแต่งังไม่ได้รับเงิน(RV)(FIN11)"
         End
         Begin VB.Menu FIN12 
            Caption         =   "รายงานใบเมนิเฟสยังไม่ได้วางบิล (FIN12)"
         End
         Begin VB.Menu FIN13 
            Caption         =   "รายงานรายการวางบิล (FIN13)"
         End
         Begin VB.Menu MNUFIN13A 
            Caption         =   "รายงานรายการวางบิล (FIN13_A)"
         End
         Begin VB.Menu FIN14 
            Caption         =   "รายงานการรับเช็ค (FIN14)"
         End
         Begin VB.Menu Fin15 
            Caption         =   "รายงานค่าบริการ (FIN15)"
         End
         Begin VB.Menu Fin15A 
            Caption         =   "รายงานค่าบริการ (FIN15_A)"
         End
         Begin VB.Menu Fin16 
            Caption         =   "รายงานต้นทุน (FIN16)"
         End
         Begin VB.Menu mnu_ChkReport 
            Caption         =   "รายงานเช็คนำฝากธนาคาร"
         End
      End
      Begin VB.Menu mnu_RptCommision 
         Caption         =   "คอมมิชชั่น"
         Visible         =   0   'False
         Begin VB.Menu mnu_RptCOM1 
            Caption         =   "*การจ่ายค่าคอมมิชชั่นแยกตามใบ Manifest (COM1)"
         End
         Begin VB.Menu mnu_RptCOM2 
            Caption         =   "*สรุปการจ่ายค่าคอมภายนอก (COM2)"
         End
         Begin VB.Menu mnu_RptCom3 
            Caption         =   "ยอดขายและการคิดค่าคอมใน (COM3_In)"
         End
         Begin VB.Menu mnu_RptCom3Ex 
            Caption         =   "ยอดขายและการคิดค่าคอมนอก (COM3_Ex)"
         End
         Begin VB.Menu mnu_RptORD03 
            Caption         =   "*ค่าคอมมิชชั่น สำหรับ Sales และบุคคลภายนอก (ORD03)"
         End
         Begin VB.Menu mnu_UpTF 
            Caption         =   "Update Treatment Fee"
            Visible         =   0   'False
         End
      End
      Begin VB.Menu mnu_rdf 
         Caption         =   "RDF"
         Begin VB.Menu mnu_rdfList 
            Caption         =   "รายงานการขน Waste RDF"
         End
      End
      Begin VB.Menu Box01 
         Caption         =   "รายงาน Box"
         Begin VB.Menu Box_01 
            Caption         =   "รายงานการวาง Box ลูกค้า"
         End
      End
      Begin VB.Menu MnuCon 
         Caption         =   "รายงานภาชนะ"
         Begin VB.Menu MnuConTake 
            Caption         =   "การเบิกภาชนะ"
         End
         Begin VB.Menu MnuConAdd 
            Caption         =   "การเพิ่ม - ลด ภาชนะ"
         End
      End
      Begin VB.Menu ReportPR 
         Caption         =   "PR"
         Begin VB.Menu MnuCustVisit 
            Caption         =   "รายงานลูกค้าเข้าศูนย์ (PR01)"
         End
      End
   End
   Begin VB.Menu mnu_Income 
      Caption         =   "Income"
      Begin VB.Menu mnu_RptIncome 
         Caption         =   "รายงานยอดขาย"
      End
      Begin VB.Menu mnu_RptIncomeSub 
         Caption         =   "รายงานยอดขาย Sub"
      End
      Begin VB.Menu IncomePST 
         Caption         =   "รายงานยอดขาย TK"
         Visible         =   0   'False
      End
      Begin VB.Menu incomeGE 
         Caption         =   "รายงานยอดขาย GE"
         Visible         =   0   'False
      End
      Begin VB.Menu incomeRDF 
         Caption         =   "รายงานยอดขาย RDF"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnu_ITApplication 
      Caption         =   "IT"
      Begin VB.Menu mnu_ITCompareIncome 
         Caption         =   "เปรียบเทียบยอดขาย"
      End
      Begin VB.Menu mnu_ITRecoverQuo 
         Caption         =   "กู้ข้อมูล ใบเสนอราคา"
      End
      Begin VB.Menu mnu_ITRecoverOrder 
         Caption         =   "กู้ข้อมูล Order"
      End
      Begin VB.Menu mnu_ITRecoverMenifest 
         Caption         =   "กู้ข้อมูล ใบเมนิเฟส"
      End
      Begin VB.Menu mnu_ITRecoverInv 
         Caption         =   "กู้ขอมูล INV ที่ถูกยกเลิก"
      End
      Begin VB.Menu mnu_ITRecoverRV 
         Caption         =   "กู้ข้อมูล ใบเสร็จรับเงิน"
      End
      Begin VB.Menu mnu_ITSaleChanged 
         Caption         =   "เปลี่ยนแปลงพนักงานขาย"
      End
   End
   Begin VB.Menu mnu_Monitor 
      Caption         =   "Monitor"
      Begin VB.Menu mnu_Monitor_Truck 
         Caption         =   "Truck Monitor"
      End
   End
   Begin VB.Menu mnu_Window 
      Caption         =   "Window"
      WindowList      =   -1  'True
   End
   Begin VB.Menu mnu_WhatNew 
      Caption         =   "What New?"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private ImageTime As Integer
Private ImageIndex As Integer
Private IntervalTime As Long

Public Sub ActivateForm()
    On Error Resume Next
    Dim TotalWidth As Long

    Stb1.Panels(1).Text = App.Title
    Stb1.Panels(2).Text = ""
    Stb1.Panels(3).Text = ""
    Stb1.Panels(3).Text = "ชื่อผู้เข้าใช้งาน: " & CurrentUser.StaffFName & "  " & CurrentUser.StaffLName
    Stb1.Panels(4).Text = "Regional: " & GetUserLocaleInfo(LCID, LOCALE_SLANGUAGE) & " (" & GetUserLocaleInfo(LCID, CAL_SSHORTDATE) & ")"
    
    TotalWidth = Stb1.Width - (Stb1.Panels(4).Width + Stb1.Panels(5).Width + Stb1.Panels(5).Width + Stb1.Panels(7).Width)
    Stb1.Panels(1).Width = (70 / 100 * TotalWidth)
    Stb1.Panels(2).Width = (15 / 100 * TotalWidth)
    Stb1.Panels(3).Width = (15 / 100 * TotalWidth)
    Stb1.Panels(4).Width = 2300
    Stb1.Panels(5).Width = 500
    Stb1.Panels(6).Width = 500
    Stb1.Panels(7).Width = 500
    Stb1.Refresh
    Me.Caption = "AKP" & " " & App.Title & "   " & App.Major & "." & App.Minor & "." & App.Revision & "  Running on " & strDb
End Sub

Public Sub InitializeMenu()
Dim ShowMenu As Boolean
    If LCase(CurrentUser.LoginName) <> "admin" Then
        ShowMenu = False
    Else
        ShowMenu = True
    End If
    Me.mnu_CompanyData.Visible = ShowMenu
    Me.mnu_UserGroup.Visible = ShowMenu
    Me.mnu_User.Visible = ShowMenu
'    Me.Sep1002.Visible = ShowMenu
    Me.Sep1003.Visible = ShowMenu
    Me.mnu_Unit.Visible = ShowMenu
    Me.mnu_LABUnit.Visible = ShowMenu
    Me.Sep1004.Visible = ShowMenu
    Me.mnu_Province.Visible = ShowMenu
    Me.mnu_IndustryType.Visible = ShowMenu
    Me.mnu_Zone.Visible = ShowMenu
    Me.Sep1005.Visible = ShowMenu
    Me.mnu_SaleZone.Visible = ShowMenu
    Me.mnu_SaleDistrict.Visible = ShowMenu
    Me.Sep1006.Visible = ShowMenu
    Me.mnu_TruckType.Visible = ShowMenu
    Me.mnu_SubTruckType.Visible = ShowMenu
    Me.Sep1007.Visible = ShowMenu
    Me.mnu_EUCodeData.Visible = ShowMenu
'    Me.mnu_TreatmentData.Visible = ShowMenu
    Me.Sep1008.Visible = ShowMenu
    Me.mnu_Alert.Visible = ShowMenu
    Me.mnu_InternalMail.Visible = ShowMenu
    Me.Sep1009.Visible = ShowMenu
    Me.mnu_DecryptPwd.Visible = ShowMenu
    Me.mnu_OrderList.Visible = ShowMenu
'    Me.mnuCustBase.Visible = ShowMenu
End Sub

Private Sub Box_01_Click()
frmPrintBox01.Show 1
End Sub

Private Sub CR05_Click()
frmPrintCR05.Show 1
End Sub

Private Sub CRBigbag_Click()
FrmCRBigbag.Show 1
End Sub

Private Sub Fin092DIW_Click()
    FrmFin09DIW.Show 1
End Sub

Private Sub FiN10_Click()
    frmPrintTRN14.Show 1
End Sub

Private Sub FIN11_Click()
    frmPrintFIN11.Show 1
End Sub

Private Sub FIN12_Click()
    frmPrintFIN12.Show 1
End Sub

Private Sub FIN13_Click()
    frmPrintFIN13.Show 1
End Sub

Private Sub FIN14_Click()
       frmPrintFIN14.Show 1
End Sub

Private Sub Fin15_Click()
    FrmPrintFIN15.Show 1
End Sub

Private Sub Fin15A_Click()
    FrmPrintFIN15A.Show 1
End Sub

Private Sub Fin16_Click()
    FrmPrintFIN16.Show 1
End Sub

'Private Sub incomeAP_Click()
'frmIncomeAKP.Show 1
'End Sub

Private Sub incomeGE_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmIncome", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
frmIncomeGE.Show 1
End Sub

Private Sub IncomePST_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmIncome", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
frmIncomepst.Show 1
End Sub

Private Sub incomeRDF_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmIncome", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    frmIncomeRDF.Show 1
End Sub

Private Sub MDIForm_Load()
Dim Appv As String
Dim Appn As String
Dim tmpRec As ADODB.Recordset
Set tmpRec = New ADODB.Recordset
'Dim tmpStr As String
'    tmpStr = CurrentUser.VerifyAccessRole("frmMenifestStatusList", AccessRole)
'    If Trim(tmpStr) <> "" Then
'        MsgBox tmpStr, vbExclamation
'        Exit Sub
'    End If
    InitializeMenu
    ActivateForm
    Me.Caption = "AKP" & " " & App.Title & "   " & App.Major & "." & App.Minor & "." & App.Revision '& "  Database on " & strDb
    PopulateUnit
    Dim sqlStr As String
    sqlStr = "Select vApp From ApplicationTB Where ApplicationName = 'AKPApp' "
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec.EOF = False Then
    Appv = tmpRec(0)
    End If
    Me.Caption = Me.Caption & "    V." & App.Major & "." & App.Minor & "." & App.Revision
    Appn = App.Major & "." & App.Minor & "." & App.Revision
    If Appn <> Appv Then
        MsgBox ("โปรแกรมที่ท่านใช้อยู่  v. " & Appn & "ไม่ใช่ Version ปัจจุบัน ท่านต้องทำการ Update Program ก่อนครับ ปัจจุบันคือ Version " & Appv & " ")
        End
    End If
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = CloseAllForms
    Unload frmPicture
'    Localization.SetDefaultRegionalSettings DefaultLCID
End Sub

Private Sub MDIForm_Resize()
On Error Resume Next
'    ActivateForm
    Dim TotalWidth As Long
    TotalWidth = Stb1.Width - (Stb1.Panels(4).Width + Stb1.Panels(5).Width + Stb1.Panels(5).Width + Stb1.Panels(7).Width)
    Stb1.Panels(1).Width = (70 / 100 * TotalWidth)
    Stb1.Panels(2).Width = (15 / 100 * TotalWidth)
    Stb1.Panels(3).Width = (15 / 100 * TotalWidth)
    
    Picture1.Height = Me.Height - (Stb1.Height * 3)
    Image1.Width = Picture1.Width
    Image1.Height = Picture1.Height
End Sub

Private Sub mmu_CalOil_Click()
    frmCalOil.Show
End Sub

Private Sub mnu_Acc_PreInvoice_Click()
    Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmPreInvoice", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    VisibleImage False
    frmPreInvoice.Show
    frmPreInvoice.ZOrder
End Sub

Private Sub Mnu_AccIvcBill_Click()
    FrmIvcBill.Show 1
End Sub

Private Sub mnu_AccountPayableGerneral_Click()
    Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmAccountPayable", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    frmAccountPayable.Show 1
End Sub

Private Sub mnu_AccountPayableLAB_Click()
    Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmLABVender", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    frmLABVender.Show 1
End Sub

Private Sub mnu_AccountPayableVendor_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmTransportVender", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    frmTransportVender.Show 1
End Sub

Private Sub mnu_AddDept_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmAdjustDetp", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    frmAdjustDetp.Show 1
End Sub

Private Sub mnu_Alert_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmAlert", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    frmAlert.Show 1
End Sub

Private Sub mnu_AllCustomerPrice_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmIncome", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    
    Dim UserGrp As New AWSUserGroup
    Dim isMD As Boolean, i As Integer
    If Not CurrentUser Is Nothing Then
        For i = 1 To CurrentUser.getUserGroups.Count
            Set UserGrp = CurrentUser.getUserGroups(i)
            If Trim(UserGrp.Name) = "MD" Then
                isMD = True
                Exit For
            Else
                isMD = False
            End If
        Next i
    End If
    frmPrintAllCustomerPrice.Show 1
End Sub

Private Sub mnu_Auction_Click()
frmPrintAuction.Show 1
End Sub

Private Sub mnu_BillingNote_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmInvoice", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    frmInvoice.Show 1
End Sub

Private Sub mnu_CalInternalCom_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmComProcess", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    frmComProcess.Show 1
End Sub

Private Sub mnu_BoxNo_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmTransportRate", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    frmBoxManaulGen.Show 1
End Sub

Private Sub mnu_BoxOwners_Click()
    Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmTransportRate", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    frmBoxOwners.Show 1
End Sub

Private Sub mnu_BoxType_Click()
    Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmTransportRate", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    frmBoxType.Show 1
End Sub

Private Sub mnu_CarStatus_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmCarWorkingQue", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    VisibleImage False
    frmCarWorkingQue.Show
    frmCarWorkingQue.ZOrder
End Sub

Private Sub mnu_CarVendor_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmTransportVender", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    frmTransportVender.Show 1
End Sub

Private Sub mnu_CHangePwd_Click()
    Load frmChangePwd
    With frmChangePwd
        .SetShowUser CurrentUser
        .Show 1
    End With
End Sub

Private Sub mnu_ChkReport_Click()
    frmRERptCheque.Show 1
End Sub

Private Sub mnu_CloseJob_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmMenifestClosedList", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    VisibleImage False
    With frmMenifestClosedList
        '.cboJobStatus.ListIndex = 3
        .Show
        .ZOrder
    End With
End Sub

Private Sub mnu_ComConfig_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmComRateConfig", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    frmComRateConfig.Show 1
End Sub

Private Sub mnu_ComDeduct_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmComRateDeduct", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    frmComRateDeduct.Show 1
End Sub

Private Sub mnu_CompanyData_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmCompany", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    frmCompany.Show 1
End Sub

Private Sub mnu_CR03_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmPrintCR03", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    frmPrintCR03.Show 1
End Sub

Private Sub mnu_CustGroup_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmCustomerGroup", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    frmCustomerGroup.Show 1
End Sub

Private Sub mnu_Customer_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmCustomer", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    frmCustomer.Show 1
End Sub

Private Sub mnu_CustomerReport01_Click()
    frmCustomerReport01.Show 1
    frmCustomerReport01.ZOrder
End Sub

Private Sub mnu_CustomerReport02_Click()
    frmCustomerReport02.Show 1
    frmCustomerReport02.ZOrder
End Sub

Private Sub mnu_DecryptPwd_Click()
    frmInitialLogin.Show 1
End Sub

Private Sub mnu_DeptorList_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmeDeptorList", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    VisibleImage False
    frmeDeptorList.Show
    frmeDeptorList.ZOrder
End Sub

Private Sub mnu_EUCodeData_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmEUCode", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    frmEUCode.Show 1
End Sub

Private Sub mnu_ExternalCom5_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmComExternal", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    frmComExternal.Show 1
End Sub

Private Sub mnu_FnRatio_Click()
    FrmFinancialRatio.lblNew.Caption = "Y"
    FrmFinancialRatio.Show
End Sub

Private Sub mnu_FrontEnd_Click()
    FrmFrontEnd.Show 1
End Sub

Private Sub mnu_FuelCost_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmFuelBuyRecord", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    frmFuelBuyRecord.Show 1
End Sub

Private Sub mnu_ImportDataFix_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmImportDataFix", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    frmImportDataFix.Show 1
End Sub

Private Sub mnu_IndustryType_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmIndustryType", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    frmIndustryType.Show 1
End Sub

Private Sub mnu_IntSample_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmSendSampleInt", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    frmSendSampleInt.Show 1
End Sub

Private Sub mnu_IntSampleType_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmInternalSampleType", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    frmInternalSampleType.Show 1
End Sub

Private Sub mnu_ITCompareIncome_Click()
    Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmITApplication", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    VisibleImage False
    frmCompareIncome.Show
    frmCompareIncome.ZOrder
End Sub

Private Sub mnu_ITRecoverInv_Click()
    Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmITApplication", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    Load frmRecoverDocument
    frmRecoverDocument.DocType = "BillingNo"
    frmRecoverDocument.Show 1
'    frmRecoverDocument.ZOrder
End Sub

Private Sub mnu_ITRecoverMenifest_Click()
    Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmITApplication", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    Load frmRecoverDocument
    frmRecoverDocument.DocType = "Menifest"
    frmRecoverDocument.Show 1
'    frmRecoverDocument.ZOrder
End Sub

Private Sub mnu_ITRecoverOrder_Click()
    Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmITApplication", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    Load frmRecoverDocument
    frmRecoverDocument.DocType = "JobNo"
    frmRecoverDocument.Show 1
'    frmRecoverDocument.ZOrder
End Sub

Private Sub mnu_ITRecoverQuo_Click()
    Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmITApplication", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    Load frmRecoverDocument
    frmRecoverDocument.DocType = "Quotation"
    frmRecoverDocument.Show 1
'    frmRecoverDocument.ZOrder
End Sub

Private Sub mnu_ITRecoverRV_Click()
    Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmITApplication", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    Load frmRecoverDocument
    frmRecoverDocument.DocType = "InvNo"
    frmRecoverDocument.Show 1
'    frmRecoverDocument.ZOrder
End Sub

Private Sub mnu_ITSaleChanged_Click()
    Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmITApplication", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    VisibleImage False
    frmSaleChanged.Show
    frmSaleChanged.ZOrder
End Sub

Private Sub mnu_JobListForCar_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmJobForCarAssignList", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    VisibleImage False
    frmJobForCarAssignList.Show
    frmJobForCarAssignList.ZOrder
End Sub

Private Sub Mnu_JobQuotationDet_Click()
        FrmJobQuo.Show
End Sub

Private Sub mnu_LabType_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmLABTestType", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    frmLABTestType.Show 1
End Sub

Private Sub mnu_LABUnit_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmLABUnit", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    frmLABUnit.Show 1
End Sub

Private Sub mnu_LABVendor_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmLABVender", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    frmLABVender.Show 1
End Sub

Private Sub mnu_ListInv_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmMenifestClosedList", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    VisibleImage False
    With frmListInv
        .Show
        .ZOrder
    End With
End Sub

Private Sub mnu_ListInvoiceCustRecieve_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmMenifestStatusListChkRe", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    VisibleImage False
    With frmMenifestStatusListChkRe
        .Show
        .ZOrder
    End With
End Sub

Private Sub mnu_Logout_Click()
    SetLogin frmMain
End Sub

Private Sub mnu_ManifestBuild_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmMenifest", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    frmMenifest.SSTab1.TabVisible(3) = False
    frmMenifest.Show 1
End Sub

Private Sub mnu_MenifestCheckDoc_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmMenifestStatusListChkDoc", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    VisibleImage False
    With frmMenifestStatusListChkDoc
        .Show
        .ZOrder
    End With
End Sub

Private Sub mnu_MenifestClose_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmMenifestStatusList", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    VisibleImage False
    frmMenifestStatusListClose.Show
    frmMenifestStatusListClose.ZOrder
End Sub

Private Sub mnu_MenifestClosedToApprove_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmMenifestStatusListApprove", ApproveRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    VisibleImage False
    frmMenifestStatusListApprove.Show
    frmMenifestStatusListApprove.ZOrder
End Sub

Private Sub mnu_MenifestQue_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmMenifestStatusList", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    VisibleImage False
    frmMenifestStatusList.Show
    frmMenifestStatusList.ZOrder
End Sub

Private Sub mnu_MnfBill_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmListMenifestBill", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    VisibleImage False
    frmListMenifestBill.Show
End Sub

Private Sub mnu_Monitor_Truck_Click()
    frmTruckMonitor.Show 1
End Sub

Private Sub mnu_MoreCost_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmVendorExpense", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    frmVendorExpense.Show 1
End Sub

Private Sub mnu_ORD01_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmPrintORD01", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    frmPrintORD01.Show 1
End Sub

Private Sub mnu_OrderList_Click()
'    If MsgBox("ท่านต้องการ Download ข้อมูล Menifest ที่รอดำเนินการหรือไม่?", vbQuestion + vbYesNo) = vbYes Then
'        MsgBox "เสร็จสิ้นการ Download ข้อมูล", vbInformation
'    End If
End Sub

Private Sub mnu_OrderOpen_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmJobOpen", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    frmJobOpen.gVal = "1"
    frmJobOpen.Show 1
End Sub

Private Sub mnu_PaidManagerRate_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmComRateManager", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    frmComRateManager.Show 1
End Sub

Private Sub mnu_PaidSupRate_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmComRateSup", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    frmComRateSup.Show 1
End Sub

Private Sub mnu_Parameters_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmLABParam", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    frmLABParam.Show 1
End Sub

Private Sub mnu_PayDeptor_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmPaytoDeptor", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    frmPaytoDeptor.Show 1
End Sub

Private Sub mnu_PaymentVoucher_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmPaytoDeptor", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    frmPaymentVoucher.Show 1
End Sub

Private Sub mnu_PreInv_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmInvoice", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    frmInvoice.Label1(3).Caption = "Pre INV"
    frmInvoice.Show 1
End Sub

Private Sub mnu_Province_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmProvince", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    frmProvince.Show 1
End Sub

Private Sub mnu_QuickOrder_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmMenifestUrgent", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    frmMenifestUrgent.SSTab1.TabVisible(2) = False
    frmMenifestUrgent.Show 1
    'frmJobCarAssign.Show 1
End Sub
Private Sub mnu_QuotationDiscount_click()
'Dim tmpStr As String
'     tmpStr = CurrentUser.VerifyAccessRole("frmQuotationDiscout", AccessRole)
'      If Trim(tmpStr) <> "" Then
'        MsgBox tmpStr, vbExclamation
'        Exit Sub
'    End If
    frmQuotationDiscout.Show 1
End Sub
Private Sub mnu_QUO001_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmPrintQUO01", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
     frmPrintQUO01.Show 1
End Sub

Private Sub mnu_Quotation_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmQuotation", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    frmQuotation.Show 1
End Sub

Private Sub mnu_QuotationList_Click()
    Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmQuotationList", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    VisibleImage False
    frmQuotationList.Show
    frmQuotationList.ZOrder
End Sub

Private Sub mnu_QuotationListHeader_Click()
    Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmQuotationList", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    VisibleImage False
    FrmQuotationListHeader.Show
    FrmQuotationListHeader.ZOrder
End Sub

Private Sub mnu_rdfList_Click()
'Dim tmpStr As String
'    tmpStr = CurrentUser.VerifyAccessRole("frmRemDesc", AccessRole)
'    If Trim(tmpStr) <> "" Then
'        MsgBox tmpStr, vbExclamation
'        Exit Sub
'    End If
    FrmPrintRDF01.Show 1
End Sub

Private Sub mnu_RemDesc_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmRemDesc", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    frmRemDesc.lvList.Checkboxes = False
    frmRemDesc.Show 1
End Sub

Private Sub mnu_Rpt9_2_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmPrintFIN09P2", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    frmPrintFIN09P2.Show 1
End Sub

Private Sub mnu_RptCR01_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmPrintCR01", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    frmPrintCR01.Show 1
End Sub

Private Sub mnu_RptCR02_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmPrintCR02", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    frmPrintCR02.Show 1
End Sub

Private Sub mnu_RptENV01_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmPrintENV01", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    frmPrintENV01.Show 1
End Sub

Private Sub mnu_RptEnv02_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmPrintENV02", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    frmPrintENV02.Show 1
End Sub

Private Sub mnu_RptENV03_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmPrintENV03", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    frmPrintENV03.Show 1
End Sub

Private Sub mnu_RptENV04_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmPrintENV04", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    frmPrintENV04.Show 1
End Sub

Private Sub mnu_RPTFIN01_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmPrintFIN01", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    frmPrintFIN01.Show 1
End Sub

Private Sub mnu_RptFIN02_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmPrintFIN02", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    frmPrintFIN02.Show 1
End Sub

Private Sub mnu_RptFIN03_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmPrintFIN03", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    frmPrintFIN03.Show 1
End Sub

Private Sub mnu_RptFIN04_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmPrintFIN04", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    frmPrintFIN04.Show 1
End Sub

Private Sub mnu_RptFIN05_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmPrintFIN05", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    frmPrintFIN05.Show 1
End Sub

Private Sub mnu_RptFIN06_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmPrintFIN06", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    frmPrintFIN06.Show 1
End Sub

Private Sub mnu_RptFIN07_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmPrintFIN07", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    frmPrintFIN07.Show 1
End Sub

Private Sub mnu_RptFin08_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmPrintFIN08", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    frmPrintFIN08.Show 1
End Sub

Private Sub mnu_RptFIn09_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmPrintFIN09", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    frmPrintFIN09.Show 1
End Sub

Private Sub mnu_RptIncome_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmIncome", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    
    Dim UserGrp As New AWSUserGroup
    Dim isMD As Boolean, i As Integer
    If Not CurrentUser Is Nothing Then
        For i = 1 To CurrentUser.getUserGroups.Count
            Set UserGrp = CurrentUser.getUserGroups(i)
            If Trim(UserGrp.Name) = "MD" Then
                isMD = True
                Exit For
            Else
                isMD = False
            End If
        Next i
    End If
    frmIncome.LblForm.Caption = ""
    frmIncome.isMD = isMD
    frmIncome.Show 1
    
End Sub

Private Sub mnu_RptIncomeSub_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmIncome", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    
    Dim UserGrp As New AWSUserGroup
    Dim isMD As Boolean, i As Integer
    If Not CurrentUser Is Nothing Then
        For i = 1 To CurrentUser.getUserGroups.Count
            Set UserGrp = CurrentUser.getUserGroups(i)
            If Trim(UserGrp.Name) = "MD" Then
                isMD = True
                Exit For
            Else
                isMD = False
            End If
        Next i
    End If
    frmIncome.LblForm.Caption = "ME"
    frmIncome.isMD = isMD
    frmIncome.Show 1
End Sub

Private Sub mnu_RptLAB01_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmPrintLAB01", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    frmPrintLAB01.Show 1
End Sub

Private Sub mnu_RptLAB02_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmPrintLAB02", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    frmPrintLAB02.Show 1
End Sub

Private Sub mnu_RptLAB03_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmPrintLAB03", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    frmPrintLAB03.Show 1
End Sub

Private Sub mnu_RptLAB05_Click()
    'frPrintLAB05.Show 1
End Sub

Private Sub mnu_RptORD02_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmPrintORD02", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    frmPrintORD02.Show 1
End Sub

Private Sub mnu_RptORD03_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmPrintORD03", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    frmPrintORD03.Show 1
End Sub

Private Sub mnu_RptOrd04_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmPrintORD04", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    frmPrintORD04.Show 1
End Sub

Private Sub mnu_RptOrd05_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmPrintORD05", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    frmPrintORD05.Show 1
End Sub

Private Sub mnu_RptQUO02_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmPrintQU002", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    frmPrintQU002.Show 1
End Sub

Private Sub mnu_RptQuo03_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmPrintQUO03", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    frmPrintQUO03.Show 1
End Sub

Private Sub mnu_RptQuo04_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmPrintQU004", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    frmPrintQU004.Show 1
End Sub

Private Sub mnu_RptTRN01_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole(frmPrintTRN01.Name, AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    frmPrintTRN01.Show 1
End Sub

Private Sub mnu_RptTRN02_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole(frmPrintTRN02.Name, AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    frmPrintTRN02.Show 1
End Sub

Private Sub mnu_RptTRN03_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole(frmPrintTRN03.Name, AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    frmPrintTRN03.Show 1
End Sub

Private Sub mnu_RptTRN03A_Click()
    frmPrintTRN03A.Show 1
End Sub


Private Sub mnu_RptTRN04_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole(frmPrintTRN04.Name, AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    frmPrintTRN04.Show 1
End Sub

Private Sub mnu_RPTTRN05_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole(frmPrintTRN05.Name, AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    frmPrintTRN05.Show 1
End Sub

Private Sub mnu_RPtTRN06_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole(frmPrintTRN06.Name, AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    frmPrintTRN06.Show 1
End Sub

Private Sub mnu_RptTRN07_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole(frmPrintTRN07.Name, AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    frmPrintTRN07.Show 1
End Sub

Private Sub mnu_RptTRN08_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole(frmPrintTRN08.Name, AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    frmPrintTRN08.Show 1
End Sub

Private Sub mnu_RptTRN09_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole(frmPrintTRN09.Name, AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    frmPrintTRN09.Show 1
End Sub

Private Sub mnu_RptTRN10_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole(frmPrintTRN10.Name, AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    frmPrintTRN10.Show 1
End Sub

Private Sub mnu_RptTRN10A_Click()
    FrmPrintTRN10A.Show 1
End Sub

Private Sub mnu_RptTRN10A1_Click()
    FrmPrintTRN10A1.Show 1
End Sub

Private Sub mnu_RptTRN11_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole(frmPrintTRN11.Name, AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    frmPrintTRN11.Show 1
End Sub

Private Sub mnu_RptTRN12_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole(frmPrintTRN12.Name, AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    
    Dim UserGrp As New AWSUserGroup
    Dim isMD As Boolean, i As Integer
    If Not CurrentUser Is Nothing Then
        For i = 1 To CurrentUser.getUserGroups.Count
            Set UserGrp = CurrentUser.getUserGroups(i)
            If Trim(UserGrp.Name) = "MD" Then
                isMD = True
                Exit For
            Else
                isMD = False
            End If
        Next i
    End If
    frmPrintTRN12.isMD = isMD
    frmPrintTRN12.Show 1
End Sub

Private Sub mnu_RptTRN13_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole(frmPrintTRN13.Name, AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    
    Dim UserGrp As New AWSUserGroup
    Dim isMD As Boolean, i As Integer
    If Not CurrentUser Is Nothing Then
        For i = 1 To CurrentUser.getUserGroups.Count
            Set UserGrp = CurrentUser.getUserGroups(i)
            If Trim(UserGrp.Name) = "MD" Then
                isMD = True
                Exit For
            Else
                isMD = False
            End If
        Next i
    End If
    
    frmPrintTRN13.isMD = isMD
    frmPrintTRN13.Show 1
End Sub

Private Sub mnu_SaleDistrict_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmSaleDistrict", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    frmSaleDistrict.Show 1
End Sub

Private Sub mnu_SalePaidRate_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmComRateSale", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    frmComRateSale.Show 1
End Sub

Private Sub mnu_SaleTarget_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmSaleTarget", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    frmSaleTarget.Show 1
End Sub

Private Sub mnu_SaleTarget_List_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmSaleTarget", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    VisibleImage False
    frmSalesTarget.Show
End Sub


Private Sub mnu_SaleZone_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmSaleZone", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    frmSaleZone.Show 1
End Sub

Private Sub mnu_Sample_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmSendSample", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    frmSendSample.Show 1
End Sub

Private Sub mnu_SamplingPoint_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmSamplingPoint", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    frmSamplingPoint.Show 1
End Sub

Private Sub mnu_Schedule_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmJobCalendar", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    frmJobCalendar.Show 1
End Sub

Private Sub mnu_SubTruckType_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmTruckSubType", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    frmTruckSubType.Show 1
End Sub

Private Sub mnu_TaxInvoice_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmReceipt", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    frmReceipt.Show 1
End Sub

Private Sub mnu_TestData_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmLABResult", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    frmLABResult.Show 1
End Sub

Private Sub mnu_TestForm_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmLABTemplate", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    frmLABTemplate.Show 1

End Sub

Private Sub mnu_TestMethod_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmLABMethod", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    frmLABMethod.Show 1
End Sub

Private Sub mnu_TestQue_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmWasteLABQue", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    VisibleImage False
    frmWasteLABQue.Show
    frmWasteLABQue.ZOrder
End Sub

Private Sub mnu_TranCustomerRate_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmTransportRate", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    VisibleImage False
    frmTransportChargRate.TransType "vw_TransportCustChargeRate"
    frmTransportChargRate.VSFlexGrid1.Clear
    frmTransportChargRate.Caption = "ตารางค่าขนส่งที่เก็บกับลูกค้า"
    frmTransportChargRate.Show
    frmTransportChargRate.ZOrder
End Sub

Private Sub mnu_TransDocCheck_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmMenifest", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    frmMenifest.Show 1
End Sub

Private Sub mnu_TranVendorRate_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmTransportRate", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    VisibleImage False
    frmTransportChargRate.TransType "vw_TransportVendorChargeRate"
    frmTransportChargRate.VSFlexGrid1.Clear
    frmTransportChargRate.Caption = "ตารางค่าขนส่งที่จ่าย Vendor"
    frmTransportChargRate.Show
    frmTransportChargRate.ZOrder
End Sub

Private Sub mnu_TranVendorFee_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmTransportRate", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    frmTransportRate.Show 1
End Sub

Private Sub mnu_TreatmentData_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmTreatmentData", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    frmTreatmentData.Show 1
End Sub

Private Sub mnu_TripControlBox_Click()
    Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmTransportRate", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    VisibleImage False
    frmJobForBoxAssignList.Show
    frmJobForBoxAssignList.ZOrder
End Sub

Private Sub mnu_TripControlCarBoxAssign_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmJobForCarAssignList", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    VisibleImage False
    frmJobCarBoxAssign.Show 1
End Sub

Private Sub mnu_TripControlCarBoxStatus_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmBoxControlStatusList", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    VisibleImage False
    frmBoxControlStatusList.Show
    frmBoxControlStatusList.ZOrder
End Sub

Private Sub MnuCBaseChange_Click()
    FrmCBaseChange.Show 1
End Sub

Private Sub MnuCarDaily_Click()
    FrmCarDaily.Show 1
End Sub

Private Sub MnuConAdd_Click()
    FrmRptCEdit.Show 1
End Sub

Private Sub MnuConTake_Click()
    FrmRptConTake.Show 1
End Sub

Private Sub mnuCustBase_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmTruckType", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    FrmCBaseChange.Show 1
End Sub

Private Sub MnuCustIn_Click()
    FrmCustCalendar.Show 1
End Sub

Private Sub MnuCustVisit_Click()
    FrmRptCustVisit.Show 1
End Sub

Private Sub MNUFIN13A_Click()
    frmPrintFIN13_A.Show 1
End Sub
Private Sub mnuJobProject_Click()
    FrmJobProject.Show
End Sub

Private Sub mnuOPCost_Click()
    FrmBurnCost.Show
End Sub

Private Sub MnuRptCovidbin_Click()
FrmCRBigbag.Label1(1).Caption = "รายงานถังขยะโควิด"
FrmCRBigbag.Frame1.Caption = "รายงานถังขยะโควิด"
FrmCRBigbag.Caption = "รายงานถังขยะโควิด"
FrmCRBigbag.Show 1
End Sub

Private Sub MnuRptQuotation_Click()
 FrmPrintCR06.Show 1
End Sub

Private Sub MnuTRN10EX_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole(FrmPrintTrn10EX.Name, AccessRole)
'    If Trim(tmpStr) <> "" Then
'        MsgBox tmpStr, vbExclamation
'        Exit Sub
'    End If
    FrmPrintTrn10EX.Show 1
End Sub

Private Sub Mu_trn14_Click()
    frmPrintTRN14.Show 1
End Sub
Private Sub mnu_TruckManage_Click()
    frmTruckManage.Show 1
End Sub

Private Sub mnu_TruckType_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmTruckType", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    frmTruckType.Show 1
End Sub

Private Sub mnu_TSDFCompany_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmTSDF", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    frmTSDF.Show 1
End Sub

Private Sub mnu_Unit_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmUnit", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    frmUnit.Show 1
End Sub

Private Sub mnu_UpTF_Click()
Dim tmpCol As Collection, tmpS As New BWGSearchManager
Dim x As Long
    Set tmpCol = tmpS.JobDataTimeTableSearch("TripTreatFee=-999")
    pgb1.Min = 0
    pgb1.Max = tmpCol.Count + 1
    pgb1.Visible = True
    For x = 1 To tmpCol.Count
        DBConnExc "Update tbJobDataTimeTable Set TripTreatFee=" & CalCustTreatFee(tmpCol(x)) & " WHERE JobDataCarID='" & tmpCol(x).JobDataCarID & "' And TripNo=" & tmpCol(x).TripNo
        DoEvents
        pgb1.Value = x
        Me.Stb1.Panels(1).Text = x & "/" & pgb1.Max
        'pgb1.ZOrder
    Next
    pgb1.Visible = False
    MsgBox "Done"
End Sub

Private Sub mnu_User_Click()
On Error Resume Next
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmUserProfile", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    frmUserProfile.Show 1
End Sub

Private Sub mnu_UserGroup_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmUserGroup", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    frmUserGroup.Show 1
End Sub

Private Sub mnu_WasteQue_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmWasteDataQue", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    VisibleImage False
    frmWasteDataQue.ZOrder
    frmWasteDataQue.Show
End Sub

Private Sub mnu_WasteStable_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmSendSampleStable", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    frmSendSampleStable.Show 1
End Sub

Private Sub mnu_WhatNew_Click()
    frmWhatNew.Show
End Sub

Private Sub mnu_Zone_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmZoneArea", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    frmZoneArea.Show 1
End Sub

Private Sub Mu_Trn15_Click()
    frmPrintTRN15.Show
End Sub
Private Sub Mu_Trn16_Click()
       frmPrintTRN16.Show
End Sub


Private Sub ORD06_Click()
       FrmRptord06.Show 1
End Sub



Private Sub rptEnv05_Click()
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmPrintENV05", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    frmPrintENV05.Show 1
End Sub

Public Sub VisibleImage(ByVal bool As Boolean)
    Picture1.Visible = bool
End Sub


Private Sub Timer1_Timer()
On Error Resume Next
    Dim oFS As FileSystemObject
    Dim PicFile As String
    Dim ImageLst() As String
    Set oFS = New FileSystemObject
    
    If ImageTime = 0 Then ImageTime = 1
    IntervalTime = IntervalTime + Timer1.Interval
    If IntervalTime / 60000 = ImageTime Then
        IntervalTime = 0
        ImageLst = ListFiles(App.Path & "\Image")
        If UBound(ImageLst) <= 0 Then Exit Sub
        ImageIndex = ImageIndex + 1
        If ImageIndex > UBound(ImageLst) Then ImageIndex = 1
        PicFile = ImageLst(ImageIndex)
        If oFS.FileExists(PicFile) = True Then
            Image1.Picture = LoadPicture(PicFile)
            Image1.Stretch = True
            DoEvents
        End If
    End If
End Sub

Function ListFiles(ByVal Path As String, Optional ByVal NestedDirs As Boolean) _
    As String()
    Dim FSO As New Scripting.FileSystemObject
    Dim fld As Scripting.Folder
    Dim fileList As String
    
    ' get the starting folder
    Set fld = FSO.GetFolder(Path)
    ' let the private subroutine do all the work
    fileList = ListFilesPriv(fld, NestedDirs)
    ' convert to a string array
    ' (the first element will be a null string)
    ListFiles = Split(fileList, ";")
    
End Function

' private procedure that returns a file list
' as a comma-delimited list of files
Function ListFilesPriv(ByVal fld As Scripting.Folder, _
    ByVal NestedDirs As Boolean) As String
    Dim fil As Scripting.File
    Dim subfld As Scripting.Folder
    
    ' list all the files in this directory
    For Each fil In fld.Files
        If fil.Type = "JPG File" Then
            ListFilesPriv = ListFilesPriv & ";" & fil.Path
        End If
    Next
        
    ' if requested, search also subdirectories
    If NestedDirs Then
        For Each subfld In fld.SubFolders
            ListFilesPriv = ListFilesPriv & ListFilesPriv(subfld, NestedDirs)
        Next
    End If
    
End Function

Private Sub TRN17_Click()
frnPrintTRN17.Show 1
End Sub
