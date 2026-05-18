VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "Commission"
   ClientHeight    =   7185
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   9540
   Icon            =   "frmMDI.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar Stb1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   0
      Top             =   6780
      Width           =   9540
      _ExtentX        =   16828
      _ExtentY        =   714
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   7
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2461
            MinWidth        =   2469
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
            Object.Width           =   4057
            MinWidth        =   4057
            Text            =   "Regional Setting"
            TextSave        =   "Regional Setting"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
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
   Begin VB.Menu MNU_IntCom 
      Caption         =   "ค่าคอมฯ ใน"
      Begin VB.Menu MNU_IntCom1 
         Caption         =   "การตั้งค่าคอมฯ ใน"
      End
      Begin VB.Menu MNU_IntCom2 
         Caption         =   "การตัดจ่ายค่าคอมฯ ใน"
      End
      Begin VB.Menu MNU_Dat1 
         Caption         =   "-"
      End
      Begin VB.Menu MNU_Compay 
         Caption         =   "ประมวลผลค่าคอมฯ ใน"
      End
   End
   Begin VB.Menu MNU_ExtCom 
      Caption         =   "ค่าคอมฯ นอก"
      Begin VB.Menu MNU_ExtCom1 
         Caption         =   "การตั้งค่าคอมฯ นอก"
         Visible         =   0   'False
      End
      Begin VB.Menu MNU_ExtComm_CustomerList 
         Caption         =   "รายชื่อลูกค้าที่ต้องจ่ายค่าคอมฯ นอก"
      End
      Begin VB.Menu MNU_ExtCom3 
         Caption         =   "ประมวลผลค่าคอมฯ นอก"
      End
   End
   Begin VB.Menu MNU_Menifest 
      Caption         =   "รายการ Menifest"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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
    Me.Caption = App.Title & "   " & App.Major & "." & App.Minor & "." & App.Revision
End Sub

Public Sub InitializeMenu()
    '
End Sub

Private Sub MDIForm_Load()
    Me.Caption = Me.Caption & "  " & App.Major & "." & App.Minor & "." & App.Revision
    frmComPay.Show
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'    Localization.SetDefaultRegionalSettings DefaultLCID
End Sub

Private Sub MDIForm_Resize()
    On Error Resume Next
    Dim TotalWidth As Long
    TotalWidth = Stb1.Width - (Stb1.Panels(4).Width + Stb1.Panels(5).Width + Stb1.Panels(5).Width + Stb1.Panels(7).Width)
    Stb1.Panels(1).Width = (70 / 100 * TotalWidth)
    Stb1.Panels(2).Width = (15 / 100 * TotalWidth)
    Stb1.Panels(3).Width = (15 / 100 * TotalWidth)
    Stb1.Panels(4).Width = 2300
    Stb1.Panels(5).Width = 500
    Stb1.Panels(6).Width = 500
    Stb1.Panels(7).Width = 500
    Stb1.Refresh
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    Dim frm As Form
    For Each frm In Forms
        Unload frm
    Next
End Sub

Private Sub MNU_Compay_Click()
    frmComPay.Show
    frmComPay.ZOrder
End Sub

Private Sub MNU_ExtCom1_Click()
    frmComExt1.Show 1
End Sub

Private Sub MNU_ExtCom3_Click()
    frmExtCom2.Show
    frmExtCom2.ZOrder
End Sub

Private Sub MNU_ExtComm_CustomerList_Click()
    frmCustomerExtComm.Show
    frmCustomerExtComm.ZOrder
End Sub

Private Sub MNU_IntCom1_Click()
    frmComInt1.Show 1
End Sub

Private Sub MNU_IntCom2_Click()
    frmComInt2.Show 1
End Sub

Private Sub MNU_Menifest_Click()
    frListMenifest.Show
    frListMenifest.ZOrder
End Sub
