VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmMain 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Preliminary Industrial - Waste Data Form"
   ClientHeight    =   7965
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   14175
   Icon            =   "MDIMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar Stb1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   7590
      Width           =   14175
      _ExtentX        =   25003
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            Bevel           =   0
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnu_Customer 
      Caption         =   "ข้อมูลลูกค้า"
   End
   Begin VB.Menu mnu_CustomerQue 
      Caption         =   "ลูกค้าที่รอการตรวจสอบ"
   End
   Begin VB.Menu mnu_WasteDataQue 
      Caption         =   "กากของเสียที่รอการตรวจรับ"
   End
   Begin VB.Menu mnu_Quotation 
      Caption         =   "Quotation"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CKuser_ As New CKuser

Public Sub InitializeMenu()
    Call WasteType
End Sub

Public Sub ActivateForm()
    '
End Sub

Private Sub MDIForm_Load()
Dim Appv As String
Dim Appn As String
Dim tmpRec As ADODB.Recordset
Set tmpRec = New ADODB.Recordset
Dim tmpCK As ADODB.Recordset
Set tmpCK = New ADODB.Recordset
Dim sqlStr As String
    sqlStr = "Select vApp From ApplicationTB Where ApplicationName = '0AWasteDataFrom' "
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec.EOF = False Then
    Appv = tmpRec(0)
    End If
    'Me.Caption = Me.Caption & "    V." & App.Major & "." & App.Minor & "." & App.Revision
    Me.Caption = "AKP" & " " & App.Title & "   " & App.Major & "." & App.Minor & "." & App.Revision & "  Server : " & strSvr & "  Running on " & strDb
    Appn = App.Major & "." & App.Minor & "." & App.Revision
    If Appn <> Appv Then
        MsgBox ("โปรแกรมที่ท่านใช้อยู่  v. " & Appn & "ไม่ใช่ Version ปัจจุบัน ท่านต้องทำการ Update Program ก่อนครับ ปัจจุบันคือ Version " & Appv & " ")
        End
    End If
    Stb1.Panels(2).Width = Stb1.Width - Stb1.Panels(1).Width
    Stb1.Panels(2).Text = "ชื่อผู้เข้าใช้งาน : " & CurrentUser.StaffFName & "  " & CurrentUser.StaffLName
    tmpRec.Close
    Set tmpRec = Nothing
    'ตรวจสอบ บริษัท
     sqlStr = "Select  * From tbStaffInCompany Where (StaffID = '" & CurrentUser.ID & "') And (CompanyID = '776EDA5830F744A6B3551348D851BC22')"
     Set tmpCK = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    Dim i As Integer
    i = tmpCK.RecordCount
    If i = 0 Then
        isAkp = False
    ElseIf i > 0 Then
        isAkp = True
    End If
    tmpCK.Close
    Set tmpCK = Nothing
    
     sqlStr = "Select  * From tbStaffInCompany Where (StaffID = '" & CurrentUser.ID & "') And (CompanyID = '89F23B57E7DC4407B95953A38AC5B537')"
     Set tmpCK = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    i = tmpCK.RecordCount
    If i = 0 Then
        isBWG = False
    ElseIf i > 0 Then
        isBWG = True
    End If
    tmpCK.Close
    Set tmpCK = Nothing
    
     sqlStr = "Select  * From tbStaffInCompany Where (StaffID = '" & CurrentUser.ID & "') And (CompanyID = '3F764F1713974948A42CC0FAF926C730')"
     Set tmpCK = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    i = tmpCK.RecordCount
    If i = 0 Then
        isPST = False
    ElseIf i > 0 Then
        isPST = True
    End If
    tmpCK.Close
    Set tmpCK = Nothing

    
     sqlStr = "Select  * From tbStaffInCompany Where (StaffID = '" & CurrentUser.ID & "') And (CompanyID = 'C451AF07C9A44B4282D6E08B14AC0195')"
     Set tmpCK = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    i = tmpCK.RecordCount
    If i = 0 Then
        isBWC = False
    ElseIf i > 0 Then
        isBWC = True
    End If
    tmpCK.Close
    Set tmpCK = Nothing
        '-----------------------------------------
  frmWasteDataQue.Show
End Sub
Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim frm As Form
    For Each frm In Forms
        Unload frm
    Next
End Sub
Private Sub MDIForm_Resize()
On Error Resume Next
    Stb1.Panels(2).Width = Stb1.Width - Stb1.Panels(1).Width
    Stb1.Panels(2).Text = "ชื่อผู้เข้าใช้งาน : " & CurrentUser.StaffFName & "  " & CurrentUser.StaffLName
End Sub
Private Sub mnu_Customer_Click()
    frmCustomer.Show 1
End Sub
Private Sub mnu_CustomerQue_Click()
    frmCustomerList.Show
End Sub
'Private Sub mnu_Quotation_Click()
'    frQuotation.Show 1
'End Sub
Private Sub mnu_WasteDataQue_Click()
    frmWasteDataQue.Show
End Sub
