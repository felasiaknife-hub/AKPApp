VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmQuotationTransfer 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ย้ายใบเสนอราคา"
   ClientHeight    =   6495
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8400
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   8400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "ใช้กับทุก Job"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3360
      TabIndex        =   1
      Top             =   180
      Value           =   1  'Checked
      Width           =   1365
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ยกเลิก"
      Height          =   525
      Left            =   6660
      TabIndex        =   4
      Top             =   5910
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ยืนยันการย้าย"
      Height          =   525
      Left            =   4890
      TabIndex        =   3
      Top             =   5910
      Width           =   1695
   End
   Begin VB.TextBox txtNewQuoNo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1440
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   150
      Width           =   1755
   End
   Begin MSComctlLib.ListView lvResult 
      Height          =   5235
      Left            =   0
      TabIndex        =   2
      Top             =   600
      Width           =   8385
      _ExtentX        =   14790
      _ExtentY        =   9234
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
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Job No"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "ใบเสนอราคาเดิม"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "ใบเสนอราคาใหม่"
         Object.Width           =   3528
      EndProperty
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ใบเสนอราคาใหม่"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   1
      Left            =   150
      TabIndex        =   5
      Top             =   210
      Width           =   1215
   End
End
Attribute VB_Name = "frmQuotationTransfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim colJobDet_ As Collection
Dim selQuo_  As BWGQuotationHeader

Private Sub Command1_Click()
On Error GoTo ErrD
Dim x%, tmpJobDet As BWGJobDataDetail
Dim tmpRec As New ADODB.Recordset, tmpCon As Connection
Dim sqlStr As String
    If MsgBox("ท่านต้องการยืนยันการย้ายใบเสนอราคานี้หรือไม่?", vbQuestion + vbYesNo) = vbYes Then
        
        For x = 1 To lvResult.ListItems.Count
            If Trim(lvResult.ListItems(x).SubItems(2)) <> "" Then
                Set tmpJobDet = colJobDet_(x)
                sqlStr = "SELECT * FROM tbJobDataDetail WHERE WasteDataID not in (Select WasteDataID from tbQuotationDetail WHERE QuotationID='" & selQuo_.ID & "')"
                Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
                If tmpRec Is Nothing Then Exit Sub
                If tmpRec.RecordCount > 0 Then
                    MsgBox "มีกากของเสียบางรายการใน Job ที่เลือกไม่มีอยู่ในใบเสนอราคาเลขที่ " & txtNewQuoNo, vbExclamation
                    tmpRec.Close
                    Set tmpRec = Nothing
                    Exit Sub
                End If
                DBConnExc "Update tbJobDataDetail Set QuotationID='" & tmpJobDet.QuotationID & "',ExtraLaborFee=" & tmpJobDet.QuoLaborCharge & ",TruckTypeID='" & tmpJobDet.QuoTruckTypeID & "',TransportFee=" & tmpJobDet.QuoTransportFee & ",TransferUnitID='" & tmpJobDet.QuoTransUnitID & "',TreatmentRate=" & tmpJobDet.QuoTreatmentRate & ",TreatmentUnitID='" & tmpJobDet.QuoTreatmentUnitID & "',LaborCount=" & tmpJobDet.QuoLaborCount & ",isPriceIncTrans='" & TF_To_YN(tmpJobDet.isPriceIncTrans) & "',MinWeightPerCar=" & tmpJobDet.QuoMinWeightPerCar & ",NonMinRate=" & tmpJobDet.NonMinRate & ",HazMinRate=" & tmpJobDet.HazMinRate & ",NBLMinRate=" & tmpJobDet.NBLMinRate & ",HBLMinRate=" & tmpJobDet.HBLMinRate & " WHERE JobDetID='" & tmpJobDet.ID & "'"
                tmpRec.Close
            End If
        Next

        Set tmpRec = Nothing
        MsgBox "ย้ายใบเสนอราคาเรียบร้อยแล้ว", vbInformation
        Unload Me
    End If
    Exit Sub
ErrD:
    GetMsg "", vbCritical
    Exit Sub
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub txtNewQuoNo_KeyDown(KeyCode As Integer, Shift As Integer)
Dim tmpS As New BWGSearchManager, tmpCol As Collection
Dim x%, tmpJobDet As BWGJobDataDetail
    If KeyCode = 13 Then
        Set tmpCol = tmpS.QuotationHeaderSearch("QuotationNo Like '%" & Replace(txtNewQuoNo.Text, "*", "%") & "%'")
        If tmpCol.Count = 0 Then
            MsgBox "ไม่พบใบเสนอราคาเลขที่ที่ท่านเลือก", vbExclamation
        Else
            If tmpCol.Count = 1 Then
                Set selQuo_ = tmpCol(1)
            Else
                Load frmQuotationSearch
                With frmQuotationSearch
                    .SetCallerForm Me
                    .getCriteria txtNewQuoNo.Text
                    .Show 1
                End With
                If selQuo_ Is Nothing Then
                    MsgBox "กรุณาเลือกใบเสนอราคา", vbExclamation
                    Exit Sub
                Else
                    If Trim(selQuo_.ID) = "" Then
                        MsgBox "กรุณาเลือกใบเสนอราคา", vbExclamation
                        Exit Sub
                    End If
                End If
            End If
            If Check1.Value = 1 Then
                For x = 1 To lvResult.ListItems.Count
                    Set tmpJobDet = colJobDet_(lvResult.ListItems(x).Key)
                    Set tmpCol = tmpS.QuotationDetailSearch("WasteDataID='" & tmpJobDet.wastedataID & "' And QuotationID='" & selQuo_.ID & "'")
                    If tmpCol.Count = 0 Then
                        lvResult.ListItems(x).SubItems(2) = " "
                    Else
                        lvResult.ListItems(x).SubItems(2) = Trim(txtNewQuoNo.Text)
                        tmpJobDet.QuotationID = selQuo_.ID
                        tmpJobDet.QuoLaborCharge = selQuo_.ExtraLaborFee
                        tmpJobDet.QuoTruckTypeID = selQuo_.TruckTypeID
                        tmpJobDet.QuoTransportFee = selQuo_.TransferRate
                        tmpJobDet.QuoTransUnitID = selQuo_.TransferUnitID
                        tmpJobDet.QuoTreatmentRate = tmpCol(1).TreatmentRate
                        tmpJobDet.QuoTreatmentUnitID = tmpCol(1).TreatmentUnitID
                        tmpJobDet.QuoLaborCount = selQuo_.StaffCountPerCar
                        tmpJobDet.isPriceIncTrans = selQuo_.IsPriceIncTransport
                        tmpJobDet.QuoMinWeightPerCar = selQuo_.MinWeightPerCar
                        tmpJobDet.NonMinRate = selQuo_.NonMinRate
                        tmpJobDet.HazMinRate = selQuo_.HazMinRate
                        tmpJobDet.NBLMinRate = selQuo_.NBLMinRate
                        tmpJobDet.HBLMinRate = selQuo_.HBLMinRate
                    End If
                    Set tmpJobDet = Nothing
                Next
            Else
                If Not lvResult.SelectedItem Is Nothing Then
                    Set tmpJobDet = colJobDet_(lvResult.SelectedItem.Key)
                    Set tmpCol = tmpS.QuotationDetailSearch("WasteDataID='" & tmpJobDet.wastedataID & "' And QuotationID='" & selQuo_.ID & "'")
                    If tmpCol.Count = 0 Then
                        lvResult.SelectedItem.SubItems(2) = " "
                    Else
                        lvResult.SelectedItem.SubItems(2) = Trim(txtNewQuoNo.Text)
                        tmpJobDet.QuotationID = selQuo_.ID
                        tmpJobDet.QuoLaborCharge = selQuo_.ExtraLaborFee
                        tmpJobDet.QuoTruckTypeID = selQuo_.TruckTypeID
                        tmpJobDet.QuoTransportFee = selQuo_.TransferRate
                        tmpJobDet.QuoTransUnitID = selQuo_.TransferUnitID
                        tmpJobDet.QuoTreatmentRate = tmpCol(1).TreatmentRate
                        tmpJobDet.QuoTreatmentUnitID = tmpCol(1).TreatmentUnitID
                        tmpJobDet.QuoLaborCount = selQuo_.StaffCountPerCar
                        tmpJobDet.isPriceIncTrans = selQuo_.IsPriceIncTransport
                        tmpJobDet.QuoMinWeightPerCar = selQuo_.MinWeightPerCar
                        tmpJobDet.NonMinRate = selQuo_.NonMinRate
                        tmpJobDet.HazMinRate = selQuo_.HazMinRate
                        tmpJobDet.NBLMinRate = selQuo_.NBLMinRate
                        tmpJobDet.HBLMinRate = selQuo_.HBLMinRate
                    End If
                    Set tmpJobDet = Nothing
                End If
            End If
        End If
        Set tmpCol = Nothing
        Set tmpS = Nothing
    End If
End Sub

Public Sub SetPendingJob(selCol As Collection, oldQuoNo As String)
Dim x%, tmpList As ListItem
Dim tmpJobDet As BWGJobDataDetail
    lvResult.ListItems.Clear
    Set colJobDet_ = selCol
    For x = 1 To selCol.Count
        Set tmpJobDet = selCol(x)
        Set tmpList = lvResult.ListItems.Add(, "'" & Trim(tmpJobDet.ID) & "'", tmpJobDet.JobNo & " - " & tmpJobDet.WasteName)
        tmpList.SubItems(1) = oldQuoNo
        tmpList.SubItems(2) = " "
        Set tmpJobDet = Nothing
    Next
End Sub

Public Sub SetDataQuotation(selQuo As BWGQuotationHeader)
    Set selQuo_ = selQuo
    txtNewQuoNo.Text = selQuo.QuotationNo
End Sub
