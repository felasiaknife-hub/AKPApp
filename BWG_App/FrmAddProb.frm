VERSION 5.00
Begin VB.Form FrmAddProb 
   Caption         =   "เพิ่มข้อร้องเรียน"
   ClientHeight    =   3570
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5535
   LinkTopic       =   "Form1"
   ScaleHeight     =   3570
   ScaleWidth      =   5535
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton BtSave 
      Caption         =   "บันทึก"
      Height          =   600
      Left            =   4320
      TabIndex        =   5
      Top             =   2880
      Width           =   1050
   End
   Begin VB.TextBox TxtDet 
      Height          =   1500
      Left            =   1170
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   1260
      Width           =   4200
   End
   Begin VB.TextBox TxtName 
      Height          =   285
      Left            =   1170
      TabIndex        =   2
      Top             =   765
      Width           =   4200
   End
   Begin WasteManagment.ctlDate dtdate 
      Height          =   315
      Left            =   1170
      TabIndex        =   6
      Top             =   270
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   556
      BackColor       =   -2147483633
   End
   Begin VB.Label ProbID 
      Height          =   240
      Left            =   4635
      TabIndex        =   7
      Top             =   180
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "ข้อร้องเรียน :"
      Height          =   195
      Left            =   180
      TabIndex        =   3
      Top             =   1305
      Width           =   915
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "เรื่อง :"
      Height          =   195
      Left            =   675
      TabIndex        =   1
      Top             =   810
      Width           =   420
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "วันที่ : "
      Height          =   195
      Left            =   675
      TabIndex        =   0
      Top             =   315
      Width           =   450
   End
End
Attribute VB_Name = "FrmAddProb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub BtSave_Click()
    On Error Resume Next
    Dim RS As ADODB.Recordset
    Dim RSC As ADODB.Recordset
    Dim tSql As String
    Dim tmpList
    Dim i As Integer
    If Trim(TxtName.Text) = "" Then
        MsgBox "ระบุชื่อเรื่องก่อน", vbInformation + vbOKOnly, "ผิดพลาด"
        Exit Sub
    End If
    If Trim(TxtDet.Text) = "" Then
        MsgBox "ระบุรายละเอียดก่อน", vbInformation + vbOKOnly, "ผิดพลาด"
        Exit Sub
    End If
    If ProbID.Caption = "" Then
            tSql = "INSERT INTO tbprob(CustID, ProbName,ProbDet,ProbDate)" & vbCrLf
        tSql = tSql & "VALUES ('" & frmCustomer.txtCustCode.Text & "', '" & Trim(TxtName.Text) & "', '" & Trim(TxtDet.Text) & "', '" & dtdate.ValueYMD & "')" & vbCrLf
        DBConnExc tSql
    Else
        Dim Sql As String
        Sql = "update tbprob SET ProbName =  '" & Trim(TxtName.Text) & "',ProbDet='" & Trim(TxtDet.Text) & "',ProbDate='" & dtdate.ValueYMD & "' " & vbCrLf
        Sql = Sql & " WHERE ID = " & Trim(ProbID.Caption) & ""
        DBConnExc Sql
    End If
        frmCustomer.SetProb
            Unload Me
End Sub

