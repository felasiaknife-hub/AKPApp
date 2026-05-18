VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmInitialLogin 
   Caption         =   "Initial Login System"
   ClientHeight    =   6600
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   11730
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   6600
   ScaleWidth      =   11730
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   975
      Left            =   -840
      TabIndex        =   3
      Top             =   5700
      Width           =   12735
      Begin VB.CommandButton cmdExit 
         Caption         =   "Exit"
         Height          =   495
         Left            =   11160
         TabIndex        =   7
         Top             =   300
         Width           =   1275
      End
      Begin VB.CommandButton cmdEncrypt 
         Caption         =   "Encrypt-MD5"
         Height          =   495
         Left            =   9720
         TabIndex        =   6
         Top             =   300
         Width           =   1275
      End
      Begin VB.CommandButton cmdDecrypt 
         Caption         =   "Decrypt"
         Height          =   495
         Left            =   8280
         TabIndex        =   5
         Top             =   300
         Width           =   1275
      End
      Begin VB.CommandButton cmdLoad 
         Caption         =   "Load All User"
         Height          =   495
         Left            =   6840
         TabIndex        =   4
         Top             =   300
         Width           =   1275
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   1275
      Left            =   -540
      TabIndex        =   1
      Top             =   -180
      Width           =   12435
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   315
         Left            =   1320
         TabIndex        =   8
         Top             =   780
         Visible         =   0   'False
         Width           =   9975
         _ExtentX        =   17595
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label Label1 
         Caption         =   "Security System Login is progressing for new encrypt password."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   840
         TabIndex        =   2
         Top             =   420
         Width           =   7095
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid vsGridUser 
      Height          =   4515
      Left            =   0
      TabIndex        =   0
      Top             =   1200
      Width           =   11715
      _cx             =   20664
      _cy             =   7964
      _ConvInfo       =   1
      Appearance      =   1
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
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
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
End
Attribute VB_Name = "frmInitialLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ObjCrypt As New RC4Crypt
Dim ObjMD5Crypt As MD5Crypt.EncryptionTools

Private Sub cmdDecrypt_Click()
    Dim DeCryptPwd As String, i As Long
    Set ObjCrypt = New RC4Crypt
    With vsGridUser
        ProgressBar1.Max = vsGridUser.Rows - 1
        ProgressBar1.Min = 0
        ProgressBar1.Value = 0
        ProgressBar1.Visible = True
        For i = 1 To vsGridUser.Rows - 1
            DeCryptPwd = ObjCrypt.DecryptString(.TextMatrix(i, 5), ENC_KEY)
            .TextMatrix(i, 7) = Trim(DeCryptPwd)
            ProgressBar1.Value = i
            DoEvents
        Next i
    End With
    ProgressBar1.Visible = False
End Sub

Private Sub cmdEncrypt_Click()
    Dim EnCryptPwd As String, i As Long
    Dim RS As ADODB.Recordset
    Dim tSql As String
    Set ObjCrypt = New RC4Crypt
    Set ObjMD5Crypt = New MD5Crypt.EncryptionTools
    With vsGridUser
        ProgressBar1.Max = vsGridUser.Rows - 1
        ProgressBar1.Min = 0
        ProgressBar1.Value = 0
        ProgressBar1.Visible = True
        For i = 1 To vsGridUser.Rows - 1
            If Trim(.TextMatrix(i, 7)) <> "" Then
                EnCryptPwd = ObjMD5Crypt.MD5HashString(Trim(.TextMatrix(i, 7)))
                'EnCryptPwd = ObjCrypt.Encrypt(Trim(.TextMatrix(i, 7)), txtData.Text)
                'EnCryptPwd = ObjCrypt.EncryptString(Trim(.TextMatrix(i, 7)), ENC_KEY, True)
                .TextMatrix(i, 6) = Trim(EnCryptPwd)
                ProgressBar1.Value = i
            End If
            If Trim(EnCryptPwd) <> "" Then
                tSql = "UPDATE tbCompanyStaff SET LoginPwdWeb = '" & Trim(.TextMatrix(i, 7)) & "', MD5Pwd = '" & EnCryptPwd & "' WHERE StaffID = '" & .TextMatrix(i, 1) & "'"
            Else
                tSql = "UPDATE tbCompanyStaff SET LoginPwdWeb = '" & Trim(.TextMatrix(i, 7)) & "' WHERE StaffID = '" & .TextMatrix(i, 1) & "'"
            End If
            DBConn.Execute tSql
            Label1.Caption = "Encrypting MD5 password for " & .TextMatrix(i, 2) & " " & .TextMatrix(i, 3)
            DoEvents
        Next i
    End With
    ProgressBar1.Visible = False
    Label1.Caption = "Security System Login are encrypting to MD5  password completed."
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdLoad_Click()
    On Error GoTo ErrHandler
    Dim tSql As String
    Dim i As Long
    Dim RS As ADODB.Recordset
    vsGridUser.Clear
    SetGrid
    tSql = "SELECT StaffID, StaffFName, StaffLName, LoginName, LoginPwd, isActive" & vbCrLf
    tSql = tSql & "FROM         dbo.tbCompanyStaff" & vbCrLf
    tSql = tSql & "ORDER BY StaffFName, StaffLName"
    Set RS = GetRS(tSql, adOpenKeyset, adLockOptimistic)
    Do While Not RS.EOF
        i = i + 1
        vsGridUser.Rows = RS.RecordCount + 1
        With vsGridUser
            .TextMatrix(i, 0) = i
            .TextMatrix(i, 1) = RS!StaffID
            .TextMatrix(i, 2) = RS!StaffFName
            .TextMatrix(i, 3) = RS!StaffLName
            .TextMatrix(i, 4) = RS!LoginName
            .TextMatrix(i, 5) = RS!LoginPwd
            .TextMatrix(i, 6) = ""
            .TextMatrix(i, 7) = ""
        End With
        RS.MoveNext
    Loop
    RS.Close
    Set RS = Nothing
    MsgBox "Load all user completed."
    Exit Sub
ErrHandler:
    MsgBox err.Description
    Exit Sub
End Sub

Private Sub Form_Load()
    SetGrid
End Sub

Private Sub SetGrid()
    With vsGridUser
        .Cols = 8
        .TextMatrix(0, 0) = "No"
        .TextMatrix(0, 1) = ""
        .TextMatrix(0, 2) = "F-Name"
        .TextMatrix(0, 3) = "L-Name"
        .TextMatrix(0, 4) = "U-Name"
        .TextMatrix(0, 5) = "EncryptPwd"
        .TextMatrix(0, 6) = "EncryptPwd-MD5"
        .TextMatrix(0, 7) = "DecryptPwd"
        .ColWidth(0) = 500
        .ColWidth(1) = 0
        .ColWidth(2) = 2000
        .ColWidth(3) = 1500
        .ColWidth(4) = 1500
        .ColWidth(5) = 1800
        .ColWidth(6) = 2500
        .ColWidth(7) = 1500
        .ColAlignment(0) = flexAlignCenterCenter
        .ColAlignment(1) = flexAlignLeftCenter
        .ColAlignment(2) = flexAlignLeftCenter
        .ColAlignment(3) = flexAlignLeftCenter
        .ColAlignment(4) = flexAlignLeftCenter
        .ColAlignment(5) = flexAlignLeftCenter
        .ColAlignment(6) = flexAlignLeftCenter
        .ColAlignment(7) = flexAlignLeftCenter
    End With
End Sub
