VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmUpdateTripControl 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Update Incollect Fee"
   ClientHeight    =   2595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9795
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2595
   ScaleWidth      =   9795
   StartUpPosition =   2  'CenterScreen
   Begin prjBWG.ctlDate cdFrom 
      Height          =   315
      Left            =   120
      TabIndex        =   7
      Top             =   2160
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   556
      BackColor       =   -2147483633
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   375
      Left            =   6960
      TabIndex        =   3
      Top             =   2160
      Width           =   1275
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   8400
      TabIndex        =   2
      Top             =   2160
      Width           =   1275
   End
   Begin VB.Frame Frame1 
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   9555
      Begin MSComctlLib.ProgressBar PGBar 
         Height          =   255
         Left            =   420
         TabIndex        =   1
         Top             =   420
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Height          =   195
         Left            =   4140
         TabIndex        =   5
         Top             =   660
         Width           =   1395
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Height          =   195
         Left            =   420
         TabIndex        =   4
         Top             =   900
         Width           =   8775
      End
   End
   Begin prjBWG.ctlDate cdTo 
      Height          =   315
      Left            =   2640
      TabIndex        =   8
      Top             =   2160
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   556
      BackColor       =   -2147483633
   End
   Begin VB.Label Label3 
      Caption         =   "TO"
      Height          =   255
      Left            =   2160
      TabIndex        =   6
      Top             =   2220
      Width           =   255
   End
End
Attribute VB_Name = "frmUpdateTripControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdStart_Click()
On Error GoTo ErrD
Dim RS As ADODB.Recordset
Dim RSUpdate As ADODB.Recordset
Dim tSql As String

Dim tmpCol As Collection, tmpS As New BWGSearchManager
Dim tmpTime As BWGJobDataTimeTable
Dim aumphurID1 As String, x%


    cmdStart.Enabled = False
    PGBar.Visible = True
    Label1.Caption = "Executing update incollect fee."
    tSql = "SELECT * From vw_MenifestQue" & vbCrLf
    tSql = tSql & "WHERE WorkDate >= '" & cdFrom.ValueYMD & "' AND WorkDate <= '" & cdTo.ValueYMD & "'" & vbCrLf
    tSql = tSql & "ORDER BY DocNo"
    Set RS = New ADODB.Recordset
    DoEvents
    RS.Open tSql, DBManager.GetDBConnection, adOpenKeyset, adLockOptimistic
    If Not RS.EOF Then
        PGBar.Max = RS.RecordCount
        Label1.Caption = "Updating Incollect Fee in " & RS.RecordCount & " records."
    End If
    PGBar.Max = RS.RecordCount
    Label2.Caption = "0%"
    Do While Not RS.EOF
        DoEvents
        PGBar.Value = PGBar.Value + 1
        Label2.Caption = Format(PGBar.Value / PGBar.Max * 100, "0.00") & "%"
        
        
        '***** Search TimeTableNo
        Set tmpCol = tmpS.JobDataTimeTableQuickSearch("TimeTableNo='" & RS!TimeTableNo & "' And isMain='Y' And isCancel ='N'")
        If tmpCol.Count > 0 Then
            Label1.Caption = "Updating Manifest No# " & RS!docNo & "  "
            DoEvents
            Set tmpTime = tmpCol(1)
            Load frTripControl
            With frTripControl
                aumphurID1 = Trim(RS!CustAumphurID)
                .SetTripControl tmpTime, RS!CustomerName, Trim(RS!SubTypeDesc) & ", " & Trim(RS!TruckTypeDesc), RS!CarRegisID, aumphurID1, RS!TSDFAumphurID, RS!TruckTypeID
                If .Save = True Then
                    Label1.Caption = Label1.Caption & "Can update."
                Else
                    Label1.Caption = Label1.Caption & "Can not update."
                End If
                DoEvents
'                .Show 1
            End With
        End If
        
        
        
        
        RS.MoveNext
    Loop
    'Label2.Caption = "0%"
    Label1.Caption = "Update completed!!"
    cmdStart.Enabled = True
    PGBar.Visible = False
    PGBar.Value = 0
    Set RS = Nothing
    
    Exit Sub
ErrD:
    MsgBox Err.Description, vbCritical
    cmdStart.Enabled = True
    Set RS = Nothing
    Exit Sub
End Sub

Private Sub Form_Load()
    PGBar.Value = 0
    Label1.Caption = "Click Start button to start execute data"
    cdFrom.DayValue = "01"
End Sub

