VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmIncome 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " "
   ClientHeight    =   5790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6390
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   6390
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3975
      Left            =   120
      TabIndex        =   13
      Top             =   720
      Width           =   5985
      Begin VB.ComboBox cboBmeNew 
         Height          =   315
         Left            =   1140
         TabIndex        =   25
         Text            =   "cboBmeNew"
         Top             =   3000
         Width           =   4455
      End
      Begin VB.ComboBox CobJobtype 
         Height          =   315
         ItemData        =   "frmIncome.frx":0000
         Left            =   1140
         List            =   "frmIncome.frx":000D
         TabIndex        =   22
         Top             =   2640
         Width           =   1995
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "แสดงรายละเอียด"
         Height          =   195
         Left            =   1140
         TabIndex        =   21
         Top             =   3360
         Width           =   2535
      End
      Begin VB.TextBox txtCondition 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         Height          =   315
         Left            =   4680
         TabIndex        =   7
         Top             =   1860
         Width           =   855
      End
      Begin VB.ComboBox cboOper 
         Height          =   315
         ItemData        =   "frmIncome.frx":0036
         Left            =   3360
         List            =   "frmIncome.frx":0038
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1860
         Width           =   1305
      End
      Begin VB.ComboBox cboWasteType 
         Height          =   315
         Left            =   1140
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1860
         Width           =   2205
      End
      Begin VB.ComboBox cboISClosed 
         Height          =   315
         Left            =   1140
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1500
         Width           =   4425
      End
      Begin VB.TextBox txtSale 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1140
         TabIndex        =   8
         Top             =   2250
         Width           =   4425
      End
      Begin VB.ComboBox cboCompany 
         Height          =   315
         Left            =   1140
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   750
         Width           =   4425
      End
      Begin VB.ComboBox cboCarType 
         Height          =   315
         Left            =   1140
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1140
         Width           =   4425
      End
      Begin WasteManagment.ctlDate dtFrom 
         Height          =   315
         Left            =   1140
         TabIndex        =   0
         Top             =   360
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         BackColor       =   -2147483633
      End
      Begin WasteManagment.ctlDate dtTo 
         Height          =   315
         Left            =   3780
         TabIndex        =   1
         Top             =   360
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         BackColor       =   -2147483633
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000005&
         Caption         =   "ลูกค้า BME"
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   3000
         Width           =   855
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000005&
         Caption         =   "ประเภทงานขน"
         Height          =   255
         Left            =   60
         TabIndex        =   23
         Top             =   2700
         Width           =   1035
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ราคาเวสต์"
         Height          =   195
         Index           =   6
         Left            =   300
         TabIndex        =   20
         Top             =   1920
         Width           =   705
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "การปิดงาน"
         Height          =   195
         Index           =   5
         Left            =   255
         TabIndex        =   19
         Top             =   1560
         Width           =   750
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ผู้แทนขาย"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   285
         TabIndex        =   18
         Top             =   2310
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "บริษัท"
         Height          =   195
         Index           =   4
         Left            =   585
         TabIndex        =   17
         Top             =   810
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ประเภทรถ"
         Height          =   195
         Index           =   0
         Left            =   255
         TabIndex        =   16
         Top             =   1200
         Width           =   750
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ช่วงวันที่ขน"
         Height          =   195
         Index           =   3
         Left            =   210
         TabIndex        =   15
         Top             =   420
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ถึงวันที่"
         Height          =   195
         Index           =   2
         Left            =   3120
         TabIndex        =   14
         Top             =   420
         Width           =   495
      End
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "พิมพ์"
      Height          =   795
      Left            =   4110
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4800
      Width           =   915
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "ปิด"
      Height          =   795
      Left            =   5070
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4800
      Width           =   915
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "Preview"
      Height          =   795
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4800
      Width           =   915
   End
   Begin Crystal.CrystalReport rpt1 
      Left            =   5700
      Top             =   60
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin VB.Label LblForm 
      Height          =   285
      Left            =   1215
      TabIndex        =   24
      Top             =   4820
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   4
      Index           =   0
      X1              =   120
      X2              =   5760
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      Index           =   1
      X1              =   210
      X2              =   5910
      Y1              =   570
      Y2              =   570
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "รายงานยอดขาย"
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
      Left            =   120
      TabIndex        =   12
      Top             =   60
      Width           =   1380
   End
   Begin VB.Image Image2 
      Height          =   735
      Index           =   1
      Left            =   0
      Picture         =   "frmIncome.frx":003A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11280
   End
End
Attribute VB_Name = "frmIncome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim curStaff_ As BWGCompanyStaff

Public isMD As Boolean

Private Sub cboOper_Validate(Cancel As Boolean)
    If cboWasteType.ListIndex = 0 Then cboOper.ListIndex = 0
End Sub

Private Sub cboWasteType_Validate(Cancel As Boolean)
    If cboWasteType.ListIndex <= 0 Then
        cboOper.ListIndex = 0
        txtCondition.Text = ""
    End If
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdPreview_Click()
    If cboCarType.ListIndex = -1 Then
        MsgBox "กรุณาระบุประเภทรถ", vbExclamation
        cboCarType.SetFocus
        Exit Sub
    End If
    PrintMain 0
End Sub

Private Sub cmdPrint_Click()
    If cboCarType.ListIndex = -1 Then
        MsgBox "กรุณาระบุประเภทรถ", vbExclamation
        cboCarType.SetFocus
        Exit Sub
    End If
    PrintMain 1
End Sub

Private Sub Form_Load()
    dtFrom.DayValue = "01"
    PopulateCarType cboCarType
    PopulateCompany cboCompany, cboIndex

    With cboISClosed
        .AddItem "==ทุกสถานะ=="
        .AddItem "ปิดงานแล้ว"
        .AddItem "ยังไม่ปิดงาน"
        .ListIndex = 0
    End With
    With cboWasteType
        .AddItem "==ทุกประเภท=="
        .AddItem "('NON', 'NBL')"
        .AddItem "('Haz', 'HBL')"
        .AddItem "('HIC', 'HBI')"
        .AddItem "('NIC', 'NBI')"
        .AddItem "('BBH', 'BBL')"
        .AddItem "('HIC','HBI','NIC','NBI','BBH','BBL')"
        .ListIndex = 0
    End With
    With cboOper
        .AddItem "==ทุกราคา=="
        .AddItem "="
        .AddItem "<"
        .AddItem "<="
        .AddItem ">"
        .AddItem ">="
        .ListIndex = 0
    End With
    With cboBmeNew
        .AddItem "==ทุกประเภท=="
        .AddItem "Commission"
        .AddItem "No Commission"
        .ListIndex = 0
    End With
     With frmPicture
        cmdPreview.Picture = .picPreview.Picture
        cmdPrint.Picture = .picPrint.Picture
        cmdClose.Picture = .picExit.Picture
    End With
    CobJobtype.ListIndex = 0
End Sub

Private Sub PrintMain(printType As Integer)
    Dim tSql As String, tmpSql As String, i As Integer, sCase As String
    Dim tSqlisClose As String, tSqlGroup As String, tSqlOrder As String
    Dim colWasteType As Collection, tSqlView As String
    If WasteTypes Is Nothing Then Call WasteType
    sCase = ""
    For i = 1 To WasteTypes.Count
        sCase = sCase & WasteTypes(i).WasteType & "= CASE NonOrHz WHEN '" & WasteTypes(i).WasteType & "' THEN SUM(ISNULL(IncollectFee, 0)) ELSE 0 END"
        If i < WasteTypes.Count Then sCase = sCase & "," & vbCrLf
    Next i
    sCase = sCase & vbCrLf
    Dim comp As String
    Dim compID As String
'    If cboCompany.ListIndex = 1 Then
'        comp = "AKP"
'        compID = "776EDA5830F744A6B3551348D851BC22"
'    End If
    comp = CurrentUser.WorkCompanies(cboCompany.ListIndex).CompanyCode
    compID = CurrentUser.WorkCompanies(cboCompany.ListIndex).ID
    With rpt1
        .LogonInfo(0) = "DSN = " & DSN_NAME & ";UID = " & strUid & ";PWD = " & strPwd
        If isMD = True Then
            .ReportFileName = App.Path & "\reports\Incomes_MD.rpt"
        Else
            
            If comp = "AKP" Then
'            If cboCompany.Text = "อัคคีปราการ จำกัด (มหาชน)" Then
                .ReportFileName = App.Path & "\reports\Incomes_AKP.rpt"
            Else
                .ReportFileName = App.Path & "\reports\Incomes.rpt"
            End If
        End If
'        tmpSql = "SELECT T1.CompanyNameTH AS CompanyName, ISNULL(T1.ZoneName, 'NoZone') AS ZoneName, ISNULL(T1.SDName, 'NoSD') AS SDName, ISNULL(T1.SaleName, 'NoSale') AS SaleName, MAX(T1.Target) AS [Target], T1.TimeTableNo, T1.WorkDate, SUM(T1.IncollectFee) AS IncollectFee, MAX(T1.TripTranFee) AS TripTranFee," & vbCrLf
'        tmpSql = tmpSql & "MAX(T1.TripTransFee_Inc) AS TripTransFee_Inc, MAX(T1.VDTransFee) AS VDTransFee, ISNULL(T1.isClosed ,'N') AS isClosed," & vbCrLf
        
'***** ORG Code *********************
        tSqlView = "SELECT * " & vbCrLf
        If cboISClosed.ListIndex = 2 Then
            tSqlView = tSqlView & ",                          (SELECT     SUM(InCollectFee) AS Expr1" & vbCrLf
            tSqlView = tSqlView & "                            From dbo.vw_MenifestSearch T2" & vbCrLf
            tSqlView = tSqlView & "                            WHERE      (WorkDate BETWEEN '" & dtFrom.ValueYMD & "' AND '" & dtTo.ValueYMD & "') AND (ISNULL(isCanceled, 'N') <> 'Y') " & vbCrLf
            
            If dtFrom.ValueYMD >= "2561/08/01" Then
                If LblForm.Caption = "" Then
                    tSqlView = tSqlView & "                            AND (SDName_New like 'A_AK%'    or  SDName_New like 'B_AK%' or SDName_New like 'C_AK%' or SDName_New like 'D_AK%' or SDName_New like 'BWC%'  or SDName_New like 'AP_AP%' )                " & vbCrLf
                Else
                     tSqlView = tSqlView & "                            AND (SDName_New like 'A_ME%'    or  SDName_New like 'B_ME%' or SDName_New like 'C_ME%' or SDName_New like 'D_ME%'  or SDName_New like 'AP_ME%'   )                " & vbCrLf
                End If
            End If
            
            tSqlView = tSqlView & "                                                AND   (CompanyID='" & compID & "') AND (isClosed = 'Y') AND (T2.SaleName = vw_MenifestSearch.SaleName)) AS FeeClosed" & vbCrLf
            tSqlView = tSqlView & ",                          (SELECT COUNT (DISTINCT(TimeTableNo)) AS TimeTableNo" & vbCrLf
            tSqlView = tSqlView & "                            From dbo.vw_MenifestSearch T2" & vbCrLf
            tSqlView = tSqlView & "                            WHERE      (WorkDate BETWEEN '" & dtFrom.ValueYMD & "' AND '" & dtTo.ValueYMD & "') AND (ISNULL(isCanceled, 'N') <> 'Y') " & vbCrLf
            
            If dtFrom.ValueYMD >= "2561/08/01" Then
                If LblForm.Caption = "" Then
                    tSqlView = tSqlView & "                            AND (SDName_New like 'A_AK%'    or  SDName_New like 'B_AK%' or SDName_New like 'C_AK%' or SDName_New like 'D_AK%'  or SDName_New like 'BWC%'  or SDName_New like 'AP_AP%'  )                " & vbCrLf
                Else
                     tSqlView = tSqlView & "                            AND (SDName_New like 'A_ME%'    or  SDName_New like 'B_ME%' or SDName_New like 'C_ME%' or SDName_New like 'D_ME%'  or SDName_New like 'AP_ME%'   )                " & vbCrLf
                End If
            End If
            
            tSqlView = tSqlView & "                                                   AND (CompanyID='" & compID & "') AND (isClosed = 'Y') AND (T2.SaleName = vw_MenifestSearch.SaleName)) AS TripCount" & vbCrLf
        End If
        tSqlView = tSqlView & "From vw_MenifestSearch " & vbCrLf
        tSqlView = tSqlView & "WHERE WorkDate BETWEEN '" & dtFrom.ValueYMD & "' AND '" & dtTo.ValueYMD & "'" & vbCrLf
        tSqlView = tSqlView & "AND UnderSaleStaffID NOT IN ('A2EEC74864FC482E9D24AEC9847E2C98', '3DFB44CD947C400C9A0478F0FAC91CEA')" & vbCrLf
        
        If dtFrom.ValueYMD >= "2561/08/01" Then
                If LblForm.Caption = "" Then
                    tSqlView = tSqlView & "                            AND (SDName_New like 'A_AK%'    or  SDName_New like 'B_AK%' or SDName_New like 'C_AK%' or SDName_New like 'D_AK%'  or SDName_New like 'BWC%' or SDName_New like 'AP_AP%'   )                " & vbCrLf
                Else
                     tSqlView = tSqlView & "                            AND (SDName_New like 'A_ME%'    or  SDName_New like 'B_ME%' or SDName_New like 'C_ME%' or SDName_New like 'D_ME%' or SDName_New like 'AP_ME%'   )                " & vbCrLf
                End If
            End If
        
'        If LblForm.Caption <> "" Then
'            tSqlView = tSqlView & "AND customercode like 'M%' " & vbCrLf
'            Else
'            tSqlView = tSqlView & "AND customercode like 'C%' or customercode like 'S%' or customercode like '%F%' or customercode like '%L%' or customercode like '%P%' or customercode like '%T%' or customercode like '%K%' or customercode like '%G%' or customercode like '%E%' or customercode like '%R%'" & vbCrLf
'        End If
'
        
        If cboCompany.ListCount > 1 Then 'บริษัท
            If cboCompany.ListIndex > 0 Then
                tSqlView = tSqlView & "AND CompanyID = '" & compID & "'" & vbCrLf
            End If
        Else
            If cboCompany.ListIndex >= 0 Then
                tSqlView = tSqlView & "AND CompanyID = '" & compID & "'" & vbCrLf
            End If
        End If
        
        If cboISClosed.ListIndex = 1 Then
            tSqlView = tSqlView & "AND isClosed = 'Y'" & vbCrLf
        End If
        tSqlView = tSqlView & "AND ISNULL(isCanceled, 'N') <> 'Y'" & vbCrLf
        If CobJobtype.ListIndex = 0 Then
                tSqlView = tSqlView & "  AND (WasteCaseID IS NULL  Or  WasteCaseID  = '' )" & vbCrLf
        ElseIf CobJobtype.ListIndex = 1 Then
                 tSqlView = tSqlView & "  AND  WasteCaseID IS NOT NULL  AND  WasteCaseID  <> ''" & vbCrLf
       End If
        If cboCarType.ListIndex > 0 Then tSqlView = tSqlView & "AND TruckTypeID='" & coltruckType_(cboCarType.ListIndex).ID & "'" & vbCrLf
        If cboWasteType.ListIndex > 0 Then
            tSqlView = tSqlView & "AND MenifestID IN (SELECT MenifestID FROM vw_MenifestDetailSearch WHERE  (WorkDate BETWEEN '" & dtFrom.ValueYMD & "' AND '" & dtTo.ValueYMD & "') AND (WasteType IN " & cboWasteType.Text & ")"
            If cboOper.ListIndex > 0 Then tSqlView = tSqlView & " AND (TreatmentRate " & cboOper.Text & " " & txtCondition.Text & ")"
            tSqlView = tSqlView & ")" & vbCrLf
        End If
        If cboBmeNew.ListIndex = 1 Then
           tSqlView = tSqlView & "  AND  ComBme = 'Y' " & vbCrLf
        ElseIf cboBmeNew.ListIndex = 2 Then
           tSqlView = tSqlView & " AND ComBme  <> 'Y' " & vbCrLf
        End If
        tSql = "ALTER VIEW vw_MenifestSearch02 AS" & vbCrLf
        tSql = tSql & tSqlView
        DBConn.Execute tSql
' *****************************************
        tmpSql = "SELECT *" & vbCrLf
        If cboISClosed.ListIndex = 2 Then
            tmpSql = tmpSql & ",                          (SELECT     SUM(InCollectFee) AS Expr1" & vbCrLf
            tmpSql = tmpSql & "                            From dbo.vw_MenifestSearch_SaleTarget T2" & vbCrLf
            tmpSql = tmpSql & "                            WHERE      (WorkDate BETWEEN '" & dtFrom.ValueYMD & "' AND '" & dtTo.ValueYMD & "') AND (ISNULL(isCanceled, 'N') <> 'Y') AND" & vbCrLf
            tmpSql = tmpSql & "                                                   (CompanyID='" & compID & "') AND (isClosed = 'Y') AND (T2.SaleName = vw_MenifestSearch_SaleTarget.SaleName)) AS FeeClosed" & vbCrLf
            tmpSql = tmpSql & ",                          (SELECT COUNT (DISTINCT(TimeTableNo)) AS TimeTableNo" & vbCrLf
            tmpSql = tmpSql & "                            From dbo.vw_MenifestSearch_SaleTarget T2" & vbCrLf
            tmpSql = tmpSql & "                            WHERE      (WorkDate BETWEEN '" & dtFrom.ValueYMD & "' AND '" & dtTo.ValueYMD & "') AND (ISNULL(isCanceled, 'N') <> 'Y') AND" & vbCrLf
            tmpSql = tmpSql & "                                                   (CompanyID='" & compID & "') AND (isClosed = 'Y') AND (T2.SaleName = vw_MenifestSearch_SaleTarget.SaleName)) AS TripCount" & vbCrLf
        End If
'        tmpSql = tmpSql & "From vw_MenifestSearch" & vbCrLf
        tmpSql = tmpSql & "FROM         dbo.vw_MenifestSearch_SaleTarget" & vbCrLf
        tmpSql = tmpSql & "WHERE ForYearMonth BETWEEN '" & dtFrom.YearValue & Format(dtFrom.MonthValue, "00") & "' AND '" & dtTo.YearValue & Format(dtTo.MonthValue, "00") & "'" & vbCrLf

'        If dtFrom.ValueYMD >= "2561/08/01" Then
'            tmpSql = tmpSql & " AND (SDName_New like 'A_AK%'    or  SDName_New like 'B_AK%' or SDName_New like 'C_AK%' or SDName_New like 'D_AK%'   )  " & vbCrLf
'
'        Else
'              tmpSql = tmpSql & " AND (ZoneName like 'AKP%'  Or ZoneName like 'BWG%' Or ZoneName like 'BWC%' or ZoneName like 'KWG%' or ZoneName like 'KWC%' Or ZoneName like 'BKP%' Or ZoneName like 'BME%'  )  " & vbCrLf
'        End If
        If dtFrom.ValueYMD >= "2561/08/01" Then
                If LblForm.Caption = "" Then
                    tmpSql = tmpSql & " AND (ZoneName like 'AKP%'  Or ZoneName like 'BWG%' Or ZoneName like 'BWC%' or ZoneName like 'KWG%' or ZoneName like 'KWC%' Or ZoneName like 'BKP%' )  " & vbCrLf
                Else
                    tmpSql = tmpSql & " AND (ZoneName like 'BME%')  " & vbCrLf
                End If
        Else
                     tmpSql = tmpSql & " AND (ZoneName like 'AKP%'  Or ZoneName like 'BWG%' Or ZoneName like 'BWC%' or ZoneName like 'KWG%' or ZoneName like 'KWC%' Or ZoneName like 'BKP%' )  " & vbCrLf
        End If
        
        tmpSql = tmpSql & "AND SaleStaffID NOT IN ('A2EEC74864FC482E9D24AEC9847E2C98', '3DFB44CD947C400C9A0478F0FAC91CEA')" & vbCrLf
        
        If cboCompany.ListCount > 1 Then 'บริษัท
            If cboCompany.ListIndex > 0 Then
                tmpSql = tmpSql & "AND CompanyID = '" & compID & "'" & vbCrLf
            End If
        Else
            If cboCompany.ListIndex >= 0 Then
                tmpSql = tmpSql & "AND CompanyID = '" & compID & "'" & vbCrLf
            End If
        End If

        If Trim(txtSale.Text) <> "" Then tmpSql = tmpSql & "AND SaleStaffID='" & curStaff_.ID & "' " & vbCrLf
                
    tSqlOrder = "ORDER BY isClosed Desc, CompanyNameTH, ZoneName, SDName, SaleName, WorkDate, TimeTableNo, IncollectFee Desc" & vbCrLf
        
        If cboISClosed.ListIndex = 2 Then
            .ReportFileName = App.Path & "\Reports\Incomes_IsClosedN.rpt"
            tSqlView = tSqlView & "AND isClosed = 'N'" & vbCrLf
            tSql = "ALTER VIEW vw_MenifestSearch_ISClosedN AS" & vbCrLf
            tSql = tSql & tSqlView '& tSqlGroup
            DBConn.Execute tSql
            tmpSql = "SELECT * FROM vw_MenifestSearch_ISClosedN" & vbCrLf
        End If
        
        .Formulas(0) = "DateRange=" & Chr(34) & "ในช่วงวันที่ขนตั้งแต่ " & dtFrom.ValueDMY & " ถึงวันที่ " & dtTo.ValueDMY & Chr(34)
        .Formulas(2) = "ShowCompany=" & Chr(34) & cboCompany.Text & Chr(34)
        .Formulas(3) = "ReportTitle = " & Chr(34) & cboCarType.Text & ", สถานะใบเมนิเฟส " & cboISClosed.Text & ", ประเภทเวสต์ " & cboWasteType.Text & ", Commission " & cboBmeNew.Text & _
                                                                                            IIf(cboOper.ListIndex > 0, " " & cboOper.Text & " " & Format(txtCondition.Text, "#,##0") & " บาท/ตัน", "") & Chr(34)
        If cboCarType.ListIndex <= 0 Then .Formulas(4) = "ShowDetail = " & Chr(34) & "False" & Chr(34)
        If Check1.Value = 1 Then .Formulas(4) = "ShowRecords = " & Chr(34) & "True" & Chr(34)
        
        .WindowTitle = "รายงานยอดขาย"
        If cboISClosed.ListIndex = 2 Then .WindowTitle = "Sales Forecast average weight"
        .WindowShowCloseBtn = True
        .WindowShowExportBtn = True
        .DiscardSavedData = True
        .SQLQuery = tmpSql & tSqlOrder
        .Destination = printType
        .PrintReport
        If Trim(.LastErrorString) <> "" Then MsgBox .LastErrorString
    End With
End Sub



Private Sub txtSale_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then txtSale_Validate False
End Sub

Private Sub txtSale_Validate(Cancel As Boolean)
    If curStaff_ Is Nothing Then Set curStaff_ = New BWGCompanyStaff
    If Trim(txtSale.Text) = "" And Trim(curStaff_.StaffFName & " " & curStaff_.StaffLName) <> Trim(txtSale.Text) Then
        Set curStaff_ = Nothing
        Set curStaff_ = New BWGCompanyStaff
    Else
        If Trim(curStaff_.StaffFName & " " & curStaff_.StaffLName) <> Trim(txtSale.Text) Then
            Set curStaff_ = Nothing
            Set curStaff_ = New BWGCompanyStaff
            Load frmStaffSearch
            With frmStaffSearch
                .getCriteria (txtSale.Text)
                .SetCallerForm Me
                .Show 1
            End With
            If curStaff_.ID = "" Then txtSale.Text = ""
        End If
    End If
End Sub

Public Sub SetDataStaff(selStaff As BWGCompanyStaff, Optional isAssigned As Boolean = True)
Dim tmpRec As New ADODB.Recordset
Dim x%
    Set curStaff_ = Nothing
    Set curStaff_ = selStaff
    If curStaff_ Is Nothing Then Exit Sub
    With curStaff_
        txtSale.Text = .StaffFName & " " & .StaffLName
    End With
End Sub

