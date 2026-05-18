VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form FrmFinancialRatio 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Financial Ratios"
   ClientHeight    =   7065
   ClientLeft      =   4215
   ClientTop       =   2370
   ClientWidth     =   12090
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7065
   ScaleWidth      =   12090
   Begin TabDlg.SSTab SSTab1 
      Height          =   6900
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   11985
      _ExtentX        =   21140
      _ExtentY        =   12171
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Finacial Ratios"
      TabPicture(0)   =   "FrmFinancialRatio.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblID"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblNew"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label21"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "CmdNew"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "CmdSearch"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cbQ"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cbYears"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Command2"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmdSave"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Frame4"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Frame3"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Frame2"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Frame1"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).ControlCount=   14
      TabCaption(1)   =   "รายได้"
      TabPicture(1)   =   "FrmFinancialRatio.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cbYear2"
      Tab(1).Control(1)=   "cbMonth2"
      Tab(1).Control(2)=   "Command5"
      Tab(1).Control(3)=   "cmdSearchInc"
      Tab(1).Control(4)=   "cmdColse"
      Tab(1).Control(5)=   "cmdSaveInc"
      Tab(1).Control(6)=   "cbMonth"
      Tab(1).Control(7)=   "cbYear"
      Tab(1).Control(8)=   "grIncome"
      Tab(1).Control(9)=   "grIncomeCompany"
      Tab(1).Control(10)=   "Label54"
      Tab(1).Control(11)=   "Label49"
      Tab(1).Control(12)=   "Label48"
      Tab(1).Control(13)=   "Label47"
      Tab(1).Control(14)=   "Label46"
      Tab(1).Control(15)=   "Label45"
      Tab(1).Control(16)=   "Label44"
      Tab(1).Control(17)=   "Label43"
      Tab(1).Control(18)=   "Label42"
      Tab(1).Control(19)=   "Label41"
      Tab(1).Control(20)=   "Label40"
      Tab(1).Control(21)=   "Label39"
      Tab(1).Control(22)=   "Label38"
      Tab(1).Control(23)=   "Label37"
      Tab(1).Control(24)=   "Label36"
      Tab(1).Control(25)=   "Label35"
      Tab(1).Control(26)=   "Label30"
      Tab(1).Control(27)=   "Label29"
      Tab(1).Control(28)=   "Label28"
      Tab(1).Control(29)=   "Label27"
      Tab(1).Control(30)=   "Label26"
      Tab(1).Control(31)=   "Label24"
      Tab(1).Control(32)=   "Label22"
      Tab(1).Control(33)=   "lblyy"
      Tab(1).ControlCount=   34
      TabCaption(2)   =   "ต้นทุน"
      TabPicture(2)   =   "FrmFinancialRatio.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtOutWaste"
      Tab(2).Control(1)=   "txtInWaste"
      Tab(2).Control(2)=   "cbYear3"
      Tab(2).Control(3)=   "cbMonth3"
      Tab(2).Control(4)=   "Command6"
      Tab(2).Control(5)=   "Command4"
      Tab(2).Control(6)=   "Command3"
      Tab(2).Control(7)=   "Command1"
      Tab(2).Control(8)=   "cbMonth4"
      Tab(2).Control(9)=   "cbYear4"
      Tab(2).Control(10)=   "grCost"
      Tab(2).Control(11)=   "grBypassCost"
      Tab(2).Control(12)=   "Label60"
      Tab(2).Control(13)=   "Label59"
      Tab(2).Control(14)=   "lblSumStaff"
      Tab(2).Control(15)=   "Label58"
      Tab(2).Control(16)=   "Label53"
      Tab(2).Control(17)=   "Label52"
      Tab(2).Control(18)=   "Label57"
      Tab(2).Control(19)=   "Label56"
      Tab(2).Control(20)=   "Label55"
      Tab(2).Control(21)=   "lblSum"
      Tab(2).Control(22)=   "Label51"
      Tab(2).Control(23)=   "Label50"
      Tab(2).ControlCount=   24
      Begin VB.TextBox txtOutWaste 
         Height          =   285
         Left            =   -64740
         TabIndex        =   127
         Top             =   405
         Width           =   1005
      End
      Begin VB.TextBox txtInWaste 
         Height          =   285
         Left            =   -67440
         TabIndex        =   125
         Top             =   405
         Width           =   1005
      End
      Begin VB.ComboBox cbYear3 
         Height          =   315
         Left            =   -73695
         Style           =   2  'Dropdown List
         TabIndex        =   116
         Top             =   405
         Width           =   1275
      End
      Begin VB.ComboBox cbMonth3 
         Height          =   315
         Left            =   -71445
         Style           =   2  'Dropdown List
         TabIndex        =   115
         Top             =   405
         Width           =   1950
      End
      Begin VB.CommandButton Command6 
         Caption         =   "บันทึก"
         Height          =   510
         Left            =   -66855
         TabIndex        =   114
         Top             =   6300
         Width           =   1185
      End
      Begin VB.CommandButton Command4 
         Caption         =   "ปิด"
         Height          =   510
         Left            =   -64425
         TabIndex        =   113
         Top             =   6300
         Width           =   1185
      End
      Begin VB.CommandButton Command3 
         Caption         =   "แสดง"
         Height          =   510
         Left            =   -70500
         TabIndex        =   112
         Top             =   6300
         Width           =   1185
      End
      Begin VB.CommandButton Command1 
         Caption         =   "สร้างใหม่"
         Height          =   510
         Left            =   -65640
         TabIndex        =   111
         Top             =   6300
         Width           =   1185
      End
      Begin VB.ComboBox cbMonth4 
         Height          =   315
         Left            =   -72480
         Style           =   2  'Dropdown List
         TabIndex        =   110
         Top             =   6390
         Width           =   1950
      End
      Begin VB.ComboBox cbYear4 
         Height          =   315
         Left            =   -74370
         Style           =   2  'Dropdown List
         TabIndex        =   109
         Top             =   6390
         Width           =   1275
      End
      Begin VB.ComboBox cbYear2 
         Height          =   315
         Left            =   -74370
         Style           =   2  'Dropdown List
         TabIndex        =   106
         Top             =   6390
         Width           =   1275
      End
      Begin VB.ComboBox cbMonth2 
         Height          =   315
         Left            =   -72480
         Style           =   2  'Dropdown List
         TabIndex        =   105
         Top             =   6390
         Width           =   1950
      End
      Begin VB.CommandButton Command5 
         Caption         =   "สร้างใหม่"
         Height          =   510
         Left            =   -65820
         TabIndex        =   85
         Top             =   6300
         Width           =   1185
      End
      Begin VB.CommandButton cmdSearchInc 
         Caption         =   "แสดง"
         Height          =   510
         Left            =   -70500
         TabIndex        =   84
         Top             =   6300
         Width           =   1185
      End
      Begin VB.CommandButton cmdColse 
         Caption         =   "ปิด"
         Height          =   510
         Left            =   -64605
         TabIndex        =   83
         Top             =   6300
         Width           =   1185
      End
      Begin VB.CommandButton cmdSaveInc 
         Caption         =   "บันทึก"
         Height          =   510
         Left            =   -67035
         TabIndex        =   82
         Top             =   6300
         Width           =   1185
      End
      Begin VB.ComboBox cbMonth 
         Height          =   315
         Left            =   -71445
         Style           =   2  'Dropdown List
         TabIndex        =   79
         Top             =   405
         Width           =   1950
      End
      Begin VB.ComboBox cbYear 
         Height          =   315
         Left            =   -73695
         Style           =   2  'Dropdown List
         TabIndex        =   78
         Top             =   405
         Width           =   1275
      End
      Begin VB.Frame Frame1 
         Caption         =   "Comprehensive Income (Service Income)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   4740
         Left            =   90
         TabIndex        =   38
         Top             =   1125
         Width           =   4515
         Begin VB.TextBox RevenueBuring 
            Height          =   330
            Left            =   1890
            TabIndex        =   49
            Top             =   405
            Width           =   1950
         End
         Begin VB.TextBox Transportation 
            Height          =   330
            Left            =   1890
            TabIndex        =   48
            Top             =   765
            Width           =   1950
         End
         Begin VB.TextBox OtherIncome 
            Height          =   330
            Left            =   1890
            TabIndex        =   47
            Top             =   1125
            Width           =   1950
         End
         Begin VB.TextBox NetIncome 
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
            Height          =   330
            Left            =   1890
            Locked          =   -1  'True
            TabIndex        =   46
            Top             =   1485
            Width           =   1950
         End
         Begin VB.TextBox CostOfService 
            Height          =   330
            Left            =   1890
            TabIndex        =   45
            Top             =   1845
            Width           =   1950
         End
         Begin VB.TextBox GrossProfit 
            Height          =   330
            Left            =   1890
            TabIndex        =   44
            Top             =   2205
            Width           =   1950
         End
         Begin VB.TextBox Expense 
            Height          =   330
            Left            =   1890
            TabIndex        =   43
            Top             =   2565
            Width           =   1950
         End
         Begin VB.TextBox EBITDA 
            Height          =   330
            Left            =   1890
            TabIndex        =   42
            Top             =   2925
            Width           =   1950
         End
         Begin VB.TextBox EBIT 
            Height          =   330
            Left            =   1890
            TabIndex        =   41
            Top             =   3285
            Width           =   1950
         End
         Begin VB.TextBox ShareofProfit 
            Height          =   330
            Left            =   1890
            TabIndex        =   40
            Top             =   3870
            Width           =   1950
         End
         Begin VB.TextBox NetProfit 
            Height          =   330
            Left            =   1890
            TabIndex        =   39
            Top             =   4230
            Width           =   1950
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Revenue Buring :"
            Height          =   195
            Left            =   585
            TabIndex        =   71
            Top             =   450
            Width           =   1245
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Transportation Rubbish :"
            Height          =   195
            Left            =   90
            TabIndex        =   70
            Top             =   810
            Width           =   1740
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Other income :"
            Height          =   195
            Left            =   810
            TabIndex        =   69
            Top             =   1170
            Width           =   1035
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Net Income :"
            Height          =   195
            Left            =   900
            TabIndex        =   68
            Top             =   1530
            Width           =   915
         End
         Begin VB.Label LblP1 
            AutoSize        =   -1  'True
            Height          =   195
            Left            =   4230
            TabIndex        =   67
            Top             =   450
            Width           =   45
         End
         Begin VB.Label LblP2 
            AutoSize        =   -1  'True
            Height          =   195
            Left            =   4230
            TabIndex        =   66
            Top             =   810
            Width           =   45
         End
         Begin VB.Label LblP3 
            AutoSize        =   -1  'True
            Height          =   195
            Left            =   4230
            TabIndex        =   65
            Top             =   1170
            Width           =   45
         End
         Begin VB.Label LblP4 
            AutoSize        =   -1  'True
            Height          =   195
            Left            =   4230
            TabIndex        =   64
            Top             =   1530
            Width           =   45
         End
         Begin VB.Label LblP5 
            AutoSize        =   -1  'True
            Height          =   195
            Left            =   4230
            TabIndex        =   63
            Top             =   1890
            Width           =   45
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Cost Of Service :"
            Height          =   195
            Left            =   630
            TabIndex        =   62
            Top             =   1890
            Width           =   1200
         End
         Begin VB.Label LblP6 
            AutoSize        =   -1  'True
            Height          =   195
            Left            =   4230
            TabIndex        =   61
            Top             =   2250
            Width           =   45
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Gross Profit :"
            Height          =   195
            Left            =   900
            TabIndex        =   60
            Top             =   2250
            Width           =   900
         End
         Begin VB.Label LblP7 
            AutoSize        =   -1  'True
            Height          =   195
            Left            =   4230
            TabIndex        =   59
            Top             =   2610
            Width           =   45
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Expenses :"
            Height          =   195
            Left            =   1035
            TabIndex        =   58
            Top             =   2610
            Width           =   780
         End
         Begin VB.Label LblP8 
            AutoSize        =   -1  'True
            Height          =   195
            Left            =   4230
            TabIndex        =   57
            Top             =   2970
            Width           =   45
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "EBITDA :"
            Height          =   195
            Left            =   1125
            TabIndex        =   56
            Top             =   2970
            Width           =   675
         End
         Begin VB.Label LblP9 
            AutoSize        =   -1  'True
            Height          =   195
            Left            =   4230
            TabIndex        =   55
            Top             =   3330
            Width           =   45
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "EBIT :"
            Height          =   195
            Left            =   1350
            TabIndex        =   54
            Top             =   3330
            Width           =   450
         End
         Begin VB.Label LblP10 
            AutoSize        =   -1  'True
            Height          =   195
            Left            =   4230
            TabIndex        =   53
            Top             =   3960
            Width           =   45
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Share of profit from investment in associates :"
            Height          =   195
            Left            =   585
            TabIndex        =   52
            Top             =   3645
            Width           =   3195
         End
         Begin VB.Label LblP11 
            AutoSize        =   -1  'True
            Height          =   195
            Left            =   4230
            TabIndex        =   51
            Top             =   4320
            Width           =   45
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Net profit :"
            Height          =   195
            Left            =   1035
            TabIndex        =   50
            Top             =   4275
            Width           =   735
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Financial Ratios"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   2490
         Left            =   4680
         TabIndex        =   25
         Top             =   1125
         Width           =   3165
         Begin VB.TextBox EPS 
            Height          =   330
            Left            =   1395
            TabIndex        =   31
            Top             =   2070
            Width           =   1545
         End
         Begin VB.TextBox BV 
            Height          =   330
            Left            =   1395
            TabIndex        =   30
            Top             =   1710
            Width           =   1545
         End
         Begin VB.TextBox DebtRatio 
            Height          =   330
            Left            =   1395
            TabIndex        =   29
            Top             =   1350
            Width           =   1545
         End
         Begin VB.TextBox DE 
            Height          =   330
            Left            =   1395
            TabIndex        =   28
            Top             =   990
            Width           =   1545
         End
         Begin VB.TextBox ROA 
            Height          =   330
            Left            =   1395
            TabIndex        =   27
            Top             =   630
            Width           =   1545
         End
         Begin VB.TextBox ROE 
            Height          =   330
            Left            =   1395
            TabIndex        =   26
            Top             =   270
            Width           =   1545
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "EPS. (B.) :"
            Height          =   195
            Left            =   585
            TabIndex        =   37
            Top             =   2115
            Width           =   735
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            Caption         =   "BV. :"
            Height          =   195
            Left            =   945
            TabIndex        =   36
            Top             =   1755
            Width           =   375
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            Caption         =   "Debt Ratio (x) :"
            Height          =   195
            Left            =   270
            TabIndex        =   35
            Top             =   1395
            Width           =   1065
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            Caption         =   "D/E (x) :"
            Height          =   195
            Left            =   720
            TabIndex        =   34
            Top             =   1035
            Width           =   600
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            Caption         =   "ROA (%) :"
            Height          =   195
            Left            =   630
            TabIndex        =   33
            Top             =   675
            Width           =   690
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            Caption         =   "ROE (%) :"
            Height          =   195
            Left            =   630
            TabIndex        =   32
            Top             =   315
            Width           =   690
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Financial Position"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   1410
         Left            =   4680
         TabIndex        =   18
         Top             =   3645
         Width           =   3165
         Begin VB.TextBox Asset 
            Height          =   330
            Left            =   1395
            TabIndex        =   21
            Top             =   270
            Width           =   1545
         End
         Begin VB.TextBox Liability 
            Height          =   330
            Left            =   1395
            TabIndex        =   20
            Top             =   630
            Width           =   1545
         End
         Begin VB.TextBox Equity 
            Height          =   330
            Left            =   1395
            TabIndex        =   19
            Top             =   990
            Width           =   1545
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Asset :"
            Height          =   195
            Left            =   630
            TabIndex        =   24
            Top             =   315
            Width           =   480
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Liability :"
            Height          =   195
            Left            =   630
            TabIndex        =   23
            Top             =   675
            Width           =   600
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Equity :"
            Height          =   195
            Left            =   720
            TabIndex        =   22
            Top             =   1035
            Width           =   525
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Cash Flow"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   2175
         Left            =   7920
         TabIndex        =   7
         Top             =   1125
         Width           =   3840
         Begin VB.TextBox OperatingAct 
            Height          =   330
            Left            =   2115
            TabIndex        =   12
            Top             =   270
            Width           =   1545
         End
         Begin VB.TextBox InvestingAct 
            Height          =   330
            Left            =   2115
            TabIndex        =   11
            Top             =   630
            Width           =   1545
         End
         Begin VB.TextBox FinancingAct 
            Height          =   330
            Left            =   2115
            TabIndex        =   10
            Top             =   990
            Width           =   1545
         End
         Begin VB.TextBox NetIncrease 
            Height          =   330
            Left            =   2115
            TabIndex        =   9
            Top             =   1350
            Width           =   1545
         End
         Begin VB.TextBox EndingCash 
            Height          =   330
            Left            =   2115
            TabIndex        =   8
            Top             =   1710
            Width           =   1545
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Operating activities (CFO) :"
            Height          =   195
            Left            =   135
            TabIndex        =   17
            Top             =   315
            Width           =   1890
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Investing activities (CFI) :"
            Height          =   195
            Left            =   270
            TabIndex        =   16
            Top             =   675
            Width           =   1770
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Financing acties (CFF) :"
            Height          =   195
            Left            =   360
            TabIndex        =   15
            Top             =   1035
            Width           =   1665
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Net Increase :"
            Height          =   195
            Left            =   1035
            TabIndex        =   14
            Top             =   1395
            Width           =   1005
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "Ending cash :"
            Height          =   195
            Left            =   1080
            TabIndex        =   13
            Top             =   1755
            Width           =   975
         End
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "บันทึก"
         Height          =   510
         Left            =   8100
         TabIndex        =   6
         Top             =   6075
         Width           =   1185
      End
      Begin VB.CommandButton Command2 
         Caption         =   "ปิด"
         Height          =   510
         Left            =   10530
         TabIndex        =   5
         Top             =   6075
         Width           =   1185
      End
      Begin VB.ComboBox cbYears 
         Height          =   315
         Left            =   6120
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   675
         Width           =   1590
      End
      Begin VB.ComboBox cbQ 
         Height          =   315
         Left            =   4860
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   675
         Width           =   870
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "ค้นหา"
         Height          =   510
         Left            =   90
         TabIndex        =   2
         Top             =   6075
         Width           =   1185
      End
      Begin VB.CommandButton CmdNew 
         Caption         =   "สร้างใหม่"
         Height          =   510
         Left            =   9315
         TabIndex        =   1
         Top             =   6075
         Width           =   1185
      End
      Begin WasteManagment.ctlGrid grIncome 
         Height          =   2115
         Left            =   -74955
         TabIndex        =   80
         Top             =   810
         Width           =   9405
         _ExtentX        =   16589
         _ExtentY        =   3731
      End
      Begin WasteManagment.ctlGrid grIncomeCompany 
         Height          =   2430
         Left            =   -74955
         TabIndex        =   86
         Top             =   3375
         Width           =   9405
         _ExtentX        =   16589
         _ExtentY        =   4286
      End
      Begin WasteManagment.ctlGrid grCost 
         Height          =   3420
         Left            =   -74955
         TabIndex        =   117
         Top             =   1080
         Width           =   6120
         _ExtentX        =   10795
         _ExtentY        =   6033
      End
      Begin WasteManagment.ctlGrid grBypassCost 
         Height          =   3420
         Left            =   -68430
         TabIndex        =   129
         Top             =   1080
         Width           =   5085
         _ExtentX        =   7144
         _ExtentY        =   6033
      End
      Begin VB.Label Label60 
         Caption         =   "ต้นทุนบุคลากร(ได้รวมอยู่ในต้นทุนข้อ 1-3 แล้ว)"
         Height          =   240
         Left            =   -68115
         TabIndex        =   133
         Top             =   855
         Width           =   3660
      End
      Begin VB.Label Label59 
         Caption         =   "ต้นทุน"
         Height          =   240
         Left            =   -74730
         TabIndex        =   132
         Top             =   855
         Width           =   1275
      End
      Begin VB.Label lblSumStaff 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000FF00&
         Caption         =   "0"
         Height          =   195
         Left            =   -65505
         TabIndex        =   131
         Top             =   4590
         Width           =   2040
      End
      Begin VB.Label Label58 
         AutoSize        =   -1  'True
         Caption         =   "รวม"
         Height          =   195
         Left            =   -66225
         TabIndex        =   130
         Top             =   4590
         Width           =   285
      End
      Begin VB.Label Label54 
         Caption         =   $"FrmFinancialRatio.frx":0054
         Height          =   5145
         Left            =   -65460
         TabIndex        =   128
         Top             =   810
         Width           =   2265
      End
      Begin VB.Label Label53 
         AutoSize        =   -1  'True
         Caption         =   "ปริมาณกากเผา(ตัน) :"
         Height          =   195
         Left            =   -66225
         TabIndex        =   126
         Top             =   450
         Width           =   1455
      End
      Begin VB.Label Label52 
         AutoSize        =   -1  'True
         Caption         =   "ปริมาณกากรับ(ตัน) :"
         Height          =   195
         Left            =   -68925
         TabIndex        =   124
         Top             =   450
         Width           =   1395
      End
      Begin VB.Label Label57 
         AutoSize        =   -1  'True
         Caption         =   "ปี :"
         Height          =   195
         Left            =   -73920
         TabIndex        =   123
         Top             =   450
         Width           =   195
      End
      Begin VB.Label Label56 
         AutoSize        =   -1  'True
         Caption         =   "เดือน :"
         Height          =   195
         Left            =   -72030
         TabIndex        =   122
         Top             =   450
         Width           =   480
      End
      Begin VB.Label Label55 
         AutoSize        =   -1  'True
         Caption         =   "รวม"
         Height          =   195
         Left            =   -72075
         TabIndex        =   121
         Top             =   4590
         Width           =   285
      End
      Begin VB.Label lblSum 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000FF00&
         Caption         =   "0"
         Height          =   195
         Left            =   -70995
         TabIndex        =   120
         Top             =   4590
         Width           =   2040
      End
      Begin VB.Label Label51 
         AutoSize        =   -1  'True
         Caption         =   "เดือน :"
         Height          =   195
         Left            =   -73065
         TabIndex        =   119
         Top             =   6435
         Width           =   480
      End
      Begin VB.Label Label50 
         AutoSize        =   -1  'True
         Caption         =   "ปี :"
         Height          =   195
         Left            =   -74640
         TabIndex        =   118
         Top             =   6435
         Width           =   195
      End
      Begin VB.Label Label49 
         AutoSize        =   -1  'True
         Caption         =   "ปี :"
         Height          =   195
         Left            =   -74640
         TabIndex        =   108
         Top             =   6435
         Width           =   195
      End
      Begin VB.Label Label48 
         AutoSize        =   -1  'True
         Caption         =   "เดือน :"
         Height          =   195
         Left            =   -73065
         TabIndex        =   107
         Top             =   6435
         Width           =   480
      End
      Begin VB.Label Label47 
         AutoSize        =   -1  'True
         Caption         =   "รวม"
         Height          =   195
         Left            =   -73560
         TabIndex        =   104
         Top             =   5850
         Width           =   285
      End
      Begin VB.Label Label46 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000FF00&
         Caption         =   "0"
         Height          =   195
         Left            =   -71490
         TabIndex        =   103
         Top             =   5850
         Width           =   1905
      End
      Begin VB.Label Label45 
         BackColor       =   &H0000FF00&
         Caption         =   "0"
         Height          =   195
         Left            =   -70230
         TabIndex        =   102
         Top             =   6210
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.Label Label44 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000FF00&
         Caption         =   "0"
         Height          =   195
         Left            =   -69555
         TabIndex        =   101
         Top             =   5850
         Width           =   1905
      End
      Begin VB.Label Label43 
         BackColor       =   &H0000FF00&
         Caption         =   "0"
         Height          =   195
         Left            =   -68025
         TabIndex        =   100
         Top             =   6210
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.Label Label42 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000FF00&
         Caption         =   "0"
         Height          =   195
         Left            =   -67620
         TabIndex        =   99
         Top             =   5850
         Width           =   1905
      End
      Begin VB.Label Label41 
         BackColor       =   &H0000FF00&
         Caption         =   "0"
         Height          =   195
         Left            =   -65955
         TabIndex        =   98
         Top             =   6210
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.Label Label40 
         BackColor       =   &H0000FF00&
         Caption         =   "0"
         Height          =   195
         Left            =   -65145
         TabIndex        =   97
         Top             =   6210
         Visible         =   0   'False
         Width           =   1230
      End
      Begin VB.Label Label39 
         BackColor       =   &H0000FF00&
         Caption         =   "0"
         Height          =   195
         Left            =   -63885
         TabIndex        =   96
         Top             =   6210
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.Label Label38 
         BackColor       =   &H0000FF00&
         Caption         =   "0"
         Height          =   195
         Left            =   -63885
         TabIndex        =   95
         Top             =   3285
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.Label Label37 
         BackColor       =   &H0000FF00&
         Caption         =   "0"
         Height          =   195
         Left            =   -65145
         TabIndex        =   94
         Top             =   3285
         Visible         =   0   'False
         Width           =   1230
      End
      Begin VB.Label Label36 
         BackColor       =   &H0000FF00&
         Caption         =   "0"
         Height          =   195
         Left            =   -65955
         TabIndex        =   93
         Top             =   3285
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.Label Label35 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000FF00&
         Caption         =   "0"
         Height          =   195
         Left            =   -67845
         TabIndex        =   92
         Top             =   2970
         Width           =   2265
      End
      Begin VB.Label Label30 
         BackColor       =   &H0000FF00&
         Caption         =   "0"
         Height          =   195
         Left            =   -68025
         TabIndex        =   91
         Top             =   3285
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.Label Label29 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000FF00&
         Caption         =   "0"
         Height          =   195
         Left            =   -70185
         TabIndex        =   90
         Top             =   2970
         Width           =   2310
      End
      Begin VB.Label Label28 
         BackColor       =   &H0000FF00&
         Caption         =   "0"
         Height          =   195
         Left            =   -70230
         TabIndex        =   89
         Top             =   3285
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.Label Label27 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000FF00&
         Caption         =   "0"
         Height          =   195
         Left            =   -71985
         TabIndex        =   88
         Top             =   2970
         Width           =   1770
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "รวม"
         Height          =   195
         Left            =   -73560
         TabIndex        =   87
         Top             =   2970
         Width           =   285
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "เดือน :"
         Height          =   195
         Left            =   -72030
         TabIndex        =   81
         Top             =   450
         Width           =   480
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "ปี :"
         Height          =   195
         Left            =   -73920
         TabIndex        =   77
         Top             =   450
         Width           =   195
      End
      Begin VB.Label lblyy 
         Height          =   195
         Left            =   -64425
         TabIndex        =   76
         Top             =   495
         Visible         =   0   'False
         Width           =   870
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "ปี :"
         Height          =   195
         Left            =   5850
         TabIndex        =   75
         Top             =   720
         Width           =   195
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "ไตรมาส :"
         Height          =   195
         Left            =   4140
         TabIndex        =   74
         Top             =   720
         Width           =   645
      End
      Begin VB.Label lblNew 
         Height          =   195
         Left            =   10305
         TabIndex        =   73
         Top             =   3870
         Visible         =   0   'False
         Width           =   570
      End
      Begin VB.Label lblID 
         Height          =   195
         Left            =   10350
         TabIndex        =   72
         Top             =   4185
         Visible         =   0   'False
         Width           =   570
      End
   End
End
Attribute VB_Name = "FrmFinancialRatio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Asset_Change()
'    If Not IsNumeric(Asset.Text) Then SendKeys "{backspace}"
End Sub
Private Sub Asset_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        With Liability
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
        End With
            Asset.Text = Format(Asset.Text, "#,##0.00")
    End If
End Sub
Private Sub BV_Change()
'    If Not IsNumeric(BV.Text) Then SendKeys "{backspace}"
End Sub
Private Sub BV_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        With EPS
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
        End With
            BV.Text = Format(BV.Text, "#,##0.0000")
    End If
End Sub
Private Sub cbYears_Click()
                    Dim tSql As String
                Dim rs As New ADODB.Recordset
                Dim i As Integer
                    tSql = "SELECT * FROM Financial_Ratios " & vbCrLf
                    tSql = tSql & "where qname= '" & Trim(cbYears.Text) & "' " & vbCrLf
                    Set rs = GetRS(tSql, adOpenForwardOnly, adLockOptimistic)
                    If rs.RecordCount > 0 Then
                        ROE.Text = rs!ROE
                        ROA.Text = rs!ROA
                        DE.Text = rs!dex
                        DebtRatio.Text = rs!DebtRatio
                        BV.Text = rs!BV
                        EPS.Text = rs!EPS
                    Else
                        ROE.Text = "0"
                        ROA.Text = "0"
                        DE.Text = "0"
                        DebtRatio.Text = "0"
                        BV.Text = "0"
                        EPS.Text = "0"
                    End If
End Sub

Private Sub cmdColse_Click()
    Unload Me
End Sub

Private Sub cmdNew_Click()
    setObj
    lblID.Caption = ""
    lblNew.Caption = "Y"
End Sub

Private Sub CmdSave_Click()
    Dim tSql As String
    If Trim(RevenueBuring.Text) = "" Then RevenueBuring.Text = "0"
    If Trim(Transportation.Text) = "" Then Transportation.Text = "0"
    If Trim(OtherIncome.Text) = "" Then OtherIncome.Text = "0"
    If Trim(CostOfService.Text) = "" Then CostOfService.Text = "0"
    If Trim(GrossProfit.Text) = "" Then GrossProfit.Text = "0"
    If Trim(Expense.Text) = "" Then Expense.Text = "0"
    If Trim(EBITDA.Text) = "" Then EBITDA.Text = "0"
    If Trim(EBIT.Text) = "" Then EBIT.Text = "0"
    If Trim(ShareofProfit.Text) = "" Then ShareofProfit.Text = "0"
    If Trim(NetProfit.Text) = "" Then NetProfit.Text = "0"
    If Trim(ROE.Text) = "" Then ROE.Text = "0"
    If Trim(ROA.Text) = "" Then ROA.Text = "0"
    If Trim(DE.Text) = "" Then DE.Text = "0"
    If Trim(DebtRatio.Text) = "" Then DebtRatio.Text = "0"
    If Trim(BV.Text) = "" Then BV.Text = "0"
    If Trim(EPS.Text) = "" Then EPS.Text = "0"
    If Trim(Asset.Text) = "" Then Asset.Text = "0"
    If Trim(Liability.Text) = "" Then Liability.Text = "0"
    If Trim(Equity.Text) = "" Then Equity.Text = "0"
    If Trim(OperatingAct.Text) = "" Then OperatingAct.Text = "0"
    If Trim(InvestingAct.Text) = "" Then InvestingAct.Text = "0"
    If Trim(FinancingAct.Text) = "" Then FinancingAct.Text = "0"
    If Trim(NetIncrease.Text) = "" Then NetIncrease.Text = "0"
    If Trim(EndingCash.Text) = "" Then EndingCash.Text = "0"
    
    If lblNew.Caption = "Y" Then
    If cbQ.ListIndex = 0 Then
        MsgBox "กรุณาระบุไตรมาสก่อน", vbExclamation + vbOKOnly, "บันทึกข้อมูลไม่ได้"
        Exit Sub
    End If
                Dim rs As New ADODB.Recordset
                Dim i As Integer
                    tSql = "SELECT * FROM Financial_Ratios " & vbCrLf
                    tSql = tSql & "where q='" & Trim(cbQ.Text) & "' and qname= '" & Trim(cbYears.Text) & "' " & vbCrLf
                    Set rs = GetRS(tSql, adOpenForwardOnly, adLockOptimistic)
                    If rs.RecordCount > 0 Then
                        MsgBox "ข้อมูลไตรมาสของปีที่ระบุมีข้อมูลอยู่ในฐานข้อมูลแล้ว กรุณาตรวจสอบ", vbExclamation + vbOKOnly, "บันทึกข้อมูลไม่ได้"
                        Exit Sub
                    End If

        tSql = "INSERT INTO Financial_Ratios(Q,QName,RevenueBuring,TransportationRubbish,OtherIncome,NetIncome,CostOfService,GrossProfit,Expenses,EBITDA"
        tSql = tSql & ",EBIT,ShareOfProfit,NetProfit,ROE,ROA,DEX,DebtRatio,BV,EPS,Assete,Liabilities,Equity,OperatingAc,InvestingAc,FinancingAc,NetIncrease,EndingCash) "
        tSql = tSql + " Values ('" & Trim(cbQ.Text) & "','" & Trim(cbYears.Text) & "', " & CDbl(RevenueBuring.Text) & ", " & CDbl(Transportation.Text) & "," & CDbl(OtherIncome.Text) & ","
        tSql = tSql + " " & CDbl(NetIncome.Text) & "," & CDbl(CostOfService.Text) & "," & CDbl(GrossProfit.Text) & "," & CDbl(Expense.Text) & "," & CDbl(EBITDA.Text) & " "
        tSql = tSql + "," & CDbl(EBIT.Text) & "," & CDbl(ShareofProfit.Text) & "," & CDbl(NetProfit.Text) & "," & CDbl(ROE.Text) & "," & CDbl(ROA.Text) & "," & CDbl(DE.Text) & ""
        tSql = tSql + "," & CDbl(DebtRatio.Text) & "," & CDbl(BV.Text) & "," & CDbl(EPS.Text) & "," & CDbl(Asset.Text) & "," & CDbl(Liability.Text) & "," & CDbl(Equity.Text) & ""
        tSql = tSql + "," & CDbl(OperatingAct.Text) & "," & CDbl(InvestingAct.Text) & "," & CDbl(FinancingAct.Text) & "," & CDbl(NetIncrease.Text) & "," & CDbl(EndingCash.Text) & ")"
        DBConnExc tSql
    ElseIf lblNew.Caption = "N" Then
        DBConnExc "Update Financial_Ratios Set RevenueBuring=" & CDbl(Trim(RevenueBuring.Text)) & ",TransportationRubbish=" & CDbl(Trim(Transportation.Text)) & ",OtherIncome=" & CDbl(Trim(OtherIncome.Text)) & " " & _
        ",NetIncome=" & CDbl(Trim(NetIncome.Text)) & ",CostOfService=" & CDbl(Trim(CostOfService.Text)) & ",GrossProfit=" & CDbl(Trim(GrossProfit.Text)) & " " & _
        ",Expenses=" & CDbl(Trim(Expense.Text)) & ",EBITDA=" & CDbl(Trim(EBITDA.Text)) & ",EBIT=" & CDbl(Trim(EBIT.Text)) & ",ShareOfProfit=" & CDbl(Trim(ShareofProfit.Text)) & " " & _
        ",NetProfit=" & CDbl(Trim(NetProfit.Text)) & ",ROE=" & CDbl(Trim(ROE.Text)) & ",ROA=" & CDbl(Trim(ROA.Text)) & ",DEX=" & CDbl(Trim(DE.Text)) & ",DebtRatio=" & CDbl(Trim(DebtRatio.Text)) & " " & _
        ",BV=" & CDbl(Trim(BV.Text)) & ",EPS=" & CDbl(Trim(EPS.Text)) & ",Assete=" & CDbl(Trim(Asset.Text)) & ",Liabilities=" & CDbl(Trim(Liability.Text)) & ",Equity=" & CDbl(Trim(Equity.Text)) & " " & _
        ",OperatingAc=" & CDbl(Trim(OperatingAct.Text)) & ",InvestingAc=" & CDbl(Trim(InvestingAct.Text)) & ",FinancingAc=" & CDbl(Trim(FinancingAct.Text)) & ",NetIncrease=" & CDbl(Trim(NetIncrease.Text)) & " " & _
        ",EndingCash=" & CDbl(Trim(EndingCash.Text)) & " WHERE ID='" & Trim(lblID.Caption) & "'"
    End If
        MsgBox "เก็บข้อมูลเรียบร้อย", vbInformation + vbOKOnly, "บันทึกข้อมูล"
        setObj
End Sub

Private Sub cmdSaveInc_Click()
Dim sqlstm, SqlStm2 As String
Dim RSC, rsC2 As New ADODB.Recordset
Dim a As Integer
Dim i As Integer
    If cbYear.ListIndex = 0 Then
        MsgBox "กรุณาระบุปี", vbExclamation + vbOKOnly, "บันทึกข้อมูลไม่ได้"
        Exit Sub
    End If
            sqlstm = "SELECT * FROM tbIncomeResult " & vbCrLf
            sqlstm = sqlstm & "where IncYear='" & cbYear.Text & "' and IncMonth='" & cbMonth.Text & "'" & vbCrLf
            Set RSC = GetRS(sqlstm, adOpenForwardOnly, adLockOptimistic)
                If RSC.RecordCount > 0 And grIncome.getGridObj.TextMatrix(1, 5) = "" Then
                    MsgBox "ค้นพบข้อมูลปีและเดือนที่ระบุในฐานข้อมูล กรุณาตรวจสอบ", vbExclamation + vbOKOnly, "พบข้อมูลซ้ำ"
                    Exit Sub
                End If
            SqlStm2 = "SELECT * FROM tbIncCompanyResult " & vbCrLf
            SqlStm2 = SqlStm2 & "where IncYear='" & cbYear.Text & "' and IncMonth='" & cbMonth.Text & "'" & vbCrLf
            Set rsC2 = GetRS(SqlStm2, adOpenForwardOnly, adLockOptimistic)
                If rsC2.RecordCount > 0 And grIncomeCompany.getGridObj.TextMatrix(1, 5) = "" Then
                    MsgBox "ค้นพบข้อมูลปีและเดือนที่ระบุในฐานข้อมูล กรุณาตรวจสอบ", vbExclamation + vbOKOnly, "พบข้อมูลซ้ำ"
                    Exit Sub
                End If
                
        For i = 1 To grIncome.getGridObj.Rows - 1
            With grIncome.getGridObj
                Dim Sql, tSql As String
                Dim rsG As New ADODB.Recordset
                Dim x As Integer
                    Sql = "SELECT * FROM tbIncomeResult " & vbCrLf
                    Sql = Sql & "where ID='" & .TextMatrix(i, 5) & "' " & vbCrLf
                    Set rsG = GetRS(Sql, adOpenForwardOnly, adLockOptimistic)
                    If rsG.RecordCount > 0 Then
                        DBConnExc "Update tbIncomeResult Set Income=" & CDbl(.TextMatrix(i, 2)) & ",IncomeComp=" & CDbl(.TextMatrix(i, 3)) & ",IncomeBroke=" & CDbl(.TextMatrix(i, 4)) & " where ID='" & .TextMatrix(i, 5) & "' "
                    Else
                        tSql = "INSERT INTO tbIncomeResult(IncID,Income,IncomeComp,IncomeBroke,IncYear,IncMonth) "
                        tSql = tSql + " Values (" & .TextMatrix(i, 0) & "," & CDbl(.TextMatrix(i, 2)) & ", " & CDbl(.TextMatrix(i, 3)) & ", " & CDbl(.TextMatrix(i, 4)) & ",'" & cbYear.Text & "','" & cbMonth.Text & "')"
                        DBConnExc tSql
                    End If
            End With
        Next
        For i = 1 To grIncomeCompany.getGridObj.Rows - 1
            With grIncomeCompany.getGridObj
                Dim Sql1, tSql1 As String
                Dim rsG1 As New ADODB.Recordset
                Dim X1 As Integer
                    Sql1 = "SELECT * FROM tbIncCompanyResult " & vbCrLf
                    Sql1 = Sql1 & "where ID='" & .TextMatrix(i, 5) & "' " & vbCrLf
                    Set rsG1 = GetRS(Sql1, adOpenForwardOnly, adLockOptimistic)
                    If rsG1.RecordCount > 0 Then
                        DBConnExc "Update tbIncCompanyResult Set Income=" & CDbl(.TextMatrix(i, 2)) & ",SK9=" & CDbl(.TextMatrix(i, 3)) & ",TrnInc=" & CDbl(.TextMatrix(i, 4)) & " where ID='" & .TextMatrix(i, 5) & "' "
                    Else
                        tSql1 = "INSERT INTO tbIncCompanyResult(CompanyID,Income,SK9,TrnInc,IncYear,IncMonth) "
                        tSql1 = tSql1 + " Values (" & .TextMatrix(i, 0) & "," & CDbl(.TextMatrix(i, 2)) & ", " & CDbl(.TextMatrix(i, 3)) & ", " & CDbl(.TextMatrix(i, 4)) & ",'" & cbYear.Text & "','" & cbMonth.Text & "')"
                        DBConnExc tSql1
                    End If
            End With
        Next
        MsgBox "บันทึกข้อมูลเรียบร้อย", vbInformation + vbOKOnly, "บันทึกข้อมูล"
        Exit Sub
End Sub
Private Sub cmdSearch_Click()
    FrmFSearch.Show 1
End Sub

Private Sub cmdSearchInc_Click()
                Dim Sql, tSql As String
                Dim rsG As New ADODB.Recordset
                Dim i As Integer
                
                If cbYear2.ListIndex = 0 Then
                    MsgBox "เลือกปีก่อนค้นหา", vbInformation + vbOKOnly, "ค้นหาข้อมูล"
                    Exit Sub
                End If
                Label27.Caption = "0"
                Label29.Caption = "0"
                Label35.Caption = "0"
                Label46.Caption = "0"
                Label44.Caption = "0"
                Label42.Caption = "0"
                    Sql = "SELECT * FROM vw_IncomeResult " & vbCrLf
                    Sql = Sql & "where IncYear='" & cbYear2.Text & "' and IncMonth='" & cbMonth2.Text & "' order by id" & vbCrLf
                    Set rsG = GetRS(Sql, adOpenForwardOnly, adLockOptimistic)
                    If rsG.RecordCount > 0 Then
                        With grIncome.getGridObj
                        grIncome.ClearAllData
                            .Rows = rsG.RecordCount + 1
                            For i = 1 To rsG.RecordCount
                                .TextMatrix(i, 0) = rsG!IncID
                                .TextMatrix(i, 1) = rsG!IncName
                                .TextMatrix(i, 2) = Format(rsG!Income, "#,##0.00")
                                .TextMatrix(i, 3) = Format(rsG!IncomeComp, "#,##0.00")
                                .TextMatrix(i, 4) = Format(rsG!IncomeBroke, "#,##0.00")
                                .TextMatrix(i, 5) = rsG!ID
                                
                                Label27.Caption = Format(CDbl(Label27.Caption) + .TextMatrix(i, 2), "#,##0.00")
                                Label29.Caption = Format(CDbl(Label29.Caption) + .TextMatrix(i, 3), "#,##0.00")
                                Label35.Caption = Format(CDbl(Label35.Caption) + .TextMatrix(i, 4), "#,##0.00")
                                rsG.MoveNext
                            Next
                        End With
                        Else
                        MsgBox "ไม่พบข้อมูลที่ต้องการแสดง", vbInformation + vbOKOnly, "ไม่มีข้อมูลที่ค้นหา"
                    End If
                Dim Sql1, tSql1 As String
                Dim rsG1 As New ADODB.Recordset
                Dim X1 As Integer
                    Sql1 = "SELECT * FROM vw_IncomeCompany " & vbCrLf
                    Sql1 = Sql1 & "where IncYear='" & cbYear2.Text & "' and IncMonth='" & cbMonth2.Text & "' order by id" & vbCrLf
                    Set rsG1 = GetRS(Sql1, adOpenForwardOnly, adLockOptimistic)
                    If rsG1.RecordCount > 0 Then
                        With grIncomeCompany.getGridObj
                        grIncomeCompany.ClearAllData
                            .Rows = rsG1.RecordCount + 1
                            For i = 1 To rsG1.RecordCount
                                .TextMatrix(i, 0) = rsG1!CompanyID
                                .TextMatrix(i, 1) = rsG1!CompanyName
                                .TextMatrix(i, 2) = Format(rsG1!Income, "#,##0.00")
                                .TextMatrix(i, 3) = Format(rsG1!SK9, "#,##0.00")
                                .TextMatrix(i, 4) = Format(rsG1!TrnInc, "#,##0.00")
                                .TextMatrix(i, 5) = rsG1!ID
                                rsG1.MoveNext
                                Label46.Caption = Format(CDbl(Label46.Caption) + .TextMatrix(i, 2), "#,##0.00")
                                Label44.Caption = Format(CDbl(Label44.Caption) + .TextMatrix(i, 3), "#,##0.00")
                                Label42.Caption = Format(CDbl(Label42.Caption) + .TextMatrix(i, 4), "#,##0.00")
                            Next
                        End With
                        Else
                    MsgBox "ไม่พบข้อมูลที่ต้องการแสดง", vbInformation + vbOKOnly, "ไม่มีข้อมูลที่ค้นหา"
                    Exit Sub
                    End If
                    rsG.MoveFirst
                    rsG1.MoveFirst
                    If rsG!IncYear <> "" Then
                        cbYear.Text = rsG!IncYear
                        cbMonth.Text = rsG!IncMonth
                    ElseIf rsG1!IncYear <> "" Then
                        cbYear.Text = rsG1!IncYear
                        cbMonth.Text = rsG1!IncMonth
                    Else
                        cbYear.ListIndex = 0
                        cbMonth.ListIndex = 0
                    End If
End Sub

Private Sub Command1_Click()
    cbYear3.ListIndex = 0
    cbMonth3.ListIndex = 0
    txtInWaste.Text = ""
    txtOutWaste.Text = ""
    grCost.ClearAllData
    grBypassCost.ClearAllData
    setGridCost
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Command3_Click()
                Dim Sql, tSql As String
                Dim rs As New ADODB.Recordset
                Dim i As Integer
                
                If cbYear4.ListIndex = 0 Then
                    MsgBox "เลือกปีก่อนค้นหา", vbInformation + vbOKOnly, "ค้นหาข้อมูล"
                    Exit Sub
                End If
                    lblSumStaff.Caption = "0"
                    lblSum.Caption = "0"
                    Sql = "SELECT * FROM vw_CostAcc " & vbCrLf
                    Sql = Sql & "where CostYear='" & cbYear4.Text & "' and CostMonth='" & cbMonth4.Text & "' order by id" & vbCrLf
                    Set rs = GetRS(Sql, adOpenForwardOnly, adLockOptimistic)
                    If rs.RecordCount > 0 Then
                        With grCost.getGridObj
                        grCost.ClearAllData
                            .Rows = rs.RecordCount - 1
                            For i = 1 To rs.RecordCount
                            If rs!Costdescid <= "6" Then
                                .TextMatrix(i, 0) = rs!Costdescid
                                .TextMatrix(i, 1) = rs!CostDesc
                                .TextMatrix(i, 2) = Format(rs!Cost, "#,##0.00")
                                .TextMatrix(i, 3) = rs!ID
                                lblSum.Caption = Format(CDbl(lblSum.Caption) + .TextMatrix(i, 2), "#,##0.00")
                            End If
                                rs.MoveNext
                            Next
                    End With
                rs.MoveFirst
                    Dim x, a As Integer
                    a = 1
                    For x = 1 To rs.RecordCount
                    If rs!Costdescid > "6" Then
                        With grBypassCost.getGridObj
                            .Rows = 3
                            .TextMatrix(a, 0) = rs!Costdescid
                            .TextMatrix(a, 1) = rs!CostDesc
                            .TextMatrix(a, 2) = Format(rs!Cost, "#,##0.00")
                            .TextMatrix(a, 3) = rs!ID
                            lblSumStaff.Caption = Format(CDbl(lblSumStaff.Caption) + .TextMatrix(a, 2), "#,##0.00")
                        End With
                    a = a + 1
                    End If
                    rs.MoveNext
                    Next
                    rs.MoveFirst
                            txtInWaste.Text = rs!InWaste
                            txtOutWaste.Text = rs!OutWaste
                            cbYear3.Text = rs!CostYear
                            cbMonth3.Text = rs!CostMonth
                        
                        Else
                    MsgBox "ไม่พบข้อมูลที่ต้องการแสดง", vbInformation + vbOKOnly, "ไม่มีข้อมูลที่ค้นหา"
                    End If
End Sub

Private Sub Command4_Click()
    Unload Me
End Sub

Private Sub Command5_Click()
    cbYear.ListIndex = 0
    cbMonth.ListIndex = 0
    grIncome.ClearAllData
    grIncomeCompany.ClearAllData
    SetGrid
End Sub

Private Sub Command6_Click()
Dim sqlstm As String
Dim RSC As New ADODB.Recordset
Dim i As Integer
    If cbYear3.ListIndex = 0 Then
        MsgBox "กรุณาระบุปี", vbExclamation + vbOKOnly, "บันทึกข้อมูลไม่ได้"
        Exit Sub
    End If
                sqlstm = "SELECT * FROM tbCost " & vbCrLf
            sqlstm = sqlstm & "where CostYear='" & cbYear3.Text & "' and CostMonth='" & cbMonth3.Text & "'" & vbCrLf
            Set RSC = GetRS(sqlstm, adOpenForwardOnly, adLockOptimistic)
                If RSC.RecordCount > 0 And grCost.getGridObj.TextMatrix(1, 3) = "" Then
                    MsgBox "ค้นพบข้อมูลปีและเดือนที่ระบุในฐานข้อมูล กรุณาตรวจสอบ", vbExclamation + vbOKOnly, "พบข้อมูลซ้ำ"
                    Exit Sub
                End If
        For i = 1 To grCost.getGridObj.Rows - 1
                Dim Sql, tSql As String
                Dim rsG As New ADODB.Recordset
                Dim x As Integer
            With grCost.getGridObj
                    Sql = "SELECT * FROM tbCost " & vbCrLf
                    Sql = Sql & "where ID='" & .TextMatrix(i, 3) & "' " & vbCrLf
                    Set rsG = GetRS(Sql, adOpenForwardOnly, adLockOptimistic)
                    If rsG.RecordCount > 0 Then
                        DBConnExc "Update tbCost Set Cost=" & CDbl(.TextMatrix(i, 2)) & ",InWaste=" & txtInWaste.Text & ",OutWaste=" & txtOutWaste.Text & " where ID='" & .TextMatrix(i, 3) & "' "
                    Else
                        tSql = "INSERT INTO tbCost(CostDescID,Cost,InWaste,OutWaste,CostYear,CostMonth) "
                        tSql = tSql + " Values (" & .TextMatrix(i, 0) & "," & CDbl(.TextMatrix(i, 2)) & ", " & txtInWaste.Text & ", " & txtOutWaste.Text & ",'" & cbYear3.Text & "','" & cbMonth3.Text & "')"
                        DBConnExc tSql
                    End If
            End With
        Next
        
                For i = 1 To grBypassCost.getGridObj.Rows - 1
            With grBypassCost.getGridObj
                    Sql = "SELECT * FROM tbCost " & vbCrLf
                    Sql = Sql & "where ID='" & .TextMatrix(i, 3) & "' " & vbCrLf
                    Set rsG = GetRS(Sql, adOpenForwardOnly, adLockOptimistic)
                    If rsG.RecordCount > 0 Then
                        DBConnExc "Update tbCost Set Cost=" & CDbl(.TextMatrix(i, 2)) & ",InWaste=" & txtInWaste.Text & ",OutWaste=" & txtOutWaste.Text & " where ID='" & .TextMatrix(i, 3) & "' "
                    Else
                        tSql = "INSERT INTO tbCost(CostDescID,Cost,InWaste,OutWaste,CostYear,CostMonth) "
                        tSql = tSql + " Values (" & .TextMatrix(i, 0) & "," & CDbl(.TextMatrix(i, 2)) & ", " & txtInWaste.Text & ", " & txtOutWaste.Text & ",'" & cbYear3.Text & "','" & cbMonth3.Text & "')"
                        DBConnExc tSql
                    End If
            End With
        Next

        MsgBox "บันทึกข้อมูลเรียบร้อย", vbInformation + vbOKOnly, "บันทึกข้อมูล"

End Sub

Private Sub CostOfService_Change()
'    If Not IsNumeric(CostOfService.Text) Then SendKeys "{backspace}"
End Sub
Private Sub CostOfService_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        With GrossProfit
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
        End With
            CostOfService.Text = Format(CostOfService.Text, "#,##0.00")
    End If
End Sub
Private Sub DE_Change()
'    If Not IsNumeric(DE.Text) Then SendKeys "{backspace}"
End Sub
Private Sub DE_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        With DebtRatio
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
        End With
            DE.Text = Format(DE.Text, "#,##0.0000")
    End If
End Sub
Private Sub DebtRatio_Change()
'    If Not IsNumeric(DebtRatio.Text) Then SendKeys "{backspace}"
End Sub
Private Sub DebtRatio_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        With BV
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
        End With
            DebtRatio.Text = Format(DebtRatio.Text, "#,##0.0000")
    End If
End Sub

Private Sub EBIT_Change()
'    If Not IsNumeric(EBIT.Text) Then SendKeys "{backspace}"
End Sub
Private Sub EBIT_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        With ShareofProfit
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
        End With
            EBIT.Text = Format(EBIT.Text, "#,##0.00")
    End If
End Sub
Private Sub EBITDA_Change()
'    If Not IsNumeric(EBITDA.Text) Then SendKeys "{backspace}"
End Sub
Private Sub EBITDA_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        With EBIT
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
        End With
            EBITDA.Text = Format(EBITDA.Text, "#,##0.00")
    End If
End Sub
Private Sub EndingCash_Change()
'    If Not IsNumeric(EndingCash.Text) Then SendKeys "{backspace}"
End Sub
Private Sub EndingCash_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
            EndingCash.Text = Format(EndingCash.Text, "#,##0.00")
        cmdSave.SetFocus
    End If
End Sub
Private Sub EPS_Change()
'    If Not IsNumeric(EPS.Text) Then SendKeys "{backspace}"
End Sub
Private Sub EPS_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        With Asset
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
        End With
            EPS.Text = Format(EPS.Text, "#,##0.0000")
    End If
End Sub
Private Sub Equity_Change()
'    If Not IsNumeric(Equity.Text) Then SendKeys "{backspace}"
End Sub
Private Sub Equity_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        With OperatingAct
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
        End With
            Equity.Text = Format(Equity.Text, "#,##0.00")
    End If
End Sub
Private Sub Expense_Change()
'    If Not IsNumeric(Expense.Text) Then SendKeys "{backspace}"
End Sub
Private Sub Expense_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        With EBITDA
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
        End With
            Expense.Text = Format(Expense.Text, "#,##0.00")
    End If
End Sub
Private Sub FinancingAct_Change()
'    If Not IsNumeric(FinancingAct.Text) Then SendKeys "{backspace}"
End Sub
Private Sub FinancingAct_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        With NetIncrease
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
        End With
            FinancingAct.Text = Format(FinancingAct.Text, "#,##0.00")
    End If
End Sub
Private Sub Form_Load()
    setObj
    SetGrid
    setGridCost
End Sub
Sub setGridCost()
    With grCost.getGridObj
        .Rows = 6
        .Cols = 4
        .TextMatrix(0, 1) = "ต้นทุน"
        .TextMatrix(0, 2) = "จำนวนเงิน"
        .TextMatrix(0, 3) = "ID"
        .ColWidth(0) = 0
        .ColWidth(1) = 4000
        .ColWidth(2) = 2000
        .ColWidth(3) = 0
    End With
                Dim tSql2 As String
                Dim RS2 As New ADODB.Recordset
                Dim i2 As Integer
                    tSql2 = "SELECT * FROM tbCostDesc where id<='6' and isdelete<>'Y' order by id" & vbCrLf
                    Set RS2 = GetRS(tSql2, adOpenForwardOnly, adLockOptimistic)
                    If RS2.RecordCount > 0 Then
                        With grCost.getGridObj
                            grCost.ClearAllData
                            .Rows = RS2.RecordCount + 1
                            For i2 = 1 To RS2.RecordCount
                            .TextMatrix(i2, 0) = RS2!ID
                            .TextMatrix(i2, 1) = RS2!CostDesc
                            .TextMatrix(i2, 2) = "0"
                            .TextMatrix(i2, 3) = "0"
                            RS2.MoveNext
                            Next
                        End With
                    End If
                    
    With grBypassCost.getGridObj
        .Rows = 2
        .Cols = 4
        .TextMatrix(0, 1) = "ต้นทุน"
        .TextMatrix(0, 2) = "จำนวนเงิน"
        .TextMatrix(0, 3) = "ID"
        .ColWidth(0) = 0
        .ColWidth(1) = 3500
        .ColWidth(2) = 1500
        .ColWidth(3) = 0
    End With
                Dim tSql3 As String
                Dim RS3 As New ADODB.Recordset
                Dim i3 As Integer
                    tSql3 = "SELECT * FROM tbCostDesc where id in('7','8') and isdelete<>'Y' order by id" & vbCrLf
                    Set RS3 = GetRS(tSql3, adOpenForwardOnly, adLockOptimistic)
                    If RS3.RecordCount > 0 Then
                        With grBypassCost.getGridObj
                            grBypassCost.ClearAllData
                            .Rows = RS3.RecordCount + 1
                            For i3 = 1 To RS3.RecordCount
                            .TextMatrix(i3, 0) = RS3!ID
                            .TextMatrix(i3, 1) = RS3!CostDesc
                            .TextMatrix(i3, 2) = "0"
                            .TextMatrix(i3, 3) = "0"
                            RS3.MoveNext
                            Next
                        End With
                    End If

End Sub

Sub SetGrid()
    With grIncome.getGridObj
        .Rows = 6
        .Cols = 6
        .TextMatrix(0, 1) = "รายได้"
        .TextMatrix(0, 2) = "จำนวนเงิน (Net)"
        .TextMatrix(0, 3) = "บริษัทเกี่ยวข้องกัน (BWG,BWC)"
        .TextMatrix(0, 4) = "Broker(PST GGE PAFU)"
        .TextMatrix(0, 5) = "ID"
        .ColWidth(0) = 0
        .ColWidth(1) = 3000
        .ColWidth(2) = 1800
        .ColWidth(3) = 2300
        .ColWidth(4) = 2300
        .ColWidth(5) = 0
    End With
                Dim tSql As String
                Dim rs As New ADODB.Recordset
                Dim i As Integer
                    tSql = "SELECT * FROM tbinc where isdelete<>'Y' " & vbCrLf
                    Set rs = GetRS(tSql, adOpenForwardOnly, adLockOptimistic)
                    If rs.RecordCount > 0 Then
                        With grIncome.getGridObj
                            grIncome.ClearAllData
                            .Rows = rs.RecordCount + 1
                            For i = 1 To rs.RecordCount
                            .TextMatrix(i, 0) = rs!IncID
                            .TextMatrix(i, 1) = rs!IncName
                            .TextMatrix(i, 2) = "0"
                            .TextMatrix(i, 3) = "0"
                            .TextMatrix(i, 4) = "0"
                            .TextMatrix(i, 5) = "0"
                            rs.MoveNext
                            Next
                        End With
                    End If
                    
    With grIncomeCompany.getGridObj
        .Rows = 6
        .Cols = 6
        .TextMatrix(0, 1) = "บริษัท"
        .TextMatrix(0, 2) = "รายได้ค่าบริการ"
        .TextMatrix(0, 3) = "สก.9"
        .TextMatrix(0, 4) = "รับจ้างรถขนส่ง"
        .TextMatrix(0, 5) = "ID"
        .ColWidth(0) = 0
        .ColWidth(1) = 3500
        .ColWidth(2) = 1900
        .ColWidth(3) = 1900
        .ColWidth(4) = 1900
        .ColWidth(5) = 0
    End With
                Dim tSql1 As String
                Dim RS1 As New ADODB.Recordset
                Dim i1 As Integer
                    tSql1 = "SELECT * FROM tbIncCompany where isdelete<>'Y' " & vbCrLf
                    Set RS1 = GetRS(tSql1, adOpenForwardOnly, adLockOptimistic)
                    If RS1.RecordCount > 0 Then
                        With grIncomeCompany.getGridObj
                            grIncomeCompany.ClearAllData
                            .Rows = RS1.RecordCount + 1
                            For i = 1 To RS1.RecordCount
                            .TextMatrix(i, 0) = RS1!IncID
                            .TextMatrix(i, 1) = RS1!CompanyName
                            .TextMatrix(i, 2) = "0"
                            .TextMatrix(i, 3) = "0"
                            .TextMatrix(i, 4) = "0"
                            .TextMatrix(i, 5) = "0"
                            RS1.MoveNext
                            Next
                        End With
                    End If
End Sub
Sub setObj()
Dim YY As String
Dim i As Integer
    RevenueBuring.Text = "0"
    Transportation.Text = "0"
    OtherIncome.Text = "0"
    CostOfService.Text = "0"
    GrossProfit.Text = "0"
    NetIncome.Text = "0"
    Expense.Text = "0"
    EBITDA.Text = "0"
    EBIT.Text = "0"
    ShareofProfit.Text = "0"
    NetProfit.Text = "0"
    ROE.Text = "0"
    ROA.Text = "0"
    DE.Text = "0"
    DebtRatio.Text = "0"
    BV.Text = "0"
    EPS.Text = "0"
    Asset.Text = "0"
    Liability.Text = "0"
    Equity.Text = "0"
    OperatingAct.Text = "0"
    InvestingAct.Text = "0"
    FinancingAct.Text = "0"
    NetIncrease.Text = "0"
    EndingCash.Text = "0"
    With cbQ
        .Clear
        .AddItem "เลือก"
        .AddItem "1"
        .AddItem "2"
        .AddItem "3"
        .AddItem "4"
        .ListIndex = 0
    End With
    If Year(Date) <= 2500 Then
        YY = Year(Date) + 543
    Else
        YY = Year(Date)
    End If
    lblyy.Caption = YY
    With cbYears
    .Clear
        For i = 2550 To YY
        Dim ListDet As String
            ListDet = CDbl(YY)
            .AddItem ListDet
            YY = YY - 1
        Next
        .ListIndex = 0
    End With
    With cbYear
    YY = lblyy.Caption
    .Clear
        .AddItem "เลือก"
        For i = 2550 To YY
        Dim ListDet1 As String
            ListDet1 = CDbl(YY)
            .AddItem ListDet1
            YY = YY - 1
        Next
        .ListIndex = 0
    End With
    With cbYear2
    YY = lblyy.Caption
    .Clear
        .AddItem "เลือก"
        For i = 2550 To YY
        Dim ListDet2 As String
            ListDet2 = CDbl(YY)
            .AddItem ListDet2
            YY = YY - 1
        Next
        .ListIndex = 0
    End With
    With cbYear3
    YY = lblyy.Caption
    .Clear
        .AddItem "เลือก"
        For i = 2550 To YY
        Dim ListDet3 As String
            ListDet3 = CDbl(YY)
            .AddItem ListDet3
            YY = YY - 1
        Next
        .ListIndex = 0
    End With
    With cbYear4
    YY = lblyy.Caption
    .Clear
        .AddItem "เลือก"
        For i = 2550 To YY
        Dim ListDet4 As String
            ListDet4 = CDbl(YY)
            .AddItem ListDet4
            YY = YY - 1
        Next
        .ListIndex = 0
    End With
   With cbMonth
        .Clear
        .AddItem "มกราคม"
        .AddItem "กุมภาพันธ์"
        .AddItem "มีนาคม"
        .AddItem "เมษายน"
        .AddItem "พฤษภาคม"
        .AddItem "มิถุนายน"
        .AddItem "กรกฎาคม"
        .AddItem "สิงหาคม"
        .AddItem "กันยายน"
        .AddItem "ตุลาคม"
        .AddItem "พฤศจิกายน"
        .AddItem "ธันวาคม"
        .ListIndex = 0
    End With
   With cbMonth2
        .Clear
        .AddItem "มกราคม"
        .AddItem "กุมภาพันธ์"
        .AddItem "มีนาคม"
        .AddItem "เมษายน"
        .AddItem "พฤษภาคม"
        .AddItem "มิถุนายน"
        .AddItem "กรกฎาคม"
        .AddItem "สิงหาคม"
        .AddItem "กันยายน"
        .AddItem "ตุลาคม"
        .AddItem "พฤศจิกายน"
        .AddItem "ธันวาคม"
        .ListIndex = 0
    End With
   With cbMonth3
        .Clear
        .AddItem "มกราคม"
        .AddItem "กุมภาพันธ์"
        .AddItem "มีนาคม"
        .AddItem "เมษายน"
        .AddItem "พฤษภาคม"
        .AddItem "มิถุนายน"
        .AddItem "กรกฎาคม"
        .AddItem "สิงหาคม"
        .AddItem "กันยายน"
        .AddItem "ตุลาคม"
        .AddItem "พฤศจิกายน"
        .AddItem "ธันวาคม"
        .ListIndex = 0
    End With
   With cbMonth4
        .Clear
        .AddItem "มกราคม"
        .AddItem "กุมภาพันธ์"
        .AddItem "มีนาคม"
        .AddItem "เมษายน"
        .AddItem "พฤษภาคม"
        .AddItem "มิถุนายน"
        .AddItem "กรกฎาคม"
        .AddItem "สิงหาคม"
        .AddItem "กันยายน"
        .AddItem "ตุลาคม"
        .AddItem "พฤศจิกายน"
        .AddItem "ธันวาคม"
        .ListIndex = 0
    End With
End Sub

Private Sub grBypassCost_BeforeUpdateValue(NewValue As String, row As Integer, col As Integer, Cancel As Boolean)
    If (col = 2) And IsNumeric(NewValue) = False Then
        MsgBox "กรุณาใส่ข้อมูลเป็นตัวเลขเท่านั้น", vbExclamation
        Cancel = True
    End If

End Sub

Private Sub grBypassCost_RolColChanged(row As Integer, col As Integer)
    grBypassCost.ColEnabled(1) = False
End Sub

Private Sub grBypassCost_UpdateValue(row As Integer, col As Integer)
Dim i, x, a As Integer
        Select Case col
            Case 2
                lblSumStaff.Caption = "0"
                With grBypassCost.getGridObj
                    For i = 1 To .Rows - 1
                        .TextMatrix(i, 2) = Format(.TextMatrix(i, 2), "#,##0.00")
                        lblSumStaff.Caption = Format(CDbl(lblSumStaff.Caption) + .TextMatrix(i, 2), "#,##0.00")
                    Next
                End With
            End Select
End Sub

Private Sub grCost_BeforeUpdateValue(NewValue As String, row As Integer, col As Integer, Cancel As Boolean)
    If (col = 2 Or col = 3 Or col = 4) And IsNumeric(NewValue) = False Then
        MsgBox "กรุณาใส่ข้อมูลเป็นตัวเลขเท่านั้น", vbExclamation
        Cancel = True
    End If
End Sub
Private Sub grCost_RolColChanged(row As Integer, col As Integer)
    grCost.ColEnabled(1) = False
End Sub
Private Sub grCost_UpdateValue(row As Integer, col As Integer)
Dim i, x, a As Integer
        Select Case col
            Case 2
                lblSum.Caption = "0"
                With grCost.getGridObj
                    For i = 1 To .Rows - 1
                        .TextMatrix(i, 2) = Format(.TextMatrix(i, 2), "#,##0.00")
                        lblSum.Caption = Format(CDbl(lblSum.Caption) + .TextMatrix(i, 2), "#,##0.00")
                    Next
                End With
            End Select
End Sub

Private Sub grIncome_BeforeUpdateValue(NewValue As String, row As Integer, col As Integer, Cancel As Boolean)
    If (col = 2 Or col = 3 Or col = 4) And IsNumeric(NewValue) = False Then
        MsgBox "กรุณาใส่ข้อมูลเป็นตัวเลขเท่านั้น", vbExclamation
        Cancel = True
    End If
End Sub

Private Sub grIncome_RolColChanged(row As Integer, col As Integer)
    grIncome.ColEnabled(1) = False
End Sub

Private Sub grIncome_UpdateValue(row As Integer, col As Integer)
Dim i, x, a As Integer
        Select Case col
            Case 2
                Label27.Caption = "0"
'                Label28.Caption = "0"
'                Label37.Caption = "0"
'                Label38.Caption = "0"
                With grIncome.getGridObj
                        .TextMatrix(row, 2) = Format(.TextMatrix(row, 2), "#,##0.00")
                    For i = 1 To .Rows - 1
                        Label27.Caption = Format(CDbl(Label27.Caption) + .TextMatrix(i, 2), "#,##0.00")
                    Next
'                    For x = 1 To .Rows - 1
'                        .TextMatrix(x, 3) = Format((CDbl(.TextMatrix(x, 2)) / CDbl(Label27.Caption)) * 100, "#0")
'                    Next
'                    For a = 1 To .Rows - 1
'                        Label28.Caption = CDbl(Label28.Caption) + .TextMatrix(a, 3)
'                        .TextMatrix(a, 8) = CDbl(.TextMatrix(a, 2)) - CDbl(.TextMatrix(a, 4)) - CDbl(.TextMatrix(a, 6))
'                        Label37.Caption = CDbl(Label37.Caption) + .TextMatrix(a, 8)
'                        .TextMatrix(a, 9) = Format((CDbl(.TextMatrix(a, 8)) / CDbl(Label37.Caption)) * 100, "#0")
'                         Label38.Caption = CDbl(Label38.Caption) + .TextMatrix(a, 9)
'                    Next
                End With
            Case 3
                Label29.Caption = "0"
'                Label30.Caption = "0"
'                Label37.Caption = "0"
'                Label38.Caption = "0"
                With grIncome.getGridObj
                        .TextMatrix(row, 3) = Format(.TextMatrix(row, 3), "#,##0.00")
                    For i = 1 To .Rows - 1
                        Label29.Caption = Format(CDbl(Label29.Caption) + .TextMatrix(i, 3), "#,##0.00")
                    Next
'                    For x = 1 To .Rows - 1
'                        .TextMatrix(x, 5) = Format((CDbl(.TextMatrix(x, 4)) / CDbl(Label29.Caption)) * 100, "#0")
'                    Next
'                    For a = 1 To .Rows - 1
'                        Label30.Caption = CDbl(Label30.Caption) + .TextMatrix(a, 5)
'                        .TextMatrix(a, 8) = CDbl(.TextMatrix(a, 2)) - CDbl(.TextMatrix(a, 4)) - CDbl(.TextMatrix(a, 6))
'                        Label37.Caption = CDbl(Label37.Caption) + .TextMatrix(a, 8)
'                        .TextMatrix(a, 9) = Format((CDbl(.TextMatrix(a, 8)) / CDbl(Label37.Caption)) * 100, "#0")
'                         Label38.Caption = CDbl(Label38.Caption) + .TextMatrix(a, 9)
'                    Next
                End With
            Case 4
                Label35.Caption = "0"
'                Label36.Caption = "0"
'                Label37.Caption = "0"
'                Label38.Caption = "0"
                With grIncome.getGridObj
                        .TextMatrix(row, 4) = Format(.TextMatrix(row, 4), "#,##0.00")
                    For i = 1 To .Rows - 1
                        Label35.Caption = Format(CDbl(Label35.Caption) + .TextMatrix(i, 4), "#,##0.00")
                  Next
'                    For x = 1 To .Rows - 1
'                        .TextMatrix(x, 7) = Format((CDbl(.TextMatrix(x, 6)) / CDbl(Label35.Caption)) * 100, "#0")
'                    Next
'                    For a = 1 To .Rows - 1
'                        Label36.Caption = CDbl(Label36.Caption) + .TextMatrix(a, 7)
'                        .TextMatrix(a, 8) = CDbl(.TextMatrix(a, 2)) - CDbl(.TextMatrix(a, 4)) - CDbl(.TextMatrix(a, 6))
'                        Label37.Caption = CDbl(Label37.Caption) + .TextMatrix(a, 8)
'                        .TextMatrix(a, 9) = Format((CDbl(.TextMatrix(a, 8)) / CDbl(Label37.Caption)) * 100, "#0")
'                         Label38.Caption = CDbl(Label38.Caption) + .TextMatrix(a, 9)
'                   Next
                End With
                End Select
End Sub
Private Sub grIncomeCompany_BeforeUpdateValue(NewValue As String, row As Integer, col As Integer, Cancel As Boolean)
    If (col = 2 Or col = 3 Or col = 4) And IsNumeric(NewValue) = False Then
        MsgBox "กรุณาใส่ข้อมูลเป็นตัวเลขเท่านั้น", vbExclamation
        Cancel = True
    End If
End Sub

Private Sub grIncomeCompany_RolColChanged(row As Integer, col As Integer)
    grIncomeCompany.ColEnabled(1) = False
End Sub

Private Sub grIncomeCompany_UpdateValue(row As Integer, col As Integer)
On Error Resume Next
Dim i, x, a As Integer
        Select Case col
            Case 2
                Label46.Caption = "0"
'                Label45.Caption = "0"
                With grIncomeCompany.getGridObj
                        .TextMatrix(row, 2) = Format(.TextMatrix(row, 2), "#,##0.00")
                    For i = 1 To .Rows - 1
                        Label46.Caption = Format(CDbl(Label46.Caption) + .TextMatrix(i, 2), "#,##0.00")
                    Next
'                    For x = 1 To .Rows - 1
'                        .TextMatrix(x, 3) = Format((CDbl(.TextMatrix(x, 2)) / CDbl(Label27.Caption)) * 100, "#0.#0")
'                    Next
'                    For a = 1 To .Rows - 1
'                        .TextMatrix(a, 9) = Format((CDbl(.TextMatrix(a, 8)) / CDbl(Label27.Caption)) * 100, "#0.#0")
'                        Label45.Caption = Format((CDbl(Label46.Caption) / CDbl(Label27.Caption)) * 100, "#0.#0")
'                        Label40.Caption = CDbl(Label46.Caption) + CDbl(Label44.Caption) + CDbl(Label42.Caption)
'                        .TextMatrix(a, 8) = CDbl(.TextMatrix(a, 2)) + CDbl(.TextMatrix(a, 4)) + CDbl(.TextMatrix(a, 6))
'                    Next
                End With
            Case 3
                Label44.Caption = "0"
'                Label43.Caption = "0"
                With grIncomeCompany.getGridObj
                        .TextMatrix(row, 3) = Format(.TextMatrix(row, 3), "#,##0.00")
                    For i = 1 To .Rows - 1
                        Label44.Caption = Format(CDbl(Label44.Caption) + .TextMatrix(i, 3), "#,##0.00")
                    Next
'                    For x = 1 To .Rows - 1
'                        .TextMatrix(x, 5) = Format((CDbl(.TextMatrix(x, 4)) / CDbl(Label27.Caption)) * 100, "#0.#0")
'                    Next
'                    For a = 1 To .Rows - 1
'                        .TextMatrix(a, 9) = Format((CDbl(.TextMatrix(a, 8)) / CDbl(Label27.Caption)) * 100, "#0.#0")
'                        .TextMatrix(a, 8) = CDbl(.TextMatrix(a, 2)) + CDbl(.TextMatrix(a, 4)) + CDbl(.TextMatrix(a, 6))
'                        Label43.Caption = Format((CDbl(Label44.Caption) / CDbl(Label27.Caption)) * 100, "#0.#0")
'                        Label40.Caption = CDbl(Label46.Caption) + CDbl(Label44.Caption) + CDbl(Label42.Caption)
'                    Next
                End With
            Case 4
                Label42.Caption = "0"
'                Label41.Caption = "0"
'                Label40.Caption = "0"
'                Label39.Caption = "0"
                With grIncomeCompany.getGridObj
                        .TextMatrix(row, 4) = Format(.TextMatrix(row, 4), "#,##0.00")
                    For i = 1 To .Rows - 1
                        Label42.Caption = CDbl(Label42.Caption) + .TextMatrix(i, 4)
                    Next
'                    For x = 1 To .Rows - 1
'                        .TextMatrix(x, 7) = Format((CDbl(.TextMatrix(x, 6)) / CDbl(Label27.Caption)) * 100, "#0.#0")
'                    Next
'                    For a = 1 To .Rows - 1
'                        Label41.Caption = CDbl(Label41.Caption) + .TextMatrix(a, 7)
'                        .TextMatrix(a, 8) = CDbl(.TextMatrix(a, 2)) + CDbl(.TextMatrix(a, 4)) + CDbl(.TextMatrix(a, 6))
'                        Label40.Caption = CDbl(Label46.Caption) + CDbl(Label44.Caption) + CDbl(Label42.Caption)
'                        .TextMatrix(a, 9) = Format((CDbl(.TextMatrix(a, 8)) / CDbl(Label27.Caption)) * 100, "#0.#0")
'                        If Val(Label27.Caption) > 0 Then
'                            Label39.Caption = Format((CDbl(Label40.Caption) / CDbl(Label27.Caption)) * 100, "#0.#0")
'                        End If
'                    Next
                End With
                End Select

End Sub

Private Sub GrossProfit_Change()
'    If Not IsNumeric(GrossProfit.Text) Then SendKeys "{backspace}"
End Sub
Private Sub GrossProfit_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        With Expense
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
        End With
            GrossProfit.Text = Format(GrossProfit.Text, "#,##0.00")
    End If
End Sub
Private Sub InvestingAct_Change()
'    If Not IsNumeric(InvestingAct.Text) Then SendKeys "{backspace}"
End Sub
Private Sub InvestingAct_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        With FinancingAct
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
        End With
            InvestingAct.Text = Format(InvestingAct.Text, "#,##0.00")
    End If
End Sub
Private Sub Liability_Change()
'    If Not IsNumeric(Liability.Text) Then SendKeys "{backspace}"
End Sub
Private Sub Liability_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        With Equity
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
        End With
            Liability.Text = Format(Liability.Text, "#,##0.00")
    End If
End Sub

Private Sub NetIncrease_Change()
'    If Not IsNumeric(NetIncrease.Text) Then SendKeys "{backspace}"
End Sub
Private Sub NetIncrease_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        With EndingCash
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
        End With
            NetIncrease.Text = Format(NetIncrease.Text, "#,##0.00")
    End If
End Sub
Private Sub NetProfit_Change()
'    If Not IsNumeric(NetProfit.Text) Then SendKeys "{backspace}"
End Sub
Private Sub NetProfit_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        With ROE
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
        End With
            NetProfit.Text = Format(NetProfit.Text, "#,##0.00")
    End If
End Sub
Private Sub OperatingAct_Change()
'    If Not IsNumeric(OperatingAct.Text) Then SendKeys "{backspace}"
End Sub
Private Sub OperatingAct_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        With InvestingAct
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
        End With
            OperatingAct.Text = Format(OperatingAct.Text, "#,##0.00")
    End If
End Sub
Private Sub OtherIncome_Change()
    If Not IsNumeric(OtherIncome.Text) Then SendKeys "{backspace}"
End Sub
Private Sub OtherIncome_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        NetIncome.Text = Format(CDbl(Trim(RevenueBuring.Text)) + CDbl(Trim(Transportation.Text)) + CDbl(Trim(OtherIncome)), "#,##0")
        With CostOfService
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
        End With
            OtherIncome.Text = Format(OtherIncome.Text, "#,##0.00")
    End If
End Sub
Private Sub RevenueBuring_Change()
    If Not IsNumeric(RevenueBuring.Text) Then SendKeys "{backspace}"
End Sub
Private Sub RevenueBuring_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        With Transportation
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
        End With
            RevenueBuring.Text = Format(RevenueBuring.Text, "#,##0.00")
    End If
End Sub
Private Sub ROA_Change()
'    If Not IsNumeric(ROA.Text) Then SendKeys "{backspace}"
End Sub
Private Sub ROA_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        With DE
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
        End With
            ROA.Text = Format(ROA.Text, "#,##0.0000")
    End If
End Sub
Private Sub ROE_Change()
'    If Not IsNumeric(ROE.Text) Then SendKeys "{backspace}"
End Sub
Private Sub ROE_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        With ROA
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
        End With
            ROE.Text = Format(ROE.Text, "#,##0.0000")
    End If
End Sub
Private Sub ShareofProfit_Change()
'    If Not IsNumeric(ShareofProfit.Text) Then SendKeys "{backspace}"
End Sub
Private Sub ShareofProfit_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        With NetProfit
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
        End With
            ShareofProfit.Text = Format(ShareofProfit.Text, "#,##0.00")
    End If
End Sub
Private Sub Transportation_Change()
    If Not IsNumeric(Transportation.Text) Then SendKeys "{backspace}"
End Sub
Private Sub Transportation_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        With OtherIncome
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
        End With
            Transportation.Text = Format(Transportation.Text, "#,##0.00")
    End If
End Sub
Private Sub txtQYear_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        With RevenueBuring
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
        End With
    End If
End Sub
