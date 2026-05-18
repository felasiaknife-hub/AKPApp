VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmCustCalendar 
   Caption         =   "ลูกค้าเข้าเยี่ยมชมประจำเดือน"
   ClientHeight    =   9315
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11625
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9315
   ScaleWidth      =   11625
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   525
      Left            =   30
      TabIndex        =   168
      Top             =   -75
      Width           =   11565
      Begin VB.ComboBox CboType 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "FrmCustCalendar.frx":0000
         Left            =   2790
         List            =   "FrmCustCalendar.frx":0010
         TabIndex        =   190
         Top             =   150
         Width           =   1530
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "ปิด"
         Height          =   375
         Left            =   10695
         Style           =   1  'Graphical
         TabIndex        =   175
         Top             =   120
         Width           =   840
      End
      Begin VB.ComboBox cboMonth 
         Height          =   315
         Left            =   7500
         Style           =   2  'Dropdown List
         TabIndex        =   174
         Top             =   150
         Width           =   1455
      End
      Begin VB.ComboBox cboYear 
         Height          =   315
         Left            =   8970
         Style           =   2  'Dropdown List
         TabIndex        =   173
         Top             =   150
         Width           =   825
      End
      Begin VB.CommandButton Command1 
         Caption         =   "เพิ่มข้อมูล"
         Height          =   375
         Left            =   9840
         Style           =   1  'Graphical
         TabIndex        =   172
         Top             =   120
         Width           =   840
      End
      Begin VB.TextBox txtSale 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   615
         TabIndex        =   171
         Top             =   -45
         Visible         =   0   'False
         Width           =   1425
      End
      Begin VB.TextBox txtCustomer 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   4770
         TabIndex        =   170
         Top             =   150
         Width           =   1290
      End
      Begin VB.ComboBox cboCustType 
         Height          =   315
         ItemData        =   "FrmCustCalendar.frx":0058
         Left            =   6450
         List            =   "FrmCustCalendar.frx":005A
         TabIndex        =   169
         Top             =   150
         Width           =   1065
      End
      Begin WasteManagment.ctlDate dtToday 
         Height          =   345
         Left            =   6675
         TabIndex        =   176
         Top             =   420
         Visible         =   0   'False
         Width           =   1755
         _extentx        =   3096
         _extenty        =   609
         backcolor       =   -2147483633
      End
      Begin VB.Label aaaa 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "วัตถุประสงค์"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   1935
         TabIndex        =   189
         Top             =   225
         Width           =   840
      End
      Begin VB.Label lbToday 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "XXXXXXXXX"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   90
         TabIndex        =   180
         Top             =   180
         Width           =   1230
      End
      Begin VB.Label aaa 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Sale"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   42
         Left            =   225
         TabIndex        =   179
         Top             =   15
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.Label aaaa 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "ลูกค้า"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   43
         Left            =   4335
         TabIndex        =   178
         Top             =   210
         Width           =   375
      End
      Begin VB.Label aaa 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "ฐาน"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   6120
         TabIndex        =   177
         Top             =   195
         Width           =   285
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H80000008&
      Height          =   1455
      Index           =   0
      Left            =   30
      ScaleHeight     =   1425
      ScaleWidth      =   1635
      TabIndex        =   164
      Top             =   660
      Width           =   1665
      Begin MSComctlLib.ListView lvResult 
         Height          =   1155
         Index           =   0
         Left            =   0
         TabIndex        =   165
         TabStop         =   0   'False
         Top             =   270
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   2037
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   12648447
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Code"
            Object.Width           =   2355
         EndProperty
      End
      Begin VB.Label lbJobCount 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         Caption         =   "20 Jobs"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   0
         Left            =   30
         TabIndex        =   167
         Top             =   30
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   1470
         TabIndex        =   166
         Top             =   30
         Width           =   105
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H80000008&
      Height          =   1455
      Index           =   1
      Left            =   1680
      ScaleHeight     =   1425
      ScaleWidth      =   1635
      TabIndex        =   160
      Top             =   660
      Width           =   1665
      Begin MSComctlLib.ListView lvResult 
         Height          =   1155
         Index           =   1
         Left            =   0
         TabIndex        =   161
         TabStop         =   0   'False
         Top             =   270
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   2037
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   12648447
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Code"
            Object.Width           =   2355
         EndProperty
      End
      Begin VB.Label lbJobCount 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         Caption         =   "20 Jobs"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   1
         Left            =   30
         TabIndex        =   163
         Top             =   30
         Width           =   1395
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   1470
         TabIndex        =   162
         Top             =   30
         Width           =   105
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H80000008&
      Height          =   1455
      Index           =   2
      Left            =   3330
      ScaleHeight     =   1425
      ScaleWidth      =   1635
      TabIndex        =   156
      Top             =   660
      Width           =   1665
      Begin MSComctlLib.ListView lvResult 
         Height          =   1155
         Index           =   2
         Left            =   0
         TabIndex        =   157
         TabStop         =   0   'False
         Top             =   270
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   2037
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   12648447
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Code"
            Object.Width           =   2355
         EndProperty
      End
      Begin VB.Label lbJobCount 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         Caption         =   "20 Jobs"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   2
         Left            =   30
         TabIndex        =   159
         Top             =   30
         Width           =   1395
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   1470
         TabIndex        =   158
         Top             =   30
         Width           =   105
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H80000008&
      Height          =   1455
      Index           =   3
      Left            =   4980
      ScaleHeight     =   1425
      ScaleWidth      =   1635
      TabIndex        =   152
      Top             =   660
      Width           =   1665
      Begin MSComctlLib.ListView lvResult 
         Height          =   1155
         Index           =   3
         Left            =   0
         TabIndex        =   153
         TabStop         =   0   'False
         Top             =   270
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   2037
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   12648447
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Code"
            Object.Width           =   2355
         EndProperty
      End
      Begin VB.Label lbJobCount 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         Caption         =   "20 Jobs"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   3
         Left            =   30
         TabIndex        =   155
         Top             =   30
         Width           =   1395
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   3
         Left            =   1470
         TabIndex        =   154
         Top             =   30
         Width           =   105
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H80000008&
      Height          =   1455
      Index           =   4
      Left            =   6630
      ScaleHeight     =   1425
      ScaleWidth      =   1635
      TabIndex        =   148
      Top             =   660
      Width           =   1665
      Begin MSComctlLib.ListView lvResult 
         Height          =   1155
         Index           =   4
         Left            =   0
         TabIndex        =   149
         TabStop         =   0   'False
         Top             =   270
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   2037
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   12648447
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Code"
            Object.Width           =   2355
         EndProperty
      End
      Begin VB.Label lbJobCount 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         Caption         =   "20 Jobs"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   4
         Left            =   30
         TabIndex        =   151
         Top             =   30
         Width           =   1395
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   4
         Left            =   1470
         TabIndex        =   150
         Top             =   30
         Width           =   105
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H80000008&
      Height          =   1455
      Index           =   5
      Left            =   8280
      ScaleHeight     =   1425
      ScaleWidth      =   1635
      TabIndex        =   144
      Top             =   660
      Width           =   1665
      Begin MSComctlLib.ListView lvResult 
         Height          =   1155
         Index           =   5
         Left            =   0
         TabIndex        =   145
         TabStop         =   0   'False
         Top             =   270
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   2037
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   12648447
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Code"
            Object.Width           =   2355
         EndProperty
      End
      Begin VB.Label lbJobCount 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         Caption         =   "20 Jobs"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   5
         Left            =   30
         TabIndex        =   147
         Top             =   30
         Width           =   1395
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   5
         Left            =   1470
         TabIndex        =   146
         Top             =   30
         Width           =   105
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H80000008&
      Height          =   1455
      Index           =   6
      Left            =   9930
      ScaleHeight     =   1425
      ScaleWidth      =   1635
      TabIndex        =   140
      Top             =   660
      Width           =   1665
      Begin MSComctlLib.ListView lvResult 
         Height          =   1155
         Index           =   6
         Left            =   0
         TabIndex        =   141
         TabStop         =   0   'False
         Top             =   270
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   2037
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   12648447
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Code"
            Object.Width           =   2355
         EndProperty
      End
      Begin VB.Label lbJobCount 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         Caption         =   "20 Jobs"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   6
         Left            =   30
         TabIndex        =   143
         Top             =   30
         Width           =   1395
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   6
         Left            =   1470
         TabIndex        =   142
         Top             =   30
         Width           =   105
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H80000008&
      Height          =   1455
      Index           =   7
      Left            =   30
      ScaleHeight     =   1425
      ScaleWidth      =   1635
      TabIndex        =   136
      Top             =   2100
      Width           =   1665
      Begin MSComctlLib.ListView lvResult 
         Height          =   1155
         Index           =   7
         Left            =   0
         TabIndex        =   137
         TabStop         =   0   'False
         Top             =   270
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   2037
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   12648447
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Code"
            Object.Width           =   2355
         EndProperty
      End
      Begin VB.Label lbJobCount 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         Caption         =   "20 Jobs"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   7
         Left            =   30
         TabIndex        =   139
         Top             =   30
         Width           =   1395
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   7
         Left            =   1470
         TabIndex        =   138
         Top             =   30
         Width           =   105
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H80000008&
      Height          =   1455
      Index           =   8
      Left            =   1680
      ScaleHeight     =   1425
      ScaleWidth      =   1635
      TabIndex        =   132
      Top             =   2100
      Width           =   1665
      Begin MSComctlLib.ListView lvResult 
         Height          =   1155
         Index           =   8
         Left            =   0
         TabIndex        =   133
         TabStop         =   0   'False
         Top             =   270
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   2037
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   12648447
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Code"
            Object.Width           =   2355
         EndProperty
      End
      Begin VB.Label lbJobCount 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         Caption         =   "20 Jobs"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   8
         Left            =   30
         TabIndex        =   135
         Top             =   30
         Width           =   1395
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   8
         Left            =   1470
         TabIndex        =   134
         Top             =   30
         Width           =   105
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H80000008&
      Height          =   1455
      Index           =   9
      Left            =   3330
      ScaleHeight     =   1425
      ScaleWidth      =   1635
      TabIndex        =   128
      Top             =   2100
      Width           =   1665
      Begin MSComctlLib.ListView lvResult 
         Height          =   1155
         Index           =   9
         Left            =   0
         TabIndex        =   129
         TabStop         =   0   'False
         Top             =   270
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   2037
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   12648447
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Code"
            Object.Width           =   2355
         EndProperty
      End
      Begin VB.Label lbJobCount 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         Caption         =   "20 Jobs"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   9
         Left            =   30
         TabIndex        =   131
         Top             =   30
         Width           =   1395
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   9
         Left            =   1470
         TabIndex        =   130
         Top             =   30
         Width           =   105
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H80000008&
      Height          =   1455
      Index           =   10
      Left            =   4980
      ScaleHeight     =   1425
      ScaleWidth      =   1635
      TabIndex        =   124
      Top             =   2100
      Width           =   1665
      Begin MSComctlLib.ListView lvResult 
         Height          =   1155
         Index           =   10
         Left            =   0
         TabIndex        =   125
         TabStop         =   0   'False
         Top             =   240
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   2037
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   12648447
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Code"
            Object.Width           =   2355
         EndProperty
      End
      Begin VB.Label lbJobCount 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         Caption         =   "20 Jobs"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   10
         Left            =   30
         TabIndex        =   127
         Top             =   30
         Width           =   1395
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   10
         Left            =   1470
         TabIndex        =   126
         Top             =   30
         Width           =   105
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H80000008&
      Height          =   1455
      Index           =   11
      Left            =   6630
      ScaleHeight     =   1425
      ScaleWidth      =   1635
      TabIndex        =   120
      Top             =   2100
      Width           =   1665
      Begin MSComctlLib.ListView lvResult 
         Height          =   1155
         Index           =   11
         Left            =   0
         TabIndex        =   121
         TabStop         =   0   'False
         Top             =   270
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   2037
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   12648447
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "code"
            Object.Width           =   2355
         EndProperty
      End
      Begin VB.Label lbJobCount 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         Caption         =   "20 Jobs"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   11
         Left            =   30
         TabIndex        =   123
         Top             =   30
         Width           =   1395
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   11
         Left            =   1470
         TabIndex        =   122
         Top             =   30
         Width           =   105
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H80000008&
      Height          =   1455
      Index           =   12
      Left            =   8280
      ScaleHeight     =   1425
      ScaleWidth      =   1635
      TabIndex        =   116
      Top             =   2100
      Width           =   1665
      Begin MSComctlLib.ListView lvResult 
         Height          =   1155
         Index           =   12
         Left            =   0
         TabIndex        =   117
         TabStop         =   0   'False
         Top             =   270
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   2037
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   12648447
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Code"
            Object.Width           =   2355
         EndProperty
      End
      Begin VB.Label lbJobCount 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         Caption         =   "20 Jobs"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   12
         Left            =   30
         TabIndex        =   119
         Top             =   30
         Width           =   1395
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   12
         Left            =   1470
         TabIndex        =   118
         Top             =   30
         Width           =   105
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H80000008&
      Height          =   1455
      Index           =   13
      Left            =   9930
      ScaleHeight     =   1425
      ScaleWidth      =   1635
      TabIndex        =   112
      Top             =   2100
      Width           =   1665
      Begin MSComctlLib.ListView lvResult 
         Height          =   1155
         Index           =   13
         Left            =   0
         TabIndex        =   113
         TabStop         =   0   'False
         Top             =   270
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   2037
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   12648447
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Code"
            Object.Width           =   2355
         EndProperty
      End
      Begin VB.Label lbJobCount 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         Caption         =   "20 Jobs"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   13
         Left            =   30
         TabIndex        =   115
         Top             =   30
         Width           =   1395
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   13
         Left            =   1470
         TabIndex        =   114
         Top             =   30
         Width           =   105
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H80000008&
      Height          =   1455
      Index           =   14
      Left            =   30
      ScaleHeight     =   1425
      ScaleWidth      =   1635
      TabIndex        =   108
      Top             =   3540
      Width           =   1665
      Begin MSComctlLib.ListView lvResult 
         Height          =   1155
         Index           =   14
         Left            =   0
         TabIndex        =   109
         TabStop         =   0   'False
         Top             =   270
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   2037
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   12648447
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Code"
            Object.Width           =   2355
         EndProperty
      End
      Begin VB.Label lbJobCount 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         Caption         =   "20 Jobs"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   14
         Left            =   30
         TabIndex        =   111
         Top             =   30
         Width           =   1395
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   14
         Left            =   1470
         TabIndex        =   110
         Top             =   30
         Width           =   105
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H80000008&
      Height          =   1455
      Index           =   15
      Left            =   1680
      ScaleHeight     =   1425
      ScaleWidth      =   1635
      TabIndex        =   104
      Top             =   3540
      Width           =   1665
      Begin MSComctlLib.ListView lvResult 
         Height          =   1155
         Index           =   15
         Left            =   0
         TabIndex        =   105
         TabStop         =   0   'False
         Top             =   270
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   2037
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   12648447
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Code"
            Object.Width           =   2355
         EndProperty
      End
      Begin VB.Label lbJobCount 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         Caption         =   "20 Jobs"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   15
         Left            =   30
         TabIndex        =   107
         Top             =   30
         Width           =   1395
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   15
         Left            =   1470
         TabIndex        =   106
         Top             =   30
         Width           =   105
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H80000008&
      Height          =   1455
      Index           =   16
      Left            =   3330
      ScaleHeight     =   1425
      ScaleWidth      =   1635
      TabIndex        =   100
      Top             =   3540
      Width           =   1665
      Begin MSComctlLib.ListView lvResult 
         Height          =   1155
         Index           =   16
         Left            =   0
         TabIndex        =   101
         TabStop         =   0   'False
         Top             =   270
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   2037
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   12648447
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Code"
            Object.Width           =   2355
         EndProperty
      End
      Begin VB.Label lbJobCount 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         Caption         =   "20 Jobs"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   16
         Left            =   30
         TabIndex        =   103
         Top             =   30
         Width           =   1395
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   16
         Left            =   1470
         TabIndex        =   102
         Top             =   30
         Width           =   105
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H80000008&
      Height          =   1455
      Index           =   17
      Left            =   4980
      ScaleHeight     =   1425
      ScaleWidth      =   1635
      TabIndex        =   96
      Top             =   3540
      Width           =   1665
      Begin MSComctlLib.ListView lvResult 
         Height          =   1155
         Index           =   17
         Left            =   0
         TabIndex        =   97
         TabStop         =   0   'False
         Top             =   270
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   2037
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   12648447
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Code"
            Object.Width           =   2355
         EndProperty
      End
      Begin VB.Label lbJobCount 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         Caption         =   "20 Jobs"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   17
         Left            =   30
         TabIndex        =   99
         Top             =   30
         Width           =   1395
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   17
         Left            =   1470
         TabIndex        =   98
         Top             =   30
         Width           =   105
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H80000008&
      Height          =   1455
      Index           =   18
      Left            =   6630
      ScaleHeight     =   1425
      ScaleWidth      =   1635
      TabIndex        =   92
      Top             =   3540
      Width           =   1665
      Begin MSComctlLib.ListView lvResult 
         Height          =   1155
         Index           =   18
         Left            =   0
         TabIndex        =   93
         TabStop         =   0   'False
         Top             =   270
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   2037
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   12648447
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Code"
            Object.Width           =   2355
         EndProperty
      End
      Begin VB.Label lbJobCount 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         Caption         =   "20 Jobs"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   18
         Left            =   30
         TabIndex        =   95
         Top             =   30
         Width           =   1395
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   18
         Left            =   1470
         TabIndex        =   94
         Top             =   30
         Width           =   105
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H80000008&
      Height          =   1455
      Index           =   19
      Left            =   8280
      ScaleHeight     =   1425
      ScaleWidth      =   1635
      TabIndex        =   88
      Top             =   3540
      Width           =   1665
      Begin MSComctlLib.ListView lvResult 
         Height          =   1155
         Index           =   19
         Left            =   0
         TabIndex        =   89
         TabStop         =   0   'False
         Top             =   270
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   2037
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   12648447
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Code"
            Object.Width           =   2355
         EndProperty
      End
      Begin VB.Label lbJobCount 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         Caption         =   "20 Jobs"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   19
         Left            =   30
         TabIndex        =   91
         Top             =   30
         Width           =   1395
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   19
         Left            =   1470
         TabIndex        =   90
         Top             =   30
         Width           =   105
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H80000008&
      Height          =   1455
      Index           =   20
      Left            =   9930
      ScaleHeight     =   1425
      ScaleWidth      =   1635
      TabIndex        =   84
      Top             =   3540
      Width           =   1665
      Begin MSComctlLib.ListView lvResult 
         Height          =   1155
         Index           =   20
         Left            =   0
         TabIndex        =   85
         TabStop         =   0   'False
         Top             =   270
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   2037
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   12648447
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Code"
            Object.Width           =   2355
         EndProperty
      End
      Begin VB.Label lbJobCount 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         Caption         =   "20 Jobs"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   20
         Left            =   30
         TabIndex        =   87
         Top             =   30
         Width           =   1395
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   20
         Left            =   1470
         TabIndex        =   86
         Top             =   30
         Width           =   105
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H80000008&
      Height          =   1455
      Index           =   21
      Left            =   30
      ScaleHeight     =   1425
      ScaleWidth      =   1635
      TabIndex        =   80
      Top             =   4980
      Width           =   1665
      Begin MSComctlLib.ListView lvResult 
         Height          =   1155
         Index           =   21
         Left            =   0
         TabIndex        =   81
         TabStop         =   0   'False
         Top             =   270
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   2037
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   12648447
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Code"
            Object.Width           =   2355
         EndProperty
      End
      Begin VB.Label lbJobCount 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         Caption         =   "20 Jobs"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   21
         Left            =   30
         TabIndex        =   83
         Top             =   30
         Width           =   1395
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   21
         Left            =   1470
         TabIndex        =   82
         Top             =   30
         Width           =   105
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H80000008&
      Height          =   1455
      Index           =   22
      Left            =   1680
      ScaleHeight     =   1425
      ScaleWidth      =   1635
      TabIndex        =   76
      Top             =   4980
      Width           =   1665
      Begin MSComctlLib.ListView lvResult 
         Height          =   1155
         Index           =   22
         Left            =   0
         TabIndex        =   77
         TabStop         =   0   'False
         Top             =   270
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   2037
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   12648447
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Code"
            Object.Width           =   2355
         EndProperty
      End
      Begin VB.Label lbJobCount 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         Caption         =   "20 Jobs"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   22
         Left            =   30
         TabIndex        =   79
         Top             =   30
         Width           =   1395
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   22
         Left            =   1470
         TabIndex        =   78
         Top             =   30
         Width           =   105
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H80000008&
      Height          =   1455
      Index           =   23
      Left            =   3330
      ScaleHeight     =   1425
      ScaleWidth      =   1635
      TabIndex        =   72
      Top             =   4980
      Width           =   1665
      Begin MSComctlLib.ListView lvResult 
         Height          =   1155
         Index           =   23
         Left            =   0
         TabIndex        =   73
         TabStop         =   0   'False
         Top             =   270
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   2037
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   12648447
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Code"
            Object.Width           =   2355
         EndProperty
      End
      Begin VB.Label lbJobCount 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         Caption         =   "20 Jobs"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   23
         Left            =   30
         TabIndex        =   75
         Top             =   30
         Width           =   1395
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   23
         Left            =   1470
         TabIndex        =   74
         Top             =   30
         Width           =   105
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H80000008&
      Height          =   1455
      Index           =   24
      Left            =   4980
      ScaleHeight     =   1425
      ScaleWidth      =   1635
      TabIndex        =   68
      Top             =   4980
      Width           =   1665
      Begin MSComctlLib.ListView lvResult 
         Height          =   1155
         Index           =   24
         Left            =   0
         TabIndex        =   69
         TabStop         =   0   'False
         Top             =   270
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   2037
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   12648447
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Code"
            Object.Width           =   2355
         EndProperty
      End
      Begin VB.Label lbJobCount 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         Caption         =   "20 Jobs"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   24
         Left            =   30
         TabIndex        =   71
         Top             =   30
         Width           =   1395
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   24
         Left            =   1470
         TabIndex        =   70
         Top             =   30
         Width           =   105
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H80000008&
      Height          =   1455
      Index           =   25
      Left            =   6630
      ScaleHeight     =   1425
      ScaleWidth      =   1635
      TabIndex        =   64
      Top             =   4980
      Width           =   1665
      Begin MSComctlLib.ListView lvResult 
         Height          =   1155
         Index           =   25
         Left            =   0
         TabIndex        =   65
         TabStop         =   0   'False
         Top             =   270
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   2037
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   12648447
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Code"
            Object.Width           =   2355
         EndProperty
      End
      Begin VB.Label lbJobCount 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         Caption         =   "20 Jobs"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   25
         Left            =   30
         TabIndex        =   67
         Top             =   30
         Width           =   1395
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   25
         Left            =   1470
         TabIndex        =   66
         Top             =   30
         Width           =   105
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H80000008&
      Height          =   1455
      Index           =   26
      Left            =   8280
      ScaleHeight     =   1425
      ScaleWidth      =   1635
      TabIndex        =   60
      Top             =   4980
      Width           =   1665
      Begin MSComctlLib.ListView lvResult 
         Height          =   1155
         Index           =   26
         Left            =   0
         TabIndex        =   61
         TabStop         =   0   'False
         Top             =   270
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   2037
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   12648447
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Code"
            Object.Width           =   2355
         EndProperty
      End
      Begin VB.Label lbJobCount 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         Caption         =   "20 Jobs"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   26
         Left            =   30
         TabIndex        =   63
         Top             =   30
         Width           =   1395
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   26
         Left            =   1470
         TabIndex        =   62
         Top             =   30
         Width           =   105
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H80000008&
      Height          =   1455
      Index           =   27
      Left            =   9930
      ScaleHeight     =   1425
      ScaleWidth      =   1635
      TabIndex        =   56
      Top             =   4980
      Width           =   1665
      Begin MSComctlLib.ListView lvResult 
         Height          =   1155
         Index           =   27
         Left            =   0
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   270
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   2037
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   12648447
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Code"
            Object.Width           =   2355
         EndProperty
      End
      Begin VB.Label lbJobCount 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         Caption         =   "20 Jobs"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   27
         Left            =   30
         TabIndex        =   59
         Top             =   30
         Width           =   1395
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   27
         Left            =   1470
         TabIndex        =   58
         Top             =   30
         Width           =   105
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H80000008&
      Height          =   1455
      Index           =   28
      Left            =   30
      ScaleHeight     =   1425
      ScaleWidth      =   1635
      TabIndex        =   52
      Top             =   6420
      Width           =   1665
      Begin MSComctlLib.ListView lvResult 
         Height          =   1155
         Index           =   28
         Left            =   0
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   270
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   2037
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   12648447
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Code"
            Object.Width           =   2355
         EndProperty
      End
      Begin VB.Label lbJobCount 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         Caption         =   "20 Jobs"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   28
         Left            =   30
         TabIndex        =   55
         Top             =   30
         Width           =   1395
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   28
         Left            =   1470
         TabIndex        =   54
         Top             =   30
         Width           =   105
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H80000008&
      Height          =   1455
      Index           =   29
      Left            =   1680
      ScaleHeight     =   1425
      ScaleWidth      =   1635
      TabIndex        =   48
      Top             =   6420
      Width           =   1665
      Begin MSComctlLib.ListView lvResult 
         Height          =   1155
         Index           =   29
         Left            =   0
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   270
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   2037
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   12648447
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Code"
            Object.Width           =   2355
         EndProperty
      End
      Begin VB.Label lbJobCount 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         Caption         =   "20 Jobs"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   29
         Left            =   30
         TabIndex        =   51
         Top             =   30
         Width           =   1395
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   29
         Left            =   1470
         TabIndex        =   50
         Top             =   30
         Width           =   105
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H80000008&
      Height          =   1455
      Index           =   30
      Left            =   3330
      ScaleHeight     =   1425
      ScaleWidth      =   1635
      TabIndex        =   44
      Top             =   6420
      Width           =   1665
      Begin MSComctlLib.ListView lvResult 
         Height          =   1155
         Index           =   30
         Left            =   0
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   270
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   2037
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   12648447
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Code"
            Object.Width           =   2355
         EndProperty
      End
      Begin VB.Label lbJobCount 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         Caption         =   "20 Jobs"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   30
         Left            =   30
         TabIndex        =   47
         Top             =   30
         Width           =   1395
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   30
         Left            =   1470
         TabIndex        =   46
         Top             =   30
         Width           =   105
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H80000008&
      Height          =   1455
      Index           =   31
      Left            =   4980
      ScaleHeight     =   1425
      ScaleWidth      =   1635
      TabIndex        =   40
      Top             =   6420
      Width           =   1665
      Begin MSComctlLib.ListView lvResult 
         Height          =   1155
         Index           =   31
         Left            =   0
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   270
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   2037
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   12648447
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Code"
            Object.Width           =   2355
         EndProperty
      End
      Begin VB.Label lbJobCount 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         Caption         =   "20 Jobs"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   31
         Left            =   30
         TabIndex        =   43
         Top             =   30
         Width           =   1395
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   31
         Left            =   1470
         TabIndex        =   42
         Top             =   30
         Width           =   105
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H80000008&
      Height          =   1455
      Index           =   32
      Left            =   6630
      ScaleHeight     =   1425
      ScaleWidth      =   1635
      TabIndex        =   36
      Top             =   6420
      Width           =   1665
      Begin MSComctlLib.ListView lvResult 
         Height          =   1155
         Index           =   32
         Left            =   0
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   270
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   2037
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   12648447
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Code"
            Object.Width           =   2355
         EndProperty
      End
      Begin VB.Label lbJobCount 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         Caption         =   "20 Jobs"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   32
         Left            =   30
         TabIndex        =   39
         Top             =   30
         Width           =   1395
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   32
         Left            =   1470
         TabIndex        =   38
         Top             =   30
         Width           =   105
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H80000008&
      Height          =   1455
      Index           =   33
      Left            =   8280
      ScaleHeight     =   1425
      ScaleWidth      =   1635
      TabIndex        =   32
      Top             =   6420
      Width           =   1665
      Begin MSComctlLib.ListView lvResult 
         Height          =   1155
         Index           =   33
         Left            =   0
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   270
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   2037
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   12648447
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Code"
            Object.Width           =   2355
         EndProperty
      End
      Begin VB.Label lbJobCount 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         Caption         =   "20 Jobs"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   33
         Left            =   30
         TabIndex        =   35
         Top             =   30
         Width           =   1395
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   33
         Left            =   1470
         TabIndex        =   34
         Top             =   30
         Width           =   105
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H80000008&
      Height          =   1455
      Index           =   34
      Left            =   9930
      ScaleHeight     =   1425
      ScaleWidth      =   1635
      TabIndex        =   28
      Top             =   6420
      Width           =   1665
      Begin MSComctlLib.ListView lvResult 
         Height          =   1155
         Index           =   34
         Left            =   0
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   270
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   2037
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   12648447
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Code"
            Object.Width           =   2355
         EndProperty
      End
      Begin VB.Label lbJobCount 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         Caption         =   "20 Jobs"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   34
         Left            =   30
         TabIndex        =   31
         Top             =   30
         Width           =   1395
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   34
         Left            =   1470
         TabIndex        =   30
         Top             =   30
         Width           =   105
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H80000008&
      Height          =   1455
      Index           =   35
      Left            =   30
      ScaleHeight     =   1425
      ScaleWidth      =   1635
      TabIndex        =   24
      Top             =   7860
      Width           =   1665
      Begin MSComctlLib.ListView lvResult 
         Height          =   1155
         Index           =   35
         Left            =   0
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   270
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   2037
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   12648447
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Code"
            Object.Width           =   2355
         EndProperty
      End
      Begin VB.Label lbJobCount 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         Caption         =   "20 Jobs"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   35
         Left            =   30
         TabIndex        =   27
         Top             =   30
         Width           =   1395
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   35
         Left            =   1470
         TabIndex        =   26
         Top             =   30
         Width           =   105
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H80000008&
      Height          =   1455
      Index           =   36
      Left            =   1680
      ScaleHeight     =   1425
      ScaleWidth      =   1635
      TabIndex        =   20
      Top             =   7860
      Width           =   1665
      Begin MSComctlLib.ListView lvResult 
         Height          =   1155
         Index           =   36
         Left            =   0
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   270
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   2037
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   12648447
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Code"
            Object.Width           =   2355
         EndProperty
      End
      Begin VB.Label lbJobCount 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         Caption         =   "20 Jobs"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   36
         Left            =   30
         TabIndex        =   23
         Top             =   30
         Width           =   1395
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   36
         Left            =   1470
         TabIndex        =   22
         Top             =   30
         Width           =   105
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H80000008&
      Height          =   1455
      Index           =   37
      Left            =   3330
      ScaleHeight     =   1425
      ScaleWidth      =   1635
      TabIndex        =   16
      Top             =   7860
      Width           =   1665
      Begin MSComctlLib.ListView lvResult 
         Height          =   1155
         Index           =   37
         Left            =   0
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   270
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   2037
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   12648447
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Code"
            Object.Width           =   2355
         EndProperty
      End
      Begin VB.Label lbJobCount 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         Caption         =   "20 Jobs"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   37
         Left            =   30
         TabIndex        =   19
         Top             =   30
         Width           =   1395
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   37
         Left            =   1470
         TabIndex        =   18
         Top             =   30
         Width           =   105
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H80000008&
      Height          =   1455
      Index           =   38
      Left            =   4980
      ScaleHeight     =   1425
      ScaleWidth      =   1635
      TabIndex        =   12
      Top             =   7860
      Width           =   1665
      Begin MSComctlLib.ListView lvResult 
         Height          =   1155
         Index           =   38
         Left            =   0
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   270
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   2037
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   12648447
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Code"
            Object.Width           =   2355
         EndProperty
      End
      Begin VB.Label lbJobCount 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         Caption         =   "20 Jobs"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   38
         Left            =   30
         TabIndex        =   15
         Top             =   30
         Width           =   1395
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   38
         Left            =   1470
         TabIndex        =   14
         Top             =   30
         Width           =   105
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H80000008&
      Height          =   1455
      Index           =   39
      Left            =   6630
      ScaleHeight     =   1425
      ScaleWidth      =   1635
      TabIndex        =   8
      Top             =   7860
      Width           =   1665
      Begin MSComctlLib.ListView lvResult 
         Height          =   1155
         Index           =   39
         Left            =   0
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   270
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   2037
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   12648447
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Code"
            Object.Width           =   2355
         EndProperty
      End
      Begin VB.Label lbJobCount 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         Caption         =   "20 Jobs"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   39
         Left            =   30
         TabIndex        =   11
         Top             =   30
         Width           =   1395
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   39
         Left            =   1470
         TabIndex        =   10
         Top             =   30
         Width           =   105
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H80000008&
      Height          =   1455
      Index           =   40
      Left            =   8280
      ScaleHeight     =   1425
      ScaleWidth      =   1635
      TabIndex        =   4
      Top             =   7860
      Width           =   1665
      Begin MSComctlLib.ListView lvResult 
         Height          =   1155
         Index           =   40
         Left            =   0
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   270
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   2037
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   12648447
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Code"
            Object.Width           =   2355
         EndProperty
      End
      Begin VB.Label lbJobCount 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         Caption         =   "20 Jobs"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   40
         Left            =   30
         TabIndex        =   7
         Top             =   30
         Width           =   1395
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   40
         Left            =   1470
         TabIndex        =   6
         Top             =   30
         Width           =   105
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H80000008&
      Height          =   1455
      Index           =   41
      Left            =   9930
      ScaleHeight     =   1425
      ScaleWidth      =   1635
      TabIndex        =   0
      Top             =   7860
      Width           =   1665
      Begin MSComctlLib.ListView lvResult 
         Height          =   1245
         Index           =   41
         Left            =   0
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   270
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   2196
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   12648447
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Code"
            Object.Width           =   2355
         EndProperty
      End
      Begin VB.Label lbJobCount 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         Caption         =   "20 Jobs"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   41
         Left            =   30
         TabIndex        =   3
         Top             =   30
         Width           =   1395
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   41
         Left            =   1470
         TabIndex        =   2
         Top             =   30
         Width           =   105
      End
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "อา."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   30
      TabIndex        =   188
      Top             =   420
      Width           =   1665
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "จ."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   1680
      TabIndex        =   187
      Top             =   420
      Width           =   1665
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "อ."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   3330
      TabIndex        =   186
      Top             =   420
      Width           =   1665
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "พ."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   4980
      TabIndex        =   185
      Top             =   420
      Width           =   1665
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "พฤ."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   6630
      TabIndex        =   184
      Top             =   420
      Width           =   1665
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ศ."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   5
      Left            =   8280
      TabIndex        =   183
      Top             =   420
      Width           =   1665
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ส."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   6
      Left            =   9930
      TabIndex        =   182
      Top             =   420
      Width           =   1665
   End
   Begin VB.Label aaaa 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "ลูกค้า"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   -45
      TabIndex        =   181
      Top             =   -45
      Visible         =   0   'False
      Width           =   375
   End
End
Attribute VB_Name = "FrmCustCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim monthName_(13) As String
Dim yearList_(10) As Integer
Dim colJobData_ As Collection
Dim colCustData_ As Collection
Dim firstDayIndex_ As Integer
Dim curIndex_ As Integer
Dim srcJobKey_ As String
Dim curSale_ As BWGCompanyStaff
Dim curCustomer_ As BWGCustomer
Private Sub cboCustType_Click()
ShowCalendar
End Sub
Private Sub cboMonth_Click()
    ShowCalendar
End Sub
Private Sub CboType_Click()
    ShowCalendar
End Sub
Private Sub cboYear_Click()
    ShowCalendar
End Sub
Private Sub cmdClose_Click()
    Unload Me
End Sub
Private Sub Command1_Click()
    frmCustVisit.Show 1
    ShowCustThisMonth
End Sub
Private Sub Form_Load()
Dim x%
Dim curYear As Integer
Dim UserGrp As Collection
Dim i As Integer
'    If LCase(CurrentUser.StaffCode) = "pr" Or LCase(CurrentUser.StaffCode) = "env" Or LCase(CurrentUser.StaffCode) = "akpenvi" _
'    Or LCase(CurrentUser.StaffCode) = "admin" Or LCase(CurrentUser.StaffCode) = "env02" Then
'    Command1.Visible = True
'    Else
'    Command1.Visible = False
'    End If
    curYear = dtToday.YearValue
    cboMonth.AddItem "มกราคม"
    cboMonth.AddItem "กุมภาพันธ์"
    cboMonth.AddItem "มีนาคม"
    cboMonth.AddItem "เมษายน"
    cboMonth.AddItem "พฤษภาคม"
    cboMonth.AddItem "มิถุนายน"
    cboMonth.AddItem "กรกฎาคม"
    cboMonth.AddItem "สิงหาคม"
    cboMonth.AddItem "กันยายน"
    cboMonth.AddItem "ตุลาคม"
    cboMonth.AddItem "พฤศจิกายน"
    cboMonth.AddItem "ธันวาคม"
    cboMonth.ListIndex = dtToday.MonthValue - 1
    For x = 0 To lbJobCount.UBound
        lbJobCount(x).BackStyle = 0
        Label1(x).BackStyle = 0
    Next
    For x = -5 To 3
        cboYear.AddItem curYear + x
    Next
    cboYear.ListIndex = 5
       With cboCustType
        .AddItem "=ฐาน="
        .AddItem "(C) ,(S)"
        .AddItem "(M),(MS)"
        .AddItem "(B),(E)"
        .ListIndex = 0
    End With
    CboType.ListIndex = 0
    lbToday.Caption = dtToday.DayValue & " " & dtToday.MonthName & " " & dtToday.YearValue
    ShowCalendar
End Sub
Private Sub ShowCalendar()
Dim x%
Dim firstDate As Date, lastDate
Dim firstInd%, lastInd%
    If cboYear.ListIndex > -1 And cboMonth.ListIndex > -1 Then
        firstDate = DateSerial(cboYear.Text - 543, cboMonth.ListIndex + 1, 1)
        lastDate = DateSerial(cboYear.Text - 543, cboMonth.ListIndex + 2, 0)
        firstInd = Weekday(firstDate, vbSunday)
        lastInd = Day(lastDate)
        For x = 0 To Picture1.UBound
            If x < firstInd - 1 Or x - firstInd + 2 > lastInd Then
                Picture1(x).BackColor = vbWhite
                lbJobCount(x).Visible = False
                Label1(x).Visible = False
                lvResult(x).Visible = False
                lvResult(x).ListItems.Clear
            Else
                lvResult(x).ListItems.Clear
                If x - firstInd + 2 = dtToday.DayValue And cboMonth.ListIndex + 1 = dtToday.MonthValue And CInt(cboYear.Text) = dtToday.YearValue Then
                    Picture1(x).BackColor = &HC0C0FF
                Else
                    Picture1(x).BackColor = &HC0FFC0
                End If
                lbJobCount(x).Visible = True
                Label1(x).Visible = True
                lvResult(x).Visible = True
                Label1(x).Caption = x - firstInd + 2
                If Trim(Label1(x).Caption) = "1" Then firstDayIndex_ = x
                If lvResult(x).ListItems.Count > 0 Then
                    lbJobCount(x).Caption = lvResult(x).ListItems.Count & " Jobs"
                Else
                    lbJobCount(x).Caption = ""
                End If
            End If
        Next
        ShowCustThisMonth
    End If
End Sub
Private Sub lvResult_DblClick(Index As Integer)
On Error GoTo ErrD
    Dim tmpS As New BWGSearchManager, tmpCol As Collection
'    If LCase(CurrentUser.StaffCode) = "pr" Or LCase(CurrentUser.StaffCode) = "env" Or LCase(CurrentUser.StaffCode) = "akpenvi" _
'    Or LCase(CurrentUser.StaffCode) = "admin" Or LCase(CurrentUser.StaffCode) = "env02" Then
Dim tmpStr As String
    tmpStr = CurrentUser.VerifyAccessRole("frmCustVisit", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    If lvResult(Index).SelectedItem Is Nothing Then Exit Sub
    Load frmCustVisit
    With frmCustVisit
        Set tmpCol = tmpS.CustVisit("VisitID=" & lvResult(Index).SelectedItem.Key)
        If tmpCol.Count > 0 Then
            .gVal = "0"
            .SetDataCustVisit tmpCol(1)
            Set tmpCol = Nothing
            Set tmpS = Nothing
            .Show 1
            ShowCustThisMonth
        End If
    End With
    Set tmpCol = Nothing
    Set tmpS = Nothing
    Exit Sub
'    Else
'        Exit Sub
'    End If
ErrD:
    MsgBox err.Description
    Exit Sub
End Sub

Private Sub ShowCustThisMonth()
Dim tmpS As New BWGSearchManager, tmpStr As String
Dim x%, tmpJob As BWGJobData, tmpDay As Integer
Dim tmpCustVisit As BWGCustVisitor
Dim UserInCompany As New BWGCompanyStaff
Dim UserCompany As Collection
Dim i As Integer, CompanyID As String
    For x = 1 To lvResult.UBound
        lvResult(x).ListItems.Clear
    Next
    If cboYear.ListIndex > -1 And cboMonth.ListIndex > -1 Then
        tmpStr = "visitdate>='" & cboYear.Text & "/" & Format((cboMonth.ListIndex + 1), "00") & "/01' And visitdate<='" & cboYear.Text & "/" & Format((cboMonth.ListIndex + 1), "00") & "/31' And (isCancel<>'Y' or isCancel is null) "
        If Not curCustomer_ Is Nothing Then
            If Trim(curCustomer_.ID) <> "" Then
                tmpStr = tmpStr & " And CustomerID='" & curCustomer_.ID & "' "
            End If
        End If
        
'        Set UserCompany = CurrentUser.getUserCompany
'        CompanyID = ""
'        For i = 1 To UserCompany.Count
'            If CompanyID = "" Then
'                CompanyID = "'" & UserCompany(i).ID & "'"
'            Else
'                CompanyID = CompanyID & ", '" & UserCompany(i).ID & "'"
'            End If
'        Next
'        tmpStr = tmpStr & " AND CompanyID IN (" & CompanyID & ")"
        If cboCustType.ListIndex = 1 Then
            tmpStr = tmpStr & " AND CustTypeFlag in ('C','S') "
        ElseIf cboCustType.ListIndex = 2 Then
            tmpStr = tmpStr & " AND CustTypeFlag in ('M','MS') "
        ElseIf cboCustType.ListIndex = 3 Then
            tmpStr = tmpStr & " AND CustTypeFlag in ('B','E') "
        End If
        
        If CboType.ListIndex > 0 Then
            tmpStr = tmpStr & " AND visittype='" & CboType.Text & "' "
        End If
        Set colCustData_ = tmpS.CustVisit(tmpStr)
        For x = 1 To colCustData_.Count
            Set tmpCustVisit = colCustData_(x)
            tmpDay = CInt(Right(tmpCustVisit.VisitDate, 2))
            If isExist(lvResult(firstDayIndex_ + tmpDay - 1).ListItems, "'" & tmpCustVisit.ID & "'") Then
                lvResult(firstDayIndex_ + tmpDay - 1).ListItems("'" & tmpCustVisit.ID & "'").Text = tmpCustVisit.NameDisplay
            Else
                lvResult(firstDayIndex_ + tmpDay - 1).ListItems.Add , "'" & tmpCustVisit.ID & "'", tmpCustVisit.NameDisplay
            End If
            If tmpCustVisit.isComplete = True Then
            lvResult(firstDayIndex_ + tmpDay - 1).ListItems("'" & tmpCustVisit.ID & "'").ForeColor = vbBlue
            End If
            If tmpCustVisit.isUse = True Then
            lvResult(firstDayIndex_ + tmpDay - 1).ListItems("'" & tmpCustVisit.ID & "'").ForeColor = vbGreen
            End If
            If tmpCustVisit.IsInform = True Then
            lvResult(firstDayIndex_ + tmpDay - 1).ListItems("'" & tmpCustVisit.ID & "'").ForeColor = vbRed
            End If
            If tmpCustVisit.IsNoCome = True Then
            lvResult(firstDayIndex_ + tmpDay - 1).ListItems("'" & tmpCustVisit.ID & "'").ForeColor = &H800080
            End If
            Set tmpCustVisit = Nothing
        Next
    End If
    Set tmpS = Nothing
End Sub

Private Sub lvResult_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        curIndex_ = Index
    End If
End Sub

Private Sub TxtCustomer_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{Tab}"
End Sub

Private Sub txtCustomer_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub txtCustomer_Validate(Cancel As Boolean)
    If curCustomer_ Is Nothing Then Set curCustomer_ = New BWGCustomer
    If Trim(txtCustomer.Text) = "" And Trim(curCustomer_.CustomerName) <> Trim(txtCustomer.Text) Then
        Set curCustomer_ = Nothing
        Set curCustomer_ = New BWGCustomer
        ShowCustThisMonth
    Else
        If Trim(curCustomer_.CustomerName) <> Trim(txtCustomer.Text) Then
            Set curCustomer_ = Nothing
            Set curCustomer_ = New BWGCustomer
            Load frmCustomerSearch
            With frmCustomerSearch
                .CStatus = "N"
                .getCriteria (txtCustomer.Text)
                .SetCallerForm Me
                .Show 1
            End With
            If curCustomer_.ID = "" Then txtCustomer.Text = ""
        End If
    End If
End Sub

Private Sub txtSale_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{Tab}"
End Sub

Private Sub txtSale_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub txtSale_Validate(Cancel As Boolean)
    If curSale_ Is Nothing Then Set curSale_ = New BWGCompanyStaff
    If Trim(txtSale.Text) = "" And Trim(curSale_.StaffFName & " " & curSale_.StaffLName) <> Trim(txtSale.Text) Then
        Set curSale_ = Nothing
        Set curSale_ = New BWGCompanyStaff
        ShowCustThisMonth
    Else
        If Trim(curSale_.StaffFName & " " & curSale_.StaffLName) <> Trim(txtSale.Text) Then
            Set curSale_ = Nothing
            Set curSale_ = New BWGCompanyStaff
            Load frmStaffSearch
            With frmStaffSearch
                .getCriteria (txtSale.Text)
                .SetCallerForm Me
                .Show 1
            End With
            If curSale_.ID = "" Then txtSale.Text = ""
        End If
    End If
End Sub

Public Sub SetDataStaff(selStaff As BWGCompanyStaff)
    Set curSale_ = selStaff
    txtSale.Text = curSale_.StaffFName & " " & curSale_.StaffLName
    ShowCustThisMonth
End Sub

Public Sub SetDataCustomer(selCust As BWGCustomer)
    Set curCustomer_ = selCust
    txtCustomer.Text = curCustomer_.CustomerName
    ShowCustThisMonth
End Sub
