VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "VSFLEX7L.OCX"
Begin VB.Form frmCal 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Calendar Demo"
   ClientHeight    =   6330
   ClientLeft      =   720
   ClientTop       =   1500
   ClientWidth     =   9300
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   9300
   ShowInTaskbar   =   0   'False
   Begin VSFlex7LCtl.VSFlexGrid cal 
      Height          =   3855
      Left            =   1740
      TabIndex        =   0
      Top             =   2220
      Width           =   6975
      _cx             =   5080
      _cy             =   5080
      _ConvInfo       =   -1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483644
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   8421504
      BackColorAlternate=   -2147483644
      GridColor       =   12632256
      GridColorFixed  =   0
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   0
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   2
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   7
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   0
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
      ShowComboButton =   -1  'True
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
   Begin VB.Image Image3 
      Height          =   1110
      Left            =   480
      Picture         =   "frmCal.frx":0000
      Top             =   60
      Width           =   2280
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   1
      Left            =   360
      Shape           =   3  'Circle
      Top             =   1080
      Width           =   135
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   420
      X2              =   5160
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   465
      Left            =   1740
      TabIndex        =   2
      Top             =   1500
      Width           =   3420
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   330
      Index           =   0
      Left            =   6855
      Top             =   1560
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   330
      Index           =   1
      Left            =   7350
      Picture         =   "frmCal.frx":1D2E
      Top             =   1545
      Width           =   360
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   330
      Index           =   2
      Left            =   7755
      Picture         =   "frmCal.frx":1EB8
      Top             =   1545
      Width           =   360
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   330
      Index           =   3
      Left            =   8160
      Picture         =   "frmCal.frx":2042
      Top             =   1545
      Width           =   360
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Drag and Drop -->"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   5340
      TabIndex        =   1
      Top             =   1590
      Width           =   2040
   End
End
Attribute VB_Name = "frmCal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TempDate#

Private Sub cal_DblClick()
   
   ' avoid first row or blank days
   If cal.MouseRow = 0 Or cal = "" Then Exit Sub

   'txtNotes.Move (cal.CellLeft + cal.Left) + 100, (cal.CellTop + cal.Top) + 100
   'txtNotes.Visible = True
   'txtNotes.SetFocus
   
End Sub


Private Sub cal_DragDrop(Source As Control, X As Single, Y As Single)

   ' avoid first row and blank days
   If cal.MouseRow = 0 Or cal = "" Then Exit Sub
   
   ' drop picture to cell
   cal.Cell(flexcpPicture, cal.MouseRow, cal.MouseCol) = Source.Picture
   cal.Cell(flexcpPictureAlignment, cal.MouseRow, cal.MouseCol) = 1
   cal.Cell(flexcpBackColor, cal.MouseRow, cal.MouseCol) = &H80FFFF
End Sub


Private Sub Form_Load()
    DoCalendar Now
End Sub


Private Sub DoCalendar(AnyDate)

   Dim FirstOfMonth#, FirstOfNextMonth#, DateOffset#, NumDays#
   Dim i%, j%, k%, FiveWeeks%, NumOf%
   
   ' Basic Date Math to calculate view --------------------------
   TempDate = AnyDate                     ' save date for future reference
   FirstOfMonth = DateSerial(Year(AnyDate), Month(AnyDate), 1)
   FirstOfNextMonth = DateSerial(Year(AnyDate), Month(AnyDate) + 1, 1)
   NumDays = FirstOfNextMonth - FirstOfMonth
   DateOffset = Weekday(FirstOfMonth) - 1
   ' ------------------------------------------------------------

   cal.Clear            'Clear off the grid
   cal.Row = 0
   cal.FormatString = "^SUN|^MON|^TUE|^WED|^THU|^FRI|^SAT"
   cal.Row = 1
   
   ' place datenumbers into the cells
   k = DateOffset - 1
   For i = 1 To NumDays

       k = k + 1
       If k = 7 Then
           k = 0
           cal.Row = cal.Row + 1
       End If
       
       ' check to see if the month uses 5 or 6 weeks to display
       FiveWeeks = True
       If cal.Row = 6 Then FiveWeeks = False
       
       cal.col = k
       cal = Str$(i)
       cal.CellAlignment = 6
       cal.CellPictureAlignment = 4
   Next
   
   Label1 = Format$(FirstOfMonth, "MMMM YYYY")
   
   ' resize the grid cells to fit form
   For i = 0 To 6
       cal.ColWidth(i) = (cal.Width \ 7)
       NumOf% = 5
       If FiveWeeks = True Then NumOf = 6
       For j = 1 To 6
           cal.RowHeight(j) = (cal.Height \ NumOf) - ((cal.RowHeight(0) \ NumOf))
       Next
   Next

  ' cal.TopRow = 0

End Sub


Private Sub Image1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    'txtNotes.Visible = False
    Image1(Index).Drag 1
End Sub


