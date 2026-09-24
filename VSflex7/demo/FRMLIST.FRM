VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "VSFLEX7L.OCX"
Begin VB.Form frmList 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lists Demo"
   ClientHeight    =   6195
   ClientLeft      =   885
   ClientTop       =   1395
   ClientWidth     =   10020
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6195
   ScaleWidth      =   10020
   Begin VSFlex7LCtl.VSFlexGrid vsTodo 
      Height          =   2415
      Left            =   4920
      TabIndex        =   1
      Top             =   1440
      Width           =   3855
      _cx             =   5080
      _cy             =   5080
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
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
      BackColorBkg    =   8421504
      BackColorAlternate=   -2147483643
      GridColor       =   12632256
      GridColorFixed  =   0
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   0
      FocusRect       =   0
      HighLight       =   2
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   0
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   2
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   2
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   1
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
      WordWrap        =   -1  'True
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
   Begin VSFlex7LCtl.VSFlexGrid vsSched 
      Height          =   4215
      Left            =   600
      TabIndex        =   0
      Top             =   1440
      Width           =   3375
      _cx             =   5080
      _cy             =   5080
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   12648447
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   8421504
      BackColorAlternate=   12648447
      GridColor       =   12632256
      GridColorFixed  =   0
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   0
      FocusRect       =   1
      HighLight       =   0
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   3
      FixedRows       =   0
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
   Begin VB.Line Line1 
      Index           =   0
      X1              =   4620
      X2              =   8280
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   0
      Left            =   4500
      Shape           =   3  'Circle
      Top             =   1080
      Width           =   135
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   360
      X2              =   3480
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   1
      Left            =   240
      Shape           =   3  'Circle
      Top             =   1080
      Width           =   135
   End
   Begin VB.Image Image2 
      Height          =   1020
      Index           =   1
      Left            =   4740
      Picture         =   "frmList.frx":0000
      Top             =   300
      Width           =   1425
   End
   Begin VB.Image image1 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   1
      Left            =   5700
      Picture         =   "frmList.frx":105B
      Top             =   4980
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image image1 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   0
      Left            =   6480
      Picture         =   "frmList.frx":1365
      Top             =   4920
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image image1 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   2
      Left            =   7500
      Picture         =   "frmList.frx":3B07
      Top             =   4920
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   1020
      Index           =   0
      Left            =   360
      Picture         =   "frmList.frx":3E11
      Top             =   300
      Width           =   3060
   End
End
Attribute VB_Name = "frmList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub DoAppointments()

   ' enter data for appointments
   Dim i%
   
   
   For i = 1 To 3
      vsSched.Cell(flexcpText, i, 1) = "Breakfast at VideoSoft"
      vsSched.Cell(flexcpBackColor, i, 1) = &H80FFFF    'yellow
   Next i

   
   For i = 9 To 14
      vsSched.Cell(flexcpPictureAlignment, i, 1) = 3
      vsSched.Cell(flexcpPicture, i, 1) = Image1(2).Picture
      vsSched.Cell(flexcpText, i, 1) = "Dentist Appointment"
      vsSched.Cell(flexcpBackColor, i, 1) = &HFFFF80  ' blue
   Next i

   
   For i = 2 To 4
      vsSched.Cell(flexcpText, i, 2) = "VSFlexGrid Press Release"
      vsSched.Cell(flexcpBackColor, i, 2) = &H80FF80  ' green
   Next i

End Sub

Private Sub DoInitialSettings()

    vsSched.WordWrap = True

   ' schedule FlexArray
   vsSched.ColWidth(1) = 1000
   vsSched.ColWidth(2) = 1000
   vsSched.MergeCol(1) = True
   vsSched.MergeCol(2) = True
   vsSched.MergeCells = flexMergeRestrictColumns

   ' todo FlexArray
   vsTodo.ColWidth(0) = 500
   vsTodo.ColWidth(1) = vsTodo.Width - vsTodo.ColWidth(0)

End Sub

Private Sub DoTimeTitles()
   
   vsSched.Rows = 0

   vsSched.AddItem " 07:00 AM"
   vsSched.AddItem " 07:30 AM"
   vsSched.AddItem " 08:00 AM"
   vsSched.AddItem " 08:30 AM"
   vsSched.AddItem " 09:00 AM"
   vsSched.AddItem " 09:30 AM"
   vsSched.AddItem " 10:00 AM"
   vsSched.AddItem " 10:30 AM"
   vsSched.AddItem " 11:00 AM"
   vsSched.AddItem " 11:30 AM"
   vsSched.AddItem " 12:00 PM"
   vsSched.AddItem " 12:30 PM"
   vsSched.AddItem " 01:00 PM"
   vsSched.AddItem " 01:30 PM"
   vsSched.AddItem " 02:00 PM"
   vsSched.AddItem " 02:30 PM"
   vsSched.AddItem " 03:00 PM"
   vsSched.AddItem " 03:30 PM"
   vsSched.AddItem " 04:00 PM"
   vsSched.AddItem " 04:30 PM"
   vsSched.AddItem " 05:00 PM"
   vsSched.AddItem " 05:30 PM"
   vsSched.AddItem " 06:00 PM"
   vsSched.AddItem " 06:30 PM"
   vsSched.AddItem " 07:00 PM"
   vsSched.AddItem " 07:30 PM"
   vsSched.AddItem " 08:00 PM"
   vsSched.AddItem " 08:30 PM"

End Sub

Private Sub Form_Load()
   
    Dim i%
    
    DoInitialSettings
    DoTimeTitles
    DoAppointments
    
    
    vsTodo.Cell(flexcpPicture, 0, 0) = Image1(0)
    vsTodo.Cell(flexcpText, 0, 1) = "Update to Windows 98"
    
    vsTodo.Cell(flexcpPicture, 1, 0) = Image1(1)
    vsTodo.Cell(flexcpPictureAlignment, 1, 0) = 5
    vsTodo.Cell(flexcpText, 1, 1) = "Remember: Call your wife"
    vsTodo.Cell(flexcpForeColor, 1, 1) = RGB(255, 0, 0)
    vsTodo.Cell(flexcpFontBold, 1, 1) = True
    
    
    vsTodo.Cell(flexcpFontName, 2, 1) = "Arial"
    vsTodo.Cell(flexcpFontSize, 2, 1) = "12"
    vsTodo.Cell(flexcpFontItalic, 2, 1) = True
    vsTodo.Cell(flexcpFontBold, 2, 1) = True
    vsTodo.Cell(flexcpForeColor, 2, 1) = RGB(0, 0, 128)
    
    
    vsTodo.Cell(flexcpText, 2, 1) = "Check tax exemptions"
    vsTodo.Cell(flexcpText, 3, 1) = "Open bank account"
    
    vsTodo.AutoSize 0, 1, True, 200

End Sub


