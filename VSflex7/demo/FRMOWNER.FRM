VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "VSFLEX7L.OCX"
Begin VB.Form frmOwner 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "This only looks good with True color"
   ClientHeight    =   4035
   ClientLeft      =   1800
   ClientTop       =   1545
   ClientWidth     =   5925
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4035
   ScaleWidth      =   5925
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   3600
      Width           =   2175
   End
   Begin VSFlex7LCtl.VSFlexGrid fa 
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5655
      _ExtentX        =   5080
      _ExtentY        =   5080
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
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   10
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   "=From|To|Gradient        "
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
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
      OwnerDraw       =   6
      Editable        =   -1  'True
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   4
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
   End
   Begin VB.Label Label1 
      Caption         =   "OwnerDraw setting:"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   3600
      Width           =   1575
   End
End
Attribute VB_Name = "frmOwner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Windows API stuff
Private Type RECT
    left As Long
    top As Long
    right As Long
    bottom As Long
End Type

Private Const ETO_CLIPPED = 4
Private Const ETO_GRAYED = 1
Private Const ETO_OPAQUE = 2

Private Declare Sub InflateRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long)
Private Declare Function ExtTextOut Lib "gdi32" Alias "ExtTextOutA" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal wOptions As Long, lpRect As RECT, ByVal lpString As String, ByVal nCount As Long, lpDx As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long


Private Sub GetRGB(clrname$, r%, g%, b%)
    Select Case clrname
        Case "Black"
        Case "White"
            r = 255
            g = 255
            b = 255
        Case "Red"
            r = 255
        Case "Green"
            g = 255
        Case "Blue"
            b = 255
        Case "Yellow"
            r = 255
            g = 255
        Case "Purple"
            r = 200
            b = 200
        Case "Aqua"
            b = 200
            g = 200
    End Select
End Sub

Private Sub Combo1_Click()
    fa.OwnerDraw = Combo1.ListIndex
    fa.Refresh
End Sub


Private Sub fa_AfterEdit(ByVal Row As Long, ByVal col As Long)
    Dim clr1$, clr2$
    
    clr1 = fa.TextMatrix(Row, 0)
    clr2 = fa.TextMatrix(Row, 1)
    
    If clr1 <> "" And clr2 <> "" Then
        fa.TextMatrix(Row, 2) = left(clr1, Len(clr1) / 2 + 1) & right(clr2, Len(clr2) / 2 + 1)
    Else
        fa.TextMatrix(Row, 2) = ""
    End If
    
End Sub

Private Sub fa_BeforeEdit(ByVal Row As Long, ByVal col As Long, Cancel As Boolean)
    fa.ComboList = "Black|White|Red|Green|Blue|Yellow|Purple|Aqua"
    If fa.col > 1 Then Cancel = True
End Sub

Private Sub fa_Click()
    fa.EditCell
End Sub

Private Sub fa_DrawCell(ByVal hDC As Long, ByVal Row As Long, ByVal col As Long, ByVal left As Long, ByVal top As Long, ByVal right As Long, ByVal bottom As Long, done As Boolean)
    Dim clr&, clrg&, l&
    Dim rc As RECT
    Dim rc1 As RECT
    Dim clr1$, clr2$
    Dim r1%, g1%, b1%
    Dim r2%, g2%, b2%
    Dim rg%, gg%, bg%
    
    With fa
    
        ' make sure we need to draw this guy
        clr1 = .TextMatrix(Row, 0)
        clr2 = .TextMatrix(Row, 1)
        If Row > 0 And col = 2 And clr1 <> "" And clr2 <> "" Then
    
            ' get rectangle
            Select Case .OwnerDraw
                Case flexODComplete
                    rc.left = left
                    rc.top = top
                    rc.right = right
                    rc.bottom = bottom
                Case flexODOver
                    rc.left = left + (right - left) / 3
                    rc.top = top + 2
                    rc.right = right - 5
                    rc.bottom = bottom - 3
                Case flexODContent
                    rc.left = left + 5
                    rc.top = top + 2
                    rc.right = right - 5
                    rc.bottom = bottom - 3
                Case Else
                    Exit Sub
            End Select
            
            ' draw frame
            clr = SetBkColor(hDC, 0)
            ExtTextOut hDC, rc.left, rc.top, ETO_OPAQUE, rc, " ", 1, l
            InflateRect rc, -1, -1
            
            ' get colors
            GetRGB clr1, r1, g1, b1
            GetRGB clr2, r2, g2, b2
            
            ' show color
            rc1 = rc
            For l = rc.left To rc.right Step 3
                rg = r1 + (l - rc.left) * (r2 - r1) / (rc.right - rc.left)
                gg = g1 + (l - rc.left) * (g2 - g1) / (rc.right - rc.left)
                bg = b1 + (l - rc.left) * (b2 - b1) / (rc.right - rc.left)
                clrg = RGB(rg, gg, bg)
                SetBkColor hDC, clrg
                rc1.left = l
                ExtTextOut hDC, rc1.left, rc1.top, ETO_OPAQUE, rc1, " ", 1, l
            Next
            
            ' done, restore hDC and quit
            SetBkColor hDC, clr
            done = True
        
        End If
    
    End With
    
End Sub

Private Sub Form_Load()

    Combo1.AddItem "flexODNone"
    Combo1.AddItem "flexODOver"
    Combo1.AddItem "flexODContent"
    Combo1.AddItem "flexODComplete"
    Combo1.ListIndex = 1
    
End Sub


