VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "VSFLEX7L.OCX"
Begin VB.Form frmHTML 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Save to HTML Demo"
   ClientHeight    =   7155
   ClientLeft      =   1755
   ClientTop       =   1215
   ClientWidth     =   7560
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7155
   ScaleWidth      =   7560
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save To HTML"
      Height          =   495
      Left            =   5520
      TabIndex        =   1
      Top             =   900
      Width           =   1635
   End
   Begin VSFlex7LCtl.VSFlexGrid flex 
      Height          =   5175
      Left            =   300
      TabIndex        =   0
      Top             =   1620
      Width           =   6975
      _cx             =   12303
      _cy             =   9128
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   12640511
      ForeColor       =   -2147483640
      BackColorFixed  =   0
      ForeColorFixed  =   8421504
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   8421504
      BackColorAlternate=   12640511
      GridColor       =   12632256
      GridColorFixed  =   0
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   0
      HighLight       =   0
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   3
      GridLines       =   1
      GridLinesFixed  =   11
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
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
      Ellipsis        =   1
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   1
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
   Begin VB.Shape Shape1 
      BackColor       =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   7
      Left            =   240
      Shape           =   3  'Circle
      Top             =   960
      Width           =   135
   End
   Begin VB.Line Line1 
      Index           =   3
      X1              =   3840
      X2              =   360
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Image Image1 
      Height          =   1035
      Left            =   240
      Picture         =   "frmHTML.frx":0000
      Top             =   180
      Width           =   3390
   End
End
Attribute VB_Name = "frmHTML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Sub cmdSave_Click()

    MousePointer = 11

    On Error Resume Next

    ' save HTML file
    SaveFlexGridToHTML flex, App.Path & "\test.html"

    ' display HTML file in browser
    ShellExecute hwnd, "open", App.Path & "\test.html", 0, 0, 0
    

    MousePointer = 0

End Sub

Private Sub Form_Load()
    Dim t$
    
    t = Chr$(9)
    
    With flex
    
            '--------------------------------------------------------------------
            '
            '--------------------------------------------------------------------
            .AddItem "OLEDB Support" & t & "Read Only" & t & "Read and Write"
            .AddItem "Printing" & t & "Through VSVIEW" & t & "Yes"
            .AddItem "MergeCells" & t & "Yes" & t & "Enhanced"
            .AddItem "In-Cell Editing" & t & "" & t & "Yes"
            .AddItem "MaskEdit" & t & "" & t & "Yes"
            .AddItem "Subtotals and Outlining" & t & "" & t & "Yes"
            .AddItem "AutoSizeColumns" & t & "" & t & "Yes"
            .AddItem "Load/Save Grid" & t & "" & t & "Yes"
            .AddItem "Cell Borders" & t & "" & t & "Yes"
            .AddItem "OLE Drag And Drop" & t & "" & t & "Yes"
            .AddItem "ScrollTips" & t & "" & t & "Yes"
            .AddItem "OwnerDraw" & t & "" & t & "Yes"
            .AddItem "ExplorerBar View" & t & "" & t & "Yes"
            .AddItem "Checkbox" & t & "" & t & "Yes"
            .AddItem "ComboLists" & t & "" & t & "Yes"
            .AddItem "Multi-Column ComboLists" & t & "" & t & "Yes"
            .AddItem "Pop-up Buttons" & t & "" & t & "Yes"
            .AddItem "ListBox-Style Selection" & t & "" & t & "Yes"
            .AddItem "Formatted, International Sorting" & t & "" & t & "Yes"
            .AddItem "Cell Flooding" & t & "" & t & "Yes"
            .AddItem "Built-in CheckBoxes" & t & "" & t & "Yes"
            .AddItem "BindToArray" & t & "" & t & "Yes"
            .AddItem "AutoSearch" & t & "" & t & "Yes"
            .AddItem "FindRow" & t & "" & t & "Yes"
            .AddItem "Improved Date Formatting" & t & "" & t & "Yes"
            .AddItem "Enumerate Selected Rows" & t & "" & t & "Yes"
            .AddItem "Automatic Auditing" & t & "" & t & "Yes"
            .AddItem "Frozen Rows & Columns" & t & "" & t & "Yes"
            .AddItem "WallPaper" & t & "" & t & "Yes"
            .AddItem "Custom Data Sources" & t & "" & t & "Yes"
            .AddItem "Translated ImageLists" & t & "" & t & "Yes"
            .AddItem "Refer to Columns by Key" & t & "" & t & "Yes"
            .AddItem "Range Aggregates" & t & "" & t & "Yes"
            
    
            '--------------------------------------------------------------------
            ' title
            '--------------------------------------------------------------------
            .Cell(flexcpText, 0, 1) = "MSFlexGrid 6.0"
            .Cell(flexcpText, 0, 2) = "VSFlexGrid Pro 7.0"
            .Cell(flexcpForeColor, 0, 0, 0, .Cols - 1) = vbWhite
            .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = vbWhite
            
            
            .Select 0, 0, .Rows - 1, 0
            .CellBackColor = vbWhite
            
            
        
            .AutoSize 0, .Cols - 1


    End With

End Sub


