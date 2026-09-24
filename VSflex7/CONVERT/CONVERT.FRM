VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "VSFlexGrid 6 -> 7 Conversion Utility"
   ClientHeight    =   4830
   ClientLeft      =   2760
   ClientTop       =   2010
   ClientWidth     =   7335
   Icon            =   "Convert.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   7335
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox lstCtrls 
      Height          =   1620
      Left            =   5535
      TabIndex        =   10
      Top             =   1740
      Visible         =   0   'False
      Width           =   1605
   End
   Begin VB.Frame Frame1 
      Height          =   2640
      Left            =   195
      TabIndex        =   4
      Top             =   1500
      Width           =   4395
      Begin VB.OptionButton Option1 
         Caption         =   "No data binding"
         Height          =   330
         Index           =   0
         Left            =   375
         TabIndex        =   11
         Top             =   525
         Width           =   3210
      End
      Begin VB.CommandButton cmd 
         Caption         =   "Select Project..."
         Height          =   315
         Left            =   405
         TabIndex        =   7
         Top             =   1935
         Width           =   1725
      End
      Begin VB.OptionButton Option1 
         Caption         =   "OLEDB/ADO data binding"
         Height          =   330
         Index           =   2
         Left            =   375
         TabIndex        =   6
         Top             =   1125
         Width           =   3210
      End
      Begin VB.OptionButton Option1 
         Caption         =   "DAO/RDO data binding"
         Height          =   330
         Index           =   1
         Left            =   375
         TabIndex        =   5
         Top             =   825
         Value           =   -1  'True
         Width           =   3630
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1) The new project should use VSFlex7 with:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   150
         TabIndex        =   9
         Top             =   270
         Width           =   4605
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2) The original project is:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   150
         TabIndex        =   8
         Top             =   1635
         Width           =   3300
         WordWrap        =   -1  'True
      End
   End
   Begin MSComDlg.CommonDialog dlg 
      Left            =   5700
      Top             =   930
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ListBox lst 
      Height          =   3225
      IntegralHeight  =   0   'False
      ItemData        =   "Convert.frx":0442
      Left            =   4815
      List            =   "Convert.frx":0449
      TabIndex        =   0
      Top             =   435
      Width           =   2310
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Project Files:"
      Height          =   195
      Left            =   4815
      TabIndex        =   3
      Top             =   195
      Width           =   1215
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Note: The VB source code for this utility is included in the VSFlexGrid Pro 7.0 distribution package."
      Height          =   420
      Left            =   195
      TabIndex        =   2
      Top             =   4380
      Width           =   7110
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   $"Convert.frx":0452
      Height          =   1095
      Index           =   0
      Left            =   195
      TabIndex        =   1
      Top             =   225
      Width           =   4410
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

' Conversion constants
'
' VSFlex6 OLEDB/ADO strings
Const F6_NAME = "VSFlex6 (OLEDB/ADO version)"
Const F6_VBPREF = "Object={8099FCC2-0A81-11D2-BAA4-04F205C10000}#1.0#0; VSFLEX6.OCX"
Const F6_FRMREF = "Object = ""{8099FCC2-0A81-11D2-BAA4-04F205C10000}#1.0#0""; ""VSOCX6.OCX"""
Const F6_OFAREF = "Begin VSFlex6Ctl.vsFlexGrid"
Const F6_OFSREF = "Begin VSFlex6Ctl.vsFlexString"
Const F6_ODTREF = "VSFlex6Ctl.vsDataObject"
    
' VSFlex6 DAO strings
Const F6D_NAME = "VSFlex6 (DAO/RDO version)"
Const F6D_VBPREF = "Object={C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0; VSFLEX6D.OCX"
Const F6D_FRMREF = "Object = ""{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0""; ""VSFLEX6D.OCX"""
Const F6D_OFAREF = "Begin VSFlex6DAOCtl.vsFlexGrid"
Const F6D_OFSREF = "Begin VSFlex6DAOCtl.vsFlexString"
Const F6D_ODTREF = "VSFlex6DAOCtl.vsDataObject"

' MSFlexGrid strings
Const MS_NAME = "MSFlexGrid"
Const MS_VBPREF = "Object={5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0; MSFLXGRD.OCX"
Const MS_FRMREF = "Object = ""{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0""; ""MSFLXGRD.OCX"""
Const MS_OFAREF = "Begin MSFlexGridLib.MSFlexGrid"
Const MS_ODTREF = "MSFlexGridLib.DataObject"

' VSFlex7 OLEDB/ADO strings
Const F7_NAME = "VSFlex7 (OLEDB/ADO version)"
Const F7_VBPREF = "Object={8099FCC2-0A81-11D2-BAA4-04F205C10000}#1.0#0; VSFLEX7.OCX"
Const F7_FRMREF = "Object = ""{8099FCC2-0A81-11D2-BAA4-04F205C10000}#1.0#0""; ""VSOCX7.OCX"""
Const F7_OFAREF = "Begin VSFlex7Ctl.VSFlexGrid"
Const F7_ODTREF = "VSFlex7Ctl.VSDataObject"
    
' VSFlex7 DAO strings
Const F7D_NAME = "VSFlex7 (DAO/RDO version)"
Const F7D_VBPREF = "Object={D76D7130-4A96-11D3-BD95-D296DC2DD072}#1.0#0; VSFLEX7D.OCX"
Const F7D_FRMREF = "Object = ""{D76D7130-4A96-11D3-BD95-D296DC2DD072}#1.0#0""; ""VSFLEX7D.OCX"""
Const F7D_OFAREF = "Begin VSFlex7DAOCtl.VSFlexGrid"
Const F7D_ODTREF = "VSFlex7DAOCtl.VSDataObject"

' VSFlex7 Light strings
Const F7L_NAME = "VSFlex7 (Light version)"
Const F7L_VBPREF = "Object={C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0; VSFLEX7L.OCX"
Const F7L_FRMREF = "Object = ""{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0""; ""VSFLEX7L.OCX"""
Const F7L_OFAREF = "Begin VSFlex7LCtl.VSFlexGrid"
Const F7L_ODTREF = "VSFlex7LCtl.VSDataObject"

' Variables to hold current conversion
Dim OLD_NAME$, NEW_NAME$
Dim OLD_VBPREF$, NEW_VBPREF$
Dim OLD_FRMREF$, NEW_FRMREF$
Dim OLD_OFAREF$, NEW_OFAREF$
Dim OLD_OFSREF$, NEW_OFSREF$
Dim OLD_ODTREF$, NEW_ODTREF$


Sub ConvertControl(s$)
    Dim bFlexGrid%, bFlexString%
    
    ' keep track of the control type
    ' we will need this to filter the properties
    '
    If StringReplace(s, OLD_OFAREF, NEW_OFAREF) Then
        bFlexGrid = True
        
        ' keep track of the VSFlexGrid control name
        ' we will need this to convert events
        '
        Dim cn$
        cn = s
        While InStr(cn, " ") > 0
            cn = Trim(Mid(cn, InStr(cn, " ") + 1))
        Wend
        lstCtrls.AddItem cn
    ElseIf StringReplace(s, OLD_OFSREF, NEW_OFSREF) Then
        bFlexString = True
    End If
    
    ' write control declaration
    '
    Print #2, s
    
    ' honor MSFlexGrid defaults for Rows/Cols
    '
    Dim r%, c%
    r = 2
    c = 2
    
    ' loop converting the control
    '
    While Not EOF(1)
    
        ' get a line
        '
        Line Input #1, s
        
        ' honor MSFlexGrid defaults for Rows/Cols
        '
        If bFlexGrid And (OLD_NAME = MS_NAME) Then
            If Left(Trim(s), 5) = "Rows " Then r = -1
            If Left(Trim(s), 5) = "Cols " Then c = -1
        End If
        
        ' handle nested controls using recursive call
        '
        If Left(Trim(s), 6) = "Begin " Then
            ConvertControl s
        
        ' at the end of the control, bail out
        '
        ElseIf Trim(s) = "End" Then
        
            ' honor MSFlexGrid defaults for Rows/Cols
            If bFlexGrid And (OLD_NAME = MS_NAME) Then
                If r > -1 Then Print #2, "Rows = "; r
                If c > -1 Then Print #2, "Cols = "; c
            End If
            
            ' finish
            Print #2, s
            Exit Sub
        
        ' filter MFC-specific properties.
        '
        Else
            FilterPropControl s
        End If
    Wend
    
    ' we should never get here
    '
    Debug.Print "unexpected EOF"

End Sub

Sub ConvertEvents(ByRef s$)
    Dim i%, f$, r$
    
    ' if this is not an event handler, we're done
    '
    If InStr(s, "Private Sub") = 0 Then Exit Sub
    If InStr(s, "_") = 0 Then Exit Sub
    
    ' get the control name
    '
    f = Left(s, InStr(s, "_") - 1)
    While InStr(f, " ") > 0
        f = Trim(Mid(f, InStr(f, " ") + 1))
    Wend
    
    ' make sure it's on our list
    '
    For i = 0 To lstCtrls.ListCount - 1
        If f = lstCtrls.List(i) Then Exit For
    Next
    If i >= lstCtrls.ListCount Then Exit Sub
    
    ' custom compare event is slightly different:
    '
    f = "_Compare(Row1 As Long, Row2 As Long, Cmp As Integer)"
    r = "_Compare(ByVal Row1 As Long, ByVal Row2 As Long, Cmp As Integer)"
    If InStr(s, f) > 0 Then StringReplace s, f, r
    
    ' UserResized event renamed for consistency
    '
    f = "_UserResized(ByVal Row As Long, ByVal Col As Long)"
    r = "_AfterUserResize(ByVal Row As Long, ByVal Col As Long)"
    If InStr(s, f) > 0 Then StringReplace s, f, r
    
    ' Scroll event renamed for consistency
    '
    f = "_Scroll()"
    r = "_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)"
    If InStr(s, f) > 0 Then StringReplace s, f, r
    
    ' Validate event renamed for compatibility with VB6
    '
    f = "_Validate(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)"
    r = "_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)"
    If InStr(s, f) > 0 Then StringReplace s, f, r

    ' OLE DragDrop events have DataObject parameter
    '
    If InStr(s, "_OLEDragDrop(") > 0 Or InStr(s, "_OLEDragOver(") > 0 Or InStr(s, "_OLEStartDrag(") > 0 Or InStr(s, "_OLESetData(") > 0 Then
        StringReplace s, OLD_ODTREF, NEW_ODTREF
        If OLD_NAME = MS_NAME Then
            StringReplace s, "Button As Integer", "ByVal Button As Integer"
            StringReplace s, "Shift As Integer", "ByVal Shift As Integer"
            StringReplace s, "X As Single", "ByVal X As Single"
            StringReplace s, "Y As Single", "ByVal Y As Single"
        End If
    End If

End Sub

Sub ConvertForm(ByVal pth$, fn$)
    Dim s$
    
    ' if the file has a path spec, don't prepend the project path
    '
    If InStr(fn, "\") = 0 Then pth = ""
    
    ' make a backup copy of the file
    ' we just append '.bak' to the file name
    ' the file will have two extensions, which is OK in Win32
    '
    FileCopy pth & fn, pth & fn & ".bak"
    
    ' open files to read (backup) and write (converted)
    '
    Open pth & fn & ".bak" For Input As #1
    Open pth & fn For Output As #2
    
    ' reset list of converted controls
    lstCtrls.Clear
    
    ' scan the file
    '
    While Not EOF(1)
        
        ' read a line
        '
        Line Input #1, s
        
        ' convert object reference
        '
        If StringReplace(s, OLD_FRMREF, NEW_FRMREF) Then
            Print #2, s
        
        ' convert controls
        '
        ElseIf Left(Trim(s), 6) = "Begin " Then
            ConvertControl s
        
        ' convert event declarations, if any are needed,
        ' then copy whatever is left to the new file
        '
        Else
            ConvertEvents s
            Print #2, s
        End If
    Wend
    
    ' done with this form, close both files
    '
    Close #1
    Close #2
End Sub

Sub FilterPropControl(s)
    Dim i%
    
    '----------------------------------------------------
    ' remove some properties
    '
    If InStr(s, "_Version ") > 0 Then Exit Sub
    If InStr(s, "_StockProps ") > 0 Then Exit Sub
    
    '----------------------------------------------------
    ' write out whatever is left as-is
    '
    Print #2, s

End Sub


Function GetPropValue(ByVal s$) As Variant
    
    ' parse a property assignment string,
    ' return portion after the equals sign.
    '
    If InStr(s, "=") > 0 Then
        GetPropValue = Trim(Mid(s, InStr(s, "=") + 1))
    End If

End Function

Function StringReplace(s$, find$, repl$) As Boolean
    Dim i%
    
    ' find target string, return False if failed
    '
    i = InStr(s, find)
    If i = 0 Then Exit Function
    If Len(find) = 0 Then Exit Function
    
    ' make all the replacements
    '
    While i > 0
        s = Left(s, i - 1) & repl & Mid(s, i + Len(find))
        i = i + Len(repl)
        i = InStr(i + 1, s, find)
    Wend
    
    ' return True to indicate replacements were made
    '
    StringReplace = True
    
End Function


Private Sub cmd_Click()
    Dim prj$, pth$, s$, i%, j%
    
    '=======================================================
    ' open existing project
    
    ' get VB project name
    '
    dlg.CancelError = False
    dlg.FileName = ""
    dlg.DefaultExt = "vbp"
    dlg.Filter = "Visual Basic Projecs (*.vbp)|*.vbp"
    dlg.Flags = cdlOFNFileMustExist
    dlg.ShowOpen
    prj = dlg.FileName
    If prj = "" Then Exit Sub
    lst.Clear
    
    ' parse project path and file name
    '
    i = 0
    While InStr(i + 1, prj, "\") > 0
        i = InStr(i + 1, prj, "\")
    Wend
    pth = Left(prj, i)
    prj = Mid(prj, i + 1)
    
    ' read the whole project file
    '
    Open pth & prj For Binary As #1
    s = Input(LOF(1), 1)
    Close #1
    
    '=======================================================
    ' scan existing project to detect library type
    
    ' see if it's a Flex6 OLEDB/ADO project
    '
    OLD_NAME = ""
    If InStr(s, F6_VBPREF) > 0 Then
        OLD_NAME = F6_NAME
        OLD_VBPREF = F6_VBPREF
        OLD_FRMREF = F6_FRMREF
        OLD_OFAREF = F6_OFAREF
        OLD_OFSREF = F6_OFSREF
        OLD_ODTREF = F6_ODTREF
    
    ' if not, see if it's a Flex6 DAO/RDO project
    '
    ElseIf InStr(s, F6D_VBPREF) > 0 Then
        OLD_NAME = F6D_NAME
        OLD_VBPREF = F6D_VBPREF
        OLD_FRMREF = F6D_FRMREF
        OLD_OFAREF = F6D_OFAREF
        OLD_OFSREF = F6D_OFSREF
        OLD_ODTREF = F6D_ODTREF
    
    ' if not, see if it's an MSFlexGrid project
    '
    ElseIf InStr(s, MS_VBPREF) > 0 Then
        OLD_NAME = MS_NAME
        OLD_VBPREF = MS_VBPREF
        OLD_FRMREF = MS_FRMREF
        OLD_OFAREF = MS_OFAREF
'        OLD_OFSREF = MS_OFSREF
        OLD_ODTREF = MS_ODTREF
    End If
    
    '=======================================================
    ' select library to convert to
    
    ' convert to LIGHT version
    '
    If Option1(0).Value Then
        NEW_NAME = F7L_NAME
        NEW_VBPREF = F7L_VBPREF
        NEW_FRMREF = F7L_FRMREF
        NEW_OFAREF = F7L_OFAREF
'todo        NEW_OFSREF = F7L_OFSREF
        NEW_ODTREF = F7L_ODTREF
    
    ' convert to DAO version
    '
    ElseIf Option1(1).Value Then
        NEW_NAME = F7D_NAME
        NEW_VBPREF = F7D_VBPREF
        NEW_FRMREF = F7D_FRMREF
        NEW_OFAREF = F7D_OFAREF
'todo        NEW_OFSREF = F7D_OFSREF
        NEW_ODTREF = F7D_ODTREF
        
    ' convert to ADO version
    Else
        NEW_NAME = F7_NAME
        NEW_VBPREF = F7_VBPREF
        NEW_FRMREF = F7_FRMREF
        NEW_OFAREF = F7_OFAREF
'todo        NEW_OFSREF = F7_OFSREF
        NEW_ODTREF = F7_ODTREF
    End If
    
    '=======================================================
    ' validate what we have so far
    If OLD_NAME = "" Then
        MsgBox "This project does not use VSFlex6, VSFlex7, or MSFlexGrid." & vbCrLf & vbCrLf & "Conversion cancelled.", vbInformation
        Exit Sub
    End If
    If OLD_NAME = NEW_NAME Then
        MsgBox "This project already uses " & NEW_NAME & "." & vbCrLf & vbCrLf & "No conversion is necessary.", vbInformation
        Exit Sub
    End If
    
    '=======================================================
    ' convert project
    
    ' replace library references.
    '
    StringReplace s, OLD_VBPREF, NEW_VBPREF
    
    ' confirm conversion
    '
    i = MsgBox("Ready to convert '" & prj & "' from " & OLD_NAME & " to " & NEW_NAME & "." & vbCrLf & vbCrLf & _
               "The original files will be backed up in files with an additional '.bak' extension." & vbCrLf & vbCrLf & _
               "Press OK to proceed or Cancel if you changed your mind.", _
               vbOKCancel Or vbInformation)
    If i <> vbOK Then Exit Sub
    
    ' don't disturb me while I work
    '
    Me.Enabled = False
    
    ' make a backup copy of the project file
    ' we just append '.bak' to the file name
    ' the file will have two extensions, which is OK in Win32
    '
    FileCopy prj, prj & ".bak"
    
    ' write out updated project file
    '
    Open prj For Output As #1
    Print #1, s
    Close #1
    lst.AddItem prj & "... OK"
    
    ' find all forms in the project
    '
    i = 1
    Do
        i = InStr(i, s, "Form=")
        If i = 0 Then Exit Do
        If Mid(s, i + 5, 1) <> """" Then
            j = InStr(i, s, ".frm")
            If j <> 0 Then lst.AddItem Mid(s, i + 5, j - i - 5 + 4)
        End If
        i = i + 1
    Loop
    
    ' find all UserControls in the project
    '
    i = 1
    Do
        i = InStr(i, s, "UserControl=")
        If i = 0 Then Exit Do
        If Mid(s, i + 12, 1) <> """" Then
            j = InStr(i, s, ".ctl")
            If j <> 0 Then lst.AddItem Mid(s, i + 12, j - i - 12 + 4)
        End If
        i = i + 1
    Loop
    
    ' process form files
    '
    For i = 1 To lst.ListCount - 1
        ConvertForm pth, lst.List(i)
        lst.ListIndex = i
        lst.List(i) = lst.List(i) & "... OK"
        Refresh
        DoEvents
    Next
    
    ' show confirmation and warnings
    '
    s = "Project '" & prj & "' has been converted from " & OLD_NAME & " to " & NEW_NAME & "." & vbCrLf & vbCrLf & _
        "Copies of the original files were saved with a 'bak' extension in case you want to restore the original project." & vbCrLf & vbCrLf & _
        "You should now open the new project in VB and compile it to make sure there are no syntax errors." & vbCrLf & vbCrLf
    MsgBox s, vbInformation Or vbOKOnly
    
    ' all done, re-activate form
    '
    Me.Enabled = True
End Sub

Private Sub Form_Load()
    
    ' just clear our file list
    '
    lst.Clear

End Sub

