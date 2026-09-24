Attribute VB_Name = "modAPI"
Option Explicit
'Public Localization As New CLSCulture  'ปิดสำหรับ Waste Data

Public Enum BTN_STYLE
    MF_CHECKED = &H8&
    MF_APPEND = &H100&
    TPM_LEFTALIGN = &H0&
    MF_DISABLED = &H2&
    MF_GRAYED = &H1&
    MF_SEPARATOR = &H800&
    MF_STRING = &H0&
    TPM_RETURNCMD = &H100&
    TPM_RIGHTBUTTON = &H2&
End Enum

Public Type POINTAPI
    x As Long
    Y As Long
End Type

Public Type LVHITTESTINFO
   Pt As POINTAPI
   Flags As Long
   iItem As Long
   iSubItem  As Long
End Type

Public Const PS_SOLID            As Long = 0

Public Type RECT
   Left                 As Long
   Top                  As Long
   Right                As Long
   Bottom               As Long
End Type

Public Declare Function CreatePopupMenu Lib "user32" () As Long
Public Declare Function TrackPopupMenuEx Lib "user32" (ByVal hMenu As Long, _
                ByVal wFlags As Long, ByVal x As Long, ByVal Y As Long, _
                ByVal hwnd As Long, ByVal lptpm As Any) As Long
Public Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" _
                (ByVal hMenu As Long, _
                ByVal wFlags As BTN_STYLE, ByVal wIDNewItem As Long, _
                ByVal lpNewItem As Any) As Long
Public Declare Function DestroyMenu Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public hMenu As Long

Public Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, _
    ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, _
    ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function FrameRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long, ByVal hBrush As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function sndPlaySound Lib "winmm" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SWP_SHOWWINDOW = &H40
Const HWND_NOTOPMOST = -2
Const HWND_TOPMOST = -1

Public Type tagInitCommonControlsEx
   lngSize As Long
   lngICC As Long
End Type
Public Declare Function InitCommonControlsEx Lib "comctl32.dll" (iccex As tagInitCommonControlsEx) As Boolean
Public Const ICC_USEREX_CLASSES = &H200

Public Declare Function RegCreateKey Lib "advapi32.dll" Alias _
"RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, _
phkResult As Long) As Long

Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias _
"RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, _
ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal _
cbData As Long) As Long

Public Declare Function RegCloseKey Lib "advapi32.dll" _
(ByVal hKey As Long) As Long


Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias _
"RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, _
ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As _
Long) As Long


' Note that if you declare the lpData parameter as String, you must pass it By Value.
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias _
"RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, _
ByVal lpReserved As Long, lpType As Long, lpData As Any, _
lpcbData As Long) As Long

Public Declare Function RegQueryValueExString Lib "advapi32.dll" Alias _
"RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As _
String, ByVal lpReserved As Long, lpType As Long, ByVal lpData _
As String, lpcbData As Long) As Long

Public Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias _
"RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As _
String, ByVal lpReserved As Long, lpType As Long, lpData As _
Long, lpcbData As Long) As Long

Public Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias _
"RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As _
String, ByVal lpReserved As Long, lpType As Long, ByVal lpData _
As Long, lpcbData As Long) As Long

' Set a form always on the top.
'
' the form can be specified as a Form or object
' or through its hWnd property
' If OnTop=False the always on the top mode is de-activated.

Public Sub SetAlwaysOnTopMode(hWndOrForm As Variant, Optional ByVal OnTop As Boolean = _
    True)
    Dim hwnd As Long
    ' get the hWnd of the form to be move on top
    If VarType(hWndOrForm) = vbLong Then
        hwnd = hWndOrForm
    Else
        hwnd = hWndOrForm.hwnd
    End If
    SetWindowPos hwnd, IIf(OnTop, HWND_TOPMOST, HWND_NOTOPMOST), 0, 0, 0, 0, _
        SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW
End Sub

Public Sub CutRR(Obj As Object, Rad As Integer)
   Dim TwipPix          As Integer
   TwipPix = Screen.TwipsPerPixelX
   SetWindowRgn Obj.hwnd, CreateRoundRectRgn(Obj.Width / TwipPix, Obj.Height / TwipPix, 0, 0, (Obj.Width / TwipPix) / Rad, (Obj.Width / TwipPix) / Rad), True
End Sub

Public Sub FormDrag(TheForm As Form)
   ReleaseCapture
   SendMessage TheForm.hwnd, &HA1, 2, 0&
End Sub

Public Sub MakeFormRounded(Obj As Object, _
   ByVal Radius As Long, _
   Optional ByVal Colour As OLE_COLOR = vbBlack, _
   Optional ByVal Bevel As Long = 0, _
   Optional Texture As StdPicture = Nothing)

   Dim hRgn             As Long
   Dim hHBr             As Long

   Obj.ScaleMode = vbPixels

   hRgn = CreateRoundRectRgn(0, 0, Obj.ScaleWidth, Obj.ScaleHeight, Radius, Radius)

   If Not Texture Is Nothing Then
      'Create a brush from picture
      hHBr = CreatePatternBrush(Texture)
   Else
      If Bevel > 0 Then
         hHBr = CreatePen(PS_SOLID, Bevel, Colour)
      End If
   End If

   'Draw accurate frame around region
   FrameRgn Obj.hdc, hRgn, hHBr, Bevel, Bevel
   SetWindowRgn Obj.hwnd, hRgn, True
   'Clean up
   DeleteObject hRgn
   DeleteObject hHBr

End Sub
