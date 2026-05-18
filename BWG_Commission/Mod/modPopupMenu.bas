Attribute VB_Name = "modPopupMenu"
Option Explicit

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
    X As Long
    Y As Long
End Type

Public Declare Function CreatePopupMenu Lib "user32" () As Long
Public Declare Function TrackPopupMenuEx Lib "user32" (ByVal hMenu As Long, _
                ByVal wFlags As Long, ByVal X As Long, ByVal Y As Long, _
                ByVal HWnd As Long, ByVal lptpm As Any) As Long
Public Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" _
                (ByVal hMenu As Long, _
                ByVal wFlags As BTN_STYLE, ByVal wIDNewItem As Long, _
                ByVal lpNewItem As Any) As Long
Public Declare Function DestroyMenu Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public hMenu As Long

