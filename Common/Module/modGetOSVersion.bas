Attribute VB_Name = "modGetOSVersion"
Option Explicit

Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

Public Const ERROR_SUCCESS = 0&

Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_CURRENT_USER = &H80000001

Public Const STANDARD_RIGHTS_ALL = &H1F0000
Public Const KEY_QUERY_VALUE = &H1
Public Const KEY_SET_VALUE = &H2
Public Const KEY_CREATE_SUB_KEY = &H4
Public Const KEY_ENUMERATE_SUB_KEYS = &H8
Public Const KEY_NOTIFY = &H10
Public Const KEY_CREATE_LINK = &H20
Public Const SYNCHRONIZE = &H100000
Public Const KEY_ALL_ACCESS = _
    ((STANDARD_RIGHTS_ALL Or _
    KEY_QUERY_VALUE Or _
    KEY_SET_VALUE Or _
    KEY_CREATE_SUB_KEY Or _
    KEY_ENUMERATE_SUB_KEYS Or _
    KEY_NOTIFY Or KEY_CREATE_LINK) And _
    (Not SYNCHRONIZE))

Public Declare Function GetVersionExA Lib "kernel32" (lpVersionInformation As OSVERSIONINFO) As Integer
Public Type OSVERSIONINFO
   dwOSVersionInfoSize As Long
   dwMajorVersion As Long
   dwMinorVersion As Long
   dwBuildNumber As Long
   dwPlatformId As Long
   szCSDVersion As String * 128
End Type

Public Const REG_PRODUCT_KEY As String = "SOFTWARE\Microsoft\Windows NT\CurrentVersion"

' Return the registry value.
Public Function GetRegistryValue(ByVal hKey As Long, ByVal subkey_name As String) As String
Dim value As String
Dim length As Long
Dim value_type As Long

    length = 256
    value = Space$(length)
    If RegQueryValueEx(hKey, subkey_name, _
        0&, value_type, ByVal value, length) _
            <> ERROR_SUCCESS _
    Then
        value = "<Error>"
    Else
        ' Remove the trailing null character.
        value = Left$(value, length - 1)
    End If
    
    GetRegistryValue = value
End Function

' Return the OS version, build, and platform ID.
'            dwPlatformID  dwMajorVersion  dwMinorVersion
' Win3.1          0               ?               ?
' Win95           1               4               0
' Win98           1               4              10
' WinME           1               4              90
' NT 3.51         2               3              51
' NT 4.0          2               4               0
' Win2000         2               5               0
' WinXP           2               5               1
' Win2003         2               5               2
' Vista           2               6               0
' Win2008 Server  2               6               0
Public Function GetOSVersion() As String
Dim version_info As OSVERSIONINFO
Dim txt As String
Dim hKey As Long

    GetOSVersion = "Unknown"

    version_info.dwOSVersionInfoSize = 148
    version_info.szCSDVersion = Space$(128)
    GetVersionExA version_info

    Select Case version_info.dwPlatformId
        Case 0
            GetOSVersion = "Win 3.1"
        Case 1
            Select Case version_info.dwMinorVersion
                Case 0
                    GetOSVersion = "Win95"
                Case 10
                    GetOSVersion = "Win98"
                Case 90
                    GetOSVersion = "WinME"
                Case Else
                    GetOSVersion = "Unknown Win32"
            End Select
        Case 2
            Select Case version_info.dwMajorVersion
                Case 3
                    GetOSVersion = "NT 3.51"
                Case 4
                    GetOSVersion = "NT 4.0"
                Case 5
                    Select Case version_info.dwMinorVersion
                        Case 0
                            GetOSVersion = "Win2000"
                        Case 1
                            GetOSVersion = "WinXP"
                        Case 2
                            GetOSVersion = "Win2003"
                    End Select
                Case 6
                    If RegOpenKeyEx(HKEY_LOCAL_MACHINE, _
                        REG_PRODUCT_KEY, _
                        0&, KEY_QUERY_VALUE, hKey) <> ERROR_SUCCESS _
                    Then
                        txt = ""
                    Else
                        txt = GetRegistryValue(hKey, "ProductName")

                        ' Close the key.
                        If RegCloseKey(hKey) <> ERROR_SUCCESS Then
                            MsgBox "Error closing key."
                        End If
                    End If

                    txt = LCase$(txt)
                    If InStr(txt, "vista") > 0 Then
                        GetOSVersion = "Vista"
                    ElseIf InStr(txt, "server") > 0 Then
                        GetOSVersion = "Win2008"
                    Else
                        GetOSVersion = "Unknown Win32NT6"
                    End If
                Case Else
                    GetOSVersion = "Unknown Win32NT"
            End Select
        Case 3
            GetOSVersion = "Win CE"
        Case 4
            GetOSVersion = "Unix"
    End Select

    GetOSVersion = "Unknown"
End Function

