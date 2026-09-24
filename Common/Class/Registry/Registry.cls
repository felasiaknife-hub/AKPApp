VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Registry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'********************************************************************
'*Class/Method Name:    Registry
'*Version:              1.0
'*File Name:            Registry.cls
'*Decription:           Registry class for 3F Aggregator
'*                      - handles all registry & crypto methods to be
'*                        called by both client app and engine
'*Author:               TEWee
'*Company:              3rdFrontier (M) Sdn Bhd
'*Project:              3F Aggregator (Copyright (c)2001)
'*Date submitted:
'********************************************************************
'*MODIFICATION HISTORY
'* Name              Date        Version      Description
'*-------           ------      ---------    --------------
'* TEW              01/08/2001  Beta 1.0    Basic functionality
'*                                          completed
'* TEW              16/08/2001  Beta 1.0m   Added DeleteRegistryKey method to
'*                                          delete key from registry
'* LSN              27/08/2001              Added
'*                                          - DeleteRegistryKeyLoop
'*                                              * Enhance over the DeleteRegistryKey to enable the key to be deleted from
'*                                                the NT.
'*                                          - DeleteSubRegKey
'*                                              * Private function to be called by DeleteRegistryKeyLoop.
'*                                                This is a recursive call which will call itself until no
'*                                                key available for deletion.
'*                                          - fEnumKey
'*                                              * This function will retrieve a list of the key
'*                                          - fEnumValue
'*                                              * Function for retrieving a list of the key value pair
'*                                          - fTopKey
'*                                              * Function to get the registry top level constant
'*                                          - fRenameRegKey
'*                                              * Function to rename the key to the specified key
'*                                          - fCopyRegKey
'*                                              * To copy from 1 key to another key, including its
'*                                                sub-key and value pair
'*                                          - DeleteRegValue
'*                                              * To delete a specified value from the registry
'*
'* LSN              09/10/2001              Added
'*                                          - ProcessNull
'*                                              * Provide the conversion on chr(0) to the 'CHR(0)' character
'*                                          - DeprocessNull
'*                                              * Provide the deconversion on 'CHR(0)' back to the chr(0)
'*                                          - Modify the read registry value method to support the ability
'*                                            to process the chr(0) in the string value.
'********************************************************************/
'

Option Explicit
Const REG_NONE = 0
Const REG_SZ = 1 ' Unicode nul terminated string
Const REG_BINARY = 3 ' Free form BINARY
Const REG_DWORD = 4  ' Free form DWORD
Const HKEY_CURRENT_USER = &H80000001
Const strBackSlash As String = "\"
Const strEqual As String = "="
 
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hkey As Long) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hkey As Long, _
        ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hkey As Long, _
        ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
'Private Declare Function RegQueryValueExW Lib "advapi32.dll" (ByVal hkey As Long, _
'        ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long

Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hkey As Long, _
        ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hkey As Long, _
        ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As String, ByVal cbData As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hkey As Long, _
        ByVal lpSubKey As String) As Long

'''''''''''''''''''''''''''''' New Registry function
'Registry module.
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hkey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hkey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegFlushKey Lib "advapi32.dll" (ByVal hkey As Long) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hkey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hkey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long
Private Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hkey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long

Public Enum REGIS_KEY_CONST
    KEY_QUERY_VALUE = &H1
    KEY_SET_VALUE = &H2
    KEY_CREATE_SUB_KEY = &H4
    KEY_ENUMERATE_SUB_KEYS = &H8
    KEY_NOTIFY = &H10
    KEY_CREATE_LINK = &H20
    KEY_ALL_ACCESS = KEY_QUERY_VALUE And KEY_ENUMERATE_SUB_KEYS And KEY_NOTIFY And KEY_CREATE_SUB_KEY And KEY_CREATE_LINK And KEY_SET_VALUE
    REG_OPTION_NON_VOLATILE = 0       ' Key is preserved when system is rebooted
    REG_OPTION_VOLATILE = 1           ' Key is not preserved when system is rebooted
    ERROR_SUCCESS = 0&
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_CONFIG = &H80000005
    HKEY_DYN_DATA = &H80000006
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_PERFORMANCE_DATA = &H80000004
    HKEY_USERS = &H80000003
End Enum
'''' 2nd Part
Const MAX_SIZE = 2048

' constant value for the process/deprocess null method
Const CONST_strCHR0 = "CHR(0)"

Public Function isRegPathExist(ByVal rPath As String, Optional regRoot As REGIS_KEY_CONST = HKEY_CURRENT_USER) As Boolean
Dim Ret
    RegOpenKey regRoot, rPath, Ret
    
    If Ret = 0 Then
        isRegPathExist = False
    Else
        isRegPathExist = True
    End If
    
    RegCloseKey Ret
End Function

Public Function GetRegisterValue(ByVal strFullPathName As String, ByVal strValueName As String) As String
    Dim Ret As Long
    Dim lResult As Long, lValueType As Long, strBuf As String, lDataBufSize As Long
    'Debug.Print strFullPathName & " - " & strValueName
    GetRegisterValue = vbNullString
    'get the key's content
    RegOpenKey HKEY_CURRENT_USER, strFullPathName, Ret
        
    'retrieve information about the key
    lResult = RegQueryValueEx(Ret, strValueName, 0, lValueType, ByVal 0, lDataBufSize)
    'lResult = FindBufferSize(Ret, strValueName, 0, lValueType, 0, lDataBufSize)
    If lResult = 0 Then
        If lValueType = REG_SZ Then
            'Create a buffer
            strBuf = String(lDataBufSize, Chr$(0))
            'retrieve the key's content
            lResult = RegQueryValueEx(Ret, strValueName, 0, 0, ByVal strBuf, lDataBufSize)
            If lResult = 0 Then
            'Remove the unnecessary chr$(0)'s
                If strBuf <> vbNullString Then
                    GetRegisterValue = Left$(strBuf, InStr(1, strBuf, Chr$(0)) - 1)
                Else
                    GetRegisterValue = vbNullString
                End If
            End If
        ElseIf lValueType = REG_BINARY Then
            Dim strData As Integer            'retrieve the key's value
            lResult = RegQueryValueEx(Ret, strValueName, 0, 0, strData, lDataBufSize * 2)
            If lResult = 0 Then
                GetRegisterValue = strData
            End If
        ElseIf lValueType = REG_DWORD Then
            Dim longdata As Long
            lResult = RegQueryValueEx(Ret, strValueName, 0, 0, longdata, lDataBufSize * 2)
            If lResult = 0 Then
                GetRegisterValue = longdata
            End If
        End If
    End If
    RegCloseKey Ret
    GetRegisterValue = DeprocessNull(GetRegisterValue)
End Function

Public Function SetRegisterValue(ByVal strFullPathName As String, ByVal strValueName As String, _
ByVal strValue As String) As Boolean
    Dim Ret As Long
    Dim lResult, lValueType As Long, lDataBufSize As Long
    strValue = ProcessNull(strValue)
        
    'Create a new key
    RegCreateKey HKEY_CURRENT_USER, strFullPathName, Ret
    lResult = RegQueryValueEx(Ret, strValueName, 0, lValueType, ByVal 0, lDataBufSize)
    'lResult = FindBufferSize(Ret, strValueName, 0, lValueType, 0, lDataBufSize)
    If lResult = 0 Then
        'Save a string to the key
        If lValueType = REG_SZ Then
            lResult = RegSetValueEx(Ret, strValueName, 0, REG_SZ, ByVal strValue, Len(strValue))
        ElseIf lValueType = REG_BINARY Then
            lResult = RegSetValueEx(Ret, strValueName, 0, REG_BINARY, CByte(strValue), 4)
        ElseIf lValueType = REG_DWORD Then
            lResult = RegSetValueEx(Ret, strValueName, 0, REG_DWORD, ByVal Chr$(strValue), 4)
        End If
    ElseIf lResult = 2 Then
            lResult = RegSetValueEx(Ret, strValueName, 0, REG_SZ, ByVal strValue, Len(strValue))
        
    End If
    'close the key
    RegCloseKey Ret
    
    If lResult = 0 Then
        SetRegisterValue = True
    End If
End Function

Public Function DeleteRegistryKey(fullPath As String, SubKeyName As String) As Boolean

    Dim Ret As Long     'Create a new key
    Dim lResult As Long
    RegCreateKey HKEY_CURRENT_USER, fullPath, Ret       'Delete the key's value
    lResult = RegDeleteKey(Ret, SubKeyName)     'close the key
    RegCloseKey Ret
    
    If lResult = 0 Then
        DeleteRegistryKey = True
    End If

End Function

Public Function DeleteRegistryKeyLoop(ByVal sTopKey As String, ByVal sSourcePath As String) As Boolean
    ' This is a recursive call to delete the registry key and value from
    ' sSourcePath to the sDestinationPath
    ' Parameter required:
    '   sTopKey = string value that indicate the top level registry key
    '             Possible value will be : "HKCU", "HKLM", "HKU", "HKDD", "HKCC", "HKCR"
    '   sSourcePath = The key path that will copy from
    ' Example:
    '   DeleteRegistryKeyLoop("HKCU", "Software\3F")
    DeleteRegistryKeyLoop = True
    Dim lResult
    Dim lTopKey  As Long

    'Tong On Error GoTo Err_General
    lResult = 99
    lTopKey = fTopKey(sTopKey)
    If lTopKey = 0 Then GoTo Err_General
    
    lResult = DeleteSubRegKey(sTopKey, sSourcePath)
    
    If Len(sSourcePath) > 0 Then
        Dim intOpset As Long
        intOpset = InStrRev(sSourcePath, strBackSlash, , vbTextCompare)
        If intOpset > 0 Then
            DeleteRegistryKeyLoop = DeleteRegistryKey(Mid$(sSourcePath, 1, intOpset - 1), Mid$(sSourcePath, intOpset + 1, Len(sSourcePath) - intOpset))
        End If
    End If
    Exit Function

Err_General:
    DeleteRegistryKeyLoop = False

End Function

Private Function DeleteSubRegKey(ByVal sTopKey As String, ByVal sSourcePath As String) As Boolean
    ' This is a recursive call to delete the registry key and value from
    ' sSourcePath to the sDestinationPath
    ' Parameter required:
    '   sTopKey = string value that indicate the top level registry key
    '             Possible value will be : "HKCU", "HKLM", "HKU", "HKDD", "HKCC", "HKCR"
    '   sSourcePath = The key path that will copy from
    '   sDestinationPath = The key path that will copy to
    ' Example:
    '   DeleteSubRegKey("HKCU", "Software\3F")
    'Tong On Error GoTo Err_General
    Dim arrKeys() As String
    Dim lResult  As Long
    Dim lCount As Long
    Dim intOpset As Integer
    Dim lTopKey As Long
    Dim Ret
    DeleteSubRegKey = True

    ReDim arrKeys(0) As String
    lTopKey = fTopKey(sTopKey)
    If lTopKey = 0 Then GoTo Err_General
    
    ' Get a list of the sub-registry folder
    lResult = fEnumKey(sTopKey, sSourcePath, arrKeys)
    If UBound(arrKeys) > 0 Then
        For lCount = 0 To UBound(arrKeys) - 1
            'Create a new key
            RegCreateKey lTopKey, sSourcePath & strBackSlash & arrKeys(lCount), Ret
            DeleteSubRegKey sTopKey, sSourcePath & strBackSlash & arrKeys(lCount)
            DeleteSubRegKey = DeleteRegistryKey(sSourcePath, arrKeys(lCount))
        Next lCount
    End If
    Exit Function

Err_General:
    DeleteSubRegKey = False
End Function

Public Function DeleteRegValue(ByVal sTopKeyOrFile As String, _
    ByVal sSubKeyOrSection As String, ByVal sValueName As String) As Long
' Use this function to:
'   -   Delete a registry value.
' sTopKeyOrIniFile
'   -   A top level registry key abbreviation {"HKCU","HKLM","HKU","HKDD","HKCC","HKCR"} or
'   -   The full path of an .ini file (ex. "C:\Windows\MyFile.ini")
' sSubKeyOrSection
'   -   A registry subkey or
' sValueName
'   -   A registry entry or
' Return Value
'   -   0 if successful, non-zero otherwise.
' Example 1   -   Delete a registry value.
'   lResult = fDeleteValue("HKCU", "Software\3F", "EntryToDelete")
Dim lTopKey As Long
Dim lHandle As Long
Dim lResult As Long

'Tong On Error GoTo fDeleteValueError
lResult = 99
lTopKey = fTopKey(sTopKeyOrFile)
If lTopKey = 0 Then GoTo fDeleteValueError

lResult = RegOpenKeyEx(lTopKey, sSubKeyOrSection, 0, KEY_SET_VALUE, lHandle)
If lResult = ERROR_SUCCESS Then
    lResult = RegDeleteValue(lHandle, sValueName)
End If
    
If lResult = ERROR_SUCCESS Then
    DeleteRegValue = ERROR_SUCCESS
Else
    DeleteRegValue = lResult
End If
Exit Function

fDeleteValueError:
    DeleteRegValue = lResult
End Function

Public Function fEnumKey(ByVal sTopKey As String, ByVal sSubKey As String, sValues() As String) As Long
' Use this function to:
'   -   Enumerate the subkeys of a registry key (folder).
' sTopKey
'   -   A top level registry key abbreviation {"HKCU","HKLM","HKU","HKDD","HKCC","HKCR"}
' sSubKey
'   -   A registry subkey
' sValues
'   -   A returned string array
' Return Value
'   -   0 if successful, non-zero otherwise.
' Example 1
'   lResult = fEnumKey("HKLM", "Software\Microsoft", sValues)
Dim bDone    As Boolean
Dim lTopKey  As Long
Dim lHandle  As Long
Dim lResult  As Long
Dim lIndex   As Long
Dim sKeyName As String
Dim lCount As Long
lCount = 0

ReDim Preserve sValues(0) As String
'Tong On Error GoTo fEnumKeyError
lResult = 99
lTopKey = fTopKey(sTopKey)
If lTopKey = 0 Then GoTo fEnumKeyError
' Open the registry SubKey.
lResult = RegOpenKeyEx(lTopKey, sSubKey, 0, KEY_ENUMERATE_SUB_KEYS, lHandle)
If lResult <> ERROR_SUCCESS Then GoTo fEnumKeyError
' Get all subkeys until ERROR_NO_MORE_ITEMS or an error occurs.
Do While Not bDone
    sKeyName = Space$(MAX_SIZE)
    lResult = RegEnumKey(lHandle, lIndex, sKeyName, MAX_SIZE)
    
    If lResult = ERROR_SUCCESS Then
        lCount = lCount + 1
        ReDim Preserve sValues(lCount) As String
        If sKeyName <> vbNullString Then
            sValues(lCount - 1) = Left$(sKeyName, InStr(1, sKeyName, Chr$(0)) - 1)
        Else
            sValues(lCount - 1) = vbNullString
        End If
        lIndex = lIndex + 1
    Else
        bDone = True
    End If
Loop
' Close the key.
fEnumKey = RegCloseKey(lHandle)
Exit Function
' Error processing.
fEnumKeyError:
    fEnumKey = lResult
End Function

Public Function fEnumValue(ByVal sTopKeyOrIniFile As String, _
    ByVal sSubKeyOrSection As String, sValues() As String) As Long
' Get a list of a Registry key value pair
' sValues
'   - return string array
' Return Value
'   -   0 if successful, non-zero otherwise.
' Example 1
'   lResult = fEnumValue("HKCU", "Software\YourKey\LastKey\YourApp", sValues)
' NOTE:
'   When enumerating registry values, only string, dword and binary values
'   with a length under 2 bytes (which allows for true/false values) are returned.
Dim lTopKey    As Long
Dim lHandle    As Long
Dim lResult    As Long
Dim lValueLen  As Long
Dim lIndex     As Long
Dim lValue     As Long
Dim lValueType As Long
Dim lData      As Long
Dim lDataLen   As Long
Dim bDone      As Boolean
Dim sValueName As String
Dim sValue     As String
Dim lCount As Long
lCount = 0
'Tong On Error GoTo fEnumValueError
lResult = 99
lTopKey = fTopKey(sTopKeyOrIniFile)
ReDim Preserve sValues(0) As String
If lTopKey = 0 Then GoTo fEnumValueError
    ' Open the registry SubKey.
    lResult = RegOpenKeyEx(lTopKey, sSubKeyOrSection, 0, KEY_QUERY_VALUE, lHandle)
    If lResult <> ERROR_SUCCESS Then GoTo fEnumValueError
    ' Get all values until ERROR_NO_MORE_ITEMS or an error occurs.
    '
    Do While Not bDone
        lDataLen = MAX_SIZE
        lValueLen = lDataLen
        sValueName = Space$(lDataLen)
        
        lResult = RegEnumValue(lHandle, lIndex, sValueName, lValueLen, 0, lValueType, ByVal lData, lDataLen)
        
        If lResult = ERROR_SUCCESS Then
            Select Case lValueType
                Case REG_SZ
                    sValue = Space$(lDataLen)
                    sValueName = Left$(sValueName, lValueLen)
                    lResult = RegQueryValueEx(lHandle, sValueName, 0, REG_SZ, ByVal sValue, lDataLen)
                    If lResult = ERROR_SUCCESS Then
                        lCount = lCount + 1
                        ReDim Preserve sValues(lCount) As String
                        If Trim$(sValue) = vbNullString Then
                            sValues(lCount - 1) = sValueName & strEqual
                        Else
                            sValues(lCount - 1) = sValueName & strEqual & DeprocessNull(Left$(sValue, InStr(1, sValue, Chr$(0)) - 1))
                        End If
                    Else
                        GoTo fEnumValueError
                    End If
                Case REG_DWORD
                    MsgBox "BB"
                    lResult = RegQueryValueEx(lHandle, sValueName, 0, REG_NONE, lValue, lDataLen)
                    If lResult = ERROR_SUCCESS Then
                        lCount = lCount + 1
                        ReDim Preserve sValues(lCount) As String
                        sValueName = Left$(sValueName, lValueLen)
                        If Trim$(sValue) = vbNullString Then
                            sValues(lCount - 1) = sValueName & strEqual
                        Else
                            sValues(lCount - 1) = Left$(sValueName, InStr(1, sValueName, Chr$(0)) - 1) & strEqual & lValue & vbNullChar
                        End If
                    Else
                        GoTo fEnumValueError
                    End If
                Case REG_BINARY
                    MsgBox "cc"
                    If lDataLen <= 2 Then
                        lResult = RegQueryValueEx(lHandle, sValueName, 0, REG_NONE, lValue, lDataLen)
                        If lResult = ERROR_SUCCESS Then
                            lCount = lCount + 1
                            ReDim Preserve sValues(lCount) As String
                            sValueName = Left$(sValueName, lValueLen)
                            If Trim$(sValue) = vbNullString Then
                                sValues(lCount - 1) = sValueName & strEqual
                            Else
                                sValues(lCount - 1) = Left$(sValueName, InStr(1, sValueName, Chr$(0)) - 1) & strEqual & lValue & vbNullChar
                            End If
                        Else
                            GoTo fEnumValueError
                        End If
                    End If
                Case Else
            End Select
            lIndex = lIndex + 1
        Else
            bDone = True
        End If
    Loop
    ' Close the key.
    lResult = RegCloseKey(lHandle)
    fEnumValue = lResult

Exit Function
'
' Error processing.
'
fEnumValueError:
    fEnumValue = lResult
End Function

Private Function fTopKey(ByVal sTopKeyOrFile As String) As Long
Dim sDir   As String
' This function returns:
'   -   the numeric value of a top level registry key or
'   -   1 if sTopKey is a valid .ini file or
'   -   0 otherwise.
'Tong On Error GoTo fTopKeyError
fTopKey = 0
Select Case UCase$(sTopKeyOrFile)
    Case "HKCU"
        fTopKey = HKEY_CURRENT_USER
    Case "HKLM"
        fTopKey = HKEY_LOCAL_MACHINE
    Case "HKU"
        fTopKey = HKEY_USERS
    Case "HKDD"
        fTopKey = HKEY_DYN_DATA
    Case "HKCC"
        fTopKey = HKEY_CURRENT_CONFIG
    Case "HKCR"
        fTopKey = HKEY_CLASSES_ROOT
End Select
Exit Function

fTopKeyError:
End Function

Public Function fRenameRegKey(ByVal sTopKey As String, ByVal sSourcePath As String, ByVal sDestinationPath As String) As Long
    ' This is a recursive call to copy the registry key and value from
    ' sSourcePath to the sDestinationPath
    ' return 0 if success
    ' Parameter required:
    '   sTopKey = string value that indicate the top level registry key
    '             Possible value will be : "HKCU", "HKLM", "HKU", "HKDD", "HKCC", "HKCR"
    '   sSourcePath = The key path that will copy from
    '   sDestinationPath = The key path that will copy to
    'Tong On Error GoTo Err_General
    fRenameRegKey = 0
    Dim bDone    As Boolean
    Dim lTopKey  As Long
    Dim lHandle  As Long
    Dim lResult  As Long
    Dim lIndex   As Long
    Dim sKeyName As String
    Dim lCount As Long
    lCount = 0

    'Tong On Error GoTo Err_General
    lResult = 99
    lTopKey = fTopKey(sTopKey)
    If lTopKey = 0 Then GoTo Err_General
    ' Open the registry SubKey.
    lResult = RegOpenKeyEx(lTopKey, sSourcePath, 0, KEY_ENUMERATE_SUB_KEYS, lHandle)
    If lResult <> ERROR_SUCCESS Then GoTo Err_General
    
    Dim Ret
    'Create a new key
    RegCreateKey lTopKey, sDestinationPath, Ret
    fCopyRegKey sTopKey, sSourcePath, sDestinationPath, True
    ' Get the last \ and break it, for deletion purpose
    If Len(sSourcePath) > 0 Then
        Dim intOpset As Long
        intOpset = InStrRev(sSourcePath, strBackSlash, , vbTextCompare)
        If intOpset > 0 Then
            If DeleteRegistryKey(Mid$(sSourcePath, 1, intOpset - 1), Mid$(sSourcePath, intOpset + 1, Len(sSourcePath) - intOpset)) Then
                fRenameRegKey = 0
            End If
        End If
    End If
    Exit Function

Err_General:
    fRenameRegKey = -1
End Function

Public Function fCopyRegKey(ByVal sTopKey As String, ByVal sSourcePath As String, ByVal sDestinationPath As String, IsDelete As Boolean) As Long
    ' This is a recursive call to copy the registry key and value from
    ' sSourcePath to the sDestinationPath
    ' Parameter required:
    '   sTopKey = string value that indicate the top level registry key
    '             Possible value will be : "HKCU", "HKLM", "HKU", "HKDD", "HKCC", "HKCR"
    '   sSourcePath = The key path that will copy from
    '   sDestinationPath = The key path that will copy to
    '   IsDelete = TRUE for delete sub-key
    ' Example:
    '   fCopyRegKey("HKCU", "Software\3F", "Software\4F", TRUE)
    'Tong On Error GoTo Err_General
    Dim arrValues() As String
    Dim arrKeys() As String
    Dim lResult  As Long
    Dim lCount As Long
    Dim intOpset As Integer
    Dim lTopKey As Long
    Dim Ret

    ReDim arrValues(0) As String
    ReDim arrKeys(0) As String
    
    ' Transfer the destination's value 1st
    lResult = fEnumValue(sTopKey, sSourcePath, arrValues)
    If UBound(arrValues) > 0 Then
        For lCount = 0 To UBound(arrValues) - 1
            intOpset = InStr(1, arrValues(lCount), strEqual)
            If intOpset > 0 Then
                SetRegisterValue sDestinationPath, Mid$(arrValues(lCount), 1, intOpset - 1), Mid$(arrValues(lCount), intOpset + 1, Len(arrValues(lCount)))
            End If
        Next lCount
    End If
    
    lTopKey = fTopKey(sTopKey)
    If lTopKey = 0 Then GoTo Err_General

    ' Get a list of the sub-registry folder
    lResult = fEnumKey(sTopKey, sSourcePath, arrKeys)
    If UBound(arrKeys) > 0 Then
        For lCount = 0 To UBound(arrKeys) - 1
            'Create a new key
            RegCreateKey lTopKey, sDestinationPath & strBackSlash & arrKeys(lCount), Ret
            fCopyRegKey sTopKey, sSourcePath & strBackSlash & arrKeys(lCount), sDestinationPath & strBackSlash & arrKeys(lCount), IsDelete
            If IsDelete Then
                If DeleteRegistryKey(sSourcePath, arrKeys(lCount)) Then
                    fCopyRegKey = 0
                End If
            End If
        Next lCount
    End If
    Exit Function

Err_General:
    fCopyRegKey = -1
End Function

Public Function ProcessNull(ByRef psToProcess As String)
'Tong On Error GoTo Err_General
    ProcessNull = vbNullString
    ProcessNull = Replace(psToProcess, Chr(0), CONST_strCHR0)
Exit Function

Err_General:
End Function

Public Function DeprocessNull(ByRef psToProcess As String)
'Tong On Error GoTo Err_General
    DeprocessNull = vbNullString
    DeprocessNull = Replace(psToProcess, CONST_strCHR0, Chr(0))
Exit Function

Err_General:
End Function

Public Function CreateNewKey(ByVal fullPath As String)
Dim Ret As Long
    CreateNewKey = RegCreateKey(HKEY_CURRENT_USER, fullPath, Ret)
    RegCloseKey Ret
End Function

