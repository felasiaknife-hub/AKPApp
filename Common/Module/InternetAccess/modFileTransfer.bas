Attribute VB_Name = "modFileTransfer"
Option Explicit

Public boolConfirm
Public boolTransfering As Boolean
Public CheckFileStatus As Boolean
Public boolStartDownload As Boolean
Public boolFileDeleted As Boolean
Public lngFreeFileNo As Long

Public Function sendFile(localFileName As String, serverFolderName As String)
On Error GoTo err

If testConnection = False Then
    Exit Function
End If

If chkFilePath(localFileName) = False Then
    MsgBox "The file name specified by you does not exist on the local system.", vbExclamation
    Exit Function
End If
'checks if the file is already there in the server
If fileExistsOnServer(serverFolderName & "\" & GetFileName(localFileName)) = True Then
    boolConfirm = MsgBox("This file already exists on the server. Do you want to overwrite this file ?", vbYesNoCancel + vbQuestion + vbDefaultButton3)
    If boolConfirm <> vbYes Then
        Exit Function
    End If
End If

Load frmSendingFile
frmSendingFile.Show
'This is the function that sends a file
Dim strData As String
Dim BlockSize As Long
lngFreeFileNo = FreeFile
Open localFileName For Binary Access Read As lngFreeFileNo 'Open the file to send
BlockSize = 4096 'Set the block size, If needed, set it higher
boolError = False

'sends the folder name to store in server\files\
frmClient.Winsock1.SendData "#^FOLDERNAME^#" & serverFolderName
Call Wait(0.1)
If boolError = False Then
    'sends the file name to store in server\files\folder name
    frmClient.Winsock1.SendData "#^FILEBEGIN^#" & GetFileName(localFileName)
    Call Wait(0.1)
    Do While Not EOF(lngFreeFileNo) And boolError = False
        'Set aside spaces to store data inside temp
        strData = Space$(BlockSize)
    
        Get lngFreeFileNo, , strData 'Get first line from file
        frmClient.Winsock1.SendData "#^FILE^#" & strData 'Send the Real data
        Call Wait(0)
        If boolError = True Then
            GoTo err
        End If
    Loop
            
        'This tell the reciever that the data file has ended
    Call Wait(0)
    frmClient.Winsock1.SendData "#^FILEEND^#"
    Close FreeFile
Else
    MsgBox "File not saved.", vbCritical
End If

Unload frmSendingFile
Exit Function

err:
boolError = False
MsgBox "File not saved.", vbCritical
Unload frmSendingFile
Exit Function
End Function

Public Function ReceiveFile(ServerFileName As String, LocalFolder As String)
On Error GoTo err
If testConnection = False Then
    Exit Function
End If

If chkFolderPath(LocalFolder, False) = False Then
    MsgBox "The local folder does not exist on the system.", vbExclamation
    Exit Function
ElseIf chkFilePath(LocalFolder & "\" & GetVirtualFileName(ServerFileName)) = True Then
    boolConfirm = MsgBox("This file already exists on your system. Do you want to overwrite this file ?", vbYesNoCancel + vbQuestion + vbDefaultButton3)
    If boolConfirm <> vbYes Then
        Exit Function
    End If
End If

'checks if the file is already there in the server
If fileExistsOnServer(ServerFileName) = False Then
    MsgBox "This file does not exists on the server.", vbExclamation
    Exit Function
Else
    Load frmSendingFile
    frmSendingFile.Show
    
    boolStartDownload = False
    frmClient.Winsock1.SendData "#^GETFILE^#" & ServerFileName
    Call Wait(0)
    If boolStartDownload = True Then
        boolTransfering = True
        strData = ""
        lngFreeFileNo = FreeFile
        Open LocalFolder & "\" & GetVirtualFileName(ServerFileName) _
        For Binary Access Write As lngFreeFileNo
        
        While boolTransfering = True And boolError = False
            Call Wait(0.1)
        Wend
        
        Unload frmSendingFile
        
        If boolError = True Then
            GoTo err
        End If
    Else
        GoTo err
    End If
End If

MsgBox "File downloaded to : " & vbCrLf & vbCrLf & LocalFolder & "\" & GetVirtualFileName(ServerFileName), vbInformation

Exit Function

err:
boolError = False
MsgBox "File not saved.", vbCritical
Unload frmSendingFile
Exit Function

End Function

Public Function DeleteFile(ServerFileName As String)
'checks if the file is already there in the server
If fileExistsOnServer(ServerFileName) = False Then
    boolConfirm = MsgBox("This file does not exist on the server.", vbExclamation)
    Exit Function
Else
    boolFileDeleted = False
    frmClient.Winsock1.SendData "#^DELETEFILE^#" & ServerFileName
    Call Wait(0)
    If boolFileDeleted = True Then
        MsgBox "File Deleted.", vbInformation
    Else
        MsgBox "File Not Deleted.", vbExclamation
    End If
End If

End Function
