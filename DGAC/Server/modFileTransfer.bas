Attribute VB_Name = "modFileTransfer"
Option Explicit

Public CheckFileStatus As Boolean

Public Function sendData(FileName As String, Index As Integer)
'This is the function that sends a file
Dim strData As String
Dim BlockSize As Long
Open FileName For Binary Access Read As FreeFile 'Open the file to send
BlockSize = 4096 'Set the block size, If needed, set it higher

frmServer.sckServer(Index).sendData "#^FILEBEGIN^#"
Do While Not EOF(FreeFile)
    'Set aside spaces to store data inside temp
    strData = Space$(BlockSize)

    Get FreeFile, , strData 'Get first line from file
    frmServer.sckServer(Index).sendData strData 'Send the Real data
Loop
        
    'This tell the reciever that the data file has ended
frmServer.sckServer(Index).sendData "#^FILEEND^#"
Close FreeFile
End Function

