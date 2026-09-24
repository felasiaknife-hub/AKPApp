VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim MailMsg, MailConf, MCFields
Const cdoBasic = 1

Public Sub sendEMail(EMailIDs As String, LocalIP As String)
If Len(frmServer.txtEMail.Text) > 0 Then
    Call AddSeverMsg("Sending startup mail...")
End If

Set MailMsg = CreateObject("CDO.Message")
Set MailConf = CreateObject("CDO.Configuration")
Set MCFields = MailConf.Fields

'Configure SMTP Mail
MCFields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2  '  SMTP
MCFields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "" 'put SMTP Server here
MCFields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "********" 'put Pwd here
MCFields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "" 'put User Name here
MCFields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = cdoBasic
MCFields.Update

Set MailMsg.Configuration = MailConf
'MailMsg.HTMLBody = txtMessage.Text
MailMsg.TextBody = LocalIP & vbCrLf & vbCrLf & vbCrLf & "---------" & vbCrLf _
    & "This is an automated mail. Please do not reply to this mail." & vbCrLf & vbCrLf _
    & "This e-mail" _
    & "correspondence is confidential and intended for the named recipient(s) only." _
    & "If you are not the named recipient and received this correspondence in error," _
    & "you must not copy, distribute or take any action in reliance on it."
MailMsg.Subject = "DGAC Server IP"
MailMsg.To = EMailIDs  ' "ABCD <abcd@abcd.com>"
MailMsg.From = "" ' "ABCD <abcd@abcd.com>
'MailMsg.AddAttachment "C:\Documents and Settings\Desktop\Test.txt"

On Error GoTo err
MailMsg.Send
If Len(frmServer.txtEMail.Text) > 0 Then
    Call AddSeverMsg("The startup mail has been sent successfully")
End If
Exit Sub

err:
If Len(frmServer.txtEMail.Text) > 0 Then
    Call AddSeverMsg(err.Description)
    Call AddSeverMsg("ERROR: The startup mail could not be sent successfully")
End If
End Sub

