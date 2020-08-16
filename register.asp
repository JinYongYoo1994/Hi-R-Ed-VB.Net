<%
Set myMail=CreateObject("CDO.Message")

dim smtpServer, yourEmail, yourPassword
smtpServer = "smtp.gmail.com"
yourEmail = "luckytaurus1225@gmail.com"     'replace with a valid gmail account
yourPassword = "Fantasylab123!@#!@#"   'replace with a valid password for account set in yourEmail 

'E-mail subject:
myMail.Subject="Sending email with CDO"

'The from address:
myMail.From="luckytaurus1225@gmail.com"

'The to address:
myMail.To="liangmi425@gmail.com"

'Text:
myMail.TextBody="Your text goes here when sending e-mails"

myMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
myMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = smtpServer
myMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
myMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
myMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = true
myMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = yourEmail
myMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = yourPassword
myMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60 
myMail.Configuration.Fields.Update

'Send the E-mail:

On Error Resume Next
myMail.Send

If Err.Number <> 0 Then
    MsgBox Err.Description,16,"Error Sending Mail"
    Response.Write(Err.Description)
    Response.Write(Err.Source)
Else 
    MsgBox "Mail was successfully sent !",64,"Information"
End If

set myMail=nothing
%>