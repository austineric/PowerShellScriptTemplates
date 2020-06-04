

####################################
# Author:       Eric Austin
# Create date:  
# Description:  Uses MailKit to send emails
#               Microsoft's System.Net.Mail.SmtpClient is marked obsolete; it still works but it's probably better to use MailKit (the recommended alternative)
#               The best way to get the latest version for each assembly is to download the package from NuGet and unzip the .nupkg using 7Zip
####################################

using namespace MimeKit
using namespace MailKit.Net

#load assemblies
#updated versions should probably be used
#have to be loaded in this specific order as far as I can tell
#I don't think I can take just the DLL's since the BouncyCastle DLL's are named the same, I just grabbed the whole folders
Add-Type -Path ".\MailKit\bouncycastle.1.8.6.1\lib\BouncyCastle.Crypto.dll"
Add-Type -Path ".\MailKit\portable.bouncycastle.1.8.6.7\lib\netstandard2.0\BouncyCastle.Crypto.dll"
Add-Type -Path ".\MailKit\mimekit.2.8.0\lib\netstandard2.0\MimeKit.dll"
Add-Type -Path ".\MailKit\mailkit.2.7.0\lib\netstandard2.0\MailKit.dll"

#check if assembly is loaded
#[System.AppDomain]::CurrentDomain.GetAssemblies() | Where-Object -Property Location -Like "*Mail*"

#message
$Message=New-Object MimeMessage

#from
$From=New-Object MailboxAddress("Sender name", "SenderEmailAddress")
$Message.From.Add($From)

#to (single recipient)
$To=New-Object MailboxAddress("Recipient name", "RecipientEmailAddress")
$Message.To.Add($To)

#to (multiple recipients)
$To=New-Object InternetAddressList
$To.Add("RecipientEmailAddres")
$To.Add("RecipientEmailAddres")
$Message.To.AddRange($To)

#subject
$Message.Subject="Email subject"

#body (html)
$BodyBuilder=New-Object BodyBuilder
$BodyBuilder.HtmlBody=
@"
<div style="background-color:green; font-size:larger; font-weight:bold;">
HTML email with bold and background color
<br />
Line 2
</div>
"@
$Message.Body=$BodyBuilder.ToMessageBody()

#smtp send
#25 is the port number in this example
$Client=New-Object MailKit.Net.Smtp.SmtpClient
$Client.Connect("SMTP server name or IP address", 25)
$Client.Send($Message)
$Client.Disconnect($true)

