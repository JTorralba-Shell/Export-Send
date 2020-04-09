Clear-Host

$WaitTime = 5

Write-Host "Preparing E-Mail"
Write-Host

########## Variables ##########

$FROM = "FirstLast@Domain.com"
$TO = "FirstLast@Domain.com"
$CC = "FirstLast@Domain.com"
$BCC = "FirstLast@Domain.com"

$SUBJECT = "Hello World from AUTOMATE server."
$BODY = "This is an automated e-mail test (no response required at this time)."
$BODY = get-content .\Body.txt
$FILE = "Body.txt"

$SERVER = "SMTP.SocketLabs.com"

########## Build & Send Message ##########

$Message = New-Object System.Net.Mail.MailMessage

$Message.From = $FROM
$Message.To.Add($TO)
$Message.CC.Add($CC)
$Message.BCC.Add($BCC)

$Message.Subject = $SUBJECT
$Message.IsBodyHtml = $False
$Message.Body = $BODY

$Attachment = New-Object System.Net.Mail.Attachment($File)
$Message.Attachments.Add($Attachment)

$SMTP = New-Object Net.Mail.SmtpClient($SERVER, 587)
$SMTP.EnableSsl = $False
$SMTP.Credentials = New-Object System.Net.NetworkCredential("username","password")

Write-Host "Sending E-Mail"
Write-Host

$SMTP.Send($Message)

$Message.Dispose()
$SMTP.Close

Start-Sleep -s $WaitTime

Read-Host -Prompt "Press any key to continue."

