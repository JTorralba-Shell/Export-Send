Clear-Host

$WaitTime = 5

Write-Host "Exporting Report"
Write-Host

$LOG = $PSScriptRoot + "\PowerShell.log"
$BODYFILE = $PSScriptRoot + "\Body.txt"

$PDFFile = $PSScriptRoot + "\Impound Logs.pdf"

Start-Sleep -s $WaitTime

Write-Host "Preparing PDF Attachement"
Write-Host

$TimeStamp = (Get-Item $PDFFile).LastWriteTime.toString("yyyy-MM-dd HHmmss")

$PDFAttachment = "Impound Logs " + $TimeStamp + ".pdf"
Write-Output $PDFAttachment > $LOG

Rename-Item $PDFFile $PDFAttachment

Start-Sleep -s $WaitTime

Write-Host "Preparing E-Mail"
Write-Host

########## Variables ##########

$FROM = "FirstLast@Domain.com"

$TO = "FirstLast@Domain.com"

$CC = "FirstLast@Domain.com"
$CC2 = "FirstLast@Domain.com"
$CC3 = "FirstLast@Domain.com"

$BCC = "FirstLast@Domain.com"
$BCC2 = "FirstLast@Domain.com"
$BCC3 = "FirstLast@Domain.com"

$SUBJECT = $PDFAttachment.Replace(".pdf","")
$BODY = get-content $BODYFILE
$FILE = $PSScriptRoot + "\" + $PDFAttachment
Write-Output $File >> $LOG

$SERVER = "SMTP.SocketLabs.com"

########## Build & Send Message ##########

$Message = New-Object System.Net.Mail.MailMessage

$Message.From = $FROM

$Message.To.Add($TO)

$Message.CC.Add($CC)
$Message.CC.Add($CC2)
$Message.CC.Add($CC3)

$Message.BCC.Add($BCC)
$Message.BCC.Add($BCC2)
#$Message.BCC.Add($BCC3)

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

Write-Host "Cleaning Up"
Write-Host

If (Test-Path $File)
{
    Remove-Item $File -Force
}

Start-Sleep -s $WaitTime
