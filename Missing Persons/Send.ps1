Clear-Host

$WaitTime = 5

Write-Host "Exporting Report"
Write-Host

$LOG = $PSScriptRoot + "\PowerShell.log"
$BODYFILE = $PSScriptRoot + "\Body.txt"

$PDFFile = $PSScriptRoot + "\Missing Persons.pdf"

Start-Sleep -s $WaitTime

Write-Host "Preparing PDF Attachement"
Write-Host

$TimeStamp = (Get-Item $PDFFile).LastWriteTime.toString("yyyy-MM-dd HHmmss")

$PDFAttachment = "Missing Persons " + $TimeStamp + ".pdf"
Write-Output $PDFAttachment > $LOG

Rename-Item $PDFFile $PDFAttachment

Start-Sleep -s $WaitTime

Write-Host "Preparing E-Mail"
Write-Host

########## Variables ##########

$FROM = "AUTOMATE01@EPTPaging.Info"

$TO = "SHRDetectivesReport@ElPasoCo.com"

$CC = "KylaGingrich@ElPasoCo.com"
$CC2 = "MeighanPowell@ElPasoCo.com"
$CC3 = "LeahStevens@ElPasoCo.com"

$BCC = "CAD@EPTC911.org"

$SUBJECT = $PDFAttachment.Replace(".pdf","")
$BODY = get-content $BODYFILE
$FILE = $PSScriptRoot + "\" + $PDFAttachment
Write-Output $File >> $LOG

$SERVER = "Mail.EPTPaging.Info"

########## Build & Send Message ##########

$Message = New-Object System.Net.Mail.MailMessage

$Message.From = $FROM

$Message.To.Add($TO)

$Message.CC.Add($CC)
$Message.CC.Add($CC2)
$Message.CC.Add($CC3)

$Message.BCC.Add($BCC)

$Message.Subject = $SUBJECT
$Message.IsBodyHtml = $False
$Message.Body = $BODY

$Attachment = New-Object System.Net.Mail.Attachment($File)
$Message.Attachments.Add($Attachment)

$SMTP = New-Object Net.Mail.SmtpClient($SERVER, 25)
$SMTP.EnableSsl = $False
# $SMTP.Credentials = New-Object System.Net.NetworkCredential("username","password")

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
