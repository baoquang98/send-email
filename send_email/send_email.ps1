#Usage: Right click -> Run with PowerShell
echo "Welcome to send email script, to exit the script, press Ctrl+C"
$EmailFrom = Read-Host "Enter your email address (not outlook)"
$List_file = Read-Host "Email list file name"
$Emaillist = Get-content $List_file
$Subject = Read-Host "Email Subject"
#$att = new-object Net.Mail.Attachment("kitty.jpg")
#$att.ContentId = "att"
$template_file = Read-Host "Email template file name"
$Body_raw = Get-Content $template_file | out-string
$SMTPServer = “smtp.gmail.com”
$SMTPClient = New-Object Net.Mail.SmtpClient($SmtpServer, 587)
$SMTPClient.EnableSsl = $true
$SecurePassword = Read-Host -Prompt "Enter password" -AsSecureString
$BSTR = `
    [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecurePassword)
$Password = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
$SMTPClient.Credentials = New-Object System.Net.NetworkCredential($EmailFrom, $Password);
foreach ($Recipent in $Emaillist)
{
    $EmailTo = $Recipent.split("`t")
    $Body = $Body_raw.replace("[FIRSTNAME]", $EmailTo[0])
    $message = New-Object Net.Mail.MailMessage($EmailFrom, $EmailTo[1], $Subject, $Body)
    $message.IsBodyHtml = $true;
    #$Message.Attachments.Add($att)
    $SMTPClient.Send($message)
    echo $EmailTo[1]
}
#$att.Dispose()
echo "Have a nice day"
Read-Host "Press Enter to continue"