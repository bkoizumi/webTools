$mail = @{
  from = "ã∆ê—ä«óùÉVÉXÉeÉÄ<Koizumi.Bunpei@trans-cosmos.co.jp>";
  to = "Koizumi.Bunpei@trans-cosmos.co.jp";
  smtp_server = "smdp.trans-cosmos.co.jp";
  smtp_port = "587";
  user = "a2015135@trans-cosmos.co.jp";
  password = "4rfv%TGB6yhn";
}

$mailtoArray=@(
 "Koizumi.Bunpei@trans-cosmos.co.jp"
 )
 $mailccArray=@(
  "Hanamura.Hideaki@trans-cosmos.co.jp",
  "Akai.Hiroaki@trans-cosmos.co.jp"
 )



$FromFilePath = $Args[0]
$ToFilePath =   $Args[0] + "_cov"
Get-Content  $FromFilePath | Set-Content -Encoding Unicode $ToFilePath

$lines = Get-Content $ToFilePath
foreach ($line in $lines) {
$MailBody = $MailBody + $line + "`n"
}

# $today = Get-Date -Format d
$subject = $Args[1]

$password = ConvertTo-SecureString $mail["password"] -AsPlainText -Force
$credential = New-Object System.Management.Automation.PSCredential $mail["user"], $password

Send-MailMessage -To $mailtoArray `
                 -Cc $mailccArray   `
                 -From $mail["from"] `
                 -SmtpServer $mail["smtp_server"] `
                 -Credential $credential `
                 -Port $mail["smtp_port"] `
                 -Subject $subject `
                 -Body $MailBody `
                 -Encoding ([System.Text.Encoding]::UTF8) `
                 -UseSsl
