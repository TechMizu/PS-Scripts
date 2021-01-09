## Check and notify through email for local expiering certificates

This script will look for certificates matching the given thumbprints, check and send an email if is expiering within the next 30 days.

This was usefull to me after a start-up company had an issue where one of the early sysadmins left months prior and no one knew about this certificate expiration.

In my case found two undocumented certificates and created this script to alert if they were expiring.

I have not had need for this ever since, but this can be helpful to someone.

This can be run daily from a task schedule.

```
#Number of days for certificate expiration threshold
$Threshold = 30
 
#Deadline in comparison with current date
$Deadline = (Get-Date).AddDays($Threshold)
 
#Email parameters
$SMTPServer = "your.smtp.server"
$From = "noreply@yourdomain.com"
$To = "email@yourdomain.com"
$Subject = "Certificate expiration check performed on $(Get-Date -Format "MM.dd.yyyy")"

# Head for html body
$a = "<style>"
$a = $a + "BODY{font-family:Arial;background-color:#fff;width: 100%}"
$a = $a + "TABLE{font-family:Arial;width: 100%; border-collapse: collapse;background-color:#f2f2f2;}"
$a = $a + "TH{background-color: #F48642; color: white;height: 35px;}"
$a = $a + "TD{display: table-cell;height: 25px;vertical-align: inherit;border-bottom: 1px solid #ddd;}"
$a = $a + "P{font-family:Arial;position: absolute;}"
$a = $a + "</style>"

#Create temp csv file
$Date = Get-Date -Format "yyMMdd"
$ReportPath = Join-Path $env:TEMP "CertCheck_$Date.CSV"

#Specific certificate thumbprint to look for
$tmbprt = "a909502dd82ae41433e6f83886b00d4277a32a7b", "886b00d4277a32a7ba909502dd82ae41433e6f83"

#Get all certificates from all stores
$Certificates = Get-ChildItem Cert: -Recurse | Where-Object {$_.Thumbprint -in $tmbprt -and $_.Subject -ne $null}
 
#List all certificates within defined treashold creating PS objects
$Report =@()
ForEach ($Certificate in $Certificates) {
 
    If ($Certificate.NotAfter -le $Deadline) {
            $Report += New-Object PSObject -Property @{
                CertificateSubject = $Certificate.Subject
                Thumbprint = $Certificate.Thumbprint
                ExpiresAfter = $Certificate.NotAfter
                ExpiresIn = ($Certificate.NotAfter - (Get-Date)).Days
            }
    }
}
 
#region If any certificates corresponds to treshold criteria create and email report...
 
If (($Report | Measure-Object).Count -gt 0) {
        $Report | Select-Object CertificateSubject, Thumbprint, ExpiresAfter, ExpiresIn | Sort ExpiresAfter | Export-CSV -Path $ReportPath -NoTypeInformation
        $out = $Report | ConvertTo-Html -Head $a -Body "<H4>Expiring Certificates</H4>" 
        $SMTPMessage = @{
            To = $To
            From = $From
            Subject = $Subject
            SmtpServer = $SMTPServer
            Attachments = $ReportPath
        }
    
    $Body = @()
    $body += "`n"   
    $body += "Hello, <br />"
    $body += ""
    $body += "Please find attached report CertCheck_$($Date).CSV which lists certificates expiring in $Threshold days. <br />"
    $body += ""
    $body += "<center>"
    $body += $out
    $body += "</center>"
    $body += "<br />"
    $body += "Regards, <br />"
    $body += "Company. | Technology Team"
    $body += "`n"
    $Body = $body | Out-String

               #html temp file path  
               $htmltemp = Join-Path $env:TEMP "mytempfile.htm" 
            
               #output to the filepath  
               $body | out-file $htmltemp
    
               #get the contents of the file
               $htmlcontents = get-content ("$htmltemp")

    Send-MailMessage @SMTPMessage -Body ("$htmlcontents") -Priority High -BodyAsHtml
}
 
#endregion
 
#region   System clenup...
 
Try {
    Remove-Item $ReportPath, $htmltemp -ErrorAction Stop
}
Catch {
    Write-Host ""
    Write-Host "No certificates will expire in defined threshold." -ForegroundColor Cyan
}
 
Remove-Variable Treshold, Deadline, Certificates, Report, Date -ErrorAction SilentlyContinue
Remove-Variable SMTPServer, From, To, Subject, Body -ErrorAction SilentlyContinue
 
#endregion

```
