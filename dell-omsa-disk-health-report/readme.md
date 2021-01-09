## Dell OpenManage Server Admnistrator Disk Health Report

I think now days and with the latest version of Dell OpenManage this script is completely irrelevant but maybe somone running and old server with an old version of it may still find this useful.

Old version of OpenManage tho it had a build in notification system for disk health issues I believe it use to simply run once when the event was trigger and if you catch it great if not, well... You will soon one way or another but hopefully before more than one disk goes bad.

I found part of the script on getting the information from powershell of the disk health and I made it into a email with the report attachment and some other custom information to meet my needs at the time, so thank you to that person who already worked on something similar for this and my apologize for not remembering your blog.

The script runs from a task schedule.

Check for the health of the disks, extract the information into a csv file and then sends it attached to an email to the specified mail such as the IT team.

```
#Email parameters
$SMTPServer = "your.smtp.server"
$From = "noreply@yourdomain.com"
$To = "inf-alerts@yourdomain.com"
$Subject = "Physical Disk Weekly report for $(Get-Date -Format "MM.dd.yyyy") on $env:COMPUTERNAME"

# Head for html body
$a = "<style>"
$a = $a + "BODY{font-family:Arial;background-color:#fff;width: 100%}"
$a = $a + "TABLE{font-family:Arial;width: 100%; border-collapse: collapse;background-color:#f2f2f2;}"
$a = $a + "TH{background-color: #F48642; color: white;height: 35px;text-align:left;}"
$a = $a + "TD{display: table-cell;height: 25px;vertical-align: inherit;border-bottom: 1px solid #ddd;}"
$a = $a + "P{font-family:Arial;position: absolute;}"
$a = $a + "</style>"

$ReportPath = Join-Path $env:TEMP "DiskCheck_$Date.CSV"

# Store the output of the omreport.exe command
$omsa = & "C:\Program Files\Dell\SysMgt\oma\bin\omreport.exe" storage pdisk controller=0

# Count the number of disks
$diskCount = $omsa | where {$_ -like 'ID*'} | measure | select -ExpandProperty Count

# Count the number of lines in the first disk
$lineCount=0 
for($i=3;$i -lt $omsa.Count; $i++){
if($omsa[$i] -ne ''){
$lineCount++
}
else{
break
}
}
 
$lineStart = 3 # First line of first disks information
$lineEnd = $lineCount + 2 # Last line of first disks information
# Empty array to store all disks in
$disks = New-Object System.Collections.ArrayList
for($i=0;$i -lt $diskCount;$i++){
# Empty object to store disk info
$disk = new-object -TypeName psobject
# Replace colons with dashes
foreach($j in $omsa[$lineStart..$lineEnd]){
if($j -match "0:0:$i"){
$j = $j -replace "0:0:$i","0-0-$i"
}
$name = $j.split(':')[0].trim().replace(' ','')
$value = $j.split(':')[1].trim()
$disk | Add-Member -MemberType NoteProperty -Name $name -Value $value
}
# Split lines into Key/Value pairs, remove excess spaces
[void]$disks.Add($disk)
# Increment start and finish lines for next disk
$lineStart = $lineEnd + 2
$lineEnd = $lineStart + ($lineCount - 1)
}
$disks | Where-Object -Property State -ne Other |  Export-CSV -Path $ReportPath -NoTypeInformation
$out = $disks | Select-Object -Property ID,Status,Media,@{Name="Capacity";Expression = {$_.Capacity.Substring(0,11)}},FailurePredicted,VendorID,Certified,SerialNo.|ConvertTo-Html -Head $a -Body "<H4>Physical Disk Report for $env:COMPUTERNAME</H4>" 
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
$body += "Please find attached the complete report DiskCheck_$($Date).CSV which lists information about the physical disks on $env:COMPUTERNAME server. <br />"
$body += ""
$body += "<center>"
$body += $out
$body += "</center>"
$body += "<br />"
$body += "Regards, <br />"
$body += "Company | Technology Team"
$body += "`n"
$Body = $body | Out-String

       #html temp file path  
       $htmltemp = Join-Path $env:TEMP "mytempfile.htm" 
    
       #output to the filepath  
       $body | out-file $htmltemp

       #get the contents of the file
       $htmlcontents = get-content ("$htmltemp")

Send-MailMessage @SMTPMessage -Body ("$htmlcontents") -Priority High -BodyAsHtml

 
Try {
    Remove-Item $ReportPath, $htmltemp -ErrorAction Stop
}
Catch {
    Write-Host ""
    Write-Host "Report." -ForegroundColor Cyan
}
 
Remove-Variable  omsa, discount, disks, Date -ErrorAction SilentlyContinue
Remove-Variable SMTPServer, From, To, Subject, Body -ErrorAction SilentlyContinue
 
#endregion
```
