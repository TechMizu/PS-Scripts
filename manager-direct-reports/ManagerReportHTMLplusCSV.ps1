<#
.NOTES

    NAME:      Manager DirectReport Email.ps1

    AUTHOR:    Tech Mizu | https://github.com/TechMizu

    CREATED:   04/06/2017
        
.DESCRIPTION

    Sends an email to each mananager with the list of users who directly reports to them. The email contains html in the body to show
    the information: Name, Title, Group in AD. It also sends as attachment a csv file with the same information.
  
#>

Import-Module ActiveDirectory

# Head
$a = "<style>"
$a = $a + "BODY{font-family:Arial;background-color:#fff;width: 100%}"
$a = $a + "TABLE{font-family:Arial;width: 90%; border-collapse: collapse;background-color:#f2f2f2;}"
$a = $a + "TH{background-color: #F48642; color: white;height: 35px;}"
$a = $a + "TD{display: table-cell;height: 25px;vertical-align: inherit;border-bottom: 1px solid #ddd;}"
$a = $a + "P{font-family:Arial;position: absolute;}"
$a = $a + "</style>"

$managers = Get-ADUser -Filter * -Properties Name, DirectReports, EmailAddress | where {$_.directreports -ne $Null}

foreach ($manager in $managers) 
{
    $managername = $manager.Name
    $manageremail = $manager.UserPrincipalName
    $dreports = $manager.directreports
 
    $today = (Get-Date).ToString()
	
	#start the csv code
	
	foreach($dr in $dreports){
		$hash = @{ 
            Manager         = $manager.Name
            DirectReport    = $dr -replace '^CN=(.+?),(?:OU|CN)=.+','$1'
            GroupMembership = (@((get-aduser $dr -Properties memberof).memberof -replace '^CN=(.+?),(?:OU|CN)=.+','$1') -join ' ;; ')
        }
        
        $Object = New-Object PSObject -Property $hash
        $out    = $Object | select Manager,DirectReport,GroupMembership 
		
        #set variable for the csv file to be used within the script
	    $csvfile = "C:\$($manager.samaccountname).csv"
        $out | Export-Csv $csvfile -NoTypeInformation -Force -Append
	}
	
	
	#start the email body code

    $body = "<center><p>Report Date $today.</p></center>"
    $body += "<p>Dear <b>$managername</b>,</p> "
    $body += "<p>Please find below the current list of users who directly report to you and each of the groups in Active Directory they are member of.</p>" 

    $body += "<p>You can also find attached the same list of users reporting to you and their access separated by two semicolons ' ;; '. If any change is required, please submit a case to the <a href='https://aservicedeskportal.ifexisting'>service desk</a>.</p>" 

    foreach ($dr in $dreports)
    {
        $user = get-aduser $dr -properties *
        $members = $user.memberof -replace '^CN=(.+?),(?:OU|CN)=.+','$1' | %{New-Object PSObject -Property @{'Group Membership'=$_}} | convertTo-html -Head $a -Body "<H4>Group Membership.</H4>" 
        $dreport = $dr -replace '^CN=(.+?),(?:OU|CN)=.+','$1'
        $desc = $user.description

        $body += "`n"
        $body += "<br />"
        $body += "<center>"
        $body += "<p><b>Direct Report:</b> $dreport</p>"
        $body += "<p><b>Title:</b> $desc</p>"
        $body += $members
        $body += "</center>"
     }
        $body += "`n"
        $body += "<br />"
        $body += "<center><h3>Company Technology Team.</h3></center>" 

           #html temp file path  
           $htmltemp = "c:\mytempfile.htm" 
            
           #output to the filepath  
           $body | out-file $htmltemp

           #get the contents of the file
           $htmlcontents = get-content ("$htmltemp")

     
      #send email using $htmlcontents as the email body and attaching $csvfile
      Send-MailMessage -From noreply@yourdomain.com -To youremail@yourdomain.com <#$manager.UserPrincipalName#> -Subject "Managers - User Entitlements" -Attachments $csvfile -Body ("$htmlcontents") -SmtpServer yourdomain.SmtpServer -Port 25 -BodyAsHtml
      
      #delete temp file
      remove-item $htmltemp -force | out-null

      #delete the csv file from c:\
      remove-item $csvfile -force | out-null
}