## Password expiration email notification

Users at your company keep ignoring, missing or "not getting" the Windows notification about their password expiring?

With this script you can spam their mailbox and reduce the chances of them not changing their password in time by 2%. Because many will still not change it in time...

The script check for two type of user in AD, one in the remote group meaning someone working from home not connected to the network physically or over a vpn and that could only update password from Office365 portal, and the other for those who are working from an office or over a vpn being able to update their password from windows.

The script will send a different email for each of the two cases.

You can remove your email from the bcc field to not get each of the emails but I had it for testing.


```
Import-Module ActiveDirectory

$maxPasswordAgeTimeSpan = (Get-ADDefaultDomainPasswordPolicy).maxPasswordAge

Get-ADUser -Filter {Enabled -eq $true -and PasswordNeverExpires -eq $false} -Properties PasswordLastSet, PasswordExpired, PasswordNeverExpires, EmailAddress, Name, Office, MemberOf |

ForEach-Object {

$today=get-date
$UserName=$_.Name
$Email=$_.EmailAddress
$NotifyAt = 14,7,3

        if (!$_.PasswordExpired -and !$_.PasswordNeverExpires) {

        $ExpiryDate=$_.PasswordLastSet + $maxPasswordAgeTimeSpan
        $DaysLeft=($ExpiryDate-$today).days
            
        if($_.Office -match 'Remote' -or
           $_.MemberOf -contains 'CN=All Remote Employees, OU=Groups, DC=yourdomain, DC=com'){
           
           $body = "
            <p style='font-family:Lucida'>Dear $UserName,</p>
            <p style='font-family:Lucida'>Your Company login password will expire in $DaysLeft.</p>

            <p style='font-family:Lucida'>Please follow the next steps to change your password.</p>
                <ul style='font-family:Lucida'>
		    <li>Login into <a href='https://portal.office.com/'>Office365</a></li>
                    <li>In the top right corner, click your picture</li>
                    <li>Select the 'Change Password' option</li>
                    <li>Type your current password</li>
                    <li>Type the new password</li>
                    <li>Re-enter the new password and click 'Submit'</li>
                </ul>

            <p style='font-family:Lucida'>Requirements for the password are as follows:</p>
                <ul style='font-family:Lucida'>
		    <li>Must contain at least 8 characters</li>
                    <li>Contain characters from three of the following four categories:</li>
                    <li>English uppercase characters (A through Z)</li>
                    <li>English lowercase characters (a through z)</li>
                    <li>Base 10 digits (0 through 9)</li>
                    <li>Non-alphabetic characters (for example, !, $, #, %)</li>
                </ul>
            <p style='font-family:Lucida'>For any assistance, please submit a <a href='https://aservicedeskportal.ifexisting'>Service Desk Ticket</a></p>

			<p style='font-family:Lucida'><b>Company Technology Team</b></p>
            "          
           }
        else{
             $body = "
            <p style='font-family:Lucida'>Dear $UserName,</p>
            <p style='font-family:Lucida'>Your Windows login password will expire in $DaysLeft days, please press <b>CTRL-ALT-DEL</b> and change your password.</p>

            <p style='font-family:Lucida'>Requirements for the password are as follows:</p>
                <ul style='font-family:Lucida'>
		    <li>Must contain at least 8 characters</li>
                    <li>Contain characters from three of the following four categories:</li>
                    <li>English uppercase characters (A through Z)</li>
                    <li>English lowercase characters (a through z)</li>
                    <li>Base 10 digits (0 through 9)</li>
                    <li>Non-alphabetic characters (for example, !, $, #, %)</li>
                </ul>
            <p style='font-family:Lucida'>For any assistance, please submit a <a href='https://aservicedeskportal.ifexisting'>Service Desk Ticket</a></p>

			<p style='font-family:Lucida'><b>Company Technology Team</b></p>
            "          
            }

            if ($DaysLeft -in $NotifyAt){
                ForEach($email in $_.EmailAddress) {
                send-mailmessage -From noreply@yourdomain.com -To $Email -Bcc youremail@yourdomain.com -Subject "Password Reminder: Your password will expire in $DaysLeft days" -Body $body  -SmtpServer yourdomain.smtp.server -Port 25 -BodyAsHtml }
            
            }

        }
    }
```
