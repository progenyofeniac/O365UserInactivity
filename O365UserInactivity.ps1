<# 
 Gathers a list of O365 email accounts which haven't been logged into for so many days.

 Written by ProgenyofEniac
 10/5/2017
 12/03/2017 revised to include Department. Added "get-msoluser" to gather departments
            Also added additional loop to add dept to the list
 7/2/2018   MFA broke this script since it was using an admin account. Set up dedicated
            report@domain.org account with complex password, gave it read-only/reporting permissions
            Updated script to use this user & password, tested OK
 7/11/2018  Revised the script to use the audit log rather than the 'LastLogonTime' listed by 'Get-MailboxStatistics'
            That value appears to be updated or refreshed by any mailbox activity, including server-side, such as receiving an email.
            The script now checks the most recent audit log activity, excluding certain measures, and uses that to judge most
            recent account activity. This requred adding the 'view-only audit log' permission to the read-only account used for
            running this report.
            This also required rewriting portions of the script to use more complicated data gathered from the audit log.
            Comments have been added throughout to clarify what's going on.
 
#>

# Set variables
$LoginDays = 30 # Accounts older than this many days will be reported as inactive
$OutFile = "C:\temp\InactiveAccounts.csv"
$SortField = "LastLogon" # Change to either "Username", "Email Address", or "Last Logged In", as desired
$From = "EmailReport@domain.org" # Who email should appear to be from
$Recipients = @("user1@domain.org","user2@domain.org") # Email recipients list, each in parenthesis, comma-separated
$O365user = "report@domain.org"
$O365pwfile = "C:\Users\adminuser\Documents\PS_Scripts\o365cred.txt"
$EndDate = (Get-Date) # Used for the 'end date' for the audit log
$StartDate = $EndDate.AddDays(- 85) # Used as the start date for the audit log. Leaving it at 90 days gave some false positives for some reason.

# Grab password from file, use an admin account to connect to O365
$password = Get-Content $O365pwfile | ConvertTo-SecureString
$O365cred = New-Object System.Management.Automation.PSCredential $O365user,$password
$O365Session = New-PSSession –ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell -Credential $O365Cred -Authentication Basic -AllowRedirection
Import-PSSession $O365Session
Connect-MsolService –Credential $O365Cred


# Calculate cutoff date for last login
$Cutoff = (Get-Date).AddDays(- $LoginDays)

# Set up output lists as array
$Output1 = @()
$Output2 = @()
$AllMailboxesStats = @()

# Gather mailbox data, filtering only user mailboxes created prior to the cutoff date
$AllMailboxes = (Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox | ?{$_.WhenMailboxCreated -lt $Cutoff})

# Gather the most recent audit log entry for each mailbox, excepting "UserLoginFailed" and "PasswordLogonInitialAuthUsingPassword"
# Those two entries don't indicate a successful login and so are not included in calculating most recent activity
# 
foreach ($SingleMailbox in $AllMailboxes) {
    $MailboxStats = Search-UnifiedAuditLog -enddate $EndDate -startdate $StartDate -UserIds $SingleMailbox.WindowsLiveID | where {$_.operations -ne "UserLoginFailed" -and $_.operations -ne "PasswordLogonInitialAuthUsingPassword"} | select UserIDs,CreationDate -First 1
    
    # In case the audit log does not show any data for a given mailbox, still add an empty entry to the $AllMailboxesStats variable
    if ($MailboxStats -eq $Null)
         {$AllMailboxesStats += $SingleMailbox | select @{Label="EmailAddress";Expression={$SingleMailbox.WindowsLiveID}},@{Label="LastLogonTime";Expression={$null}}}
    
    # Translate the audit log field names to more recognizable headers
    else {$AllMailboxesStats += $MailboxStats | select @{Label="EmailAddress";Expression={$mailboxstats.UserIds}},@{Label="LastLogonTime";Expression={$mailboxstats.CreationDate}}}
    }


# Gather Userlist which includes Departments
# Go figure, neither of the other two commands includes Department...
$DeptList = get-msoluser -All | Select UserPrincipalName,Department

# Gather list of which mailboxes haven't been  logged into for $LoginDays days or never
$InactiveMailboxes = ($AllMailboxesStats | ?{($_.LastLogonTime -lt $Cutoff) -or ($_.LastLogonTime -eq $Null)})

# Loop through list of unused mailboxes, convert last logon to string rather than long date
# Convert null values (never logged in) to the date the mailbox was created (if created less than 6 months ago)
# Or state that it's been more than 6 months since login
foreach ($Mailbox in $InactiveMailboxes) {
    $User = ($AllMailboxes | ?{$_.WindowsEmailAddress -eq $Mailbox.EmailAddress})
    if ($Mailbox.LastLogonTime -eq $Null)
        {
         if ($User.WhenMailboxCreated -gt $StartDate)
             {$LastLogon = ($User.WhenMailboxCreated).ToString("yyyy-MM-dd") + " Never"}
        else {$LastLogon = "Over 3 Months"}
        }
    else {$LastLogon = ($Mailbox.LastLogonTime).ToString("yyyy-MM-dd")}

# Compile an organized list from each item, including Name, Email, Last logon

$SingleUser = ($User | Select DisplayName,@{Label="EmailAddress";Expression={$User.PrimarySmtpAddress}},@{Label="LastLogon";Expression={$LastLogon}})
$Output1 += $SingleUser
}

# Add Departments to the list, matching users by email address
foreach ($Item in $Output1) {
    $TempUser = ($DeptList | ?{$_.UserPrincipalName -eq $Item.EmailAddress})
    $FullSingleUser = ($Item | Select DisplayName,EmailAddress,LastLogon,@{Label="Department";Expression={$TempUser.Department}})
    $Output2 += $FullSingleUser
}

# Sort final output list by $SortField and output to file
$Output2 = $Output2 | Sort $SortField
$Output2 | Export-Csv $OutFile -NoTypeInformation

# Send email to $Recipients
Send-MailMessage -SmtpServer "mail.domain.org" -Subject "Weekly Email Report" -From ($From | Out-string) -To $Recipients -Body "See attached email report." -Attachment $Outfile

# Remove the csv file that was created
Remove-Item -Path $OutFile

# Close the Powershell sesion
Remove-PSSession $O365Session
