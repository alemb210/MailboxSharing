$mailbox = Read-Host "Please enter the username of the mailbox owner. (ex: user@yourdomain.com) "
$admin = Read-Host "Please enter the username of the user being granted mailbox access. (ex: supervisor@yourdomain.com) "
$days = Read-Host "Please enter the number of days you wish to grant access. "
$csv = "MockDB.csv" #Path to a CSV file that is used as a database for mailbox access. Does not have to exist prior to runtime.
if (Get-Module -ListAvailable -Name ExchangeOnlineManagement) { #Check if EO module is installed
    Write-Host "Module exists"
} 
else {
    Write-Host "Module does not exist. Installing..."
	Install-Module -Name ExchangeOnlineManagement -Force #Install EO module if needed
}
Connect-ExchangeOnline #Authenticate to EO servers
Add-MailboxPermission -Identity $mailbox -user $admin -AccessRights FullAccess -Automapping $false #Grant full access to mailbox, do not map it to Outlook.
$permission = New-Object PsObject -Property @{MailboxOwner = $mailbox ; GrantAccessTo = $admin ; DateSet = Get-Date ; ExpirationDate = (Get-Date).AddDays($days) ; Expired = $false} #Create an array that can be easily exported as a row in CSV
echo $permission | Select-Object MailboxOwner, GrantAccessTo, DateSet, ExpirationDate, Expired | Export-CSV -Path $csv -Append #Add new row to CSV
echo "${admin} has access to ${mailbox}'s inbox for ${days} days"