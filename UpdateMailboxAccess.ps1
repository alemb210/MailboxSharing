$csv = "MockDB.csv"  #Path to a CSV file that is used as a database for mailbox access. Does not have to exist prior to runtime.
$data = Import-CSV $csv #Import CSV to be modified
$expired = ($data | Where-Object { [datetime]::Parse($_.ExpirationDate) -lt (Get-Date) -and $_.Expired -eq $false }) #Select objects that are past the expiration date and have not been modified
if (Get-Module -ListAvailable -Name ExchangeOnlineManagement) { #Check if EO module is installed
    Write-Host "Module exists"
} 
else {
    Write-Host "Module does not exist. Installing..."
	Install-Module -Name ExchangeOnlineManagement -Force #Install EO module if needed
}
Connect-ExchangeOnline #Authenticate to EO servers
foreach($perm in $expired){ #For each expired permission, remove the full access rights
	Remove-MailboxPermission -Identity $perm.MailboxOwner -User $perm.GrantAccessTo -AccessRights FullAccess
	$perm.Expired = $true #Boolean flag to determine whether access has been removed yet (avoid redundancy)
}
$data | Export-CSV -Path $csv #Update CSV with changes