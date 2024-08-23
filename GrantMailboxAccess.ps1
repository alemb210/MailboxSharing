# Gather input paramters from user
$mailbox = Read-Host "Please enter the username of the mailbox owner. (ex: user@yourdomain.com) "
$admin = Read-Host "Please enter the username of the user being granted mailbox access. (ex: supervisor@yourdomain.com) "
$days = Read-Host "Please enter the number of days you wish to grant access. "

# Set variables based on current date and given number of days
$date = "$(Get-Date)"
$expiration = "$($(Get-Date).AddDays($days))"

# Module verification and installation
if (Get-Module -ListAvailable -Name ExchangeOnlineManagement) { # Check if EO module is installed
    Write-Host "EXO module exists"
} 
else {
    Write-Host "EXO module does not exist. Installing..."
	Install-Module -Name ExchangeOnlineManagement -Force # Install EO module if needed
}

if (Get-Module -ListAvailable -Name Az.Accounts) { # Check if Azure Accounts module is installed
    Write-Host "AZ Accounts module exists"
} 
else {
    Write-Host "AZ Accounts module does not exist. Installing..."
	Install-Module -Name Az.Accounts -Force # Install Azure Accounts module if needed
}

if (Get-Module -ListAvailable -Name Az.Storage) { # Check if Azure Storage module is installed
    Write-Host "AZ Storage module exists"
} 
else {
    Write-Host "Az Storage module does not exist. Installing..."
	Install-Module -Name Az.Storage -Force # Install Azure Storage module if needed
}

# Grant mailbox access
Connect-ExchangeOnline # Authenticate to EO servers (Will open SSO prompt in browser)
Add-MailboxPermission -Identity $mailbox -user $admin -AccessRights FullAccess -Automapping $false # Grant full access to mailbox, disable mapping to avoid visual clutter.

# Authenticate to Azure (Will open SSO prompt in browser)
Connect-AzAccount -Subscription 'Azure Subscription'

# Retrieve CSV storage blob from Azure Storage
$StorageAccount = Get-AzStorageAccount -ResourceGroupName 'YourResourceGroup' -StorageAccountName 'YourStorageAccount'
$Context = $StorageAccount.Context
$ContainerName = 'StorageContainer'
$Blob = Get-AzStorageBlob -Container $ContainerName -Context $Context

# Load CSV text to memory and modify to reflect new changes
$Content = $Blob.ICloudBlob.DownloadText()
$Append = "`n`"$mailbox`",`"$admin`",`"$date`",`"$expiration`",`"$False`""
$NewContent = $Content + $Append

# Upload modified CSV back to Azure Storage
$Blob.ICloudBlob.UploadText($NewContent)
