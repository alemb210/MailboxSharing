# Input bindings are passed in via param block.
param($Timer)

# The 'IsPastDue' property is 'true' when the current function invocation is later than scheduled.
if ($Timer.IsPastDue) {
    Write-Host "PowerShell timer is running late!"
}

# Define variables for storage account, container, blob, and organization
$storageAccountName = "YourStorageAccount"
$containerName = "StorageContainer"
$blobName = "MockDB.csv"
$Org = "yourdomain.onmicrosoft.com"
$tempFilePath = "${env:HOME}\${blobName}" # Temporary path in Azure Function environment
$currentDate = [System.TimeZoneInfo]::ConvertTime($(Get-Date), [System.TimeZoneInfo]::FindSystemTimeZoneById('Eastern Standard Time')) #Adjust time to EST (Azure Functions will run on UTC)

# Create a context with the managed identity
$ctx = New-AzStorageContext -StorageAccountName $storageAccountName -UseConnectedAccount

# Download the CSV file from Blob Storage
Get-AzStorageBlobContent -Container $containerName -Blob $blobName -Destination $tempFilePath -Context $ctx -Force

# Import CSV to be modified
$data = Import-Csv -Path $tempFilePath

# Filter expired entries that are not yet marked as expired
$expired = $data | Where-Object { 
    [datetime]::Parse($_.ExpirationDate) -lt $currentDate -and $_.Expired -eq $false 
}

# Define Key Vault name and secret names
$keyVaultName = "Your-Key-Vault"
$passwordSecretName = "CertificatePassword"

# Retrieve password from Azure Key Vault and convert to secure string
$Password = Get-AzKeyVaultSecret -VaultName $keyVaultName -Name $passwordSecretName -AsPlainText | ConvertTo-SecureString -AsPlainText -Force

# Retrieve App Registration ID from Azure Key Vault
$AppID = Get-AzKeyVaultSecret -VaultName $keyVaultName -Name "YourAppID" -AsPlainText

# Retrieve the certificate from Key Vault
$secretText = Get-AzKeyVaultSecret -VaultName $keyVaultName -Name "YourCertificate" -AsPlainText
$certBytes = [Convert]::FromBase64String($secretText)
$certCollection = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2Collection
$certCollection.Import($certBytes, $null, [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::Exportable)
$cert = $certCollection[0]

# Authenticate to Exchange Online using the certificate and password
Connect-ExchangeOnline -Certificate $cert -CertificatePassword $Password -AppId $AppID -Organization $Org 

# Loop through expired entries and remove mailbox permissions
foreach ($perm in $expired) {
    Remove-MailboxPermission -Identity $perm.MailboxOwner -User $perm.GrantAccessTo -AccessRights FullAccess -Confirm:$false
    $perm.Expired = $true
}

# Export the updated CSV
$data | Export-Csv -Path $tempFilePath -NoTypeInformation -Confirm:$false

# Upload the modified CSV back to Azure Blob Storage
Set-AzStorageBlobContent -File $tempFilePath -Container $containerName -Blob $blobName -Context $ctx -Force

# Clean up local CSV file if no longer needed
Remove-Item -Path $tempFilePath
