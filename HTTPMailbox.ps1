using namespace System.Net

# Input bindings are passed in via param block.
param($Request, $TriggerMetadata)

# Write to the Azure Functions log stream.
Write-Host "PowerShell HTTP trigger function processed a request."

# Log the request body
# Write-Host "request body"
# Write-Host $Request.Body

$storageAccountName = "YourStorageAccount"
$containerName = "StorageContainer"
$blobName = "data.csv"
$Org = "yourdomain.onmicrosoft.com"
$tempFilePath = "${env:HOME}\${blobName}" # Temporary path in Azure Function environment

# Define Key Vault name and secret names
$keyVaultName = "Your-Key-Vault"
$passwordSecretName = "Password"

# Retrieve password from Azure Key Vault and convert to secure string
$Password = Get-AzKeyVaultSecret -VaultName $keyVaultName -Name $passwordSecretName -AsPlainText | ConvertTo-SecureString -AsPlainText -Force

# Retrieve App ID from Azure Key Vault
$AppID = Get-AzKeyVaultSecret -VaultName $keyVaultName -Name "YourAppID" -AsPlainText

#Retrieve Certificate as Certificate object from Azure Key Vault
$secretText = Get-AzKeyVaultSecret -VaultName $keyVaultName -Name "YourCertificate" -AsPlainText
$certBytes = [Convert]::FromBase64String($secretText)
$certCollection = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2Collection
$certCollection.Import($certBytes, $null, [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::Exportable)
$cert = $certCollection[0]


# Interact with query parameters or the body of the request.
$mailbox = $Request.Query.Mailbox
$admin = $Request.Query.Admin
$days = $Request.Query.Days

if (-not $mailbox -or -not $admin -or -not $days) {
    try {
        Write-Host "Request Body: $($Request.Body | ConvertTo-Json -Depth 10)"
        # Directly access the hashtable value
        $mailbox = $Request.Body["mailbox"]
        $admin = $Request.Body["admin"]
        $days = $Request.Body["days"]
    }
    catch {
        Write-Host "Failed to access request body as hashtable. Error: $_"
    }
}

#Ensure days is converted to integer
if ($days -is [System.Management.Automation.OrderedHashtable]) {
    $days = [int]$days["value"]
} else {
    $days = [int]$days
}

$body = "This HTTP triggered function executed successfully. Pass a mailbox and admin in the query string or in the request body for a personalized response."

if ($mailbox -and $admin -and $days) {
    try {
        # Perform the action to change the email access right
        Write-Host "Changing access right for $mailbox to $admin"
        Connect-ExchangeOnline -Certificate $cert -CertificatePassword $Password -AppId $AppID -Organization $Org 
        Add-MailboxPermission -Identity $mailbox -user $admin -AccessRights FullAccess -Automapping $false -Confirm:$false
        if ($?) {
            $body = "Successfully changed access right for $mailbox to $admin."

            # Set variables based on current date and given number of days
            $date = "$(Get-Date)"
            $expiration = "$($(Get-Date).AddDays($days))"

            # For appending to Azure Blob Storage CSV
            $csvLine = "`"$mailbox`",`"$admin`",`"$date`",`"$expiration`",`"$False`""
            # Create a context with the managed identity
            $ctx = New-AzStorageContext -StorageAccountName $storageAccountName -UseConnectedAccount

            # Download the CSV file from Blob Storage
            Get-AzStorageBlobContent -Container $containerName -Blob $blobName -Destination $tempFilePath -Context $ctx -Force

            # Read the existing CSV content
            $blob = Get-AzStorageBlobContent -Blob $blobName -Container $containerName -Context $ctx -Destination "temp.csv"
            $csvContent = Get-Content "temp.csv"
            
            # Append the new line
            $csvContent += $csvLine
            
            # Write back to the blob
            Set-Content -Path "temp.csv" -Value $csvContent
            Set-AzStorageBlobContent -File "temp.csv" -Container $containerName -Blob $blobName -Context $ctx -Confirm:$false
            
            # Clean up
            Remove-Item "temp.csv"
        }
        else {
            $body = "Failed to change access right for $mailbox."
        }
    }
    catch {
        Write-Host "Error changing access right. Error: $_"
        $body = "Error changing access right for $mailbox."
    }
}
else {
    $body = "Mailbox and admin must be provided in the query string or request body."
}

# Associate values to output bindings by calling 'Push-OutputBinding'.
Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        StatusCode = [HttpStatusCode]::OK
        Body       = $body
    })