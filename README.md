# MailboxSharing
MailboxSharing is a PowerShell-focused project intended to simplify and automate the process of sharing mailboxes between users in Exchange Online for administrators. By automatically removing lingering access rights, potential for human error among support technicians is removed, creating a more secure and decluttered domain environment.

# RESTMailbox
The folder RESTMailbox contains the code for the website that is configured to run as an Azure Static Web App. Utilizing HTTP Forms and Fetch, the site can take in parameters from an administrator and send a POST request containing necessary data to a Function App configured to respond to HTTP requests.

# HTTPMailbox.ps1
HTTPMailbox.ps1 is a PowerShell script integrated as an Azure Function App, triggered by HTTP requests. 

It parses JSON data from the body of POST requests sent by the static web app (RESTMailbox folder), specifically a user who is requesting access, the user whose inbox they are requesting access to, and the number of days they should be granted access.

It will then authenticate to Exchange Online using a certificate in Azure Key Vault and assign the necessary permissions.

# UpdateMailboxAccess.ps1
UpdateMailboxAccess.ps1 is a PowerShell script which is used as an Azure Function App linked to a timer trigger. Depending on the trigger, the script will be run at set intervals and review the access rights in the database, revoking access and updating the 'Expired' column for all rights which have passed their expiration date.

It utilizes certificate-based authentication, with a self-signed certificate stored in Azure Key Vault, as well as an app registration and the Function App's managed identity with appropriate Exchange Online management permissions. This allows for a fully automated process that does not require any user input or prompts.

Please note that extra configuration may be required, such as configuring the Key Vault/Storage Account firewalls to allow access to Microsoft Azure CIDR blocks.
