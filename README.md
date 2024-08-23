# MailboxSharing
GrantMailboxAccess.ps1 and UpdateMailboxAccess.ps1 are two scripts meant to be used in conjunction to simplify and automate the process of sharing mailboxes between users in Exchange Online for administrators. By automatically removing lingering access rights, potential for human error among support technicians is removed, creating a more secure and decluttered domain environment.

# GrantMailboxAccess.ps1
GrantMailboxAccess.ps1 is a script that should be run manually by an administrator upon receiving a support request to share a user's mailbox. It will prompt for three parameters -- the first, being the email address of the user who's inbox is being accessed. The second prompt is the email address of the user requesting access. The third is the number of days in which this user should have access, as an integer.

It will then verify the proper modules are installed, and install them if necessary. 

It will prompt the user to connect to Exchange Online, after which the mailbox access will be granted, and then Microsoft's Azure, in order to access the database.

After, the new access rights will be uploaded to the database in Azure Storage and the script will complete.

# UpdateMailboxAccess.ps1
UpdateMailboxAccess.ps1 is a PowerShell script which is used as an Azure Function App linked to a timer trigger. Depending on the trigger, the script will be run at set intervals and review the access rights in the database, revoking access and updating the 'Expired' column for all rights which have passed their expiration date.

It utilizes certificate-based authentication, with a self-signed certificate stored in Azure Key Vault, as well as an app registration and the Function App's managed identity with appropriate Exchange Online management permissions. This allows for a fully automated process that does not require any user input or prompts.

Please note that extra configuration may be required, such as configuring the Key Vault/Storage Account firewalls to allow access to Microsoft Azure CIDR blocks.
