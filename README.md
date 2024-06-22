# MailboxSharing
GrantMailboxAccess.ps1 and UpdateMailboxAccess.ps1 are two scripts meant to be used in conjunction to simplify and automate the process of sharing mailboxes in Exchange Online for administrators.


GrantMailboxAccess.ps1 is a script that should be run manually by an administrator upon receiving a request to share a user's mailbox. It will take in two user's email addresses as well as a number of days. 

Upon verifying that the module is installed and authenticating with Exchange, it will grant User 2 full access to User 1's mailbox for how ever many days the administrator requested. 

The intended use case is for when an employee goes out of office, their supervisor can be granted access to their emails to review any incoming mail from clients only for the duration necessary. The script will report to a CSV file that is used as a database for mailbox access.


UpdateMailboxAccess.ps1 is a script that should be run automatically and/or regularly. It will review the CSV file and see which permissions have exceeded the number of days initially requested. 

It will then revoke all of those permissions to avoid lingering access rights that can pose a security risk and clutter an environment. 

The CSV file is updated to reflect which permissions have been successfully revoked to avoid errors from attempting to remove the same access rights twice.
