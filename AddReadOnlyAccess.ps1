# add read-only access to a mailbox through EXO, for any email addresses that are on a txt or csv
# Use: .\AddReadOnlyAccess.ps1 filename.csv 
# If you get an OAuth error about basic authentication needing to be enabled in our env run the BasicAuth script first then try again

# Connect powershell to Exo, this can be done in the cmd window first but included here so its not forgotten
Write-Host "Connecting to EXO..."
Connect-ExchangeOnline
Write-Host "Connection Completed"

# Get cmd line arguments and import csv file
$emailCsv = $args[0]
$listOfEmails = Import-Csv -Path .\$emailCsv

# Loop over list of emails and add permissions
foreach ($row in $listOfEmails) {
    $mailbox = $row.SharedMailboxAddress
    $user = $row.UsersEmailAddress

    # Get all folders in target mailbox
    Write-Host "Gathering Folders from" $mailbox
    $exclusions = @("/Audits",
                    "/Calendar Logging",
                    "/SubstrateHolds",
                    "/Sync Issues",
                    "/Sync Issues/Conflicts",
                    "/Sync Issues/Local Failures",
                    "/Sync Issues/Server Failures",
                    "/Recoverable Items",
                    "/Deletions",
                    "/Purges",
                    "/Versions"
                    )
    $mailboxFolders = @(Get-MailboxFolderStatistics $mailbox | Where-Object {!($exclusions -icontains $_.FolderPath)} | Select-Object FolderPath)

    Write-Host "Adding Read-Only Permissions to" $mailbox "for" $user

    # full access required to access mailbox from outlook, rights are then set per folder.
    Add-MailboxPermission -Identity $mailbox -User $user -AccessRights FullAccess -Confirm:$false
    
    # Loop over each folder in mailbox and assign Reviewer role to user
    foreach ($mailboxFolder in $mailboxFolders) {
        $folder = $mailboxFolder.FolderPath.Replace("/","\")
        if ($folder -match "Top of Information Store") {
            $folder = $folder.Replace("\Top of Information Store", "\")
        }
        $identity = "$($mailbox):$folder"

        # Check if folder permissions exist before adding or setting role
        # Error handeling could be better. Most common is folder not found, its fine to continue but try catch doesn't seem to work?
        $permissions = Get-EXOMailboxFolderPermission -Identity $identity -User $user -ErrorAction SilentlyContinue
        if($null -eq $permissions) {
            Write-Host "Adding Permissions for" $identity
            Add-MailboxFolderPermission -Identity $identity -User $user -AccessRights Reviewer -ErrorAction SilentlyContinue
        }
        else {
            Write-Host "Setting Permissions for" $identity
            Set-MailboxFolderPermission -Identity $identity -User $user -AccessRights Reviewer -ErrorAction SilentlyContinue
        }
    }

    Write-Host "Granted" $user "read only access to" $mailbox
    Write-Host ""
}
Write-Host "Done!"