<# 
    .SYNOPSIS
    Sets Access Rights for multiple users on multiple mailboxes.

    .DESCRIPTION
    Sets Access Rights for multiple users on multiple mailboxes.
    Takes any number of users as strings, any number of mailboxes as srings, and the access rights to add.

    .PARAMETER users
    Users to grant the rights to, should be strings seperated by commas

    .PARAMETER mailboxes
    Mailboxes where rights are needed, should be strings seperated by commas

    .PARAMETER accessRights
    Access rights to grant (FullAccess, ReadPermissions, SendAs, etc.)
#>
param (
    [Parameter(Mandatory=$true)]
    [string[]]$users, 
    [Parameter(Mandatory=$true)]
    [string[]]$mailboxes,
    [Parameter(Mandatory=$true)]
    [string]$accessRights
)

# Try to Connect to Exchange
try {
    Connect-ExchangeOnline
}

catch [CommandNotFoundException]{
    # Exchange Not Installed --> try to install
    Write-Output "Exchange Commands not Found"
    # Check if running as an admin
    $id = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    $p = New-Object System.Security.Principal.WindowsPrincipal($id)

    if ($p.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator))
    {
        # Install Exchange Online Tools
        Write-Output "Installing Exchage Commands"
        Install-Module -Name PSWSMan
        Install-WSMan
        Connect-ExchangeOnline
    }      
 
     else
    { Write-Output "Please Run as admin to install Exchange Online Commands" }  
}
# Loop through all mailboxes
foreach ($mailbox in $mailboxes) {
    # Loop through all users
    foreach ($user in $users) {

        # Get the user's access rights for the mailbox.
        $u = Get-MailboxPermission -Identity $mailbox -User $user
        $perms = $u.AccessRights

        # If they have different rights, set the new one. If they have no rights add the rights.
        if ($perms -contains $accessRights){
            Set-MailboxPermission -Identity $mailbox -User $user -AccessRights $accessRights
        } elseif ($null -eq $perms){
            Add-MailboxPermission -Identity $mailbox -User $user -AccessRights $accessRights
        }
    }
}