function Get-MailboxPerms {
    $mailbox = Read-Host "What mailbox would you like to get?"
    Get-MailboxPermission -Identity $mailbox
}

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

$runCommands = $true
$commandOptions = "Press 1 to Manage MailboxPermissions"

while ($runCommands) {
    $select = Read-Host -Prompt $commandOptions

    switch ($select) {
        1 { 
         }
        Default { $runCommands = $false}
    }

    $continue = Read-Host "Would you like to continue? (Y/N)"
    if ($continue -like "Y*") {
        $runCommands = $true
    } elseif ($continue -like "y*"){
        $runCommands = $true
    } else {
        $runCommands = $false
    }
}

