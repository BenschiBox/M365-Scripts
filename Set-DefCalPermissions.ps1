<#
    Orginal Script written by ALI TAJRAN
    https://www.alitajran.com/set-default-calendar-permissions-for-all-users-powershell/

    Adapted by BenschiBox 11-02-2025
    https://github.com/BenschiBox/M365-Scripts

    This Script utilizes the Exchange Online PowerShell V3 commands (Get-EXO...)
    Append -whatif to the Set-MailboxFolderPermission command for testing
#>

# Start transcript
Start-Transcript -Path C:\temp\Set-DefCalPermissions.log -Append

# Set scope to entire forest. Cmdlet only available for Exchange on-premises.
#Set-ADServerSettings -ViewEntireForest $true

# Get all user mailboxes
$Users = Get-EXOMailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox

# Users exception
#$Exception = @("CEO Mailbox", "CTO Mailbox")
$Exception = @( "admin@contoso.com", 
                "CEO@contoso.com")

# Permissions
# full list: https://docs.microsoft.com/en-us/powershell/module/exchange/set-mailboxfolderpermission?view=exchange-ps#parameters
$Permission = "Reviewer"

# Calendar name languages
$FolderCalendars = @("Agenda", "Calendar", "Calendrier", "Kalender", "æ—¥åŽ†")

# Loop through each user
foreach ($User in $Users) {

    # Get calendar in every user mailbox
    $Calendars = (Get-EXOMailboxFolderStatistics $User.UserPrincipalName -FolderScope Calendar)

    # Leave permissions if user is exception
    if ($Exception -Contains ($User.UserPrincipalName)) {
        Write-Host "$User is an exception, don't touch permissions" -ForegroundColor Red
    }
    else {

        # Loop through each user calendar
        foreach ($Calendar in $Calendars) {
            $CalendarName = $Calendar.Name

            # Check if calendar exist
            if ($FolderCalendars -Contains $CalendarName) {
                $Cal = "$($User.UserPrincipalName):\$CalendarName"
                $CurrentMailFolderPermission = Get-EXOMailboxFolderPermission -Identity $Cal -User Default
                
                # Update calendar permissions if necessary
                if ($CurrentMailFolderPermission.AccessRights -ne "$Permission") {
                    # Set calendar permission / Remove -WhatIf parameter after testing
                    Set-MailboxFolderPermission -Identity $Cal -User Default -AccessRights $Permission -WarningAction:SilentlyContinue #-WhatIf
                    Write-Host $User.DisplayName added permissions $Permission -ForegroundColor Green
                }
                else {
                    Write-Host $User.DisplayName already has the permission $Permission -ForegroundColor Yellow
                }            }
        }
    }
}

Stop-Transcript