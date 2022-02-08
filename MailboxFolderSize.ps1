# Written by BenschiBox 24-01-2022
# https://github.com/BenschiBox/M365-Scripts
# 
# This Script utilizes Exchange Online PowerShell V2 commands (Get-EXO...)

#$UserCredential = Get-Credential
#Connect-ExchangeOnline -Credential $UserCredential -ShowBanner:$false
#clear

$Mailbox = Read-Host -Prompt 'Mailbox Email-address or identifier'

$TopLevelOnly = $null
while (-not ($TopLevelOnly -is [bool])) {
    $TopLevelOnly = Read-Host -Prompt "Top-Level Paths only? (true/false)"
    try {
        $TopLevelOnly = [System.Convert]::ToBoolean($TopLevelOnly)
    } catch {
        Write-Host "true/false input only!"
    }
}

if ($TopLevelOnly) {
    Get-EXOMailboxFolderStatistics -Identity $Mailbox | 
    Where-Object {$_.folderpath.Substring(1).Contains("/") -eq $false} | 
    select FolderPath, ` @{name=”FolderAndSubfolderSize (MB)”; expression={[math]::Round( ` 
    ($_.FolderAndSubfolderSize.ToString().Split(“(“)[1].Split(” “)[0].Replace(“,”,””)/1MB),2)}} | 
    sort -Property "FolderAndSubfolderSize (MB)" -Descending |
    Export-csv .\MailboxFolderSize.csv -NoTypeInformation
} else {
    Get-EXOMailboxFolderStatistics -Identity $Mailbox | 
    select FolderPath, ` @{name=”FolderAndSubfolderSize (MB)”; expression={[math]::Round( ` 
    ($_.FolderAndSubfolderSize.ToString().Split(“(“)[1].Split(” “)[0].Replace(“,”,””)/1MB),2)}} | 
    sort -Property "FolderAndSubfolderSize (MB)" -Descending |
    Export-csv .\MailboxFolderSize.csv -NoTypeInformation
}

#Disconnect-ExchangeOnline -Confirm:$false -InformationAction Ignore -ErrorAction SilentlyContinue
