# Written by BenschiBox 24-01-2022
# https://github.com/BenschiBox/M365-Scripts
# 
# adapted from:
# PurgeMessagesWithContentSearch.PS1
# https://github.com/12Knocksinna/Office365itpros/blob/master/PurgeMessagesWithContentSearch.PS1
# 
# and
# https://stackoverflow.com/users/67419/neildeadman
# https://stackoverflow.com/questions/62681477/o365-compliance-search-harddelete-not-working

Clear-Host

# Connect to Exchange Online
$credentials = get-credential
Connect-ExchangeOnline -Credential $credentials -ShowBanner:$false
# $SccSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.compliance.protection.outlook.com/powershell-liveid/ -Credential $credentials -Authentication "Basic" -AllowRedirection;
# Import-PSSession $SccSession
Connect-IPPSSession -Credential $credentials

# optional
# $mailboxes = @("mailbox1@example.net", "mailbox2@example.net")
# $monthsToKeep = 3
# $sourceDate = (Get-Date).AddMonths(-$monthsToKeep)
# $searchName = "PurgeEmails-Powershell"
# $contentQuery = "received<=$($sourceDate) AND kind:email"

$mailboxes = @("office@idontknow.com")
$searchName = "PurgeEmails-Powershell"
$contentQuery = "folderid:PASTEIDHERE AND kind:email"

# Clean-up any old searches from failed runs of this script
if (Get-ComplianceSearch -Identity $searchName) {
    Write-Host "Cleaning up any old searches from failed runs of this script"

    try {
        Remove-ComplianceSearch -Identity $searchName -Confirm:$false | Out-Null
    }
    catch {
        Write-Host "Clean-up of old script runs failed!" -ForegroundColor Red
        break
    }
}

# optional
# Write-Host "Creating new search for emails older than $($sourceDate)"

Write-Host "Creating new search $searchName with contentQuery $contentQuery"

New-ComplianceSearch -Name $searchName -ContentMatchQuery $contentQuery -ExchangeLocation $mailboxes -AllowNotFoundExchangeLocationsEnabled $true | Out-Null
                                                                            
Start-ComplianceSearch -Identity $searchName | Out-Null

Write-Host "Searching..." -NoNewline

while ((Get-ComplianceSearch -Identity $searchName).Status -ne "Completed") {
    Write-Host "." -NoNewline
    Start-Sleep -Seconds 2
}

$items = (Get-ComplianceSearch -Identity $searchName).Items

if ($items -gt 0) {
    $searchStatistics = Get-ComplianceSearch -Identity $searchName | Select-Object -Expand SearchStatistics | Convertfrom-JSON

    $sources = $searchStatistics.ExchangeBinding.Sources | Where-Object { $_.ContentItems -gt 0 }

    Write-Host ""
    Write-Host "Total Items found matching query:" $items 
    Write-Host ""
    Write-Host "Items found in the following mailboxes"
    Write-Host "--------------------------------------"

    foreach ($source in $sources) {
        Write-Host $source.Name "has" $source.ContentItems "items of size" $source.ContentSize
    }

    Write-Host ""
    $ContinueDeletion = $null
    while(-not ($ContinueDeletion -is [bool])) {
        $ContinueDeletion = Read-Host -Prompt "Continue with deletion? (true/false)"
        try {
            $ContinueDeletion = [System.Convert]::ToBoolean($ContinueDeletion)
        } catch {
            Write-Host "true/false input only!"
        }
    }
    Write-Host ""

    if($ContinueDeletion -eq $true) {
        $iterations = 0;
        $itemsProcessed = 0
        while ($itemsProcessed -lt $items) {
            $iterations++

            Write-Host "Deleting items iteration $($iterations)"

            New-ComplianceSearchAction -SearchName $searchName -Purge -PurgeType HardDelete -Confirm:$false | Out-Null

                while ((Get-ComplianceSearchAction -Identity "$($searchName)_Purge").Status -ne "Completed") { 
            Start-Sleep -Seconds 2
            }

            $itemsProcessed = $itemsProcessed + 10
        
            # Remove the search action so we can recreate it
            Remove-ComplianceSearchAction -Identity "$($searchName)_Purge" -Confirm:$false  
        }    
    } else {
        Write-Host ""
        Write-Host "Deletion Aborted!"
    }
} else {
    Write-Host ""
    Write-Host "No items found"
}

Disconnect-ExchangeOnline -Confirm:$false -InformationAction Ignore -ErrorAction SilentlyContinue

Write-Host ""
Write-Host "COMPLETED!"
