Install-Module -Name ExchangeOnlineManagement
#Requires -Modules "ExchangeOnlineManagement"

$proceed = $True

$TechnicianEmail = Read-Host "Enter your First American email address"
Connect-ExchangeOnline -UserPrincipalName $TechnicianEmail
Write-Output "Connecting to Exchange server..."
Write-Output "`n"

# Get-InstalledModule -Name "ExchangeOnlineManagement"
do {
    Write-Output "---------------------------------------------------------------------"
    Write-Output "Type '1' to check Public Folder Access"
    Write-Output "Type '11' to check Mailbox delegate permissions"
    Write-Output "Type '2' to check RA Mailbox Permissions"
    Write-Output "Type '3' to check if RA Mailbox is Office365 or OnPrem"
    Write-Output "Type '4' to check Mailbox general information"
    Write-Output "Type '5' to see a description of Inbox rules"
    Write-Output "Type '6' to check list of Folders and number of items per folder"
    Write-Output "Type '66' to check list of Items per folder and export to a CSV file"
    Write-Output "Type '7' to check mail forward settings"
    Write-Output "Type '8' to re-authenticate to the Exchange server"
    Write-Output "Type '9' to exit the program "
    Write-Output "--------------------------------------------------------------------- `n"

    $WhatToDo = Read-Host "Please enter a number"
    $WhatToDo.Trim()

    if($WhatToDo -eq "1") {
        Write-Output "Example Public Folder Path: ""\FA Title Midwest\Closing Calendars\Wyoming\Lander"" "
        $FolderPath = Read-Host "Please enter the path to the Public Folder without the quotes("")"
        $FolderPath.Trim()
        Write-Output "Checking Public Folder Access..."
        Get-PublicFolderClientPermission $FolderPath
        Write-Output "`n"
    }

    elseif($WhatToDo -eq "11") {
        $MailboxName = Read-Host "Please enter mailbox email address to check delegate permissions"
        $MailboxName = $MailboxName.Trim() + ":\Calendar"
        Write-Output "Checking Mailbox Delegate Permissions..."
        Get-MailboxFolderPermission $MailboxName
        Write-Output "`n"
    }

    elseif($WhatToDo -eq "2") {
        $MailboxName = Read-Host "Please enter mailbox email address"
        $MailboxName.Trim()
        Write-Output "Checking RA Mailbox Permissions..."
        Get-MailboxPermission $MailboxName
        Write-Output "`n"
    }

    elseif($WhatToDo -eq "3") {
        $MailboxName = Read-Host "Please enter mailbox email address"
        $MailboxName.Trim()
        $Mailboxtype = Get-Recipient $MailboxName
        # Write-Output $MailboxType
        Write-Output "Checking RA Mailbox Type..."
        if ($MailboxType.RecipientType -eq "MailUser"){
            Write-Output "Mailbox type: OnPrem `n"
        }
        elseif ($MailboxType.RecipientType -eq "UserMailbox"){
            Write-Output "Mailbox type: Office365 `n"
        }
        else {
            Write-Output "Mailbox type not found. `n"
        }
    }

    elseif($WhatToDo -eq "4") {
        $MailboxName = Read-Host "Please enter mailbox email address"
        Write-Output "Checking general mailbox information..."
        Get-Mailbox -Identity $MailboxName | Format-List 
        Write-Output "`n"
    }

    elseif($WhatToDo -eq "5") {
        $MailboxName = Read-Host "Please enter mailbox email address"
        Write-Output "Checking Inbox rules..."
        Get-InboxRule –Mailbox $MailboxName  | Select-Object Name, Description | Format-List
        Write-Output "`n"
    }

    elseif($WhatToDo -eq "6") {
        $MailboxName = Read-Host "Please enter mailbox email address"
        Write-Output "Checking folder list and number of items per folder..."
        Get-MailboxFolderStatistics -Identity $MailboxName | Sort-Object FolderPath | ForEach-Object {
            $Count = ($_.FolderPath | Select-String -Pattern '/' -AllMatches).Matches.Count
            if($Count -gt 1){
                $Indent = "`t" * $Count
            } else {
                $Indent = ''
            }
            Write-Host "$($Indent)$($_.Name)`t$($_.ItemsInFolder)"
        }
        Write-Output "`n"
    }

    elseif($WhatToDo -eq "66") {
        $MailboxName = Read-Host "Please enter mailbox email address"
        Write-Output "Checking folder list and number of items per folder..."
        Get-EXOMailboxFolderStatistics -Identity $MailboxName | select-object Identity, ItemsInFolder, FolderSize | Export-csv C:\temp\ExchangeExport\MailboxStats.csv -NoTypeInformation
        $folderPath = "C:\temp\ExchangeExport"
        if (-not (Test-Path $folderPath)){
            New-Item -Path $folderPath -ItemType Directory
        }
        Write-Output "Export Complete, check for file at 'C:\temp\ExchangeExport\MailboxStats.csv'`n"
        Invoke-Item C:\temp\ExchangeExport
    }

    elseif($WhatToDo -eq "7") {
        $MailboxName = Read-Host "Please enter mailbox email address"
        Write-Output "Checking folder list and number of items per folder..."
        get-mailbox $MailboxName | Format-List *forward*
    }

    elseif($WhatToDo -eq "8") {
        Write-Output "Authentication started..."
        Connect-ExchangeOnline -UserPrincipalName $TechnicianEmail
        Write-Output "`n"
    }

    elseif($WhatToDo -eq "9") {
        Write-Output "Ending program..."
        $proceed = $False
    }
} until ($proceed -eq $False)