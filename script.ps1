<#
.SYNOPSIS
There is a known behaviour with on-prem Exchange 2016 and Outlook 2013/2016 when Skype for Business client save multiple versions of the same conversation into Conversation History folder where all next
versions of the comversation contain all the messages which were already saved in previously plus new messages. By doing this it consumes a lot of space and generates a lot of useless items.
Outlook built-in clean up tools are not abe to handle this.  
 
This script finds all items in conversation history folder which relate to single coversation where only last saved version of such conversations
contains all the messages and all previous saved versions of the same conversation obviously contain only message up to the time when they were saved and so they
are not needed and just consume storage space.

Script was tested with on-prem Exchange 2016 and Outlook 2016 client running Windows 10 LTSC build. It is not guranteed that it will run properly in another environment. 
   
.DESCRIPTION
1. Script is using Outlook to process messages. It does not use EWS. Script should be run under user account who has access to mailbox.
2. To run the script you need to supply Outlook mailbox name, type of action (analyze/delete) and start date for filtering items. Analyze action will check items in Conversation History folder
and provide you the statistics on how many duplicate items you have and how much space they consume. Delete action will move all the duplicates to Deleted Items folder.
   
.NOTES
Author: https://github.com/dinozzaur
Created on: 19/01/2021
Modified on: 22/01/2021
Version: 1.1

.CHANGELOG
v.1.0 - initial version
v.1.1 - fixed date filtering and significantly improved speed of finding duplicates.

.TODO
Add logging
#>


function run-conversation-history-cleanup {
    Param
    (
         [ValidateNotNull()]
         [Parameter(Mandatory=$true, Position=0)]
         [string]$email,
         [ValidateSet(”analyze”,”delete”)]
         [Parameter(Mandatory=$true, Position=1)]
         [string]$action,
         [Parameter(Mandatory=$true, Position=2)]
         [string]$start_date
    )
    Process
    {
        $ol_app = New-Object -ComObject Outlook.Application
        $ol_ns = $ol_app.GetNameSpace('MAPI')

        Add-Type -AssemblyName 'Microsoft.Office.Interop.Outlook'
        $mailbox = $ol_ns.Folders |? {$_.FolderPath -eq "\\$email"}
        
        if ($mailbox) {Write-Host "Connected to $email mailbox" -ForegroundColor Green}
        else {Write-Error -Message "Unable to connect to the mailbox. Script execution will be terminated." -ErrorAction Stop}

        $conversation_history = $mailbox.Folders |? {$_.FolderPath -eq "\\$email\Conversation History"}      
        $str_start_date = ([datetime]::parseexact($start_date, 'dd.MM.yyyy', $null)).ToString("dd/MM/yyyy").Replace('.','/')
        $date_range_filter = "[SentOn] > '{$str_start_date}'"
        $total_items_count = $conversation_history.Items.Count()
        $filtered_items_count = $conversation_history.Items.Restrict($date_range_filter).Count()

        Write-Host "$total_items_count of items found in Conversation History folder of $email mailbox. " -ForegroundColor Green
        Write-Host "$filtered_items_count of items will be processed." -ForegroundColor Green

        $i = 1
        $conversation_history_items = @()
        
        foreach ($item in $conversation_history.Items.Restrict($date_range_filter)) {
    
            Write-Progress -Activity “Gathering information” -status “Processing item $i of $filtered_items_count” ` -percentComplete ($i / $filtered_items_count *100)

            $conversation_id = $item.ConversationID
            $item_date = $item.SentOn
            $item_entry_id = $item.EntryID
            $size = $item.size

            $object = [PSCustomObject]@{
                conversation_id     = $conversation_id
                sent_on = $item_date
                entry_id  = $item_entry_id
                size = $size / 1024
            }

            $conversation_history_items += $object
            $i++
        }

        Write-Progress -Activity “Gathering information” -Completed

        $i = 1 
        $size = 0
        $total_count_of_duplicates = 0

        if ($action -eq "delete") {$msgBoxInput = [System.Windows.MessageBox]::Show('Are you sure you want to move duplicates to Deleted Items folder?','Confirm deletion','YesNoCancel','Question')}
    
        switch($msgBoxInput) {
            "No" {[System.Windows.MessageBox]::Show("Script will analyze and show stats for duplicates","OK",'Information') | Out-Null }
            "cancel" {
                Write-Host "Script execution cancelled."
                exit
                }
            "Yes" {[System.Windows.MessageBox]::Show("Script will move all found duplicates to Deleted Items folder","OK",'Information') | Out-Null }  
        }

        $size = 0
        $unique_items = $conversation_history_items | Select -Property conversation_id -Unique
        $unique_items_count = $unique_items.Count
        foreach ($item in $unique_items) {
            [System.Collections.Generic.List[object]]$duplicates = ""
            $stats_for_duplicates = ""
            Write-Progress -Activity “Analyzing items” -status “Processing item $i of $unique_items_count unique items” ` -percentComplete ($i / $unique_items_count *100)
            $duplicates = $conversation_history_items.Where({$_.conversation_id -eq $item.conversation_id}) | Sort -Property sent_on -Descending
            if ($duplicates) {
                $duplicates.RemoveAt(0)
                $stats_for_duplicates = ($duplicates | Measure-Object -Property size -Sum)
                $size += $stats_for_duplicates.sum / 1024
                $total_count_of_duplicates += $stats_for_duplicates.Count
                if ($msgBoxInput -eq "Yes") {$duplicates |% {$ol_ns.GetItemFromID($_.entry_id).Delete()}}
            }
            $i++         
        }

        Write-Progress -Activity “Analyzing items” “Gathering information” -Completed
        [System.Windows.MessageBox]::Show("Total size of duplicate items: $size Mb`r`nTotal count of duplicate items: $total_count_of_duplicates",'Results','OK','Information') | Out-Null

        Write-Host "$filtered_items_count items analysis completed." -ForegroundColor Green
        Write-Host "Total size of duplicate items: $size Mb" -ForegroundColor Green
        Write-Host "Total count of duplicate items: $total_count_of_duplicates" -ForegroundColor Green
    }
}

run-conversation-history-cleanup -email (Read-Host -Prompt "Enter Outlook mailbox name") -action (Read-Host -Prompt "Enter desired action (analyze/delete)") -start_date (Read-Host -Prompt "Enter start date for filtering (use dd.MM.yyyy format)")
