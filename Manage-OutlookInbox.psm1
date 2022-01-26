[System.Reflection.Assembly]::LoadFile($PSScriptRoot + "\Microsoft.Identity.Client.dll") | Out-Null
[System.Reflection.Assembly]::LoadFile($PSScriptRoot + "\Microsoft.Exchange.WebServices.dll") | Out-Null

function Manage-OutlookInbox {

[CmdletBinding(DefaultParameterSetName='searchonly')]
Param(

[Parameter(Mandatory=$true,ParameterSetName = 'harddelete')]
[Parameter(Mandatory=$true,ParameterSetName = 'softdelete')]
[Parameter(Mandatory=$true,ParameterSetName = 'deletetodeleteditems')]
[Parameter(Mandatory=$true,ParameterSetName = 'movetofolder')]
[Parameter(Mandatory=$true,ParameterSetName = 'searchonly')]
$searchQuery,
[Parameter(ParameterSetName = 'harddelete')]
[Parameter(ParameterSetName = 'softdelete')]
[Parameter(ParameterSetName = 'deletetodeleteditems')]
[Parameter(ParameterSetName = 'movetofolder')]
[Parameter(ParameterSetName = 'searchonly')]
[Parameter(ParameterSetName = 'listCalendarInvites')]
[Parameter(ParameterSetName = 'hardDeleteCalendarInvites')]
[Parameter(ParameterSetName = 'softDeleteCalendarInvites')]
[Parameter(ParameterSetName = 'moveDeleteCalendarInvites')]
$Mailbox,


[parameter(parametersetname="harddelete")]
[Parameter(ParameterSetName = 'hardDeleteCalendarInvites')]
[switch]$HardDelete,

[parameter(parametersetname="softdelete")]
[Parameter(ParameterSetName = 'softDeleteCalendarInvites')]
[switch]$SoftDelete,

[parameter(parametersetname="deletetodeleteditems")]
[Parameter(ParameterSetName = 'moveDeleteCalendarInvites')]
[switch]$DeleteToDeletedItems,

[parameter(parametersetname="movetofolder")]
[string]$MoveToFolder,

[Parameter(ParameterSetName = 'harddelete')]
[Parameter(ParameterSetName = 'softdelete')]
[Parameter(ParameterSetName = 'deletetodeleteditems')]
[Parameter(ParameterSetName = 'movetofolder')]
[Parameter(ParameterSetName = 'searchonly')]
[parameter(parametersetname = "readCredentialsFromFile")]
[Parameter(ParameterSetName = 'listCalendarInvites')]
[Parameter(ParameterSetName = 'hardDeleteCalendarInvites')]
[Parameter(ParameterSetName = 'softDeleteCalendarInvites')]
[Parameter(ParameterSetName = 'moveDeleteCalendarInvites')]
[switch]$UseSavedCredentials,

[Parameter(ParameterSetName = 'listCalendarInvites')]
[Parameter(ParameterSetName = 'hardDeleteCalendarInvites')]
[Parameter(ParameterSetName = 'softDeleteCalendarInvites')]
[Parameter(ParameterSetName = 'moveDeleteCalendarInvites')]
[switch]$CleanupCalendarInvitesInInbox,


[Parameter(ParameterSetName = 'listCalendarInvites')]
[switch]$List,

[Parameter(ParameterSetName = 'harddelete')]
[Parameter(ParameterSetName = 'softdelete')]
[Parameter(ParameterSetName = 'deletetodeleteditems')]
[Parameter(ParameterSetName = 'movetofolder')]
[Parameter(ParameterSetName = 'searchonly')]
[string]$RootSearchFolderName,

[Parameter(ParameterSetName = 'New365OAuthToken')][switch]$newtoken

)

$functionName = $PSCmdlet.ParameterSetName

$sizeCounter = 0

if(!$newtoken){

$oauthcreds = Import-Clixml $($PSScriptRoot + "\oauth.xml")

$ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2016
$ewsURL = "https://outlook.office365.com/ews/exchange.asmx"

## Create Exchange Service Object 
$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($ExchangeVersion)
$service.Timeout = 15000

$service.Credentials = New-Object Microsoft.Exchange.WebServices.Data.OAuthCredentials -ArgumentList ($oauthcreds.AccessToken)

$service.Url = [system.URI]$ewsURL


function runSearchQuery {

    if(!$RootSearchFolderName){
        $inboxfolderid = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox,$mailbox)
        $inboxfolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$inboxfolderid)
    }
else
    {
            $view = [Microsoft.Exchange.WebServices.Data.FolderView]::new(1000)
            $view.PropertySet = [Microsoft.Exchange.WebServices.Data.PropertySet]::new([Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly)
            $view.PropertySet.Add([Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName)
            
            $searchFilter = [Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo]::new([Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName, "$RootSearchFolderName")
            
            $view.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Deep
            
            [Microsoft.Exchange.WebServices.Data.FindFoldersResults]$findFolderResults = $service.FindFolders([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Root, $searchFilter, $view)
            
            if ($findFolderResults.TotalCount -gt 1){throw "Multiple folders exist with that name. Please choose another folder."}
            
            $inboxfolderid = $findFolderResults.Folders[0].Id
            $inboxfolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$inboxfolderid)
    }
    
    $view = New-Object -TypeName Microsoft.Exchange.WebServices.Data.ItemView -ArgumentList 1000
    
    return $($inboxfolder.FindItems($searchQuery,$view))

}

function runCalendarCleanup {
Param(
[Parameter(Mandatory)]
[ValidateSet("List","HardDelete","SoftDelete","MoveToDeletedItems")]
$Action,
[int]$OlderThanDays = 7
)
    
    $view = New-Object -TypeName Microsoft.Exchange.WebServices.Data.ItemView -ArgumentList 1000
    $view.PropertySet = [Microsoft.Exchange.WebServices.Data.PropertySet]::new([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties, [Microsoft.Exchange.WebServices.Data.ItemSchema]::ItemClass)
    $filter = New-Object -TypeName Microsoft.Exchange.WebServices.Data.SearchFilter+ContainsSubstring -ArgumentList ([Microsoft.Exchange.WebServices.Data.ItemSchema]::ItemClass),("IPM.Schedule.Meeting")
    $filter2 = New-Object -TypeName Microsoft.Exchange.WebServices.Data.SearchFilter+IsLessThan -ArgumentList ([Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeReceived),((get-date).AddDays($OlderThanDays))
    
    $filterCol = [Microsoft.Exchange.WebServices.Data.SearchFilter+SearchFilterCollection]::new()
    $filterCol.LogicalOperator = [Microsoft.Exchange.WebServices.Data.LogicalOperator]::And
    $filterCol.Add($filter)
    $filterCol.Add($filter2)
    
    $results = $service.FindItems([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox,$filterCol,$view)
    $itemErrors = 0
    $i = 0

    switch($Action){
    "List" {$results | select subject,itemclass,datetimereceived | Out-GridView}
    "HardDelete" {
        $results.Items | % {
            $i++
            Write-Progress -Activity "Permanently Deleting Items" -Status ("Subject: " + $_.Subject) -PercentComplete ([Math]::Ceiling(($i/$results.Items.Count)*100))
                   try{
                        $_.Delete([Microsoft.Exchange.WebServices.Data.DeleteMode]::HardDelete) | Out-Null
                        $sizeCounter += $_.Size
                        Start-Sleep -Milliseconds 5
                       }
                   catch{
                            $itemErrors++
                        }
                   }
                   Write-Host $([string]($sizeCounter / 1048576) + " MB Deleted")
                   Write-Host $([string]$emailErrors + " Not Deleted Due to Errors")
                }
    "SoftDelete" {
        $results.Items | % {
            $i++
            Write-Progress -Activity "Soft Deleting Items" -Status ("Subject: " + $_.Subject) -PercentComplete ([Math]::Ceiling(($i/$results.Items.Count)*100))
                   try{
                        $_.Delete([Microsoft.Exchange.WebServices.Data.DeleteMode]::SoftDelete) | Out-Null
                        $sizeCounter += $_.Size
                        Start-Sleep -Milliseconds 5
                       }
                   catch{
                            $itemErrors++
                        }
                   }
                   Write-Host $([string]($sizeCounter / 1048576) + " MB Deleted")
                   Write-Host $([string]$emailErrors + " Not Deleted Due to Errors")
                }
    "MoveToDeletedItems" {
        $results.Items | % {
            $i++
            Write-Progress -Activity "Permanently Deleting Items" -Status ("Subject: " + $_.Subject) -PercentComplete ([Math]::Ceiling(($i/$results.Items.Count)*100))
                   try{
                        $_.Delete([Microsoft.Exchange.WebServices.Data.DeleteMode]::MoveToDeletedItems) | Out-Null
                        $sizeCounter += $_.Size
                        Start-Sleep -Milliseconds 5
                       }
                   catch{
                            $itemErrors++
                        }
                   }
                   Write-Host $([string]($sizeCounter / 1048576) + " MB Deleted")
                   Write-Host $([string]$emailErrors + " Not Deleted Due to Errors")
                }
    }

}


}


function New-365OAuthToken{   
    $pcaOptions = New-Object Microsoft.Identity.Client.PublicClientApplicationOptions
    $pcaOptions.ClientId = "0e4bf2e2-aa7d-46e8-aa12-263adeb3a62b"
    $pcaOptions.RedirectUri = "https://microsoft.com/EwsEditor"
    
    $pca = [Microsoft.Identity.Client.PublicClientApplicationBuilder]::CreateWithApplicationOptions($pcaOptions).Build()
    
    [string[]]$ewsScopes = "https://outlook.office365.com/EWS.AccessAsUser.All"
    
    $authResult = $pca.AcquireTokenInteractive($ewsScopes).ExecuteAsync().GetAwaiter().GetResult()
    
    $authResult | Export-Clixml $($PSScriptRoot + "\oauth.xml")
    }

switch($functionName)
    {
    "searchonly" {
                    $foundemails = runSearchQuery
                    $foundemails | select From,Subject,Size,ConversationId | Out-GridView
                    $foundemails | % {$sizeCounter += $_.Size}
                    Write-Host $([string]($sizeCounter / 1048576) + " MB in search")
                 }
    "harddelete" {
                    $foundemails = runSearchQuery
                    $deleteMode = [Microsoft.Exchange.WebServices.Data.DeleteMode]::HardDelete
                    $progressText = "! PERMANENTLY DELETING MESSAGES !"
                    deleteOutlookItem -foundemails $foundemails -progressText $progressText -deleteMode $deleteMode
                 }
    "softdelete" {
                    $foundemails = runSearchQuery
                    $deleteMode = [Microsoft.Exchange.WebServices.Data.DeleteMode]::SoftDelete
                    $progressText = "Soft-Deleting Messages"
                    deleteOutlookItem -foundemails $foundemails -progressText $progressText -deleteMode $deleteMode
                 }
    "deletetodeleteditems" {
                    $foundemails = runSearchQuery
                    $deleteMode = [Microsoft.Exchange.WebServices.Data.DeleteMode]::MoveToDeletedItems
                    $progressText = "Moving Messages to Deleted Items Folder"
                    deleteOutlookItem -foundemails $foundemails -progressText $progressText -deleteMode $deleteMode
                 }
    "movetofolder" {
                    $foundemails = runSearchQuery
                    moveItemToFolder -foundemails $foundemails -service $service -MoveToFolder $MoveToFolder
                   }
    "writeCredentialFile" {
                    writeCredentialFile
                   }
    "listCalendarInvites" {
                    runCalendarCleanup -Action List
                   }
    "hardDeleteCalendarInvites" {
                    runCalendarCleanup -Action HardDelete
                   }
    "softDeleteCalendarInvites" {
                    runCalendarCleanup -Action SoftDelete
                   }
    "moveDeleteCalendarInvites" {
                    runCalendarCleanup -Action MoveToDeletedItems
                   }
    "New365OAuthToken" {
                    New-365OAuthToken
                   }
    }
}

function deleteOutlookItem {
Param(
    $foundemails,
    $progressText,
    $deleteMode
)

    $i = 0
    $emailErrors = 0
    $foundemails | % {
        $i++
        Write-Progress -Activity $progressText -Status ("Subject: " + $_.Subject) -PercentComplete ([Math]::Ceiling(($i/$foundemails.Count)*100))
        try{
            $_.Delete($deleteMode) | Out-Null
            $sizeCounter += $_.Size
            Start-Sleep -Milliseconds 10
           }
       catch
           {
            $emailErrors++
           }
    }
    Write-Host $([string]($sizeCounter / 1048576) + " MB Deleted")
    Write-Host $([string]$emailErrors + " Not Deleted Due to Errors")

}

function moveItemToFolder {
Param(
    $foundemails,
    $service,
    $MoveToFolder
)
    $emailErrors = 0
    $view = [Microsoft.Exchange.WebServices.Data.FolderView]::new(1000)
    $view.PropertySet = [Microsoft.Exchange.WebServices.Data.PropertySet]::new([Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly)
    $view.PropertySet.Add([Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName)
    
    $searchFilter = [Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo]::new([Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName, "$MoveToFolder")
    
    $view.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Deep
    
    [Microsoft.Exchange.WebServices.Data.FindFoldersResults]$findFolderResults = $service.FindFolders([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Root, $searchFilter, $view)
    
    if ($findFolderResults.TotalCount -gt 1){throw "Multiple folders exist with that name. Please choose another folder."}
    
    $folderId = $findFolderResults.Folders[0].Id

    $i = 0
    $foundemails | % {
            $i++
            $sizeCounter += $_.Size
            Write-Progress -Activity "Moving Messages" -Status ("Subject: " + $_.Subject) -PercentComplete ([Math]::Ceiling(($i/$foundemails.Count)*100))
            try{
                $_.Move($folderId) | Out-Null
            }
            catch{
            $emailErrors++
            }
        }
        Write-Host $([string]($sizeCounter / 1048576) + " MB Moved")
        Write-Host $([string]$emailErrors + " Not Moved Due to Errors")
}

Export-ModuleMember Manage-OutlookInbox