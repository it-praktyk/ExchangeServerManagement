Function Add-MailboxSubfolderPermission {
	
<#
    .SYNOPSIS
    Function intended for add permissions to subfolders in mailbox on Exchange
    
    .PARAMETER Identity
    
    .PARAMETER User
    
    .PARAMETER AccessRights
    
    .PARAMETER SubFolder
           
    .EXAMPLE
    
    Add-MailboxSubfolderPermission -Identity A155658 -User A589879TEST -SubFolder /ToReview -AccessRights Reviewer
    
    This command add permission 'Reviewer' on the folder ToReview in Mailbox A155658 to the user A589879TEST. 
    The folder ToReview exist in the root of the mailbox and outside of Inbox 
     
    .LINK
    https://github.com/it-praktyk/
    
    .LINK
    https://www.linkedin.com/in/sciesinskiwojciech
          
    .NOTES
    AUTHOR: Wojciech Sciesinski, wojciech[at]sciesinski[dot]net
    KEYWORDS: PowerShell

    Code partially based on 
    http://exchangeserverpro.com/grant-read-access-exchange-mailbox/
   
    VERSIONS HISTORY
    - 0.1.0 - 2016-02-12 - The first draft
    - 0.2.0 - 2016-02-12 - the permission on the top of information store added, errors corrected

    TODO
    - update help
    - check permissions on the top of information store
	- handle errors in situation when the permissions for the user exist now
        
    LICENSE
    Copyright (c) 2016 Wojciech Sciesinski
    This function is licensed under The MIT License (MIT)
    Full license text: https://opensource.org/licenses/MIT

#>
    
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [alias("Mailbox")]
        [string]$Identity,
        [Parameter (Mandatory=$true)]
        [alias("Folder","Path")]
        [String]$SubFolder,
        [Parameter(Mandatory = $true)]
        [string]$User,
        [Parameter(Mandatory = $true)]
        [string]$AccessRights
    )
    
    $exclusions = @("/Sync Issues",
    "/Sync Issues/Conflicts",
    "/Sync Issues/Local Failures",
    "/Sync Issues/Server Failures",
    "/Recoverable Items",
    "/Deletions",
    "/Purges",
    "/Versions"
    )
    
    $mailboxfolders = @(Get-MailboxFolderStatistics $Mailbox | Where-Object -FilterScript { !($exclusions -icontains $_.FolderPath) } | Select-Object -Property FolderPath)
    
    foreach ($mailboxfolder in $mailboxfolders) {
    
        $folder = $mailboxfolder.FolderPath.Replace('/', '\')
        
        if ($folder -match 'Top of Information Store') {
            
            $folder = $folder.Replace('\Top of Information Store', '\')
            
        }
        
        If ($folder -eq $subfolder -or $folder -eq $subfolder) {
            
            [String]$identity = "{0}:{1}" -f $mailbox, $folder
            
            [String]$MessageText = "Adding {0} to {1} with {2} permissions" -f $user, $identity, $access
            
            Write-Verbose -Message $MessageText
            
            Add-MailboxFolderPermission -Identity $identity -User $user -AccessRights $AccessRights #-ErrorAction SilentlyContinue
            
            $permissionadded = $true
            
        }
        
    }
    
    If ($permissionadded) {
        
        #Check if the user has access to the "Top of Information Store" folder
        
        $RequiredPermissions = "FolderVisible", "ReadItems", "FolderOwner"
        
        $UserRightsOnTop = Get-MailboxFolderPermission -Identity $mailbox -User $User -ErrorAction SilentlyContinue
        
        
        If (-not $($RequiredPermissions -icontains (out-string -InputObject $UserRightsOnTop.AccessRights))) {
            
            [String]$MessageText = "Adding FolderVisible permision on {0}:\ to {1}" -f $mailbox , $user
            
            Write-Verbose -Message $MessageText
            
            [String]$TopFolder = "{0}:\" -f $Mailbox
            
            Add-MailboxFolderPermission -Identity $TopFolder -User $user -AccessRights FolderVisible #-ErrorAction SilentlyContinue | out-Null
            
        }
        
    }
    
    
    Else {
        
        [String]$MessageText = "The folder {0} was not found in the mailbox {1}" -f $SubFolder, $Mailbox
        
        Write-Error -Message $MessageText
        
    }
    
}