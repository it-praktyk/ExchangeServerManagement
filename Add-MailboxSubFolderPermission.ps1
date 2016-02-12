Function Add-MailboxSubfolderPermission {

<#
    .SYNOPSIS
    Function intended for add permissions to subfolders in mailbox on Exchange
	
   
    .EXAMPLE
	
	
     
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
    0.1.0 -  2016-02-12 - The first draft

    TODO
	- update help
	- check permissions on the top of information store
        
    LICENSE
	Copyright (c) 2016 Wojciech Sciesinski
    This function is licensed under The MIT License (MIT)
    Full license text: https://opensource.org/licenses/MIT

#>

[CmdletBinding()]
param (
	[Parameter( Mandatory=$true)]
	[string]$Mailbox,
	
	[Parameter ( Mandatory=$true)]
	[String]$SubFolder,
    
	[Parameter( Mandatory=$true)]
	[string]$User,
    
  	[Parameter( Mandatory=$true)]
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



$mailboxfolders = @(Get-MailboxFolderStatistics $Mailbox | Where {!($exclusions -icontains $_.FolderPath)} | Select FolderPath)

foreach ($mailboxfolder in $mailboxfolders) {

	

    $folder = $mailboxfolder.FolderPath.Replace("/","\")
    if ($folder -match "Top of Information Store")
    {
       $folder = $folder.Replace(“\Top of Information Store”,”\”)
	   
	   
    }
	
	If ($folder -eq $subfolder -or $folder -eq $subfolder) {
	
		$identity = "$($mailbox):$folder"
		
		[String]$MessageText = "Adding {0} to {1} with {2} permissions" -f $user , $identity, $access
		
		Write-Verbose -Message $MessageText
		
		Add-MailboxFolderPermission -Identity $identity -User $user -AccessRights $Access -ErrorAction SilentlyContinue
		
		$permissionadded = $true
	
	}

}

	If ( $permissionadded ) {
	
	<#
	
		#Check if the user has access to the "Top of Information Store" folder
		
		$RequiredPermissions ="FolderVisible","ReadItems","FolderOwner"
		
		Try {
		
			$Top = Get-MailboxFolderStatistics -Identity aa565615 | where { $_.name -eq "Top of informationstore" } | select folderpath 
		
			$UserRightsOnTop = Get-MailboxFolderPermission -Identity $mailbox -User $User -ErrorAction Continue | Out-Null
		

		Catch {
		
			[String]$MessageText = "Adding FolderVisible permision on {0} to {1}" -f "$mailbox):\" , $user
		
			Add-MailboxFolderPermission -Identity "$($mailbox):\" -User $user -AccessRights FolderVisible -ErrorAction SilentlyContinue | out-Null
		
		}
		
		#>
	
	}

	Else  {
	
		[String]$MessageText = "The folder {0} was not found in the mailbox {1}" -f $Folder, $Mailbox
		
		Write-Error -Message $MessageText
	
	}

}