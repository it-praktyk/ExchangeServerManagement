Function Start-ExchangeRecipientsObjectManagement {

<#   
    .SYNOPSIS   
    Function for manage Exchange recipient objects in batch mode
  
    .DESCRIPTION   
	Function for manage Exchange recipients objects in batch mode using csv files as input  - specially for operation on email addresses (proxyaddresses)
	
    .PARAMETER InputFilePath
    The path for file with input data
    
    .PARAMETER VerifyInputFileForDuplicates
    By default input file is verified for duplicates
    
    .PARAMETER Mode
    The switch which define action to perform - default mode is DisplayOnly
	
	Available modes
	- DisplayOnly
	- PerformActions
	- CreatePerformActionsCommandsOnly
	- Rollback
	- CreateRollbackCommnadsOnly

    .PARAMETER DomainName
    Active Directory domain name - FQDN

    .PARAMETER CreateRollbackFile
    By default roolback file is created

    .PARAMETER RollBackFileDirectoryPath
    By default rollback files are stored in subfolder rolbacks in current path, if rollbacks subfolder is missed will be created
    
    .PARAMETER RollBackFileNamePrefix
    Prefix used for creating rollback files name. Default is "Rollback-"

    .PARAMETER CreateTranscript
    By default transcript file is created

    .PARAMETER $TranscriptFileDirectoryPath
    By default transcript files are stored in subfolder transcripts in current path, if transcripts subfolder is missed will be created
    
    .PARAMETER TranscriptFileNamePrefix
    Prefix used for creating transcript files name. Default is "Rollback-"
 
    .NOTES   
    
    
    AUTHOR: Wojciech Sciesinski, wojciech[at]sciesinski[dot]net

    KEYWORDS: PowerShell, Exchange, Active Directory, SMTP

    VERSION HISTORY
    0.1.0 - Initial release - untested !

	
	LICENSE
	Copyright (C) 2015 Wojciech Sciesinski
	This program is free software: you can redistribute it and/or modify
	it under the terms of the GNU General Public License as published by
	the Free Software Foundation, either version 3 of the License, or
	(at your option) any later version.
	This program is distributed in the hope that it will be useful,
	but WITHOUT ANY WARRANTY; without even the implied warranty of
	MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
	GNU General Public License for more details.
	You should have received a copy of the GNU General Public License
	along with this program. If not, see <http://www.gnu.org/licenses/>
    
         
    TO DO

    
    .EXAMPLE


    
    
#> 

[CmdletBinding()] 

param (

    [parameter(Mandatory=$true)]
	[alias("Path")]
    [String]$InputFilePath,
    
    [parameter(Mandatory=$false)]
    [Bool]$VerifyInputFileForDuplicates=$true,

    [parameter(Mandatory=$false)]
    [ValidateSet("DisplayOnly","PerformActions", "CreatePerformActionsCommandsOnly", "Rollback", "CreateRollbackCommnadsOnly" )]
    [String]$Mode="DisplayOnly",
    
    [parameter(Mandatory=$false)]
    [ValidateSet("UserMailbox","MailNonUniversalGroup","MailUniversalDistributionGroup")]
    [String]$RecipientType="UserMailbox",
	
	[parameter(Mandatory=$false)]
	[String]$DomainName,
    
    [parameter(mandatory=$false)]
    [String]$Prefix="smtp:",
    
    [parameter(Mandatory=$false)]
    [Bool]$CreateRollbackFile=$true,

    [parameter(Mandatory=$false)]
    [String]$RollBackFileDirectoryPath=".\rollbacks\",

    [parameter(Mandatory=$false)]
    [String]$RollBackFileNamePrefix="Rollback-",

    [parameter(Mandatory=$false)]
    [Bool]$CreateTranscript=$true,

    [parameter(Mandatory=$false)]
    [String]$TranscriptFileDirectoryPath=".\transcripts\",

    [parameter(Mandatory=$false)]
    [String]$TranscriptFileNamePrefix="Transcript-"

)

BEGIN {

	[String]$StartTime = Get-Date -format yyyyMMdd-HHmm

    If ( $CreateTranscript ) {
    
        Start-NewTranscript -TranscriptFileNamePrefix "Transcript-Set-PrimarySMTPAddress-" -StartTimeSuffix $StartTime

    }

}
	


}