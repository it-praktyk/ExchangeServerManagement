Function Manage-MailboxAddresses
{
	
<#   
    .SYNOPSIS   
    Function intended to manage addresses for mailboxes in Exchange environment
  
    .DESCRIPTION   
    Function intended to manage addresses for mailboxes in Exchange Server environment based on input from csv file.
	Log and rolback file can be created.
	    
    .PARAMETER InputFilePath
    The path for file with input data
    
    .PARAMETER Mode
    The switch which define action to perform - default mode is DisplayOnly

    Available modes
    - DisplayOnly
    - PerformActions
    - CreatePerformActionsCommandsOnly
    - Rollback
    - CreateRollbackCommnadsOnly
    
    .PARAMETER Operation
    Operation which need to performed on maibox.
    
    Available operations
    - AddProxyAddress 
    - RemoveProxyAddress
    - SetSMTPPrimaryAddress
	
	.PARAMETER Recipienttype
	

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
	
	.PARAMETER CreateErrorsReportFile
	By default error report file is created
	
	.PARAMETER ErrorsReportFileDirectoryPath
	By default error report files are stored in subfolder "errors" in current path, if "errors" subfolder is missed will be created
	
	.PARAMETER ErrorsReportFileNamePrefix
	Prefix used for creating errors report files name. Default is "Errors-"	
 
    .NOTES   
    
    
    AUTHOR: Wojciech Sciesinski, wojciech[at]sciesinski[dot]net

    KEYWORDS: PowerShell, Exchnange, Active Directory, SMTP

    VERSION HISTORY
    0.1.0 - 2015-03-12 - Initial release - untested !
    0.2.0 - 2015-03-13 - Tested, corrected
    0.2.1 - 2015-03-14 - Info about license added - GNU GPLv3
    0.3.0 - 2015-03-15 - Mode set extended, Operation parameter added
    0.4.0 - 2015-03-17 - Operations partially implemented
	0.5.0 - 2015-03-17 - Function renamed from Set-PrimarySMTPAddress to Manage-MailboxAddresses, help updated
	0.5.1 - 2015-03-18 - Help updated, parameters: $VerifyInputFileForDuplicates, FQDNDomainName, prefix removed
	0.6.0 - 2015-03-25 - Errors can be saved to separate file
	0.6.1 - 2015-03-30 - Minor update for initial checks section
	0.6.2 - 2015-03-30 - Skip parameter in input file implemented
    
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
    Add checking if mailbox/recipient has enabled 'useemailaddresspolicy'
    Skipping for objects need to be implemented
    
   
    
#>
	
	[CmdletBinding()]
	
	param (
		
		[parameter(Mandatory = $true)]
		[String]$InputFilePath,
		
		[parameter(Mandatory = $true, `
				   HelpMessage = "Available modes: DisplayOnly, PerformActions, CreatePerformActionsCommandsOnly, Rollback, CreateRollbackCommnadsOnly")]
		[ValidateSet("DisplayOnly", "PerformActions", "CreatePerformActionsCommandsOnly", "Rollback", "CreateRollbackCommnadsOnly")]
		[String]$Mode,
		
		[parameter(Mandatory = $true, `
				   HelpMessage = "Available operations: AddProxyAddress, RemoveProxyAddress, SetSMTPPrimaryAddress")]
		[ValidateSet("AddProxyAddress", "RemoveProxyAddress", "SetSMTPPrimaryAddress")]
		[String]$Operation,
		
		[parameter(Mandatory = $true, `
				   HelpMessage = "Available recipients types: UserMailbox")]
		[ValidateSet("UserMailbox")]
		[String]$RecipientType,
		
		[parameter(Mandatory = $false)]
		[Bool]$CreateRollbackFile = $true,
		
		[parameter(Mandatory = $false)]
		[String]$RollBackFileDirectoryPath = ".\rollbacks\",
		
		[parameter(Mandatory = $false)]
		[String]$RollBackFileNamePrefix = "Rollback-",
		
		[parameter(Mandatory = $false)]
		[Bool]$CreateTranscript = $true,
		
		[parameter(Mandatory = $false)]
		[String]$TranscriptFileDirectoryPath = ".\transcripts\",
		
		[parameter(Mandatory = $false)]
		[String]$TranscriptFileNamePrefix = "Transcript-",
		
		[parameter(Mandatory = $false)]
		[Bool]$CreateErrorsReportFile = $true,
		
		[parameter(Mandatory = $false)]
		[String]$ErrorsReportFileDirectoryPath = ".\errors\",
		
		[parameter(Mandatory = $false)]
		[String]$ErrorsReportFileNamePrefix = "Errors-"
		
		
	)
	
	BEGIN
	{
		
		#Uncomments if you need hunt any bug
		Set-StrictMode -version 2
		
		[String]$StartTime = Get-Date -format yyyyMMdd-HHmm
		
		If ($CreateTranscript)
		{
			
			Start-NewTranscript -TranscriptFileNamePrefix "Transcript-Set-PrimarySMTPAddress-" -StartTimeSuffix $StartTime
			
		}
		
		If (Test-Path -Path $InputFilePath)
		{
			
			If ((Get-Item -Path $InputFilePath) -is [System.IO.fileinfo])
			{
				
				try
				{
					
					$RecipientsFromInputFile = (Import-CSV -Path $InputFilePath -Delimiter ";" -ErrorAction Stop | Where { $_.SkipRecipientCode -eq 0 -or $_.SkipRecipientCode -eq $false })
					
					[Int]$RecipientsCount = $($RecipientsFromInputFile | Measure-Object).count
					
				}
				catch
				{
					
					Write-Error "Read input file $InputFilePath error "
					
					Stop-Transcript -ErrorAction SilentlyContinue | Out-Null
					
					break
					
				}
				
			}
			
			Else
			{
				
				Write-Error "Provided value for InputFilePath is not a file"
				
				Stop-Transcript -ErrorAction SilentlyContinue | Out-Null
				
				break
				
			}
			
		}
		Else
		{
			
			Write-Error "Provided value for InputFilePath doesn't exist"
			
			Stop-Transcript -ErorAction SilentlyContinue
			
			break
		}
		
		If ($CreateErrorsReportFile)
		{
			
			#Check if rollback directory exist and try create if not
			If (!$((Get-Item -Path $ErrorsReportFileDirectoryPath -ErrorAction SilentlyContinue) -is [system.io.directoryinfo]))
			{
				
				New-Item -Path $ErrorsReportFileDirectoryPath -type Directory -ErrorAction Stop | Out-Null
				
			}
			
			$ErrorReportFilePath = $ErrorsReportFileDirectoryPath + $ErrorsReportFileNamePrefix + $StartTime + '.csv'
			
			Write-Verbose "Error report data will be written to $ErrorReportFilePath"
			
		}
		
		#Declare variable for store results data
		$Results = @()
		
		#Declare variable for store errors data
		$ErrorResults = @()
		
		[int]$i = 1
		
		[Array]$AcceptedRecipientTypes = @("UserMailbox")
		
	}
	
	PROCESS
	{
		
		$RecipientsFromInputFile | ForEach {
			
			$PercentCompleted = [math]::Round(($i / $RecipientsCount) * 100)
			
			$StatusText = "Percent completed $PercentCompleted%, currently the recipient {0} is checked. " -f $($_.RecipientIdentity).ToString()
			
			Write-Progress -Activity "Performing action in mode $Mode" -Status $StatusText -PercentComplete $PercentCompleted
			
			[String]$MessageText = "Performing check on object {0}  in mode: {1} ." -f $_.RecipientIdentity, $Mode
			
			Write-Verbose -Message $MessageText
			
			If (@("RemoveProxyAddress", "SetSMTPPrimaryAddress") -contains $Operation)
			{
				
				Try
				{
					
					$SelectedRecipientTest1 = $(Get-Recipient $_.RecipientIdentity -ErrorAction Stop | Where { $_.RecipientType -eq $RecipientType })
					
					$SelectedRecipientTest1Count = (Measure-Object -InputObject $SelectedRecipientTest1).Count
					
					Write-Debug "First test for recipient result: $SelectedRecipientTest1"
					
					$SelectedRecipientTest2 = $(Get-Recipient $_.NewPrimarySMTPAddress -ErrorAction Stop | Where { $_.RecipientType -eq $RecipientType })
					
					$SelectedRecipientTest2Count = (Measure-Object -InputObject $SelectedRecipientTest2).Count
					
					If (($SelectedRecipientTest1Count + $SelectedRecipientTest2Count) -ne 2)
					{
						
						[String]$MessageText = "Recipient  {0} is not {1} or email address {2} is not currently assigned to any recipient with type {1}" -f $_.RecipientIdentity, $_.RecipientType, $_.NewPrimarySMTPAddress
						
						Write-Error -Message $MessageText -ErrorAction Continue
						
						Break
						
					}
					
					Write-Debug "Second test for recipient result: $SelectedRecipientTest2"
					
					If ($SelectedRecipientTest1.Guid -ne $SelectedRecipientTest2.Guid)
					{
						
						Write-Error -Message "Email address $_.NewPrimarySMTPAddress is not currently assigned to recipient $_.RecipientIdentity with type $_.RecipientType"
						
						Break
					}
					
					Else
					{
						
						$SelectedRecipient = $SelectedRecipientTest1
						
					}
					
				}
				
				Catch
				{
					
					Write-Error "Recipient $($_).RecipientIdentity or with address $_.NewPrimarySMTPAddress doesn't exist"
					
					Break
					
				}
				
			}
			Elseif (@("AddProxyAddress") -contains $Operation)
			{
				
				Try
				{
					
					$SelectedRecipient = $(Get-Recipient $_.RecipientIdentity -ErrorAction Stop | Where { $_.RecipientType -eq $RecipientType })
					
					$EmailTestResult = Test-EmailAddress -EmailAddress $_.NewProxyAddress
					
					If ($EmailTestResult.ExitCode -ne 0)
					{
						
						$ErrorResults += $EmailTestResult
						
						[String]$MessageText = "Email address {0} is not correct. Error code: {1}, Error description: {2}, Conflicted object {3} " `
						-f $_.ProxyAddresses, $EmailTestResult.ExitCode, $EmailTestResult.ExitDescription, $EmailTestResult.ConflictedObjectAlias
						
						Write-Output $MessageText -ForegroundColor red
						
					}
					
				}
				
				Catch
				{
					
					Write-Error "Mailbox $_.RecipientName doesn't exist"
					
					Break
					
				}
				
				
			}
			
			if ($AcceptedRecipientTypes -notcontains $SelectedRecipientTest1.Recipienttype)
			{
				
				Write-Error -Message "This function can only process recipients with type UserMailbox - for Recipient $_.RecipientIdentity type is $_.RecipientIdentity"
				
				Break
				
			}
			
			
			$CurrentRecipient = Get-Mailbox -Identity $($SelectedRecipient.Alias)
			
			Write-Verbose -Message "Performing action on $CurrentRecipient.Alias in mode $Mode ."
			
			#Object properties before any changes - common part of Result objects
			
			$Result = New-Object PSObject
			
			$Result | Add-Member -MemberType NoteProperty -Name RecipientIdentity -Value $_.RecipientIdentity
			
			$Result | Add-Member -MemberType NoteProperty -name RecipientType -value $_.RecipientType
			
			$Result | Add-Member -MemberType NoteProperty -Name RecipientGuid -Value $CurrentRecipient.Guid
			
			$Result | Add-Member -MemberType NoteProperty -name RecipientAlias -value $CurrentRecipient.Alias
			
			$Result | Add-Member -MemberType NoteProperty -Name PrimarySMTPAddressBefore -Value $CurrentRecipient.PrimarySMTPAddress
			
			$AllProxyAddressesStringBefore = (@(select-Object -InputObject $CurrentRecipient -expandproperty emailaddresses) -join ',')
			
			$Result | Add-Member -MemberType NoteProperty -name ProxyAddressesBefore -value $AllProxyAddressesStringBefore
			
			if ($Operation -eq 'AddProxyAddress')
			{
				
				If ($Mode -eq 'DisplayOnly')
				{
					
					[String]$ProxyAddressStringToAdd = "{0}{1}" -f $Prefix, $_.NewProxyAddress
					
					[String]$ProxyAddressStringProposal = "{0},{1}" -f $AllProxyAddressesStringBefore, $ProxyAddressStringToAdd
					
					$Result | Add-Member -MemberType NoteProperty -name ProxyAddressesProposal -value $ProxyAddressStringProposal
					
					$Result | Add-Member -type NoteProperty -name ProxyAddressesAfter -value $AllProxyAddressesStringBefore
					
				}
				
				Elseif ($Mode -eq 'PerformActions')
				{
					
					Set-Mailbox -Identity $CurrentRecipient -EmailAddresses @{ add = ($ProxyAddressStringToAdd) } -ErrorAction Continue
					
					$CurrentRecipientAfter = Get-Mailbox -Identity $($SelectedRecipient.Alias)
					
					$AllProxyAddressesStringAfter = (@(select-Object -InputObject $CurrentRecipientAfter -ExpandProperty emailaddresses) -join ',')
					
					$Result | Add-Member -type NoteProperty -name ProxyAddressesAfter -value $AllProxyAddressesStringAfter
					
				}
				
				ElseIf ($Mode -eq 'Rollback')
				{
					
					Write-Error -Message "Rollback mode is not implemented yet"
					
				}
				
			}
			
			elseif ($Operation -eq 'SetSMTPPrimaryAddress')
			{
				
				$Result | Add-Member -MemberType NoteProperty -Name PrimarySMTPAddressProposal -Value $_.NewPrimarySMTPAddress
				
				If ($Mode -eq 'DisplayOnly')
				{
					
					$Result | Add-Member -MemberType NoteProperty -Name PrimarySMTPAddresAfter -Value $CurrentRecipient.PrimarySMTPAddress
					
					$Result | Add-Member -MemberType NoteProperty -name ProxyAddressesAfter -value $AllProxyAddressesStringBefore #This need to be changed - replace for
					
				}
				
				Elseif ($Mode -eq 'PerformActions')
				{
					
					Set-Mailbox -Identity $CurrentRecipient -PrimarySMTPAddress $_.NewPrimarySMTPAddress -ErrorAction Continue
					
					$CurrentRecipientAfter = Get-Mailbox $CurrentRecipient
					
					$Result | Add-Member -MemberType NoteProperty -Name PrimarySMTPAddressAfter -Value $CurrentRecipientAfter.PrimarySMTPAddress
					
					$AllProxyAddressesStringAfter = (@(select-Object -InputObject $CurrentRecipientAfter -ExpandProperty emailaddresses) -join ',')
					
					$Result | Add-Member -MemberType NoteProperty -name ProxyAddressesAfter -value $AllProxyAddressesStringAfter
					
				}
				
				ElseIf ($Mode -eq 'Rollback')
				{
					
					Write-Error -Message "Rollback mode is not implemented yet"
					
				}
				
			}
			
			elseif ($Operation -eq 'RemoveProxyAddress')
			{
				
				
			}
			
			$Results += $Result
			
			$i += 1
			
		}
		
	}
	
	
	END
	{
		
		#Save results to rollback file - need to be moved to external function
		
		If ($CreateRollbackFile)
		{
			
			#Check if rollback directory exist and try create if not
			If (!$((Get-Item -Path $RollBackFileDirectoryPath -ErrorAction SilentlyContinue) -is [system.io.directoryinfo]))
			{
				
				New-Item -Path $RollBackFileDirectoryPath -Type Directory -ErrorAction Stop | Out-Null
				
			}
			
			$FullRollbackFilePath = $RollBackFileDirectoryPath + $RollBackFileNamePrefix + $StartTime + '.csv'
			
			Write-Verbose "Rollback data will be written to $FullRollbackFilePath"
			
			Write-Verbose "Write rollback data to file $FullRollbackFilePath"
			
			#If export will not be unsuccessfull than display $Results to screen as the list - will be catched by Transcript
			Try
			{
				
				$Results | Export-CSV -Path $FullRollbackFilePath -NoTypeInformation -Delimiter ";" -Encoding UTF8 -ErrorAction Continue
				
			}
			
			Catch
			{
				
				If ($CreateTranscript)
				{
					
					$Results | Format-List
					
				}
				Else
				{
					
					Start-NewTranscript -TranscriptFileDirectoryPath ".\emergency-transcripts\" -TranscriptFileNamePrefix "Emergency-Transcript-"
					
				}
				
			}
			
		}
		
		#Display results to console - also can be redirected to file
		Else
		{
			
			Return $Results
			
		}
		
		#Save errors to errors report file
		If ($CreateErrorsReportFile)
		{
			
			Write-Verbose "Write errors data to file $ErrorReportFilePath"
			
			$ErrorResults | Export-CSV -Path $ErrorReportFilePath -NoTypeInformation -Delimiter ";" -Encoding UTF8 -ErrorAction Continue
			
		}
		
		#Display errors to console - also can be redirected to file
		Else
		{
			
			Return $ErrorResults
			
		}
		
		Stop-Transcript -ErrorAction SilentlyContinue | Out-Null
		
	}
}





Function Start-NewTranscript
{
<#
    .SYNOPSIS
    PowerShell function intended for start new transcript based on provided parameters
   
    .DESCRIPTION
    This function extend PowerShell transcript creation start. A transcript is created in the folder other than default with the name which can be defined as parameter,
    previous transcript is stopped if needed, etc.
    
      .PARAMETER TranscriptFileDirectoryPath
    By default transcript files are stored in subfolder named "transcripts" in current path, if transcripts subfolder is missed will be created
    
    .PARAMETER TranscriptFileNamePrefix
    Prefix used for creating transcript files name. Default is "Transcript-"
    
    .PARAMETER StartTimeSuffix
    Suffix what will be added to transcript file name
  
    .EXAMPLE
    
    Start-NewTranscript -TranscriptFileDirectoryPath "C:\Transcripts\" -TranscriptFileNamePrefix "Change_No_111_transcript-"
     
    .LINK
    https://github.com/it-praktyk/Start-NewTranscript
    
    .LINK
    https://www.linkedin.com/in/sciesinskiwojciech
          
    .NOTES
    AUTHOR: Wojciech Sciesinski, wojciech[at]sciesinski[dot]net
    KEYWORDS: PowerShell
   
    VERSIONS HISTORY
    0.1.0 - 2015-03-11 - Initial release
    0.2.0 - 2015-03-12 - additional parameter StartTimeSuffix added

    TODO
    - Check format used by hours for other cultures or formating like this '{0}' -f [datetime]::UtcNow
    - Catch situation when the "\" in path are doubled or missed
    - Suppress "Start transcript ... " message
            
    LICENSE
    This function is licensed under The MIT License (MIT)
    Full license text: http://opensource.org/licenses/MIT
    Copyright (c) 2015 Wojciech Sciesinski
    
    DISCLAIMER
    This script is provided AS IS without warranty of any kind. I disclaim all implied warranties including, without limitation,
    any implied warranties of merchantability or of fitness for a particular purpose. The entire risk arising out of the use or
    performance of the sample scripts and documentation remains with you. In no event shall I be liable for any damages whatsoever
    (including, without limitation, damages for loss of business profits, business interruption, loss of business information,
    or other pecuniary loss) arising out of the use of or inability to use the script or documentation. 
   
#>
	
	[CmdletBinding()]
	
	param (
		
		[parameter(Mandatory = $false)]
		[String]$TranscriptFileDirectoryPath = ".\transcripts\",
		
		[parameter(Mandatory = $false)]
		[String]$TranscriptFileNamePrefix = "Transcript-",
		
		[parameter(Mandatory = $false)]
		[String]$StartTimeSuffix
		
	)
	
	BEGIN
	{
		
		#Uncomments if you need hunt any bug
		Set-StrictMode -version 2
		
		If ($StartTimeSuffix)
		{
			
			[String]$StartTime = $StartTimeSuffix
			
		}
		Else
		{
			
			[String]$StartTime = Get-Date -format yyyyMMdd-HHmm
			
		}
		
		#Check if transcript directory exist and try create if not
		If (!$((Get-Item -Path $TranscriptFileDirectoryPath -ErrorAction SilentlyContinue) -is [system.io.directoryinfo]))
		{
			
			New-Item -Path $TranscriptFileDirectoryPath -Type Directory -ErrorAction Stop | Out-Null
			
			Write-Verbose -Message "Folder $TranscriptFileDirectoryPath was created."
			
		}
		
		$FullTranscriptFilePath = $TranscriptFileDirectoryPath + '\' + $TranscriptFileNamePrefix + $StartTime + '.log'
		
		#Stop previous PowerShell transcript and catch error if not started previous
		
		try
		{
			
			stop-transcript | Out-Null
			
		}
		
		catch [System.InvalidOperationException]{ }
		
	}
	
	PROCESS
	{
		
		#Start new PowerShell transcript
		
		Start-Transcript -Path $FullTranscriptFilePath -ErrorAction Stop | Out-Null
		
		Write-Verbose -Message "Transcript will be written to $FullTranscriptFilePath"
		
	}
	
	END
	{
		
	}
	
}