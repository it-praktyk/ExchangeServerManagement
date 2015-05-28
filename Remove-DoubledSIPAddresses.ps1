function Remove-DoubledSIPAddresses {
	
<#
	.SYNOPSIS
	Function intended for veryfying and removing doubled SIP addresses for any mailbox in environment
    
	.DESCRIPTION 
	
	.PARAMETER CorrectSIPDomain
	Name of domain for which correct SIPs belong
	
	.PARAMETER CreateLogFile
	By default log file is created
	
	.PARAMETER LogFileDirectoryPath
	By default log files are stored in subfolder logs in current path, if logs subfolder is missed will be created
	
	.PARAMETER LogFileNamePrefix
	Prefix used for creating rollback files name. Default is "SIPs-Removed-"
          
	.LINK
	https://github.com/it-praktyk/Remove-DoubledSIPAddresses
	
	.LINK
	https://www.linkedin.com/in/sciesinskiwojciech
	      
	.NOTES
	AUTHOR: Wojciech Sciesinski, wojciech[at]sciesinski[dot]net
	KEYWORDS: PowerShell, Exchange, SIPAddresses, ProxyAddresses, Lync, migration
   
	VERSIONS HISTORY
	0.1.0 - 2015-05-27 - Initial release


	TODO
    - check if Exchange cmdlets are available

		
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
   
#>
	
	[CmdletBinding()]
	Param (
		
		[Parameter(Mandatory = $true, Position = 0)]
		[String]$CorrectSIPDomain,
	
		[parameter(Mandatory = $false)]
		[Bool]$CreateLogFile = $true,
	
		[parameter(Mandatory = $false)]
		[String]$LogFileDirectoryPath = ".\logs\",
	
		[parameter(Mandatory = $false)]
		[String]$LogFileNamePrefix = "SIPs-Removed-",
		
		[parameter(Mandatory = $false)]
		[Bool]$DisplayProgressBar = $false
		
	
		
	)
	
	
	BEGIN {
		
		[String]$StartTime = Get-Date -format yyyyMMdd-HHmm
		
		$Results = @()
		
		[String]$MessageText = "Data about mailboxes are read from Active Directory - please wait"
		
		Write-Verbose -Message $MessageText
		
		$Mailboxes = Get-Mailbox -ResultSize Unlimited | Select -Property Alias, DisplayName, RecipientType, EmailAddresses, Guid
		
		$MailboxesCount = ($Mailboxes | measure).Count
		
		$i = 1
		
	}
	
	PROCESS {
		
		$Mailboxes | ForEach  {
			
			If ($DisplayProgressBar) {
				
				$PercentCompleted = [math]::Round(($i / $MailboxesCount) * 100)
				
				$StatusText = "Percent completed $PercentCompleted%, currently the recipient {0} is checked. " -f $($_).ToString()

				Write-Progress -Activity "Checking SIP addresses" -Status $StatusText -PercentComplete $PercentCompleted
				
			}
			
			$CurrentRecipient = $_
			
			[String]$MessageText = "Currently addresses for {0} are checked" -f $CurrentRecipient.DisplayName
			
			Write-Verbose -Message $MessageText
			
			$CurrentRecipientSIPAddresses = ($CurrentRecipient | select -ExpandProperty EmailAddresses | where { $_.prefix -match 'SIP' })
			
			$CurrentRecipientSIPAddressesCount = ($CurrentRecipientSIPAddresses | Measure-Object).Count
			
			if ($CurrentRecipientSIPAddressesCount -gt 1) {
				
				$AddToLog = $false
				
				
					
					[String]$MessageText = "Mailbox with identifier {0} resolved to {1} has assigned {2} SIP addresses." `
					-f $CurrentRecipientIdentifier, $CurrentRecipient.DisplayName, $CurrentRecipientSIPAddressesCount
					
					Write-Verbose -Message $MessageText

				
				$Result = New-Object PSObject
				
				$Result | Add-Member -type 'NoteProperty' -name MailboxAlias -value $CurrentRecipient.Alias
				
				$Result | Add-Member -type 'NoteProperty' -name MailboxDisplayName -value $CurrentRecipient.DisplayName
				
				$Result | Add-Member -Type 'NoteProperty' -Name MailboxGuid -Value $CurrentRecipient.Guid
				
				$Result | Add-Member -Type 'NoteProperty' -Name SIPAddressesBeforeCount -Value $CurrentRecipientSIPAddressesCount
				
				[String]$CurrentSIPAddressesList = [string]::Join(",", $($CurrentRecipientSIPAddresses | ForEach { $_.ProxyAddressString }))
				
				$Result | Add-Member -Type 'NoteProperty' -Name SIPAddressesBeforeList -Value $CurrentSIPAddressesList
				
				$CurrentRecipientSIPAddresses | foreach {
					
					$CurrentSIP = $_.AddressString
					
					$AtPosition = $CurrentSIP.IndexOf("@")
					
					$SIPAddressLenght = $CurrentSIP.Length
					
					[String]$CurrentSIPDomain = $CurrentSIP.Substring($AtPosition + 1, $SIPAddressLenght - ($AtPosition + 1))
					
					[String]$MessageText = [String]$MessageText = "SIP address {0} is incorrect and will be deleted" `
					-f $CurrentSIP
					
					
					If ($CurrentSIPDomain -ne $CorrectSIPDomain) {
						
						if ($CurrentRecipient.IsPrimaryAddress -eq $true) {
							
							$CurrentSIP.ToSecondary()
							
						}						
						
						$SIPToRemove = $_.ProxyAddressString
						
						Write-Verbose -Message $MessageText
						
						Set-Mailbox -Identity $CurrentRecipient.Alias -EmailAddresses @{ remove = $SIPToRemove } -ErrorAction Continue
						
						$AddToLog = $true
						
					}
					
					
				}
				
				$CurrentRecipientSIPAddressesAfter  = (Get-Mailbox Identity $($CurrentRecipient.Alias) | Select -ExpandProperty EmailAddresses | where { $_.prefix -match 'SIP' })
				
				$CurrentRecipientSIPAddressesCountAfter = ($CurrentRecipientSIPAddressesAfter | Measure-Object).Count
				
				If ($CurrentRecipientSIPAddressesCountAfter -gt 1) {
					
					[String]$CurrentSIPAddressesListAfter = [string]::Join(",", $($CurrentRecipientSIPAddresses | ForEach { $_.ProxyAddressString }))
					
				}
				
				Else {
					
					$CurrentSIPAddressesListAfter = $CurrentSIPAddressesListAfter.ProxyAddressString
					
				}
				
				$Result | Add-Member -Type 'NoteProperty' -Name SIPAddressesAfterList -Value $CurrentSIPAddressesListAfter
				
				If ($AddToLog) {
					
					$Results += $Result
					
				}
				
			}
			
			$i++
			
		}
		
	}
	
	
	End {
		
		If ($CreateLogFile) {
			
			#Check if rollback directory exist and try create if not
			If (!$((Get-Item -Path $LogFileDirectoryPath -ErrorAction SilentlyContinue) -is [system.io.directoryinfo])) {
				
				New-Item -Path $LogFileDirectoryPath -Type Directory -ErrorAction Stop | Out-Null
				
			}
			
			$FullLogFilePath = $LogFileDirectoryPath + $LogFileNamePrefix + $StartTime + '.csv'
			
			Write-Verbose "Write rollback data to file $FullLogFilePath"
			
			#If export will not be unsuccessfull than display $Results to screen as the list - will be catched by Transcript
			
			Try {
				
				$Results | Export-CSV -Path $FullLogFilePath -NoTypeInformation -Delimiter ";" -Encoding UTF8 -ErrorAction SilentlyContinue
				
			}
			
			Catch {
				
				Return $Result
				
			}
			
		}
		
		
		
	}
	
	
}