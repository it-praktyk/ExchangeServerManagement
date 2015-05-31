function Remove-DoubledSIPAddresses {
    
<#
	.SYNOPSIS
	Function intended for verifying and removing doubled SIP addresses from all mailboxes in Exchange Server environment
    
	.DESCRIPTION 
    Function intended for verifying and removing doubled SIP addresses from all mailboxes in Exchange Server environment. `
    Only address in a domain provided in a parameter CorrectSIPDomain will be keep.
	
	.PARAMETER CorrectSIPDomain
	Name of domain for which correct SIPs belong
	
	.PARAMETER CreateLogFile
	By default log file is created
	
	.PARAMETER LogFileDirectoryPath
	By default log files are stored in the subfolder "logs" in current path, if the "logs" subfolder is missed will be created.
	
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
	0.1.0 - 2015-05-27 - First version published on GitHub
	0.1.2 - 2015-05-29 - Switch address to secondary befor remove, post-report corrected


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
			
			$CurrentMailbox = $_
			
			[String]$MessageText = "Currently addresses for {0} are checked" -f $CurrentMailbox.DisplayName
			
			Write-Verbose -Message $MessageText
			
			$CurrentMailboxSIPAddresses = ($CurrentMailbox | select -ExpandProperty EmailAddresses | where { $_.prefix -match 'SIP' })
			
			$CurrentMailboxSIPAddressesCount = ($CurrentMailboxSIPAddresses | Measure-Object).Count
			
			if ($CurrentMailboxSIPAddressesCount -gt 1) {
				
				$AddToLog = $false
				
				[String]$MessageText = "Mailbox with identifier {0} resolved to {1} has assigned {2} SIP addresses." `
				-f $CurrentMailboxIdentifier, $CurrentMailbox.DisplayName, $CurrentMailboxSIPAddressesCount
				
				Write-Verbose -Message $MessageText

				$Result = New-Object PSObject
				
				$Result | Add-Member -type 'NoteProperty' -name MailboxAlias -value $CurrentMailbox.Alias
				
				$Result | Add-Member -type 'NoteProperty' -name MailboxDisplayName -value $CurrentMailbox.DisplayName
				
				$Result | Add-Member -Type 'NoteProperty' -Name MailboxGuid -Value $CurrentMailbox.Guid
				
				$Result | Add-Member -Type 'NoteProperty' -Name SIPAddressesBeforeCount -Value $CurrentMailboxSIPAddressesCount
				
				[String]$CurrentSIPAddressesList = [string]::Join(",", $($CurrentMailboxSIPAddresses | ForEach { $_.ProxyAddressString }))
				
				$Result | Add-Member -Type 'NoteProperty' -Name SIPAddressesBeforeList -Value $CurrentSIPAddressesList
				
				$CurrentMailboxSIPAddresses | foreach {
					
					$CurrentSIPObject = $_
					
					[String]$CurrentSIPAddressString = $_.AddressString
					
					$AtPosition = $CurrentSIPAddressString.IndexOf("@")
					
					$SIPAddressLenght = $CurrentSIPAddressString.Length
					
					[String]$CurrentSIPDomain = $CurrentSIPAddressString.Substring($AtPosition + 1, $SIPAddressLenght - ($AtPosition + 1))
					
					If ($CurrentSIPDomain -ne $CorrectSIPDomain) {
						
						if ($CurrentSIPObject.IsPrimaryAddress -eq $true) {
							
							$CurrentSIPObject.ToSecondary() | Out-Null
							
						}
						
						$SIPToRemove = $CurrentSIPObject.ProxyAddressString
						
						[String]$MessageText = [String]$MessageText = "SIP address {0} is incorrect and will be deleted" `
						-f $CurrentSIPAddressString
						
						Write-Verbose -Message $MessageText
						
						Set-Mailbox -Identity $CurrentMailbox.Alias -EmailAddresses @{ remove = $SIPToRemove } -ErrorAction Continue
						
						$AddToLog = $true					
						
						
					}
				
					
				}
				
				$CurrentMailboxSIPAddressesAfter = (Get-Mailbox -Identity $CurrentMailbox.Alias | select -ExpandProperty EmailAddresses | where { $_.prefix -match 'SIP' })
				
				$CurrentMailboxSIPAddressesCountAfter = ($CurrentMailboxSIPAddressesAfter | Measure-Object).Count
				
				If ($CurrentMailboxSIPAddressesCountAfter -gt 1) {
					
					[String]$CurrentSIPAddressesListAfter = [string]::Join(",", $($CurrentMailboxSIPAddressesAfter | ForEach {
						
						$_.ProxyAddressString
					}))
					
				}
				
				Else {
					
					$CurrentSIPAddressesListAfter = $CurrentMailboxSIPAddressesAfter.ProxyAddressString
					
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
				
				If (($Resulst | measure).Count -lt 1) {
					
					
					$Results | Export-CSV -Path $FullLogFilePath -NoTypeInformation -Delimiter ";" -Encoding UTF8 -ErrorAction SilentlyContinue
					
				}
				Else {
					
					$Result = New-Object PSObject
					
					$Result | Add-Member -type 'NoteProperty' -name Message -value "Nothing has not changed - no doubled SIPs found."
										
					$Results | Export-CSV -Path $FullLogFilePath -NoTypeInformation -Delimiter ";" -Encoding UTF8 -ErrorAction SilentlyContinue
					
				}
				
				
				
			}
			
			Catch {
				
				Return $Result
				
			}
			
		}
		
	}
	
}