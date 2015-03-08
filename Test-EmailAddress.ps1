Function Test-EmailAddress {
<#
	.SYNOPSIS
		Function is intended to verify the correctness of addresses email in Microsoft Exchange Enviroment
		
	.DESCRIPTION
		Function which can be used to verifing an email address before for example adding it to Microosft Exchange environment. 
		Checks perfomed: 
		a) if email address contain wrong characters e.g. % or spaces
		b) if email address is from domain which are on accepted domains list
		c) if email address is currently assigned to any object in Exchange environment (a conflicted object exist)
		As the result returned is PowerShell object which contain: EmailAddress, ExitCode, ExitDescription, ConflictedObjectAlias, ConflictedObjectType.
		
		Exit codes and descriptions:
		0 - No Error
		1 - Email doesn't contain 'at' char
		2 - Email exist now
		3 - Unsupported chars found
		4 - Not accepted domain
		5 - White chars e.g. spaces founded before/after email
		
	.PARAMETER EmailAddress
		Email address which need to be verified in Exchange environment

	.EXAMPLE
		Test-EmailAddress -EmailAddress dummy@example.com
	
	.LINK
		https://github.com/it-praktyk/Test-EmailAddress
		
	.LINK
		https://www.linkedin.com/in/sciesinskiwojciech
		
	.NOTES
		AUTHOR: Wojciech Sciesinski, wojciech[at]sciesinski[dot]net
		KEYWORDS: Windows, PowerShell, Exchange Server, email
		VERSION HISTORY
		0.1.0 - 2015-02-13 - first draft
		0.2.0 - 2015-02-16 - first working version
		0.2.1 - 2-15-02-17 - minor updates, first version published on GitHub
		0.3.0 - 2015-02-18 - exit codes added, result returned as PowerShell object
		0.3.1 - 2015-02-18 - help updated, input parameater checks added
		0.3.2 - 2015-02-19 - corrected for work with PowerShell 4.0 also (Windows Server 2012 R2)
		0.3.3 - 2015-02-27 - ommited by mistake
		0.3.4 - 2015-02-27 - regex for email parsing updated
		0.3.5 - 2015-02-27 - chars like ' and # excluded from regex for parsing email address
		0.4.0 - 2015-03-07 - verifying if function is runned in EMS added
		0.5.0 - 2015-03-08 - verifying if email contains white chars (like a spaces) at the beginning or at the end added
		

		TODO
		- add parameters to disable some checks
		- add support for verifying emails from files directly 
	

		DISCLAIMER
		This script is provided AS IS without warranty of any kind. I disclaim all implied warranties including, without limitation,
		any implied warranties of merchantability or of fitness for a particular purpose. The entire risk arising out of the use or
		performance of the sample scripts and documentation remains with you. In no event shall I be liable for any damages whatsoever
		(including, without limitation, damages for loss of business profits, business interruption, loss of business information,
		or other pecuniary loss) arising out of the use of or inability to use the script or documentation. 
#>

[cmdletbinding()]

param (

	[parameter(mandatory=$true)]
	[ValidateNotNullOrEmpty()]
	[alias("email")]
	[String]$EmailAddress

)

BEGIN {

	#Declare variable for store results data
		
	$Result = New-Object PSObject

}

PROCESS {

	$AtPosition=$EmailAddress.IndexOf("@")

	
	If ( $AtPosition -eq -1 ) {
	
		Write-Verbose "Email address $EmailAddress is not correct - at char is missed."
		
		$Result | Add-Member -type NoteProperty -Name EmailAddress -value $EmailAddress
		$Result | Add-Member -type NoteProperty -Name ExitCode  -value 1
		$Result | Add-Member -type NoteProperty -Name ExitDescription -value "Email doesn't contain 'at' char"
		$Result | Add-Member -Type NoteProperty -Name ConflictedObjectAlias -value "Not checked"
		$Result | Add-Member -Type NoteProperty -Name ConflictedObjectType -value "Not checked"
				
		Return $Result
	}
	Else { 
	
		#This try/catch block check if Exchange commands are available
		Try {
		
			$AcceptedDomains = Get-AcceptedDomain 
			
		}
		
		Catch [System.Management.Automation.CommandNotFoundException] {
		
			Throw "This function need to be run using Exchange Management Shell."
		
		}
	
		Write-Verbose "Provided email address is $EmailAddress."
		
		If ( ($EmailAddress.Trim()).Length -ne $EmailAddress.Length ) {
		
			Write-Verbose -Message "Email address $EmailAddress contains white spaces at the beginning or at the end."
				
			$Result | Add-Member -type NoteProperty -Name EmailAddress -value $EmailAddress
			$Result | Add-Member -type NoteProperty -Name ExitCode  -value 5
			$Result | Add-Member -type NoteProperty -Name ExitDescription -value "White chars e.g. spaces founded before/after email"
			$Result | Add-Member -Type NoteProperty -Name ConflictedObjectAlias -value "Not checked"
			$Result | Add-Member -Type NoteProperty -Name ConflictedObjectType -value  "Not checked"
				
			Return $Result
		
		}
			
		$EmailAddressLenght = $EmailAddress.Length

		$Domain = $EmailAddress.Substring($AtPosition+1, $EmailAddressLenght - ( $AtPosition +1 ))
		
		Write-Verbose "Email address is from domain $Domain"	
		
		If ( ($AcceptedDomains | where { $_.domainname -eq $Domain } | measure).count -eq 1) {
		
			Write-verbose -Message "Domain from $EmailAddress found in accepted domains."
			
			$SpacePosition=$EmailAddress.IndexOf(" ")
			
			#Regex source http://www.regular-expressions.info/email.html
			$EmailRegex = '[a-z0-9!#$%&''*+/=?^_`{|}~-]+(?:\.[a-z0-9!#$%&''*+/=?^_`{|}~-]+)*@(?:[a-z0-9](?:[a-z0-9-]*[a-z0-9])?\.)+[a-z0-9](?:[a-z0-9-]*[a-z0-9])?'
			
			If ( ([regex]::Match($EmailAddress, $EmailRegex, "IgnoreCase ")).Success -and $SpacePosition -eq -1 ) {
			
				$NotError = $true
			
				Write-Verbose -Message "Email address  $EmailAddress  doesn't contain any unsupported chars"
				
				
				Try {
					
					$Recipient = Get-Recipient $EmailAddress -ErrorAction Stop
					
				}
								
				Catch [System.Management.Automation.RemoteException] {
					
					Write-Verbose -Message "Email address doesn't exist in environment - finally result: is correct"
					
					$NotError = $false
					
					$Result | Add-Member -type NoteProperty -Name EmailAddress -value $EmailAddress
					$Result | Add-Member -type NoteProperty -Name ExitCode  -value 0
					$Result | Add-Member -type NoteProperty -Name ExitDescription -value "No Error"
					$Result | Add-Member -Type NoteProperty -Name ConflictedObjectAlias -value "No conflict"
					$Result | Add-Member -Type NoteProperty -Name ConflictedObjectType -value "Not checked"
					
					Return $Result
				
				}
	
				If ( $NotError ) {
					
					Write-Verbose -Message "Recipient with email address $EmailAddress exist now."
						
					$Result | Add-Member -type NoteProperty -Name EmailAddress -value $EmailAddress
					$Result | Add-Member -type NoteProperty -Name ExitCode  -value 2
					$Result | Add-Member -type NoteProperty -Name ExitDescription -value "Email exist now"
					$Result | Add-Member -Type NoteProperty -Name ConflictedObjectAlias -value $Recipient.alias
					$Result | Add-Member -Type NoteProperty -Name ConflictedObjectType -value $Recipient.RecipientType
				
					Return $Result
						
				}
	
			}
			Else {
		
				Write-Verbose -Message "Email address $EmailAddress contain unsupported chars"
				
				$Result | Add-Member -type NoteProperty -Name EmailAddress -value $EmailAddress
				$Result | Add-Member -type NoteProperty -Name ExitCode  -value 3
				$Result | Add-Member -type NoteProperty -Name ExitDescription -value "Unsupported chars found"
				$Result | Add-Member -Type NoteProperty -Name ConflictedObjectAlias -value "Not checked"
				$Result | Add-Member -Type NoteProperty -Name ConflictedObjectType -value  "Not checked"
				
				Return $Result
		
			}
		}
	
		Else {
		
			Write-Verbose "Email address $EmailAddress is not from accepted domains."
			
			$Result | Add-Member -type NoteProperty -Name EmailAddress -value $EmailAddress
			$Result | Add-Member -type NoteProperty -Name ExitCode  -value 4
			$Result | Add-Member -type NoteProperty -Name ExitDescription -value "Not accepted domain"
			$Result | Add-Member -Type NoteProperty -Name ConflictedObjectAlias -value "Not checked"
			$Result | Add-Member -Type NoteProperty -Name ConflictedObjectType -value  "Not checked"
					
			Return $Result

		}

	}
	
}

END {
	
	#Nothing yet in this section

}

}