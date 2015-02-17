Function Test-EmailAddress {

<#
	.SYNOPSIS
		Function is intended to verify the correctness of addresses email in Microsoft Exchange Enviroment
		
	.DESCRIPTION
		Function which can be used to verifing an email address before for example adding it to Microosft Exchange environment. 
		Checks perfomed: 
		a) if email address contain wrong characters e.g. % or spaces
		b) if email address is from domain which are on accepted domains list
		c) if email address is currently assigned to any object in Exchange environment
		As a result function returns true/false if a provided email address is correct/incorrect
		
	.PARAMETER EmailAddress

	.EXAMPLE
		Test-EmailAddress -EmailAddress dummy@example.com
	
	.LINK
		https://github.com/it-praktyk/Test-EmailAddress
	.LINK
		https://www.linkedin.com/in/sciesinskiwojciech
	.NOTES
		AUTHOR: Wojciech Sciesinski, wojciech[at]sciesinski[dot]net
		KEYWORDS: Windows, PowerShell, Exchange Server, email
		BASE REPOSITORY: https://github.com/it-praktyk/Test-EmailAddress
		VERSION HISTORY
		0.1.0 - 2015-02-13 - first draft
		0.2.0 - 2015-02-16 - first working version
		0.2.1 - 2-15-02-17 - minor updates, first version published to GitHub
		

		TODO
		- veryfing if Exchange cmdlets are available
		- add parameters to disable some checks
		

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
	[String]$EmailAddress
	
)

	$AtPosition=$EmailAddress.IndexOf("@")
	
	$EmailAddressLenght = $EmailAddress.Length
	
	If ( $AtPosition -eq -1 ) {
	
		Write-Verbose "Email address $EmailAddress is not correct"
		
		Return $false
	}
	Else { 
	
		Write-Verbose "Provided email address is $EmailAddress."

		$Domain = $EmailAddress.Substring($AtPosition+1, $EmailAddressLenght - ( $AtPosition +1 ))
		
		Write-Verbose "Email address is from domain $Domain"
		
		If ( (Get-AcceptedDomain | where { $_.domainname -eq $Domain } | measure).count -eq 1) {
		
			Write-verbose -Message "Domain from $EmailAddress found in accepted domains."
	
			$EmailRegex = '^[_a-z0-9-]+(\.[_a-z0-9-]+)*@[a-z0-9-]+(\.[a-z0-9-]+)*(\.[a-z]{2,4})$'
			
			If ( ([regex]::Match($EmailAddress, $EmailRegex, "IgnoreCase ")).Success ) {
			
				$NotError = $true
			
				Write-Verbose -Message "Email address  $EmailAddress  doesn't contain any unsupported chars"
				
				Try {
					
					$Recipients = Get-Recipient $EmailAddress -ErrorAction Stop
					
				}
				
				Catch {
					
					Write-Verbose -Message "Email address doesn't exist in environment - finally result: is correct"
					
					$NotError = $false
					
					Return $true
				
				}
				
				Finally {
	
					If ( $NotError ) {
					
						Write-Verbose -Message "Recipient with email address $EmailAddress exist now."
				
						Return $false
						
					}
				
				}
				
			}
			Else {
		
				Write-Verbose -Message "Email address $EmailAddress contain unsupported chars"
				
				Return $false
		
			}
		}
	
		Else {
		
			Write-Verbose "Email address $EmailAddress is not from accepted domains."
		
			Return $false

		}

	}
	
}