function Test-EmailAddress {
<#
    .SYNOPSIS
    Function is intended to verify the correctness of addresses email in Microsoft Exchange Enviroment
        
    .DESCRIPTION
    Function which can be used to verifing an email address before for example adding it to Microosft Exchange environment. 
    Checks perfomed: 
    a) if email address contain wrong characters e.g. spaces
    b) if email address is from domain which are on accepted domains list
    c) if email address is currently assigned to any object in Exchange environment (a conflicted object exist)
    
    .PARAMETER EmailAddress
    Email address which need to be verified in Exchange environment
    
    .PARAMETER TestEmailFormat
    
    .PARAMETER TestAcceptedDomains
    
    .PARAMETER TestIfExists
    
    .PARAMETER TestIsPrimary

    .PARAMETER AcceptedDomains

    .EXAMPLE
    Test-EmailAddress -EmailAddress dummy@example.com 
    
    .EXAMPLE
    Test-EmailAddress -EmailAddress "dummy@example.com","john@doe.com"
    
    .LINK
    https://github.com/it-praktyk/Test-EmailAddress
        
    .LINK
    https://www.linkedin.com/in/sciesinskiwojciech
        
    .NOTES
    AUTHOR: Wojciech Sciesinski, wojciech[at]sciesinski[dot]net
    
    KEYWORDS: Windows, PowerShell, Exchange Server, email

    VERSION HISTORY
    0.6.0 - 2015-12-22 - the function rewriten, information about license added
    0.7.0 - 2015-12-29 - validation extended and corrected
    
    TODO
    - implement parameter "$AcceptedDomains"
    - update help
    - add descriptive test result
    - implementing email validation using [Microsoft.Exchange.Data.SmtpProxyAddress]::Parse($EmailAddress).ParseException or similiar method
	- resolve PSScriptAnalyzer output "PSUseOutputTypeCorrectly - The cmdlet 'Test-EmailAddress' returns an object of type 'System.Object[]' but this type is not declared in the OutputType attribute."

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
    
    .OUTPUTS
    System.Object[]
        
#>  
    
    [cmdletbinding()]
    param (
        
        [parameter(ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [alias("email", "SmtpAddress")]
        [String[]]$EmailAddress,
        [parameter(Mandatory = $false)]
        [Bool]$TestEmailFormat = $true,
        [parameter(Mandatory = $false)]
        [Bool]$TestAcceptedDomains = $true,
        [parameter(Mandatory = $false)]
        [Bool]$TestIfExists = $true,
        [parameter(Mandatory = $false)]
        [Bool]$TestIsPrimary = $true,
        [parameter(Mandatory = $false)]
        [String[]]$AcceptedDomains
        
    )
    
    BEGIN {
        
        #region Declare variables                                                        
        
        $Results = @()
        
        $PowerShellVersion = ([version]$psversiontable.psversion).major
        
        #Workaroud to resolve issue with PowerShell 2.0 on W2K8R2
        if ($PowerShellVersion -gt 2 -and $TestAcceptedDomains) {
            
            #This try/catch block check if Exchange commands are available
            Try {
                
                $AcceptedDomains = $(Get-AcceptedDomain -verbose)
                
            }
            
            Catch [System.Management.Automation.CommandNotFoundException] {
                
                Write-Verbose -Message "Error occured $error[0]"
                
                $TestAcceptedDomains = $false
                
                $TestAcceptedDomainsResult = "SKIPPED"
                
            }
            
        }
        
        #endregion        
        
    }    
    
    PROCESS {
        
        #Workaroud to resolve issue with PowerShell 2.0 on W2K8R2
        If ($PowerShellVersion -eq 2 -and $TestAcceptedDomains) {
            
            #This try/catch block check if Exchange commands are available
            Try {
                
                $AcceptedDomains = $(Get-AcceptedDomain)
                
            }
            
            Catch [System.Management.Automation.CommandNotFoundException] {
                
                Write-Verbose -Message "Error occured $error[0]"
                
                $TestAcceptedDomains = $false
                
                $TestAcceptedDomainsResult = "SKIPPED"
                
            }
		
        }
                
        #region Main loop
        $EmailAddress | ForEach-Object -Process {
            
            $CurrentEmailAddress = $_
            
            $Result = New-Object PSObject
            
            #region Checking email pattern                                                                        
            
            if ($TestEmailFormat) {
                                
                #Check if white chars are on the begining/end
                if ($CurrentEmailAddress.Trim() -ne $CurrentEmailAddress) {
                    
                    $TestWhiteCharsResult = "FAIL"
                    
                    $CurrentEmailAddress = $CurrentEmailAddress.Trim()
                    
                }
                Else {
                    
                    $TestWhiteCharsResult = "PASS"
                    
                }
                
                #Check if space is in midle of email
                $SpacePosition = $CurrentEmailAddress.IndexOf(' ')
                
                
                #Regex source http://www.regular-expressions.info/email.html
                $EmailRegex = '[a-z0-9!#$%&''*+/=?^_`{|}~-]+(?:\.[a-z0-9!#$%&''*+/=?^_`{|}~-]+)*@(?:[a-z0-9](?:[a-z0-9-]*[a-z0-9])?\.)+[a-z0-9](?:[a-z0-9-]*[a-z0-9])?'
                
                
                If (([regex]::Match($CurrentEmailAddress, $EmailRegex, "IgnoreCase")).Success -and $SpacePosition -eq -1) {
                    
                    $TestEmailFormatResult = "PASS"
                    
                }
                Else {
                    
                    $TestEmailFormatResult = "FAIL"
                    
                }
		    
            }
            
            #endregion     
            
            #$AdditionalTest = [Microsoft.Exchange.Data.SmtpProxyADdress]::Parse($CurrentEmailAddress).ParseException
            
            #$AdditionalTest
            
            #region Splitting current email address                                    
            
            If ($TestEmailFormatResult -eq "PASS") {
                
                $AtPosition = $CurrentEmailAddress.IndexOf("@")
                
                $CurrentEmailAddressLenght = $CurrentEmailAddress.Length
                
                $CurrentEmailDomain = $CurrentEmailAddress.Substring($AtPosition + 1, $CurrentEmailAddressLenght - ($AtPosition + 1))
                
                $TestAcceptedDomainsCurrent = $true
                
                $TestIfExistsCurrent = $true
                
            }
            Else {
                
                $TestAcceptedDomainsCurrent = $false
                
                $TestIfExistsCurrent = $false
                
            }
            
            If ($TestAcceptedDomains -and $TestAcceptedDomainsCurrent) {
                
                If ( ($AcceptedDomains | Where-Object -FilterScript { $_.domainname -eq $CurrentEmailDomain } | Measure-Object ).count -eq 1) {
                    
                    Write-Verbose -Message $CurrentEmailDomain
                    
                    $TestAcceptedDomainsResult = "PASS"
                    
                }
                Else {
                    
                    $TestAcceptedDomainsResult = "FAIL"
                    
                }
                
            }
            Else {
                
                $TestAcceptedDomainsResult = "SKIPPED"
                
            }
            
            if ($TestIfExists -and $TestIfExistsCurrent) {
                
                $TestIfExistResult = "EXISTS"
                
                Try {
                    
                    $Recipient = Get-Recipient $CurrentEmailAddress -ErrorAction Stop
                    
                }
                Catch {
                    
                    $TestIfExistResult = "NON EXIST"
                    
                    $ExistingObjectAlias = "NON EXIST"
                    
                    $ExistingObjectType = "NON EXIST"
                    
                    $ExistingObjectIsPrimary = "NON EXIST"
                    
                }
                
                If ($TestIfExistResult -eq "EXISTS") {
                    
                    $ExistingObjectAlias = $Recipient.alias
                    
                    $ExistingObjectType = $Recipient.RecipientType
                    
                    $ExistingObjectIsPrimary = $Recipient.IsPrimaryAddress
                    
                }
                                
            }
            Else {
                
                $TestIfExistResult = "SKIPPED"
                
                $ExistingObjectAlias = "SKIPPED"
                
                $ExistingObjectType = "SKIPPED"
                
                $ExistingObjectIsPrimary = "SKIPPED"
                
            }
            
            #endregion                                                
            
            $Result | Add-Member -Type NoteProperty -Name EmailAddress -value $CurrentEmailAddress
            $Result | Add-Member -Type NoteProperty -Name EmailDomain -value $CurrentEmailDomain
            $Result | Add-Member -Type NoteProperty -Name TestWhiteChars -Value $TestWhiteCharsResult
            $Result | Add-Member -Type NoteProperty -Name TestEmailFormat -Value $TestEmailFormatResult
            $Result | Add-Member -Type NoteProperty -Name TestAcceptedDomain -Value $TestAcceptedDomainsResult
            $Result | Add-Member -Type NoteProperty -Name TestEmailExists -Value $TestIfExistResult
            $Result | Add-Member -Type NoteProperty -Name ExistingObjectAlias -value $ExistingObjectAlias
            $Result | Add-Member -Type NoteProperty -Name ExistingObjectType -value $ExistingObjectType
            $Result | Add-Member -Type NoteProperty -Name IsPrimaryAddress -value $ExistingObjectIsPrimary
            
            $Results += $Result
            
        } #endregion Main loop
        
    } #END PROCESS
    
    END {
        
        Return $Results
        
    } #END END :-)
    
}
