function Test-EmailAddress {
    
    <#
	.SYNOPSIS
		Function is intended to verify the correctness of addresses email in Microsoft Exchange Enviroment
		
	.DESCRIPTION
		Function which can be used to verifing an email address before for example adding it to Microosft Exchange environment. 
		Checks perfomed: 
		a) if email address contain wrong characters e.g. % or spaces
		b) if email address is from domain which are on accepted domains list
		c) if email address is currently assigned to any object in Exchange environment (a conflicted object exist)
	
	.PARAMETER EmailAddress
		Email address which need to be verified in Exchange environment
    
    .PARAMETER TestEmailFormat
    
    .PARAMETER TestAcceptedDomains
    
    .PARAMETER TestIfExists
    
    .PARAMETER TestIsPrimary
    

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
		0.5.1 - 2015-03-09 - compatibility issue on Exchange 2010 (PowerShell 2.0) resolved
        0.6.0 - 2015-12-22 - the function rewriten, information about license added
    
        TODO
        - implement test "TestIsPrimary"
        - update help
        - descriptive test result add

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
        
#>    
    
    [cmdletbinding()]
    #[OutputType(System.Object[])]
    param (
        
        [parameter(ValueFromPipeline, ValueFromPipelineByPropertyName, mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [alias("email")]
        [String[]]$EmailAddress,
        [parameter(Mandatory = $false)]
        [Bool]$TestEmailFormat = $true,
        [parameter(Mandatory = $false)]
        [Bool]$TestAcceptedDomains = $true,
        [parameter(Mandatory = $false)]
        [Bool]$TestIfExists = $true,
        [parameter(Mandatory = $false)]
        [Bool]$TestIsPrimary = $true
        
    )
    
    BEGIN {
        
        #region Declare variables                                                        
        
        $Results = @()
        
        if ($TestAcceptedDomains) {
            
            #This try/catch block check if Exchange commands are available
            Try {
                
                $AcceptedDomains = Get-AcceptedDomain
                
            }
            
            Catch [System.Management.Automation.CommandNotFoundException] {
                
                $TestAcceptedDomains = $false
                
                $TestAcceptedDomainsResult = "SKIPPED"
                
            }
            
        }
        
        #endregion        
        
    } #END BEGIN
    
    PROCESS {
        
        
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
                
                #Check if space is in midle of email
                $SpacePosition = $CurrentEmailAddress.IndexOf(" ")
                
                #Regex source http://www.regular-expressions.info/email.html
                $EmailRegex = '[a-z0-9!#$%&''*+/=?^_`{|}~-]+(?:\.[a-z0-9!#$%&''*+/=?^_`{|}~-]+)*@(?:[a-z0-9](?:[a-z0-9-]*[a-z0-9])?\.)+[a-z0-9](?:[a-z0-9-]*[a-z0-9])?'
                
                
                If (([regex]::Match($CurrentEmailAddress, $EmailRegex, "IgnoreCase ")).Success -and $SpacePosition -eq -1) {
                    
                    $TestEmailFormatResult = "PASS"
                    
                }
                Else {
                    
                    $TestEmailFormatResult = "FAIL"
                    
                }
                
                
            }
            
            #endregion            
            
            #region Splitting current email address                                    
            
            If ($TestEmailFormatResult = "PASS") {
                
                $AtPosition = $CurrentEmailAddress.IndexOf("@")
                
                $CurrentEmailAddressLenght = $CurrentEmailAddress.Length
                
                $CurrentEmailDomain = $CurrentEmailAddress.Substring($AtPosition + 1, $CurrentEmailAddressLenght - ($AtPosition + 1))
                
                $TestAcceptedDomainsCurrent = $false
                
                $TestIfExistsCurrent = $false
                
            }
            Else {
                
                $TestAcceptedDomainsCurrent = $false
                
                $TestIfExistsCurrent = $false
                
            }
            
            If ($TestAcceptedDomains -and $TestAcceptedDomainsCurrent) {
                
                If (($AcceptedDomains | Where-Object -FilterScript { $_.domainname -eq $CurrentEmailDomain } | Measure-Object).count -eq 1) {
                    
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
                    
                    $Recipient = Get-Recipient $EmailAddress -ErrorAction Stop
                    
                }
                Catch {
                    
                    $TestIfExistResult = "NON EXIST"
                    
                    $ExistingObjectAlias = "NON EXIST"
                    
                    $ExistingObjectType = "NON EXIST"
                    
                }
                
                If ($TestIfExistResult -eq "EXISTS") {
                    
                    $ExistingObjectAlias = $Recipient.alias
                    
                    $ExistingObjectType = $Recipient.RecipientType
                    
                }
                
                
            }
            Else {
                
                $TestIfExistResult = "SKIPPED"
                
            }
            
            #endregion                                                
            
            $Result | Add-Member -type NoteProperty -Name EmailAddress -value $CurrentEmailAddress
            $Result | Add-Member -type NoteProperty -Name TestWhiteChars -Value $TestWhiteCharsResult
            $Result | Add-Member -type NoteProperty -Name TestEmailFormat -Value $TestEmailFormatResult
            $Result | Add-Member -Type NoteProperty -Name TestAcceptedDomain -Value $TestAcceptedDomainsResult
            $Result | Add-Member -Type NoteProperty -Name TestEmailExists -Value $TestEmailFormatResult
            $Result | Add-Member -Type NoteProperty -Name ExistingObjectAlias -value $ExistingObjectAlias
            $Result | Add-Member -Type NoteProperty -Name ExistingObjectType -value $Recipient.RecipientType
            
            
            $Results += Result
            
        } #endregion Main loop
        
    } #END PROCESS
    
    END {
        
        Return $Results
        
    } #END END :-)
    
}
