function Test-EmailAddress {
<#
    .SYNOPSIS
    The function Test-EmailAddress is intended to verify the correctness of email addresses in Microsoft Exchange Server enviroment
        
    .DESCRIPTION
    Function which can be used to verifing email addresses in Microsoft Exchange Server environment. 
    
    Checks perfomed: 
    a) if an email address provided as parameter value contains wrong characters e.g. spaces at the begining/end
    b) if an email address format is complaint with requirements - check Wikipedia https://en.wikipedia.org/wiki/Email_address
    c) if an email address is from a domain which are added to the accepted domains list of is in the list passed as parameter value
    d) if an email address is currently assigned to any object in Exchange environment (a conflicted object exist)
    e) if an email address is currently set as PrimarySMTPAddress for existing object
    
    .PARAMETER EmailAddress
    Email address which need to be verified in Exchange environment
    
    .PARAMETER TestEmailFormat
    Set to false to skip testing email address format
    
    .PARAMETER TestAcceptedDomains
    Set to false to skip testing if domain of an email address is in accepted domain list
    
    .PARAMETER TestIfExists
    Set to false to skip testing if an email address exist in mail organization
    
    .PARAMETER TestIsPrimary
    Set to false to skip testing if email is primary for existing object

    .PARAMETER AcceptedDomains
    The list of domains used to testing if email is from accepted domains.

    .EXAMPLE
    [PS] > Test-EmailAddress -EmailAddress dummy@example.com 
    
    EmailAddress              : dummy@example.com
    EmailDomain               : example.com
    TestWhiteChars            : PASS
    TestEmailFormat           : PASS
    TestAcceptedDomain        : PASS
    TestEmailExists           : EXISTS
    ExistingObjectAlias       : dummy
    ExisitngObjectGuid        : 181ca5f1-2fc0-40ef-853d-215a2b1fd16d
    ExistingObjectType        : UserMailbox
    IsPrimaryAddress          : True
    EmailAddressPolicyEnabled : True
        
    .EXAMPLE
    
    [PS] >"postmaster@gelatto.test","new@new.pl","new@test@new.pl" | Test-EmailAddress -AcceptedDomains new.pl
    
    WARNING: Only domains passed to the AcceptedDomains parameter will be evaluated under the TestAcceptedDomains test

    EmailAddress              : postmaster@gelatto.test
    EmailDomain               : gelatto.test
    TestWhiteChars            : PASS
    TestEmailFormat           : PASS
    TestAcceptedDomain        : FAIL
    TestEmailExists           : EXISTS
    ExistingObjectAlias       : DL_mailadmin
    ExisitngObjectGuid        : d897b45f-c104-4d3b-b77e-6f4332565f8451
    ExistingObjectType        : MailUniversalSecurityGroup
    IsPrimaryAddress          : False
    EmailAddressPolicyEnabled : False

    EmailAddress              : new@new.pl
    EmailDomain               : new.pl
    TestWhiteChars            : PASS
    TestEmailFormat           : PASS
    TestAcceptedDomain        : PASS
    TestEmailExists           : NO EXISTS
    ExistingObjectAlias       : NO EXISTS
    ExisitngObjectGuid        : NO EXISTS
    ExistingObjectType        : NO EXISTS
    IsPrimaryAddress          : NO EXISTS
    EmailAddressPolicyEnabled : NO EXISTS

    EmailAddress              : new@test@new.pl
    EmailDomain               : new.pl
    TestWhiteChars            : PASS
    TestEmailFormat           : FAIL
    TestAcceptedDomain        : SKIPPED
    TestEmailExists           : SKIPPED
    ExistingObjectAlias       : SKIPPED
    ExisitngObjectGuid        : SKIPPED
    ExistingObjectType        : SKIPPED
    IsPrimaryAddress          : SKIPPED
    EmailAddressPolicyEnabled : SKIPPED

    .EXAMPLE
    
    [PS] > Get-AcceptedDomain
    
    Name                           DomainName                     DomainType                   Default
    ----                           ----------                     ----------                   -------
    gto.local                      gto.local                      Authoritative                False
    gelatto.test                   gelatto.test                   InternalRelay                True
    example.com                    example.com                    Authoritative                False
        
    [PS] > Get-Mailbox dummy | Select-Object -ExpandProperty emailaddresses | Where-Object -FilterScript { $_.prefix -match 'smtp' } | ForEach { Test-EmailAddress $_.SMTPAddress }
    
    EmailAddress              : dummy@example.com
    EmailDomain               : example.com
    TestWhiteChars            : PASS
    TestEmailFormat           : PASS
    TestAcceptedDomain        : PASS
    TestEmailExists           : EXISTS
    ExistingObjectAlias       : dummy
    ExisitngObjectGuid        : 181ca5f1-2fc0-40ef-853d-215a2b1fd16d
    ExistingObjectType        : UserMailbox
    IsPrimaryAddress          : True
    EmailAddressPolicyEnabled : True
    
    EmailAddress              : dummy.user@gelatto.test
    EmailDomain               : gelatto.test
    TestWhiteChars            : PASS
    TestEmailFormat           : PASS
    TestAcceptedDomain        : PASS
    TestEmailExists           : EXISTS
    ExistingObjectAlias       : dummy
    ExisitngObjectGuid        : 181ca5f1-2fc0-40ef-853d-215a2b1fd16d
    ExistingObjectType        : UserMailbox
    IsPrimaryAddress          : False
    EmailAddressPolicyEnabled : True
        
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
    0.8.0 - 2015-12-31 - the function tested, the parameter $AcceptedDomains implemented, help updated
	0.9.0 - 2016-05-12 - license changed to MIT, check for accepted domains corrected
    
    TODO
    - add an additional parameter AcceptOnlyEnglishLetters
    - add an additional parameter AllowedCharsExclusionList
    
    LICENSE
	Copyright (c) 2016 Wojciech Sciesinski
    This function is licensed under The MIT License (MIT)
    Full license text: https://opensource.org/licenses/MIT
    
    .OUTPUTS
    System.Object[]
        
#>  
    
    [cmdletbinding()]
    [OutputType([System.Object[]])]
    param (
        
        [parameter(ValueFromPipeline = $true, mandatory = $true)]
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
        
        [Bool]$AcceptedDomainsAsParameter = -not [string]::IsNullOrEmpty($AcceptedDomains)
        
        If ($TestAcceptedDomains) {
            
            If ($AcceptedDomainsAsParameter) {
                
                $MessageText = "Only domains passed to the AcceptedDomains parameter will be evaluated under the TestAcceptedDomains test"
                
                Write-Warning -Message $MessageText
            }
            Else {
                
                #This try/catch block check if Exchange commands are available - means function is running in Exchange Management Shell
                Try {
                    
                    $AcceptedDomains = $(Get-AcceptedDomain)
                    
                }
                
                Catch [System.Management.Automation.CommandNotFoundException] {
                    
                    Write-Verbose -Message "Error occured $error[0]"
                    
                    $TestAcceptedDomains = $false
                    
                    $TestAcceptedDomainsResult = "SKIPPED"
                    
                }
                
            }
        }
        Else {
            
            $TestAcceptedDomains = $false
            
            $TestAcceptedDomainsResult = "SKIPPED"
            
        }
        
        #endregion Declare variables      
        
    }
    
    PROCESS {
        
        #region Main loop
        $EmailAddress | ForEach-Object -Process {
            
            $CurrentEmailAddress = $_
            
            $Result = New-Object PSObject
            
            #region Checking email format                                                                        
            
            if ($TestEmailFormat) {
                
                #Check if white chars are on the begining/end of provided email
                if ($CurrentEmailAddress.Trim() -ne $CurrentEmailAddress) {
                    
                    $TestWhiteCharsResult = "FAIL"
                    
                    $CurrentEmailAddress = $CurrentEmailAddress.Trim()
                    
                }
                Else {
                    
                    $TestWhiteCharsResult = "PASS"
                    
                }
                
                Try {
                    
                    $ParsedEmailAddress = [Microsoft.Exchange.Data.SmtpProxyADdress]::Parse($CurrentEmailAddress)
                    
                    If (($ParsedEmailAddress.gettype()).name -ne 'InvalidProxyAddress') {
                        
                        $TestEmailFormatResult = "PASS"
                        
                    }
                    Else {
                        
                        $TestEmailFormatResult = "FAIL"
                        
                    }
                    
                }
                
                #If a function is not runned in EMS email address will be parsed using regex
                Catch {
                
                    #Warning that EMS is not used to the function run
                    
                    $MessageText = "The function should be runned in Exchange Management Shell, some test will be skipped, email format will be checked in unrestrictive mode."
                
                    Write-Warning -Message $MessageText
                    
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
                
            }
            
            #endregion Checking email format
            
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
                
                If (($AcceptedDomains | Where-Object -FilterScript { $_.DomainName -eq $CurrentEmailDomain } | Measure-Object).count -eq 1) {
                    
                    Write-Verbose -Message $CurrentEmailDomain
                    
                    $TestAcceptedDomainsResult = "PASS"
                    
                }
                Else {
                    
                    Write-Verbose -Message $CurrentEmailDomain
                    
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
                    
                    $TestIfExistResult = "NO EXISTS"
                    
                    $ExistingObjectAlias = "NO EXISTS"
                    
                    $ExistingObjectGuid = "NO EXISTS"
                    
                    $ExistingObjectType = "NO EXISTS"
                    
                    $ExistingObjectEmailIsPrimary = "NO EXISTS"
                    
                    $ExistingObjectEmailAddressPolicyEnabled = "NO EXISTS"
                    
                }
                
                If ($TestIfExistResult -eq "EXISTS") {
                    
                    $ExistingObjectAlias = $Recipient.alias
                    
                    $ExistingObjectType = $Recipient.RecipientType
                    
                    If ( $TestIsPrimary ) {
                    
                        $ExistingObjectEmailIsPrimary = ($Recipient.PrimarySMTPAddress -eq $CurrentEmailAddress)
                    
                    }
                    
                    $ExistingObjectEmailAddressPolicyEnabled = $Recipient.EmailAddressPolicyEnabled
                    
                    $ExistingObjectGuid = $Recipient.Guid
                    
                }
                
            }
            Else {
                
                $TestIfExistResult = "SKIPPED"
                
                $ExistingObjectAlias = "SKIPPED"
                
                $ExistingObjectGuid = "SKIPPED"
                
                $ExistingObjectType = "SKIPPED"
                
                $ExistingObjectEmailIsPrimary = "SKIPPED"
                
                $ExistingObjectEmailAddressPolicyEnabled = "SKIPPED"

            }
            
            #endregion                                                
            
            $Result | Add-Member -Type NoteProperty -Name EmailAddress -value $CurrentEmailAddress
            $Result | Add-Member -Type NoteProperty -Name EmailDomain -value $CurrentEmailDomain
            $Result | Add-Member -Type NoteProperty -Name TestWhiteChars -Value $TestWhiteCharsResult
            $Result | Add-Member -Type NoteProperty -Name TestEmailFormat -Value $TestEmailFormatResult
            $Result | Add-Member -Type NoteProperty -Name TestAcceptedDomain -Value $TestAcceptedDomainsResult
            $Result | Add-Member -Type NoteProperty -Name TestEmailExists -Value $TestIfExistResult
            $Result | Add-Member -Type NoteProperty -Name ExistingObjectAlias -value $ExistingObjectAlias
            $Result | Add-Member -Type NoteProperty -Name ExisitngObjectGuid -value $ExistingObjectGuid
            $Result | Add-Member -Type NoteProperty -Name ExistingObjectType -value $ExistingObjectType
            $Result | Add-Member -Type NoteProperty -Name IsPrimaryAddress -value $ExistingObjectEmailIsPrimary
            $Result | Add-Member -Type NoteProperty -Name EmailAddressPolicyEnabled -value $ExistingObjectEmailAddressPolicyEnabled
            
            $Results += $Result
            
        } #endregion Main loop
        
    } #END PROCESS
    
    END {
        
        Return $Results
        
    } #END END :-)
    
}