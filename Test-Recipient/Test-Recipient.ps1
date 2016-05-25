function Test-Recipient {
    
<#
    .SYNOPSIS
    Return the Exchange mail object basic data and email addresses as comma-separated string.
    
    .DESCRIPTION
    Function intended for return basic data about mail object in Exchange Server environment and  email addresses for recipient.
       
    .PARAMETER InputFilePath
    File contains e.g. alias, guid, email address or any other recipients identifiers which can be used with Identity for Get-Recipient command.
   
    .PARAMETER Prefix
	Provide address prefix, default is SMTP.
    
    .PARAMETER CheckDomain
    Provide domain name what need to be checked if a recipient have email addresses from it.
    
    .PARAMETER DisplayProgressBar
    Select if a progress bar should be displayed under checking. Displaying progress bar can increase execution time.
    
        
    .EXAMPLE
    Test-Recipient -InputFilePath .\RecipientsToGetSMTP.txt -CheckDomain "example.mail.onmicrosoft.com"
    
    Addresses                           : smtp:EX2010A_USER8@ex2010a.lab,SMTP:EX2010A_USER8@ex2013a.contoso.com,smtp:EX2010A_USER
                                          8@example.mail.onmicrosoft.com
    IsValid                             : True
    Prefix                              : smtp
    ContainsAddressFromCheckedDomain    : True
    Alias                               : EX2010A_USER8
    IsMissed                            : False
    RecipientType                       : UserMailbox
    AddressesCount                      : 3
    DisplayName                         : USER8 EX2010A

    Addresses                           : smtp:EX2010A_USER9@ex2010a.lab,SMTP:EX2010A_USER9@ex2013a.contoso.com
    IsValid                             : True
    Prefix                              : smtp
    ContainsAddressFromCheckedDomain    : False
    Alias                               : EX2010A_USER9
    IsMissed                            : False
    RecipientType                       : UserMailbox
    AddressesCount                      : 2
    DisplayName                         : USER9 EX2010A
      
	.LINK
	https://github.com/it-praktyk/Exchange/Test-Recipient
	
	.LINK
	https://www.linkedin.com/in/sciesinskiwojciech
	      
	.NOTES
	AUTHOR: Wojciech Sciesinski, wojciech[at]sciesinski[dot]net
	KEYWORDS: PowerShell, Exchange, SMTP, addresses
   
	VERSIONS HISTORY
	- 0.1.0 - 2016-05-26 - Initial version published on GitHub

	TODO
    - check if Exchange cmdlets are available
    - add support for recipients from pipeline
    - implement better domain checking
    - verify names of internal variables
    - replace ForEach aliasses
    - update help
    
		
	LICENSE
	Copyright (c) 2016 Wojciech Sciesinski
    This function is licensed under The MIT License (MIT)
    Full license text: https://opensource.org/licenses/MIT
   
#>
    
    [CmdletBinding()]
    [OutputType([System.Collections.ArrayList])]
    Param (
        
        [Parameter(Mandatory = $true)]
        [String]$InputFilePath,
        [Parameter(Mandatory = $false)]
        [String]$Prefix = "smtp",
        [Parameter(Mandatory = $false)]
        [String]$CheckDomain,
        [Parameter(Mandatory = $false)]
        [Switch]$DisplayProgressBar
        
    )
    
    
    BEGIN {
        
        $Results = New-Object System.Collections.ArrayList
        
        Write-Verbose -Message "Provided input file $InputFilePath"
        
        If (Test-Path -Path $InputFilePath) {
            
            If ((Get-Item -Path $InputFilePath) -is [System.IO.fileinfo]) {
                
                try {
                    
                    $Recipients = Get-Content -Path $InputFilePath -ErrorAction Stop | Where-Object -FilterScript { ($_).ToString().Trim() -ne "" }
                    
                    $RecipientsCount = $($Recipients | Measure-Object).count
                    
                    Write-Verbose -Message "Imported recipients:  $RecipientsCount"
                    
                }
                catch {
                    
                    
                    Write-Error -Message "Read input file $InputFilePath error "
                    
                    break
                    
                }
                
            }
            
            Else {
                
                Write-Error -Message "Provided value for InputFilePath is not a file"
                
                
                break
                
            }
            
        }
        Else {
            
            Write-Error -Message "Provided value for InputFilePath doesn't exist"
                        
            break
        }
        
        
    }
    
    PROCESS {
        
        $Recipients | ForEach  {
            
            If ($DisplayProgressBar.IsPresent) {
                
                $PercentCompleted = [math]::Round(($i / $RecipientsCount) * 100)
                
                $StatusText = "Percent completed $PercentCompleted%, currently the recipient {0} is checked. " -f $($_).ToString()
                
                Write-Progress -Activity "Checking SMTP addresses" -Status $StatusText -PercentComplete $PercentCompleted
                
            }
            
            Try {
                
                $CurrentRecipientIdentifier = $_
                
                Write-Verbose "Currently addresses for $CurrentRecipientIdentifier are checked"
                
                $Recipient = Get-Recipient -Identity ($_).tostring().trim() -ErrorAction Stop | Select-Object -Property Alias, DisplayName, RecipientType, EmailAddresses
                
                Write-Verbose "Currently addresses for $_ are checked"
                
                $Recipient | ForEach {
                    
                    $ContainsFromCheckedDomain = $false
                    
                    $CurrenRecipient = $_
                    
                    $CurrentRecipientAddresses = ($CurrenRecipient | Select-Object -ExpandProperty EmailAddresses)
                    
                    $EmailsWithPrefix = $CurrentRecipientAddresses | Where-Object -FilterScript { $_.Prefix -match $Prefix }
                    
                    $EmailsWithPrefixCount = ($EmailsWithPrefix | Measure-Object).Count
                    
                    If ($EmailsWithPrefixCount -gt 0) {
                        
                        [String]$CurrentPrefixAddresses = [string]::Join(",", $($EmailsWithPrefix | ForEach { $_.ProxyAddressString }))
                        
                        If ($CurrentPrefixAddresses -match $CheckDomain) {
                            
                            $ContainsFromCheckedDomain = $true
                            
                        }
                    }
                    Else {
                        
                        $CurrentPrefixAddresses = $null
                        
                        $ContainsFromCheckedDomain = $false
                        
                    }
                    
                    $CurrentProperties = @{
                        Alias = $_.Alias
                        DisplayName = $_.DisplayName
                        Prefix = $Prefix
                        AddressesCount = $EmailsWithPrefixCount
                        Addresses = $CurrentPrefixAddresses
                        RecipientType = $_.RecipientType
                        ContainsAddressFromCheckedDomain = $ContainsFromCheckedDomain
                        IsValid = $true
                        IsMissed = $false
                    }
                    
                    $Result = New-Object PSObject -Property $CurrentProperties
                    
                    Write-Verbose $Result
                    
                    $Results.Add($Result) | Out-Null
                    
                }
                
            }
            
            Catch {
                
                $CurrentProperties = @{
                    Alias = $_.CurrentRecipientIdentifier
                    DisplayName = "Not found"
                    Prefix = $Prefix
                    AddressesCount = 0
                    Addresses = $CurrentPrefixAddresses
                    RecipientType = "Not found"
                    ContainsAddressFromCheckedDomain = $ContainsFromCheckedDomain
                    IsValid = $false
                    IsMissed = $true
                }
                
                $Result = New-Object PSObject -Property $CurrentProperties
                
                Write-Verbose -Message $Result
                
            }
            
            Finally {
                
                $i++
                
            }
            
        }
        
    }
    
    End {
        
        Return $Results
        
    }
    
    
}