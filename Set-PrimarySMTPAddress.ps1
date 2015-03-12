Function Set-PrimarySMTPAddress {

<#   
    .SYNOPSIS   
    Set primary SMTP address for mail objects in Exchange Server environment
  
    .DESCRIPTION   
    Allow set primary SMTP address for mail objects (mailboxes, distributions groups, mail enabled security groups)
    based on input file in CSV format
    
    .PARAMETER InputFilePath
    The path for file with input data
    
    .PARAMETER VerifyInputFileForDuplicates
    By default input file is verified for duplicates
    
    .PARAMETER Mode
    The switch which define action to perform - default DisplayOnly

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

    KEYWORDS: PowerShell, Exchnange, Active Directory, SMTP

    VERSION HISTORY
    0.1.0 - Initial release - untested !
    
         
    TO DO
    The structure for input files need to be parsed.
    
    .EXAMPLE

    Input CSV file format - <FILL>

    
    
#> 

[CmdletBinding()] 

param (

    [parameter(Mandatory=$true)]
    [String]$InputFilePath,
    
    [parameter(Mandatory=$false)]
    [Bool]$VerifyInputFileForDuplicates=$true,

    [parameter(Mandatory=$false)]
    [ValidateSet("DisplayOnly","PerformActions","Rollback")]
    [String]$Mode="DisplayOnly",
    
    [parameter(Mandatory=$false)]
    [ValidateSet("UserMailbox","MailNonUniversalGroup","MailUniversalDistributionGroup")]
    [String]$RecipientType="UserMailbox",
    
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

    #Uncomments if you need hunt any bug
    Set-StrictMode -version 2

    [String]$StartTime = Get-Date -format yyyyMMdd-HHmm

    If ( $CreateTranscript ) {
    
        Start-NewTranscript -TranscriptFileNamePrefix "Transcript-Set-PrimarySMTPAddress-" -StartTimeSuffix $StartTime

    }
    
    If (Test-Path -Path $InputFilePath ) {

        If ( (Get-Item -Path $InputFilePath) -is [System.IO.fileinfo]) {
        
            try {

                $RecipientsFromInputFile = Import-CSV -Path $InputFilePath -Delimiter ";" -ErrorAction Stop
                
                [Int]$RecipientsCount = $( $Mailboxes | Measure-Object).count
            
            }                
            catch {
            
                Write-Error "Read input file $InputFilePath error "

                Stop-Transcript -ErrorAction SilentlyContinue
        
                break
            
            }
        
        }
        
        Else {
        
            Write-Error "Provided value for InputFilePath is not a file"
        
            Stop-Transcript -ErrorAction SilentlyContinue
        
            break
            
        }
        
    }        
    Else {
    
        Write-Error "Provided value for InputFilePath doesn't exist"
        
        Stop-Transcript -ErorAction SilentlyContinue

        break
    }
    
    #Declare variable for store results data
    $Results=@()
    
    [int]$i=1

}

PROCESS {

        $RecipientsFromInputFile | ForEach { 
        
        
            $PercentCompleted = [math]::Round(($i / $RecipientsCount) * 100)

            $StatusText = "Percent completed $PercentCompleted%, currently the mailbox {0} is checked. " -f $($_.RecipientIdentity).ToString()

            Write-Progress -Activity "Performing action in mode $Mode" -Status $StatusText -PercentComplete $PercentCompleted
        
            [String]$MessageText="Performing check on object {0}  in mode: {1} ." -f $_.RecipientIdentity , $Mode
        
            Write-Verbose -Message $MessageText
            
            Try {
                        
                $CurrentRecipientTest1 = $(Get-Recipient $_.RecipientIdentity -ErrorAction Stop | Where { $_.RecipientType -eq $RecipientType })
                
                $CurrentRecipientTest2 = $(Get-Recipient $_.NewPrimarySMTPAddress -ErrorAction Stop | Where { $_.RecipientType -eq $RecipientType })
                
                if ( $CurrentRecipientTest1 -ne $CurrentRecipientTest2 ) {
                
                    Write-Error -Message "Email address $_.NewPrimarySMTPAddress is not currently assigned to recipient $_.RecipientIdentity with type $_.RecipientType"
                    
                    Break
                
                }
                Else {
                
                    $CurrentRecipient = $CurrentRecipientTest1
                
                }
                
            }
            
            Catch {
            
                Write-Error "Recipient $_.RecipientIdentity or with address $_.NewPrimarySMTPAddress doesn't exist"
                
                Break
            
            }
            
            Finally {

                $iserror = $false
            
                If ( $CurrentRecipient.RecipientType -eq 'UserMailbox' ) {
                
                    $CurrentMailbox = Get-Mailbox -Identity $($CurrentRecipient.Alias)
    
                    Write-Verbose -Message "Performing action on $CurrentMailbox.Alias in mode $Mode ."
                    
                    #Object properties before any changes

                    $Result = New-Object PSObject
                    
                    $Result | Add-Member -MemberType NoteProperty -Name RecipientIdentity -Value $CurrentRecipient.RecipientIdentity
                    
                    $Result | Add-Member -MemberType NoteProperty -name RecipientType -value $CurrentRecipient.RecipientType
                    
                    $Result | Add-Member -MemberType NoteProperty -Name RecipientGuid -Value $CurrentMailbox.Guid
                    
                    $Result | Add-Member -MemberType NoteProperty -name RecipientAlias  -value $CurrentMailbox.Alias
                    
                    $Result | Add-Member -MemberType NoteProperty -Name PrimarySMTPAddressBefore -Value $CurrentMailbox.PrimarySMTPAddress
                        
                    $AllProxyAddressesStringBefore = ( @(select-Object -InputObject $CurrentMailbox -expandproperty emailaddresses) -join ',')
                        
                    $Result | Add-Member -MemberType NoteProperty -name ProxyAddressesBefore -value $AllProxyAddressesStringBefore
                    
                    $Result | Add-Member -MemberType NoteProperty -Name PrimarySMTPAddressProposal -Value $CurrentRecipient.NewPrimarySMTPAddress
                        
                    #[String]$ProxyAddressStringToAdd = "{0}{1}" -f $Prefix, $_.proxyAddresses
                        
                    #[String]$ProxyAddressStringProposal = "{0},{1}" -f $AllProxyAddressesStringBefore,$ProxyAddressStringToAdd
                        
                    #$Result | Add-Member -MemberType NoteProperty -name ProxyAddressesProposal -value $ProxyAddressStringProposal
                    
                    If ( $Mode -eq 'DisplayOnly' ) {
                    
                        $Result | Add-Member -MemberTypeName NoteProperty -Name PrimarySMTPAddressProposal -Value $CurrentRecipient.NewPrimarySMTPAddress

                        $Result | Add-Member -MemberType NoteProperty -name ProxyAddressesProposal -value $AllProxyAddressesStringBefore #This need to be changed - replace for 
                            
                    }
                    
                    Elseif ( $Mode -eq 'PerformActions') {
                            
                        Set-Mailbox -Identity $CurrentMailbox -PrimarySMTPAddress $CurrentRecipient.NewPrimarySMTPAddress -ErrorAction Continue
                            
                        $CurrentMailboxAfter = Get-Mailbox -Identity $($CurrentRecipient.Alias)
                        
                        $Result | Add-Member -MemberType NoteProperty -Name PrimarySMTPAddressAfter -Value $CurrentMailboxAfter.PrimarySMTPAddress
                                                
                        $AllProxyAddressesStringAfter = ( @(select-Object -InputObject $CurrentMailboxAfter -ExpandProperty emailaddresses) -join ',')
                        
                        $Result | Add-Member -MemberType NoteProperty -name ProxyAddressesAfter -value $AllProxyAddressesStringAfter
                                                
                    }
                        
                    ElseIf ( $Mode -eq 'Rollback' ) {
    
                        Write-Error -Message "Rollback mode is not implemented yet"
    
                    }
                        
                    $Results+=$Result
                    
                    $i+=1
                
                }
                
                <#
ElseIf ( $CurrentRecipient.RecipientType -eq 'MailNonUniversalGroup' -or $CurrentRecipient.RecipientType -eq 'MailUniversalDistributionGroup' ) {
                
                    $CurrentGroup = Get-DistributionGroup -Identity $($CurrentRecipient.Alias)
                    
                    Write-Verbose -Message "Performing action on $CurrentGroup.Alias in mode $Mode ."
                    
                    #Object properties before any changes

                    $Result = New-Object PSObject

                    $Result | Add-Member -MemberType NoteProperty -name MailboxAlias  -value $CurrentGroup.Alias
                    
                    $Result | Add-Member -MemberType NoteProperty -name RecipientType -value $CurrentRecipient.RecipientType
                        
                    $AllProxyAddressesStringBefore = ( @(select-Object -InputObject $CurrentGroup -expandproperty emailaddresses) -join ',')
                        
                    $Result | Add-Member -MemberType NoteProperty -name ProxyAddressesBefore -value $AllProxyAddressesStringBefore
                        
                    [String]$ProxyAddressStringToAdd = "{0}{1}" -f $Prefix, $_.proxyAddresses
                        
                    [String]$ProxyAddressStringProposal = "{0},{1}" -f $AllProxyAddressesStringBefore,$ProxyAddressStringToAdd
                        
                    $Result | Add-Member -MemberType NoteProperty -name ProxyAddressesProposal -value $ProxyAddressStringProposal
                    
                    If ( $Mode -eq 'DisplayOnly' ) {

                        $Result | Add-Member -MemberType NoteProperty -name ProxyAddressesAfter -value $AllProxyAddressesStringBefore
                            
                    }
                    
                    Elseif ( $Mode -eq 'PerformActions') {
                            
                        Set-DistributionGroup -Identity $CurrentGroup -EmailAddresses @{add=($ProxyAddressStringToAdd)} -ErrorAction Continue
                            
                        $CurrentMailboxAfter = Get-DistributionGroup -Identity $($CurrentRecipient.Alias)
                                                
                        $AllProxyAddressesStringAfter = ( @(select-Object -InputObject $CurrentMailboxAfter -ExpandProperty emailaddresses) -join ',')
                        
                        $Result | Add-Member -MemberType NoteProperty -name ProxyAddressesAfter -value $AllProxyAddressesStringAfter
                                                
                    }
                        
                    ElseIf ( $Mode -eq 'Rollback' ) {
    
                        Write-Error -Message "Rollback mode is not implemented yet"
    
                    }
                        
                    $Results+=$Result
                    
                    $i+=1
                
                }
#>

                
                Else {
                
                    Write-Error -Message "Currently only recipients with RecipientType: UserMailbox, MailNonUniversalGroup, MailUniversalDistributionGroup are support by script `n '
                    Current object RecipientType is $CurrentRecipient.RecipientType.ToString()"
                    
                    $IsError = $true
                
                }

            }
            
            If ($IsError) {
                    
                Break
                        
            }
    
    }

}


END {

    #Save results to rollback file - need to be moved to external function

    If ( $CreateRollbackFile ) {
        
        #Check if rollback directory exist and try create if not
        If ( !$((Get-Item -Path $RollBackFileDirectoryPath -ErrorAction SilentlyContinue) -is [system.io.directoryinfo]) ) {

            New-Item -Path $RollBackFileDirectoryPath -Type Directory -ErrorAction Stop
        
        }
            
        $FullRollbackFilePath = $RollBackFileDirectoryPath + $RollBackFileNamePrefix + $StartTime + '.csv'
            
        Write-Verbose "Rollback data will be written to $FullRollbackFilePath"
    
        Write-Verbose "Write rollback data to file $FullRollbackFilePath"

        #If export will not be unsuccessfull than display $Results to screen as the list - will be catched by Transcript
        Try {
        
            $Results | Export-CSV -Path $FullRollbackFilePath -NoTypeInformation -Delimiter ";" -Encoding UTF8 -ErrorAction Continue
            
        }
        
        Catch {
        
            If ( $CreateTranscript ) {
        
                $Results | Format-List
                
            }
            Else {
            
                Start-NewTranscript -TranscriptFileDirectoryPath ".\emergency-transcripts\" -TranscriptFileNamePrefix "Emergency-Transcript-"
            
            }
        
        }
        
    }
    
    #Display results to console - also can be redirected to file 
    Else {
    
        Return $Results

    }
        
    Stop-Transcript -ErrorAction SilentlyContinue

}
}


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
        0.5.1 - 2015-03-09 - corrected intentional error handling
        

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
        
        $Result | Add-Member -MemberType NoteProperty -Name EmailAddress -value $EmailAddress
        $Result | Add-Member -MemberType NoteProperty -Name ExitCode  -value 1
        $Result | Add-Member -MemberType NoteProperty -Name ExitDescription -value "Email doesn't contain 'at' char"
        $Result | Add-Member -MemberType NoteProperty -Name ConflictedObjectAlias -value "Not checked"
        $Result | Add-Member -MemberType NoteProperty -Name ConflictedObjectType -value "Not checked"
                
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
                
            $Result | Add-Member -MemberType NoteProperty -Name EmailAddress -value $EmailAddress
            $Result | Add-Member -MemberType NoteProperty -Name ExitCode  -value 5
            $Result | Add-Member -MemberType NoteProperty -Name ExitDescription -value "White chars e.g. spaces founded before/after email"
            $Result | Add-Member -MemberType NoteProperty -Name ConflictedObjectAlias -value "Not checked"
            $Result | Add-Member -MemberType NoteProperty -Name ConflictedObjectType -value  "Not checked"
                
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
                                
                Catch {
                    
                    Write-Verbose -Message "Email address doesn't exist in environment - finally result: is correct"
                    
                    $NotError = $false
                    
                    $Result | Add-Member -MemberType NoteProperty -Name EmailAddress -value $EmailAddress
                    $Result | Add-Member -MemberType NoteProperty -Name ExitCode  -value 0
                    $Result | Add-Member -MemberType NoteProperty -Name ExitDescription -value "No Error"
                    $Result | Add-Member -MemberType NoteProperty -Name ConflictedObjectAlias -value "No conflict"
                    $Result | Add-Member -MemberType NoteProperty -Name ConflictedObjectType -value "Not checked"
                    
                    Return $Result
                
                }
    
                If ( $NotError ) {
                    
                    Write-Verbose -Message "Recipient with email address $EmailAddress exist now."
                        
                    $Result | Add-Member -MemberType NoteProperty -Name EmailAddress -value $EmailAddress
                    $Result | Add-Member -MemberType NoteProperty -Name ExitCode  -value 2
                    $Result | Add-Member -MemberType NoteProperty -Name ExitDescription -value "Email exist now"
                    $Result | Add-Member -MemberType NoteProperty -Name ConflictedObjectAlias -value $Recipient.alias
                    $Result | Add-Member -MemberType NoteProperty -Name ConflictedObjectType -value $Recipient.RecipientType
                
                    Return $Result
                        
                }
    
            }
            Else {
        
                Write-Verbose -Message "Email address $EmailAddress contain unsupported chars"
                
                $Result | Add-Member -MemberType NoteProperty -Name EmailAddress -value $EmailAddress
                $Result | Add-Member -MemberType NoteProperty -Name ExitCode  -value 3
                $Result | Add-Member -MemberType NoteProperty -Name ExitDescription -value "Unsupported chars found"
                $Result | Add-Member -MemberType NoteProperty -Name ConflictedObjectAlias -value "Not checked"
                $Result | Add-Member -MemberType NoteProperty -Name ConflictedObjectType -value  "Not checked"
                
                Return $Result
        
            }
        }
    
        Else {
        
            Write-Verbose "Email address $EmailAddress is not from accepted domains."
            
            $Result | Add-Member -MemberType NoteProperty -Name EmailAddress -value $EmailAddress
            $Result | Add-Member -MemberType NoteProperty -Name ExitCode  -value 4
            $Result | Add-Member -MemberType NoteProperty -Name ExitDescription -value "Not accepted domain"
            $Result | Add-Member -MemberType NoteProperty -Name ConflictedObjectAlias -value "Not checked"
            $Result | Add-Member -MemberType NoteProperty -Name ConflictedObjectType -value  "Not checked"
                    
            Return $Result

        }

    }
    
}

END {
    
    #Nothing yet in this section

}

}


Function Start-NewTranscript {
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

    [parameter(Mandatory=$false)]
    [String]$TranscriptFileDirectoryPath=".\transcripts\",

    [parameter(Mandatory=$false)]
    [String]$TranscriptFileNamePrefix="Transcript-",
    
    [parameter(Mandatory=$false)]
    [String]$StartTimeSuffix

)

BEGIN {

    #Uncomments if you need hunt any bug
    Set-StrictMode -version 2
    
    If ( $StartTimeSuffix ) {
    
        [String]$StartTime = $StartTimeSuffix
        
    }
    Else {

        [String]$StartTime = Get-Date -format yyyyMMdd-HHmm
        
    }

    #Check if transcript directory exist and try create if not
        If ( !$((Get-Item -Path $TranscriptFileDirectoryPath -ErrorAction SilentlyContinue) -is [system.io.directoryinfo]) ) {

                New-Item -Path $TranscriptFileDirectoryPath -Type Directory -ErrorAction Stop | Out-Null
                
                Write-Verbose -Message "Folder $TranscriptFileDirectoryPath was created."
                
        }
        
        $FullTranscriptFilePath = $TranscriptFileDirectoryPath + '\' + $TranscriptFileNamePrefix + $StartTime + '.log'

        #Stop previous PowerShell transcript and catch error if not started previous

        try{

              stop-transcript  | Out-Null

        }

        catch [System.InvalidOperationException]{}
        
}

PROCESS {

        #Start new PowerShell transcript

        Start-Transcript -Path $FullTranscriptFilePath -ErrorAction Stop

        Write-Verbose "Transcript will be written to $FullTranscriptFilePath"
    
}

END {

}
 
}