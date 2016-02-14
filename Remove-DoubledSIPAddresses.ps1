function Remove-DoubledSIPAddresses {
<#
    .SYNOPSIS
    Function intended for verifying and removing doubled SIP addresses from all mailboxes in Exchange Server environment
    
    .DESCRIPTION 
    Function intended for verifying and removing doubled SIP addresses from all mailboxes in Exchange Server environment. `
    Only address in a domain provided in a parameter CorrectSIPDomain will be keep.
    
    .PARAMETER Identity 
    The Identity parameter specifies the identity of the mailbox. You can use one of the following values:
    * GUID
    * Distinguished name (DN)
    * Display name
    * Domain\Account
    * User principal name (UPN)
    * LegacyExchangeDN
    * SmtpAddress
    * Alias
    
    .PARAMETER CorrectSIPDomain
    Name of domain for which correct SIPs belong. If the parameter is not set domain from PrimarySMTPAddress will be used.
    
    .INPUT
    To see the input types that this cmdlet accepts, see Cmdlet Input and Output Types (http://go.microsoft.com/fwlink/p/?linkId=313798). If the Input Type field for a cmdlet is blank, the cmdlet doesn't accept input data.
    The function accept the same input types like Get-Mailbox.
    
    .OUTPUTS
    PowerShell object what contains properties: MailboxAlias, MailboxDisplayName, PrimarySMTPAddress, MailboxGuid, SIPAddressesBeforeCount, SIPAddressesList, SIPAddresses
    The SIPAddressesList property is a string with comma delimited list of SIPs. The SIPAddresses property is Microsoft.Exchange.Data.CustomProxyAddress data type and can be expanded.

    .EXAMPLE
    
    [PS] > Remove-DoubledSIPAddresses -Identity aa473815 -WhatIf -Verbose -CorrectSIPDomain contoso.com
    
    VERBOSE: Mailbox with alias AA473815 has assigned 2 SIP addresses.
    What if: Performing operation "Remove SIP address sip:ingrid.thomes@example.com" on Target "mailbox: AA473815".

    MailboxAlias              : AA473815
    MailboxDisplayName        : Wolters-van der Thomes, IAV (Ingrid)
    MailboxSMTPPrimaryAddress : ingrid.wolters-van.der.thomes@example.com
    MailboxGuid               : b201434a-1f62-4ee4-a446-e0b2bc7badc9
    SIPAddressesBeforeCount   : 2
    SIPAddressesBeforeList    : SIP:Ingrid.Wolters@contoso.com,sip:ingrid.thomes@example.com
    SIPAddressesBefore        : {SIP:Ingrid.Wolters@contoso.com,sip:ingrid.thomes@example.com}
    SIPAddressAfterCount      : 2
    SIPAddressesAfterList     : SIP:Ingrid.Wolters@contoso.com,sip:ingrid.thomes@example.com
    SIPAddressesAfter         : {SIP:Ingrid.Wolters@contoso.com,sip:ingrid.thomes@example.com}
    
    .EXAMPLE
    
    [PS] > Remove-DoubledSIPAddresses -Identity aa473815 -Verbose
    
    VERBOSE: Mailbox with alias AA473815 has assigned 2 SIP addresses.
	VERBOSE: SIP address Ingrid.Wolters@contoso.com is incorrect and will be deleted
    
    MailboxAlias              : AA473815
    MailboxDisplayName        : Wolters-van der Thomes, IAV (Ingrid)
    MailboxSMTPPrimaryAddress : ingrid.wolters-van.der.thomes@tailspintoys.com
    MailboxGuid               : b201434a-1f62-4ee4-a446-e0b2bc7badc9
    SIPAddressesBeforeCount   : 2
    SIPAddressesBeforeList    : SIP:Ingrid.Wolters@contoso.com,sip:ingrid.thomes@tailspintoys.com
    SIPAddressesBefore        : {SIP:Ingrid.Wolters@contoso.com,sip:ingrid.thomes@tailspintoys.com}
    SIPAddressAfterCount      : 1
    SIPAddressesAfterList     : SIP:ingrid.thomes@tailspintoys.com
    SIPAddressesAfter         : {SIP:ingrid.thomes@tailspintoys.com}

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
    0.1.3 - 2015-05-31 - Help updated, a script reformatted
    0.1.4 - 2015-06-09 - Primary SMTP address added to report file
    0.2.0 - 2016-02-10 - Report cappabilities removed from the function, input from pipeline added, the licence changed to MIT
        
    LICENSE
    Copyright (c) 2016 Wojciech Sciesinski
    This function is licensed under The MIT License (MIT)
    Full license text: https://opensource.org/licenses/MIT
   
#>  
    
    [CmdletBinding(SupportsShouldProcess = $true)]
    Param (
        
        [parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Position = 0)]
        [Alias("Mailbox")]
        $Identity,
        [Parameter(Mandatory = $false, Position = 1)]
        [String]$CorrectSIPDomain
        
    )
    
    
    BEGIN {
        
        #Check if the function is run in the Exchange Management Shell
        
        [String]$CmdletForCheck = "Get-Mailbox"
        
        $CmdletAvailable = Test-Path -Path Function:$CmdletForCheck
        
        If (!$CmdletAvailable) {
            
            [String]$MessageText = "The function need to be run in the Exchange Management Shell"
            
            Throw $MessageText
            
        }
        
        $Results = @()
        
    }
    
    PROCESS {
        
        ForEach ($MailboxInLoop in $Identity) {
            
            If ($MailboxInLoop.gettype().fullname -eq 'Microsoft.Exchange.Data.Directory.Management.Mailbox') {
                
                $CurrentMailbox = $MailboxInLoop | Select-Object -Property Alias, DisplayName, RecipientType, EmailAddresses, Guid, PrimarySMTPAddress
                
            }
            Else {
                
                $CurrentMailbox = Get-Mailbox $MailboxInLoop | Select-Object -Property Alias, DisplayName, RecipientType, EmailAddresses, Guid, PrimarySMTPAddress
                
            }
            
            $CurrentMailboxSIPAddressesBefore = ($CurrentMailbox | Select-Object -ExpandProperty EmailAddresses | Where-Object -FilterScript { $_.prefix -match 'SIP' })
            
            $CurrentMailboxSIPAddressesBeforeCount = ($CurrentMailboxSIPAddressesBefore | Measure-Object).Count
            
            if ($CurrentMailboxSIPAddressesBeforeCount -gt 1) {
                
                if ([String]::IsNullOrEmpty($CorrectSIPDomain)) {
                    
                    $CorrectSIPDomain = $($CurrentMailbox.PrimarySMTPAddress).Domain
                    
                }
                
                [String]$MessageText = "Mailbox with alias {0} has assigned {1} SIP addresses." -f $CurrentMailbox.Alias, $CurrentMailboxSIPAddressesBeforeCount
                
                Write-Verbose -Message $MessageText
                
                $Result = New-Object PSObject
                
                $Result | Add-Member -type 'NoteProperty' -name MailboxAlias -value $CurrentMailbox.Alias
                
                $Result | Add-Member -type 'NoteProperty' -name MailboxDisplayName -value $CurrentMailbox.DisplayName
                
                $Result | Add-Member -type 'NoteProperty' -Name MailboxSMTPPrimaryAddress -Value $CurrentMailbox.PrimarySMTPAddress
                
                $Result | Add-Member -Type 'NoteProperty' -Name MailboxGuid -Value $CurrentMailbox.Guid
                
                $Result | Add-Member -Type 'NoteProperty' -Name SIPAddressesBeforeCount -Value $CurrentMailboxSIPAddressesBeforeCount
                
                [String]$CurrentSIPAddressesList = [string]::Join(",", $($CurrentMailboxSIPAddressesBefore | ForEach-Object -Process { $_.ProxyAddressString }))
                
                $Result | Add-Member -Type 'NoteProperty' -Name SIPAddressesBeforeList -Value $CurrentSIPAddressesList
                
                $Result | Add-Member -Type 'NoteProperty' -Name SIPAddressesBefore -Value $CurrentMailboxSIPAddressesBefore
                
                foreach ($CurrentSIPObject in $CurrentMailboxSIPAddressesBefore) {
                    
                    [String]$CurrentSIPAddressString = $CurrentSIPObject.AddressString
                    
                    $AtPosition = $CurrentSIPAddressString.IndexOf("@")
                    
                    $SIPAddressLenght = $CurrentSIPAddressString.Length
                    
                    [String]$CurrentSIPDomain = $CurrentSIPAddressString.Substring($AtPosition + 1, $SIPAddressLenght - ($AtPosition + 1))
                    
                    
                    If ($CurrentSIPDomain -ne $CorrectSIPDomain) {
                        
                        if ($CurrentSIPObject.IsPrimaryAddress -eq $true) {
                            
                            $CurrentSIPObject.ToSecondary() | Out-Null
                            
                        }
                        
                        [String]$SIPToRemove = $CurrentSIPObject.ProxyAddressString
                        
                        [String]$MessageText = [String]$MessageText = "SIP address {0} is incorrect and will be deleted." `
                        -f $SIPToRemove
                        
                        if ($PSCmdlet.ShouldProcess("mailbox: $($CurrentMailbox.Alias)", "Remove SIP address $SIPToRemove")) {
                            
                            Set-Mailbox -Identity $CurrentMailbox.Alias -EmailAddresses @{ remove = $SIPToRemove } -ErrorAction Continue
                            
                        }

                    }
                    
                }
                
                $CurrentMailboxSIPAddressesAfter = (Get-Mailbox -Identity $CurrentMailbox.Alias | Select-Object -ExpandProperty EmailAddresses | Where-Object -FilterScript { $_.prefix -match 'SIP' })
                
                $CurrentMailboxSIPAddressesCountAfter = ($CurrentMailboxSIPAddressesAfter | Measure-Object).Count
                
                If ($CurrentMailboxSIPAddressesCountAfter -gt 1) {
                    
                    [String]$CurrentSIPAddressesListAfter = [string]::Join(",", $($CurrentMailboxSIPAddressesAfter | ForEach-Object -Process { $_.ProxyAddressString }))
                    
                }
                
                elseif ($CurrentMailboxSIPAddressesCountAfter -eq 1) {
                    
                    $CurrentSIPAddressesListAfter = $CurrentMailboxSIPAddressesAfter.ProxyAddressString
                    
                }
                
                Else {
                    
                    $CurrentMailboxSIPAddressesCountAfter = 0
                    
                }
                
                $Result | Add-Member -Type 'NoteProperty' -Name SIPAddressesAfterCount -Value $CurrentMailboxSIPAddressesCountAfter
                
                $Result | Add-Member -Type 'NoteProperty' -Name SIPAddressesAfterList -Value $CurrentSIPAddressesListAfter
                
                $Result | Add-Member -Type 'NoteProperty' -Name SIPAddressesAfter -Value $CurrentMailboxSIPAddressesAfter
                
                $Results += $Result
                
            }
            
        }
        
    }
    
    
    End {
        
        Return $Results
        
    }
    
}