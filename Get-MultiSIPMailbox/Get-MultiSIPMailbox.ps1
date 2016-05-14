function Get-MultiSIPMailbox {
<#
    .SYNOPSIS
    Function intended for verifying if mailbox in the Exchange Server environment has assigned more than one SIP address.
    
    .DESCRIPTION 
    Function intended for verifying if the mailbox in the Exchange Server environment has assigned more than one SIP address.
    Only mailboxes with multiplie SIP addresses are returned.
    
    .PARAMETER Identity 
    The Identity parameter specifies the identity of the mailbox. You can use one of the following values:
    - GUID
    - Distinguished name (DN)
    - Display name
    - Domain\Account
    - User principal name (UPN)
    - LegacyExchangeDN
    - SmtpAddress
    - Alias
    
    .INPUTS
    To see the input types that this cmdlet accepts, see Cmdlet Input and Output Types (http://go.microsoft.com/fwlink/p/?linkId=313798). If the Input Type field for a cmdlet is blank, the cmdlet doesn't accept input data.
    The function accept the same input types like Get-Mailbox.
    
    .OUTPUTS
    PowerShell object what contains properties: MailboxAlias, MailboxDisplayName, PrimarySMTPAddress, MailboxGuid, SIPAddressesCount, SIPAddressesList, SIPAddresses
    The SIPAddressesList property is a string with comma delimited list of SIPs. The SIPAddresses property is Microsoft.Exchange.Data.CustomProxyAddress data type and can be expanded.
    
    .LINK
    https://github.com/it-praktyk/Get-MultiSIPMailbox
    
    .LINK
    https://www.linkedin.com/in/sciesinskiwojciech
    
    .EXAMPLE
    
    Check if the mailbox has assigned more than one SIP address - direct providing the mailbox identity
    
    [PS] > Get-MultiSIPMailbox -Identity AA473815

    MailboxAlias              : AA473815
    MailboxDisplayName        : Wolters-van der Thomes, IAV (Ingrid)
    PrimarySMTPAddress        : ingrid.wolters-van.der.thomes@example.com
    MailboxGuid               : b201434a-1f62-4ee4-a446-e0b2bc7badc9
    SIPAddressesCount         : 2
    SIPAddressesList          : SIP:Ingrid.Wolters@example.com,sip:ingrid.thomes@example.com
    SIPAddresses              : {SIP:Ingrid.Wolters@example.com, sip:ingrid.thomes@example.com}
    
    .EXAMPLE
    
    Check if the mailbox has assigned more than one SIP address - providing the mailbox identity by pipeline
    
    [PS] > Get-Mailbox AA473815 | Get-MultiSIPMailbox

    MailboxAlias              : AA473815
    MailboxDisplayName        : Wolters-van der Thomes, IAV (Ingrid)
    PrimarySMTPAddress        : ingrid.wolters-van.der.thomes@example.com
    MailboxGuid               : b201434a-1f62-4ee4-a446-e0b2bc7badc9
    SIPAddressesCount         : 2
    SIPAddressesList          : SIP:Ingrid.Wolters@example.com,sip:ingrid.thomes@example.com
    SIPAddresses              : {SIP:Ingrid.Wolters@example.com, sip:ingrid.thomes@example.com}
          
    .NOTES
    AUTHOR: Wojciech Sciesinski, wojciech[at]sciesinski[dot]net
    KEYWORDS: PowerShell, Exchange, SIPAddresses, ProxyAddresses, Lync, migration
   
    VERSIONS HISTORY
    - 0.1.0 - 2016-02-12 - First version published on GitHub
    - 0.1.1 - 2016-02-12 - Help updated
    - 0.1.2 - 2016-02-14 - Example updated
	- 0.2.0 - 2016-05-12 - Help structure corrected, script reformatted, verbose message corrected, improved execution time
            
    LICENSE
    Copyright (c) 2016 Wojciech Sciesinski
    This function is licensed under The MIT License (MIT)
    Full license text: https://opensource.org/licenses/MIT
   
#>
    
    [CmdletBinding()]
    Param (
        
        [parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Position = 0)]
        [alias("Mailbox")]
        $Identity
        
    )
    
    BEGIN {
        
        #Check if the function is run in the Exchange Management Shell
        
        [String]$CmdletForCheck = "Get-Mailbox"
        
        $CmdletAvailable = Test-Path -Path Function:$CmdletForCheck
        
        If (!$CmdletAvailable) {
            
            [String]$MessageText = "The function need to be run in the Exchange Management Shell"
            
            Throw $MessageText
            
        }

		$Results = New-Object System.Collections.ArrayList
		
    }
	
    PROCESS {
        
        ForEach ($MailboxInLoop in $Identity) {
            
            If ($MailboxInLoop.gettype().fullname -eq 'Microsoft.Exchange.Data.Directory.Management.Mailbox') {
                
                $CurrentMailbox = $MailboxInLoop | Select-Object -Property Alias, DisplayName, RecipientType, EmailAddresses, Guid, PrimarySMTPAddress
                
            }
            Else {
                
                $CurrentMailbox = Get-Mailbox $MailboxInLoop | Select-Object -Property Alias, DisplayName, RecipientType, EmailAddresses, Guid, PrimarySMTPAddress
                
            }
            
            $CurrentMailboxSIPAddresses = ($CurrentMailbox | Select-Object -ExpandProperty EmailAddresses | Where-Object -FilterScript { $_.prefix -match 'SIP' })
            
            $CurrentMailboxSIPAddressesCount = ($CurrentMailboxSIPAddresses | Measure-Object).Count
            
            If ($CurrentMailboxSIPAddressesCount -gt 1) {
                
                [String]$MessageText = "Mailbox with alias {0} has assigned {1} SIP addresses." -f $CurrentMailbox.Alias, $CurrentMailboxSIPAddressesCount
                
                Write-Verbose -Message $MessageText
                
                $Result = New-Object PSObject
                
                $Result | Add-Member -type 'NoteProperty' -name MailboxAlias -value $CurrentMailbox.Alias
                
                $Result | Add-Member -type 'NoteProperty' -name MailboxDisplayName -value $CurrentMailbox.DisplayName
                
                $Result | Add-Member -type 'NoteProperty' -Name PrimarySMTPAddress -Value $CurrentMailbox.PrimarySMTPAddress
                
                $Result | Add-Member -Type 'NoteProperty' -Name MailboxGuid -Value $CurrentMailbox.Guid
                
                $Result | Add-Member -Type 'NoteProperty' -Name SIPAddressesCount -Value $CurrentMailboxSIPAddressesCount
                
                [String]$CurrentSIPAddressesList = [string]::Join(",", $($CurrentMailboxSIPAddresses | ForEach-Object -Process { $_.ProxyAddressString }))
                
                $Result | Add-Member -Type 'NoteProperty' -Name SIPAddressesList -Value $CurrentSIPAddressesList
                
                $Result | Add-Member -Type 'NoteProperty' -Name SIPAddresses -Value $CurrentMailboxSIPAddresses
                
                $Results.Add($Result) | Out-Null
                
            }
            
        }
        
    }
    
    End {
        
        Return $Results
        
    }
    
}