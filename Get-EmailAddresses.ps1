function Get-EmailAddresses {
    
<#
    .SYNOPSIS
    Function intended for get email addresses  with any prefix for Exchange server recipients taken from flat text file
    
    .DESCRIPTION
       
  
    .PARAMETER InputFilePath
     File contains e.g. alias, guid, email address or any other recipients identifiers which can be used with Identity for Get-Recipient command 
   
        
    .EXAMPLE
    Get-EmailAddresses -InputFilePath .\RecipientsToGetSMTP.txt   
      
	.LINK
	https://github.com/it-praktyk/Get-EmailAddresses
	
	.LINK
	https://www.linkedin.com/in/sciesinskiwojciech
	      
	.NOTES
	AUTHOR: Wojciech Sciesinski, wojciech[at]sciesinski[dot]net
	KEYWORDS: PowerShell, Exchange, EmailAddresses, ProxyAddresses
   
	VERSIONS HISTORY
	0.6.0 - 2015-05-14 - First version published on GitHub, partially based on Get-SMTPAddresses v. 0.7.0
	0.6.1 - 2015-05-14 - AppendPrefixToResultAddresses parameter added, emails subgroup for prefixes added to output object
	0.7.0 - 2015-05-14 - Updated to achieve compatibility with Exchange 2010 and PowerShell 2.0

	TODO
    - check if Exchange cmdlets are available
    - add verification recipient type 
        # Recipients types - can be used in the next version
        # "DynamicDistributionGroup","UserMailbox","MailUser","MailContact","MailUniversalDistributionGroup","MailUniversalSecurityGroup","MailNonUniversalGroup","PublicFolder"
    - update help
    
		
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
        [String]$InputFilePath,
        
        [Parameter(Mandatory = $false, Position = 1)]
        [Bool]$AppendPrefixToResultAddresses = $false
        
    )
    
    
    BEGIN {
        
        $Results = @()
        
        Write-Verbose -Message "Provided input file $InputFilePath"
        
        If (Test-Path -Path $InputFilePath) {
            
            If ((Get-Item -Path $InputFilePath) -is [System.IO.fileinfo]) {
                
                try {
                    
                    $RecipientsFromFile = Get-Content -Path $InputFilePath -ErrorAction Stop | Where { ($_).ToString().Trim() -ne "" }
                    
                    $RecipientsFromFileCount = $($RecipientsFromFile | Measure-Object).count
                    
                    Write-Verbose -Message "Imported recipients:  $RecipientsFromFileCount"
                    
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
        
        $RecipientsFromFile | ForEach  {
            
            $PercentCompleted = [math]::Round(($i / $RecipientsFromFileCount) * 100)
            
            $StatusText = "Percent completed $PercentCompleted%, currently the recipient {0} is checked. " -f $($_).ToString()
            
            Write-Progress -Activity "Checking SMTP addresses" -Status $StatusText -PercentComplete $PercentCompleted
            
            #Try {
            
            $CurrentRecipientIdentifier = $_
            
            Write-Verbose "Currently addresses for $CurrentRecipientIdentifier are checked"
            
            #Be carefully - the result can be more than one recipient !
            $FoundRecipients = Get-Recipient -Identity ($CurrentRecipientIdentifier).tostring().trim() -ErrorAction Stop | Select-Object -Property Alias, DisplayName, RecipientType, EmailAddresses, Guid
            
            $FoundRecipientsCount = ($FoundRecipients | Measure-Object).Count
            
            if ($FoundRecipientsCount -gt 0) {
                
                Write-Verbose "Amount of recipients found for $CurrentRecipientIdentifier is $FoundRecipientsCount - email addresses are checked now"
                
                $FoundRecipients | ForEach {
                    
                    $CurrentRecipient = $_
                    
                    $CurrentRecipientAddresses = ($CurrentRecipient | select -ExpandProperty EmailAddresses)
                    
                    $CurrentRecipientAddressesCount = ($CurrentRecipientAddresses | Measure-Object).Count
                    
                    $SortedPrefixes = @()
                    
                    ($CurrentRecipientAddresses | select Prefix -Unique | ForEach { $SortedPrefixes += ($_.Prefix) })
                    
                    $PrefixesCount = ($SortedPrefixes | Measure-Object).Count
                    
                    [String]$MessageText = "Recipient with identifier {0} resolved to {1} has assigned {2} email addresses in {3} prefixes groups." `
                    -f $CurrentRecipientIdentifier, $CurrentRecipient.DisplayName, $CurrentRecipientAddressesCount, $PrefixesCount
                    
                    Write-Verbose -Message $MessageText
                    
                    $Result = New-Object PSObject
                    
                    $Result | Add-Member -Type 'NoteProperty' -Name RecipientIdentifier -Value $CurrentRecipientIdentifier
                    
                    $Result | Add-Member -type 'NoteProperty' -name RecipientAlias -value $CurrentRecipient.Alias
                    
                    $Result | Add-Member -type 'NoteProperty' -name RecipientDisplayName -value $CurrentRecipient.DisplayName
                    
                    $Result | Add-Member -Type 'NoteProperty' -Name RecipientType -Value $CurrentRecipient.RecipientType
                    
                    $Result | Add-Member -Type 'NoteProperty' -Name RecipientGuid -Value $CurrentRecipient.Guid
                    
                    $Result | Add-Member -Type 'NoteProperty' -Name RecipientEmailsCount -Value $CurrentRecipientAddressesCount
                    
                    $Result | Add-Member -Type 'NoteProperty' -Name PrefixesGroupsCountforRecipient -value $PrefixesCount
                    
                    $Result | Add-Member -Type 'NoteProperty' -Name IsValid -Value $true
                    
                    $Result | Add-Member -Type 'NoteProperty' -Name IsMissed -Value $false
                    
                    #Loop for prefixes
                    $SortedPrefixes | foreach {
                        
                        $CurrentPrefix = $_.PrimaryPrefix
                        
                        $EmailsWithPrefix = $CurrentRecipientAddresses | where { $_.Prefix -match $CurrentPrefix }
                        
                        $EmailsWithPrefixCount = ($EmailsWithPrefix | Measure-Object).Count
                        
                        [String]$CurrentPrefixAddressCountColumnName = "{0}{1}{2}" -f "AddressesForPrefix_", $CurrentPrefix, "_Count"
                        
                        $Result | Add-Member -Type 'NoteProperty' -Name $CurrentPrefixAddressCountColumnName -Value $EmailsWithPrefixCount
                        
                        [String]$CurrentPrefixAddressessColumnName = "{0}{1}" -f "AddressesForPrefix_", $CurrentPrefix
                        
                        [String]$CurrentPrefixAddresses = [string]::Join(",", $($CurrentRecipientAddresses | where { $_.Prefix -match $CurrentPrefix } | ForEach { $_.ProxyAddressString }))
                        
                        $Result | Add-Member -Type 'NoteProperty' -Name $CurrentPrefixAddressessColumnName -Value $CurrentPrefixAddresses
                        
                        $e = 1
                        
                        #Loop for emails with current prefix
                        
                        $EmailsWithPrefix | foreach {
                            
                            $CurrentEmailAddressWithPrefix = $_
                            
                            [String]$CurrentEmailAddressColumnName = "{0}{1}{2}" -f $CurrentPrefix, "_Address_", $e
                            
                            If ($AppendPrefixToResultAddresses) {
                                
                                $Result | Add-Member -Type 'NoteProperty' -Name $CurrentEmailAddressColumnName -Value $CurrentEmailAddressWithPrefix.ProxyAddressString
                                
                            }
                            Else {
                                
                                $Result | Add-Member -Type 'NoteProperty' -Name $CurrentEmailAddressColumnName -Value $CurrentEmailAddressWithPrefix.AddressString
                                
                            }
                            
                            $e++
                            
                        }
                        
                        
                    }
                    
                    Write-Verbose $Result
                    
                    $Results += $Result
                    
                }
                
            }
            
            #}
            
			<#
			
            Catch {
			
				Result = New-Object PSObject
				
				$Result | Add-Member -Type 'NoteProperty' -Name RecipientIdentifier -Value $CurrentRecipientIdentifier
				
				$Result | Add-Member -Type 'NoteProperty' -Name IsValid -Value $false
                        
				$Result | Add-Member -Type 'NoteProperty' -Name IsMissed -Value $true
                
                Write-Verbose -Message $Result
                
                $Results += $Result
                
            }
			
			#>
            
            #Finally {
            
            $i++
            
            #}
            
        }
        
    }
    
    End {
        
        Return $Results
        
    }
    
    
}