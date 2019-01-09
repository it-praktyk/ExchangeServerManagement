function Get-SMTPAddresses {
    
<#
    .SYNOPSIS
    Function intended for get email addresses  with "smtp" prefix for Exchange server recipients taken from flat text file
    
    .DESCRIPTION
       
  
    .PARAMETER InputFilePath
     File contains e.g. alias, guid, email address or any other recipients identifiers which can be used with Identity for Get-Recipient command 
   
    .PARAMETER Prefix
   
     .PARAMETER Scope
      Currently not used but can be added validation function implemented
        
    .EXAMPLE
    Get-SMTPAddresses -InputFilePath .\RecipientsToGetSMTP.txt   
      
	.LINK
	https://github.com/it-praktyk/Get-SMTPAddresses
	
	.LINK
	https://www.linkedin.com/in/sciesinskiwojciech
	      
	.NOTES
	AUTHOR: Wojciech Sciesinski, wojciech[at]sciesinski[dot]net
	KEYWORDS: PowerShell, Exchange, SMTP
   
	VERSIONS HISTORY
	0.7.0 -  2015-05-13 - First version published on GitHub

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
        [String]$Prefix = "smtp"
        
    )
    
    
    BEGIN {
        
        $Results = @()
        
        Write-Verbose -Message "Provided input file $InputFilePath"
        
        If (Test-Path -Path $InputFilePath) {
            
            If ((Get-Item -Path $InputFilePath) -is [System.IO.fileinfo]) {
                
                try {
                    
                    $Recipients = Get-Content -Path $InputFilePath -ErrorAction Stop | Where { ($_).ToString().Trim() -ne "" }
                    
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
            
            $PercentCompleted = [math]::Round(($i / $RecipientsCount) * 100)
            
            $StatusText = "Percent completed $PercentCompleted%, currently the recipient {0} is checked. " -f $($_).ToString()
            
            Write-Progress -Activity "Checking SMTP addresses" -Status $StatusText -PercentComplete $PercentCompleted
            
            
            
            Try {
                
                $CurrentRecipientIdentifier = $_
                
                Write-Verbose "Currently addresses for $CurrentRecipientIdentifier are checked"
                
                $Recipient = Get-Recipient -Identity ($_).tostring().trim() -ErrorAction Stop | Select-Object -Property Alias, DisplayName, RecipientType -ExpandProperty EmailAddresses
                
                Write-Verbose "Currently addresses for $_ are checked"
                
                $Recipient | ForEach {
                    
                    
                    if ($_.PrefixString -match "smtp") {
                        
                        
                        $CurrentProperties = @{
                            Alias = $_.Alias
                            DisplayName = $_.DisplayName
                            SMTPAddress = $_.AddressString
                            IsPrimaryAddress = $_.IsPrimaryAddress
                            RecipientType = $_.RecipientType
                            IsValid = $true
                            IsMissed = $false
                        }
                        
                        $Result = New-Object PSObject -Property $CurrentProperties
                        
                        Write-Verbose $Result
                        
                        $Results += $Result
                        
                    }
                    
                }
                
            }
            
            Catch {
                
                $CurrentProperties = @{
                    Alias = $_.CurrentRecipientIdentifier
                    IsValid = $false
                    IsMissed = $true
                }
                
                $Result = New-Object PSObject -Property $CurrentProperties
                
                Write-Verbose -Message $Result
                
            }
            
            Finally {
                
                $i += 1
                
            }
            
        }
        
    }
    
    End {
        
        Return $Results
        
    }
    
    
}