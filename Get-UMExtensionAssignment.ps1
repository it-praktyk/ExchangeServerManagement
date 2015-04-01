Function Get-UMExtensionAssignment
{
    
<#
	.SYNOPSIS
	Function intended for searching current UM extension assigment in Exchange Server environment
    
    .DESCRIPTION
    Function intended for searching current UM extension assigment in Exchange Server environment,
    as a result the unified enabled mailbox is returned for which extension is assigned.
  
	.PARAMETER UMExtensionsFile
	Raw text file with UM extensions to check, one extension by line
    
    .LINK
	https://github.com/it-praktyk/Get-UMExtensionAssigment
	
	.LINK
	https://www.linkedin.com/in/sciesinskiwojciech
        
	.EXAMPLE
   	Get-UMExtensionAssigment -UMExtensionsFile .\umextensions.txt
      
	.NOTES
	AUTHOR: Wojciech Sciesinski, wojciech[at]sciesinski[dot]net
	KEYWORDS: PowerShell, UM, Exchange, Lync, Active Directory
	VERSION HISTORY
	0.5.1 - 2015-04-02 - First version uploaded to GitHub
	
    TODO
    - check if Exchange cmdlets are available
    - add parameter to provide extension from CLI
   
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
    
    param (
        
        [parameter(mandatory = $false)]
        [alias("Path", "UMExtensionsFilePath")]
        [String]$InputFilePath
                
    )
    
    BEGIN
    {
        
        If (Test-Path -Path $InputFilePath)
        {
            
            If ((Get-Item -Path $InputFilePath) -is [System.IO.fileinfo])
            {
                
                try
                {
                    
                    $Extensions = get-Content -Path $InputFilePath
                    
                    [Int]$ExtensionsCount = $($Extensions | Measure-Object).count
                    
                    Write-Verbose "$ExtensionsCount extensions to check"
                    
                }
                catch
                {
                    
                    Write-Error "Read input file $InputFilePath error "
                    
                    break
                    
                }
                
            }
            
            Else
            {
                
                Write-Error "Provided value for InputFilePath is not a file"
                
                break
                
            }
            
        }
        Else
        {
            
            Write-Error "Provided value for InputFilePath doesn't exist"
            
            break
        }
        
        Write-Verbose -Message "Read UM enabled mailboxes from Active Directory"
        
        $UMMailboxes = Get-UMMailbox -ResultSize Unlimited
        
        $UMMailboxesCount = $($UMMailboxes | Measure-Object).Count
        
        Write-Verbose -Message "$UMMailboxesCount read"
        
        $Results = @()
        
        $i = 1
        
    }
    
    PROCESS
    {
        
        $Extensions | ForEach
        {
            
            $PercentCompleted = [math]::Round(($i / $ExtensionsCount) * 100)
            
            $StatusText = "Percent completed $PercentCompleted%, currently the extension {0} is checked. " -f $($_).ToString()
            
            Write-Progress -Activity "Searching current extension assignment. " -Status $StatusText -PercentComplete $PercentCompleted
            
            $CurrentUMExtension = $_
            
            Write-Verbose -Message "Looking for $_"
            
            $CurrentUMMailbox = ($UMMailboxes | where { $_.Extensions -eq $CurrentUMExtension } | select -Property Name, DisplayName, Guid, LinkedMasterAccount, UMEnabled, PrimarySMTPAddress, Extensions)
            
            $SortedExtensions = ($CurrentUMMailbox | Select -ExpandProperty Extensions | Sort)
            
            $Result = New-Object PSObject
            
            $Result | Add-Member -type NoteProperty -name MailboxAlias -value $CurrentUMMailbox.Name
            
            $Result | Add-Member -type NoteProperty -name MailboxDisplayName -value $CurrentUMMailbox.DisplayName
            
            $Result | Add-Member -type NoteProperty -name MailboxGuid -value $CurrentUMMailbox.Guid
            
            $Result | Add-Member -type NoteProperty -name LinkedMasterAccount -value $CurrentUMMailbox.LinkedMasterAccount
            
            $Result | Add-Member -type NoteProperty -name PrimarySMTPAddress -value $CurrentUMMailbox.PrimarySMTPAddress
            
            $Result | Add-Member -type NoteProperty -name MailboxUMEnabled -value $CurrentUMMailbox.UMEnabled
            
            $Result | Add-Member -type NoteProperty -name MailboxUMExtensionsCount -value $($SortedExtensions | Measure).Count
            
            $Result | Add-Member -type NoteProperty -name MailboxUMExtensions -value $($SortedExtensions -join ',')
            
            $e = 1
            
            $SortedExtensions | ForEach {
                
                $Result | Add-Member -type NoteProperty -name MailboxUMExtensions$e -value $_
                
                $e++
                                
            }
            
            [String]$MessageText = "UM extension {0} currently assigned to mailbox {1} with PrimarySMTPAddress {2} ; mailbox is UMEnabled: {3}" `
            -f $CurrentUMExtension, $CurrentUMMailbox.Name, $CurrentUMMailbox.PrimarySMTPAddress, $CurrentUMMailbox.UMEnabled
            
            Write-Verbose -Message $MessageText
            
            $Results += $Result
            
            $i++
            
        }
        
    }
    
    END
    {
        
        Return $Results
        
    }
    
}