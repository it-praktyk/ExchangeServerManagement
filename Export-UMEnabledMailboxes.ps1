Function Export-UMEnabledMailboxes {
    
<#
	.SYNOPSIS
	Function intended for export all UM enabled mailboxes to csv file. Default output directory is subdirectory with a name "UMMailboxes"
  
	.PARAMETER OutputFileDirectoryPath
	Directory path where outputfile need to be saved, if directory doesn't exists will be created
	
	.PARAMETER OutputFileNamePrefix
	Name prefix for the output file name
	
	.PARAMETER 
	Date and time which will be added in file name, if not provided currect date and time will be added
	        
	.EXAMPLE
   
	Export-UMEnabledMailboxes -OutputFileDirectoryPath "c:\UMMailboxes\" -OutputFileNamePrefix "MyDearCustomerName-UMMailboxes-" 
	
	As a result in directory c:\UMMailboxes\ will be created the file with a name MyDearCustomerName-UMMailboxes-2010331-1225.csv
	
	Columns in a file: MailboxAlias, MailboxDisplayName, MailboxGuid, LinkedMasterAccount, PrimarySMTPAddress, MailboxUMEnabled, MailboxUMExtensionsCount, MailboxUMExtensions1, MailboxUMExtensions2

      
	.NOTES
	AUTHOR: Wojciech Sciesinski, wojciech[at]sciesinski[dot]net
	KEYWORDS: PowerShell, UM, Exchange, Lync, Active Directory
	VERSION HISTORY
	0.1.0 - 2015-03-31 - An initial release mostly based on Check-UMExtensionAssignment.ps1 v. 0.5.0, the first version uploaded to GitHub


    TODO
    - check if Exchange cmdlets are available
   
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
        
        [parameter(Mandatory = $false)]
        [String]$OutputFileDirectoryPath = ".\UMMailboxes\",
        
        [parameter(Mandatory = $false)]
        [String]$OutputFileNamePrefix = "UMEnabledMailboxes-",
        
        [parameter(Mandatory = $false)]
        [String]$StartTimeSuffix
        
    )
    
    BEGIN {
        
        If ($StartTimeSuffix) {
            
            [String]$StartTime = $StartTimeSuffix
            
        }
        Else {
            
            [String]$StartTime = Get-Date -format yyyyMMdd-HHmm
            
        }
        
        If (!$((Get-Item -Path $OutputFileDirectoryPath -ErrorAction SilentlyContinue) -is [system.io.directoryinfo])) {
            
            New-Item -Path $OutputFileDirectoryPath -Type Directory -ErrorAction Stop | Out-Null
            
            Write-Verbose -Message "Folder $OutputFileDirectoryPath was created."
            
        }
        
        $FullOutputFilePath = $OutputFileDirectoryPath + '\' + $OutputFileNamePrefix + $StartTime + '.csv'
        
        
        Write-Verbose -Message "Read UM enabled mailboxes from Active Directory"
        
        $UMMailboxes = Get-UMMailbox -ResultSize Unlimited
        
        $UMMailboxesCount = $($UMMailboxes | Measure-Object).Count
        
        Write-Verbose -Message "$UMMailboxesCount read"
        
        $Results = @()
        
        $i = 1
        
    }
    
    PROCESS {
        
        $UMMailboxes | ForEach {
            
            $PercentCompleted = [math]::Round(($i / $UMMailboxesCount) * 100)
            
            $StatusText = "Percent completed $PercentCompleted%, currently the {0} is checked. " -f $($_).ToString()
            
            Write-Progress -Activity "Gathering data for . " -Status $StatusText -PercentComplete $PercentCompleted
            
            $CurrentUMMailbox = ($_ | select -Property Name, DisplayName, Guid, LinkedMasterAccount, UMEnabled, PrimarySMTPAddress, Extensions)
            
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
            
            [String]$MessageText = "UM enabled mailbox {0} with PrimarySMTPAddress {1} has assigned {2} extension(s), extensions list: {3} " `
            -f $CurrentUMMailbox.Name, $CurrentUMMailbox.PrimarySMTPAddress, $MailboxUMExtensionsCount, $($SortedExtensions -join ',')
            
            Write-Verbose -Message $MessageText
            
            $Results += $Result
            
            $i++
            
        }
        
    }
    
    END {
        
        $Results | Export-CSV -Path $FullOutputFilePath -Delimiter ";" -Encoding UTF8 -NoTypeInformation
        
    }
    
}