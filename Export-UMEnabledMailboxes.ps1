Function Export-UMEnabledMailboxesNew {
    
<#
	.SYNOPSIS
		Function intended for export all UM enabled mailboxes to csv file. Default output directory is subdirectory with a name "UMMailboxes"
	
	.DESCRIPTION
		A detailed description of the Export-UMEnabledMailboxesNew function.
	
	.PARAMETER OutputFileDirectoryPath
		Directory path where outputfile need to be saved, if directory doesn't exists will be created
	
	.PARAMETER OutputFileNamePrefix
		Name prefix for the output file name
	
	.PARAMETER StartTimeSuffix
		A description of the StartTimeSuffix parameter.
	
	.PARAMETER ResolveLinkedMasterAccounts
		A description of the ResolveLinkedMasterAccounts parameter.
	
	.PARAMETER LinkedDomainSearchBase
		A description of the LinkedDomainSearchBase parameter.
	
	.PARAMETER ResultSize
		A description of the ResultSize parameter.
	
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
		0.2.0 - 2016-03-08 - Added resolving LinkedMasterAccount feature - need to be tested
		
		
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
        [String]$StartTimeSuffix,
        [parameter(Mandatory = $false)]
        [Bool]$ResolveLinkedMasterAccounts = $true,
        [parameter(Mandatory = $false)]
        [String]$LinkedDomainSearchBase,
        [Parameter(Mandatory = $false)]
        $ResultSize = 100
        
    )
    
    BEGIN {
        
        if ($ResolveLinkedMasterAccounts) {
            
            if ((Get-Module -name 'ActiveDirectory' -ErrorAction SilentlyContinue) -eq $null) {
                
                Import-Module -Name 'ActiveDirectory' -ErrorAction Stop | Out-Null
                
            }
            
        }
        
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
        
        $UMMailboxes = Get-UMMailbox -ResultSize $ResultSize | select Name, DisplayName, guid, LinkedMasterAccount, Description, UMEnabled, PrimarySMTPAddress, Extensions
        
        $UMMailboxesCount = $($UMMailboxes | Measure-Object).Count
        
        Write-Verbose -Message "$UMMailboxesCount read"
        
        $Results = @()
        
        $i = 1
        
    }
    
    PROCESS {
        
        $UMMailboxes | ForEach {
            
            $CurrentLinkedMasterDomain = $null
            
            $CurrentLinkedSAMAccountName = $null
            
            $PercentCompleted = [math]::Round(($i / $UMMailboxesCount) * 100)
            
            $StatusText = "Percent completed $PercentCompleted%, currently the {0} is checked. " -f $($_).ToString()
            
            Write-Progress -Activity "Gathering data for . " -Status $StatusText -PercentComplete $PercentCompleted
            
            $CurrentUMMailbox = $_
            
            If ($ResolveLinkedMasterAccounts -and $CurrentUMMailbox.LinkedMasterAccount -ne $null) {
                
                [String]$LinkedMasterAccountString = ($CurrentUMMailbox.LinkedMasterAccount).ToString()
                
                [Int]$BackslashPosition = $LinkedMasterAccountString.IndexOf("\")
                
                [String]$CurrentLinkedMasterDomain = $LinkedMasterAccountString.Substring(0, $BackslashPosition)
                
                [String]$CurrentLinkedSAMAccountName = ($LinkedMasterAccountString.Substring($BackslashPosition + 1, $LinkedMasterAccountString.Length - $BackslashPosition - 1))
                
                [String]$DCName = "{0}DC" -f $CurrentLinkedMasterDomain
                
                #Try {
                
                $LinkedDomainController = ((Get-Variable -Name $DCName).Value) | Out-Null
                
                Write-Host $LinkedDomainController + " in Try"
                
                #}
                
                #Catch {
                
                $CurrentLinkedMasterDomainController = (Get-ADDomainController -DomainName $CurrentLinkedMasterDomain -Discover).HostName
                
                Write-Host $CurrentLinkedMasterDomainController
                
                New-Variable -Name $DCName -Value $CurrentLinkedMasterDomainController
                
                $LinkedDomainController = $((Get-Variable -Name $DCName).Value)
                
                #}
                
                
                #Finally {
                
                
                Write-Verbose -Message "Trying find user $CurrentLinkedSAMAccountName in the Active Directory domain $CurrentLinkedMasterDomain using the server $LinkedDomainController"
                
                $LinkedADUser = Get-ADUser -Identity "$CurrentLinkedSAMAccountName" -Server $LinkedDomainController -Properties enabled
                
                
                #}
                
            }
            
            $SortedExtensions = ($CurrentUMMailbox | Select -ExpandProperty Extensions | Sort)
            
            $Result = New-Object PSObject
            
            $Result | Add-Member -type NoteProperty -name MailboxAlias -value $CurrentUMMailbox.Name
            
            $Result | Add-Member -type NoteProperty -name MailboxDisplayName -value $CurrentUMMailbox.DisplayName
            
            $Result | Add-Member -type NoteProperty -name MailboxGuid -value $CurrentUMMailbox.Guid
            
            $Result | Add-Member -type NoteProperty -name LinkedMasterAccount -value $CurrentUMMailbox.LinkedMasterAccount
            
            $Result | Add-Member -Type NoteProperty -Name LinkedMasterDomain -Value $CurrentLinkedMasterDomain
            
            $Result | Add-Member -Type NoteProperty -Name LikedMasterSAMAccountName -Value $CurrentLinkedSAMAccountName
            
            $Result | Add-Member -Type NoteProperty -Name LinkedMasterAccountEnabled -Value $LinkedADUser.Enabled
            
            $Result | Add-Member -Type NoteProperty -Name LinkedMasterAccountDescription -Value $LinkedADUser.Description
            
            $Result | Add-Member -type NoteProperty -name PrimarySMTPAddress -value $CurrentUMMailbox.PrimarySMTPAddress
            
            $Result | Add-Member -type NoteProperty -name Description -value $CurrentUMMailbox.Description
            
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