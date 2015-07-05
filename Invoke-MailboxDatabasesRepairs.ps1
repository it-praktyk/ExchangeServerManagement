﻿function Invoke-MailboxDatabasesRepairs {
    
    <#
	.SYNOPSIS
	Function intended for 
   
	.DESCRIPTION
        
    White Paper: Database Integrity Checking in Exchange Server 2010 SP1
    https://technet.microsoft.com/en-us/library/hh547017%28v=exchg.141%29.aspx
	
  	.PARAMETER FirstParameter
  
	.EXAMPLE
     
	.LINK
	https://github.com/it-praktyk/Invoke-MailboxDatabasesRepairs
	
	.LINK
	https://www.linkedin.com/in/sciesinskiwojciech
	      
	.NOTES
	AUTHOR: Wojciech Sciesinski, wojciech[at]sciesinski[dot]net
	KEYWORDS: PowerShell, Exchange, 
   
	VERSIONS HISTORY
	0.1.0 - 2015-07-05 - Initial release
    0.1.1 - 2015-07-06 - Help updated, TO DO updated
    
    DEPENDENCIES
    -   Function Test-ExchangeCmdletsAvailability - minimum 0.1.2
        https://github.com/it-praktyk/Test-ExchangeCmdletsAvailability
    -   Function Function Get-EventsBySource - minimum 0.3.2
        https://github.com/it-praktyk/Get-EvenstBySource

	TODO
    - mail summary report
    - parse output for application events 10062
    - Exchange Server version checking (at least 2010 SP1 need to be)
    - Add support for other errors - https://blogs.it.ox.ac.uk/nexus/2012/06/11/new-mailboxrepairrequest/
		
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
    
    [cmdletbinding()]
    
    param
    (
        [parameter(Mandatory = $false)]
        [String]$ComputerName = 'localhost',
        
        [parameter(Mandatory = $false)]
        $CorruptionType = @("SearchFolder", "AggregateCounts", "ProvisionedFolder", "FolderView", "MessagePTagCn"),
        
        [parameter(Mandatory = $false)]
        $Database = "All",
        
        [parameter(Mandatory = $false)]
        [switch]$DetectOnly = $false,
        
        [parameter(Mandatory = $false)]
        [Int]$CheckProgressEverySeconds = 30,
        
        [parameter(Mandatory = $false)]
        [switch]$DisplaySummary = $false
                
    )
    
    Begin {
        
        $ActiveDatabases = @()
        
        If ($ComputerName -eq 'localhost') {
            
            $ComputerFQDNName = ([Net.DNS]::GetHostEntry("localhost")).HostName
            
            $ComputerNetBIOSName = $ComputerFQDNName.Split(".")[0]
            
        }
        elseif ($ComputerName.Contains(".")) {
            
            $ComputerNetBIOSName = $ComputerName.Split(".")[0]
            
        }
        
        
        If ((Test-ExchangeCmdletsAvailability) -ne $true) {
            
            Throw "The function Invoke-MailboxDatabasesReapairs need to be run using Exchange Management Shell"
        }
        
        
        #Test if target Exchange server is availble need to be add (?)
        
        If ($Database -eq 'All') {
            
            $ActiveDatabases = (Get-MailboxDatabase -Server $ComputerNetBIOSName | where { $_.Server -eq $ComputerNetBIOSName } | select Name)
            
            $ActiveDatabasesCount = ($ActiveDatabases | measure).Count
            
        }
        Else {
            
            $Database | foreach {
                
                [Bool]$waserror = $false
                
                Try {
                    
                    $CurrentDatabase = (Get-MailboxDatabase -Identity $_ -Server $ComputerNetBIOSName | where { $_.Server -eq $ComputerNetBIOSName } | select Name)
                    
                }
                Catch {
                    
                    $waserror = $true
                    
                    [String]$MessageText = "Database {0} is not corrently active on {1} and can't be checked" -f $CurrentDatabase, $ComputerNetBIOSName
                    Write-Error -Message $MessageText
                    
                }
                
                Finally {
                    
                    if (-not $waserror) {
                        
                        $ActiveDatabases += $CurrentDatabase
                        
                    }
                    
                }
                
            }
            
            $ActiveDatabasesCount = ($ActiveDatabases | measure).Count
            
        }
        
        If ($ActiveDatabasesCount -lt 1) {
            
            [String]$MessageText = "Any database was not found on the server {0}" -f $ComputerNetBIOSName
            
            Write-Verbose $MessageText
            
        }
        
    }
    
    Process {
        
        $ActiveDatabases | foreach {
            
            $StartRepairEventFound = $false
            
            $StopRepairEventFound = $false
            
            $StartTimeForDatabase = Get-Date
            
            $CurrentRepairRequest = New-MailboxRepairRequest -Database $_.Name -CorruptionType $($CorruptionType -join ",") -DetectOnly:$DetectOnly
            
            Start-Sleep -Seconds 1
            
            do {
                
                $StartRepairEvent = Get-EventsBySource -ComputerName $ComputerFQDNName -LogName "Application" -ProviderName "MSExchangeIS Mailbox Store" -EventID 10059 -StartTime $StartTimeForDatabase
                
                $StartRepairEventFound = (($StartRepairEvent | measure).count -eq 1)
                
                If (-not $StartRepairEventFound) {
                    
                    Start-Sleep -Seconds $CheckProgressEverySeconds
                    
                }
                
                Else {
                    
                    [String]$MessageText = "Repair request for database {0} started at {1}" -f $_.Name, $StartRepairEvent.TimeGenerated
                    
                    Write-Verbose -Message $MessageText
                    
                }
                
            }
            while ($StartRepairEventFound -eq $false)
            
            Start-Sleep -Seconds 1
            
            do {
                
                $StopRepairEvent = Get-EventsBySource -ComputerName $ComputerFQDNName -LogName "Application" -ProviderName "MSExchangeIS Mailbox Store" -EventID 10048 -StartTime $StartTimeForDatabase
                
                $StopRepairEventFound = (($StopRepairEvent | measure).count -eq 1)
                
                If (-not $StopRepairEventFound) {
                    
                    Start-Sleep -Seconds $CheckProgressEverySeconds
                    
                }
                Else {
                    
                    [String]$MessageText = "Repair request for database {0} end at {1}" -f $_.Name, $StopRepairEvent.TimeGenerated
                    
                    Write-Verbose -Message $MessageText
                    
                    If ($DisplaySummary) {
                        
                        $CorruptionFoundEvents = Get-EventsBySource -ComputerName $ComputerFQDNName -LogName "Application" -ProviderName "MSExchangeIS Mailbox Store" -EventID 10062 -StartTime $StartTimeForDatabase
                        
                        $CorruptionFoundEventsCount = ($CorruptionFoundEvents | measure).count
                        
                        If ($CorruptionFoundEventsCount -ge 1) {
                            
                            Write-Output $CorruptionFoundEvents
                            
                        }
                        
                    }
                    
                }
                
                
            }
            while ($StopRepairEventFound -eq $false)
            
        }
        
    }
    
    End {
        
    }
    
}

Function Test-ExchangeCmdletsAvailability {
    
<#
    .SYNOPSIS
    Function intended for veryfing if in current PowerShell session Exchange cmdlets and Exchange servers are available
   
    .DESCRIPTION
    
    .PARAMETER CmdletForCheck
    Cmdlet which availability will be tested
    
    .PARAMETER CheckExchangeServersAvailability
    Try read list of available Exchange servers
  
    .EXAMPLE
    
    Test-ExchangeCmdletsAvailability -CmdletForCheck Get-Mailbox
     
    .LINK
    https://github.com/it-praktyk/Test-ExchangeCmdletsAvailability
    
    .LINK
    https://www.linkedin.com/in/sciesinskiwojciech
          
    .NOTES
    AUTHOR: Wojciech Sciesinski, wojciech[at]sciesinski[dot]net
    KEYWORDS: PowerShell,Exchange
   
    VERSIONS HISTORY
    0.1.0 - 2015-05-25 - Initial release
    0.1.1 - 2015-05-25 - Variable renamed, help updated, simple error handling added
    0.1.2 - 2015-07-06 - Corrected

    TODO
        
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
        [String]$CmdletForCheck = "Get-ExchangeServer",
        
        [parameter(Mandatory = $false)]
        [Bool]$CheckExchangeServersAvailability = $false
        
    )
    
    BEGIN {
        
        
        
        
    }
    
    PROCESS {
        
        $CmdletAvailable = Test-Path -Path Function:$CmdletForCheck
        
        if ($CmdletAvailable -and $CheckExchangeServersAvailability) {
            
            Try {
                
                $ReturnedServers = Get-ExchangeServer
                
                $ReturnedServerCount = ($ReturnedServers | measure).Count
                
                $Result = ($CmdletAvailable -and ($ReturnedServerCount -ge 1))
            }
            
            Catch {
                
                $Result = $false
                
            }
            
        }
        Else {
            
            $Result = $CmdletAvailable
            
        }
        
    }
    
    END {
        
        Return $Result
        
    }
    
}

Function Get-EventsBySource {
<#
	.SYNOPSIS
	Function intended for remote gathering events data 
  
	.PARAMETER ComputerName
   
	.PARAMETER LogName
	
	.PARAMETER ProviderName
	
	.PARAMETER EventID
	
	.PARAMETER ConcatenateMessageLines
	
	.PARAMETER ConcatenatedLinesSeparator
	
	.PARAMETER MessageCharsAmount
	
     
    .EXAMPLE
    Get-EventsBySource
         
    .LINK
    https://github.com/it-praktyk/Get-EvenstBySource
    
    .LINK
    https://www.linkedin.com/in/sciesinskiwojciech
          
    .NOTES
   
    AUTHOR: Wojciech Sciesinski, wojciech.sciesinski@atos.net
    KEYWORDS: Windows, Event logs
    VERSION HISTORY
    0.3.1 - 2015-07-03 - Support for time span corrected, the first version published on GitHub
    0.3.2 - 2015-07-05 - Help updated, function corrected
    

    TODO
    - help update needed
    
        
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
        [parameter(mandatory = $true)]
        [String]$ComputerName,
        
        [parameter(mandatory = $true)]
        [String]$LogName,
        
        [parameter(mandatory = $true)]
        [String]$ProviderName,
        
        [parameter(mandatory = $true)]
        [Int]$EventID,
        
        [parameter(mandatory = $false, ParameterSetName = "StartEndTime")]
        [DateTime]$StartTime,
        
        [parameter(mandatory = $false, ParameterSetName = "StartEndTime")]
        [DateTime]$EndTime,
        
        [parameter(mandatory = $false, ParameterSetName = "ForLast")]
        [int]$ForLastTimeSpan = 24,
        
        [parameter(mandatory = $false, ParameterSetName = "ForLast")]
        [ValidateSet("minutes", "hours", "days")]
        [string]$ForLastTimeUnit = "hours",
        
        [parameter(mandatory = $false)]
        [Bool]$ConcatenateMessageLines = $true,
        
        [parameter(mandatory = $false)]
        [String]$ConcatenatedLinesSeparator = "^",
        
        [parameter(mandatory = $false)]
        [Int]$MessageCharsAmount = -1
        
    )
    
    BEGIN {
        
        $Results = @()
        
    }
    
    PROCESS {
        
        $SkipServer = $false
        
        Try {
            
            Write-Verbose -Message "Checking logs on the server $ComputerName"
            
            If ($StartTime -ne $null -or $EndTime -ne $null) {
                
                If ($StartTime -and $EndTime) {
                    
                    [Array]$FilterHashTable = @{ "Logname" = $LogName; "Id" = $EventID; "ProviderName" = $ProviderName; "StartTime" = $StartTime; "EndTime" = $EndTime }
                    
                }
                Elseif ($EndTime) {
                    
                    [Array]$FilterHashTable = @{ "Logname" = $LogName; "Id" = $EventID; "ProviderName" = $ProviderName; "EndTime" = $EndTime }
                    
                }
                Else {
                    
                    [Array]$FilterHashTable = @{ "Logname" = $LogName; "Id" = $EventID; "ProviderName" = $ProviderName; "StartTime" = $StartTime }
                    
                }
                
            }
            
            elseif ($ForLastTimeSpan -ne $null -or $ForLastTimeSpan -ne $null) {
                
                $EndTime = Get-Date
                
                switch ($ForLastTimeUnit) {
                    "minutes" {
                        
                        $StartTime = $EndTime.AddMinutes(- $ForLastTimeSpan)
                        
                    }
                    "hours" {
                        
                        $StartTime = $EndTime.AddHours(- $ForLastTimeSpan)
                        
                    }
                    "days" {
                        
                        $StartTime = $EndTime.AddDays(- $ForLastTimeSpan)
                        
                    }
                    
                }
                
                [Array]$FilterHashTable = @{ "Logname" = $LogName; "Id" = $EventID; "ProviderName" = $ProviderName; "StartTime" = $StartTime; "EndTime" = $EndTime }
                
                
            }
            
            Else {
                
                [Array]$FilterHashTable = @{ "Logname" = $LogName; "Id" = $EventID; "ProviderName" = $ProviderName }
                
            }
            
            $Events = $(Get-WinEvent -ComputerName $ComputerName -FilterHashtable $FilterHashTable -ErrorAction 'SilentlyContinue' | Select-Object -Property MachineName, Providername, ID, TimeCreated, Message)
            
        }
        
        Catch {
            
            Write-Verbose -Message "Computer $ComputerName not accessible or error with access to $LogName event log."
            
            [Bool]$SkipServer = $true
            
        }
        
        Finally {
            
            
            If (-not $SkipServer) {
                
                
                $Found = $($Events | Measure-Object).Count
                
                If ($Found -ne 0) {
                    
                    [String]$MessageText = "For the computer $ComputerName events $Found found"
                    
                    Write-Verbose -Message $MessageText
                    
                    $Events | ForEach  {
                        
                        $Result = New-Object -TypeName PSObject
                        $Result | Add-Member -type NoteProperty -name ComputerName -value $_.MachineName
                        $Result | Add-Member -type NoteProperty -name Source -value $_.Providername
                        $Result | Add-Member -type NoteProperty -name EventID -Value $_.ID
                        $Result | Add-Member -type NoteProperty -name TimeGenerated -Value $_.TimeCreated
                        
                        $MessageLength = $($_.Message).Length
                        
                        If (($MessageCharsAmount -eq -1) -or $MessageCharsAmount -gt $MessageLength) {
                            
                            $MessageCharsAmount = $MessageLength
                            
                        }
                        
                        if ($ConcatenateMessageLines) {
                            
                            $MessageFields = $_.Message.Substring(0, $MessageCharsAmount - 1).Replace("`r`n", $ConcatenatedLinesSeparator)
                            
                            $Result | Add-Member -type NoteProperty -name Message -Value $MessageFields
                            
                        }
                        else {
                            
                            $Result | Add-Member -type NoteProperty -name Message -Value $_.Message.Substring(0, $MessageCharsAmount - 1)
                            
                        }
                        
                        $Results += $Result
                        
                    }
                    
                }
                
            }
            
        }
        
    }
    
    
    END {
        
        Return $Results
        
    }
    
}