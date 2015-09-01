function Invoke-MailboxDatabasesRepairs {
    
	<#
	.SYNOPSIS
	Function intended for 
   
	.DESCRIPTION

	The Exchange Team Blog: New Support Policy for Repaired Exchange Databases
	http://blogs.technet.com/b/exchange/archive/2015/05/01/new-support-policy-for-repaired-exchange-databases.aspx
	
	White Paper: Database Integrity Checking in Exchange Server 2010 SP1
	https://technet.microsoft.com/en-us/library/hh547017%28v=exchg.141%29.aspx
	   
	Nexus: News, Messages about messaging, Matthew Gaskin blog
	Using the New-MailboxRepairRequest cmdlet
	https://blogs.it.ox.ac.uk/nexus/2012/06/11/new-mailboxrepairrequest/
	
	.PARAMETER ComputerName
	
	.PARAMETER Database
	
	.PARAMETER DetectOnly
	
	.PARAMETER CheckProgressEverySeconds
	
	.PARAMETER DisplaySummary
	
	.PARAMETER DisplayProgressBar
	
	.PARAMETER ExpectedDurationTimeMinutes
	
	.PARAMETER CreateReportFile
    By default report file is created
    
    .PARAMETER ReportFileDirectoryPath
    By default report files are stored in subfolder "reports" in current path, if "reports" subfolder is missed will be created
    
    .PARAMETER ReportFileNamePrefix
    Prefix used for creating errors report files name. Default is "Report-" 

	.
  
	.EXAMPLE
	
	[PS] >Invoke-MailboxDatabasesRepairs -ComputerName XXXXXXMBX03 -Database All -DisplaySummary:$true -ExpectedDurationTimeMinutes 120 -DetectOnly:$true
	 
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
	0.1.2 - 2015-07-15 - Progress bar added, verbose messages partially suppressed, help next update
	0.1.3 - 2015-08-11 - Additional checks added to verify provided Exchange server, help and TO DO updated
	0.2.0 - 2015-08-31 - Corrected checking of Exchange version, output redirected to per mailbox database reports
	
	DEPENDENCIES
	-	Function Test-ExchangeCmdletsAvailability - minimum 0.1.2
		https://github.com/it-praktyk/Test-ExchangeCmdletsAvailability
	-	Function Function Get-EventsBySource - minimum 0.3.2
		https://github.com/it-praktyk/Get-EvenstBySource
	-	Function New-ReportFileNameFullPath - minimum 0.1.1
		https://github.com/it-praktyk/New-ReportFileNameFullPath

	TODO
	- additional events need to be checked
		a) 10045 -	The database repair request failed for provisioned folders. This event ID is created in conjunction with event ID 10049
		b) 10049 -	The mailbox or database repair request failed because Exchange encountered a problem with the database or another task 
					is running against the database. (Fix for this is ESEUTIL then contact Microsoft Product Support Services)
		c) 10050 -	The database repair request couldn’t run against the database because the database doesn’t support the corruption types 
					specified in the command. This issue can occur when you run the command from a server that’s running a later version 
					of Exchange than the database you’re scanning.
	
		d) 10051 -	The database repair request was cancelled because the database was dismounted.
	- store and/or mail summary report
	- parse output for application events 10062
	- Exchange Server version checking (at least 2010 SP1 need to be)
		
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
        
        #This parameter is not currently used - 
        [parameter(Mandatory = $false)]
        $CorruptionType = @("SearchFolder", "AggregateCounts", "ProvisionedFolder", "FolderView", "MessagePTagCn"),
        [parameter(Mandatory = $false)]
        $Database = "All",
        [parameter(Mandatory = $false)]
        [switch]$DetectOnly = $false,
        [parameter(Mandatory = $false)]
        [Int]$CheckProgressEverySeconds = 120,
        [parameter(Mandatory = $false)]
        [switch]$DisplaySummary = $false,
        [parameter(Mandatory = $false)]
        [switch]$DisplayProgressBar = $true,
        [Parameter(mandatory = $false, Position = 3)]
        [int]$ExpectedDurationTimeMinutes = 15,
        [parameter(Mandatory = $false)]
        [Bool]$CreateReportFile = $true,
        [parameter(Mandatory = $false)]
        [String]$ReportFileDirectoryPath = ".\reports\",
        [parameter(Mandatory = $false)]
        [String]$ReportFileNamePrefix = "Report-",
        [parameter(Mandatory = $false)]
        [String]$ReportFileNameMidPart,
        [parameter(Mandatory = $false)]
        [Switch]$IncludeDateTimePartInFileName = $true,
        [parameter(Mandatory = $false)]
        [String]$DateTimePartInFileName,
        [parameter(Mandatory = $false)]
        [String]$ReportFileNameExtension = ".csv"
        
        
        
    )
    
    Begin {
        
        $ActiveDatabases = @()
        
        $EventsToReport = @()
        
        
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
        
        $MailboxServer = (Get-MailboxServer -Identity $ComputerNetBIOSName)
        
        $MailboxServerCount = ($MailboxServer | Measure).Count
        
        
        If ($MailboxServerCount -gt 1) {
            
            [String]$MessageText = "You can use this function to perform actions at only one server at once."
            
            Throw $MessageText
            
        }
        Elseif ($MailboxServerCount -ne 1) {
            
            [String]$MessageText = "Server {0} is not a Exchange mailbox server" -f $MailboxServer
            
            Throw $MessageText
        }
        
        #Determine Exchange server version
        #List of build version: https://technet.microsoft.com/library/hh135098.aspx
        Try {
            
            if ($ComputerName -match 'localhost') {
                
                $ExchangeSetupFileVersion = Get-Command Exsetup.exe | select FileversionInfo
                
            }
            Else {
                
                $ExchangeSetupFileVersion = Invoke-Command -ComputerName $MailboxServer -ScriptBlock { Get-Command Exsetup.exe | select FileversionInfo }
                
            }
            
            [Version]$MailboxServerVersion = ($ExchangeSetupFileVersion.FileVersionInfo).FileVersion
            
        }
        Catch {
            
            [String]$MessageText = "Server {0} is not reachable or PowerShell remoting is not enabled on it."
            
            #Decission based on 
            
        }
        
        Finally {
            
            If (($ExchangeSetupFileVersion.Major -eq 14 -and $ExchangeSetupFileVersion.Minor -lt 1) -or $ExchangeSetupFileVersion.Major -lt 14) {
                
                [String]$MessageText = "This function can be used only on Exchange Server 2010 SP1 or newer version."
                
                Throw $MessageText
                
            }
            
        }
        
        
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
                    
                    [String]$MessageText = "Database {0} is not currently active on {1} and can't be checked" -f $CurrentDatabase, $ComputerNetBIOSName
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
            
            If ($CreateReportFile) {
                
                
                
            }
            
            $StartRepairEventFound = $false
            
            $StopRepairEventFound = $false
            
            $StartTimeForDatabase = Get-Date
            
            $CurrentRepairRequest = New-MailboxRepairRequest -Database $_.Name -CorruptionType "SearchFolder", "AggregateCounts", "ProvisionedFolder", "FolderView", "MessagePTagCn" -DetectOnly:$DetectOnly
            
            Start-Sleep -Seconds 1
            
            do {
                
                $StartRepairEvent = Get-EventsBySource -ComputerName $ComputerFQDNName -LogName "Application" -ProviderName "MSExchangeIS Mailbox Store" -EventID 10059 -StartTime $StartTimeForDatabase -Verbose:$false
                
                $StartRepairEventFound = (($StartRepairEvent | measure).count -eq 1)
                
                If (-not $StartRepairEventFound) {
                    
                    Start-Sleep -Seconds $CheckProgressEverySeconds
                    
                }
                
                Else {
                    
                    [String]$MessageText = "Repair request for database {0} started at {1}" -f $_.Name, $StartRepairEvent.TimeGenerated
                    
                    [String]$StartTime = Get-Date $($StartRepairEvent.TimeGenerated) -format yyyyMMdd-HHmm
                    
                    Write-Verbose -Message $MessageText
                    
                }
                
            }
            while ($StartRepairEventFound -eq $false)
            
            
            Start-Sleep -Seconds 1
            
            [int]$i = $CheckProgressEverySeconds
            
            do {
                
                $StopRepairEvent = Get-EventsBySource -ComputerName $ComputerFQDNName -LogName "Application" -ProviderName "MSExchangeIS Mailbox Store" -EventID 10048 -StartTime $StartTimeForDatabase -Verbose:$false
                
                $StopRepairEventFound = (($StopRepairEvent | measure).count -eq 1)
                
                
                If (-not $StopRepairEventFound) {
                    
                    If ($DisplayProgressBar) {
                        
                        [String]$MessageText = "Database {0} repair request	 is in progress." -f $_.Name
                        
                        Write-Progress -Activity $MessageText -Status "Completion percentage is only approximate." -PercentComplete (($i / ($ExpectedDurationTimeMinutes * 60)) * 100)
                        
                        if (($i += $CheckProgressEverySeconds) -ge ($ExpectedDurationTimeMinutes * 60)) {
                            
                            $i = $CheckProgressEverySeconds
                        }
                        Else {
                            
                            $i += $CheckProgressEverySeconds
                            
                        }
                        
                    }
                    
                    
                    Start-Sleep -Seconds $CheckProgressEverySeconds
                    
                }
                Else {
                    
                    [String]$MessageText = "Repair request for database {0} end at {1}" -f $_.Name, $StopRepairEvent.TimeGenerated
                    
                    Write-Verbose -Message $MessageText
                    
                    If ($DisplaySummary) {
                        
                        $CorruptionFoundEvents = Get-EventsBySource -ComputerName $ComputerFQDNName -LogName "Application" -ProviderName "MSExchangeIS Mailbox Store" -EventID 10062 -StartTime $StartTimeForDatabase -Verbose:$false
                        
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

Function New-ReportFileNameFullPath {
    
<#

	.SYNOPSIS
	Function intended for 
   
	.DESCRIPTION
	
	.PARAMETER CreateReportFileDirectory
	
	.PARAMETER ReportFileDirectoryPath
	
	.PARAMETER ReportFileNamePrefix
	
	.PARAMETER ReportFileNameMidPart
	
	.PARAMETER IncludeDateTimePartInFileName
	
	.PARAMETER DateTimePartInFileName
	
	.PARAMETER ReportFileNameExtension
	
	.PARAMETER CheckIfReportFileExist
	
	.PARAMETER BreakIfError

	.EXAMPLE
	
	[PS] > New-ReportFileNameFullPath 
	 
	.LINK
	https://github.com/it-praktyk/New-ReportFileNameFullPath
	
	.LINK
	https://www.linkedin.com/in/sciesinskiwojciech
		  
	.NOTES
	AUTHOR: Wojciech Sciesinski, wojciech[at]sciesinski[dot]net
	KEYWORDS: PowerShell
   
	VERSIONS HISTORY
	0.1.0 - 2015-09-01 - Initial release
    0.1.1 - 2015-09-01 - Minor update
	
	TODO
	Update help

		
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
    
    
    param (
        
        [parameter(Mandatory = $false)]
        [Switch]$CreateReportFileDirectory = $true,
        [parameter(Mandatory = $false)]
        [String]$ReportFileDirectoryPath = ".\reports\",
        [parameter(Mandatory = $false)]
        [String]$ReportFileNamePrefix = "Report-",
        [parameter(Mandatory = $false)]
        [String]$ReportFileNameMidPart,
        [parameter(Mandatory = $false)]
        [Switch]$IncludeDateTimePartInFileName = $true,
        [parameter(Mandatory = $false)]
        [String]$DateTimePartInFileName,
        [parameter(Mandatory = $false)]
        [String]$ReportFileNameExtension = ".csv",
        [parameter(Mandatory = $false)]
        [Switch]$CheckIfReportFileExist = $true,
        [parameter(Mandatory = $false)]
        [Switch]$BreakIfError = $true
        
    )
    
    #Declare variable
    
    [Int]$ExitCode = 0
    
    [String]$ErrorDescription = $null
    
    $Result = New-Object PSObject
    
    #Convert relative path to absolute path
    [String]$ReportFileDirectoryPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ReportFileDirectoryPath)
    
    #Assign value to the variable $IncludeDateTimePartInFileName if is not initialized
    If ($IncludeDateTimePartInFileName -and $DateTimePartInFileName -eq "") {
        
        [String]$DateTimePartInFileName = $(Get-Date -format yyyyMMdd-HHmm)
        
    }
    
    #Check if report directory exist and try create if not
    
    If ($CreateReportFileDirectory -and !$((Get-Item -Path $ReportFileDirectoryPath -ErrorAction SilentlyContinue) -is [system.io.directoryinfo])) {
        
        Try {
            
            $ErrorActionPreference = 'Stop'
            
            New-Item -Path $ReportFileDirectoryPath -type Directory | Out-Null
            
        }
        Catch {
            
            [String]$MessageText = "Provided path {0} doesn't exist and can't be created" -f $ReportFileDirectoryPath
            
            If ($BreakIfError) {
                
                Throw $MessageText
                
            }
            Else {
                
                Write-Error -Message $MessageText
                
                [Int]$ExitCode = 1
                
                [String]$ErrorDescription = $MessageText
                
            }
            
        }
        
    }
    ElseIf (!$((Get-Item -Path $ReportFileDirectoryPath -ErrorAction SilentlyContinue) -is [system.io.directoryinfo])) {
        
        [String]$MessageText = "Provided patch {0} doesn't exist and value for the parameter CreateReportFileDirectory is set to False" -f $ReportFileDirectoryPath
        
        If ($BreakIfError) {
            
            Throw $MessageText
            
        }
        Else {
            
            Write-Error -Message $MessageText
            
            [Int]$ExitCode = 2
            
            [String]$ErrorDescription = $MessageText
            
        }
        
    }
    
    #Try if report directory is writable - temporary file is stored
    Try {
        
        $ErrorActionPreference = 'Stop'
        
        [String]$TempFileName = [System.IO.Path]::GetTempFileName() -replace '.*\\', ''
        
        [String]$TempFilePath = "{0}{1}" -f $ReportFileDirectoryPath, $TempFileName
        
        New-Item -Path $TempFilePath -type File | Out-Null
        
    }
    Catch {
        
        [String]$MessageText = "Provided patch {0} is not writable" -f $ReportFileDirectoryPath
        
        If ($BreakIfError) {
            
            Throw $MessageText
            
        }
        Else {
            
            Write-Error -Message $MessageText
            
            [Int]$ExitCode = 3
            
            [String]$ErrorDescription = $MessageText
            
        }
        
    }
    
    Remove-Item $TempFilePath -ErrorAction SilentlyContinue | Out-Null
    
    
    #Constructing the file name
    If (!($IncludeDateTimePartInFileName) -and ($ReportFileNameMidPart -ne $null)) {
        
        [String]$ReportFilePathTemp = "{0}\{1}-{2}.{3}" -f $ReportFileDirectoryPath, $ReportFileNamePrefix, $ReportFileNameMidPart, $ReportFileNameExtension
        
    }
    Elseif (!($IncludeDateTimePartInFileName) -and ($ReportFileNameMidPart -eq $null)) {
        
        [String]$ReportFilePathTemp = "{0}\{1}.{2}" -f $ReportFileDirectoryPath, $ReportFileNamePrefix, $ReportFileNameExtension
        
    }
    ElseIf ($IncludeDateTimePartInFileName -and ($ReportFileNameMidPart -ne $null)) {
        
        [String]$ReportFilePathTemp = "{0}\{1}-{2}-{3}.{4}" -f $ReportFileDirectoryPath, $ReportFileNamePrefix, $ReportFileNameMidPart, $DateTimePartInFileName, $ReportFileNameExtension
        
    }
    Else {
        
        [String]$ReportFilePathTemp = "{0}\{1}-{2}.{3}" -f $ReportFileDirectoryPath, $ReportFileNamePrefix, $DateTimePartInFileName, $ReportFileNameExtension
        
    }
    
    #Replacing doubled chars \\ , -- , ..
    [String]$ReportFilePath = "{0}{1}" -f $ReportFilePathTemp.substring(0, 2), (($ReportFilePathTemp.substring(2, $ReportFilePathTemp.length - 2).replace("\\", '\')).replace("--", "-")).replace("..", ".")
    
    If ($CheckIfReportFileExist -and (Test-Path -Path $ReportFilePath -PathType Leaf)) {
        
        [String]$MessageText = "The file {0} already exist" -f $ReportFilePath
        
        If ($BreakIfError) {
            
            Throw $MessageText
            
        }
        Else {
            
            Write-Error -Message $MessageText
            
            [Int]$ExitCode = 4
            
            [String]$ErrorDescription = $MessageText
            
        }
    }
    
    $Result | Add-Member -MemberType NoteProperty -Name ExitCode -Value
    
    $Result | Add-Member -MemberType NoteProperty -Name ReportFilePath -Value $ReportFilePath
    
    Return $Result
    
}


Function Parse10062Events {
    
    param (
        
        [parameter(mandatory = $true)]
        $Events
        
    )
    
    
    $CorruptionFoundEvents = $Events
    
    $Results = @()
    
    $option = [System.StringSplitOptions]::RemoveEmptyEntries
    
    $CorruptionFoundEvents | ForEach {
        
        If ($_.EventID -eq 10062) {
            
            $separator = '^'
            
            $MessageLines = ($_.Message).Split($separator, $option)
            
            Write-Verbose "Lines to parse ($MessageLines | Measure).count"
            
            [Int]$i = 1
            
            $MessageLines | ForEach {
                
                [String]$Line = $_
                
                $Result = New-Object -TypeName PSObject
                
                If ($i -eq 2) {
                    
                    $ColonPosition = $Line.IndexOf(":")
                    
                    $SpacePosition = $Line.IndexOf(" ")
                    
                    $LineLength = $Line.Length
                    
                    $MailboxGuid = $Line.Substring($ColonPosition + 1, $SpacePosition - $ColonPosition - 1)
                    
                    $MailboxDisplayName = ($Line.Substring($SpacePosition + 2, $LineLength - $SpacePosition - 3)).Trim()
                    
                }
                ElseIf ($i -gt 4) {
                    
                    $Result | Add-Member -type NoteProperty -name MailboxGuid -value $MailboxGuid
                    
                    $Result | Add-Member -type NoteProperty -name MailboxDisplayName -value $MailboxDisplayName
                    
                    $option = [System.StringSplitOptions]::RemoveEmptyEntries
                    
                    [String]$separator = ','
                    
                    $Fields = ($Line).Split($separator, $option)
                    
                    $f = 1
                    
                    $Fields | ForEach {
                        
                        $Field = $_
                        
                        Switch ($f) {
                            
                            1 {
                                
                                $Result | Add-Member -type NoteProperty -name CorruptionType -value $Field.trim()
                                
                            }
                            
                            2 {
                                
                                $Result | Add-Member -type NoteProperty -name IsFixed -value $Field.trim()
                                
                            }
                            
                            3 {
                                
                                $Result | Add-Member -type NoteProperty -name FID -value $Field.trim()
                                
                            }
                            
                            4 {
                                
                                $Result | Add-Member -type NoteProperty -name Property -value $Field.trim()
                                
                            }
                            
                            5 {
                                
                                $Result | Add-Member -type NoteProperty -name Resolutions -value $Field.trim()
                                
                            }
                            
                        }
                        
                        $f++
                        
                    }
                    
                    $Results += $Result
                    
                }
                
                $i++
                
            }
            
        }
        
    }
    
    Return $Results
    
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