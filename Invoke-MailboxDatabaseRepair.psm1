function Invoke-MailboxDatabaseRepair {
<#
    .SYNOPSIS
    Function intended for performing checks and repairs operation on Exchange Server 2010 SP1 (or newer) mailbox databases
   
    .DESCRIPTION
    Function invokes New-MailboxDatabaseRepair cmdlet for all active mailbox database copies on the server. Mailbox databases can be also
    provided by name in function parameter
    
    Informations about repair operations you can find on pages

    The Exchange Team Blog: New Support Policy for Repaired Exchange Databases
    http://blogs.technet.com/b/exchange/archive/2015/05/01/new-support-policy-for-repaired-exchange-databases.aspx
    
    White Paper: Database Integrity Checking in Exchange Server 2010 SP1
    https://technet.microsoft.com/en-us/library/hh547017%28v=exchg.141%29.aspx
       
    Nexus: News, Messages about messaging, Matthew Gaskin blog
    Using the New-MailboxRepairRequest cmdlet
    https://blogs.it.ox.ac.uk/nexus/2012/06/11/new-mailboxrepairrequest/
    
    Possible events for Exchange Server 2010 SP1 and newer
    
    - normal operation
    a) 10048 -  The mailbox or database repair request completed successfully.
    b) 10059 -  A database-level repair request started.

    -errors
     a) 10045 - The database repair request failed for provisioned folders. This event ID is created in conjunction with event ID 10049
     b) 10049 - The mailbox or database repair request failed because Exchange encountered a problem with the database or another task 
                is running against the database. (Fix for this is ESEUTIL then contact Microsoft Product Support Services)
     c) 10050 - The database repair request couldn’t run against the database because the database doesn’t support the corruption types 
                specified in the command. This issue can occur when you run the command from a server that’s running a later version 
                of Exchange than the database you’re scanning.
    d) 10051 -  The database repair request was cancelled because the database was dismounted.
    
    .PARAMETER ComputerName
    Exchange server for which actions should be performed - need to be a mailbox server
    
    .PARAMETER Database
    Database identifier - e.g. name - for which action need to be performed. If more than one identifiers need to be separated by commas
    
    .PARAMETER DetectOnly
    Set to TRUE if any repair action shouldn't be started
    
    .PARAMETER DisplayProgressBar
    If function is used in interactive mode progress bar can be displayed to provide overall information that something is happend. 
    
    .PARAMETER CheckProgressEverySeconds
    Set interval for progress checking, by default operation progress is checked every 120 seconds
        
    .PARAMETER DisplaySummary
    Set to TRUE if summary should be displayed - summary will contain data about performed operations
        
    .PARAMETER ExpectedDurationTimeMinutes
    Time in minutes used for displaing progress bar
    
    .PARAMETER CreateReportFile
    By default report file per server is created
    
    .PARAMETER ReportFileDirectoryPath
    By default report files are stored in subfolder "reports" in current path, if "reports" subfolder is missed will be created
    
    .PARAMETER ReportFileNamePrefix
    Prefix used for creating report files name. Default is "MBDBs_IntegrityChecks_<SERVER_NETBIOS_NAME>"
    
    .PARAMETER ReportFileNameMidPart
    Part of the name which will be used in midle of name
    
    .PARAMETER IncludeDateTimePartInReportFileName
    Set to TRUE if report file name should contains part based on current date and time - format yyyyMMdd-HHmm is used
    
    .PARAMETER DateTimePartInReportFileName
    Set to date and time which should be used in report file name, by default current date and time is used
    
    .PARAMETER ReportFileNameExtension
    Set to extension which need to be used for report file, by default ".txt" is used
    
    .PARAMETER BreakOnReportCreationError
    Break function execution if parameters provided for report file creation are not correct or destination file path is not writables
    
    .EXAMPLE
    
    [PS] >Invoke-MailboxDatabaseRepair -ComputerName XXXXXXMBX03 -Database All -DisplaySummary:$true -ExpectedDurationTimeMinutes 120 -DetectOnly:$true
     
    .LINK
    https://github.com/it-praktyk/Invoke-MailboxDatabaseRepair
    
    .LINK
    https://www.linkedin.com/in/sciesinskiwojciech
          
    .NOTES
    AUTHOR: Wojciech Sciesinski, wojciech[at]sciesinski[dot]net
    KEYWORDS: PowerShell, Exchange, New-MailboxRepairRequest
   
    VERSIONS HISTORY
    0.1.0 - 2015-07-05 - Initial release
    0.1.1 - 2015-07-06 - Help updated, TO DO updated
    0.1.2 - 2015-07-15 - Progress bar added, verbose messages partially suppressed, help next update
    0.1.3 - 2015-08-11 - Additional checks added to verify provided Exchange server, help and TO DO updated
    0.2.0 - 2015-08-31 - Corrected checking of Exchange version, output redirected to per mailbox database reports
    0.3.0 - 2015-09-04 - Added support for Exchange 2013, added support for the database repair errors
    0.3.1 - 2015-09-05 - Corrected but still required testing on Exchange 2013
    0.4.0 - 2015-09-07 - Support for Exchange 2013 removed, help partially updated, report creation partially implemented
                         TODO section updated
    0.4.1 - 2015-09-08 - Function reformated
    0.5.0 - 2015-09-13 - Added support for creation per server log
    0.5.1 - 2015-09-14 - Help updated, TO DO section updated, DEPENDENCIES section updated
    0.6.0 - 2015-09-14 - Log creation capabilities updates, parsing 10062 events added
    0.6.1 - 2015-09-15 - Logging per database corrected
    0.6.2 - 2015-10-20 - Named regions partially added, function Parse10062Events corrected based on PSScriptAnalyzer rules,
                         function New-OutputFileNameFullPath updated to version 0.4.0, reports per server changed,
                         function Get-EventsBySource updated to version 0.5.0
    0.6.3 - 2015-10-21 - Date for version 0.6.2 corrected
    0.7.0 - 2015-10-21 - Reports per database changed, function corrected based on PSScriptAnalyzer rules, TO DO updated
    0.7.1 - 2015-10-22 - Logging podsystem partially updated to use PSLogging module,  function New-OutputFileNameFullPath updated to version 0.4.0 - need to be tested
    0.8.0 - 2015-10-27 - Major updates especially logging fully updated to use PSLogging module
    0.8.1 - 2015-10-28 - Corrected, tested
	0.8.2 - 2015-10-28 - Script reformated
    0.9.0 - 2015-11-11 - Script switched to module, main function renamed to Invoke-MailboxDatabaseRepair
        

    DEPENDENCIES
    -   Function Test-ExchangeCmdletsAvailability - minimum 0.1.2
        https://github.com/it-praktyk/Test-ExchangeCmdletsAvailability
    -   Function Function Get-EventsBySource - minimum 0.5.0
        https://github.com/it-praktyk/Get-EvenstBySource
    -   Function New-OutputFileNameFullPath - minimum 0.5.0
        https://github.com/it-praktyk/New-OutputFileNameFullPath
    -   Module PSLogging - minimum 3.1.0 - original author: Luca Sturlese, http://9to5it.com
        https://github.com/it-praktyk/PSLogging
    

    TO DO
    - Current time and timezone need to be compared between localhost and destination host to avoid mistakes
    - exit code return need to be implemented
    - add support for Exchange 2013 (?) and 2016 (?)
    - add named regions to easier navigation in code 
    - summary for detected corruption need to be implemented
    - summary per server e.g. checked databases need to be implemented
        
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
    
    #region parameters
    
    param
    (
        [parameter(Mandatory = $false)]
        [alias("server", "cn")]
        [String]$ComputerName = 'localhost',
        
        #This parameter is not currently used - 
        #[parameter(Mandatory = $false)]
        #$CorruptionType = @("SearchFolder", "AggregateCounts", "ProvisionedFolder", "FolderView", "MessagePTagCn"),
        
        [parameter(Mandatory = $false)]
        [String[]]$Database = 'All',
        [parameter(Mandatory = $false)]
        [Bool]$DetectOnly = $false,
        [parameter(Mandatory = $false)]
        [Bool]$DisplayProgressBar = $false,
        [parameter(Mandatory = $false)]
        [Int]$CheckProgressEverySeconds = 120,
        [parameter(Mandatory = $false)]
        [Bool]$DisplaySummary = $false,
        [Parameter(mandatory = $false)]
        [int]$ExpectedDurationTimeMinutes = 150,
        [parameter(Mandatory = $false, ParameterSetName = "Reports")]
        [ValidateSet("CreatePerServer", "CreatePerDatabase", "None")]
        [String]$CreateReportFile = "CreatePerServer",
        [parameter(Mandatory = $false, ParameterSetName = "Reports")]
        [String]$ReportFileDirectoryPath = ".\reports\",
        [parameter(Mandatory = $false, ParameterSetName = "Reports")]
        [String]$ReportFileNamePrefix,
        [parameter(Mandatory = $false, ParameterSetName = "Reports")]
        [String]$ReportFileNameMidPart,
        [parameter(Mandatory = $false, ParameterSetName = "Reports")]
        [Bool]$IncludeDateTimePartInReportFileName = $true,
        [parameter(Mandatory = $false, ParameterSetName = "Reports")]
        [DateTime]$DateTimePartInReportFileName,
        [parameter(Mandatory = $false, ParameterSetName = "Reports")]
        [String]$ReportFileNameExtension = ".txt",
        [parameter(Mandatory = $false, ParameterSetName = "Reports")]
        [Bool]$BreakOnReportCreationError = $true
        
    )
    
    #endregion
    
    Begin {
        
        #region Initialize variables
        
        [Version]$ScriptVersion = "0.8.2"
        
        $ActiveDatabases = @()
        
        $EventsToReport = @()
        
        [Bool]$WriteToFile = $false
        
        $Events10062DetailsToReport = @()
        
        [Bool]$IsRunningOnLocalhost = $false
        
        [Bool]$StartRepairEventFound = $false
        
        [Bool]$StopRepairEventFound = $false
        
        [DateTime]$StartTimeForServer = $([DateTime]::Now)
        
        #endregion
        
        #region Load External modules and dependencies
        
        
        
        #endregion
        
        #region Initialize a computer names
        
        If ($ComputerName -eq 'localhost') {
            
            $ComputerFQDNName = ([Net.DNS]::GetHostEntry("localhost")).HostName
            
            $ComputerNetBIOSName = $ComputerFQDNName.Split(".")[0]
            
        }
        elseif ($ComputerName.Contains(".")) {
            
            $ComputerNetBIOSName = $ComputerName.Split(".")[0]
            
        }
        else {
            
            $ComputerNetBIOSName = $ComputerName
            
            $ComputerFQDNName = ([Net.DNS]::GetHostEntry($ComputerName)).HostName
            
        }
        
        If (([Net.DNS]::GetHostEntry("localhost")).HostName -eq ([Net.DNS]::GetHostEntry($ComputerName)).HostName) {
            
            [Bool]$IsRunningOnLocalhost = $true
            
        }
        
        [String]$MessageText = "Resolved Exchange server names FQDN: {0} , NETBIOS: {1}" -f $ComputerFQDNName, $ComputerNetBIOSName
        
        $MessageText = Write-LogEntry -ToFile:$false -MessageType INFO -Message $MessageText -TimeStamp -ToScreen
        
        Write-Verbose -Message $MessageText
        
        #endregion
        
        #region Initialize reports files names and files
        
        #Creating name for the report, a report file will be used for save initial errors or all messages if CreatePerServer report will be selected
        
        if ($CreateReportFile -eq 'CreatePerServer') {
            
            [Bool]$WriteToFile = $true
            
            if ($PSBoundParameters.ContainsKey('$ReportFileNamePrefix')) {
                
                [String]$ReportPerServerNamePrefix = $ReportFileNamePrefix
                
            }
            Else {
                
                [String]$ReportPerServerNamePrefix = "MBDBs_IntegrityChecks_{0}" -f $ComputerNetBIOSName
                
            }
            
            $PerServerMessagesReportFile = $(New-OutputFileNameFullPath -OutputFileDirectoryPath $ReportFileDirectoryPath -OutputFileNamePrefix $ReportPerServerNamePrefix `
                                                                        -OutputFileNameMidPart $ReportFileNameMidPart `
                                                                        -IncludeDateTimePartInOutputFileName $IncludeDateTimePartInReportFileName `
                                                                        -DateTimePartInOutputFileName $StartTimeForServer `
                                                                        -OutputFileNameSuffix 'messages' -BreakIfError $BreakOnReportCreationError).OutputFilePath
            
            Start-Log -LogPath $PerServerMessagesReportFile.DirectoryName -LogName $PerServerMessagesReportFile.Name -ScriptVersion $ScriptVersion.ToString() | Out-Null
            
            $PerServerEventsReportFile = $(New-OutputFileNameFullPath -OutputFileDirectoryPath $ReportFileDirectoryPath -OutputFileNamePrefix $ReportPerServerNamePrefix `
                                                                      -OutputFileNameMidPart $ReportFileNameMidPart -OutputFileNameExtension "csv" `
                                                                      -IncludeDateTimePartInOutputFileName $IncludeDateTimePartInReportFileName `
                                                                      -DateTimePartInOutputFileName $StartTimeForServer `
                                                                      -OutputFileNameSuffix 'events' -BreakIfError $BreakOnReportCreationError).OutputFilePath
            
            $PerServersCorruptionDetailsReportFile = $(New-OutputFileNameFullPath -OutputFileDirectoryPath $ReportFileDirectoryPath -OutputFileNamePrefix $ReportPerServerNamePrefix `
                                                                                  -OutputFileNameMidPart $ReportFileNameMidPart -OutputFileNameExtension "csv" `
                                                                                  -IncludeDateTimePartInOutputFileName $IncludeDateTimePartInReportFileName `
                                                                                  -DateTimePartInOutputFileName $StartTimeForServer `
                                                                                  -OutputFileNameSuffix 'corruptions_details' -BreakIfError $BreakOnReportCreationError).OutputFilePath
            
        }
        
        #endregion
        
        [String]$MessageText = "Invoke-MailboxDatabaseRepair.ps1 started - version {0} on the server {1}" -f $ScriptVersion.ToString(), $ComputerFQDNName #,  $StartTimeForServer, "Not implemented yet :-(" #, $PSBoundParameters.GetEnumerator()
        
        $MessageText = Write-LogEntry -ToFile:$WriteToFile -LogPath $PerServerMessagesReportFile.FullName -MessageType INFO -Message $MessageText -TimeStamp -EntryDateTime $StartTimeForServer -ToScreen
        
        Write-Verbose -Message $MessageText
        
        #region Initial EMS test
        
        If ((Test-ExchangeCmdletsAvailability) -ne $true) {
            
            [String]$MessageText = "The function Invoke-MailboxDatabasesReapairs need to be run using Exchange Management Shell"
            
            $MessageText = Write-LogEntry -ToFile:$WriteToFile -LogPath $PerServerMessagesReportFile.FullName -MessageType FAILURE -Message $MessageText -TimeStamp -ToScreen
            
            Throw $MessageText
            
        }
        
        #endregion
        
        $MailboxServer = (Get-MailboxServer -Identity $ComputerNetBIOSName)
        
        [Int]$MailboxServerCount = (Measure-Object -InputObject $MailboxServer).Count
        
        If ($MailboxServerCount -gt 1) {
            
            [String]$MessageText = "You can use this function to perform actions on only one server at once."
            
            $MessageText = Write-LogEntry -ToFile:$WriteToFile -LogPath $PerServerMessagesReportFile.FullName -MessageType FAILURE -Message $MessageText -TimeStamp -ToScreen
            
            Throw $MessageText
            
        }
        Elseif ($MailboxServerCount -ne 1) {
            
            [String]$MessageText = "Server {0} is not a Exchange mailbox server" -f $MailboxServer
            
            $MessageText = Write-LogEntry -ToFile:$WriteToFile -LogPath $PerServerMessagesReportFile.FullName -MessageType FAILURE -Message $MessageText -TimeStamp -ToScreen
            
            Throw $MessageText
        }
        
        #Determine Exchange server version
        #List of build version: https://technet.microsoft.com/library/hh135098.aspx
        Try {
            
            if ($IsRunningOnLocalhost) {
                
                
                $ExchangeSetupFileVersion = Select-Object -InputObject $(Get-Command -Name Exsetup.exe) -Property FileversionInfo
                
            }
            Else {
                
                $ExchangeSetupFileVersion = Invoke-Command -ComputerName $MailboxServer.Name -ScriptBlock { Select-Object -InputObject $(Get-Command -Name Exsetup.exe) -Property FileversionInfo }
                
            }
            
            [Version]$MailboxServerVersion = ($ExchangeSetupFileVersion.FileVersionInfo).FileVersion
            
            [String]$MessageText = "Discovered version of Exchange Server: {0} " -f $MailboxServerVersion.ToString()
            
            $MessageText = Write-LogEntry -ToFile:$WriteToFile -LogPath $PerServerMessagesReportFile.FullName -MessageType INFO -Message $MessageText -TimeStamp -ToScreen
            
            Write-Verbose -Message $MessageText
            
        }
        Catch {
            
            
            [String]$MessageText = "Server {0} is not reachable or PowerShell remoting is not enabled on it." -f $ComputerFQDNName
            
            $MessageText = Write-LogEntry -ToFile:$WriteToFile -LogPath $PerServerMessagesReportFile.FullName -MessageType FAILURE -Message $MessageText -TimeStamp -ToScreen
            
            Throw $MessageText
            
        }
        
        Finally {
            
            If (($MailboxServerVersion.Major -eq 14 -and $MailboxServerVersion.Minor -lt 1) -or $MailboxServerVersion.Major -ne 14) {
                
                [String]$MessageText = "This function can be used only on Exchange Server 2010 SP1 or newer version of Exchange Server 2010."
                
                $MessageText = Write-LogEntry -ToFile:$WriteToFile -LogPath $PerServerMessagesReportFile.FullName -MessageType FAILURE -Message $MessageText -TimeStamp -ToScreen
                
                Throw $MessageText
                
            }
            
        }
        
        If ($Database -eq 'All') {
            
            $ActiveDatabases = (Get-MailboxDatabase -Server $ComputerNetBIOSName | Where-Object -FilterScript { $_.Server -match $ComputerNetBIOSName } | Select-Object -Property Name)
            
        }
        Else {
            
            $Database | ForEach-Object -Process {
                
                Try {
                    
                    $CurrentDatabase = (Get-MailboxDatabase -Identity $_ | Where-Object -FilterScript { $_.Server -match $ComputerNetBIOSName } | Select-Object -Property Name)
                    
                }
                Catch {
                    
                    [String]$MessageText = "Database {0} is not currently active on {1} and can't be checked" -f $CurrentDatabase.Name, $ComputerFQDNName
                    
                    $MessageText = Write-LogEntry -ToFile:$WriteToFile -LogPath $PerServerMessagesReportFile.FullName -MessageType WARNING -Message $MessageText -TimeStamp -ToScreen
                    
                    Write-Warning -Message $MessageText
                    
                    Continue
                    
                }
                
                Finally {
                    
                    $ActiveDatabases += $CurrentDatabase
                    
                }
                
            }
        }
        
        [Int]$ActiveDatabasesCount = (Measure-Object -InputObject $ActiveDatabases).Count
        
        If ($ActiveDatabasesCount -lt 1) {
            
            [String]$MessageText = "Any database was not found on the server {0}" -f $ComputerNetBIOSName
            
            $MessageText = Write-LogEntry -ToFile:$WriteToFile -LogPath $PerServerMessagesReportFile.FullName -MessageType WARNING -Message $MessageText -TimeStamp -ToScreen
            
            Write-Warning -Message $MessageText
            
        }
        
    }
    
    Process {
        
        $ActiveDatabases | ForEach-Object -Process {
            
            #Current time need to be compared between localhost and destination host to avoid mistakes
            $StartTimeForDatabase = $([DateTime]::Now)
            
            If ($CreateReportFile -eq 'CreatePerDatabase') {
                
                [Bool]$WriteToFile = $true
                
                if ($PSBoundParameters.ContainsKey('$ReportFileNamePrefix')) {
                    
                    [String]$ReportPerDatabaseNamePrefix = $ReportFileNamePrefix
                    
                }
                Else {
                    
                    [String]$ReportPerDatabaseNamePrefix = "{0}_{1}_IntegrityChecks" -f $_.Name, $_.Server
                    
                }
                
                
                $PerDatabaseMessagesReportFile = $(New-OutputFileNameFullPath -OutputFileDirectoryPath $ReportFileDirectoryPath -OutputFileNamePrefix $ReportPerDatabaseNamePrefix `
                                                                              -IncludeDateTimePartInOutputFileName $IncludeDateTimePartInReportFileName `
                                                                              -DateTimePartInOutputFileName $StartTimeForDatabase `
                                                                              -OutputFileNameSuffix 'messages' -BreakIfError $BreakOnReportCreationError).OutputFilePath
                
                Start-Log -LogPath $PerDatabaseEventsReportFile.DirectoryName -LogName $PerDatabaseEventsReportFile.Name -ScriptVersion $ScriptVersion.ToString()
                
                $PerDatabaseEventsReportFile = $(New-OutputFileNameFullPath -OutputFileDirectoryPath $ReportFileDirectoryPath -OutputFileNamePrefix $ReportPerDatabaseNamePrefix `
                                                                            -IncludeDateTimePartInOutputFileName $IncludeDateTimePartInReportFileName `
                                                                            -DateTimePartInOutputFileName $StartTimeForDatabase -OutputFileNameExtension "csv" `
                                                                            -OutputFileNameSuffix 'events' -BreakIfError $BreakOnReportCreationError).OutputFilePath
                
                $PerDatabaseCorruptionsDetailsReportFile = $(New-OutputFileNameFullPath -OutputFileDirectoryPath $ReportFileDirectoryPath -OutputFileNamePrefix $ReportPerDatabaseNamePrefix `
                                                                                        -IncludeDateTimePartInOutputFileName $IncludeDateTimePartInReportFileName `
                                                                                        -DateTimePartInOutputFileName $StartTimeForDatabase -OutputFileNameExtension "csv" `
                                                                                        -OutputFileNameSuffix 'corruption_details' -BreakIfError $BreakOnReportCreationError).OutputFilePath
                
            }
            
            
            #Check current status of database - if is still mounted on correct server - if not exit from current loop iteration
            
            Try {
                
                $CurrentDatabase = (Get-MailboxDatabase -Identity $_.Name | Where-Object -FilterScript { $_.Server -match $ComputerNetBIOSName } | Select-Object -Property Name)
                
            }
            Catch {
                
                [String]$MessageText = "Database {0} is not currently active on {1} and can't be checked" -f $CurrentDatabase, $ComputerFQDNName
                
                $MessageText = Write-LogEntry -ToFile:$WriteToFile -LogPath $PerServerMessagesReportFile.FullName -MessageType WARNING -Message $MessageText -TimeStamp -ToScreen
                
                Write-Warning -Message $MessageText
                
                #Exit from current loop iteration - check the next database
                Continue
                
            }
            
            [String]$MessageText = "Invoking command for repair database {0} on mailbox server {1}" -f $CurrentDatabase.Name, $ComputerFQDNName
            
            switch ($CreateReportFile) {
                
                'CreatePerServer' {
                    
                    $MessageText = Write-LogEntry -ToFile:$true -LogPath $PerServerMessagesReportFile.FullName -MessageType INFO -Message $MessageText -TimeStamp -ToScreen
                    
                }
                
                'CreatePerDatabase' {
                    
                    $MessageText = Write-LogEntry -ToFile:$true -LogPath $PerDatabaseMessagesReportFile.FullName -MessageType INFO -Message $MessageText -TimeStamp -ToScreen
                    
                }
                
                'None' {
                    
                    $MessageText = Write-LogEntry -MessageType INFO -Message $MessageText -TimeStamp -ToScreen
                    
                }
                
            }
            
            Write-Verbose -Message $MessageText
            
            Try {
                
                
                $RepairRequest = New-MailboxRepairRequest -Database $_.Name -CorruptionType "SearchFolder", "AggregateCounts", "ProvisionedFolder", "FolderView", "MessagePTagCn" -DetectOnly:$DetectOnly -ErrorAction Stop
                
            }
            Catch {
                
                [String]$MessageText = "Under invoking New-MailboxRepairRequest on {0} for the database {1} error occured: {2} " -f $ComputerFQDNName, $CurrentDatabase.Name, $Error[0]
                
                switch ($CreateReportFile) {
                    
                    'CreatePerServer' {
                        
                        $MessageText = Write-LogEntry -ToFile:$true -LogPath $PerServerMessagesReportFile.FullName -MessageType CRITICAL -Message $MessageText -TimeStamp -ToScreen
                        
                    }
                    
                    'CreatePerDatabase' {
                        
                        $MessageText = Write-LogEntry -ToFile:$true -LogPath $PerDatabaseMessagesReportFile.FullName -MessageType CRITICAL -Message $MessageText -TimeStamp -ToScreen
                        
                    }
                    
                    'None' {
                        
                        $MessageText = Write-LogEntry -MessageType INFO -Message $MessageText -TimeStamp -ToScreen
                        
                    }
                    
                    
                }
                
                Write-Error -Message $MessageText
                
                Continue
                
            }
            
            Start-Sleep -Seconds 1
            
            [Int]$ExpectedDurationStartWait = 5
            
            [int]$i = 1
            
            [Bool]$MonitoredEventsFound = $false
            
            while ($MonitoredEventsFound -eq $false) {
                
                If ($DisplayProgressBar) {
                    
                    [String]$MessageText = "Waiting for start repair operation on the database {0}." -f $CurrentDatabase.Name
                    
                    Write-Progress -Activity $MessageText -Status "Completion percentage is only confirmation that something is happening :-)" -PercentComplete (($i / ($ExpectedDurationStartWait * 60)) * 100)
                    
                    if (($i += $CheckProgressEverySeconds) -ge ($ExpectedDurationStartWait * 60)) {
                        
                        $i = $CheckProgressEverySeconds
                        
                    }
                    Else {
                        
                        $i += $CheckProgressEverySeconds
                        
                    }
                    
                }
                
                $MonitoredEvents = Get-EventsBySource -ComputerName $ComputerFQDNName -LogName "Application" -ProviderName "MSExchangeIS Mailbox Store" -EventID 10049, 10050, 10051, 10059 -StartTime $StartTimeForDatabase -Verbose:$false
                
                $MonitoredEventsFound = ((Measure-Object -InputObject $MonitoredEvents).count -ge 1)
                
                If ($MonitoredEventsFound) {
                    
                    $EventsToReport += $MonitoredEvents
                    
                    Try {
                        
                        #Filter for events which are errors
                        $ErrorEvents = ($MonitoredEvents | Where-Object -FilterScript { $_.EventId -ne 10059 })
                        
                    }
                    Catch { }
                    
                    $ErrorEventsFound = ((Measure-Object -InputObject $ErrorEvents).count -ge 1)
                    
                    Try {
                        
                        #Filter for event which confirm start of repair operations
                        $StartRepairEvent = ($MonitoredEvents | Where-Object -FilterScript { $_.EventId -eq 10059 })
                        
                    }
                    Catch { }
                    
                    $StartRepairEventFound = ((Measure-Object -InputObject $StartRepairEvent).count -eq 1)
                    
                }
                
                # Operations if errors found
                If ($ErrorEventsFound) {
                    
                    [String]$MessageText = "Under checking database the {0} on {1}  error occured - event ID  {2} " -f $_.Name, $ComputerFQDNName, $ErrorEvents.EventId
                    
                    
                    switch ($CreateReportFile) {
                        
                        'CreatePerServer' {
                            
                            $MessageText = Write-LogEntry -ToFile:$true -LogPath $PerServerMessagesReportFile.FullName -MessageType ERROR -Message $MessageText -TimeStamp -ToScreen
                            
                        }
                        
                        'CreatePerDatabase' {
                            
                            $MessageText = Write-LogEntry -ToFile:$true -LogPath $PerDatabaseMessagesReportFile.FullName -MessageType ERROR -Message $MessageText -TimeStamp -ToScreen
                            
                        }
                        
                        'None' {
                            
                            $MessageText = Write-LogEntry -MessageType INFO -Message $MessageText -TimeStamp -ToScreen
                            
                        }
                        
                    }
                    
                    Write-Error -Message $MessageText
                    
                    $StopTimeForDatabase = Get-Date
                    
                    $DurationTimeForDatabase = New-TimeSpan -Start $StartTimeForDatabase -End $StopTimeForDatabase
                    
                    [String]$MessageText = "Operation for the database {0} server {1} end with error at {2}, operation duration time: {3} days, {4} hours, {5} minutes, {6} seconds" `
                    -f $_.Name, $ComputerFQDNName, $StopTimeForDatabase, $DurationTimeForDatabase.Days, $DurationTimeForDatabase.Hours, $DurationTimeForDatabase.Minutes, $DurationTimeForDatabase.Seconds
                    
                    switch ($CreateReportFile) {
                        
                        'CreatePerServer' {
                            
                            $MessageText = Write-LogEntry -ToFile:$true -LogPath $PerServerMessagesReportFile.FullName -MessageType INFO -Message $MessageText -TimeStamp -ToScreen
                            
                        }
                        
                        'CreatePerDatabase' {
                            
                            $MessageText = Write-LogEntry -ToFile:$true -LogPath $PerDatabaseMessagesReportFile.FullName -MessageType INFO -Message $MessageText -TimeStamp -ToScreen
                            
                        }
                        
                        'None' {
                            
                            $MessageText = Write-LogEntry -MessageType INFO -Message $MessageText -TimeStamp -ToScreen
                            
                        }
                        
                    }
                    
                    Write-Verbose -Message $MessageText
                    
                    #Exit from current loop iteration - check the next database
                    Continue
                    
                    
                }
                
                elseif ($StartRepairEventFound) {
                    
                    [DateTime]$StartTimeRepair = $StartRepairEvent.TimeGenerated
                    
                    [String]$MessageText = "Repair request for the database {0} on the server {1} started at {2}" -f $_.Name, $ComputerFQDNName, $StartTimeRepair
                    
                    switch ($CreateReportFile) {
                        
                        'CreatePerServer' {
                            
                            $MessageText = Write-LogEntry -ToFile:$true -LogPath $PerServerMessagesReportFile.FullName -MessageType INFO -Message $MessageText -TimeStamp -ToScreen
                            
                        }
                        
                        'CreatePerDatabase' {
                            
                            $MessageText = Write-LogEntry -ToFile:$true -LogPath $PerDatabaseMessagesReportFile.FullName -MessageType INFO -Message $MessageText -TimeStamp -ToScreen
                            
                        }
                        
                        'None' {
                            
                            $MessageText = Write-LogEntry -MessageType INFO -Message $MessageText -TimeStamp -ToScreen
                            
                        }
                        
                    }
                    
                    Write-Verbose -Message $MessageText
                    
                }
                
                Else {
                    
                    Start-Sleep -Seconds $CheckProgressEverySeconds
                    
                    $MonitoredEventsFound = $false
                    
                }
                
            }
            
            
            Start-Sleep -Seconds 1
            
            [int]$i = 1
            
            $MonitoredEventsFound = $false
            
            $ErrorEventsFound = $false
            
            $StopRepairEventFound = $false
            
            #Loop responsible to check if repair operation finished
            while ($MonitoredEventsFound -eq $false) {
                
                $MonitoredEvents = Get-EventsBySource -ComputerName $ComputerFQDNName -LogName "Application" -ProviderName "MSExchangeIS Mailbox Store" -EventID 10045, 10048, 10049, 10050, 10051 -StartTime $StartTimeForDatabase -Verbose:$false
                
                $MonitoredEventsFound = ((Measure-Object -InputObject $MonitoredEvents).count -ge 1)
                
                If ($MonitoredEventsFound) {
                    
                    #Can be more than 1 object
                    $MonitoredEvents | ForEach-Object -Process {
                        
                        $EventsToReport += $_
                        
                    }
                    
                    Try {
                        
                        #Filter for events which are errors
                        $ErrorEvents = ($MonitoredEvents | Where-Object -FilterScript { $_.EventId -ne 10048 })
                        
                    }
                    Catch { }
                    
                    $ErrorEventsFound = ((Measure-Object -InputObject $ErrorEvents).count -ge 1)
                    
                    Try {
                        
                        #Filter for event which confirm start of repair operations
                        $StopRepairEvent = ($MonitoredEvents | Where-Object -FilterScript { $_.EventId -eq 10048 })
                        
                    }
                    Catch { }
                    
                    $StopRepairEventFound = ((Measure-Object -InputObject $StopRepairEvent).count -eq 1)
                    
                }
                
                If ($ErrorEventsFound) {
                    
                    [String]$MessageText = "Under checking the database {0} on the server {1} error occured - event ID  {2} " -f $_.Name, $ComputerFQDNName, $ErrorEvents.EventId
                    
                    switch ($CreateReportFile) {
                        
                        'CreatePerServer' {
                            
                            $MessageText = Write-LogEntry -ToFile:$true -LogPath $PerServerMessagesReportFile.FullName -MessageType ERROR -Message $MessageText -TimeStamp -ToScreen
                            
                        }
                        
                        'CreatePerDatabase' {
                            
                            $MessageText = Write-LogEntry -ToFile:$true -LogPath $PerDatabaseMessagesReportFile.FullName -MessageType ERROR -Message $MessageText -TimeStamp -ToScreen
                            
                        }
                        
                        'None' {
                            
                            $MessageText = Write-LogEntry -MessageType ERROR -Message $MessageText -TimeStamp -ToScreen
                            
                        }
                        
                    }
                    
                    Write-Error -Message $MessageText
                    
                    #Check if any 10062 errors occured before error                    
                    $CorruptionFoundEvents = Get-EventsBySource -ComputerName $ComputerFQDNName -LogName "Application" -ProviderName "MSExchangeIS Mailbox Store" -EventID 10062 -StartTime $StartTimeForDatabase -Verbose:$false
                    
                    $CorruptionFoundEventsCount = (Measure-Object -InputObject $CorruptionFoundEvents).count
                    
                    if ($CorruptionFoundEventsCount -ge 1) {
                        
                        $EventsToReport += $CorruptionFoundEvents
                        
                        $Events10062Details = Parse10062Events -Events $CorruptionFoundEvents
                        
                        $Events10062Details | ForEach-Object -Process {
                            
                            $Events10062DetailsToReport += $_
                            
                        }
                        
                    }
                    
                    $StopTimeForDatabase = Get-Date
                    
                    $DurationTimeForDatabase = New-TimeSpan -Start $StartTimeForDatabase -End $StopTimeForDatabase
                    
                    [String]$MessageText = "Operation for the database {0} server {1} end with error2 at {2}, operation duration time: {3} days, {4} hours, {5} minutes, {6} seconds" `
                    -f $_.Name, $ComputerFQDNName, $StopTimeForDatabase, $DurationTimeForDatabase.Days, $DurationTimeForDatabase.Hours, $DurationTimeForDatabase.Minutes, $DurationTimeForDatabase.Seconds
                    
                    switch ($CreateReportFile) {
                        
                        'CreatePerServer' {
                            
                            $MessageText = Write-LogEntry -ToFile:$true -LogPath $PerServerMessagesReportFile.FullName -MessageType INFO -Message $MessageText -TimeStamp -ToScreen
                            
                        }
                        
                        'CreatePerDatabase' {
                            
                            $MessageText = Write-LogEntry -ToFile:$true -LogPath $PerDatabaseMessagesReportFile.FullName -MessageType INFO -Message $MessageText -TimeStamp -ToScreen
                            
                        }
                        
                        'None' {
                            
                            $MessageText = Write-LogEntry -MessageType INFO -Message $MessageText -TimeStamp -ToScreen
                            
                        }
                        
                    }
                    
                    Write-Verbose -Message $MessageText
                    
                    #Exit from current loop iteration - check the next database
                    
                    Continue
                    
                }
                
                # Stop event found
                elseif ($StopRepairEventFound) {
                    
                    [String]$MessageText = "Repair request for the database {0} on the server {1} end successfully at {1}" -f $_.Name, $ComputerFQDNName, $StopRepairEvent.TimeGenerated
                    
                    switch ($CreateReportFile) {
                        
                        'CreatePerServer' {
                            
                            $MessageText = Write-LogEntry -ToFile:$true -LogPath $PerServerMessagesReportFile.FullName -MessageType INFO -Message $MessageText -TimeStamp -ToScreen
                            
                        }
                        
                        'CreatePerDatabase' {
                            
                            $MessageText = Write-LogEntry -ToFile:$true -LogPath $PerDatabaseMessagesReportFile.FullName -MessageType INFO -Message $MessageText -TimeStamp -ToScreen
                            
                        }
                        
                        'None' {
                            
                            $MessageText = Write-LogEntry -MessageType INFO -Message $MessageText -TimeStamp -ToScreen
                            
                        }
                        
                    }
                    
                    Write-Verbose -Message $MessageText
                    
                    #Check if any 10062 errors occured under check           
                    $CorruptionFoundEvents = Get-EventsBySource -ComputerName $ComputerFQDNName -LogName "Application" -ProviderName "MSExchangeIS Mailbox Store" -EventID 10062 -StartTime $StartTimeForDatabase -Verbose:$false
                    
                    
                    
                    $CorruptionFoundEventsCount = (Measure-Object -InputObject $CorruptionFoundEvents).count
                    
                    if ($CorruptionFoundEventsCount -ge 1) {
                        
                        $EventsToReport += $CorruptionFoundEvents
                        
                        $Events10062Details = Parse10062Events -Events $CorruptionFoundEvents
                        
                        $Events10062Details | ForEach-Object -Process {
                            
                            $Events10062DetailsToReport += $_
                            
                        }
                        
                    }
                    
                    $StopTimeForDatabase = Get-Date
                    
                    $DurationTimeForDatabase = New-TimeSpan -Start $StartTimeForDatabase -End $StopTimeForDatabase
                    
                    [String]$MessageText = "Operation for the database {0} server {1} end at {2}, operation duration time: {3} days, {4} hours, {5} minutes, {6} seconds" `
                    -f $_.Name, $ComputerFQDNName, $StopTimeForDatabase, $DurationTimeForDatabase.Days, $DurationTimeForDatabase.Hours, $DurationTimeForDatabase.Minutes, $DurationTimeForDatabase.Seconds
                    
                    switch ($CreateReportFile) {
                        
                        'CreatePerServer' {
                            
                            $MessageText = Write-LogEntry -ToFile:$true -LogPath $PerServerMessagesReportFile.FullName -MessageType INFO -Message $MessageText -TimeStamp -ToScreen
                            
                        }
                        
                        'CreatePerDatabase' {
                            
                            $MessageText = Write-LogEntry -ToFile:$true -LogPath $PerDatabaseMessagesReportFile.FullName -MessageType INFO -Message $MessageText -TimeStamp -ToScreen
                            
                            $EventsToReport | Export-Csv -Path $PerDatabaseEventsReportFile.FullName -Encoding UTF8 -NoTypeInformation -Delimiter ";"
                            
                            if ($CorruptionFoundEventsCount -ge 1) {
                                
                                $Events10062DetailsToReport | Export-Csv -Path $PerDatabaseCorruptionsDetailsReportFile.FullName -Encoding UTF8 -NoTypeInformation -Delimiter ";" -ErrorAction SilentlyContinue
                                
                            }
                            Else {
                                
                                Set-Content -Path $PerDatabaseCorruptionsDetailsReportFile.FullName -value "No corruption events found"
                                
                            }
                            
                            Clear-Variable -Name MessagesToReport -ErrorAction 'SilentlyContinue'
                            
                            Clear-Variable -Name EventsToReport -ErrorAction 'SilentlyContinue'
                            
                            Clear-Variable -Name Event10062DetailsToReport -ErrorAction 'SilentlyContinue'
                            
                        }
                        
                        'None' {
                            
                            $MessageText = Write-LogEntry -MessageType INFO -Message $MessageText -TimeStamp -ToScreen
                            
                        }
                        
                    }
                    
                    Write-Verbose -Message $MessageText
                    
                    <#
                    
                    If ($DisplaySummary) {
                        
                        Write-Output -InputObject $CorruptionFoundEvents
                        
                    }
                    
                    #>
                    
                }
                
                else {
                    
                    If ($DisplayProgressBar) {
                        
                        [String]$MessageText = "The database {0} repair on the server {1} request  is in progress." -f $_.Name, $ComputerFQDNName
                        
                        Write-Progress -Activity $MessageText -Status "Completion percentage is only confirmation that something is happening :-)" -PercentComplete (($i / ($ExpectedDurationTimeMinutes * 60)) * 100)
                        
                        if (($i += $CheckProgressEverySeconds) -ge ($ExpectedDurationTimeMinutes * 60)) {
                            
                            $i = $CheckProgressEverySeconds
                        }
                        
                        Else {
                            
                            $i += $CheckProgressEverySeconds
                            
                        }
                        
                    }
                    
                    Start-Sleep -Seconds $CheckProgressEverySeconds
                    
                    $MonitoredEventsFound = $false
                    
                }
                
            }
            
            
        }
        
        
    }
    
    End {
        
        $StopTimeForServer = Get-Date
        
        $DurationTimeForServer = New-TimeSpan -Start $StartTimeForServer -End $StopTimeForServer
        
        [String]$MessageText = "Operation for the server {0} ended at {1}, operation duration time: {2} days, {3} hours, {4} minutes, {5} seconds" `
        -f $ComputerFQDNName, $StopTimeForServer, $DurationTimeForServer.Days, $DurationTimeForServer.Hours, $DurationTimeForServer.Minutes, $DurationTimeForServer.Seconds
        
        $MessageText = Write-LogEntry -ToFile:$WriteToFile -LogPath $PerServerMessagesReportFile.FullName -MessageType INFO -Message $MessageText -TimeStamp -ToScreen
        
        Write-Verbose -Message $MessageText
        
        #region Write reports per server
        
        If ($CreateReportFile -eq "CreatePerServer") {
            
            $EventsToReport | Export-Csv -Path $PerServerEventsReportFile.FullName -Encoding UTF8 -NoTypeInformation -Delimiter ";"
            
            if ($CorruptionFoundEventsCount -ge 1) {
                
                $Events10062DetailsToReport | Export-Csv -Path $PerServersCorruptionDetailsReportFile.FullName -Encoding UTF8 -NoTypeInformation -Delimiter ";" -ErrorAction SilentlyContinue
                
            }
            Else {
                
                Set-Content -Path $PerServersCorruptionDetailsReportFile.FullName -value "No corruption events found"
                
            }
            
            
            
        }
        
        #endregion
        
    }
    
}

Function New-OutputFileNameFullPath {
<#

    .SYNOPSIS
    Function intended for preparing filename for output files like reports or logs
   
    .DESCRIPTION
    Function intended for preparing filename for output files like reports or logs based on prefix, middle name part, suffix, date, etc. with verification if provided path is writable
    
    Returned object contains properties
    - OutputFilePath - to use it please check an examples - as a [System.IO.FileInfo]
    - ExitCode
    - ExitCodeDescription
    
    Exit codes and descriptions
    0 = "Everything is fine :-)"
    1 = "Provided path <PATH> doesn't exist and can't be created
    2 = "Provided patch <PATH> doesn't exist and value for the parameter CreateOutputFileDirectory is set to False"
    3 = "Provided patch <PATH> is not writable"
    4 = "The file <PATH>\<FILE_NAME> already exist"
    
    .PARAMETER OutputFileDirectoryPath
    By default output files are stored in subfolder "outputs" in current path
    
    .PARAMETER CreateOutputFileDirectory
    Set tu TRUE if provided output file directory should be created if is missed
    
    .PARAMETER OutputFileNamePrefix
    Prefix used for creating output files name
    
    .PARAMETER OutputFileNameMidPart
    Part of the name which will be used in midle of output file name
    
    .PARAMETER OutputFileNameSuffixPart
    Part of the name which will be used at the end of output file name
    
    .PARAMETER IncludeDateTimePartInOutputFileName
    Set to TRUE if report file name should contains part based on date and time - format yyyyMMdd-HHmm is used
    
    .PARAMETER DateTimePartInOutputFileName
    Set to date and time which should be used in output file name, by default current date and time is used
    
    .PARAMETER OutputFileNameExtension
    Set to extension which need to be used for output file, by default ".txt" is used
    
    .PARAMETER ErrorIfOutputFileExist
    Generate error if output file already exist
    
    .PARAMETER BreakIfError
    Break function execution if parameters provided for output file creation are not correct or destination file path is not writables
    
    .EXAMPLE
    
    PS \> $PerServerMessagesReportFile = New-OutputFileNameFullPath -OutputFileDirectoryPath 'C:\USERS\Wojtek\' -OutputFileNamePrefix 'Messages' `
                                                                    -OutputFileNameMidPart 'COMPUTERNAME' `
                                                                    -IncludeDateTimePartInOutputFileName:$true `
                                                                    -BreakIfError:$true
    
    PS \> $PerServerMessagesReportFile | Format-List
    
    OutputFilePath                                           ExitCode ExitCodeDescription
    --------------                                           -------- -------------------
    C:\users\wojtek\Messages-COMPUTERNAME-20151021-0012-.txt        0 Everything is fine :-)
    
    .EXAMPLE
    
    PS \> $PerServerMessagesReportFile = New-OutputFileNameFullPath -OutputFileDirectoryPath 'C:\USERS\Wojtek\' -OutputFileNamePrefix 'Messages' `
                                                                    -OutputFileNameMidPart 'COMPUTERNAME' -IncludeDateTimePartInOutputFileName:$true 
                                                                    -OutputFileNameExtension rxc -OutputFileNameSuffix suffix `
                                                                    -BreakIfError:$true
    
    
    PS \> $PerServerMessagesReportFile.OutputFilePath | select name,extension,Directory | Format-List

    Name      : Messages-COMPUTERNAME-20151022-235607-suffix.rxc
    Extension : .rxc
    Directory : C:\USERS\Wojtek
    
    PS \> ($PerServerMessagesReportFile.OutputFilePath).gettype()

    IsPublic IsSerial Name                                     BaseType
    -------- -------- ----                                     --------
    True     True     FileInfo                                 System.IO.FileSystemInfo
     
    .LINK
    https://github.com/it-praktyk/New-OutputFileNameFullPath
    
    .LINK
    https://www.linkedin.com/in/sciesinskiwojciech
          
    .NOTES
    AUTHOR: Wojciech Sciesinski, wojciech[at]sciesinski[dot]net
    KEYWORDS: PowerShell
   
    VERSIONS HISTORY
    0.1.0 - 2015-09-01 - Initial release
    0.1.1 - 2015-09-01 - Minor update
    0.2.0 - 2015-09-08 - Corrected, function renamed to New-OutputFileNameFullPath from New-ReportFileNameFullPath
    0.3.0 - 2015-09-13 - implementation for DateTimePartInFileName parameter corrected, help updated, some parameters renamed
    0.4.0 - 2015-10-20 - additional OutputFileNameSuffix parameter added, help updated, TODO updated
    0.4.1 - 2015-10-21 - help corrected
    0.5.0 - 2015-10-22 - Returned OutputFilePath changed to type [System.IO.FileInfo], help updated
    
    TODO
    Change/extend behavior if file exist ?
    Trim provided parameters, replace not standard chars ? 

        
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
    param (
        [parameter(Mandatory = $false)]
        [String]$OutputFileDirectoryPath = ".\Outputs\",
        [parameter(Mandatory = $false)]
        [Bool]$CreateOutputFileDirectory = $true,
        [parameter(Mandatory = $false)]
        [String]$OutputFileNamePrefix = "Output-",
        [parameter(Mandatory = $false)]
        [String]$OutputFileNameMidPart = $null,
        [parameter(Mandatory = $false)]
        [String]$OutputFileNameSuffix = $null,
        [parameter(Mandatory = $false)]
        [Bool]$IncludeDateTimePartInOutputFileName = $true,
        [parameter(Mandatory = $false)]
        [Nullable[DateTime]]
        $DateTimePartInOutputFileName = $null,
        [parameter(Mandatory = $false)]
        [String]$OutputFileNameExtension = ".txt",
        [parameter(Mandatory = $false)]
        [Bool]$ErrorIfOutputFileExist = $true,
        [parameter(Mandatory = $false)]
        [Bool]$BreakIfError = $true
        
    )
    
    #Declare variable
    
    [Int]$ExitCode = 0
    
    [String]$ExitCodeDescription = "Everything is fine :-)"
    
    $Result = New-Object -TypeName PSObject
    
    #Convert relative path to absolute path
    [String]$OutputFileDirectoryPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputFileDirectoryPath)
    
    #Assign value to the variable $IncludeDateTimePartInOutputFileName if is not initialized
    If ($IncludeDateTimePartInOutputFileName -and $DateTimePartInOutputFileName -eq $null) {
        
        [String]$DateTimePartInFileNameString = $(Get-Date -format yyyyMMdd-HHmmss)
        
    }
    Else {
        
        [String]$DateTimePartInFileNameString = $(Get-Date -Date $DateTimePartInOutputFileName -format yyyyMMdd-HHmmss)
        
    }
    
    #Check if Output directory exist and try create if not
    If ($CreateOutputFileDirectory -and !$((Get-Item -Path $OutputFileDirectoryPath -ErrorAction SilentlyContinue) -is [system.io.directoryinfo])) {
        
        Try {
            
            $ErrorActionPreference = 'Stop'
            
            New-Item -Path $OutputFileDirectoryPath -type Directory | Out-Null
            
        }
        Catch {
            
            [String]$MessageText = "Provided path {0} doesn't exist and can't be created" -f $OutputFileDirectoryPath
            
            If ($BreakIfError) {
                
                Throw $MessageText
                
            }
            Else {
                
                Write-Error -Message $MessageText
                
                [Int]$ExitCode = 1
                
                [String]$ExitCodeDescription = $MessageText
                
            }
            
        }
        
    }
    ElseIf (!$((Get-Item -Path $OutputFileDirectoryPath -ErrorAction SilentlyContinue) -is [system.io.directoryinfo])) {
        
        [String]$MessageText = "Provided patch {0} doesn't exist and value for the parameter CreateOutputFileDirectory is set to False" -f $OutputFileDirectoryPath
        
        If ($BreakIfError) {
            
            Throw $MessageText
            
        }
        Else {
            
            Write-Error -Message $MessageText
            
            [Int]$ExitCode = 2
            
            [String]$ExitCodeDescription = $MessageText
            
        }
        
    }
    
    #Try if Output directory is writable - temporary file is stored
    Try {
        
        $ErrorActionPreference = 'Stop'
        
        [String]$TempFileName = [System.IO.Path]::GetTempFileName() -replace '.*\\', ''
        
        [String]$TempFilePath = "{0}{1}" -f $OutputFileDirectoryPath, $TempFileName
        
        New-Item -Path $TempFilePath -type File | Out-Null
        
    }
    Catch {
        
        [String]$MessageText = "Provided patch {0} is not writable" -f $OutputFileDirectoryPath
        
        If ($BreakIfError) {
            
            Throw $MessageText
            
        }
        Else {
            
            Write-Error -Message $MessageText
            
            [Int]$ExitCode = 3
            
            [String]$ExitCodeDescription = $MessageText
            
        }
        
    }
    
    Remove-Item -Path $TempFilePath -ErrorAction SilentlyContinue | Out-Null
    
    #Constructing the file name
    If (!($IncludeDateTimePartInOutputFileName) -and !([String]::IsNullOrEmpty($OutputFileNameMidPart))) {
        
        [String]$OutputFilePathTemp1 = "{0}\{1}-{2}" -f $OutputFileDirectoryPath, $OutputFileNamePrefix, $OutputFileNameMidPart
        
    }
    Elseif (!($IncludeDateTimePartInOutputFileName) -and [String]::IsNullOrEmpty($OutputFileNameMidPart)) {
        
        [String]$OutputFilePathTemp1 = "{0}\{1}" -f $OutputFileDirectoryPath, $OutputFileNamePrefix
        
    }
    ElseIf ($IncludeDateTimePartInOutputFileName -and !([String]::IsNullOrEmpty($OutputFileNameMidPart))) {
        
        [String]$OutputFilePathTemp1 = "{0}\{1}-{2}-{3}" -f $OutputFileDirectoryPath, $OutputFileNamePrefix, $OutputFileNameMidPart, $DateTimePartInFileNameString
        
    }
    Else {
        
        [String]$OutputFilePathTemp1 = "{0}\{1}-{2}" -f $OutputFileDirectoryPath, $OutputFileNamePrefix, $DateTimePartInFileNameString
        
    }
    
    If ([String]::IsNullOrEmpty($OutputFileNameSuffix)) {
        
        [String]$OutputFilePathTemp = "{0}.{1}" -f $OutputFilePathTemp1, $OutputFileNameExtension
        
    }
    Else {
        
        [String]$OutputFilePathTemp = "{0}-{1}.{2}" -f $OutputFilePathTemp1, $OutputFileNameSuffix, $OutputFileNameExtension
        
    }
    
    #Replacing doubled chars \\ , -- , .. - except if \\ is on begining - means that path is UNC share
    [System.IO.FileInfo]$OutputFilePath = "{0}{1}" -f $OutputFilePathTemp.substring(0, 2), (($OutputFilePathTemp.substring(2, $OutputFilePathTemp.length - 2).replace("\\", '\')).replace("--", "-")).replace("..", ".")
    
    If ($ErrorIfOutputFileExist -and (Test-Path -Path $OutputFilePath -PathType Leaf)) {
        
        [String]$MessageText = "The file {0} already exist" -f $OutputFilePath
        
        If ($BreakIfError) {
            
            Throw $MessageText
            
        }
        Else {
            
            Write-Error -Message $MessageText
            
            [Int]$ExitCode = 4
            
            [String]$ExitCodeDescription = $MessageText
            
        }
    }
    
    $Result | Add-Member -MemberType NoteProperty -Name OutputFilePath -Value $OutputFilePath
    
    $Result | Add-Member -MemberType NoteProperty -Name ExitCode -Value $ExitCode
    
    $Result | Add-Member -MemberType NoteProperty -Name ExitCodeDescription -Value $ExitCodeDescription
    
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
    
    $CorruptionFoundEvents | ForEach-Object -Process {
        
        If ($_.EventID -eq 10062) {
            
            $separator = '^'
            
            $MessageLines = ($_.Message).Split($separator, $option)
            
            Write-Verbose -Message "Lines to parse ($MessageLines | Measure).count"
            
            [Int]$i = 1
            
            $MessageLines | ForEach-Object -Process {
                
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
                    
                    $Fields | ForEach-Object -Process {
                        
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
    Function intended to gather data from Windows events logs - the function Get-EventsBySource is wrapper for Get-WinEvent function
    
    .DESCRIPTION
    Function intended to gather data from Windows events logs - the function Get-EventsBySource is wrapper for Get-WinEvent function. Generally
    the HashQuerySet parameter set is used but time span can be constructed based not only on start/end time but also 
    
    Function offer additional capabilities to merge multilines event description (using defined char as a lines separator)
    and can limit amount of returned 
  
    .PARAMETER ComputerName
    Gets events from the event logs on the specified computer. Type the NetBIOS name, an Internet Protocol (IP) address,
    or the fully qualified domain name of the computer. The default value is the local computer.
       
    .PARAMETER LogName
    Gets events from the specified event logs. Enter the event log names in a comma-separated list.
    
    .PARAMETER ProviderName
    Gets events written by the specified event log providers. Enter the provider names in a comma-separated list, or use wildcard characters to create provider name patterns.
    An event log provider is a program or service that writes events to the event log. It is not a Windows PowerShell provider.
    Please remember that ProviderName is usually not equal with a source for event - please check an event XML to check used provider.
    
    .PARAMETER EventID
    
    
    .PARAMETER StartTime
    Date and time which will be used as the begining of a time period to query
    
    .PARAMETER EndTime
    Date and time which will be used as the end of a time period to query
    
    .PARAMETER ForLastTimeSpan
    Use number for which logs need to be queried - please select also correct "ForLastTimeUnit" value
    
    .PARAMETER ForLastTimeUnit
    Use the name of units for construct query.
    
    .PARAMETER ConcatenateMessageLines
    For multilines events description lines will be merged by default. Please change to $false if you would not like this behaviour, than only first line can be handled.
    
    .PARAMETER ConcatenatedLinesSeparator
    A char used to separated merged multilines event description. By default "^" is used due that is not usually used in events descriptions.
    
    .PARAMETER MessageCharsAmount
    The number of chars which will be returned from event description.
    
    .INPUT
    Cos
    
    .OUTPUT
    Costam
     
    .EXAMPLE
    Get-EventsBySource -ComputerName localhost -LogName application -ProviderName SecurityCenter -EventID 1,16 -ForLastTimeSpan 160 -ForLastTimeUnit minutes

    ComputerName  : COMPUTERNAME.wojteks.lab
    Source        : SecurityCenter
    EventID       : 16
    TimeGenerated : 10/19/2015 10:48:26 PM
    Message       : The Windows Security Center Service could not stop Windows Defender

    ComputerName  : COMPUTERNAME.wojteks.lab
    Source        : SecurityCenter
    EventID       : 1
    TimeGenerated : 10/19/2015 10:48:23 PM
    Message       : The Windows Security Center Service has started
    
    .LINK
    https://github.com/it-praktyk/Get-EvenstBySource
    
    .LINK
    https://www.linkedin.com/in/sciesinskiwojciech
          
    .NOTES
   
    AUTHOR: Wojciech Sciesinski, wojciech[at]sciesinski[dot]net
            Parameters description partially based on Get-WinEvent help from PowerShell 3.0
    
    KEYWORDS: Windows, Event logs, PowerShell
    
    VERSION HISTORY
    0.3.1 - 2015-07-03 - Support for time span corrected, the first version published on GitHub
    0.3.2 - 2015-07-05 - Help updated, function corrected
    0.3.3 - 2015-08-25 - Help updated, to do updated
    0.4.0 - 2015-09-08 - Code reformated, Added support for more than one event id, minor update
    0.5.0 - 2015-10-19 - Code corrected based on PSScriptAnalyzer 1.1.0 output, support for more than logs (by name) added,
                         help partially updated
                         

    TODO
    - help update needed
    - handle situation like
    
    PS > [Array]$FilterHashTable = @{ "Logname" = "Application"; "Id" = 900; "ProviderName" = "Microsoft-Windows-Security-SPP" }
    PS > Get-WinEvent -FilterHashtable $FilterHashTable
    
    PS > Get-WinEvent -FilterHashtable $FilterHashTable
    Get-WinEvent : The specified providers do not write events to any of the specified logs.
    At line:1 char:1
    + Get-WinEvent -FilterHashtable $FilterHashTable
    + ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    + CategoryInfo          : InvalidArgument: (:) [Get-WinEvent], Exception
    + FullyQualifiedErrorId : LogsAndProvidersDontOverlap,Microsoft.PowerShell.Commands.GetWinEventCommand
    
        
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
    [OutputType("System.Object[]")]
    param (
        [parameter(mandatory = $true)]
        [String]$ComputerName,
        [parameter(mandatory = $true)]
        [String[]]$LogName,
        [parameter(mandatory = $true)]
        [String]$ProviderName,
        [parameter(mandatory = $true)]
        [alias("ID")]
        [Int[]]$EventID,
        [parameter(mandatory = $false, ParameterSetName = "StartEndTime")]
        [Nullable[DateTime]]
        $StartTime = $null,
        [parameter(mandatory = $false, ParameterSetName = "StartEndTime")]
        [Nullable[DateTime]]
        $EndTime = $null,
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
            
            Continue
            
        }
        
        Finally {
            
            $Found = $($Events | Measure-Object).Count
            
            If ($Found -ne 0) {
                
                [String]$MessageText = "For the computer $ComputerName events $Found found"
                
                Write-Verbose -Message $MessageText
                
                $Events | ForEach-Object -Process {
                    
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
    
    
    END {
        
        Return $Results
        
    }
    
}

###
# Author: Luca Sturlese
# URL: http://9to5IT.com
# Author: Wojciech Sciesinski
# URL: https://www.linkedin.com/in/sciesinskiwojciech
###

Set-StrictMode -Version Latest


Function Start-Log {
  <#
  .SYNOPSIS
    Creates a new log file

  .DESCRIPTION
    Creates a log file with the path and name specified in the parameters. Checks if log file exists, and if it does deletes it and creates a new one.
    Once created, writes initial logging data

  .PARAMETER LogPath
    Mandatory. Path of where log is to be created. Example: C:\Windows\Temp

  .PARAMETER LogName
    Mandatory. Name of log file to be created. Example: Test_Script.log

  .PARAMETER ScriptVersion
    Mandatory. Version of the running script which will be written in the log. Example: 1.5

  .PARAMETER ToScreen
    Optional. When parameter specified will display the content to screen as well as write to log file. This provides an additional
    another option to write content to screen as opposed to using debug mode.

  .INPUTS
    Parameters above

  .OUTPUTS
    Log file created

  .NOTES
    Version:        1.0
    Author:         Luca Sturlese
    Creation Date:  10/05/12
    Purpose/Change: Initial function development.

    Version:        1.1
    Author:         Luca Sturlese
    Creation Date:  19/05/12
    Purpose/Change: Added debug mode support.

    Version:        1.2
    Author:         Luca Sturlese
    Creation Date:  02/09/15
    Purpose/Change: Changed function name to use approved PowerShell Verbs. Improved help documentation.

    Version:        1.3
    Author:         Luca Sturlese
    Creation Date:  07/09/15
    Purpose/Change: Resolved issue with New-Item cmdlet. No longer creates error. Tested - all ok.

    Version:        1.4
    Author:         Luca Sturlese
    Creation Date:  12/09/15
    Purpose/Change: Added -ToScreen parameter which will display content to screen as well as write to the log file.
    
    TODO
    Information about time zone need to be added to header
    Script version should be changed to optional

  .LINK
    http://9to5IT.com/powershell-logging-v2-easily-create-log-files

  .EXAMPLE
    Start-Log -LogPath "C:\Windows\Temp" -LogName "Test_Script.log" -ScriptVersion "1.5"

    Creates a new log file with the file path of C:\Windows\Temp\Test_Script.log. Initialises the log file with
    the date and time the log was created (or the calling script started executing) and the calling script's version.
  #>
    
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $true, Position = 0)]
        [string]$LogPath,
        [Parameter(Mandatory = $true, Position = 1)]
        [string]$LogName,
        [Parameter(Mandatory = $true, Position = 2)]
        [string]$ScriptVersion,
        [Parameter(Mandatory = $false, Position = 3)]
        [switch]$ToScreen
    )
    
    Process {
        $sFullPath = Join-Path -Path $LogPath -ChildPath $LogName
        
        #Check if file exists and delete if it does
        If ((Test-Path -Path $sFullPath)) {
            Remove-Item -Path $sFullPath -Force
        }
        
        #Create file and start logging
        New-Item -Path $sFullPath –ItemType File
        
        Add-Content -Path $sFullPath -Value "***************************************************************************************************"
        Add-Content -Path $sFullPath -Value "Started processing at [$([DateTime]::Now)]."
        Add-Content -Path $sFullPath -Value "***************************************************************************************************"
        Add-Content -Path $sFullPath -Value ""
        Add-Content -Path $sFullPath -Value "Running script version [$ScriptVersion]."
        Add-Content -Path $sFullPath -Value ""
        Add-Content -Path $sFullPath -Value "***************************************************************************************************"
        Add-Content -Path $sFullPath -Value ""
        
        #Write to screen for debug mode
        Write-Debug "***************************************************************************************************"
        Write-Debug "Started processing at [$([DateTime]::Now)]."
        Write-Debug "***************************************************************************************************"
        Write-Debug ""
        Write-Debug "Running script version [$ScriptVersion]."
        Write-Debug ""
        Write-Debug "***************************************************************************************************"
        Write-Debug ""
        
        #Write to scren for ToScreen mode
        If ($ToScreen -eq $True) {
            Write-Output "***************************************************************************************************"
            Write-Output "Started processing at [$([DateTime]::Now)]."
            Write-Output "***************************************************************************************************"
            Write-Output ""
            Write-Output "Running script version [$ScriptVersion]."
            Write-Output ""
            Write-Output "***************************************************************************************************"
            Write-Output ""
        }
    }
}

Function Write-LogEntry {
<#
    .SYNOPSIS
    Writes a message to specified log file

    .DESCRIPTION
    Appends a new message to the specified log file

    .PARAMETER LogPath
    Mandatory. Full path of the log file you want to write to. Example: C:\Windows\Temp\Test_Script.log
    
    .PARAMETER MessageType
    Mandatory. Allowed message types: INFO, WARNING, ERROR, CRITICAL, START, STOP, SUCCESS, FAILURE

    .PARAMETER Message
    Mandatory. The string that you want to write to the log

    .PARAMETER TimeStamp
    Optional. When parameter specified will append the current date and time to the end of the line. Useful for knowing
    when a task started and stopped.
    
    .PARAMETER EntryDateTime
    Optional. By default a current date and time is used but it is possible provide any other correct date/time.
    
    .PARAMETER ConvertTimeToUTF
    # Need be filled

    .PARAMETER ToScreen
    Optional. When parameter specified will display the content to screen as well as write to log file. This provides an additional
    another option to write content to screen as opposed to using debug mode.
    
    .PARAMETER

    .INPUTS
    Parameters above

    .OUTPUTS
    None or String

    .NOTES
    
    Version:        1.0.0
    Author:         Wojciech Sciesinski, wojciech[at]sciesinski[dot]net
    Purpose/Change: Initial function development.
    Creation Date:  25/10/2015
    
    Version         1.1.0
    Author:         Wojciech Sciesinski, wojciech[at]sciesinski[dot]net
    Purpose/Change: A date for a message can be declared as parameter, version number corrected 2.0.0 > 1.0.0 corrected
    Creation Date:  26/10/2015
    
    Inspired and partially based on PSLogging module authored by Luca Sturlese - https://github.com/9to5IT/PSLogging
    
    TODO
    Updated examples - add additional with new implemented parameters
    Implement converting day/time to UTF
    Output with colors (?) - Write-Host except Write-Output need to be used
    
    .LINK
    http://9to5IT.com/powershell-logging-v2-easily-create-log-files

    .LINK
    https://github.com/it-praktyk/PSLogging
    
    .LINK
    https://www.linkedin.com/in/sciesinskiwojciech
    

  .EXAMPLE
    Write-LogEntry -LogPath "C:\Windows\Temp\Test_Script.log" -MessageType CRITICAL -Message "This is a new line which I am appending to the end of the log file."

    Writes a new critical log message to a new line in the specified log file.
    
  #>
    
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $false, ParameterSetName = "WriteToFile")]
        [Switch]$ToFile,
        [Parameter(Mandatory = $false, ParameterSetName = "WriteToFile")]
        [string]$LogPath,
        [Parameter(Mandatory = $true, HelpMessage = "Allowed values: INFO, WARNING, ERROR, CRITICAL, START, STOP, SUCCESS, FAILURE")]
        [ValidateSet("INFO", "WARNING", "ERROR", "CRITICAL", "START", "STOP", "SUCCESS", "FAILURE")]
        [String]$MessageType,
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [Alias("EventMessage", "EntryMessage")]
        [string]$Message,
        [Parameter(Mandatory = $false)]
        [switch]$TimeStamp,
        [Parameter(Mandatory = $false)]
        [Alias("EventDateTime", "EntryDate", "MessageDate")]
        [DateTime]$EntryDateTime = $([DateTime]::Now),
        [Parameter(Mandatory = $false)]
        [switch]$ToScreen
        
    )
    
    
    
    Process {
        
        #Capitalize MessageType value
        [String]$CapitalizedMessageType = $MessageType.ToUpper()
        
        #A padding used to allign columns in output file
        [String]$Padding = " " * $(10 - $CapitalizedMessageType.Length)
        
        #Add TimeStamp to message if specified
        If ($TimeStamp -eq $True) {
            
            [String]$MessageToFile = "[{0}][{1}{2}][{3}]" -f $EntryDateTime, $CapitalizedMessageType, $Padding, $Message
            
            [String]$MessageToScreen = "[{0}] {1}: {2}" -f $EntryDateTime, $CapitalizedMessageType, $Message
            
        }
        Else {
            
            [String]$MessageToFile = "[{0}{1}][{2}]" -f $type, $Padding, $Message
            
            [String]$MessageToScreen = "{0}: {1}" -f $type, $Message
        }
        
        #Write Content to Log
        
        If ($ToFile -eq $true) {
            
            Add-Content -Path $LogPath -Value $MessageToFile
            
        }
        
        #Write to screen for debug mode
        Write-Debug $MessageToScreen
        
        #Write to scren for ToScreen mode
        If ($ToScreen -eq $True) {
            
            Write-Output $MessageToScreen
            
        }
        
    }
}