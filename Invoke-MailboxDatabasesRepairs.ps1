function Invoke-MailboxDatabasesRepairs {
<#
    .SYNOPSIS
    Function intended for performing checks and repairs operation on Exchange Server 2010 SP1 (or newer) mailbox databases
   
    .DESCRIPTION
    Function invokes New-MailboxDatabaseRepair cmdlet for all active mailbox database copies on server. Mailbox databases can be also
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
    
    [PS] >Invoke-MailboxDatabasesRepairs -ComputerName XXXXXXMBX03 -Database All -DisplaySummary:$true -ExpectedDurationTimeMinutes 120 -DetectOnly:$true
     
    .LINK
    https://github.com/it-praktyk/Invoke-MailboxDatabasesRepairs
    
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
    0.3.0 - 2015-09-04 - Added support for Exchange 2013, added support for database repair errors
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
        

    DEPENDENCIES
    -   Function Test-ExchangeCmdletsAvailability - minimum 0.1.2
        https://github.com/it-praktyk/Test-ExchangeCmdletsAvailability
    -   Function Function Get-EventsBySource - minimum 0.5.0
        https://github.com/it-praktyk/Get-EvenstBySource
    -   Function New-OutputFileNameFullPath - minimum 0.4.0
        https://github.com/it-praktyk/New-OutputFileNameFullPath

    TO DO
    - improve store and/or mail summary report - export as csv/xml/JSON
    - Current time and timezone need to be compared between localhost and destination host to avoid mistakes
    - exit code return need to be implemented
    - add support for Exchange 2013 (?) and 2016 (?)
    - add named regions to easier navigation in code 
    - normalize displayed/logged messages
        
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
        [switch]$DetectOnly = $false,
        [parameter(Mandatory = $false)]
        [switch]$DisplayProgressBar = $true,
        [parameter(Mandatory = $false)]
        [Int]$CheckProgressEverySeconds = 120,
        [parameter(Mandatory = $false)]
        [switch]$DisplaySummary = $false,
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
        [Switch]$IncludeDateTimePartInReportFileName = $true,
        [parameter(Mandatory = $false, ParameterSetName = "Reports")]
        [DateTime]$DateTimePartInReportFileName,
        [parameter(Mandatory = $false, ParameterSetName = "Reports")]
        [String]$ReportFileNameExtension = ".txt",
        [parameter(Mandatory = $false, ParameterSetName = "Reports")]
        [Switch]$BreakOnReportCreationError = $true
        
    )
    
    #endregion
    
    Begin {
        
        #region Initialize variables
        
        $ActiveDatabases = @()
        
        $MessagesToReport = @()
        
        $EventsToReport = @()
        
        $Events10062DetailsToReport = @()
        
        [Bool]$IsRunningOnLocalhost = $false
        
        [Bool]$StartRepairEventFound = $false
        
        [Bool]$StopRepairEventFound = $false
        
        $StartTimeForServer = Get-Date
        
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
        
        $MessagesToReport += "`n$MessageText"
        
        Write-Verbose -Message $MessageText
        
        #endregion
        
        #region Initialize reports files names
        
        #Creating name for the report, a report file will be used for save initial errors or all messages if CreatePerServer report will be selected
        
        if ($CreateReportFile -ne 'None') {
            
            if ($PSBoundParameters.ContainsKey('$ReportFileNamePrefix')) {
                
                [String]$ReportPerServerNamePrefix = $ReportFileNamePrefix
                
            }
            Else {
                
                [String]$ReportPerServerNamePrefix = "MBDBs_IntegrityChecks_{0}" -f $ComputerNetBIOSName
                
            }
            
            $PerServerReportFileMessages = New-OutputFileNameFullPath -OutputFileDirectoryPath $ReportFileDirectoryPath -OutputFileNamePrefix $ReportPerServerNamePrefix `
                                                                      -OutputFileNameMidPart $ReportFileNameMidPart `
                                                                      -IncludeDateTimePartInOutputFileName:$IncludeDateTimePartInReportFileName `
                                                                      -DateTimePartInOutputFileName $StartTimeForServer `
                                                                      -OutputFileNameSuffix 'messages' -BreakIfError:$BreakOnReportCreationError
            
            $PerServerReportFileEvents = New-OutputFileNameFullPath -OutputFileDirectoryPath $ReportFileDirectoryPath -OutputFileNamePrefix $ReportPerServerNamePrefix `
                                                                    -OutputFileNameMidPart $ReportFileNameMidPart `
                                                                    -IncludeDateTimePartInOutputFileName:$IncludeDateTimePartInReportFileName `
                                                                    -DateTimePartInOutputFileName $StartTimeForServer `
                                                                    -OutputFileNameSuffix 'events' -BreakIfError:$BreakOnReportCreationError
            
            $PerServerReportFileCorruptionsDetails = New-OutputFileNameFullPath -OutputFileDirectoryPath $ReportFileDirectoryPath -OutputFileNamePrefix $ReportPerServerNamePrefix `
                                                                                -OutputFileNameMidPart $ReportFileNameMidPart `
                                                                                -IncludeDateTimePartInOutputFileName:$IncludeDateTimePartInReportFileName `
                                                                                -DateTimePartInOutputFileName $StartTimeForServer `
                                                                                -OutputFileNameSuffix 'corruptions_details' -BreakIfError:$BreakOnReportCreationError
            
            
        }
        
        #endregion
        
        #region Initial EMS test
        
        If ((Test-ExchangeCmdletsAvailability) -ne $true) {
            
            [String]$MessageText = "The function Invoke-MailboxDatabasesReapairs need to be run using Exchange Management Shell"
            
            $MessagesToReport += "`n$MessageText"
            
            If ($CreateReportFile -ne "None") {
                
                $MessagesToReport | Out-File -FilePath $PerServerReportFileMessages.OutputFilePath
                
            }
            
            Throw $MessageText
            
        }
        
        #endregion
        
        #If parameter has value which is array something like this "[Text2, System.String[]]" will be added to log/message
        [String]$MessageText = "Operation started at {0} with parameters {1} " -f $StartTimeForServer, "Not implemented yet :-(" #, $PSBoundParameters.GetEnumerator()
        
        $MessagesToReport += $MessageText
        
        Write-Verbose -Message $MessageText
        
        $MailboxServer = (Get-MailboxServer -Identity $ComputerNetBIOSName)
        
        [Int]$MailboxServerCount = (Measure-Object -InputObject $MailboxServer).Count
        
        If ($MailboxServerCount -gt 1) {
            
            [String]$MessageText = "You can use this function to perform actions on only one server at once."
            
            $MessagesToReport += "`n$MessageText"
            
            
            If ($CreateReportFile -ne "None") {
                
                
                $MessagesToReport | Out-File -FilePath $PerServerReportFileMessages.OutputFilePath
                
            }
            
            Throw $MessageText
            
        }
        Elseif ($MailboxServerCount -ne 1) {
            
            [String]$MessageText = "Server {0} is not a Exchange mailbox server" -f $MailboxServer
            
            $MessagesToReport += "`n$MessageText"
            
            
            If ($CreateReportFile -ne "None") {
                
                
                $MessagesToReport | Out-File -FilePath $PerServerReportFileMessages.OutputFilePath
                
            }
            
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
            
            $MessagesToReport += "`n$MessageText"
            
            Write-Verbose -Message $MessageText
            
        }
        Catch {
            
            
            [String]$MessageText = "Server {0} is not reachable or PowerShell remoting is not enabled on it." -f $ComputerNetBIOSName
            
            $MessagesToReport += "`n$MessageText"
            
            Write-Verbose -Message $MessageText
            
            #Decision based on 
            
        }
        
        Finally {
            
            If (($MailboxServerVersion.Major -eq 14 -and $MailboxServerVersion.Minor -lt 1) -or $MailboxServerVersion.Major -ne 14) {
                
                [String]$MessageText = "This function can be used only on Exchange Server 2010 SP1 or newer version of Exchange Server 2010."
                
                $MessagesToReport += "`n$MessageText"
                
                If ($CreateReportFile -ne "None") {
                    
                    $MessagesToReport | Out-File -FilePath $PerServerReportFileMessages.OutputFilePath
                    
                }
                
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
                    
                    [String]$MessageText = "Database {0} is not currently active on {1} and can't be checked" -f $CurrentDatabase.Name, $ComputerNetBIOSName
                    
                    $MessagesToReport += "`n$MessageText"
                    
                    Write-Error -Message $MessageText
                    
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
            
            $MessagesToReport += "`n$MessageText"
            
            Write-Error -Message $MessageText
            
        }
        
    }
    
    Process {
        
        $ActiveDatabases | ForEach-Object -Process {
            
            #Current time need to be compared between localhost and destination host to avoid mistakes
            $StartTimeForDatabase = Get-Date
            
            If ($CreateReportFile -eq 'CreatePerDatabase') {
                
                if ($PSBoundParameters.ContainsKey('$ReportFileNamePrefix')) {
                    
                    [String]$ReportPerDatabaseNamePrefix = $ReportFileNamePrefix
                    
                }
                Else {
                    
                    [String]$ReportPerDatabaseNamePrefix = "{0}_{1}_IntegrityChecks" -f $_.Name, $_.Server
                    
                }
                
                
                $PerDatabaseReportFileMessages = New-OutputFileNameFullPath -OutputFileDirectoryPath $ReportFileDirectoryPath -OutputFileNamePrefix $ReportPerDatabaseNamePrefix `
                                                                            -IncludeDateTimePartInOutputFileName:$IncludeDateTimePartInReportFileName `
                                                                            -DateTimePartInOutputFileName $StartTimeForDatabase `
                                                                            -OutputFileNameSuffix 'messages' -BreakIfError:$BreakOnReportCreationError
                
                $PerDatabaseReportFileEvents = New-OutputFileNameFullPath -OutputFileDirectoryPath $ReportFileDirectoryPath -OutputFileNamePrefix $ReportPerDatabaseNamePrefix `
                                                                          -IncludeDateTimePartInOutputFileName:$IncludeDateTimePartInReportFileName `
                                                                          -DateTimePartInOutputFileName $StartTimeForDatabase `
                                                                          -OutputFileNameSuffix 'events' -BreakIfError:$BreakOnReportCreationError
                
                $PerDatabaseReportFileCorruptionsDetails = New-OutputFileNameFullPath -OutputFileDirectoryPath $ReportFileDirectoryPath -OutputFileNamePrefix $ReportPerDatabaseNamePrefix `
                                                                                      -IncludeDateTimePartInOutputFileName:$IncludeDateTimePartInReportFileName `
                                                                                      -DateTimePartInOutputFileName $StartTimeForDatabase `
                                                                                      -OutputFileNameSuffix 'corruption_details' -BreakIfError:$BreakOnReportCreationError
                
            }
            
            #Check current status of database - if is still mounted on correct server - if not exit from current loop iteration
            
            Try {
                
                $CurrentDatabase = (Get-MailboxDatabase -Identity $_.Name | Where-Object -FilterScript { $_.Server -match $ComputerNetBIOSName } | Select-Object -Property Name)
                
            }
            Catch {
                
                [String]$MessageText = "Database {0} is not currently active on {1} and can't be checked" -f $CurrentDatabase, $ComputerNetBIOSName
                
                $MessagesToReport += "`n$MessageText"
                
                Write-Error -Message $MessageText
                
                #Exit from current loop iteration - check the next database
                Continue
                
            }
            
            [String]$MessageText = "Invoking command for repair database {0} on mailbox server {1}" -f $CurrentDatabase.Name, $ComputerFQDNName
            
            $MessagesToReport += "`n$MessageText"
            
            Write-Verbose -Message $MessageText
            
            Try {
                
                
                if ($MailboxServerVersion.Major -eq 14) {
                    
                    New-MailboxRepairRequest -Database $_.Name -CorruptionType "SearchFolder", "AggregateCounts", "ProvisionedFolder", "FolderView", "MessagePTagCn" -DetectOnly:$DetectOnly -ErrorAction Stop
                    
                }
                
                else {
                    
                    [String]$MessageText = "Something goes wrong - Exchange version unknown"
                    
                    $MessagesToReport += "`n$MessageText"
                    
                    If ($CreateReportFile -ne "None") {
                        
                        $MessagesToReport | Out-File -FilePath $PerServerReportFileMessages.OutputFilePath
                        
                    }
                    
                    Throw $MessageText
                    
                }
                
            }
            Catch {
                
                [String]$MessageText = "Under invoking New-MailboxRepairRequest on {0} error occured: {1} " -f $CurrentDatabase.Name, $Error[0]
                
                $MessagesToReport += "`n$MessageText"
                
                
                If ($CreateReportFile -ne "None") {
                    
                    
                    $MessagesToReport | Out-File -FilePath $PerServerReportFileMessages.OutputFilePath
                    
                }
                
                Throw $MessageText
                
            }
            
            Start-Sleep -Seconds 1
            
            [Int]$ExpectedDurationStartWait = 5
            
            [int]$i = 1
            
            do {
                
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
                
                If ((Measure-Object -InputObject $MonitoredEvents).count -ge 1) {
                    
                    [String]$MessageText = "Events Found {0}" -f $MonitoredEvents
                    
                    $MessagesToReport += "`n$MessageText"
                    
                    Write-Verbose -Message $MessageText
                    
                    $EventsToReport += $MonitoredEvents
                    
                }
                
                $ErrorEvents = ($MonitoredEvents | Where-Object -FilterScript { $_.EventId -ne 10059 })
                
                $ErrorEventsFound = ((Measure-Object -InputObject $ErrorEvents).count -ge 1)
                
                If ($ErrorEventsFound) {
                    
                    [String]$MessageText = "Under checking database {0} error occured - event ID  {1} " -f $_.Name, $ErrorEvents.EventId
                    
                    $MessagesToReport += "`n$MessageText"
                    
                    Write-Error -Message $MessageText
                    
                    Clear-Variable -Name MonitoredEvents
                    
                    $StopTimeForDatabase = Get-Date
                    
                    $DurationTimeForDatabase = New-TimeSpan -Start $StartTimeForDatabase -End $StopTimeForDatabase
                    
                    #Exit from current loop iteration - check the next database
                    Continue
                    
                }
                
                $StartRepairEvent = ($MonitoredEvents | Where-Object -FilterScript { $_.EventId -eq 10059 })
                
                $StartRepairEventFound = ((Measure-Object -InputObject $StartRepairEvent).count -eq 1)
                
                If (-not $MonitoredEvents) {
                    
                    Start-Sleep -Seconds $CheckProgressEverySeconds
                    
                }
                
                Else {
                    
                    [String]$StartTime = Get-Date -Date $($StartRepairEvent.TimeGenerated) -format yyyyMMdd-HHmm
                    
                    [String]$MessageText = "Repair request for database {0} started at {1}" -f $_.Name, $StartRepairEvent.TimeGenerated
                    
                    $MessagesToReport += "`n$MessageText"
                    
                    Write-Verbose -Message $MessageText
                    
                    Clear-Variable -Name MonitoredEvents
                    
                }
                
            }
            while ($StartRepairEventFound -eq $false -and $ErrorEventsFound -eq $false)
            
            Start-Sleep -Seconds 1
            
            [int]$i = 1
            
            do {
                
                $MonitoredEvents = Get-EventsBySource -ComputerName $ComputerFQDNName -LogName "Application" -ProviderName "MSExchangeIS Mailbox Store" -EventID 10045, 10048, 10049, 10050, 10051 -StartTime $StartTimeForDatabase -Verbose:$false
                
                If ((Measure-Object -InputObject $MonitoredEvents).count -ge 1) {
                    
                    [String]$MessageText = "Events Found {0}" -f $MonitoredEvents
                    
                    $MessagesToReport += "`n$MessageText"
                    
                    Write-Verbose -Message $MessageText
                    
                    $MonitoredEvents | ForEach-Object -Process {
                        
                        $EventsToReport += $_
                        
                    }
                    
                }
                
                $ErrorEvents = ($MonitoredEvents | Where-Object -FilterScript { $_.EventId -ne 10048 })
                
                $ErrorEventsFound = ((Measure-Object -InputObject $ErrorEvents).count -ge 1)
                
                If ($ErrorEventsFound) {
                    
                    [String]$MessageText = "Under checking database {0} error occured - event ID  {1} " -f $_.Name, $ErrorEvents.EventId
                    
                    Write-Error -Message $MessageText
                    
                    $MessagesToReport += "`n$MessageText"
                    
                    Clear-Variable -Name MonitoredEvents
                    
                    
                    #Exit from current loop iteration - check the next database
                    
                    $StopTimeForDatabase = Get-Date
                    
                    $DurationTimeForDatabase = New-TimeSpan -Start $StartTimeForDatabase -End $StopTimeForDatabase
                    
                    Continue
                    
                }
                
                $StopRepairEvent = ($MonitoredEvents | Where-Object -FilterScript { $_.EventId -eq 10048 })
                
                $StopRepairEventFound = ((Measure-Object -InputObject $StopRepairEvent).count -eq 1)
                
                If (-not $StopRepairEventFound) {
                    
                    If ($DisplayProgressBar) {
                        
                        [String]$MessageText = "Database {0} repair request  is in progress." -f $_.Name
                        
                        Write-Progress -Activity $MessageText -Status "Completion percentage is only confirmation that something is happening :-)" -PercentComplete (($i / ($ExpectedDurationTimeMinutes * 60)) * 100)
                        
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
                    
                    $MessagesToReport += "`n$MessageText"
                    
                    Write-Verbose -Message $MessageText
                    
                    $CorruptionFoundEvents = Get-EventsBySource -ComputerName $ComputerFQDNName -LogName "Application" -ProviderName "MSExchangeIS Mailbox Store" -EventID 10062 -StartTime $StartTimeForDatabase -Verbose:$false
                    
                    $CorruptionFoundEventsCount = (Measure-Object -InputObject $CorruptionFoundEvents).count
                    
                    if ($CorruptionFoundEventsCount -ge 1) {
                        
                        $Events10062Details = Parse10062Events -Events $CorruptionFoundEvents
                        
                        $Events10062Details | ForEach-Object -Process {
                            
                            $Events10062DetailsToReport += $_
                            
                        }
                        
                    }
                    
                    If ($CreateReportFile -eq 'CreatePerDatabase') {
                        
                        $MessagesToReport | Set-Content -Path $PerDatabaseReportFileMessages.OutputFilePath
                        
                        $EventsToReport | Add-Content -Path $PerDatabaseReportFileEvents.OutputFilePath
                        
                        $Events10062DetailsToReport | Add-Content -Path $PerDatabaseReportFileCorruptionsDetails.OutputFilePath
                        
                        Clear-Variable -Name MessagesToReport -ErrorAction 'SilentlyContinue'
                        
                        Clear-Variable -Name EventsToReport -ErrorAction 'SilentlyContinue'
                        
                        Clear-Variable -Name Event10062DetailsToReport -ErrorAction 'SilentlyContinue'
                        
                    }
                    
                    If ($DisplaySummary) {
                        
                        Write-Output -InputObject $CorruptionFoundEvents
                        
                    }
                    
                    
                }
                
            }
            
            while ($StopRepairEventFound -eq $false -and $ErrorEventsFound -eq $false)
            
        }
        
    }
    
    End {
        
        $StopTimeForServer = Get-Date
        
        $DurationTimeForServer = New-TimeSpan -Start $StartTimeForServer -End $StopTimeForServer
        
        [String]$MessageText = "Operation for server {0} ended at {1}, operation duration time: {2} days, {3} hours, {4} minutes, {5} seconds" `
        -f $ComputerNetBIOSName, $StopTimeForServer, $DurationTimeForServer.Days, $DurationTimeForServer.Hours, $DurationTimeForServer.Minutes, $DurationTimeForServer.Seconds
        
        $MessagesToReport += "`n$MessageText"
        
        Write-Verbose -Message $MessageText
        
        #region Write reports files to disk
        
        If ($CreateReportFile -eq "CreatePerServer") {
            
            $MessagesToReport | Set-Content -Path $PerServerReportFileMessages.OutputFilePath
            
            $EventsToReport | Add-Content -Path $PerServerReportFileEvents.OutputFilePath
            
            $Events10062DetailsToReport | Add-Content -Path $PerServerReportFileCorruptionsDetails.OutputFilePath
            
        }
        
        #endregion
        
    }
    
}

Function New-OutputFileNameFullPath {
    
<#

    .SYNOPSIS
    Function intended for preparing filename for output files like reports or logs
   
    .DESCRIPTION
    Function intended for preparing filename for output files like reports or logs based on prefix, middle name part, date, etc. with verification if provided path is writable
    
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
    
    $PerServerReportFileMessages = New-OutputFileNameFullPath -OutputFileDirectoryPath 'C:\Reports' -OutputFileNamePrefix 'Messages' `
                                                              -OutputFileNameMidPart 'COMPUTERNAME' `
                                                              -IncludeDateTimePartInOutputFileName:$true `
                                                              -BreakIfError:$true
    
    $PerServerReportFileMessages | Format-List
    
    OutputFilePath                                           ExitCode ExitCodeDescription
    --------------                                           -------- -------------------
    C:\users\wojtek\Messages-COMPUTERNAME-20151021-0012-.txt        0
    
    
     
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
    
    TODO
    Change/extend type of returned object 
    Change/extend behavior if file exist
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
    
    [String]$ExitCodeDescription = $null
    
    $Result = New-Object -TypeName PSObject
    
    #Convert relative path to absolute path
    [String]$OutputFileDirectoryPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputFileDirectoryPath)
    
    #Assign value to the variable $IncludeDateTimePartInOutputFileName if is not initialized
    If ($IncludeDateTimePartInOutputFileName -and $DateTimePartInOutputFileName -eq $null) {
        
        [String]$DateTimePartInFileNameString = $(Get-Date -format yyyyMMdd-HHmm)
        
    }
    Else {
        
        [String]$DateTimePartInFileNameString = $(Get-Date -Date $DateTimePartInOutputFileName -format yyyyMMdd-HHmm)
        
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
    If (!($IncludeDateTimePartInOutputFileName) -and ($null -ne $OutputFileNameMidPart)) {
        
        [String]$OutputFilePathTemp1 = "{0}\{1}-{2}" -f $OutputFileDirectoryPath, $OutputFileNamePrefix, $OutputFileNameMidPart
        
    }
    Elseif (!($IncludeDateTimePartInOutputFileName) -and ($null -eq $OutputFileNameMidPart)) {
        
        [String]$OutputFilePathTemp1 = "{0}\{1}" -f $OutputFileDirectoryPath, $OutputFileNamePrefix
        
    }
    ElseIf ($IncludeDateTimePartInOutputFileName -and ($null -ne $OutputFileNameMidPart)) {
        
        [String]$OutputFilePathTemp1 = "{0}\{1}-{2}-{3}" -f $OutputFileDirectoryPath, $OutputFileNamePrefix, $OutputFileNameMidPart, $DateTimePartInFileNameString
        
    }
    Else {
        
        [String]$OutputFilePathTemp1 = "{0}\{1}-{2}" -f $OutputFileDirectoryPath, $OutputFileNamePrefix, $DateTimePartInFileNameString
        
    }
    
    If ($null -ne $OutputFileNameSuffix) {
        
        [String]$OutputFilePathTemp = "{0}-{1}.{2}" -f $OutputFilePathTemp1, $OutputFileNameSuffix, $OutputFileNameExtension
        
    }
    Else {
        
        [String]$OutputFilePathTemp = "{0}.{1}" -f $OutputFilePathTemp1, $OutputFileNameExtension
        
    }
    
    #Replacing doubled chars \\ , -- , ..
    [String]$OutputFilePath = "{0}{1}" -f $OutputFilePathTemp.substring(0, 2), (($OutputFilePathTemp.substring(2, $OutputFilePathTemp.length - 2).replace("\\", '\')).replace("--", "-")).replace("..", ".")
    
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