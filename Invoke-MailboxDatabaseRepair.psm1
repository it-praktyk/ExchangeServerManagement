function Invoke-MailboxDatabaseRepair {
    
    <#
    For help check en-us\Invoke-MailboxDatabaseRepair.psm1-Help.xml    
    Current version: 0.9.5 - 2015-11-20          
    
    Please remember about changing version number also in other files like:
    psd1, a xml help file, a README file and in variable $Version
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
        [Parameter(mandatory = $false)]
        [int]$ExpectedDurationTimeMinutes = 150,
        [parameter(Mandatory = $false, ParameterSetName = "Reports")]
        [ValidateSet("CreatePerServer", "CreatePerDatabase", "None")]
        [String]$CreateReportFile = "CreatePerServer",
        [parameter(Mandatory = $false, ParameterSetName = "Reports")]
        [String]$ReportFileDirectoryPath = ".\reports\",
        [parameter(Mandatory = $false, ParameterSetName = "Reports")]
        [String]$ReportFileNamePrefix = $null,
        [parameter(Mandatory = $false, ParameterSetName = "Reports")]
        [String]$ReportFileNameMidPart = $null,
        [parameter(Mandatory = $false, ParameterSetName = "Reports")]
        [Bool]$IncludeDateTimePartInReportFileName = $true,
        [parameter(Mandatory = $false, ParameterSetName = "Reports")]
        [Nullable[DateTime]]$DateTimePartInReportFileName = $null,
        [parameter(Mandatory = $false, ParameterSetName = "Reports")]
        [String]$ReportFileNameExtension = ".txt",
        [parameter(Mandatory = $false, ParameterSetName = "Reports")]
        [Bool]$BreakOnReportCreationError = $true
        
    )
    
    #endregion
    
    Begin {
        
        Write-Verbose -Message "The script is running from $PSScriptRoot"
        
        #region Initialize variables
        
        [Version]$ScriptVersion = "0.9.5"
        
        $ActiveDatabases = @()
        
        $EventsToReport = @()
        
        [Bool]$WriteToFile = $false
        
        $Events10062DetailsToReport = @()
        
        [Bool]$IsRunningOnLocalhost = $false
        
        [Bool]$StartRepairEventFound = $false
        
        [Bool]$StopRepairEventFound = $false
        
        [DateTime]$StartTimeForServer = $([DateTime]::Now)
        
        [Int]$CorruptionFoundEventstPerServerCount = 0
        
        [Int]$DatabaseCount = 0
        
        [Int]$DatabaseFailCount = 0
        
        [Int]$DatabaseSuccessCount = 0
        
        [String]$RunMode = "DetectAndFix"
        
        if ($DetectOnly) {
            
            $RunMode = "DetectOnly"
            
        }
        
        #endregion
        
        #region Load External modules and dependencies
        
        If ($PSVersionTable.psversion.major -lt 3) {
            
            $RequiredFiles = Get-ChildItem -Path "$PSScriptRoot\Nested\" -Filter "*.ps1"
            
            $RequiredFiles | Foreach-Object -Process { Import-Module $_.FullName -ErrorAction Stop -verbose:$false }
            
        }
        
        #endregion
        
        #region Initialize a computer names
        
        If ($ComputerName -eq 'localhost') {
            
            [String]$ComputerFQDNName = (([Net.DNS]::GetHostEntry("localhost")).HostName).ToUpper()
            
            [String]$ComputerNetBIOSName = ($ComputerFQDNName.Split(".")[0]).ToUpper()
            
        }
        ElseIf ($ComputerName.Contains(".")) {
            
            [String]$ComputerNetBIOSName = ($ComputerName.Split(".")[0]).ToUpper()
            
        }
        else {
            
            [String]$ComputerNetBIOSName = ($ComputerName).ToUpper()
            
            [String]$ComputerFQDNName = (([Net.DNS]::GetHostEntry($ComputerName)).HostName).ToUpper()
            
        }
        
        If (([Net.DNS]::GetHostEntry("localhost")).HostName -eq ([Net.DNS]::GetHostEntry($ComputerName)).HostName) {
            
            [Bool]$IsRunningOnLocalhost = $true
            
        }
        
        [String]$MessageText = "Resolved Exchange server names FQDN: {0} , NETBIOS: {1}" -f $ComputerFQDNName, $ComputerNetBIOSName
        
        $MessageText = Write-LogEntry -ToFile:$false -MessageType INFO -Message $MessageText -TimeStamp -ToScreen
        
        Write-Verbose -Message $MessageText
        
        #endregion
        
        #region Initialize reports files names for reports in PerServer mode
        
        #Creating name for the report, a report file will be used for save initial errors or all messages if CreatePerServer report will be selected
        
        If ($CreateReportFile -ne 'None') {
            
            [Bool]$WriteToFile = $true
            
            If ($PSBoundParameters.ContainsKey('$ReportFileNamePrefix')) {
                
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
        

        
        [String]$MessageText = "Invoke-MailboxDatabaseRepair.ps1 started - version {0} on the server {1} in mode {2}" -f $ScriptVersion.ToString(), $ComputerFQDNName, $RunMode #,  $StartTimeForServer, "Not implemented yet :-(" #, $PSBoundParameters.GetEnumerator()
        
        $MessageText = Write-LogEntry -ToFile:$WriteToFile -LogPath $PerServerMessagesReportFile.FullName -MessageType INFO -Message $MessageText -TimeStamp -EntryDateTime $StartTimeForServer -ToScreen
        
        Write-Verbose -Message $MessageText
        
        #region Initial EMS test
        
        If ((Test-ExchangeCmdletsAvailability) -ne $true) {
            
            [String]$MessageText = "The function Invoke-MailboxDatabasesReapairs need to be run using Exchange Management Shell"
            
            $MessageText = Write-LogEntry -ToFile:$WriteToFile -LogPath $PerServerMessagesReportFile.FullName -MessageType FAILURE -Message $MessageText -TimeStamp -ToScreen
            
            Throw $MessageText
            
        }
        
        #endregion        
        
        #region Initial test for ComputerName parameter
        
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
            
            If ($IsRunningOnLocalhost) {
                
                
                $ExchangeSetupFileVersion = Select-Object -InputObject $(Get-Command -Name Exsetup.exe) -Property FileversionInfo
                
            }
            Else {
                
                $ExchangeSetupFileVersion = Invoke-Command -ComputerName $MailboxServer.Name -ScriptBlock { Select-Object -InputObject $(Get-Command -Name Exsetup.exe) -Property FileversionInfo }
                
            }
            
            [Version]$MailboxServerVersion = ($ExchangeSetupFileVersion.FileVersionInfo).FileVersion
            
            [String]$MessageText = "Discovered version of Exchange Server: {0} " -f $MailboxServerVersion.ToString()
            
            $MessageText = Write-LogEntry -ToFile:$WriteToFile -LogPath $PerServerMessagesReportFile.FullName -MessageType INFO -Message $MessageText -TimeStamp -ToScreen
            
            Write-Verbose -Message $MessageText
            
        } #Try
        Catch {
            
            
            [String]$MessageText = "Server {0} is not reachable or PowerShell remoting is not enabled on it." -f $ComputerFQDNName
            
            $MessageText = Write-LogEntry -ToFile:$WriteToFile -LogPath $PerServerMessagesReportFile.FullName -MessageType FAILURE -Message $MessageText -TimeStamp -ToScreen
            
            Throw $MessageText
            
        } #Catch
        
        Finally {
            
            If (($MailboxServerVersion.Major -eq 14 -and $MailboxServerVersion.Minor -lt 1) -or $MailboxServerVersion.Major -ne 14) {
                
                [String]$MessageText = "This function can be used only on Exchange Server 2010 SP1 or newer version of Exchange Server 2010."
                
                $MessageText = Write-LogEntry -ToFile:$WriteToFile -LogPath $PerServerMessagesReportFile.FullName -MessageType FAILURE -Message $MessageText -TimeStamp -ToScreen
                
                Throw $MessageText
                
            }
            
        } #Finally
        
        #endregion        
        
        #region Initial tests for Database parameter
        
        If ($Database -eq 'All') {
            
            $ActiveDatabases = (Get-MailboxDatabase -Server $ComputerNetBIOSName | Where-Object -FilterScript { $_.Server -match $ComputerNetBIOSName } | Select-Object -Property Name)
            
        }
        Else {
            
            $Database | ForEach-Object -Process {
                
                Try {
                    
                    $CurrentDatabase = (Get-MailboxDatabase -Identity $_ | Where-Object -FilterScript { $_.Server -match $ComputerNetBIOSName } | Select-Object -Property Name)
                    
                }
                Catch {
                    
                    [String]$MessageText = "Database {0} is not currently active on {1} and can't be checked -1 " -f $CurrentDatabase, $ComputerFQDNName
                    
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
        
        [Int]$DatabaseCount = $ActiveDatabasesCount
        
        If ($ActiveDatabasesCount -lt 1) {
            
            [String]$MessageText = "Any database was not found on the server {0}" -f $ComputerNetBIOSName
            
            $MessageText = Write-LogEntry -ToFile:$WriteToFile -LogPath $PerServerMessagesReportFile.FullName -MessageType WARNING -Message $MessageText -TimeStamp -ToScreen
            
            Write-Warning -Message $MessageText
            
        }
        
    }
    
    #endregion
    
    Process {
        
        #region Operations for all databases
        $ActiveDatabases | ForEach-Object -Process {
            
            [String]$CurrentDatabaseName = $_.Name
            
            #Current time need to be compared between localhost and destination host to avoid mistakes
            $StartTimeForDatabase = $([DateTime]::Now)
            
            #region Initialize reports files names for reports in PerDatabase mode            
            
            If ($CreateReportFile -eq 'CreatePerDatabase') {
                
                $EventsToReport = @()
                
                [Bool]$WriteToFile = $true
                
                If ($PSBoundParameters.ContainsKey('$ReportFileNamePrefix')) {
                    
                    [String]$ReportPerDatabaseNamePrefix = $ReportFileNamePrefix
                    
                }
                Else {
                    
                    [String]$ReportPerDatabaseNamePrefix = "{0}_{1}_IntegrityChecks" -f $CurrentDatabaseName, $ComputerNetBIOSName
                    
                }
                
                
                $PerDatabaseMessagesReportFile = $(New-OutputFileNameFullPath -OutputFileDirectoryPath $ReportFileDirectoryPath -OutputFileNamePrefix $ReportPerDatabaseNamePrefix `
                                                                              -IncludeDateTimePartInOutputFileName $IncludeDateTimePartInReportFileName `
                                                                              -DateTimePartInOutputFileName $StartTimeForDatabase `
                                                                              -OutputFileNameSuffix 'messages' -BreakIfError $BreakOnReportCreationError).OutputFilePath
                
                Start-Log -LogPath $PerDatabaseMessagesReportFile.DirectoryName -LogName $PerDatabaseMessagesReportFile.Name -ScriptVersion $ScriptVersion.ToString() | Out-Null
                
                $PerDatabaseEventsReportFile = $(New-OutputFileNameFullPath -OutputFileDirectoryPath $ReportFileDirectoryPath -OutputFileNamePrefix $ReportPerDatabaseNamePrefix `
                                                                            -IncludeDateTimePartInOutputFileName $IncludeDateTimePartInReportFileName `
                                                                            -DateTimePartInOutputFileName $StartTimeForDatabase -OutputFileNameExtension "csv" `
                                                                            -OutputFileNameSuffix 'events' -BreakIfError $BreakOnReportCreationError).OutputFilePath
                
                $PerDatabaseCorruptionsDetailsReportFile = $(New-OutputFileNameFullPath -OutputFileDirectoryPath $ReportFileDirectoryPath -OutputFileNamePrefix $ReportPerDatabaseNamePrefix `
                                                                                        -IncludeDateTimePartInOutputFileName $IncludeDateTimePartInReportFileName `
                                                                                        -DateTimePartInOutputFileName $StartTimeForDatabase -OutputFileNameExtension "csv" `
                                                                                        -OutputFileNameSuffix 'corruption_details' -BreakIfError $BreakOnReportCreationError).OutputFilePath
                
            }
            
            #endregion
            
            #region Check current status of database - if is still mounted on correct server - if not then exit from current loop iteration
            
            Try {
                
                $CurrentDatabaseStatus = (Get-MailboxDatabase -Identity $CurrentDatabaseName | Get-MailboxDatabaseCopyStatus | where { $_.MailboxServer -eq $ComputerNetBIOSName })
                
            }
            Finally {
                
                [String]$MessageText = "Database {0} is not currently active on {1} and can't be checked - 2" -f $CurrentDatabaseName, $ComputerFQDNName
                
            }
            
            If ($CurrentDatabaseStatus.ActiveDatabaseCopy -ne $ComputerNetBIOSName) {
                
                $MessageText = Write-LogEntry -ToFile:$WriteToFile -LogPath $PerServerMessagesReportFile.FullName -MessageType WARNING -Message $MessageText -TimeStamp -ToScreen
                
                Write-Warning -Message $MessageText
                
                $DatabaseFailCount++
                
                #Exit from current loop iteration - check the next database                    
                Continue
                
            }
            
            #endregion            
            
            #region Invoking repair command
            
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
                
                $RepairRequest = New-MailboxRepairRequest -Database $CurrentDatabaseName -CorruptionType "SearchFolder", "AggregateCounts", "ProvisionedFolder", "FolderView", "MessagePTagCn" -DetectOnly:$DetectOnly -ErrorAction Stop
                
                
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
                
                $DatabaseFailCount++
                
                Continue
                
            }
            
            #endregion
            
            Start-Sleep -Seconds 1
            
            [Int]$ExpectedDurationStartWait = 5
            
            [int]$i = 1
            
            #region Checking if repair operation started
            
            [Bool]$MonitoredEventsFound = $false
            
            while ($MonitoredEventsFound -eq $false) {
                
                If ($DisplayProgressBar) {
                    
                    [String]$MessageText = "Waiting for start check operation in {0} operation on the database {1} on the server {2}" -f $RunMode, $CurrentDatabaseName, $ComputerFQDNName
                    
                    Write-Progress -Activity $MessageText -Status "Completion percentage is only confirmation that something is happening :-)" -PercentComplete (($i / ($ExpectedDurationStartWait * 60)) * 100)
                    
                    If (($i += $CheckProgressEverySeconds) -ge ($ExpectedDurationStartWait * 60)) {
                        
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
                    Finally {
                        
                        $ErrorEventsFound = ((Measure-Object -InputObject $ErrorEvents).count -ge 1)
                        
                    }
                    
                    Try {
                        
                        #Filter for event which confirm start of repair operations
                        $StartRepairEvent = ($MonitoredEvents | Where-Object -FilterScript { $_.EventId -eq 10059 })
                        
                    }
                    Finally {
                        
                        $StartRepairEventFound = ((Measure-Object -InputObject $StartRepairEvent).count -eq 1)
                        
                    }
                    
                }
                
                # Operations if errors events found
                If ($ErrorEventsFound) {
                    
                    [String]$MessageText = "Under checking database the {0} on {1}  error occured - event ID  {2} " -f $CurrentDatabaseName, $ComputerFQDNName, $ErrorEvents.EventId
                    
                    
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
                    -f $CurrentDatabaseName, $ComputerFQDNName, $StopTimeForDatabase, $DurationTimeForDatabase.Days, $DurationTimeForDatabase.Hours, $DurationTimeForDatabase.Minutes, $DurationTimeForDatabase.Seconds
                    
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
                    
                    $DatabaseFailCount++
                    
                    #Exit from current loop iteration - check the next database
                    Continue
                    
                }
                
                ElseIf ($StartRepairEventFound) {
                    
                    [DateTime]$StartTimeRepair = $StartRepairEvent.TimeGenerated
                    
                    [String]$MessageText = "Repair request for the database {0} on the server {1} started at {2} with RequestID: {3}" -f $CurrentDatabaseName, $ComputerFQDNName, $StartTimeRepair, $RepairRequest.RequestID
                    
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
            
            #endregion
            
            Start-Sleep -Seconds 1
            
            [int]$i = 1
            
            $MonitoredEventsFound = $false
            
            $ErrorEventsFound = $false
            
            $StopRepairEventFound = $false
            
            #region Loop responsible to check if repair operation finished
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
                    Finally {
                        
                        $ErrorEventsFound = ((Measure-Object -InputObject $ErrorEvents).count -ge 1)
                        
                    }
                    
                    Try {
                        
                        #Filter for event which confirm start of repair operations
                        $StopRepairEvent = ($MonitoredEvents | Where-Object -FilterScript { $_.EventId -eq 10048 })
                        
                    }
                    Finally {
                        
                        $StopRepairEventFound = ((Measure-Object -InputObject $StopRepairEvent).count -eq 1)
                        
                    }
                }
                
                If ($ErrorEventsFound) {
                    
                    [String]$MessageText = "Under checking the database {0} on the server {1} error occured - event ID  {2} " -f $CurrentDatabaseName, $ComputerFQDNName, $ErrorEvents.EventId
                    
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
                    
                    #Check if any 10062 errors occured before termination error                    
                    $CorruptionFoundEvents = Get-EventsBySource -ComputerName $ComputerFQDNName -LogName "Application" -ProviderName "MSExchangeIS Mailbox Store" -EventID 10062 -StartTime $StartTimeForDatabase -Verbose:$false
                    
                    $CorruptionFoundEventsCount = (Measure-Object -InputObject $CorruptionFoundEvents).count
                    
                    If ($CorruptionFoundEventsCount -ge 1) {
                        
                        $CorruptionFoundEventstPerServerCount += $CorruptionFoundEventsCount
                        
                        $EventsToReport += $CorruptionFoundEvents
                        
                        $Events10062Details = Parse10062Events -Events $CorruptionFoundEvents -ComputerName $ComputerFQDNName -DatabaseName $CurrentDatabaseName
                        
                        $Events10062Details | ForEach-Object -Process {
                            
                            $Events10062DetailsToReport += $_
                            
                        }
                        
                    }
                    
                    $StopTimeForDatabase = Get-Date
                    
                    $DurationTimeForDatabase = New-TimeSpan -Start $StartTimeForDatabase -End $StopTimeForDatabase
                    
                    [String]$MessageText = "Operation for the database {0} server {1} end with error2 at {2}, operation duration time: {3} days, {4} hours, {5} minutes, {6} seconds" `
                    -f $CurrentDatabaseName, $ComputerFQDNName, $StopTimeForDatabase, $DurationTimeForDatabase.Days, $DurationTimeForDatabase.Hours, $DurationTimeForDatabase.Minutes, $DurationTimeForDatabase.Seconds
                    
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
                    
                    $DatabaseFailCount++
                    
                    #Exit from current loop iteration - check the next database
                    Continue
                    
                }
                
                # Stop event found
                ElseIf ($StopRepairEventFound) {
                    
                    [String]$MessageText = "Repair request for the database {0} on the server {1} end successfully at {1}" -f $CurrentDatabaseName, $ComputerFQDNName, $StopRepairEvent.TimeGenerated
                    
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
                    
                    If ($CorruptionFoundEventsCount -ge 1) {
                        
                        $CorruptionFoundEventstPerServerCount += $CorruptionFoundEventsCount
                        
                        $EventsToReport += $CorruptionFoundEvents
                        
                        $Events10062Details = Parse10062Events -Events $CorruptionFoundEvents -ComputerName $ComputerFQDNName -DatabaseName $CurrentDatabaseName
                        
                        $Events10062Details | ForEach-Object -Process {
                            
                            $Events10062DetailsToReport += $_
                            
                        }
                        
                    }
                    
                    $StopTimeForDatabase = Get-Date
                    
                    $DurationTimeForDatabase = New-TimeSpan -Start $StartTimeForDatabase -End $StopTimeForDatabase
                    
                    [String]$MessageText = "Operation for the database {0} server {1} end at {2}, operation duration time: {3} days, {4} hours, {5} minutes, {6} seconds; corrupted mailboxes found: {7}" `
                    -f $CurrentDatabaseName, $ComputerFQDNName, $StopTimeForDatabase, $DurationTimeForDatabase.Days, $DurationTimeForDatabase.Hours, $DurationTimeForDatabase.Minutes, `
                    $DurationTimeForDatabase.Seconds, $CorruptionFoundEventsCount
                    
                    switch ($CreateReportFile) {
                        
                        'CreatePerServer' {
                            
                            $MessageText = Write-LogEntry -ToFile:$true -LogPath $PerServerMessagesReportFile.FullName -MessageType INFO -Message $MessageText -TimeStamp -ToScreen
                            
                        }
                        
                        'CreatePerDatabase' {
                            
                            $MessageText = Write-LogEntry -ToFile:$true -LogPath $PerDatabaseMessagesReportFile.FullName -MessageType INFO -Message $MessageText -TimeStamp -ToScreen
                            
                            $EventsToReport | Export-Csv -Path $PerDatabaseEventsReportFile.FullName -Encoding UTF8 -NoTypeInformation -Delimiter ";"
                            
                            If ($CorruptionFoundEventsCount -ge 1) {
                                
                                $Events10062DetailsToReport | Export-Csv -Path $PerDatabaseCorruptionsDetailsReportFile.FullName -Encoding UTF8 -NoTypeInformation -Delimiter ";" -ErrorAction SilentlyContinue
                                
                            }
                            Else {
                                
                                Set-Content -Path $PerDatabaseCorruptionsDetailsReportFile.FullName -value "No corruption events found"
                                
                            }
                            
                            Clear-Variable -Name MessagesToReport -ErrorAction SilentlyContinue
                            
                            Clear-Variable -Name $EventsToReport -ErrorAction SilentlyContinue
                            
                            Clear-Variable -Name Event10062DetailsToReport -ErrorAction SilentlyContinue
                            
                        }
                        
                        'None' {
                            
                            $MessageText = Write-LogEntry -MessageType INFO -Message $MessageText -TimeStamp -ToScreen
                            
                        }
                        
                    }
                    
                    Write-Verbose -Message $MessageText
                    
                    $DatabaseSuccessCount++
                    
                }
                
                else {
                    
                    If ($DisplayProgressBar) {
                        
                        [String]$MessageText = "The database {0} check in mode {1} on the server {2} request  is in progress." -f $CurrentDatabaseName, $RunMode, $ComputerFQDNName
                        
                        Write-Progress -Activity $MessageText -Status "Completion percentage is only confirmation that something is happening :-)" -PercentComplete (($i / ($ExpectedDurationTimeMinutes * 60)) * 100)
                        
                        If (($i += $CheckProgressEverySeconds) -ge ($ExpectedDurationTimeMinutes * 60)) {
                            
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
            
            #endregion Loop responsible to check if repair operation finished
            
            
        }
        
        #endregion
        
        
    }
    
    End {
        
        
        
        #region Write reports per server
        
        If ($CreateReportFile -eq "CreatePerServer") {
            
            $EventsToReport | Export-Csv -Path $PerServerEventsReportFile.FullName -Encoding UTF8 -NoTypeInformation -Delimiter ";"
            
            If ($CorruptionFoundEventstPerServerCount -ge 1) {
                
                $Events10062DetailsToReport | Export-Csv -Path $PerServersCorruptionDetailsReportFile.FullName -Encoding UTF8 -NoTypeInformation -Delimiter ";" -ErrorAction SilentlyContinue
                
            }
            Else {
                
                Set-Content -Path $PerServersCorruptionDetailsReportFile.FullName -value "No corruption events found"
                
            }
            
        }
        
        #endregion                
        
        $StopTimeForServer = Get-Date
        
        $DurationTimeForServer = New-TimeSpan -Start $StartTimeForServer -End $StopTimeForServer
        
        [String]$MessageText = "Operation for the server {0} ended at {1}, operation duration time: {2} days, {3} hours, {4} minutes, {5} seconds; databaseses checks success: {6}, databases checks failed: {7}; corrupted mailboxes found: {8} " `
        -f $ComputerFQDNName, $StopTimeForServer, $DurationTimeForServer.Days, $DurationTimeForServer.Hours, $DurationTimeForServer.Minutes, `
        $DurationTimeForServer.Seconds, $DatabaseSuccessCount, $DatabaseFailCount, $CorruptionFoundEventstPerServerCount
        
        $MessageText = Write-LogEntry -ToFile:$WriteToFile -LogPath $PerServerMessagesReportFile.FullName -MessageType INFO -Message $MessageText -TimeStamp -ToScreen
        
        Write-Verbose -Message $MessageText
        
    }
    
}


Function Parse10062Events {
    
    [cmdletbinding()]
    param (
        
        [parameter(mandatory = $true)]
        $Events,
        [parameter(mandatory = $true)]
        [String]$ComputerName,
        [parameter(mandatory = $true)]
        [String]$DatabaseName
        
    )
    
    BEGIN {
        
        $CorruptionFoundEvents = $Events
        
        $Results = @()
        
        $option = [System.StringSplitOptions]::RemoveEmptyEntries
        
    }
    
    PROCESS {
        
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
                        
                        $Result | Add-Member -type NoteProperty -name ComputerName -value $ComputerName
                        
                        $Result | Add-Member -type NoteProperty -name DatabaseName -value $DatabaseName
                        
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
        
    }
    
    END {
        
        Return $Results
        
    }
    
}

<#
function Parse10062EventsSummaries {
    
    [cmdletbinding()]
    param (
        
        [parameter(mandatory = $true)]
        $Events10062Details
        
    )
    
    #$10062FixedCount = (Measure-Object -InputObject ( $Events10062Details | Where-Object -in
    
}
#>