# Invoke-MailboxDatabaseRepair

## BASE REPOSITORY
https://github.com/it-praktyk/Invoke-MailboxDatabaseRepair

### ComputerName
Exchange server for which actions should be performed - need to be a mailbox server

### Database
Database identifier - e.g. name - for which action need to be performed. If more than one identifiers need to be separated by commas

### DetectOnly
Set to TRUE if any repair action shouldn't be started

### DisplayProgressBar
If function is used in interactive mode progress bar can be displayed to provide overall information that something is happend.

### CheckProgressEverySeconds
Set interval for progress checking, by default operation progress is checked every 120 seconds

### ExpectedDurationTimeMinutes
Time in minutes used for displaing progress bar

### CreateReportFile
By default report file per server is created

### ReportFileDirectoryPath
By default report files are stored in subfolder "reports" in current path, if "reports" subfolder is missed will be created

### ReportFileNamePrefix
Prefix used for creating report files name. Default is "MBDBs_IntegrityChecks_<SERVER_NETBIOS_NAME>"

### ReportFileNameMidPart
Part of the name which will be used in midle of name

### IncludeDateTimePartInReportFileName
Set to TRUE if report file name should contains part based on current date and time - format yyyyMMdd-HHmm is used

### DateTimePartInReportFileName
Set to date and time which should be used in report file name, by default current date and time is used

### BreakOnReportCreationError
Break function execution if parameters provided for report file creation are not correct or destination file path is not writables