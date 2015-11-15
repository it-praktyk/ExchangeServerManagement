# Invoke-MailboxDatabaseRepair

## SYNOPSIS
The Function intended for performing checks and repairs operations on Exchange Server 2010 SP1 (or newer) mailbox databases

## DESCRIPTION
The function invokes New-MailboxDatabaseRepair cmdlet for all active mailbox database copies on server. Mailbox databases can be also provided by name in function parameter.

Native Exchange New-MailboxDatabaseRepair has limitation - to avoid performance issue - that only one repair database request can be run at once. Additionally output is only directed to Windows Application event log than checking all databases can be time and work consuming.  The Invoke-MailboxDatabases function is a wrapper which allow do it - even as scheduled tasks - for all active databases on mailbox server.

Informations about repair operations you can find on pages
* The Exchange Team Blog: New Support Policy for Repaired Exchange Databases
http://blogs.technet.com/b/exchange/archive/2015/05/01/new-support-policy-for-repaired-exchange-databases.aspx
* White Paper: Database Integrity Checking in Exchange Server 2010 SP1
https://technet.microsoft.com/en-us/library/hh547017%28v=exchg.141%29.aspx
* Nexus: News, Messages about messaging, Matthew Gaskin blog
Using the New-MailboxRepairRequest cmdlet
https://blogs.it.ox.ac.uk/nexus/2012/06/11/new-mailboxrepairrequest/

Possible events for Exchange Server 2010 SP1 and newer
- normal operation
  * 10048 -  The mailbox or database repair request completed successfully.
  * 10059 -  A database-level repair request started.
- errors
  * 10045 - The database repair request failed for provisioned folders. This event ID is created in conjunction with event ID 10049
  * 10049 - The mailbox or database repair request failed because Exchange encountered a problem with the database or another task is running against the database. (Fix for this is ESEUTIL then contact Microsoft Product Support Services)
  * 10050 - The database repair request couldn’t run against the database because the database doesn’t support the corruption types specified in the command. This issue can occur when you run the command from a server that’s running a later version of Exchange than the database you’re scanning.
  * 10051 -  The database repair request was cancelled because the database was dismounted.

## PARAMETERS

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

### DisplaySummary
Set to TRUE if summary should be displayed - summary will contain data about performed operations

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

### ReportFileNameExtension
Set to extension which need to be used for report file, by default ".txt" is used

### BreakOnReportCreationError
Break function execution if parameters provided for report file creation are not correct or destination file path is not writables

## EXAMPLES

[PS] >Invoke-MailboxDatabaseRepair -ComputerName XXXXXXMBX03 -Database All -DisplaySummary:$true -ExpectedDurationTimeMinutes 120 -DetectOnly:$true


## BASE REPOSITORY
https://github.com/it-praktyk/Invoke-MailboxDatabaseRepair

## NOTES
AUTHOR: Wojciech Sciesinski, wojciech[at]sciesinski[dot]net
KEYWORDS: PowerShell, Exchange, New-MailboxRepairRequest

### VERSIONS HISTORY
- 0.1.0 - 2015-07-05 - Initial release
- 0.1.1 - 2015-07-06 - Help updated, TO DO updated
- 0.1.2 - 2015-07-15 - Progress bar added, verbose messages partially suppressed, help next update
- 0.1.3 - 2015-08-11 - Additional checks added to verify provided Exchange server, help and TO DO updated
- 0.2.0 - 2015-08-31 - Corrected checking of Exchange version, output redirected to per mailbox database reports
- 0.3.0 - 2015-09-04 - Added support for Exchange 2013, added support for database repair errors
- 0.3.1 - 2015-09-05 - Corrected but still required testing on Exchange 2013
- 0.4.0 - 2015-09-07 - Support for Exchange 2013 removed, help partially updated, report creation partially implemented TODO section updated
- 0.4.1 - 2015-09-08 - Function reformated
- 0.5.0 - 2015-09-13 - Added support for creation per server log
- 0.5.1 - 2015-09-14 - Help updated, TO DO section updated, DEPENDENCIES section updated
- 0.6.0 - 2015-09-14 - Log creation capabilities updates, parsing 10062 events added
- 0.6.1 - 2015-09-15 - Logging per database corrected
- 0.6.2 - 2015-10-20 - Named regions partially added, function Parse10062Events corrected based on PSScriptAnalyzer rules, function New-OutputFileNameFullPath updated to version 0.4.0, reports per server changed,                         function Get-EventsBySource updated to version 0.5.0
- 0.6.3 - 2015-10-21 - Date for version 0.6.2 corrected
- 0.7.0 - 2015-10-21 - Reports per database changed, function corrected based on PSScriptAnalyzer rules, TO DO updated
- 0.7.1 - 2015-10-22 - Logging podsystem partially updated to use PSLogging module,  function New-OutputFileNameFullPath updated to version 0.4.0 - need to be tested
- 0.8.0 - 2015-10-27 - Major updates especially logging fully updated to use PSLogging module
- 0.8.1 - 2015-10-28 - Corrected, tested
- 0.8.2 - 2015-10-28 - Script reformated
- 0.9.0 - 2015-11-11 - Script switched to module, main function renamed to Invoke-MailboxDatabaseRepair
- 0.9.1 - 2015-11-13 - Functions described in Dependencies moved to subfolder Nested, help moved to xml help file
- 0.9.2 - 2015-11-15 - Function reformated, corrected based on PSScriptAnalyzer rules

### DEPENDENCIES
-   Function Test-ExchangeCmdletsAvailability - minimum 0.1.2
https://github.com/it-praktyk/Test-ExchangeCmdletsAvailability
-   Function Function Get-EventsBySource - minimum 0.5.0
https://github.com/it-praktyk/Get-EvenstBySource
-   Function New-OutputFileNameFullPath - minimum 0.5.0
https://github.com/it-praktyk/New-OutputFileNameFullPath
-   Module PSLogging - minimum 3.1.0 - original author: Luca Sturlese, http://9to5it.com
https://github.com/it-praktyk/PSLogging

### TO DO
- Current time and timezone need to be compared between localhost and destination host to avoid mistakes
- exit code return need to be implemented
- add support for Exchange 2013 (?) and 2016 (?)
- add named regions to easier navigation in code
- summary for detected corruption need to be implemented
- summary per server e.g. checked databases need to be implemented

## LICENSE
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
