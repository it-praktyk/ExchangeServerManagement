# Invoke-MailboxDatabaseRepair

## SYNOPSIS
The Function intended for performing checks and repairs operations on Exchange Server 2010 SP1 (or newer) mailbox databases

## DESCRIPTION
The function invokes New-MailboxDatabaseRepair cmdlet for all active mailbox database provided by name in function parameter or for all databases on current or provided server.

Native Exchange New-MailboxDatabaseRepair has limitation - to avoid performance issue - that the only one repair database request can be run at once. Additionally output is only directed to Windows Application event log than checking all databases can be time and work consuming.  The Invoke-MailboxDatabases function is a wrapper which allow do it - even as scheduled tasks - for all active databases on mailbox server.

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

## [PARAMETERS](./SYNTAX.md) 

## EXAMPLES
```powershell

[PS] >Invoke-MailboxDatabaseRepair -ComputerName XXXXXXMBX03 -Database All -DisplaySummary:$true -ExpectedDurationTimeMinutes 120 -DetectOnly:$true
```

## BASE REPOSITORY
https://github.com/it-praktyk/Invoke-MailboxDatabaseRepair

## NOTES
AUTHOR: Wojciech Sciesinski, wojciech[at]sciesinski[dot]net  
KEYWORDS: PowerShell, Exchange, New-MailboxRepairRequest

## CURRENT VERSION
- 0.9.6 - 2015-11-22

## [VERSIONS HISTORY](./VERSIONS.md) 


## DEPENDENCIES
-   Function Test-ExchangeCmdletsAvailability - minimum 0.1.2
https://github.com/it-praktyk/Test-ExchangeCmdletsAvailability
-   Function Function Get-EventsBySource - minimum 0.5.0
https://github.com/it-praktyk/Get-EvenstBySource
-   Function New-OutputFileNameFullPath - minimum 0.5.0
https://github.com/it-praktyk/New-OutputFileNameFullPath
-   Module PSLogging - minimum 3.1.0 - original author: Luca Sturlese, http://9to5it.com
https://github.com/it-praktyk/PSLogging

## TO DO
- Current time and timezone need to be compared between localhost and destination host to avoid mistakes
- exit code return need to be implemented
- add support for Exchange 2013 (?) and 2016 (?)
- summary for detected corruption need to be implemented


## LICENSE
Copyright (C) 2015 Wojciech Sciesinski<br />
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
