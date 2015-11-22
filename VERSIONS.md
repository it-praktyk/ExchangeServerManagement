# Invoke-MailboxDatabaseRepair

## BASE REPOSITORY
https://github.com/it-praktyk/Invoke-MailboxDatabaseRepair

## VERSIONS HISTORY
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
- 0.9.3 - 2015-11-16 - Compatibility with PowerShell 2.0 corrected, reporting for details of events corrected, DisplaySummary parameter deleted
- 0.9.4 - 2015-11-18 - Names of code regions added, logging improved
- 0.9.5 - 2015-11-20 - Minor updates
- 0.9.6 - 2015-11-22 - Help updated, the parameter ReportFileNameExtension removed
