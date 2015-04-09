# ExchangeActiveSyncStatistics#

PowerShell project intended to create a tool to preparing statistics for ActiveSync usage in Exchange 2007/2010/2013 servers

Extracts from IIS logs will be used, extracts will be prepared using Export-ActiveSyncLog cmdlet

As the result of Export-ActiveSyncLog six files are created
- Users.csv
- Servers.csv
- Hourly.csv
- StatusCodes.csv
- PolicyCompliance.csv
- UserAgents.csv

Unfortunately Export-ActiveSyncLog can handle the only one file at once, so additional agregation is needed. ExchangeActiveSyncStatistics project will be contain tools to automate this operations.


Help for Exchange 2013
- https://technet.microsoft.com/en-us/library/bb123821%28v=exchg.150%29.aspx

Help for Exchange 201
- https://technet.microsoft.com/en-us/library/bb123821%28v=exchg.141%29.aspx

Some links about Export ActiveSyncLog tool
- https://www.simple-talk.com/sysadmin/exchange/reporting-on-mobile-device-activity-using-exchange-2007-activesync-logs/
- http://www.exchangeitpro.com/2012/02/02/understanding-export-activesynclog-part-1-2/


###README.md###

Versions history
- 0.1.0 - 2015-04-10 - basic project explanation added
