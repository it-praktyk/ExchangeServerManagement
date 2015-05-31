# Remove-DoubledSIPAddresses
Function intended for verifying and removing doubled SIP addresses from all mailboxes in Exchange Server environment

## SYNOPSIS
Function used to remove SIP addresses from mailboxes in Exchange Server environment which have duplicated SIPs
    
##DESCRIPTION 
Function intended for verifying and removing doubled SIP addresses from all mailboxes in Exchange Server environment. Only address in a domain provided in a parameter CorrectSIPDomain will be keep.
	
##PARAMETER CorrectSIPDomain
Name of domain for which correct SIPs belong
	
##PARAMETER CreateLogFile
By default log file is created
	
##PARAMETER LogFileDirectoryPath
By default log files are stored in the sub-folder "logs" in current path, if the "logs" subfolder is missed will be created
	
##PARAMETER LogFileNamePrefix
Prefix used for creating rollback files name. Default is "SIPs-Removed-"

##Base repository
https://github.com/it-praktyk/Remove-DoubledSIPAddresses

##NOTES
AUTHOR: Wojciech Sciesinski, wojciech[at]sciesinski[dot]net
KEYWORDS: PowerShell, Exchange, SIPAddresses, ProxyAddresses, Lync, migration
   
##VERSIONS HISTORY
0.1.0 - 2015-05-27 - First version published on GitHub
0.1.2 - 2015-05-29 - Switch address to secondary befor remove, post-report corrected
0.1.3 - 2015-05-31 - Help updated
	
##TODO
- check if Exchange cmdlets are available
- check function behaviour if email address policies are enabled
	
##LICENSE
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
