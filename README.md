# Set-SIPAddressLikePrimarySMTPAddress

##SYNOPSIS
Function intended for verifying and setting SIP addresses equal to PrimarySMTPAddress for all mailboxes in Exchange Server environment
    
##DESCRIPTION 
Function intended for verifying and setting SIP addresses equal to PrimarySMTPAddress for all mailboxes in Exchange Server environment,
any other addresses will be removed, also if more than one SIP address was assigned to a mailbox

##PARAMETERS
        
### CreateLogFile
By default log file is created
    
### LogFileDirectoryPath
By default log files are stored in the subfolder "logs" in current path, if the "logs" subfolder is missed will be created.
    
### LogFileNamePrefix
Prefix used for creating rollback/report files name. Default is "SIPs-Corrected-"

##Base repository
https://github.com/it-praktyk/Set-SIPAddressLikePrimarySMTPAddress

##NOTES
AUTHOR: Wojciech Sciesinski, wojciech[at]sciesinski[dot]net
KEYWORDS: PowerShell, Exchange, SIPAddresses, ProxyAddresses, Lync, migration
   
##VERSIONS HISTORY
0.1.0 - 2015-06-09 - First version published on GitHub, based mostly on Remove-DoubledSIPAddresses v. 0.1.4
0.1.1 - 2015-06-10 - Verbose logging corrected, WhatIf implemented
	
##TODO
- check if Exchange cmdlets are available

	
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
