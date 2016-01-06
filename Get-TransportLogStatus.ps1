function Get-TransportLogStatus {

<#
    .SYNOPSIS
    Function intended for gather data about transport logs used for message tracking
   
    .DESCRIPTION
    
    .PARAMETER FirstParameter
        
    .OUTPUTS
    System.Object[]
  
    .EXAMPLE
     
    .LINK
    https://github.com/it-praktyk/Get-ExchangeLogsStatus
    
    .LINK
    https://www.linkedin.com/in/sciesinskiwojciech
          
    .NOTES
    AUTHOR: Wojciech Sciesinski, wojciech[at]sciesinski[dot]net
    KEYWORDS: PowerShell, Exchange, logs
   
    Partially based on
	https://mjolinor.wordpress.com/2011/02/11/how-far-back-do-your-message-tracking-logs-really-go/#comments
   
    VERSIONS HISTORY
    0.1.0 -  2016-01-06 - the first working draft

    TODO
	Test connectivity to server(s)
	Server name(s) from command line
	
        
    LICENSE
	Copyright (c) 2016 Wojciech Sciesinski
    This function is licensed under The MIT License (MIT)
    Full license text: http://opensource.org/licenses/MIT
        
    DISCLAIMER
    This script is provided AS IS without warranty of any kind. I disclaim all implied warranties including, without limitation,
    any implied warranties of merchantability or of fitness for a particular purpose. The entire risk arising out of the use or
    performance of the sample scripts and documentation remains with you. In no event shall I be liable for any damages whatsoever
    (including, without limitation, damages for loss of business profits, business interruption, loss of business information,
    or other pecuniary loss) arising out of the use of or inability to use the script or documentation. 
    
    
   
#>

[CmdletBinding()] 
[OutputType([System.Object[]])]

	$TransportServers = get-exchangeserver | Where { $_.serverrole -match 'hubtransport' }
	
	$properties = @('Name', 'MessageTrackingLogEnabled', 'MessageTrackingLogMaxAge', 'MessageTrackingLogMaxDirectorySize',`
	'MessageTrackingLogMaxFileSize', 'MessageTrackingLogPath', 'MessageTrackingLogSubjectLoggingEnabled', 'NewestLog', 'OldestLog',`
	'CurrentLogCount', 'CurrentLogSize', 'CurrentLogHistDepth')
	
	$TransportServers | ForEach-Object { get-transportserver $_.name | select $properties } | % {
		if ($_.messagetrackinglogenabled) {
			$log_unc = “\\$($_.name)\$($_.messagetrackinglogpath -replace “:”, ”$”)”
			$logs = gci $log_unc\*.log | sort lastwritetime
			$newest_log = $logs | select -last 1
			$oldest_log = $logs | select -first 1
			$_.newestlog = $newest_log.name
			$_.oldestlog = $oldest_log.name
			$_.currentlogcount = $logs.count
			$_.currentlogsize = “$([int](($logs | measure length -sum).sum / 1MB)) MB”
			$_.currentloghistdepth = “$([int](($newest_log.LastWriteTime – $oldest_log.creationtime).totaldays)) days”
			$_
		}
	}
}
