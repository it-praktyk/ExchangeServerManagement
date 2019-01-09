Function CumulateIISLogsFiles {
    
<#
	.SYNOPSIS
	Function intended for 
   
	.DESCRIPTION
	
  	.PARAMETER FirstParameter
    
    .PARAMETER UseGMT
    IIS logs can be rotated based on UTC/GMT time - there is no time difference between Greenwich Mean Time (GMT) and Coordinated Universal Time (UTC)
    
  
	.EXAMPLE
     
	.LINK
	https://github.com/it-praktyk/ExchangeActiveSyncStatistics
	
	.LINK
	https://www.linkedin.com/in/sciesinskiwojciech
	      
	.NOTES
	AUTHOR: Wojciech Sciesinski, wojciech[at]sciesinski[dot]net
	KEYWORDS: PowerShell, Exchange, IIS, ActiveSync
   
	VERSIONS HISTORY
	0.1.0 -  2015-04-09 - initial release

		
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

    TO DO

    Check if Exchange commands are available
    Implement copy using BitTransfer

	
	
   
#>
    
    param (
        
        [parameter(mandatory = $true)]
        [ValidateSet("All", "List", "CAS", "MBX")]
        [String]$ServersScope = "CAS",
        
        [parameter(Mandatory = $true)]
        [Bool]$UseGMT,
        
        [parameter(Mandatory = $true)]
        [DateTime]$NewerThan = '03/20/2015',
        
        [parameter(Mandatory = $true)]
        [DateTime]$OlderThan = '07/04/2015',
        
        [parameter(Mandatory = $true)]
        [String]$LogSourceDirectoryPath = '\\f$\inetpub\logs\LogFiles',
        
        [parameter(Mandatory = $true)]
        [String]$DestinationComputerName = "localhost",
        
        [parameter(Mandatory = $true)]
        [String]$DestinationRootDirectoryPath = 'O$\IISLogsArchive',
        
        [parameter(Mandatory = $false)]
        [Bool]$CreateDestinationFolderIfNotExist = $true
        
        
        
    )
    
    BEGIN {
        
        $CasServers = $(Get-ExchangeServer | Where { $_.IsClientAccessServer -eq $true } | Sort Name)
        
        [String]$FullDestinationRootDirectoryPath = '\\' + $DestinationComputerName + '\' + $DestinationRootDirectoryPath
        
    }
    
    PROCESS {
        
        $CasServers | ForEach {
            
            $CurrentServerName = $_.Name
            
            [String]$FullCurrentLogPath = '\\' + $CurrentServerName + '\' + $LogSourceDirectoryPath + '\'
            
            #Only one level of subfolders is supported now - not recurese
            $SubfoldersInSource = Get-ChildItem -Path $FullCurrentLogPath | Where { $_.PSIsContainer }
            
            #SubfoldersInSourceCount add ?
            
            $SubfoldersInSource | ForEach {
                
                $CurrentDestinationPath = $FullDestinationRootDirectoryPath + '\' + $CurrentServername + '\' + $_.Name
                
                if (Test-Path -LiteralPath $CurrentDestinationPath -PathType Container) {
                    
                    
                    
                }
                Elseif (!(Test-Path $CurrentDestinationPath -PathType Container) -and $CreateIfNotExist) {
                    
                    New-Item -Path $CurrentDestinationPath -Force -Type container | Out-Null
                    
                }
                
                #This need to be checked - probably two files by day need to be parsed and next cumulated
                $FilesInSource = Get-ChildItem -Path $_.FullName | Where { $_.LastWriteTimeUTC -gt $NewerThan -AND $_.LastWriteTimeUTC -lt $OlderThan -and $_.Extension -eq '.log' }
                
                If ($FilesInSource) {
                    
                    
                    $FilesInSource | ForEach {
                        
                        Write-Verbose "The file $_.FullName is coppied to $CurrentDestinationPath"
                        
                        Copy-Item $_.FullName -Destination $CurrentDestinationPath
                        
                    }
                    
                }
                
            }
            
        }
        
    }
    
}