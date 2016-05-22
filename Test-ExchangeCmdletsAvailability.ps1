Function Test-ExchangeCmdletsAvailability {
    
<#
    .SYNOPSIS
    Verify if the function is run in Exchange Management Shell and the session to Exchange server is established.
   
    .DESCRIPTION
    
    Function to verify if the function is run in Exchange Management Shell by test the Function PSProvider content for selected cmdlet - default is Get-ExchangeServer. 
    Additionally state of the session can be checked by invoking Get-ExchangeServer and checking amount of returned objects.
        
    .PARAMETER CmdletForCheck
    Cmdlet which availability will be tested
    
    .PARAMETER CheckExchangeServersAvailability
    Try read list of available Exchange servers
  
    .EXAMPLE
    [PS] > Test-ExchangeCmdletsAvailability -CmdletForCheck Get-Mailbox
    
    If the function is run in Exchange Management Shell.
    
    .EXAMPLE
    [PS] >Test-ExchangeCmdletsAvailability
    1

    If the function is doesn't run in Exchange Management Shell.

    .OUTPUTS
    
    The function return codes like below
    
    0 = everythink OK 
    1 = cmdlets don't available
    2 = cmdlets available but Exchange servers don't available in the session
     
    .LINK
    https://github.com/it-praktyk/Test-ExchangeCmdletsAvailability
    
    .LINK
    https://www.linkedin.com/in/sciesinskiwojciech
          
    .NOTES
    AUTHOR: Wojciech Sciesinski, wojciech[at]sciesinski[dot]net
    KEYWORDS: PowerShell, Exchange
   
    VERSIONS HISTORY
    - 0.1.0 - 2015-05-25 - Initial release
    - 0.1.1 - 2015-05-25 - Variable renamed, help updated, simple error handling added
    - 0.1.2 - 2015-07-06 - Corrected
	- 0.2.0 - 2016-05-22 - The license changed to MIT, returned types extended
    - 0.2.1 - 2016-05-22 - Workaround for pass Pester test added
        
	LICENSE
	Copyright (c) 2016 Wojciech Sciesinski
    This function is licensed under The MIT License (MIT)
    Full license text: https://opensource.org/licenses/MIT
   
#>
    
    
    [CmdletBinding()]
    param (
        
        [parameter(mandatory = $false)]
        [String]$CmdletForCheck = "Get-ExchangeServer",
        [parameter(Mandatory = $false)]
        [Switch]$CheckExchangeServersAvailability
        
    )
    
    BEGIN {
        
        $ReturnCode = 2
        
    }
    
    PROCESS {
        
        $CmdletAvailable = Test-Path -Path Function:$CmdletForCheck
        
        if ($CmdletAvailable -and ($CheckExchangeServersAvailability.IsPresent)) {
            
            Try {
                
                $ReturnedServers = Get-ExchangeServer
                
                $ReturnedServerCount = ($ReturnedServers | Measure-Object).Count
                
                If ($CmdletAvailable -and ($ReturnedServerCount -ge 1)) {
                    
                    $ReturnCode = 0
                    
                }
                
            }
            Catch {
                
                $ReturnCode = 2
                
            }
            
        }
        elseif ($CmdletAvailable) {
            
            $ReturnCode = 0
            
        }
        Else {
            
            $ReturnCode = 1
            
        }
        
    }
    
    END {
        
        Return $ReturnCode
        
    }
    
}