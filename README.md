# Test-ExchangeCmdletsAvailability
## SYNOPSIS
Verify if the function is run in Exchange Management Shell and the session to Exchange server is established.

## SYNTAX
```powershell
Test-ExchangeCmdletsAvailability [[-CmdletForCheck] <String>] [-CheckExchangeServersAvailability] [<CommonParameters>]
```

## DESCRIPTION
Function to verify if the function is run in Exchange Management Shell by test the Function PSProvider content for selected cmdlet - default is Get-ExchangeServer.
Additionally state of the session can be checked by invoking Get-ExchangeServer and checking amount of returned objects.

## PARAMETERS
### -CmdletForCheck &lt;String&gt;
Cmdlet which availability will be tested
```
Required?                    false
Position?                    1
Default value                Get-ExchangeServer
Accept pipeline input?       false
Accept wildcard characters?  false
```

### -CheckExchangeServersAvailability &lt;SwitchParameter&gt;
Try read list of available Exchange servers
```
Required?                    false
Position?                    named
Default value                False
Accept pipeline input?       false
Accept wildcard characters?  false
```

## OUTPUTS

The function return codes like below.

0 = everythink OK  
1 = cmdlets don't available  
2 = cmdlets available but Exchange servers don't available in the session  


## NOTES
AUTHOR: Wojciech Sciesinski, wojciech[at]sciesinski[dot]net  
KEYWORDS: PowerShell, Exchange

VERSIONS HISTORY
- 0.1.0 - 2015-05-25 - Initial release
- 0.1.1 - 2015-05-25 - Variable renamed, help updated, simple error handling added
- 0.1.2 - 2015-07-06 - Corrected
- 0.2.0 - 2016-05-22 - The license changed to MIT, returned types extended

LICENSE  
Copyright (c) 2016 Wojciech Sciesinski  
This function is licensed under The MIT License (MIT)  
Full license text: https://opensource.org/licenses/MIT

## EXAMPLES
### EXAMPLE 1
```powershell
PS C:\>Test-ExchangeCmdletsAvailability
0

```
If the function is run in Exchange Management Shell.

### EXAMPLE 2
```powershell
PS C:\>Test-ExchangeCmdletsAvailability
1

```
If the function is doesn't run in Exchange Management Shell.
