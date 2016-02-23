# ConvertFrom-O365IPAddressesXMLFile
## SYNOPSIS
Convert the O365IPAddresses.xml file to the custom PowerShell object

## SYNTAX
```powershell
ConvertFrom-O365IPAddressesXMLFile [[-Path] <Object>] [<CommonParameters>]
```

## DESCRIPTION
Function intended for converting to the custom PowerShell object the list of hosts used for Office 365 services published as the O365IPAddresses.xml file.

The list contains addresses (IPv4, IPv6, URL) for what communication can't be proxied on customer/client side.

More information on the Microsoft support page: "Office 365 URLs and IP address ranges", http://bit.ly/1LD8fYv

## PARAMETERS
### -Path &lt;Object&gt;
The xml file containing data like O365IPAddresses.xml downloaded manually.
If the the parameter is ommited the file O365IPAddresses.xml will be downloaded from the Microsoft site and saved with
```
Required?                    false
Position?                    1
Default value                .\O365IPAddresses.xml
Accept pipeline input?       false
Accept wildcard characters?  false
```

## INPUTS


## NOTES
AUTHOR: Wojciech Sciesinski, wojciech[at]sciesinski[dot]net  


KEYWORDS: PowerShell, Exchange, Office 365, XML, proxy

VERSIONS HISTORY
- 0.1.0 - 2016-02-23 - The first working version

TODO
- update help - INPUT/OUTPUTS
- add only summary mode/switch
- add support for downloading the file via proxy with authentication
- add parameter to custom naming downloaded file
- handle errors of a download operation (?)
- add whatif (?)
- check/correct verbose and debug mode

LICENSE  
Copyright (c) 2016 Wojciech Sciesinski  
This function is licensed under The MIT License (MIT)  
Full license text: https://opensource.org/licenses/MIT  

## EXAMPLES
### EXAMPLE 1
```powershell
PS C:\>ConvertFrom-O365IPAddressesXMLFile
```
