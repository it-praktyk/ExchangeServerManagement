# Test-Recipient
## SYNOPSIS
Return the Exchange mail object basic data and email addresses as comma-separated string.

## SYNTAX
```powershell
Test-Recipient [-InputFilePath] <String> [[-Prefix] <String>] [[-CheckDomain] <String>] [-DisplayProgressBar] [<CommonParameters>]
```

## DESCRIPTION
Function intended for return basic data about mail object in Exchange Server environment and  email addresses for recipient.

## PARAMETERS
### -InputFilePath &lt;String&gt;
File contains e.g. alias, guid, email address or any other recipients identifiers which can be used with Identity for Get-Recipient command.
```
Required?                    true
Position?                    1
Default value
Accept pipeline input?       false
Accept wildcard characters?  false
```
 
### -Prefix &lt;String&gt;
Provide address prefix, default is SMTP.
```
Required?                    false
Position?                    2
Default value                smtp
Accept pipeline input?       false
Accept wildcard characters?  false
```
 
### -CheckDomain &lt;String&gt;
Provide domain name what need to be checked if a recipient have email addresses from it.
```
Required?                    false
Position?                    3
Default value
Accept pipeline input?       false
Accept wildcard characters?  false
```
 
### -DisplayProgressBar &lt;SwitchParameter&gt;
Select if a progress bar should be displayed under checking. Displaying progress bar can increase execution time.
```
Required?                    false
Position?                    named
Default value                False
Accept pipeline input?       false
Accept wildcard characters?  false
```

## INPUTS


## NOTES
AUTHOR: Wojciech Sciesinski, wojciech[at]sciesinski[dot]net
KEYWORDS: PowerShell, Exchange, SMTP, addresses
  
VERSIONS HISTORY
- 0.1.0 - 2016-05-26 - Initial version published on GitHub

TODO  
- check if Exchange cmdlets are available
- add support for recipients from pipeline
- implement better domain checking
- verify names of internal variables
- replace ForEach aliasses
- update help
   
    
LICENSE  
Copyright (c) 2016 Wojciech Sciesinski  
This function is licensed under The MIT License (MIT)  
Full license text: https://opensource.org/licenses/MIT

## EXAMPLES
### EXAMPLE 1
```powershell
PS C:\>Test-Recipient -InputFilePath .\RecipientsToGetSMTP.txt -CheckDomain "example.mail.onmicrosoft.com"

Addresses                           : smtp:EX2010A_USER8@ex2010a.lab,SMTP:EX2010A_USER8@ex2013a.contoso.com,smtp:EX2010A_USER
                                      8@example.mail.onmicrosoft.com
IsValid                             : True
Prefix                              : smtp
ContainsAddressFromCheckedDomain    : True
Alias                               : EX2010A_USER8
IsMissed                            : False
RecipientType                       : UserMailbox
AddressesCount                      : 3
DisplayName                         : USER8 EX2010A

Addresses                           : smtp:EX2010A_USER9@ex2010a.lab,SMTP:EX2010A_USER9@ex2013a.contoso.com
IsValid                             : True
Prefix                              : smtp
ContainsAddressFromCheckedDomain    : False
Alias                               : EX2010A_USER9
IsMissed                            : False
RecipientType                       : UserMailbox
AddressesCount                      : 2
DisplayName                         : USER9 EX2010A
```


