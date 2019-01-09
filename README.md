# Remove-MultiSIPFromMailbox
## SYNOPSIS
Function intended for verifying and removing doubled SIP addresses from all mailboxes in Exchange Server environment


## SYNTAX
```powershell
Remove-MultiSIPFromMailbox [-Identity] <Object> [[-CorrectSIPDomain] <string>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION  
Function intended for verifying and removing doubled SIP addresses from all mailboxes in Exchange Server environment.  
Only address in a domain provided in a parameter CorrectSIPDomain will be kept.

## PARAMETERS  
### -Identity &lt;Object&gt;  
The Identity parameter specifies the identity of the mailbox. You can use one of the following values:
- GUID
- Distinguished name (DN)
- Display name
- Domain\Account
- User principal name (UPN)
- LegacyExchangeDN
- SmtpAddress
- Alias

```
Required?                    true
Position?                    0
Accept pipeline input?       true (ByValue, ByPropertyName)  
Parameter set name           (All)  
Aliases                      Mailbox
Dynamic?                     false
```  

### -CorrectSIPDomain &lt;String&gt;
The name of domain for what correct SIPs belong. If the parameter is not set the domain name from PrimarySMTPAddress will be used.

```
Required?                    false
Position?                    1
Accept pipeline input?       false
Parameter set name           (All)
Aliases                      None
Dynamic?                     false
```


## NOTES
AUTHOR: Wojciech Sciesinski, wojciech[at]sciesinski[dot]net  
KEYWORDS: PowerShell, Exchange, SIPAddresses, ProxyAddresses, Lync, migration

VERSIONS HISTORY
- 0.1.0 - 2015-05-27 - First version published on GitHub
- 0.1.2 - 2015-05-29 - Switch address to secondary befor remove, post-report corrected
- 0.1.3 - 2015-05-31 - Help updated
- 0.1.4 - 2015-06-09 - Primary SMTP address added to report file
- 0.2.0 - 2016-02-10 - Report cappabilities removed from the function, input from pipeline added, the license changed to MIT
- 0.2.1 - 2016-02-14 - Help corrected
- 0.3.0 - 2016-02-14 - The function renamed from Remove-DoubledSIPAddresses to Remove-MultiSIPFromMailbox

TODO
- check function behaviour if email address policies are enabled


LICENSE  
Copyright (c) 2016 Wojciech Sciesinski  
This function is licensed under The MIT License (MIT)  
Full license text: https://opensource.org/licenses/MIT

## EXAMPLES

### EXAMPLE 1

Remove doubled SIP based on the correct domain provided as the parameter.  
Operation in the WhatIf mode so SIPAddressesBefore and SIPAddressesAfter are equal.

```powershell

	[PS] > Remove-MultiSIPFromMailbox -Identity aa473815 -WhatIf -Verbose -CorrectSIPDomain contoso.com

    VERBOSE: Mailbox with alias AA473815 has assigned 2 SIP addresses.
    What if: Performing operation "Remove SIP address sip:ingrid.thomes@example.com" on Target "mailbox: AA473815".

    MailboxAlias              : AA473815
    MailboxDisplayName        : Wolters-van der Thomes, IAV (Ingrid)
    MailboxSMTPPrimaryAddress : ingrid.wolters-van.der.thomes@example.com
    MailboxGuid               : b201434a-1f62-4ee4-a446-e0b2bc7badc9
    SIPAddressesBeforeCount   : 2
    SIPAddressesBeforeList    : SIP:Ingrid.Wolters@contoso.com,sip:ingrid.thomes@example.com
    SIPAddressesBefore        : {SIP:Ingrid.Wolters@contoso.com,sip:ingrid.thomes@example.com}
    SIPAddressAfterCount      : 2
    SIPAddressesAfterList     : SIP:Ingrid.Wolters@contoso.com,sip:ingrid.thomes@example.com
    SIPAddressesAfter         : {SIP:Ingrid.Wolters@contoso.com,sip:ingrid.thomes@example.com}

```

### EXAMPLE 2

Remove doubled SIP based on the domain used in PrimarySMTPAddress.

```powershell

	[PS] > Remove-MultiSIPFromMailbox -Identity aa473815 -Verbose

    VERBOSE: Mailbox with alias AA473815 has assigned 2 SIP addresses.
	VERBOSE: SIP address Ingrid.Wolters@contoso.com is incorrect and will be deleted

    MailboxAlias              : AA473815
    MailboxDisplayName        : Wolters-van der Thomes, IAV (Ingrid)
    MailboxSMTPPrimaryAddress : ingrid.wolters-van.der.thomes@tailspintoys.com
    MailboxGuid               : b201434a-1f62-4ee4-a446-e0b2bc7badc9
    SIPAddressesBeforeCount   : 2
    SIPAddressesBeforeList    : SIP:Ingrid.Wolters@contoso.com,sip:ingrid.thomes@tailspintoys.com
    SIPAddressesBefore        : {SIP:Ingrid.Wolters@contoso.com,sip:ingrid.thomes@tailspintoys.com}
    SIPAddressAfterCount      : 1
    SIPAddressesAfterList     : SIP:ingrid.thomes@tailspintoys.com
    SIPAddressesAfter         : {SIP:ingrid.thomes@tailspintoys.com}

```
