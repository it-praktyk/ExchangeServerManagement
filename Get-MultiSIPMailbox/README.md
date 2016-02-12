# Get-MultiSIPMailbox
## SYNOPSIS
Function intended for verifying if mailbox in the Exchange Server environment has assigned more than one SIP address.  


## SYNTAX  
```powershell
Get-MultiSIPMailbox [-Identity]<Object> [<CommonParameters>]
```


## DESCRIPTION
Function intended for verifying if the mailbox in the Exchange Server environment has assigned more than one SIP address.
Only mailboxes with multiplie SIP addresses are returned.  


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


## INPUTS
System.Object

## NOTES
AUTHOR: Wojciech Sciesinski, wojciech[at]sciesinski[dot]net  
KEYWORDS: PowerShell, Exchange, SIPAddresses, ProxyAddresses, Lync, migration

VERSIONS HISTORY
- 0.1.0 - 2016-02-12 - First version published on GitHub
- 0.1.1 - 2016-02-12 - Help updated

LICENSE  
Copyright (c) 2016 Wojciech Sciesinski  
This function is licensed under The MIT License (MIT)  
Full license text: https://opensource.org/licenses/MIT

## EXAMPLES

### EXAMPLE 1
```powershell

    Check if the mailbox has assigned more than one SIP address - direct providing the mailbox identity

    [PS] > Get-MultiSIPMailbox -Identity AA473815

    MailboxAlias              : AA473815
    MailboxDisplayName        : Wolters-van der Thomes, IAV (Ingrid)
    PrimarySMTPAddress        : ingrid.wolters-van.der.thomes@example.nl
    MailboxGuid               : b201434a-1f62-4ee4-a446-e0b2bc7badc9
    SIPAddressesCount         : 2
    SIPAddressesList          : SIP:Ingrid.van.der.thomes-Wolters@example.com,sip:ingrid.wolters-van.der.thomes@example.com
    SIPAddresses              : {SIP:Ingrid.van.der.thomes-Wolters@example.com, sip:ingrid.wolters-van.der.thomes@example.com}
```    

### EXAMPLE 2
```powershell
Check if the mailbox has assigned more than one SIP address - providing the mailbox identity by pipeline

    [PS] > Get-Mailbox AA473815 | Get-MultiSIPMailbox

    MailboxAlias              : AA473815
    MailboxDisplayName        : Wolters-van der Thomes, IAV (Ingrid)
    PrimarySMTPAddress        : ingrid.wolters-van.der.thomes@example.nl
    MailboxGuid               : b201434a-1f62-4ee4-a446-e0b2bc7badc9
    SIPAddressesCount         : 2
    SIPAddressesList          : SIP:Ingrid.van.der.thomes-Wolters@example.com,sip:ingrid.wolters-van.der.thomes@example.com
    SIPAddresses              : {SIP:Ingrid.van.der.thomes-Wolters@example.com, sip:ingrid.wolters-van.der.thomes@example.com}
```
