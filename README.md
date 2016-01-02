# Test-EmailAddress
## SYNOPSIS
The function Test-EmailAddress is intended to verify the correctness of email addresses in Microsoft Exchange Server environment

## SYNTAX
```powershell
Test-EmailAddress [-EmailAddress] <String[]> [[-TestEmailFormat] <Boolean>] [[-TestAcceptedDomains] <Boolean>] [[-TestIfExists] <Boolean>] [[-TestIsPrimary] <Boolean>] [[-AcceptedDomains] <String[]>] [<CommonParameters>]
```

## DESCRIPTION
Function which can be used to verifing email addresses in Microsoft Exchange Server environment.

Checks perfomed:
a) if an email address provided as parameter value contains wrong characters e.g. spaces at the begining/end
b) if an email address format is complaint with requirements - check Wikipedia https://en.wikipedia.org/wiki/Email_address
c) if an email address is from a domain which are added to the accepted domains list of is in the list passed as parameter value
d) if an email address is currently assigned to any object in Exchange environment (a conflicted object exist)
e) if an email address is currently set as PrimarySMTPAddress for existing object

## PARAMETERS
### -EmailAddress &lt;String[]&gt;
Email address which need to be verified in Exchange environment
```
Required?                    true
Position?                    1
Default value
Accept pipeline input?       true (ByValue)
Accept wildcard characters?  false
```

### -TestEmailFormat &lt;Boolean&gt;
Set to false to skip testing email address format
```
Required?                    false
Position?                    2
Default value                True
Accept pipeline input?       false
Accept wildcard characters?  false
```

### -TestAcceptedDomains &lt;Boolean&gt;
Set to false to skip testing if domain of an email address is in accepted domain list
```
Required?                    false
Position?                    3
Default value                True
Accept pipeline input?       false
Accept wildcard characters?  false
```

### -TestIfExists &lt;Boolean&gt;
Set to false to skip testing if an email address exist in mail organization
```
Required?                    false
Position?                    4
Default value                True
Accept pipeline input?       false
Accept wildcard characters?  false
```

### -TestIsPrimary &lt;Boolean&gt;
Set to false to skip testing if email is primary for existing object
```
Required?                    false
Position?                    5
Default value                True
Accept pipeline input?       false
Accept wildcard characters?  false
```

### -AcceptedDomains &lt;String[]&gt;
The list of domains used to testing if email is from accepted domains.
```
Required?                    false
Position?                    6
Default value
Accept pipeline input?       false
Accept wildcard characters?  false
```

## NOTES
AUTHOR: Wojciech Sciesinski, wojciech[at]sciesinski[dot]net

KEYWORDS: Windows, PowerShell, Exchange Server, email

VERSION HISTORY
0.6.0 - 2015-12-22 - the function rewriten, information about license added
0.7.0 - 2015-12-29 - validation extended and corrected
0.8.0 - 2015-12-31 - the function tested, the parameter $AcceptedDomains implemented, help updated

TODO
- add a description of errors for the test TestEmailFormat results (?)
- add an additional parameter AcceptOnlyEnglishLetters
- add an additional parameter AllowedCharsExclusionList

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

## EXAMPLES
### EXAMPLE 1
```powershell
[PS] >Test-EmailAddress -EmailAddress dummy@example.com

EmailAddress              : dummy@example.com
EmailDomain               : example.com
TestWhiteChars            : PASS
TestEmailFormat           : PASS
TestAcceptedDomain        : PASS
TestEmailExists           : EXISTS
ExistingObjectAlias       : dummy
ExisitngObjectGuid        : 181ca5f1-2fc0-40ef-853d-215a2b1fd16d
ExistingObjectType        : UserMailbox
IsPrimaryAddress          : True
EmailAddressPolicyEnabled : True
```


### EXAMPLE 2
```powershell
[PS] >"postmaster@gelatto.test","new@new.pl","new@test@new.pl" | Test-EmailAddress -AcceptedDomains new.pl

WARNING: Only domains passed to the AcceptedDomains parameter will be evaluated under the TestAcceptedDomains test

EmailAddress              : postmaster@gelatto.test
EmailDomain               : gelatto.test
TestWhiteChars            : PASS
TestEmailFormat           : PASS
TestAcceptedDomain        : FAIL
TestEmailExists           : EXISTS
ExistingObjectAlias       : DL_mailadmin
ExisitngObjectGuid        : d897b45f-c104-4d3b-b77e-6f4332565f8451
ExistingObjectType        : MailUniversalSecurityGroup
IsPrimaryAddress          : False
EmailAddressPolicyEnabled : False

EmailAddress              : new@new.pl
EmailDomain               : new.pl
TestWhiteChars            : PASS
TestEmailFormat           : PASS
TestAcceptedDomain        : PASS
TestEmailExists           : NO EXISTS
ExistingObjectAlias       : NO EXISTS
ExisitngObjectGuid        : NO EXISTS
ExistingObjectType        : NO EXISTS
IsPrimaryAddress          : NO EXISTS
EmailAddressPolicyEnabled : NO EXISTS

EmailAddress              : new@test@new.pl
EmailDomain               : new.pl
TestWhiteChars            : PASS
TestEmailFormat           : FAIL
TestAcceptedDomain        : SKIPPED
TestEmailExists           : SKIPPED
ExistingObjectAlias       : SKIPPED
ExisitngObjectGuid        : SKIPPED
ExistingObjectType        : SKIPPED
IsPrimaryAddress          : SKIPPED
EmailAddressPolicyEnabled : SKIPPED
```


### EXAMPLE 3
```powershell
[PS] >Get-AcceptedDomain

Name                           DomainName                     DomainType                   Default
----                           ----------                     ----------                   -------
gto.local                      gto.local                      Authoritative                False
gelatto.test                   gelatto.test                   InternalRelay                True
example.com                    example.com                    Authoritative                False

[PS] > Get-Mailbox dummy | Select-Object -ExpandProperty emailaddresses | Where-Object -FilterScript { $_.prefix -match 'smtp' } | ForEach { Test-EmailAddress $_.SMTPAddress }

EmailAddress              : dummy@example.com
EmailDomain               : example.com
TestWhiteChars            : PASS
TestEmailFormat           : PASS
TestAcceptedDomain        : PASS
TestEmailExists           : EXISTS
ExistingObjectAlias       : dummy
ExisitngObjectGuid        : 181ca5f1-2fc0-40ef-853d-215a2b1fd16d
ExistingObjectType        : UserMailbox
IsPrimaryAddress          : True
EmailAddressPolicyEnabled : True

EmailAddress              : dummy.user@gelatto.test
EmailDomain               : gelatto.test
TestWhiteChars            : PASS
TestEmailFormat           : PASS
TestAcceptedDomain        : PASS
TestEmailExists           : EXISTS
ExistingObjectAlias       : dummy
ExisitngObjectGuid        : 181ca5f1-2fc0-40ef-853d-215a2b1fd16d
ExistingObjectType        : UserMailbox
IsPrimaryAddress          : False
EmailAddressPolicyEnabled : True
```
