# Test-EmailAddress
PowerShell function intended to verifing the correctness of addresses email in Microsoft Exchange enviroment

Function which can be used to veryfing an email address before for example adding it to Microsoft Exchange environment as next proxy address. 

Checks perfomed: 
a) if an email address contain wrong characters e.g. % or spaces
b) if an email address is from domain which are on accepted domains list
c) if an email address is currently assigned to any object in Exchange environment

As a result function returns true/false if a provided email address is correct/incorrect.
