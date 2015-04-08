# ManageExchangeRecipients

Tool intended to manage mail addresses (proxy addresses) for mail objects in Exchange Server environment

##General input file Structure##
- RecipientIdentity
- RecipientType
- SkipRecipientCode - true/false or 1/0
- SkipRecipientReason - ?
- NewPrimaryProxyAddresss
- NewProxyAddressPrefix
- NewProxyAddress
- RemoveProxyAddressPrefix
- RemoveProxyAddress
- NewDisplayName - Active Directory or Exchange ?


##Modes##
* DisplayOnly
* PerformActions
* PerformActionsCommandsOnly
* Rollback
* RollbackCommandsOnly

##Currently supported Exchange related objects and operations##

###UserMailbox###
- AddProxyAddress - in modes: DisplayOnly, PerformActions
- RemoveProxyAddress - in modes: DisplayOnly, PerformActions
- SetSMTPPrimaryAddress - in modes: DisplayOnly, PerformActions

Main file: Manage-MailboxAddresses.ps1

###MailUniversalDistributionGroup###
- AddProxyAddress - in modes: DisplayOnly, PerformActions
- RemoveProxyAddress - in modes: DisplayOnly, PerformActions
- SetSMTPPrimaryAddress - in modes: DisplayOnly, PerformActions

Main file: Manage-DistributionGroupAddresses.ps1

###MailUniversalSecurityGroup,MailNonUniversalGroup###



##Output/Rollback file Structure##
- RecipientIdentity
- RecipientGUID
- RecipientAlias
- RecipientType - UserMailbox,
- RecipientWasSkipped - true/false or codes
- RecipientSkipReason - Reason description
- PrimarySMTPAddressBefore
- PrimarySMTPAddressAfter
- ProxyAddressesBefore
- ProxyAddressesProposal
- ProxyAddressesAfter

Output will be saved as csv file with semicolon (";") as as delimiter, encoded as UTF8

README.md file version history
0.1 - 2015-04-09 - initial release - mostly based on ideas.txt file
