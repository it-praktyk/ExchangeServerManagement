# ConvertFrom-O365AddressesRSS
## SYNOPSIS
Download and convert to custom PowerShell object the RSS channel data about planned changes to Office 365 networks/hosts.

## SYNTAX
```powershell

ConvertFrom-O365AddressesRSS [-Path <string>] [-Start <datetime>] [-End <datetime>] [-RemoveFileAfterParsing] [<CommonParameters>]

ConvertFrom-O365AddressesRSS [-DownloadRSSOnly] [-PassThru] [<CommonParameters>]

```

## DESCRIPTION
Function intended for downloading and converting to the custom PowerShell object the list of changes to Office 365 networks and hosts published by Microsoft as a RSS channel https://support.office.com/en-us/o365ip/rss.

More information on the Microsoft support page: "Office 365 URLs and IP address ranges", http://bit.ly/1LD8fYv .

## PARAMETERS
### -Path &lt;String&gt;
The xml file containing RSS data downloaded manually.
If the parameter is omitted the RSS data content will be downloaded from the Microsoft site and saved in current location with the name containing the date and time of download.
```
Required?                    false
Position?                    named
Default value                
Accept pipeline input?       false
Accept wildcard characters?  false
```

### -Start &lt;DateTime&gt;
```
Required?                    false
Position?                    named
Default value
Accept pipeline input?       false
Accept wildcard characters?  false
```

### -End &lt;DateTime&gt;
```
Required?                    false
Position?                    named
Default value
Accept pipeline input?       false
Accept wildcard characters?  false
```

### -RemoveFileAfterParsing &lt;SwitchParameter&gt;
Remove file used to parsing after all operations.
```
Required?                    false
Position?                    named
Default value                False
Accept pipeline input?       false
Accept wildcard characters?  false
```

### -DownloadRSSOnly &lt;SwitchParameter&gt;
Select if only RSS content need to be downloaded and stored to disk.
```
Required?                    false
Position?                    named
Default value                False
Accept pipeline input?       false
Accept wildcard characters?  false
```

### -PassThru &lt;SwitchParameter&gt;
Returns an object representing the file containing RSS content data.
```
Required?                    false
Position?                    named
Default value                False
Accept pipeline input?       false
Accept wildcard characters?  false
```

## INPUTS
None. The xml data published as RSS channel under url https://support.office.com/en-us/o365ip/rss.

## OUTPUTS
None. The custom PowerShell object what contains properties: OperationType, Title, PublicationDate, Guid, Description, DescriptionIsParsable, QuickDescription, Notes, SubChanges.  
The Subchanges property is array of objects (so can be expanded) to object what contains properties: EffectiveDate, Required, ExpressRoute, Value.  
If the parameter DownloadRSSOnly is used the file containing downloaded RSS data is returned.

## NOTES
AUTHOR: Wojciech Sciesinski, wojciech[at]sciesinski[dot]net  
KEYWORDS: PowerShell, Exchange, Office 365, O365, XML, proxy, RSS  

VERSIONS HISTORY  
- 0.1.0 - 2016-06-17 - The first version published to GitHub
- 0.1.1 - 2016-06-19 - A case when the parameter Path is used corrected, TODO updated
- 0.1.2 - 2016-06-19 - Handling input file rewrote partially, help updated
- 0.2.0 - 2016-06-21 - Support for Protocol,Port,Status means:Required/Optional added in SubChanges, help updated
- 0.2.1 - 2016-06-21 - Parsing description to SubChanges corrected
- 0.2.2 - 2016-06-21 - Parsing 'Updating' items added
- 0.2.3 - 2016-06-21 - Description will be trimmed at the begining of processing, TODO updated
- 0.3.0 - 2016-06-23 - Workarounds for inconsistent descriptions added, the parameters Start, End added to limit parse between dates
- 0.4.0 - 2016-06-24 - Parsing notes only RSS items added, verbose corrected
- 0.4.1 - 2016-06-24 - Workarounds for inconsistent descriptions corrected, TODO updated
- 0.5.0 - 2016-06-26 - Output for non parsable items changed, now is more descriptive
- 0.5.1 - 2016-06-26 - Corrected output for subchanges
- 0.6.0 - 2016-06-26 - The parameters DownloadRSSOnly,PassThru,RemoveFileAfterParsing added, the parameters set added, TODO updated, help updated
- 0.6.1 - 2016-06-26 - The default value for the parameter Path removed, help corrected
- 0.6.2 - 2016-06-02 - Help updated, code cleaned


TODO
- add support for downloading the file via proxy with authentication (?)
  #https://dscottraynsford.wordpress.com/2016/06/24/allow-powershell-to-traverse-a-secure-proxy/
- add parameter to custom naming downloaded file  
  #https://github.com/it-praktyk/New-OutputObject
- implement downloadable overwrites for non-parsable RSS items (?)
- add support for PowerShell 2.0 - Invoke-WebRequest need to be replaced

LICENSE  
Copyright (c) 2016 Wojciech Sciesinski  
This function is licensed under The MIT License (MIT)  
Full license text: https://opensource.org/licenses/MIT  

## EXAMPLES
### EXAMPLE 1
```powershell
[PS] >ConvertFrom-O365AddressesRSS

	<Output partially omitted>

	OperationType         : Adding
	Title                 : Exchange Online
	PublicationDate       : 6/13/2016 3:06:43 PM
	Guid                  : 7ef9205d-fb30-43bf-9501-9fe8106dfa20
	Description           : Adding 1 New FQDNs; 1/[Effective 8/1/2016. Required: Exchange Online Protection. ExpressRoute:
	                        No. 40mshrcstorageprod.blob.core.windows.net]. Notes: removing the wildcard for this endpoint.
	DescriptionIsParsable : True
	QuickDescription      : Adding 1 New FQDNs
	Notes                 : removing the wildcard for this endpoint.
	SubChanges            : {@{EffectiveDate=8/1/2016 12:00:00 AM; Required=Exchange Online Protection;
	                        ExpressRoute=False; Value=40mshrcstorageprod.blob.core.windows.net}}

	OperationType         : Adding
	Title                 : Office Online
	PublicationDate       : 6/13/2016 3:06:45 PM
	Guid                  : 8ef9105d-fb30-43bf-9502-9fe7106efa20
	Description           : Adding 1 New IP_Sets; 1/[Effective 6/13/2016. Required: Office Online. ExpressRoute: Yes.
	                        13.94.209.165]. Notes: Infrastructure change for a small component of Office Online, minimal
	                        (if any) customer impact; additionally, this endpoint won˘t be available via ExpressRoute until 8/1/2016.
	DescriptionIsParsable : True
	QuickDescription      : Adding 1 New IP_Sets
	Notes                 :
	SubChanges            : {@{EffectiveDate=6/13/2016 12:00:00 AM; Required=Office Online; ExpressRoute=True;
	                        Value=13.94.209.165}}

	<Output partially omitted>

	Automatically parsed RSS items, general view without expanding SubChanges.
```


### EXAMPLE 2
```powershell
[PS] >ConvertFrom-O365AddressesRSS -Path .\O365AddressesRSS.xml | Get-Member

	TypeName: Selected.System.String

	Name                  MemberType   Definition
	----                  ----------   ----------
	Equals                Method       bool Equals(System.Object obj)
	GetHashCode           Method       int GetHashCode()
	GetType               Method       type GetType()
	ToString              Method       string ToString()
	Description           NoteProperty string Description=If 134.170.0.0/16 has already been added from the Office 365 l...
	DescriptionIsParsable NoteProperty bool DescriptionIsParsable=False
	Guid                  NoteProperty string Guid=fa204cfc-402a-4fe6-818e-f9105dfb303b
	Notes                 NoteProperty object Notes=null
	OperationType         NoteProperty object OperationType=null
	PublicationDate       NoteProperty datetime PublicationDate=7/15/2014 7:00:00 AM
	QuickDescription      NoteProperty object QuickDescription=null
	SubChanges            NoteProperty object SubChanges=null
	Title                 NoteProperty string Title=Office Online


    Output for the RSS item what is not parsable
```

### EXAMPLE 3

```powershell
[PS] >ConvertFrom-O365AddressesRSS -Path .\O365AddressesRSS.xml | get-member

       TypeName: Selected.System.String

    Name                  MemberType   Definition
    ----                  ----------   ----------
    Equals                Method       bool Equals(System.Object obj)
    GetHashCode           Method       int GetHashCode()
    GetType               Method       type GetType()
    ToString              Method       string ToString()
    Description           NoteProperty string Description=Adding 1 New IP_Sets; 1/[Effective 7/1/2016. Required: Exchang...
    DescriptionIsParsable NoteProperty bool DescriptionIsParsable=True
    Guid                  NoteProperty string Guid=029fe710-7ef9-4205-8fb4-03afd6018ef8
    Notes                 NoteProperty string Notes=adding consolidated range.
    OperationType         NoteProperty string OperationType=Adding
    PublicationDate       NoteProperty datetime PublicationDate=6/1/2016 12:22:56 PM
    QuickDescription      NoteProperty string QuickDescription=Adding 1 New IP_Sets
    SubChanges            NoteProperty System.Collections.ArrayList SubChanges=
    Title                 NoteProperty string Title=Exchange Online Protection


    Output for the RSS item what is parsable

```

### EXAMPLE 4
```powershell
[PS] >ConvertFrom-O365AddressesRSS -Start 5/31/2016 -End 6/2/2016 | Select-Object -Property PublicationDate,Description,Guid -ExpandProperty Subchanges

    <Output partially omitted>

    EffectiveDate   : 7/1/2016 12:00:00 AM
    Status          : Required
    SubService      : Yammer
    ExpressRoute    : False
    Protocol        :
    Port            :
    Value           : 13.107.6.158/31
    PublicationDate : 6/1/2016 12:22:50 PM
    Description     : Adding 2 New IP_Sets; 1/[Effective 7/1/2016. Required: Yammer. ExpressRoute: No. 13.107.6.158/31],
                      2/[Effective 7/1/2016. Required: Yammer. ExpressRoute: No. 13.107.9.158/31]. Notes: adding
                      consolidated range.
    Guid            : 17ef9105-dfb3-403b-bd50-19fe8106efa2

    EffectiveDate   : 7/1/2016 12:00:00 AM
    Status          : Required
    SubService      : Yammer
    ExpressRoute    : False
    Protocol        :
    Port            :
    Value           : 13.107.9.158/31
    PublicationDate : 6/1/2016 12:22:50 PM
    Description     : Adding 2 New IP_Sets; 1/[Effective 7/1/2016. Required: Yammer. ExpressRoute: No. 13.107.6.158/31],
                      2/[Effective 7/1/2016. Required: Yammer. ExpressRoute: No. 13.107.9.158/31]. Notes: adding
                      consolidated range.
    Guid            : 17ef9105-dfb3-403b-bd50-19fe8106efa2

    EffectiveDate   : 7/1/2016 12:00:00 AM
    Status          : Required
    SubService      : Exchange Online Protection
    ExpressRoute    : True
    Protocol        :
    Port            :
    Value           : 157.56.111.0/24
    PublicationDate : 6/1/2016 12:22:53 PM
    Description     : Removing 1 Old IP_Sets; 1/[Effective 7/1/2016. Required: Exchange Online Protection. ExpressRoute:
                      YES. 157.56.111.0/24]. Notes: removing range to support consolidation.
    Guid            : 106dfa30-4bfc-4502-9fe7-017ef9205cfb

    EffectiveDate   : 7/1/2016 12:00:00 AM
    Status          : Required
    SubService      : Exchange Online Protection
    ExpressRoute    : True
    Protocol        :
    Port            :
    Value           : 216.32.180.0/23
    PublicationDate : 6/1/2016 12:22:56 PM
    Description     : Adding 1 New IP_Sets; 1/[Effective 7/1/2016. Required: Exchange Online Protection. ExpressRoute:
                      YES. 216.32.180.0/23]. Notes: adding consolidated range.
    Guid            : 029fe710-7ef9-4205-8fb4-03afd6018ef8


    Automatically parsed RSS items with details about planned subchanges.

```
