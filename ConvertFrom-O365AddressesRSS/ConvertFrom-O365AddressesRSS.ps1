function ConvertFrom-O365AddressesRSS {
    <#
    .SYNOPSIS
    Download and convert to custom PowerShell object the RSS channel data about planned changes to Office 365 networks/hosts.
    
    .DESCRIPTION
    Function intended for converting to the custom PowerShell object the list of changes published by Microsoft as RSS items.
    
    More information on the Microsoft support page: "Office 365 URLs and IP address ranges", http://bit.ly/1LD8fYv
    
    .PARAMETER Path
    The xml file containing data like O365IPAddresses.xml downloaded manually. 
    If the the parameter is ommited the file O365IPAddresses.xml will be downloaded from the Microsoft site and saved with
    
    .INPUTS
    None. The xml data published as RSS channel under url https://support.office.com/en-us/o365ip/rss. 
    
    .OUTPUTS
    None. The custom PowerShell object what contains properties: OperationType, Title, PublicationDate, Guid, Description, DescriptionIsParsable, QuickDescription, Notes, SubChanges. The Subchanges property is array of objects (so can be expanded) to object what contains properties:  EffectiveDate, Required, ExpressRoute, Value.
    
    .EXAMPLE
    [PS] > ConvertFrom-O365AddressesRSS
    
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
                            (if any) customer impact; additionally, this endpoint wonĂ˘â'¬â"˘t be available via
                            ExpressRoute until 8/1/2016.
    DescriptionIsParsable : True
    QuickDescription      : Adding 1 New IP_Sets
    Notes                 :
    SubChanges            : {@{EffectiveDate=6/13/2016 12:00:00 AM; Required=Office Online; ExpressRoute=True;
                            Value=13.94.209.165}}
    
    Automatically parsed RSS items, general view without expanding SubChanges
    
    
    .EXAMPLE
    
    [PS] > ConvertFrom-O365AddressesRSS -Path .\O365AddressesRSS.xml | get-member

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

    .EXAMPLE
    
    Output for the RSS item what is not parsable
    
    [PS] > ConvertFrom-O365AddressesRSS -Path .\O365AddressesRSS.xml | get-member

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
    
    .EXAMPLE
    
    [PS] > ConvertFrom-O365AddressesRSS | Select-Object -Property Guid,OperationType,PublicationDate,Title -ExpandProperty SubChanges
    
    <Output partially omitted>
    
    EffectiveDate   : 7/1/2016 12:00:00 AM
    Required        : Exchange Online Protection
    ExpressRoute    : True
    Value           : 216.32.180.0/23
    Guid            : 029fe710-7ef9-4205-8fb4-03afd6018ef8
    OperationType   : Adding
    PublicationDate : 6/1/2016 12:22:56 PM
    Title           : Exchange Online Protection

    EffectiveDate   : 8/1/2016 12:00:00 AM
    Required        : Office 365 Authentication and identity
    ExpressRoute    : True
    Value           : 2a01:111:2005:6::/64
    Guid            : 29fe7107-ef92-404c-bc40-3afd6018ef81
    OperationType   : Adding
    PublicationDate : 6/13/2016 3:06:37 PM
    Title           : Authentication and Identity

    EffectiveDate   : 8/1/2016 12:00:00 AM
    Required        : Exchange Online Protection
    ExpressRoute    : True
    Value           : 207.46.101.128/26
    Guid            : ef9205cf-b403-4afd-a018-fe8106dfa304
    OperationType   : Removing
    PublicationDate : 6/13/2016 3:06:39 PM
    Title           : Exchange Online Protection
    
    <Output partially omitted>
    
    Automatically parsed RSS items with details about planned changes.
     
    .LINK
    https://github.com/it-praktyk/Convert-Office365NetworksData
    
    .LINK
    https://www.linkedin.com/in/sciesinskiwojciech
          
    .NOTES
    AUTHOR: Wojciech Sciesinski, wojciech[at]sciesinski[dot]net
    KEYWORDS: PowerShell, Exchange, Office 365, O365, XML, proxy, RSS
   
    VERSIONS HISTORY
    - 0.1.0 - 2016-06-17 - The first version published to GitHub
    
    TODO
	- implement parameters DownloadRSSOnly, CleanFileAfterParsing
    - add suport to return/parse RSS items between selected dates only
    - add support for downloading the file via proxy with authentication (?)
    - add parameter to custom naming downloaded file
    
        
    LICENSE
    Copyright (c) 2016 Wojciech Sciesinski
    This function is licensed under The MIT License (MIT)
    Full license text: https://opensource.org/licenses/MIT
   
#>
	
	[cmdletbinding()]
	param (
		[Parameter(Mandatory = $false)]
		#[System.IO.File]

		$Path = ".\O365AddressesRSS.xml"#,
		#[Parameter(Mandatory = $false)]
		#[Switch]$DownloadRSSOnly,
		#[Parameter(Mandatory = $false)]
		#[Switch]$CleanFileAfterParsing 
	)
	
	BEGIN {
		
		[Bool]$InternalVerbose = $false
		
		[Bool]$ParameterVerbose = ($PSCmdlet.MyInvocation.BoundParameters["Verbose"].IsPresent)
		
		Try {
			
			If (!(Test-Path -Path $Path -PathType Leaf)) {
				
				# If not provided file then download file from the internet
				
				[String]$UrlToDownload = "https://support.office.com/en-us/o365ip/rss"
				
				[String]$OutputFileName = "O365IPAddressesChanges-{0}.xml" -f (Get-Date -Format "yyyyMMdd-HHmmss")
				
				Invoke-WebRequest -uri $UrlToDownload -OutFile ".\$OutputFileName" -Verbose:$InternalVerbose
				
				[XML]$CurrentO365AddressesChanges = Get-Content -Path ".\$OutputFileName"
				
				[String]$MessageText = "The RSS {0} content downloaded and stored as a file {1}" -f $UrlToDownload, $OutputFileName
				
				Write-Verbose -Message $MessageText
				
			}
			Else {
				
				$OutputFileName = $Path
				
				[XML]$CurrentO365AddressesChangesFile = Get-Content -Path $Path
				
			}
			
			$Results = New-Object System.Collections.ArrayList
			
		}
		Catch {
			
			Throw $error[0]
			
		}
		
	}
	
	PROCESS {
		
		$RSSItem = $CurrentO365AddressesChanges.rss.channel
		
		$RSSItemCount = $(($RSSItem.Item | Measure-Object).Count)
		
		[String]$MessageText = "{0} RSS items found in the xml file {1}" -f $RSSItemCount, $OutputFileName
		
		Write-Verbose -Message $MessageText
		
		$i = 1
		
		ForEach ($CurrentItem in $RSSItem.Item) {
			
			if ($i -gt 197 -and $i -lt 203) {
				
				#Prepoulating properties for the Result object
				
				$Result = "" | Select-Object -Property OperationType, Title, PublicationDate, Guid, Description, DescriptionIsParsable, QuickDescription, Notes, SubChanges
				
				$CurrentItemGuid = $CurrentItem.guid
				
				$DescriptionParsable = $false
				
				$CurrentItemDescription = $($CurrentItem.Description).Replace("$([char][int]10)", " ")
				
				$ParsedDescription = Parse-O365IPAddressChangesDescription -Description $CurrentItemDescription -Guid $CurrentItemGuid -Verbose:$ParameterVerbose
				
				[datetime]$CurrentItemPubDate = $CurrentItem.pubDate
				
				$CurrentItemTitle = $($($CurrentItem.Title).trim()).Replace("$([char][int]10)", " ")
				
				$Result.Title = $CurrentItemTitle
				
				$Result.PublicationDate = $CurrentItemPubDate
				
				$Result.Guid = $CurrentItemGuid
				
				$Result.Description = $CurrentItemDescription
			
			If ($ParsedDescription.DescriptionIsParsable) {
				
				$Result.OperationType = $ParsedDescription.OperationType
				
				$Result.DescriptionIsParsable = $ParsedDescription.DescriptionIsParsable
				
				$Result.QuickDescription = $ParsedDescription.QuickChangeDescription
				
				$Result.Notes = $ParsedDescription.Notes
				
				$Result.SubChanges = $ParsedDescription.SubChanges
				
			}
			Else {
				
				$Result.DescriptionIsParsable = $ParsedDescription.DescriptionIsParsable
				
			}
			
			
				$Results.Add($Result) | Out-Null
				
			}
			
			$i++
			
		}
		
	}
	
	END {
		
		Return $Results
		
	}
	
}

Function Parse-O365IPAddressChangesDescription {
	
	[cmdletbinding()]
	param (
		
		[parameter(Mandatory = $true)]
		[String]$Description,
		[parameter(Mandatory = $true)]
		[String]$Guid
		
	)
	
	begin {
		
		$DescriptionIsParsable = $false
		
	}
	
	Process {
		
		Try {
			
			$DescriptionSplittedParts = $Description.Split(';')
			
			$DescriptionSplittedPartsCount = ($DescriptionSplittedParts | Measure-Object).Count
			
			#Add something to catch the semicolons in a Notes part like in 8ef9105d-fb30-43bf-9502-9fe7106efa20
			
			#Replace end of the line chars
			$QuickDescription = $($DescriptionSplittedParts[0]).Replace("$([char][int]10)", " ")
			
			$QuickDescriptionPart0 = $($QuickDescription.Split(' '))[0]
			
			[String]$MessageText = "QuickDescription separated from the Description field: {0}" -f $QuickDescription
			
			Write-Verbose -Message $MessageText
			
			
			If (@("Adding", "Removing") -contains $QuickDescriptionPart0) {
				
				
				[String]$MessageText = "Recognized operations in the RSS item {0} is {1} - means Adding or Removing. Description will be parsed to extract SubChanges." -f $Guid, $QuickDescriptionPart0
				
				Write-Verbose -Message $MessageText
				
				$Operations = $($($($($Description.Split(';'))[1]).trim()).Replace("$([char][int]10)", " ")).Split(',')
				
				$OperationsCount = ($Operations | Measure-Object).Count
				
				For ($i = 0; $i -lt $OperationsCount; $i++) {
					
					$CurrentOperation = $($Operations[$i]).Trim()
					
					[String]$MessageText = "Parsing subchange: {0}" -f $CurrentOperation
					
					Write-Verbose -Message $MessageText
					
					If ($i -eq $OperationsCount - 1 -and $CurrentOperation -match 'Notes:') {
						
						#Try find Notes
						
						$RawNotes = $CurrentOperation.Split(']')[1]
						
						$Notes = $RawNotes.Substring(9, $RawNotes.length - 9)
						
					}
					
					$OpenBracket = $CurrentOperation.IndexOf('[')
					
					$CloseBracket = $CurrentOperation.IndexOf(']')
					
					If ($OpenBracket -eq -1 -or $CloseBracket -eq -1) {
						
						Break
						
					}
					Else {
						
						$SubResults = New-Object System.Collections.ArrayList
						
						#Clean data from data outside brackets
						$CurrentOperation = $CurrentOperation.Substring($OpenBracket + 1, ($CloseBracket - $OpenBracket) - 1)
						
						#Split data to fields
						$CurrentOperationSplited = $CurrentOperation.Split('.')
						
						[DateTime]$EffectiveDate = Get-Date -Date $($($CurrentOperationSplited[0]).Trim()).Replace('Effective ', '') -Format 'M/d/yyyy'
						
						$Required = $($($CurrentOperationSplited[1]).Trim()).Replace('Required: ', '')
						
						If ($(($CurrentOperationSplited[2]).Trim()).Replace('ExpressRoute: ', '') -eq 'Yes') {
							
							$ExpressRoute = $true
							
						}
						Else {
							
							$ExpressRoute = $false
							
						}
						
						$LastSpaceIndex = $CurrentOperation.LastIndexOf(' ')
						
						$Value = $($CurrentOperation.Substring($LastSpaceIndex + 1, $($CurrentOperation.length - $LastSpaceIndex) - 1)).Trim()
						
						$SubResult = New-Object -TypeName System.Management.Automation.PSObject
						
						$SubResult | Add-Member -MemberType NoteProperty -Name EffectiveDate -Value $EffectiveDate
						
						$SubResult | Add-Member -MemberType NoteProperty -Name Required -Value $Required
						
						$SubResult | Add-Member -MemberType NoteProperty -Name ExpressRoute -Value $ExpressRoute
						
						$SubResult | Add-Member -MemberType NoteProperty -Name Value -Value $Value
						
						$DescriptionIsParsable = $true
						
						$SubResults.Add($SubResult) | Out-Null
						
						If ($DescriptionIsParsable) {
							
							[String]$MessageText = "Subchange {0} from RSS item {1} parsed successfully to: {2}" -f $i, $Guid, $SubResult
							
						}
						Else {
							
							[String]$MessageText = "Subchange {0} from RSS item {1} not parsed successfully" -f $i, $Guid
							
						}
						
						Write-Verbose -Message $MessageText
						
						Remove-Variable -Name SubResult | Out-Null
						
					}
					
				}
				
			}
			
		}
		Catch {
			
			$DescriptionIsParsable = $false
			
		}
		
		Finally {
			
			If ($DescriptionIsParsable) {
				
				[String]$MessageText = "All subchanges from RSS item {0} have parsed successfully." -f $Guid
				
			}
			Else {
				
				[String]$MessageText = "Subchange {0} from RSS item {1} not parsed successfully." -f $i, $Guid
				
			}
			
			Write-Verbose -Message $MessageText
			
			
		}
		
	}
	
	
	End {
		
		
		$Result = New-Object -TypeName System.Management.Automation.PSObject
		
		$Result | Add-Member -MemberType NoteProperty -Name DescriptionIsParsable -Value $DescriptionIsParsable
		
		$Result | Add-Member -MemberType NoteProperty -Name QuickChangeDescription -Value $QuickDescription
		
		$Result | Add-Member -MemberType NoteProperty -Name OperationType -Value $QuickDescriptionPart0
		
		$Result | Add-Member -MemberType NoteProperty -Name Notes -Value $Notes
		
		$Result | Add-Member -MemberType NoteProperty -Name SubChanges -Value $SubResults
		
		Return $Result
		
	}
	
}