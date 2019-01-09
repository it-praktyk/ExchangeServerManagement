function ConvertFrom-O365AddressesXMLFile {
    <#
    .SYNOPSIS
    Download and convert the O365IPAddresses.xml file to the custom PowerShell object.
    
    .DESCRIPTION
    Function intended for converting to the custom PowerShell object the list of hosts used for Office 365 services published as the O365IPAddresses.xml file.
        
    The list contains addresses (IPv4, IPv6, URL) for what communication can't be proxied on customer/client side.
    
    More information on the Microsoft support page: "Office 365 URLs and IP address ranges", http://bit.ly/1LD8fYv
    
    .PARAMETER Path
    The xml file containing data like O365IPAddresses.xml downloaded manually. 
    If the the parameter is omitted the file O365IPAddresses.xml will be downloaded from the Microsoft site and saved in current location with the name containing the date and time of download.
    
    .PARAMETER RemoveFileAfterParsing
    Remove file used to parsing after returning results.
    
    .PARAMETER DownloadXMLOnly
    Select if only O365IPAddresses.xml content need to be downloaded and stored to disk.
    
    .PARAMETER PassThru
    Returns an object representing the file containing O365IPAddresses-yyyyMMdd-HHmmss.xml file.
    
    .INPUTS
    None. The xml file published by Microsoft what contains the list of IP addresses ranges and names used by Office 365 services. 
    
    .OUTPUTS
    None. The custom PowerShell object what contains properties: Service,Type,IPAddress,SubNetMaskLength,SubnetMask,Url
  
    .EXAMPLE
    ConvertFrom-O365AddressesXMLFile
    
    Service          : o365
    Type             : IPv6
    IPAddress        : 2603:1030:800:5::bfee:a0ad
    SubNetMaskLength : 128
    SubnetMask       :
    Url              :
    
    <Output partially omitted>
    
    Service          : LYO
    Type             : IPv4
    IPAddress        : 23.103.129.128
    SubNetMaskLength : 25
    SubnetMask       : 255.255.255.128
    Url              :
    
    <Output partially omitted>
    
    Service          : ProPlus
    Type             : URL
    IPAddress        :
    SubNetMaskLength : 0
    SubnetMask       :
    Url              : go.microsoft.com
    
    <Output partially omitted>
    
    
    .EXAMPLE
    
    [PS] > ConvertFrom-O365AddressesXMLFile -Path .\O365IPAddresses.xml | get-member

    TypeName: System.Management.Automation.PSCustomObject

    Name             MemberType   Definition
    ----             ----------   ----------
    Equals           Method       bool Equals(System.Object obj)
    GetHashCode      Method       int GetHashCode()
    GetType          Method       type GetType()
    ToString         Method       string ToString()
    IPAddress        NoteProperty ipaddress IPAddress=2603:1030:800:5::bfee:a0ad
    Service          NoteProperty string Service=o365
    SubnetMask       NoteProperty string SubnetMask=
    SubNetMaskLenght NoteProperty int SubNetMaskLenght=128
    Type             NoteProperty string Type=IPv6
    Url              NoteProperty string Url=
    
     
    .LINK
    https://github.com/it-praktyk/Convert-Office365NetworksData
    
    .LINK
    https://www.linkedin.com/in/sciesinskiwojciech
          
    .NOTES
    AUTHOR: Wojciech Sciesinski, wojciech[at]sciesinski[dot]net
    KEYWORDS: PowerShell, Exchange, Office 365, O365, XML, proxy
   
    VERSIONS HISTORY
    - 0.1.0 - 2016-02-23 - The first working version
    - 0.1.1 - 2016-02-23 - The parameter name in the helper function ConvertTo-Mask corrected
    - 0.1.2 - 2016-02-23 - The output spelling corrected for SubNetMaskLength, help update, the function reformatted
    - 0.1.3 - 2016-02-23 - Small correction of code in an example
    - 0.1.4 - 2016-02-24 - Dates for versions 0.1.1 - 0.1.3 corrected, alliases for some cmdlets expanded to full names
    - 0.2.0 - 2016-06-17 - Support for handling download errors added, help updated, the main repository renamed
    - 0.3.0 - 2016-06-17 - The function name changed from ConvertFrom-O365IPAddressesXMLFile to ConvertFrom-O365AddressesXMLFile
	- 0.3.1 - 2016-06-18 - The script reformatted, TODO updated
	- 0.3.2 - 2016-06-19 - The help corrected
    - 0.4.0 - 2016-06-26 - The parameters DownloadRSSOnly,PassThru,RemoveFileAfterParsing added, the parameters sets added, TODO updated, help updated
    - 0.4.1 - 2016-06-26 - Information about required PowerShell version added, help updated
    - 0.5.0 - 2016-07-03 - The parameter DownloadRSSLOnly corrected to DownloadXMLOnly

    TODO
    - add only summary mode/switch - display info a last modification date, and sums IPs/URLs for products
    - check/correct verbose and debug mode
    - add support for downloading the file via proxy with authentication (?)
      #https://dscottraynsford.wordpress.com/2016/06/24/allow-powershell-to-traverse-a-secure-proxy/
    - add parameter to custom naming downloaded file (?)
      #https://github.com/it-praktyk/New-OutputObject
    - add support for PowerShell 2.0 - Invoke-WebRequest need to be replaced
    
        
    LICENSE
    Copyright (c) 2016 Wojciech Sciesinski
    This function is licensed under The MIT License (MIT)
    Full license text: https://opensource.org/licenses/MIT
   
#>
    
    #The cmdlet Invoke-WebRequest is used
    #Requires -Version 3.0
    
    [cmdletbinding(DefaultParameterSetName = 'Parse')]
    [outputtype(ParameterSetName = 'Parse', [System.Collections.ArrayList])]
    [outputtype(ParameterSetName = 'Download', [System.IO.FileInfo])]
    param (
        [Parameter(Mandatory = $false, ParameterSetName = 'Parse')]
        [String]$Path = ".\O365IPAddresses.xml",
        [Parameter(Mandatory = $false, ParameterSetName = 'Parse')]
        [Switch]$RemoveFileAfterParsing,
        [Parameter(Mandatory = $false, ParameterSetName = 'Download')]
        [Switch]$DownloadXMLOnly,
        [Parameter(Mandatory = $false, ParameterSetName = 'Download')]
        [switch]$PassThru
    )
    
    BEGIN {
        
        Try {
            
            If (!(Test-Path -Path $Path -Type Leaf)) {
                
                # If not provided file then download file from the internet
                
                [String]$UrlToDownload = "https://support.content.office.net/en-us/static/O365IPAddresses.xml"
                
                [String]$OutputFileName = "O365IPAddresses-{0}.xml" -f (Get-Date -f "yyyyMMdd-HHmmss")
                
                Invoke-WebRequest -uri $UrlToDownload -OutFile ".\$OutputFileName" | Out-Null
                
                [System.IO.FileInfo]$FileToProcess = Get-Item -Path $OutputFileName
                
                [String]$MessageText = "The data from {0} content downloaded and stored as a file {1}." -f $UrlToDownload, $FileToProcess.FullName
                
                Write-Verbose -Message $MessageText
                
            }
            Else {
                
                [System.IO.FileInfo]$FileToProcess = Get-Item -Path $Path
                
                [String]$MessageText = "The file {0} found and will be processed." -f $FileToProcess.FullName
                
                Write-Verbose -Message $MessageText
                
            }
            
            [XML]$CurrentO365AddressesFile = Get-Content -Path $FileToProcess.FullName
            
            $Results = @()
            
        }
        Catch {
            
            Throw $error[0]
            
        }
    }
    
    PROCESS {
        
        If ($DownloadXMLOnly.IsPresent) {
            
            If ($PassThru.IsPresent) {
                
                Return $FileToProcess
                
            }
            Else {
                
                break
                
            }
            
        }
        
        $O365Services = $CurrentO365AddressesFile.products.product
        
        $O365ServicesCount = $(($O365Services | Measure-Object).Count)
        
        [String]$MessageText = "{0} products found in the xml file {1}" -f $O365ServicesCount, $Path
        
        Write-Verbose -Message $MessageText
        
        For ($i = 0; $i -lt $O365ServicesCount; $i++) {
            
            $CurrentService = $O365Services[$i]
            
            $CurrentServiceName = $CurrentService.name
            
            $CurrentAddressList = $O365Services[$i] | Select-Object -ExpandProperty addresslist
            
            $CurrentListCount = ($CurrentAddressList | Measure-Object).count
            
            [String]$MessageText = "Start processing for {0}, {1} addressess lists found" -f $CurrentServiceName, $CurrentListCount
            
            Write-Verbose -Message $MessageText
            
            For ($n = 0; $n -lt $CurrentListCount; $n++) {
                
                if ($CurrentListCount -eq 1) {
                    
                    $CurrentAddresses = $CurrentAddressList | Select-Object -expandproperty address -ErrorAction SilentlyContinue
                    
                    $CurrentListType = $CurrentAddressList.type
                    
                }
                
                else {
                    
                    $CurrentAddresses = $CurrentAddressList[$n] | Select-object -expandproperty address -ErrorAction SilentlyContinue
                    
                    $CurrentListType = $CurrentAddressList[$n].type
                    
                }
                
                $CurrentAddressCount = $(($CurrentAddresses | Measure-Object).count)
                
                [String]$MessageText = "For the service {0} on the list {1} {2} addresses found." -f $CurrentServiceName, $CurrentListType, $CurrentAddressCount
                
                Write-Verbose -Message $MessageText
                
                For ($m = 0; $m -lt $CurrentAddressCount; $m++) {
                    
                    [String]$O365ServicesCurrentAddress = $CurrentAddresses[$m]
                    
                    [String]$MessageText = "Processing the address {0} from the list {1} for the service {2}" -f $O365ServicesCurrentAddress, $CurrentListType, $CurrentServiceName
                    
                    Write-Debug -Message $MessageText
                    
                    If ($O365ServicesCurrentAddress.contains("/")) {
                        
                        if (@("IPv4", "IPv6") -contains $CurrentListType) {
                            
                            [ipaddress]$IPAddress = $O365ServicesCurrentAddress.Split("/")[0]
                            
                            [int]$SubNetMaskLength = $O365ServicesCurrentAddress.Split("/")[1]
                            
                            If ($CurrentListType -match "IPv4") {
                                
                                [String]$SubNetMask = ConvertTo-Mask -MaskLength $SubNetMaskLength
                                
                            }
                            
                            [String]$Url = $null
                            
                        }
                        
                    }
                    
                    elseif (-not ($O365ServicesCurrentAddress.contains("/")) -and $CurrentListType -eq "IPv4") {
                        
                        [ipaddress]$IPAddress = $O365ServicesCurrentAddress
                        
                        [int]$SubNetMaskLength = 32
                        
                        [String]$SubNetMask = ConvertTo-Mask -MaskLength $SubNetMaskLength
                        
                        [String]$Url = $null
                        
                    }
                    elseif (-not ($O365ServicesCurrentAddress.contains("/")) -and $CurrentListType -eq "IPv6") {
                        
                        [ipaddress]$IPAddress = $O365ServicesCurrentAddress
                        
                        [int]$SubNetMaskLength = 128
                        
                        [String]$SubNetMask = $null
                        
                        [String]$Url = $null
                        
                    }
                    
                    else {
                        
                        [ipaddress]$IPAddress = $null
                        
                        [int]$SubNetMaskLength = $null
                        
                        [String]$SubNetMask = $null
                        
                        [String]$Url = $O365ServicesCurrentAddress
                        
                    }
                    
                    $Result = New-Object -TypeName System.Management.Automation.PSObject
                    
                    $Result | Add-Member -MemberType NoteProperty -Name Service -Value $CurrentServiceName
                    
                    $Result | Add-Member -MemberType NoteProperty -Name Type -Value $CurrentListType
                    
                    $Result | Add-Member -MemberType NoteProperty -Name IPAddress -Value $IPAddress
                    
                    $Result | Add-Member -MemberType NoteProperty -Name SubNetMaskLength -Value $SubNetMaskLength
                    
                    $Result | Add-Member -MemberType NoteProperty -Name SubnetMask -value $SubNetMask
                    
                    $Result | Add-Member -MemberType NoteProperty -Name Url -Value $Url
                    
                    $Results += $Result
                    
                    #Remove previously set variables
                    $VariablesToClear = "SubnetMask", "SubNetMaskLength", "IPAddress", "Url"
                    
                    $VariablesToClear | ForEach-Object -Process {
                        
                        Clear-Variable -Name $_ -ErrorAction SilentlyContinue
                        
                    }
                    
                }
                
            }
            
        }
        
        Remove-Variable CurrentListCount
        
    }
    
    END {
        
        
        If ($RemoveFileAfterParsing.IsPresent) {
            
            Remove-Item  -Path $FileToProcess -errorAction Continue
            
        }
        
        Return $Results
        
    }
    
}

function ConvertTo-Mask {
    
<#
    .SYNOPSIS
    Convert mask length to dotted binary format
    
    .DESCRIPTION
    The function what convert mask length to binary mask for IPv4 address
    
    .PARAMETER MaskLength
    The length of mask in IPv4 address
            
    .OUTPUTS
    The scring containging dotted mask for IPv4 addresses
  
    .EXAMPLE
    [PS] >ConvertTo-Mask -MaskLength 23
    
    255.255.254.0
     
    .LINK
    https://github.com/it-praktyk/Convert-Office365NetworksData
    
    .LINK
    https://www.linkedin.com/in/sciesinskiwojciech
          
    .NOTES
    AUTHOR: Wojciech Sciesinski, wojciech[at]sciesinski[dot]net
    KEYWORDS: PowerShell, network, network mask, IPv4
   
    VERSIONS HISTORY
    - 0.1.0 - 2016-02-23 - The first working version
    - 0.1.1 - 2016-02-24 - The parameter name corrected

    TODO
        
    LICENSE
    Copyright (c) 2016 Wojciech Sciesinski
    This function is licensed under The MIT License (MIT)
    Full license text: https://opensource.org/licenses/MIT
   
#>
    
    [cmdletbinding()]
    param (
        [Parameter(Mandatory = $true)]
        [ValidateRange(1, 32)]
        [Int]$MaskLength
        
    )
    
    [Int]$FullOctetsCounts = [math]::Truncate($MaskLength/8)
    
    [String]$FullOctetsText = "255." * $FullOctetsCounts
    
    If ($FullOctetsCounts -eq 4) {
        
        Return $FullOctetsText.Substring(0, 15)
        
    }
    Else {
        
        [Int]$MiddleBites = $($MaskLength - ($FullOctetsCounts * 8))
        
        switch ($MiddleBites) {
            
            0 { [String]$MiddleOctetString = "0" }
            1 { [String]$MiddleOctetString = "128" }
            2 { [String]$MiddleOctetString = "192" }
            3 { [String]$MiddleOctetString = "224" }
            4 { [String]$MiddleOctetString = "240" }
            5 { [String]$MiddleOctetString = "248" }
            6 { [String]$MiddleOctetString = "252" }
            7 { [String]$MiddleOctetString = "254" }
            
        }
        
        [Int]$LastOctetsCounts = 4 - ($FullOctetsCounts + 1)
        
        [String]$LastOctetsString = ".0" * $LastOctetsCounts
        
    }
    
    [String]$DottedMask = ("{0}{1}{2}" -f $FullOctetsText, $MiddleOctetString, $LastOctetsString).Replace("..", ".")
    
    Return $DottedMask
    
}