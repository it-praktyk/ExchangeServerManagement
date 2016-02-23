function ConvertFrom-O365IPAddressesXMLFile {
    
    <#
    .SYNOPSIS
    Convert the O365IPAddresses.xml file to the custom PowerShell object
    
    .DESCRIPTION
    Function intended for converting to the custom PowerShell object the list of hosts used for Office 365 services published as the O365IPAddresses.xml file.
        
    The list contains addresses (IPv4, IPv6, URL) for what communication can't be proxied on customer/client side.
    
    More information on the Microsoft support page: "Office 365 URLs and IP address ranges", http://bit.ly/1LD8fYv
    
    .PARAMETER Path
    The xml file containing data like O365IPAddresses.xml downloaded manually. 
    If the the parameter is ommited the file O365IPAddresses.xml will be downloaded from the Microsoft site and saved with
        
    .OUTPUTS
    System.Object[]
  
    .EXAMPLE
    ConvertFrom-O365IPAddressesXMLFile
     
    .LINK
    https://github.com/it-praktyk/ConvertFrom-O365IPAddressesXMLFile
    
    .LINK
    https://www.linkedin.com/in/sciesinskiwojciech
          
    .NOTES
    AUTHOR: Wojciech Sciesinski, wojciech[at]sciesinski[dot]net
    KEYWORDS: PowerShell, Exchange, Office 365, XML, proxy
   
    VERSIONS HISTORY
    - 0.1.0 - 2016-02-23 - The first working version
	- 0.1.1 - 2016-02-24 - The parameter name in the helper function ConvertTo-Mask corrected

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
   
#>
    
    [cmdletbinding()]
    param (
        [Parameter(Mandatory = $false)]
        #[System.IO.File]
        $Path = ".\O365IPAddresses.xml"
    )
    
    BEGIN {
        
        If (!(Test-Path -Path $Path -Type Leaf)) {
            
            # If not provided file than download file from the internet
            
            [String]$OutputFileName = "O365IPAddresses-{0}.xml" -f (Get-Date -f "yyyyMMdd-HHmmss")
            
            Invoke-WebRequest -uri "https://support.content.office.net/en-us/static/O365IPAddresses.xml" -OutFile ".\$OutputFileName" -verbose
            
            [XML]$CurrentO365AddressesFile = Get-Content -Path ".\$OutputFileName"
            
        }
        Else {
            
            [XML]$CurrentO365AddressesFile = Get-Content -Path $Path
            
        }
        
        $Results = @()
        
    }
    
    PROCESS {
        
        $O365Services = $CurrentO365AddressesFile.products.product
        
        $O365ServicesCount = $(($O365Services | measure).Count)
        
        [String]$MessageText = "{0} products found in the xml file {1}" -f $O365ServicesCount, $Path
        
        Write-Verbose -Message $MessageText
        
        For ($i = 0; $i -lt $O365ServicesCount; $i++) {
            
            $CurrentService = $O365Services[$i]
            
            $CurrentServiceName = $CurrentService.name
            
            $CurrentAddressList = $O365Services[$i] | Select-Object -ExpandProperty addresslist
            
            $CurrentListCount = ($CurrentAddressList | measure).count
            
            [String]$MessageText = "Start processing for {0}, {1} addressess lists found" -f $CurrentServiceName, $CurrentListCount
            
            Write-Verbose -Message $MessageText
            
            For ($n = 0; $n -lt $CurrentListCount; $n++) {
                
                if ($CurrentListCount -eq 1) {
                    
                    $CurrentAddresses = $CurrentAddressList | Select -expandproperty address -ErrorAction SilentlyContinue
                    
                    $CurrentListType = $CurrentAddressList.type
                    
                }
                
                else {
                    
                    $CurrentAddresses = $CurrentAddressList[$n] | Select -expandproperty address -ErrorAction SilentlyContinue
                    
                    $CurrentListType = $CurrentAddressList[$n].type
                    
                }
                
                $CurrentAddressCount = $(($CurrentAddresses | measure).count)
                
                [String]$MessageText = "For the service {0} on the list {1} {2} addresses found." -f $CurrentServiceName, $CurrentListType, $CurrentAddressCount
                
                Write-Verbose -Message $MessageText
                
                For ($m = 0; $m -lt $CurrentAddressCount; $m++) {
                    
                    [String]$O365ServicesCurrentAddress = $CurrentAddresses[$m]
                    
                    [String]$MessageText = "Processing the address {0} from the list {1} for the service {2}" -f $O365ServicesCurrentAddress, $CurrentListType, $CurrentServiceName
                    
                    Write-Debug -Message $MessageText
                    
                    If ($O365ServicesCurrentAddress.contains("/")) {
                        
                        if (@("IPv4", "IPv6") -contains $CurrentListType) {
                            
                            [ipaddress]$IPAddress = $O365ServicesCurrentAddress.Split("/")[0]
                            
                            [int]$SubNetMaskLenght = $O365ServicesCurrentAddress.Split("/")[1]
                            
                            If ($CurrentListType -match "IPv4") {
                                
                                [String]$SubNetMask = ConvertTo-Mask -MaskLength $SubNetMaskLenght
                                
                            }
                            
                            [String]$Url = $null
                            
                        }
                    }
                    
                    elseif (-not ($O365ServicesCurrentAddress.contains("/")) -and $CurrentListType -eq "IPv4") {
                        
                        [ipaddress]$IPAddress = $O365ServicesCurrentAddress
                        
                        [int]$SubNetMaskLenght = 32
                        
                        [String]$SubNetMask = ConvertTo-Mask -MaskLength $SubNetMaskLenght
                        
                        [String]$Url = $null
                        
                    }
                    elseif (-not ($O365ServicesCurrentAddress.contains("/")) -and $CurrentListType -eq "IPv6") {
                        
                        [ipaddress]$IPAddress = $O365ServicesCurrentAddress
                        
                        [int]$SubNetMaskLenght = 128
                        
                        [String]$SubNetMask = $null
                        
                        [String]$Url = $null
                        
                    }
                    
                    else {
                        
                        [ipaddress]$IPAddress = $null
                        
                        [int]$SubNetMaskLenght = $null
                        
                        [String]$SubNetMask = $null
                        
                        [String]$Url = $O365ServicesCurrentAddress
                        
                    }
                    
                    $Result = New-Object -TypeName System.Management.Automation.PSObject
                    
                    $Result | Add-Member -MemberType NoteProperty -Name Service -Value $CurrentServiceName
                    
                    $Result | Add-Member -MemberType NoteProperty -Name Type -Value $CurrentListType
                    
                    $Result | Add-Member -MemberType NoteProperty -Name IPAddress -Value $IPAddress
                    
                    $Result | Add-Member -MemberType NoteProperty -Name SubNetMaskLenght -Value $SubNetMaskLenght
                    
                    $Result | Add-Member -MemberType NoteProperty -Name SubnetMask -value $SubNetMask
                    
                    $Result | Add-Member -MemberType NoteProperty -Name Url -Value $Url
                    
                    $Results += $Result
                    
                    #Remove previously set variables
                    $VariablesToClear = "SubnetMask", "SubnetMaskLenght", "IPAddress", "Url"
                    
                    $VariablesToClear | ForEach-Object -Process {
                        
                        Clear-Variable -Name $_ -ErrorAction SilentlyContinue
                        
                    }
                    
                }
                
            }
           
        }
        
        Remove-Variable CurrentListCount
        
    }
    
    END {
        
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
    https://github.com/it-praktyk/ConvertFrom-O365IPAddressesXMLFile
    
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
        
        Return $FullOctetsText.Substring(0,15)
        
    }
    Else {
        
        [Int]$MiddleBites = $($MaskLength - ($FullOctetsCounts * 8))
        
        switch ( $MiddleBites) {
            
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