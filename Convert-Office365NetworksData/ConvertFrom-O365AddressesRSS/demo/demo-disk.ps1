$version = "0.6.0"

$xmlFileName = "O365IPAddressesChanges-20160626-030036.xml"

Import-Module .\ConvertFrom-O365AddressesRSS.ps1 -Force -ErrorAction Stop

If ( -not (Test-Path ".\demo\$version" -type container)) {

	New-item -Path ".\demo\$version" -Type Directory | Out-Null

}



[String]$xmlFilePath = ".\demo\{0}\{1}" -f $version, $xmlFileName

If (-not (Test-Path ".\demo\$version\$xmlFileName") -and (Test-Path -Path .\$xmlFileName) ) {

	Move-Item -Path .\$xmlFileName -Destination ".\demo\$version\$xmlFileName"

}

$generalOutputFileName = "{0}-Output.csv" -f $xmlFileName.Replace('.xml', '')

[String]$generalOutputFilePath = ".\demo\{0}\{1}" -f $version, $generalOutputFileName

$detailOutputFileName = "{0}-Output-Details.csv" -f $xmlFileName.Replace('.xml', '')

[String]$detailOutputFilePath = ".\demo\{0}\{1}" -f $version, $detailOutputFileName

ConvertFrom-O365AddressesRSS -Path $xmlFilePath | Export-Csv -Path $generalOutputFilePath -NoTypeInformation -Encoding UTF8 -delimiter ";"

ConvertFrom-O365AddressesRSS -Path $xmlFilePath  | select-object -Property Guid,PublicationDate,DescriptionIsParsable,OperationType,QuickDescription -ExpandProperty SubChanges | Export-Csv -Path $detailOutputFilePath -NoTypeInformation -Encoding UTF8 -delimiter ";"