<#
	.SYNOPSIS
		Function to query mailbox data on multiply mailbox servers synchronously
	
	.DESCRIPTION
		A detailed description of the Invoke-MultiMailboxServerQuery function.
	
	.PARAMETER ComputerName
		A description of the ComputerName parameter.
	
	.PARAMETER Filter
		A description of the Filter parameter.
	
	.PARAMETER Site
		A description of the Site parameter.
	
	.PARAMETER LimitSession
		A description of the LimitSession parameter.
	
	.PARAMETER Credential
		A description of the Credential parameter.
	
	.OUTPUTS
		pscustomobject, pscustomobject, pscustomobject
	
	.NOTES
	
    AUTHOR: Wojciech Sciesinski, wojciech[at]sciesinski[dot]net  
    KEYWORDS: PowerShell, Exchange, report
   
    VERSIONS HISTORY
    - 0.1.0 - 2016-06-30 - First draft published on GitHub
    
	TODO
	- Add posibilites to multisite search


    LICENSE
    Copyright (c) 2016 Wojciech Sciesinski  
    This function is licensed under The MIT License (MIT)  
    Full license text: https://opensource.org/licenses/MIT  
   
#>
	

#>
function Invoke-MultiMailboxServerQuery
{
	[CmdletBinding(DefaultParameterSetName = 'ByName')]
	[OutputType([pscustomobject], ParameterSetName = 'BySite')]
	[OutputType([pscustomobject], ParameterSetName = 'ByFilter')]
	[OutputType([pscustomobject], ParameterSetName = 'ByName')]
	[OutputType([pscustomobject])]
	param
	(
		[Parameter(ParameterSetName = 'ByName',
				   Mandatory = $false,
				   ValueFromPipeline = $true,
				   ValueFromPipelineByPropertyName = $true)]
		[Alias('CN', '__Server', 'Server')]
		[String[]]$ComputerName = @('All'),
		[Parameter(ParameterSetName = 'ByFilter',
				   Mandatory = $false)]
		[ScriptBlock]$Filter,
		[Parameter(ParameterSetName = 'BySite',
				   Mandatory = $false)]
		[String]$Site,
		[Parameter(Mandatory = $false)]
		[Int]$LimitSession = 5,
		[Parameter(Mandatory = $false)]
		[System.Management.Automation.Credential()]$Credential,
        [Parameter(Mandatory = $false)]
        [ValidateSet("Basic", "Custom")]
        [String]$PropertiesSet
	)
	
	
	
	Write-Verbose $PsCmdlet.ParameterSetName
	
	switch ($PsCmdlet.ParameterSetName)
	{
		
		'ByName' { }
		
		'ByFilter' { }
		
		'BySite' {  }
		
	}
	
	foreach ($CurrentMailboServer in $MailboxServers)
	{
		
		
		
		
	}
}

Function New-ExchangePSSession {
	
	<#

    
	TODO
	- Add support for multiply servers
	
	#>
	
	[cmdletbinding()]
	param (
		
		[parameter(Mandatory = $true, ValueFromPipeline = $true)]
		[Alias('CN', '__Server', 'Server')]
		[String[]]$ComputerName,
		[parameter(Mandatory = $true)]
		[System.Management.Automation.Credential()]$Credential
		
	)
	
	begin
	{
		
	}
	
	process
	{

			Try
			{
				
				$FQDName = [System.Net.Dns]::GetHostByName($ComputerName).HostName
				
				$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://<FQDN of Exchange 2010 server>/PowerShell/ -Authentication Kerberos -Credential $Credential
				
			}
			Catch
			{
				
				
				
			}

		
	}
	
	End
	{
		
		
		
	}
	
}