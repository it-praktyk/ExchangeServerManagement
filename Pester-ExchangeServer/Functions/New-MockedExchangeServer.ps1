<#
    .SYNOPSIS
    Function intended to create ExchangeServer object to use with Mock in Pester

    .DESCRIPTION
    Function intended to create ExchangeServer object to use with Mock in Pester.

    .PARAMETER Name
    Name of Exchange server.

    .PARAMETER AdminDisplaVersion
    Value for AdminDisplaVersion .

    .PARAMETER IsClientAccessServer
    Set to true if server has installed Client Access Role.

    .PARAMETER IsFrontendTransportServer
    Set to true if server has installed Client Access Role.

    .PARAMETER IsHubTransportServer
    Set to true if server has installed Mailbox Role.

    .PARAMETER IsMailboxServer
    Set to true if server has installed Mailbox Role.

    .PARAMETER IsUnifiedMessagingServer
    Set to true if server has installed Mailbox Role.

    .PARAMETER IsEdgeServer
    Set to true if a server is a Edge.

    .EXAMPLE
    PS > New-MockedExchangeServer | Get-Member

    TypeName: System.Management.Automation.PSCustomObject

    Name                      MemberType   Definition
    ----                      ----------   ----------
    Equals                    Method       bool Equals(System.Object obj)
    GetHashCode               Method       int GetHashCode()
    GetType                   Method       type GetType()
    ToString                  Method       string ToString()
    AdminDisplayVersion       NoteProperty string AdminDisplayVersion=Version 15.0 (Build 1178.4)
    IsClientAccessServer      NoteProperty bool IsClientAccessServer=True
    IsEdgeServer              NoteProperty bool IsEdgeServer=False
    IsFrontendTransportServer NoteProperty bool IsFrontendTransportServer=True
    IsHubTransportServer      NoteProperty bool IsHubTransportServer=True
    IsMailboxServer           NoteProperty bool IsMailboxServer=True
    IsUnifiedMessagingServer  NoteProperty bool IsUnifiedMessagingServer=True
    Name                      NoteProperty string Name=EX-1

    .NOTES
    AUTHOR: Wojciech Sciesinski, wojciech[at]sciesinski[dot]net
    KEYWORDS: PowerShell,Pester,Mock,Exchange

    VERSIONS HISTORY
    0.2.0 - 2016-05-10 - initial working version
	0.2.1 - 2016-05-11 - the first version published to GitHub

    LICENSE
    Copyright (c) 2016 Wojciech Sciesinski
    This function is licensed under The MIT License (MIT)
    Full license text: https://opensource.org/licenses/MIT

#>
function New-MockedExchangeServer {
    [CmdletBinding(DefaultParameterSetName = 'IsNonEdgeServer')]
    [OutputType('System.Management.Automation.PSObject')]
    param (
        [Parameter(ParameterSetName = 'IsEdgeServer',
                   ValueFromPipelineByPropertyName = $true)]
        [Parameter(ParameterSetName = 'IsNonEdgeServer',
                   ValueFromPipelineByPropertyName = $true)]
        [String]$Name = 'EX-1',
        [Parameter(ParameterSetName = 'IsEdgeServer')]
        [Parameter(ParameterSetName = 'IsNonEdgeServer')]
        [String]$AdminDisplaVersion = 'Version 15.0 (Build 1178.4)',
        [Parameter(ParameterSetName = 'IsNonEdgeServer')]
        [bool]$IsClientAccessServer = $true,
        [Parameter(ParameterSetName = 'IsNonEdgeServer')]
        [bool]$IsFrontendTransportServer = $true,
        [Parameter(ParameterSetName = 'IsNonEdgeServer')]
        [bool]$IsHubTransportServer = $true,
        [Parameter(ParameterSetName = 'IsNonEdgeServer')]
        [bool]$IsMailboxServer = $true,
        [Parameter(ParameterSetName = 'IsNonEdgeServer')]
        [bool]$IsUnifiedMessagingServer = $true,
        [Parameter(ParameterSetName = 'IsEdgeServer')]
        [ValidateSet($true)]
        [bool]$IsEdgeServer
    )

    Process {

        switch ($PsCmdlet.ParameterSetName) {

            'IsEdgeServer' {

                $properties = @{
                    'Name' = $Name;
                    'AdminDisplayVersion' = $AdminDisplaVersion;
                    'IsEdgeServer' = $true;
                    'IsClientAccessServer' = $false;
                    'IsFrontendTransportServer' = $false;
                    'IsHubTransportServer' = $false;
                    'IsMailboxServer' = $false;
                    'IsUnifiedMessagingServer' = $false
                }

            }

            'IsNonEdgeServer' {

                $properties = @{
                    'Name' = $Name;
                    'AdminDisplayVersion' = $AdminDisplaVersion;
                    'IsEdgeServer' = $false;
                    'IsClientAccessServer' = $IsClientAccessServer;
                    'IsFrontendTransportServer' = $IsFrontendTransportServer;
                    'IsHubTransportServer' = $IsHubTransportServer;
                    'IsMailboxServer' = $IsMailboxServer;
                    'IsUnifiedMessagingServer' = $IsUnifiedMessagingServer
                }

            }

        }

        $ExchangeServer = New-Object -TypeName System.Management.Automation.PSObject -Property $properties

    }

    End {

        Return $ExchangeServer

    }

}