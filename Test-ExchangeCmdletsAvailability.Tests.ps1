$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$sut = (Split-Path -Leaf $MyInvocation.MyCommand.Path) -replace '\.Tests\.', '.'
. "$here\$sut"

Describe "Test-ExchangeCmdletsAvailability" {
    
    function Get-ExchangeServer { }
    
    Context "Exchange cmdlets available" {
        
        Mock -CommandName Test-Path -MockWith { Return $true }
        
        It "Cmdlets available, CheckExchangeServersAvailability not selected" {
            
            Test-ExchangeCmdletsAvailability | Should Be 0            
            
        }
        
    }
    
    Context "Exchange cmdlets available" {
        
        Mock -CommandName Test-Path -MockWith { Return $true }
        
        Mock -CommandName Get-ExchangeServer -MockWith { Return @( 'Server1', 'Server2' ) }
        
        It "Cmdlets available, CheckExchangeServersAvailability selected, Exchange Server available" {
            
            Test-ExchangeCmdletsAvailability -CheckExchangeServersAvailability | Should Be 0
            
        }
        
    }
    
    Context "Exchange cmdlets available" {
        
        Mock -CommandName Test-Path -MockWith { Return $true }
        
        Mock -CommandName Get-ExchangeServer -MockWith { Return $null }
        
        It "Cmdlets available, CheckExchangeServersAvailability selected, Exchange Server don't available" {
            
            Test-ExchangeCmdletsAvailability -CheckExchangeServersAvailability | Should Be 2
            
        }
        
    }
    
    Context "Exchange cmdlets don't available" {
        
        Mock -CommandName Test-Path -MockWith { Return $false }
        
        It "Cmdlets available, CheckExchangeServersAvailability not selected " {
            
            Test-ExchangeCmdletsAvailability | Should Be 1
            
        }
        
    }
    
}
