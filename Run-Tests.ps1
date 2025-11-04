# Run all Pester tests in the tests folder

param(
    [string]$TestPath = "$PSScriptRoot\Tests"
)

Import-Module Pester -ErrorAction Stop

Invoke-Pester -Path $TestPath -Output Detailed
