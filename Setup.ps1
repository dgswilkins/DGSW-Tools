$guid = [guid]::NewGuid().guid
$path = "$env:userprofile\documents\WindowsPowerShell\Modules\DGSW-Tools\DGSW-Tools.psd1"
$paramHash = @{
 Path = $path
 RootModule = "DGSW-Tools.psm1"
 Author = "Douglas Wilkins"
 CompanyName = ""
 ModuleVersion = "1.0"
 Guid = $guid
 PowerShellVersion = "5.0"
 Description = "My Tools module"
 FormatsToProcess = ""
 FunctionsToExport = "Get-SavedCred"
 AliasesToExport = ""
 VariablesToExport = ""
 CmdletsToExport = ""
}
New-ModuleManifest @paramHash
