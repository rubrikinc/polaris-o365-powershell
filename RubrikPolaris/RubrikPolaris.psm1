
$Core = @(Get-ChildItem -Path $PSScriptRoot\Core\*.ps1 -Recurse -ErrorAction SilentlyContinue) | Sort-Object Name
$M365 = @(Get-ChildItem -Path $PSScriptRoot\M365\*.ps1 -Recurse -ErrorAction SilentlyContinue) | Sort-Object Name

# Dots source the private files
foreach ($import in @($M365 + $Core)) {
	try {
		. $import.fullName
		Write-Verbose -Message ("Imported private function {0}" -f $import.fullName)
	} catch {
		Write-Error -Message ("Failed to import private function {0}: {1}" -f $import.fullName, $_)
	}
}

Export-ModuleMember -Function $Public.BaseName