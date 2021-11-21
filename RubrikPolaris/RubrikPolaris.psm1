
$Functions = @(Get-ChildItem -Path $PSScriptRoot\*\*.ps1 -Recurse -ErrorAction SilentlyContinue) | Sort-Object Name


# Dots source the private files
foreach ($import in $Functions) {
	try {
		. $import.fullName
		Write-Verbose -Message ("Imported private function {0}" -f $import.fullName)
	} catch {
		Write-Error -Message ("Failed to import private function {0}: {1}" -f $import.fullName, $_)
	}
}

Export-ModuleMember -Function $Public.BaseName