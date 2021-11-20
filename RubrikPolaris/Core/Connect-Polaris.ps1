function Connect-Polaris() {
    <#
    .SYNOPSIS

     Connect to a Polaris Account using a Service Account. This is the recommended connection method. 

    .DESCRIPTION

    Returns an API access token for a given Polaris account. The cmdlet requires a Service Account JSON file stored at ~/.rubrik/polaris-service-account.json.

    .INPUTS

    None. You cannot pipe objects to Get-PolarisTokenServiceAccount.

    .OUTPUTS

    System.String. Get-PolarisTokenServiceAccount returns a string containing the access token required to connect to the Polaris GraphQL API.
 
    .EXAMPLE

    PS> $token = Get-PolarisTokenServiceAccount
    #>

    Write-Information -Message "Info: Attempting to read the Service Account file located at ~/.rubrik/polaris-service-account.json "
    try {
        $serviceAccountFile = Get-Content -Path "~/.rubrik/polaris-service-account.json" -ErrorAction Stop | ConvertFrom-Json 
    }
    catch {
        $errorMessage = $_.Exception | Out-String

        if($errorMessage.Contains('because it does not exist')) {
            throw "The Service Account JSON secret file was not found. Ensure the file is location at ~/.rubrik/polaris-service-account.json."
        } 
        
        throw $_.Exception
        
    }


    $payload = @{
        grant_type = "client_credentials";
        client_id = $serviceAccountFile.client_id;
        client_secret = $serviceAccountFile.client_secret
    }   

    Write-Debug -Message "Determing if the Service Account file contains all required variables."
    $missingServiceAccount = @()
    if ($serviceAccountFile.client_id -eq $null) {
        $missingServiceAccount += "'client_id'"
    }

    if ($serviceAccountFile.client_secret -eq $null) {
        $missingServiceAccount += "'client_secret'"
    }

    if ($serviceAccountFile.access_token_uri -eq $null) {
        $missingServiceAccount += "'access_token_uri'"
    }


    if ($missingServiceAccount.count -gt 0){
        throw "The Service Account JSON secret file is missing the required paramaters: $missingServiceAccount"
    }


    $headers = @{
        'Content-Type' = 'application/json';
        'Accept'       = 'application/json';
    }
    
    Write-Debug -Message "Connecting to the Polaris GraphQL API using the Service Account JSON file."
    $response = Invoke-RestMethod -Method POST -Uri $serviceAccountFile.access_token_uri -Body $($payload | ConvertTo-JSON -Depth 100) -Headers $headers

    Write-Verbose -Message "Creating the Rubrik Polaris Connection Global variable."
    $global:rubrikPolarisConnection = @{
        accessToken      = $response.access_token;
        PolarisURL  = $serviceAccountFile.access_token_uri.Replace("/api/client_token", "")
    }

}
Export-ModuleMember -Function Connect-Polaris

