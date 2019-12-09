function Get-Presence {
    [CmdletBinding(DefaultParameterSetName = "Default")]
    param (
        [Parameter(Position = 0, Mandatory = $true)]
        [string]
        $UPN,
        
        [Parameter(Position = 1, Mandatory = $false)]
        [string]
        $ClientId = "0f7120fe-24e2-49fc-a492-2d8032e41b68",

        [Parameter(Position = 1, Mandatory = $false)]
        [string]
        $TargetUser 
        )

    Process {
        if (Test-Path ($script:ModuleRoot + "/Microsoft.Identity.Client.dll")) {
            Import-Module ($script:ModuleRoot + "/Microsoft.Identity.Client.dll")
            write-verbose ("Using MSAL dll from Local Directory")
        }
        $scope = "https://graph.microsoft.com/Presence.Read";
        $redirectUri = "https://login.microsoftonline.com/common/oauth2/nativeclient";
        $domainName = $UPN.Split('@')[1];
        $Scopes = New-Object System.Collections.Generic.List[string]
        $Scopes.Add($Scope)
        $TenantId = (Invoke-WebRequest https://login.windows.net/$domainName/v2.0/.well-known/openid-configuration | ConvertFrom-Json).token_endpoint.Split('/')[3]
        $pcaConfig = [Microsoft.Identity.Client.PublicClientApplicationBuilder]::Create($ClientId).WithTenantId($TenantId).WithRedirectUri($redirectUri)
        $TokenResult = $pcaConfig.Build().AcquireTokenInteractive($Scopes).WithPrompt([Microsoft.Identity.Client.Prompt]::Never).WithLoginHint($UPN).ExecuteAsync().Result;
        $Header = @{
            'Content-Type'  = 'application\json'
            'Authorization' = $TokenResult.CreateAuthorizationHeader()
        }
        if([String]::IsNullOrEmpty($TargetUser)){
            return (Invoke-RestMethod -Headers $Header -Uri ("https://graph.microsoft.com/beta/me/presence") -Method Get -ContentType "Application/json")
        }else{
            return (Invoke-RestMethod -Headers $Header -Uri ("https://graph.microsoft.com/beta/users('$TargetUser')/presence") -Method Get -ContentType "Application/json")
        }
        
    }

}
$script:ModuleRoot = $PSScriptRoot


