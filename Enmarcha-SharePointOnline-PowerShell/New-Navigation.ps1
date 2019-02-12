Param
(
    [Parameter(Mandatory = $true)]
    [string]$Path,
    [Parameter(Mandatory = $true)]
    [string]$urlWebapplication,
    [Parameter(Mandatory = $true)]
    [string]$tenant,
    [Parameter(Mandatory = $true)]
    [System.Management.Automation.PSCredential]$credentials 
)
Process {
    function Update-WebNavigationConfig {
        param (
            [Parameter(Mandatory = $true)]
            [string]$WebUrl,
            [Parameter(Mandatory = $true)]
            [System.Management.Automation.PSCredential]$Credentials
        )

        $cred = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($Credentials.UserName, $Credentials.Password) 
        $context = New-Object Microsoft.SharePoint.Client.ClientContext($WebUrl)
        $context.Credentials = $cred
 
        $web = $context.Web
        $context.Load($web)
        $context.ExecuteQuery() 
        $navigationSettings = New-Object Microsoft.SharePoint.Client.Publishing.Navigation.WebNavigationSettings($context, $web)
        $context.Load($navigationSettings)
        $context.Load($navigationSettings.GlobalNavigation)
        $context.Load($navigationSettings.CurrentNavigation)
        $context.ExecuteQuery()
        $navigationSettings.GlobalNavigation.Source = "PortalProvider"
        $navigationSettings.CurrentNavigation.Source = "PortalProvider"
        $navigationSettings.Update($null)
        $context.ExecuteQuery()
    }
	  
    $ctx = Get-PnPContext

    $strFileName = "$Path\manifest.xml"
    If (Test-Path $strFileName) {
        [xml]$manifest = Get-Content "$Path\manifest.xml"
        $url = ($tenant + $UrlWebApplication + "/" + $manifest.Site.RelativeUrl)

        if ($manifest.Site.Navigation -ne $null) {
            Update-WebNavigationConfig -WebUrl $url -Credentials $credentials    
            Connect-PnPOnline -Url $url -Credentials $credentials 
            
            Write-Host -ForegroundColor Yellow "Modificando la navegación"
            if ($manifest.Site.Navigation.Add -ne $null) {
                $manifest.Site.Navigation.Add | % {
                    Remove-PnPNavigationNode -Title $_.Title -Location $_.Location -Force
                    $completeUrl = $_.Url
                    if (-not $completeUrl.ToLower().StartsWith("http")) {
                        $completeUrl = $UrlWebApplication + $_.Url
                    }
                    if ($_.IsExternal -ne $null -and $_.IsExternal.ToLower() -eq "true") {
                        Add-PnPNavigationNode -Location $_.Location -Title $_.Title -External -Url $completeUrl
                    }
                    else {
                        Add-PnPNavigationNode -Location $_.Location -Title $_.Title -Url $completeUrl
                    }
                }
            }
            if ($manifest.Site.Navigation.Remove -ne $null) {
                $manifest.Site.Navigation.Remove | % {
                    Remove-PnPNavigationNode -Title $_.Title -Location $_.Location -Force
                }
            }
    
            Write-Host -ForegroundColor Green "Navegación actualizada"
        
            Set-PnPContext -Context $ctx # switch back to site A
        }
    }
}