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
    Import-Module "$currentPath\EnmarchaFunctions.psm1" -PassThru -Force -DisableNameChecking | Out-Null
    Write-Host -ForegroundColor Yellow "Creando la lista $Path"
	
    $ctx = Get-PnPContext
    $currentPath = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition

    $strFileName = "$Path\manifest.xml"
    If (Test-Path $strFileName) {

        [xml]$manifest = Get-Content "$Path\manifest.xml"
        $url = "$tenant$($UrlWebApplication)"
        $urlsite = "$url$($manifest.List.Url)"
        Write-Host -ForegroundColor DarkMagenta $urlsite

        Connect-PnPOnline -Url $urlsite -Credentials $credentials
		
        $existingList = Get-PnpList | Where-Object { $_.Title -eq $manifest.List.Name} 
        if ($existingList.Title -ne $null) {
            Write-Host -ForegroundColor Yellow  "La lista '"$manifest.List.Name"' ya existe"
        }
        else {
            New-PnPList -Title $manifest.List.Name -Template DocumentLibrary  -EnableContentTypes  -OnQuickLaunch 
        }
        if ($manifest.List.ContentTypes -ne $null) {
            if ($manifest.List.ContentTypes.Add -ne $null) {
                $manifest.List.ContentTypes.Add | % {
                    Write-Host -ForegroundColor Green "Agregando el Content Type "$_.Name" a la lista $Path"
                    if ($_.DefaultContentType -ne $null -and $_.DefaultContentType.ToLower() -eq "true") {
                        Add-PnPContentTypeToList -List $manifest.List.Name -ContentType $_.Name -DefaultContentType
                    }
                    else {
                        Add-PnPContentTypeToList -List $manifest.List.Name -ContentType $_.Name
                    }
                }

                $context = New-Object Microsoft.SharePoint.Client.ClientContext($urlsite)
                $cred = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($credentials.UserName, $credentials.Password)
                $context.Credentials = $cred
                $web = $context.Web
                $availableCTs = $web.AvailableContentTypes
                $lists = $web.Lists
                $list = $lists.GetByTitle($manifest.List.Name)
                $listCTs = $list.ContentTypes
                $context.Load($web)
                $context.Load($availableCTs)
                $context.Load($lists)
                $context.Load($list)
                $context.Load($listCTs)
                $context.ExecuteQuery()

                # Fix names in variations webs
                $manifest.List.ContentTypes.Add | % {
                    $ctName = $_.Name
                    $ctWeb = $availableCTs | where {$_.Name -eq $ctName}

                    if ($ctWeb) {
                        $ctId = ($ctWeb.Id.StringValue+"*")
                        $ctList = $listCTs | where {$_.Id.StringValue -clike $ctId}
                        if ($ctWeb.Name -ne $ctList.Name) {
                            $ctList.Name = $ctWeb.Name
                            $ctList.Update($false)
                        }

                        $listFields = $list.Fields
                        $ctFields = $ctWeb.Fields
                        $context.Load($listFields)
                        $context.Load($ctFields)
                        $context.ExecuteQuery()

                        $ctFields | % {
                            $internalName = $_.InternalName
                            $fieldList = $listFields | where {$_.InternalName -eq $internalName}
                            if ($fieldList.Title -ne $_.Title) {
                                $fieldList.Title = $_.Title
                                $fieldList.Update()
                            }
                        }

                        $context.ExecuteQuery()
                    }
                }
            }
            if ($manifest.List.ContentTypes.Remove -ne $null) {
                $manifest.List.ContentTypes.Remove | % {
                    Remove-PnPContentTypeFromList -List $manifest.List.Name -ContentType $_.Name
                }
            }
        }
		
        if ($manifest.List.DocumentSets -ne $null) {
            $contador = 1;
            $manifest.List.DocumentSets.ContentTypes.ContentType | % {			   
                if ($contador -eq 1) {
                    Write-Host -ForegroundColor Green "Creo el Document Set $manifest.List.DocumentSets.name"
                    Write-Host $manifest.List.Name
                    Write-Host $_.name
                    Write-Host $manifest.List.DocumentSets.Name
                    Add-PnPDocumentSet -List $manifest.List.Name -ContentType $_.Name -Name $manifest.List.DocumentSets.Name -ErrorAction SilentlyContinue
                    $contador = 2;
                }			   
                else {
                    Write-Host -ForegroundColor Green "Agregando el Content Type {$_.Name} al DocumentSet $Path"
                    Add-PnPContentTypeToDocumentSet -ContentType $_.Name -DocumentSet $manifest.List.DocumentSets.Name -ErrorAction SilentlyContinue
                }
            }
        }

        if ($manifest.List.Versioning -ne $null) {
            Write-Host -ForegroundColor Green "Configurando el versionado de la lista"
            $enableVersioning = $false
            if ($manifest.List.Versioning.EnableVersioning.ToLower() -eq "true") {
                $enableVersioning = $true
            }
            $enableMinorVersioning = $false
            if ($manifest.List.Versioning.EnableMinorVersioning.ToLower() -eq "true") {
                $enableMinorVersioning = $true
            }
            if ($manifest.List.Versioning.MajorVersions -ne $null) {
                if ($manifest.List.Versioning.MinorVersions -ne $null) {
                    Set-PnPList -Identity $manifest.List.Name -EnableVersioning $enableVersioning -EnableMinorVersions $enableMinorVersioning -MajorVersions $manifest.List.Versioning.MajorVersions -MinorVersions $manifest.List.Versioning.MinorVersions
                }
                else {
                    Set-PnPList -Identity $manifest.List.Name -EnableVersioning $enableVersioning -EnableMinorVersions $enableMinorVersioning -MajorVersions $manifest.List.Versioning.MajorVersions
                }
            }
            else {
                if ($manifest.List.Versioning.MinorVersions -ne $null) {
                    Set-PnPList -Identity $manifest.List.Name -EnableVersioning $enableVersioning -EnableMinorVersions $enableMinorVersioning -MinorVersions $manifest.List.Versioning.MinorVersions
                }
                else {
                    Set-PnPList -Identity $manifest.List.Name -EnableVersioning $enableVersioning -EnableMinorVersions $enableMinorVersioning
                }
            }
			
            $ctx2 = New-Object Microsoft.SharePoint.Client.ClientContext($urlsite)
            $ctx2.Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($credentials.UserName, $credentials.Password)
            $ctx2.Load($ctx2.Web)
            $list = $ctx2.Web.Lists.GetByTitle($manifest.List.Name)
            $ctx2.Load($list)

            if ($manifest.List.Versioning.ForceCheckout -ne $null) {
                if ($manifest.List.Versioning.ForceCheckout.ToLower() -eq "true") {
                    $list.ForceCheckout = $true
                }
                else {
                    $list.ForceCheckout = $false
                }
            }
			
            if ($manifest.List.Versioning.DraftVersionVisibility -ne $null) {
                if ($manifest.List.Versioning.DraftVersionVisibility.ToLower() -eq "reader") {
                    $list.DraftVersionVisibility = 0
                }
                else {
                    if ($manifest.List.Versioning.DraftVersionVisibility.ToLower() -eq "approver") {
                        $list.DraftVersionVisibility = 2
                    }
                    else {
                        $list.DraftVersionVisibility = 1
                    }
                }
            }
			
            if ($manifest.List.Versioning.EnableModeration -ne $null) {
                if ($manifest.List.Versioning.EnableModeration.ToLower() -eq "true") {
                    $list.EnableModeration = $true
                }
                else {
                    $list.EnableModeration = $false
                }
            }

            $list.Update()
            $ctx2.ExecuteQuery()
            $ctx2.Dispose()
        }

        if ($manifest.List.Views -ne $null) {
            $manifest.List.Views.View | % {	
                $view = Get-PnPView -List $manifest.List.Name -Identity $_.Name -ErrorAction SilentlyContinue
                if (-not $view) {
                    $query = Convert-XmlElementToString($_.Query) 
                    $query = $query.Replace("<Query>", "")
                    $query = $query.Replace("</Query>", "")
                    $field = $_.Fields -split ","
                    $resultField = New-Object string[] $field.Count
                    For ($i = 0; $i -le $field.Count - 1; $i++) {					 
                        $resultField[$i] = $field[$i]
                    }

                    if ($_.Default -eq "true") {
                        Write-Host "Creando vista por defecto" $_.Name
                        Add-PnPView -List $manifest.List.Name -Title $_.Name -Query $query -Fields $resultField  -Paged -SetAsDefault
                            
                    }
                    else {
                        Write-Host "Creando vista" $_.Name
                        Add-PnPView -List $manifest.List.Name -Title $_.Name -Query $query  -Fields $resultField -Paged
                    }
                }
            }
        }
        if ($manifest.List.DefaultValues -ne $null) {
            $manifest.List.DefaultValues.DefaultValue | % {
                $taxonomyItem = $_.Value -split ";"
                Write-Host $_.Field " Count" $taxonomyItem.Count
                if ($taxonomyItem.Count -gt 1 ) {
                    $item = Get-PnPTerm -Identity $taxonomyItem[2] -TermGroup $taxonomyItem[0] -TermSet $taxonomyItem[1]
                    Write-Host "Id de la Taxonomia es " $item.Id "-" $item.Name
	
                    Write-Host "Set-PnPDefaultColumnValues -List" $manifest.List.Name " -Field " $_.Field " -Value " $item.Id -ForegroundColor Green
                    Set-PnPDefaultColumnValues -List $manifest.List.Name -Field $_.Field -Value $item.Id
                }
                else {
                    Write-Host "Set-PnPDefaultColumnValues -List" $manifest.List.Name " -Field " $_.Field " -Value " $_.Value -ForegroundColor Green
                    Set-PnPDefaultColumnValues -List $manifest.List.Name -Field $_.Field -Value $_.Value
                }
            }
        }

        if ($manifest.List.Permissions -ne $null) {
            $copyRoleAssignments = $false
            if ($manifest.List.Permissions.CopyRoleAssignments -ne $null -and $manifest.List.Permissions.CopyRoleAssignments.ToLower() -eq "true") {
                $copyRoleAssignments = $true
            }
            if ($copyRoleAssignments) {
                Set-PnPList -Identity $manifest.List.Name -BreakRoleInheritance -CopyRoleAssignments
            } else {
                Set-PnPList -Identity $manifest.List.Name -BreakRoleInheritance
            }

            if ($manifest.List.Permissions.Add -ne $null) {
                Write-Host "Agregando permisos a" $manifest.List.Name
                $manifest.List.Permissions.Add | % {
                    if ($_.User -ne $null) {
                        Set-PnPListPermission -Identity $manifest.List.Name -User $_.User -AddRole $_.Role    
                    }
                    if ($_.Group -ne $null) {
                        Set-PnPListPermission -Identity $manifest.List.Name -Group $_.Group -AddRole $_.Role    
                    }
                }
            }
            if ($manifest.List.Permissions.Remove -ne $null) {
                Write-Host "Quitando permisos a" $manifest.List.Name
                $manifest.List.Permissions.Remove | % {
                    if ($_.User -ne $null) {
                        Set-PnPListPermission -Identity $manifest.List.Name -User $_.User -RemoveRole $_.Role    
                    }
                    if ($_.Group -ne $null) {
                        Set-PnPListPermission -Identity $manifest.List.Name -Group $_.Group -RemoveRole $_.Role    
                    }
                }
            }
        }
    }
    Set-PnPContext -Context $ctx # switch back to site A
}