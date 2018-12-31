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
    Write-Host -ForegroundColor Yellow "Creando la lista $Path"
	  
    $currentPath = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition
    $ctx = Get-PnPContext

    $strFileName = "$Path\manifest.xml"
    If (Test-Path $strFileName) {
        [xml]$manifest = Get-Content "$Path\manifest.xml"
        $url = "$tenant$($UrlWebApplication)"
        $urlsite = "$url$($manifest.List.Url)"

        Connect-PnPOnline -Url $urlsite -Credentials $credentials 
		
        $existingList = Get-PnpList | Where-Object { $_.Title -eq $manifest.List.Name} 
        if ($existingList.Title -ne $null) {
            Write-Host -ForegroundColor Yellow  "La lista '" $manifest.List.Name"' ya existe"
        }
        else {
            $list = New-PnPList -Title $manifest.List.Name -Template GenericList   -EnableContentTypes  -OnQuickLaunch 
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
                    $ctId = ($ctWeb.Id.StringValue + "*")
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
            if ($manifest.List.Versioning.MajorVersions -ne $null) {
                Set-PnPList -Identity $manifest.List.Name -EnableVersioning $enableVersioning -MajorVersions $manifest.List.Versioning.MajorVersions
            }
            else {
                Set-PnPList -Identity $manifest.List.Name -EnableVersioning $enableVersioning
            }
			
            $ctx2 = New-Object Microsoft.SharePoint.Client.ClientContext($urlsite)
            $ctx2.Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($credentials.UserName, $credentials.Password)
            $ctx2.Load($ctx2.Web)
            $list = $ctx2.Web.Lists.GetByTitle($manifest.List.Name)
            $ctx2.Load($list)
			
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
                try {
                    Remove-PnPView -List $manifest.List.Name -Identity $_.Name -Force
                }
                catch {}

                $query = Convert-XmlElementToString($_.Query)
                $query = $query.Replace("<Query>", "")
                $query = $query.Replace("</Query>", "")
                $field = $_.Fields -split ","
                $resultField = New-Object string[] $field.Count
                For ($i = 0; $i -le $field.Count - 1; $i++) {					 
                    $resultField[$i] = $field[$i]
                }

                if ($_.Default -eq "true") {
                    Add-PnPView -List $manifest.List.Name -Title $_.Name -Query $query -Fields $resultField  -Paged -SetAsDefault
						
                }
                else {
                    Add-PnPView -List $manifest.List.Name -Title $_.Name -Query $query  -Fields $resultField -Paged
                }
            }
		
        }

        if ($manifest.List.ItemLevelPermissions -ne $null) {
            $existingList = Get-PnPList | Where-Object { $_.Title -eq $manifest.List.Name}

            if ($manifest.List.ItemLevelPermissions.ReadSecurity -eq "ReadAllItems") {
                $existingList.ReadSecurity = 1
            }
            if ($manifest.List.ItemLevelPermissions.ReadSecurity -eq "ReadCreatedByUser") {
                $existingList.ReadSecurity = 2
            }

            if ($manifest.List.ItemLevelPermissions.WriteSecurity -eq "EditAllItems") {
                $existingList.WriteSecurity = 1
            }
            if ($manifest.List.ItemLevelPermissions.WriteSecurity -eq "EditCreatedByUser") {
                $existingList.WriteSecurity = 2
            }
            if ($manifest.List.ItemLevelPermissions.WriteSecurity -eq "None") {
                $existingList.WriteSecurity = 4
            }

            $existingList.Update()
            $existingList = Get-PnPList | Where-Object { $_.Title -eq $manifest.List.Name} 
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

        #Write-Host "Insertando en la lista " $manifest.List.Name
        #$result= & "$currentPath\New-Item.ps1" -UrlWebApplication $UrlWebApplication -tenant $tenant  -Path "$Path"  -ListName $manifest.List.Name
        Set-PnPContext -Context $ctx # switch back to site A
    }
}