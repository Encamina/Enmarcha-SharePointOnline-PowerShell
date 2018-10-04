Function New-SiteColumn() {
    Param
    (
        [Parameter(Mandatory = $true)]
        [ValidateSet("Boolean", "Number", "Text", "LookupMulti", "DateTime", "URL", "Image", "Link", "Note", "HTML", "Calculated", "User", "Choice", "MultiChoice", "TaxonomyFieldType", "TaxonomyFieldTypeMulti", "SummaryLinks", "MediaFieldType", "Currency")]
        [string]$FieldType,

        [Parameter(Mandatory = $true)]
        [string]$InternalName,

        [Parameter(Mandatory = $true)]
        [string]$DisplayName,

        [Parameter(Mandatory = $false)]
        [guid]$Id = [System.Guid]::NewGuid(),

        [Parameter(Mandatory = $false)]
        [System.Collections.Generic.Dictionary[System.Globalization.CultureInfo, string]]$LocalizedDisplayNames,

        [Parameter(Mandatory = $true)]
        [string]$StaticName,

        [Parameter(Mandatory = $true)]
        [string]$Group,

        [Parameter(Mandatory = $false)]
        [Nullable[bool]]$Hidden = $null,

        [Parameter(Mandatory = $false)]
        [Nullable[bool]]$Required = $null,

        [Parameter(Mandatory = $false)]
        [Nullable[bool]]$Sealed = $null,

        [Parameter(Mandatory = $false)]
        [string]$MaxLength = $null,

        [Parameter(Mandatory = $false)]
        [Nullable[bool]]$ShowInDisplayForm = $null,

        [Parameter(Mandatory = $false)]
        [Nullable[bool]]$ShowInEditForm = $null,

        [Parameter(Mandatory = $false)]
        [Nullable[bool]]$ShowInListSettings = $null,

        [Parameter(Mandatory = $false)]
        [Nullable[bool]]$ShowInNewForm = $null,

        [Parameter(Mandatory = $false)]
        [ValidateSet("", "DateOnly", "DateTime")]
        [string]$DateTimeFormat = "",

        [Parameter(Mandatory = $false)]
        [ValidateSet("", "Image", "Hyperlink")]
        [string]$UrlFormat = "",

        [Parameter(Mandatory = $false)]
        [ValidateSet("", "TRUE", "FALSE")]
        [string]$ImageRichText = "",

        [Parameter(Mandatory = $false)]
        [ValidateSet("", "FullHtml", "ThemeHtml")]
        [string]$ImageRichTextMode = "",

        [Parameter(Mandatory = $false)]
        [ValidateSet("", "TRUE", "FALSE")]
        [string]$LinkRichText = "",

        [Parameter(Mandatory = $false)]
        [ValidateSet("", "FullHtml", "ThemeHtml")]
        [string]$LinkRichTextMode = "",
		
        [Parameter(Mandatory = $false)]
        [ValidateSet("", "TRUE", "FALSE")]
        [string]$NoteRichText = "",

        [Parameter(Mandatory = $false)]
        [ValidateSet("", "FullHtml", "ThemeHtml")]
        [string]$NoteRichTextMode = "",

        [Parameter(Mandatory = $false)]
        [string]$UnlimitedLengthInDocumentLibrary = "",

        [Parameter(Mandatory = $false)]
        [ValidateSet("", "TRUE", "FALSE")]
        [string]$HtmlRichText = "",

        [Parameter(Mandatory = $false)]
        [ValidateSet("", "FullHtml", "ThemeHtml")]
        [string]$HtmlRichTextMode = "",

        [Parameter(Mandatory = $false)]
        [ValidateSet("", "Boolean", "Integer", "Text", "Note", "DateTime")]   # http://msdn.microsoft.com/en-us/library/microsoft.sharepoint.spfieldtype(v=office.15).aspx
        [string]$CalculatedFormulaValueType = "",

        [Parameter(Mandatory = $false)]
        [string]$CalculatedFormula = "",

        [Parameter(Mandatory = $false)]
        [ValidateSet("", "PeopleOnly", "PeopleAndGroups")]
        [string]$UserSelectionMode = "",

        [Parameter(Mandatory = $false)]
        [Nullable[bool]]$UserAllowMultipleValues = $null,

        [Parameter(Mandatory = $false)]
        [string[]]$Choices = $null,

        [Parameter(Mandatory = $false)]
        [Nullable[bool]]$IsPathRendered = $null,

        [Parameter(Mandatory = $false)]
        [Nullable[bool]]$IsOpen = $null,

        [Parameter(Mandatory = $false)]
        [string]$TermStoreGroupName,

        [Parameter(Mandatory = $false)]
        [string]$TermSetName,

        [Parameter(Mandatory = $false)]
        [Nullable[bool]]$AllowMultipleValues = $null,

        [Parameter(Mandatory = $false)]
        [string]$CurrencyFormat,

        [Parameter(Mandatory = $false)]
        [string]$DecimalFormat,

        [Parameter(Mandatory = $false)]
        [string]$Description,

        [Parameter(Mandatory = $true)]
        [System.Management.Automation.PSCredential]$Credentials	
    )
    Process {
        $existingField = Get-PnpField | Where-Object { $_.InternalName -eq $InternalName} 		
        if ($existingField.Title -ne $null) {
            Write-Host -ForegroundColor Yellow  "El campo '$InternalName' ya existe"
            if ($LocalizedDisplayNames -ne $null) {
                $cxt = Get-PnPContext
                $creden = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($Credentials.UserName, $Credentials.Password)
                $context = New-Object Microsoft.SharePoint.Client.ClientContext($cxt.Url)
                $context.Credentials = $creden
                $web = $context.Web
                $fields = $context.Web.Fields
                $myfield = $fields.GetByInternalNameOrTitle($InternalName)
                $context.Load($web)
                $context.Load($fields)
                $context.Load($myfield)
                $context.ExecuteQuery()
    
                $LocalizedDisplayNames.GetEnumerator() | % {
                    if ($web.Language -eq $_.Key.LCID) {
                        $myfield.Title = $_.Value
                    }
                    $myfield.TitleResource.SetValueForUICulture($_.Key, $_.Value)
                }
                $myfield.UpdateAndPushChanges($true)
                $context.ExecuteQuery()
            }
            return
        }
		
        $fieldXml = "<Field Type='$FieldType' ID='$Id' Name='$InternalName' DisplayName='$DisplayName' Description='$Description' Required='$Required'></Field>"		
        if ($FieldType -eq "TaxonomyFieldType") {
            $termSetPath = $TermStoreGroupName + "|" + $TermSetName
            Write-Host -ForegroundColor Yellow "El TermsetPath es $termSetPath"
            $taxField = $null
            if ($Required) {
                if ($AllowMultipleValues) {
                    $taxField = Add-PnPTaxonomyField -Id $Id -DisplayName $DisplayName -InternalName $InternalName -Group $Group -TermSetPath $termSetPath -Required -MultiValue
                }
                else {
                    $taxField = Add-PnPTaxonomyField -Id $Id -DisplayName $DisplayName -InternalName $InternalName -Group $Group -TermSetPath $termSetPath -Required
                }
            }
            else {
                if ($AllowMultipleValues) {
                    $taxField = Add-PnPTaxonomyField -Id $Id -DisplayName $DisplayName -InternalName $InternalName -Group $Group -TermSetPath $termSetPath -MultiValue
                }
                else {
                    $taxField = Add-PnPTaxonomyField -Id $Id -DisplayName $DisplayName -InternalName $InternalName -Group $Group -TermSetPath $termSetPath
                }
            }

            if ($taxField -and $IsOpen) {
                $taxField.Open = $true
                $taxField.CreateValuesInEditForm = $true
                $taxField.Update()
                $taxField.Context.ExecuteQuery()
            }
        }
        else {
            if ($FieldType -eq "Choice" -or $FieldType -eq "MultiChoice") {
                if ($Required) {
                    Add-PnPField -Id $Id -DisplayName $DisplayName -InternalName $InternalName -Type $FieldType -Group $Group  -choice $Choices -Required
                }
                else {
                    Add-PnPField -Id $Id -DisplayName $DisplayName -InternalName $InternalName -Type $FieldType -Group $Group  -choice $Choices 
                }
            }

            else {
                $requiredText = "FALSE"
                if ($Required) {
                    $requiredText = "TRUE"
                }

                if ($FieldType -eq "DateTime") {
                    $schema = "<Field ID='" + $Id + "' Type='DateTime' Name='" + $InternalName + "' StaticName='" + $InternalName + "' Group='" + $Group + "'
					DisplayName='" + $DisplayName + "' Required='" + $requiredText + "' Format='DateOnly' ></Field>"
                    Add-PnPFieldFromXml -FieldXml $schema
                }
                else {
                    if ($FieldType -eq "Note") {
                        $schema = "<Field ID='" + $Id + "' Type='Note' Name='" + $InternalName + "' StaticName='" + $InternalName + "' DisplayName='" + $DisplayName + "' Group='" + $Group + "' Required='" + $requiredText +
                        "' NumLines='6'  RichText='" + $NoteRichText + "' RichTextMode='" + $NoteRichTextMode + "' UnlimitedLengthInDocumentLibrary='" + $UnlimitedLengthInDocumentLibrary + "' />"
                        Add-PnPFieldFromXml -FieldXml $schema
                    }
                    else {
                        if ($Required) {
                            Add-PnPField -Id $Id -DisplayName $DisplayName -InternalName $InternalName -Type $FieldType -Group $Group -Required
                        }
                        else {
                            Add-PnPField -Id $Id -DisplayName $DisplayName -InternalName $InternalName -Type $FieldType -Group $Group 
                        }
                    }
                }
            }
        }

        if ($FieldType -eq "User") {
            if ($UserAllowMultipleValues -eq $true) {
                $userField = Get-PnPField -Identity $InternalName
                $userField.AllowMultipleValues = $true
                $userField.Update()
                $userField = Get-PnPField -Identity $InternalName
            }

            if ($UserSelectionMode -eq "PeopleOnly") {
                $userField = Get-PnPField -Identity $InternalName
                $userField.SelectionMode = "PeopleOnly"
                $userField.Update()
                $userField = Get-PnPField -Identity $InternalName
            }
        }

        if ($LocalizedDisplayNames -ne $null) {
            $cxt = Get-PnPContext
            $creden = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($Credentials.UserName, $Credentials.Password)
            $context = New-Object Microsoft.SharePoint.Client.ClientContext($cxt.Url)
            $context.Credentials = $creden
            $web = $context.Web
            $fields = $context.Web.Fields
            $myfield = $fields.GetByInternalNameOrTitle($InternalName)
            $context.Load($web)
            $context.Load($fields)
            $context.Load($myfield)
            $context.ExecuteQuery()

            $LocalizedDisplayNames.GetEnumerator() | % {
                if ($web.Language -eq $_.Key.LCID) {
                    $myfield.Title = $_.Value
                }
                $myfield.TitleResource.SetValueForUICulture($_.Key, $_.Value)
            }
            $myfield.UpdateAndPushChanges($true)
            $context.ExecuteQuery()
        }

        Write-Host   "Campo '$InternalName' creado" -ForegroundColor Green
    }	
}


Function New-SiteContentType() {
    Param
    (
        [Parameter(Mandatory = $true)]
        [string]$ContentTypeId,

        [Parameter(Mandatory = $true)]
        [string]$Name,

        [Parameter(Mandatory = $false)]
        [string]$Description = $null,

        [Parameter(Mandatory = $false)]
        [System.Collections.Generic.Dictionary[System.Globalization.CultureInfo, string]]$LocalizedNames = $null,

        [Parameter(Mandatory = $false)]
        [System.Collections.Generic.Dictionary[System.Globalization.CultureInfo, string]]$LocalizedDescriptions = $null,

        [Parameter(Mandatory = $true)]
        [string]$Group,

        [Parameter(Mandatory = $false)]
        [string]$DisplayFormUrl = $null,

        [Parameter(Mandatory = $false)]
        [string]$EditFormUrl = $null,

        [Parameter(Mandatory = $false)]
        [string]$NewFormUrl = $null,

        [Parameter(Mandatory = $false)]
        [bool]$Hidden = $false,

        [Parameter(Mandatory = $false)]
        [System.Xml.XmlElement]$Fields = $null,

        [Parameter(Mandatory = $true)]
        [System.Management.Automation.PSCredential]$Credentials	
    )
    Process {
        $contentType = Get-PnPContentType | Where-Object { $_.Id.StringValue -eq $ContentTypeId }
        if ($contentType.Name -ne $null) {
            Write-Host -ForegroundColor Yellow  "El tipo de contenido '$ContentTypeId' ya existe, se modificará..."            
        }
        else {
            Add-PnPContentType -Name $Name -ContentTypeId $ContentTypeId -Description $Description -Group $Group 
            Write-Host -ForegroundColor Green  "Tipo de contenido '$ContentTypeId' creado"
        }

        if ($LocalizedNames -ne $null) {
            $cxt = Get-PnPContext
            $creden = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($Credentials.UserName, $Credentials.Password)
            $context = New-Object Microsoft.SharePoint.Client.ClientContext($cxt.Url)
            $context.Credentials = $creden
            $web = $context.Web
            $contentTypes = $context.Web.ContentTypes
            $myContentType = $contentTypes.GetById($ContentTypeId)
            $context.Load($web)
            $context.Load($contentTypes)
            $context.Load($myContentType)
            $context.ExecuteQuery()

            $LocalizedNames.GetEnumerator() | % {
                if ($web.Language -eq $_.Key.LCID) {
                    $myContentType.Name = $_.Value
                }
                $myContentType.NameResource.SetValueForUICulture($_.Key, $_.Value)
            }
            $myContentType.Update($true)
            $context.ExecuteQuery()
        }

        if ($Fields) {
            Foreach ($item in $Fields.Add) {			
                if ($item.Required -eq "True") {
                    Add-PnPFieldToContentType -Field $item.InternalName -ContentType $Name -Required
                }
                else {
                    if ($item.Hidden -eq "True") {
                        Add-PnPFieldToContentType -Field $item.InternalName -ContentType $Name -Hidden
                    }
                    else {
                        Add-PnPFieldToContentType -Field $item.InternalName -ContentType $Name 
                    }
                }
                Write-Host -ForegroundColor Green "Campo anyadido al Tipo de Contenido '"$item.InternalName"'"
            }  
        } 
    }
}