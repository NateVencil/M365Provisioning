## Logging Functions ## 
function New-SiteError {
    <#
        .Synopsis 
        This fucntion is called Write-SiteError. This handles much the error logging and messaging for any exceptions and failures in scripts. 
        .Description 
        This function, Get-Template, extracts the site template, page layout, and content identified in accompanying config.json file.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [String]$ErrorMessage,
        [Parameter(Mandatory = $true)]
        [String]$ErrorType,
        [Parameter(Mandatory = $true)]
        [String]$ErrorURL,
        [Parameter(Mandatory = $true)]
        [String]$ErrorSiteTitle,
        [Parameter(Mandatory = $true)]
        [String]$ErrorSiteType,
        [Parameter(Mandatory = $true)]
        [String]$ErrorCSVFilePath

    )
    $ErrorItem = [Ordered] @{
        'URL'          = $ErrorURL
        'SiteTitle'    = $ErrorSiteTitle
        'SiteType'     = $ErrorSiteType
        'ErrorMessage' = $ErrorMessage
        'ErrorType'    = $ErrorType
    }    
    $ErrorItemsObject = New-Object PSObject -Property $ErrorItem    
    $ErrorItemsObject | Select-Object -Property URL, SiteTitle, SiteType, ErrorMessage | Export-Csv $ErrorCSVFilePath -Encoding UTF8 -NoTypeInformation -Append
}


## Site Template Functions ##
function Get-Template {
    <#
        .Synopsis 
        This fucntion is called Get-Template. This function has been created custom for the purposes of the SPO Provisioning Engine.
        .Description 
        This function, Get-Template, extracts the site template, page layout, and content identified in accompanying config.json file.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [String]$TemplateSite,
        [Parameter(Mandatory = $true)]
        [String]$TemplateName,
        [Parameter(Mandatory = $true)]
        [String]$TemplateSettings
    )
    $TemplateName = "$($TemplateName).pnp"

    try {
        Connect-PnPOnline -Url $TemplateSite
        Get-PnPSiteTemplate -ExcludeHandlers ApplicationLifecycleManagement, AuditSettings, ContentTypes, CustomActions, ExtensibilityProviders, Features, Fields, Files, ImageRenditions, Lists, PageContents, Pages, Publishing, RegionalSettings, SearchSettings, SiteFooter, SiteHeader, SitePolicy, SupportedUILanguages, SyntexModels, Tenant, TermGroups, WebApiPermissions, Workflows -Configuration "config.json" -Out $TemplateName -Force -ErrorAction Stop
    }
    catch {
        Write-SiteError -ErrorType TemplateExtractionFail -Message "Error in Extracting Template for $($TemplateSite)"
        $_.Exception.Message
        
    }
}

function Set-Template {
    <#
        .Synopsis 
        This fucntion is called Set-Template. This function has been created custom for the purposes of the SPO Provisioning Engine.
        .Description 
        This function, Set-Template, applies the extracted template to the destination site identified the accompanying site manifest .csv file.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [String]$TargetSiteUrl,
        [Parameter(Mandatory = $true)]
        [String]$TemplateName
    )
    try {
        Connect-PnPOnline -Url $TargetSiteUrl -Interactive
        Invoke-PnPSiteTemplate -Path $TemplateName -ErrorAction Stop
    }
    catch {
        Write-SiteError -ErrorType TemplateApplicationFail -Message "Error in applying $($TemplateName) to $($SiteUrl)"
    }
}

## Site Collection Creation Functions ## 

function New-SiteCollection {
    <#
        .Synopsis 
        This fucntion is called New-SiteCollection. It will create Modern site collections for M365 Group connect or standalone team sites and Communication Sites
        .Description 
        This function, New-SiteCollection, creates site collections using PnP and the type is dictated by what choice the user makes, specifying in a CSV or form. 
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [String]$SiteCollectionType,
        [Parameter(Mandatory = $true)]
        [String]$SiteCollectionAlias,
        [Parameter(Mandatory = $true)]
        [String]$SiteCollectionTitle,
        [Parameter(Mandatory = $true)]
        [String]$SiteCollectionURL

    )
    
    try {
        switch ($SiteCollectionType) {
            TeamSite { 
                Write-Host "Provisioning Team Site for $($SiteCollectionTitle)" -ForegroundColor Magenta
                New-PnPSite -Type TeamSite -Title $SiteCollectionTitle -Alias $SiteCollectionAlias -Wait -ErrorAction Stop
                ; Break 
            }
            CommunicationSite { 
                Write-Host "Provisioning Communication Site for $($SiteCollectionTitle)" -ForegroundColor Magenta
                New-PnPSite -Type CommunicationSite -Title $SiteCollectionTitle -Url $SiteCollectionURL -Wait -ErrorAction Stop 
                ; Break 
            }
            TeamSiteNoGroup {
                Write-Host "Provisioning Team Site - No M365 group for $($SiteCollectionTitle)" -ForegroundColor Magenta
                New-PnPSite -Type TeamSiteWithoutMicrosoft365Group -Title $SiteCollectionTitle -Url $SiteCollectionURL -Wait -ErrorAction Stop
                ; Break 
            }
            Default {
                throw "Invalid Site Type"
            }
        }   
    }
    catch {
        Write-Error $_.Exception.Message
        #Write-SiteError -ErrorType SiteCollectionCreateFail -Message "The Site $($SiteCollectionURL) has failed to provision or attempts to retrieve site information failed."
    }
}

## Hub Site Registration and Association Functions ##

function New-HubSiteRegistration {
    <#
        .Synopsis 
        This fucntion is called 
        .Description 
        This function, 
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [String]$HubSiteURL,
        [Parameter(Mandatory = $true)]
        [String]$HubSiteTitle 
    )
    Write-Host "Registering $($HubSiteURL) as Hub Site $($HubSiteTitle)"
    try {
        $RegisteredAsHub = Get-PnPHubSite -Identity $HubSiteURL -ErrorAction SilentlyContinue
        if ($RegisteredAsHub.SiteId) {
            Write-Host "This site is already registered as a hub within the tenant" -ForegroundColor Yellow
        }
        else {
            $HubRegisterComplete = Register-PnPHubSite -Site $HubSiteURL -ErrorAction Stop 
        }
        if (!$HubRegisterComplete.SiteId) {
            throw "Hub Registration failed"
        }
    }
    catch {
        $_.Exception.Message
        #Write-SiteError -Message "The Site $($HubSiteURL) has failed to be registered as a hub within the Tenant"
    }   
}

function New-HubSiteAssociation {
    <#
        .Synopsis 
        This fucntion is called 
        .Description 
        This function, 
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [String]$HubSiteURL,
        [Parameter(Mandatory = $true)]
        [String]$SpokeSiteURL,
        [Parameter(Mandatory = $true)]
        [String]$HubSiteRegistration
    )

    Write-Host "Associating $($SpokeSiteURL) To Target hub $($HubSiteURL)" -ForegroundColor Cyan
    try {
        if ($HubSiteRegistration -eq "Yes") {
            Write-Host "$($SpokeSiteURL) is already registered as a hub Site. Please manually associate hub to parent hub." -ForegroundColor Yellow
            ##Add-PnPHubToHubAssociation -SourceUrl $NewSiteProperties -TargetUrl $row.Association
            ## Current Get-PnPHubSiteChild only returns regular associations, not hub to hub. Will update script once this funtionality is available. We'll just need to assume they're associated if an exception is not thrown ##
        }
        else {
            Add-PnPHubSiteAssociation -Site $SpokeSiteURL -HubSite $HubSiteURL
            $HubAssociationComplete = Get-PnPHubSiteChild -Identity $HubSiteURL -ErrorAction SilentlyContinue
            if ($HubAssociationComplete -notcontains $SpokeSiteURL) {
                throw "Hub Association failed"
            }	
        }
    }
    catch {
        $_.Exception.Message
        #Write-SiteError -Message "Hub site association for $($SpokeSiteURL) to $($HubSiteURL) failed."
    }	 
    
}

# Add Microsoft Teams and Channels ##
function Add-MicrosoftTeam {
    <#
        .Synopsis 
        This fucntion is called 
        .Description 
        This function, 
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [String]$TeamSiteURL,
        [Parameter(Mandatory = $false)]
        [System.Array]$ChannelNames,
        [Parameter(Mandatory = $false)]
        [System.Array]$PrivateChannelNames
    )
    Write-Host "Adding Microsoft Team to $($TeamSiteURL)" -ForegroundColor Cyan
    $SiteforTeamsConversion = Get-PnPTenantSite -Url $TeamSiteURL
    try {
        Write-Host "The Group ID for $($TeamSiteURL) is $($SiteforTeamsConversion.GroupId)" -ForegroundColor Green
        if ($SiteforTeamsConversion.GroupId) {
            New-PnPTeamsTeam -GroupId $SiteforTeamsConversion.GroupId -ErrorAction Stop
        }
        else {
            throw "Unable to retrieve Group ID"
        }
    }
    catch {
        $_.Exception.Message
        #Write-SiteError -Message "Adding Microsoft Teams for $($TeamSiteURL) has failed"
    }
    Write-Host "Adding Prviate Channels to $($SiteforTeamsConversion.Title)" -ForegroundColor Magenta
    if ($PrivateChannelNames) {
        foreach ($name in $PrivateChannelNames) {
            try {
                Add-PnPTeamsChannel -Team $SiteforTeamsConversion.GroupId.Guid -DisplayName $Name -ChannelType Private -OwnerUPN $SCAAccountUPN -ErrorAction Stop
            }
            catch {
                $_.Exception.Message
            }   
        } 
    }else {
        Write-Host "No private channels for creation" -ForegroundColor Yellow 
    }
    Write-Host "Adding Standard Channels to $($SiteforTeamsConversion.Title)" -ForegroundColor Magenta
    if ($ChannelNames) {
        foreach ($name in $ChannelNames) {
            try {
                Add-PnPTeamsChannel -Team $SiteforTeamsConversion.GroupId.Guid -DisplayName $Name -ChannelType Standard -OwnerUPN $SCAAccountUPN -ErrorAction Stop
            }
            catch {
                $_.Exception.Message
            }   
        } 
    }else {
        Write-Host "No Standard channels for creation" -ForegroundColor Yellow 
    }
}