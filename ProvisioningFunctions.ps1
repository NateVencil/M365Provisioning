## Logging Functions ## 
function New-SiteError {
        <#
        .SYNOPSIS
            Write-SiteError function logs and handles errors occurring during script execution.
        
        .DESCRIPTION
            This function logs errors and exceptions encountered during script execution. It takes various parameters such as ErrorMessage, ErrorType, ErrorURL, ErrorSiteTitle, ErrorSiteType, and ErrorCSVFilePath to create a structured error log.
        
        .PARAMETER ErrorMessage
            Mandatory parameter specifying the error message encountered during script execution.
        
        .PARAMETER ErrorType
            Mandatory parameter specifying the type of error encountered during script execution.
        
        .PARAMETER ErrorURL
            Mandatory parameter specifying the URL related to the error encountered during script execution.
        
        .PARAMETER ErrorSiteTitle
            Mandatory parameter specifying the title of the site where the error occurred during script execution.
        
        .PARAMETER ErrorSiteType
            Mandatory parameter specifying the type of the site where the error occurred during script execution.
        
        .PARAMETER ErrorCSVFilePath
            Mandatory parameter specifying the file path where the error log will be exported in CSV format.
        
        .EXAMPLE
            New-SiteError -ErrorMessage "Page not found" -ErrorType "404" -ErrorURL "http://example.com" -ErrorSiteTitle "Example Site" -ErrorSiteType "Intranet" -ErrorCSVFilePath "C:\ErrorLog.csv"
            
            This example demonstrates how to use the New-SiteError function to log an error encountered during script execution.
        
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
    .SYNOPSIS
        Get-Template function extracts site template, page layout, and content based on settings specified in the config.json file.
    
    .DESCRIPTION
        This function, Get-Template, is custom-built for the SPO Provisioning Engine. It extracts the specified site template, page layout, and content based on the settings provided in the accompanying config.json file.
    
    .PARAMETER TemplateSite
        Mandatory parameter specifying the URL of the site from which the template will be extracted.
    
    .PARAMETER TemplateName
        Mandatory parameter specifying the name of the template file to be created.
    
    .PARAMETER TemplateSettings
        Mandatory parameter specifying the settings file (config.json) that contains configuration details for the template extraction process.
    
    .EXAMPLE
        Get-Template -TemplateSite "https://contoso.sharepoint.com/sites/TemplateSite" -TemplateName "Template1" -TemplateSettings "config.json"
        
        This example demonstrates how to use the Get-Template function to extract a site template, page layout, and content from a SharePoint site specified in the TemplateSite parameter. The extracted template will be saved as "Template1.pnp" based on the settings provided in the "config.json" file.
    
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
        $_.Exception.Message
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
