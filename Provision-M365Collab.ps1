[CmdletBinding()]
Param
(
    [Parameter(Mandatory = $true)]
    [System.String]$SPOSiteCSV,
    [Parameter(Mandatory = $true)]
    [System.String]$SPADMinUrl,
    [Parameter(Mandatory = $true)]
    [System.String]$SCAAccountUPN,
    [Parameter(Mandatory = $true)]
    [System.String]$ErrorCSVFilePath,
    [Parameter(Mandatory = $false)]
    [System.Array]$ChannelNames,
    [Parameter(Mandatory = $false)]
    [System.Array]$PrivateChannelNames,
    [Parameter(Mandatory = $false)]
    [System.String]$TemplateSiteURL,
    [Parameter(Mandatory = $false)]
    [System.String]$SiteTemplateName
)


## Run AzureADPreview in Powershell 7 Workaround ##
if ($PSVersionTable.PSVersion.Major -eq 7) {
    Import-Module -Name AzureADPreview -UseWindowsPowerShell
}
## Import CSV and Validate Values ##
$SiteManifest = Import-Csv $SPOSiteCSV
if (!$SiteManifest) {
    $TerminateDate = $EndDate = Get-Date -Format "MM/dd/yyyy HH:mm"
    Write-Host "Script terminated at $($TerminateDate) - CSV did not import" -ForegroundColor Red
    End
}
## Connect to M365 and terminate if unsuccessful ## 
try {
    Connect-PnPOnline -Interactive -Url $SPADMinUrl 
}
catch {
    $TerminateDate = $EndDate = Get-Date -Format "MM/dd/yyyy HH:mm"
    Write-Host "Script terminated at $($TerminateDate) - Unsuccessful M365 Connection" -ForegroundColor Red
    End
}
## Create PnP Site Templates ##
if ($TemplateSiteURL) {
    try {
        Get-Template -TemplateSite $TemplateSiteURL -TemplateName $SiteTemplateName
    }
    catch {
        $_.Exception.Message
    }  
}


## Import Provisioning Functions for M365 artifact creation ##
. ".\ProvisioningFunctions.ps1"

foreach ($site in $SiteManifest) {
    $SiteExists = $null
    $NewSiteProperties = $null

    ## Create Site Collection for Current Site ## 
    Write-Host "Validating URL is available" -ForegroundColor Magenta
    $SiteExists = Get-PnPTenantSite -Url $site.URL -ErrorAction SilentlyContinue
    if (!$SiteExists) {
        try {
            $NewSiteProperties = New-SiteCollection -SiteCollectionType $site.SiteType -SiteCollectionAlias $site.Alias -SiteCollectionTitle $site.Title -SiteCollectionURL $site.URL -ErrorAction Stop
        }
        catch {
            $_.Exception.Message
            New-SiteError -ErrorType "SiteCreationError" -ErrorURL $site.URL -ErrorSiteTitle $site.Title -ErrorSiteType $site.SiteType -ErrorMessage "$($site.Title) has failed to provision - Please check provisioning file and try again" -ErrorCSVFilePath $ErrorCSVFilePath
        }
    }
    else {
        Write-Host "This site has already been provisioned within the Tenant. Skipping to next site" -ForegroundColor Yellow
        continue
    }
  
    ## Register Current Site as Hub ## 
    if ($site.HubSite -eq "Yes") {
        Write-Host "Registering $($site.Title) as hub site within tenant" -ForegroundColor Magenta
        try {
            New-HubSiteRegistration -HubSiteURL $NewSiteProperties -HubSiteTitle $site.Title
        }
        catch {
            $_.Exception.Message
            New-SiteError -ErrorType "HubRegisterError" -ErrorURL $site.URL -ErrorSiteTitle $site.Title -ErrorSiteType $site.SiteType -ErrorMessage "$($site.Title) has failed to register as Hub - Please check provisioning file and try again" -ErrorCSVFilePath $ErrorCSVFilePath
        }
    }

    ## Associate Current Site with specified Hub Site URL ##
    if ($site.Association) {
        Write-Host "Associating $($NewSiteProperties) to hub $($site.Association)" -ForegroundColor Magenta
        try {
            New-HubSiteAssociation -SpokeSiteURL $NewSiteProperties -HubSiteURL $site.Association
        }
        catch {
            $_.Exception.Message
            New-SiteError -ErrorType "HubAssociationError" -ErrorURL $site.URL -ErrorSiteTitle $site.Title -ErrorSiteType $site.SiteType -ErrorMessage "$($site.Title) has failed to associate with target hub $($site.Association) - Please check provisioning file and try again" -ErrorCSVFilePath $ErrorCSVFilePath
        }
    }

    ## Create Microsoft Team for Current site with specified Channels and Private Channels ## 
    if ($site.CreateTeam) {
        Write-Host "Creating Microsoft Team for $($site.Title)" -ForegroundColor Magenta
        try {
            Add-MicrosoftTeam -TeamSiteURL $NewSiteProperties -ChannelNames $ChannelNames -PrivateChannelNames $PrivateChannelNames
        }
        catch {
            $_.Exception.Message
            New-SiteError -ErrorType TeamCreationError -ErrorURL $site.URL -ErrorSiteTitle $site.Title -ErrorSiteType $site.SiteType -ErrorMessage "$($site.Title) has failed to create a Microsoft Team - Please Check provisioning file and try again" -ErrorCSVFilePath $ErrorCSVFilePath
        }
    }

}

Disconnect-PnPOnline

$EndDate = Get-Date -Format "MM/dd/yyyy HH:mm"

Write-Host "Site Provisioning Script successfully completed at $($EndDate)" -ForegroundColor Green