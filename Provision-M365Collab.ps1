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
    [Parameter(Mandatory = $true)]
    [System.String]$TemplateSiteURL,
    [Parameter(Mandatory = $true)]
    [System.String]$SiteTemplateName
)

Begin {
    ## Run AzureADPreview in Powershell 7 Workaround ##
    Import-Module -Name AzureADPreview -UseWindowsPowerShell

    $SiteManifest = Import-Csv $SPOSiteCSV

    Connect-PnPOnline -Interactive -Url $SPADMinUrl

    ## Create PnP Site Templates ##
    try {
        Get-Template -TemplateSite $TemplateSiteURL -TemplateName $SiteTemplateName
    }
    catch {
        $_.Exception.Message
    }
}

Process {
    ## Import Common Functions for M365 artifact creation ##
    . ".\ProvisioningFunctions.ps1"

    foreach ($site in $SiteManifest) {

        ## Create Site Collection for Current Site ## 
        Write-Host "Validating URL is available" -ForegroundColor Yellow
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
  
        ## Register Current Site as Hub ## 
        if ($site.HubSite -eq "Yes") {
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
            try {
                Add-MicrosoftTeam -TeamSiteURL $NewSiteProperties -ChannelNames $ChannelNames -PrivateChannelNames $PrivateChannelNames
            }
            catch {
                $_.Exception.Message
                New-SiteError -ErrorType TeamCreationError -ErrorURL $site.URL -ErrorSiteTitle $site.Title -ErrorSiteType $site.SiteType -ErrorMessage "$($site.Title) has failed to create a Microsoft Team - Please Check provisioning file and try again" -ErrorCSVFilePath $ErrorCSVFilePath
            }
        }

    }
}