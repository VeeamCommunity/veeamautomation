Set-ExecutionPolicy Bypass -Scope Process -Force

#ignore SSL warnings
if (-not("dummy" -as [type])) {
    add-type -TypeDefinition @"
using System;
using System.Net;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;

public static class Dummy {
    public static bool ReturnTrue(object sender,
        X509Certificate certificate,
        X509Chain chain,
        SslPolicyErrors sslPolicyErrors) { return true; }

    public static RemoteCertificateValidationCallback GetDelegate() {
        return new RemoteCertificateValidationCallback(Dummy.ReturnTrue);
    }
}
"@
}
[System.Net.ServicePointManager]::ServerCertificateValidationCallback = [dummy]::GetDelegate()

function Say($r,$m){ Write-Host "${r}: $m" }

function Group-By {
    param($Items, $Key)

    $m = @{}
    foreach ($i in $Items) {
        $k = $i.$Key
        if ($null -eq $k) { continue }   # skip null keys

        if (-not $m.ContainsKey($k)) {
            $m[$k] = @()
        }
        $m[$k] += $i
    }
    $m
}

function Index-By {
    param($Items, $Key)

    $m = @{}
    foreach ($i in $Items) {
        $k = $i.$Key
        if ($null -eq $k) { continue }   # skip null keys

        $m[$k] = $i
    }
    $m
}


function Get-PagedResults {
    param (
        [string]$Url,
        [hashtable]$Headers
    )

    $out = [System.Collections.ArrayList]::new()
    $offset = 0

    $separator = if ($Url -match '\?') { '&' } else { '?' }

    do {
        $pagedUrl = "$Url${separator}limit=500&offset=$offset"
        $r = Invoke-RestMethod -Uri $pagedUrl -Headers $Headers

        if ($r.data) {
            [void]$out.AddRange($r.data)
        }

        $offset += 500
    }
    while ($r.meta.pagingInfo.count -ge 500)

    return ,$out
}

function Get-PagedResultsForParents {
    param (
        [string]$BaseUrl,
        [string]$Template,   # 'infrastructure/vb365Servers/{0}/backupRepositories'
        [array]$Parents,
        [string]$IdField,
        [hashtable]$Headers
    )

    $out = [System.Collections.ArrayList]::new()

    foreach ($p in $Parents) {
        $url = $BaseUrl + ($Template -f $p.$IdField)
        $data = Get-PagedResults -Url $url -Headers $Headers
        [void]$out.AddRange($data)
    }

    return ,$out
}

$regions = @("US", "NL", "NO")

#VSPC token and header
$USToken = '<API_TOKEN>'
$NLToken = '<API_TOKEN>'
$NOToken = '<API_TOKEN>'
$USHeaders = @{
    Authorization="Bearer $USToken"
    "x-client-version"="3.6.1"
}
$NLHeaders = @{
    Authorization="Bearer $NLToken"
    "x-client-version"="3.6.1"
}
$NOHeaders = @{
    Authorization="Bearer $NOToken"
    "x-client-version"="3.6.1"
}
$headers = @($USHeaders, $NLHeaders, $NOHeaders)


#VSPC base url
$USBaseURL = "https://vspc01:1280/api/v3/"
$NLBaseURL = "https://vspc02:1280/api/v3/"
$NOBaseURL = "https://vspc03:1280/api/v3/"
$baseURL = @($USBaseURL, $NLBaseURL, $NOBaseURL)

#File path for data export
$USFilePath = $env:USERPROFILE + "\Downloads\US_ConsumptionReport_" + (Get-Date -Format "yyyy-MM-dd_HH-mm-ss") + ".xlsx"
$NLFilePath = $env:USERPROFILE + "\Downloads\NL_ConsumptionReport_" + (Get-Date -Format "yyyy-MM-dd_HH-mm-ss") + ".xlsx"
$NOFilePath = $env:USERPROFILE + "\Downloads\NO_ConsumptionReport_" + (Get-Date -Format "yyyy-MM-dd_HH-mm-ss") + ".xlsx"
$FilePath = @($USFilePath, $NLFilePath, $NOFilePath)

$regions= @(
    @{ Name='US'; Base=$USBaseURL; Headers=$USHeaders; FilePath=$USFilePath},
    @{ Name='NL'; Base=$NLBaseURL; Headers=$NLHeaders; FilePath=$NLFilePath},
    @{ Name='NO'; Base=$NOBaseURL; Headers=$NOHeaders; FilePath=$NOFilePath}
)

#Check for/install the ImportExcel module
if (Get-Module -ListAvailable -Name ImportExcel) {
    Import-Module ImportExcel
} 
else {
    Install-PackageProvider NuGet -Force;
    Set-PSRepository PSGallery -InstallationPolicy Trusted
    Install-Module ImportExcel -Repository PSGallery

    Import-Module ImportExcel
}

foreach($r in $regions) {
    Say $r.Name "Getting active company list from VSPC..."
    #get list of active VSPC companies
    $companies = Get-PagedResults -Url ($r.Base + 'organizations/companies') -Headers $r.Headers

    #get VBR CC host information
    Say $r.Name "Getting VCC host information..."
    $CChosts = Get-PagedResults -Url ($r.Base + 'infrastructure/sites') -Headers $r.Headers

    #get VCC tenant information
    Say $r.Name "Getting tenant information from VBR servers..."
    $tenants = Get-PagedResults -Url ($r.Base + 'infrastructure/sites/tenants') -Headers $r.Headers

    #get backup repositories
    Say $r.Name "Getting VCC repositories..."
    $backupRepos = Get-PagedResultsForParents `
        -BaseUrl $r.Base `
        -Template 'infrastructure/backupServers/{0}/repositories?expand=BackupRepositoryInfo' `
        -Parents $CChosts `
        -IdField 'siteUid' `
        -Headers $r.Headers

    #get license reporting info for all customers
    Say $r.Name "Getting list of licenses in use..."
    $licensingInfo = Get-PagedResults -Url ($r.Base + 'licensing/usage/organizations') -Headers $r.Headers

    #get VBR licensing so we can display license owner later
    $licensingVBRInfo = Get-PagedResults -Url ($r.Base + 'licensing/backupServers') -Headers $r.Headers

    #get usage metrics for all companies (for things like Insider Protection)
    Say $r.Name "Getting basic company usage information..."
    $allResources = Get-PagedResults -Url ($r.Base + 'organizations/companies/usage') -Headers $r.Headers

    #get BaaS resources
    Say $r.Name "Getting all BaaS resources..."
    $baaSResourceItems = Get-PagedResults -Url ($r.Base + 'infrastructure/sites/tenants/backupResources') -Headers $r.Headers

    #get BaaS resource usage
    Say $r.Name "Getting all BaaS usage..."
    $baaSUsageItems = Get-PagedResults -Url ($r.Base + 'infrastructure/sites/tenants/backupResources/usage') -Headers $r.Headers

    #get DRaaS native resource usage
    Say $r.Name "Getting all DRaaS native usage..."
    $draaSNativeUsageItems = Get-PagedResults -Url ($r.Base + 'infrastructure/sites/tenants/replicationResources/usage') -Headers $r.Headers

    #get DRaaS VCD resource usage
    Say $r.Name "Getting all DRaaS VCD usage..."
    $draaSVCDUsageItems = Get-PagedResults -Url ($r.Base + 'infrastructure/sites/tenants/vcdReplicationResources/usage') -Headers $r.Headers

    #get vb365 server list
    Say $r.Name "Getting VB365 server list..."
    $vb365Servers = Get-PagedResults -Url ($r.Base + 'infrastructure/vb365Servers') -Headers $r.Headers

    #get vb365 proxies
    Say $r.Name "Getting VB365 proxy list..."
    $vb365Proxies = Get-PagedResultsForParents `
        -BaseUrl $r.Base `
        -Template 'infrastructure/vb365Servers/{0}/backupProxies' `
        -Parents $vb365Servers `
        -IdField 'instanceUid' `
        -Headers $r.Headers

    #get vb365 proxy pools
    Say $r.Name "Getting VB365 proxy pool list..."
    $vb365ProxyPools = Get-PagedResultsForParents `
        -BaseUrl $r.Base `
        -Template 'infrastructure/vb365Servers/{0}/backupProxyPools' `
        -Parents $vb365Servers `
        -IdField 'instanceUid' `
        -Headers $r.Headers

    #get vb365 repositories
    Say $r.Name "Getting VB365 repository list..."
    $vb365Repos = Get-PagedResultsForParents `
        -BaseUrl $r.Base `
        -Template 'infrastructure/vb365Servers/{0}/backupRepositories' `
        -Parents $vb365Servers `
        -IdField 'instanceUid' `
        -Headers $r.Headers

    #get vb365 resources
    Say $r.Name "Getting VB365 resource information..."
    $vb365Resources = Get-PagedResults -Url ($r.Base + 'organizations/companies/hostedResources/vb365') -Headers $r.Headers

    #get M365 backup resources
    Say $r.Name "Getting all VB365 backup resource information..."
    $vb365BackupResourceList = Get-PagedResults -Url ($r.Base + 'organizations/companies/hostedResources/vb365/backupResources') -Headers $r.Headers

    Say $r.Name "Building hash tables for collected data..."
    $CCHostByUid           = Index-By $CChosts 'siteUid'
    $BackupRepoByUid       = Index-By $backupRepos 'instanceUid'
    $VB365ResourceByUid    = Index-By $vb365Resources 'instanceUid'
    $VB365ProxyByUid       = Index-By $vb365Proxies 'instanceUid'
    $VB365ProxyPoolByUid   = Index-By $vb365ProxyPools 'instanceUid'
    $VB365RepoByUid        = Index-By $vb365Repos 'instanceUid'
    $BaaSResourceByUid     = Index-By $baaSResourceItems 'instanceUid'
    $LicenseOwnerByServerUid = Index-By $licensingVBRInfo 'backupServerUid'

    $LicensingByCompany    = Group-By $licensingInfo 'organizationUid'
    $UsageByCompany        = Group-By $allResources 'companyUid'
    $TenantsByCompany      = Group-By $tenants 'assignedForCompany'
    $VB365BackupByCompany  = Group-By $vb365BackupResourceList 'companyUid'
    $BaaSUsageByTenant     = Group-By $baaSUsageItems 'tenantUid'
    $DraaSNativeByTenant   = Group-By $draaSNativeUsageItems 'tenantUid'
    $DraaSVCDByTenant      = Group-By $draaSVCDUsageItems 'tenantUid'
    $LicensedUsersByCompany = Group-By $licensingInfo 'organizationUid'
    $TenantsUsingM365 = Group-By $vb365BackupResourceList 'tenantUid'

    Say $r.Name "Compiling full resource usage information..."
    #get backup, replication, and M365 information and save to array
    $UsageResources  = [System.Collections.Generic.List[object]]::new()
    $BaaSResources   = [System.Collections.Generic.List[object]]::new()
    $DRaaSResources  = [System.Collections.Generic.List[object]]::new()
    $NoResources     = [System.Collections.Generic.List[object]]::new()
    $vb365Backups    = [System.Collections.Generic.List[object]]::new()
    $Licenses        = [System.Collections.Generic.List[object]]::new()
    $ResourceInfo = ""

    $counterMap = @{
        Vb365ProtectedUsers="protectedUsers"
        Vb365ProtectedTeams="protectedTeams"
        Vb365ProtectedSites="protectedSites"
        Vb365BackupSize="M365BackupSize"
        CloudInsiderProtectionBackupSize="RIPUsage"
        VmCloudBackups="backedUpVMs"
        ServerCloudBackups="backedUpVMs"
        WorkstationCloudBackups="backedUpVMs"
        VmCloudReplicas="replicatedVMs"
        CloudPerformanceTierBackupSize="performanceTierUsage"
        CloudCapacityTierBackupSize="capacityTierUsage"
        CloudArchiveTierBackupSize="archiveTierUsage"
        CloudTotalUsage="totalBackupUsage"
        VmCloudReplicaStorageUsage="replicationStorageUsed"
    }
    $tbCounters = @{
        M365BackupSize = 1
        RIPUsage = 1
        performanceTierUsage = 1
        capacityTierUsage = 1
        archiveTierUsage = 1
        totalBackupUsage = 1
        replicationStorageUsed = 1
    }

    $count = 1
    foreach($company in $companies) {
        Say $r.Name "Collecting information for tenant ($count/$($companies.Count)) $($company.name) - $($company.instanceUid)..."

        $counterAcc = @{
            protectedUsers = 0
            protectedTeams = 0
            protectedSites = 0
            M365BackupSize = 0
            RIPUsage = 0
            backedUpVMs = 0
            replicatedVMs = 0
            performanceTierUsage = 0
            capacityTierUsage = 0
            archiveTierUsage = 0
            totalBackupUsage = 0
            replicationStorageUsed = 0
        }

        #get license consumption
        if ($LicensingByCompany.ContainsKey($company.instanceUid)) {
            foreach($license in $LicensingByCompany[$company.instanceUid]) {
                foreach($server in $license.servers) {
                    foreach($workload in $server.workloads) {
                        if($server.serverType -ne 'VB365' -and $server.serverType -ne 'CloudConnect') {
                            # Get who owns the license
                            $licenseOwner = $LicenseOwnerByServerUid[$server.serverUid]
                        }
                        else {
                            $licenseOwner = "CyberFortress"
                        }

                        $LicenseInfo = [PSCustomObject]@{
                            'Company' = $company.Name
                            'ID' = $company.instanceUid
                            'Type' = $server.serverType
                            'License Owner' = $licenseOwner
                            'Instance Type' = $workload.description
                            'Count' = $workload.usedCount
                            'PPU' = $workload.workloadsByPlatform[0].weight
                            'Points Used' = $workload.usedUnits
                        }


                        $Licenses.Add($LicenseInfo)
                    }
                }
            }
        }

        # Process VB365 backup resources (join VB365 server, proxy, and repository info)
        if ($VB365BackupByCompany.ContainsKey($company.instanceUid)) {
            foreach ($backupResource in $VB365BackupByCompany[$company.instanceUid]) {
                $resource = $VB365ResourceByUid[$backupResource.vb365ResourceUid]
                $vb365Proxy = ""
                $vb365ProxyPool = ""
                $vb365Retention = ""
                $vb365Repo = ""
                $vb365ObjectRepo = $false
                $vb365CopyRepo = $false
                $vb365PrimaryRepo = $false

                if ($backupResource.proxyUid -and $VB365ProxyByUid.ContainsKey($backupResource.proxyUid)) {
                    $vb365Proxy = $VB365ProxyByUid[$backupResource.proxyUid].hostName
                }

                # Proxy pool logic
                if ($backupResource.proxyPoolUid -and $VB365ProxyPoolByUid.ContainsKey($backupResource.proxyPoolUid)) {
                    $vb365ProxyPool = $VB365ProxyPoolByUid[$backupResource.proxyPoolUid].name
                }


                $repo = $VB365RepoByUid[$backupResource.repositoryUid]
                $vb365Repo = $repo.name
                $vb365Retention = $repo.retentionType
                $p = $repo.yearlyRetentionPeriod
                $t = $repo.retentionPeriodType

                if ($p -eq "Keep") {
                    $vb365RetentionPeriod = "Forever"
                }
                elseif ($t -in @("Yearly","Monthly","Daily")) {
                    $vb365RetentionPeriod = (
                        $p -replace '^(Year|Month|Day)s?(\d+)$', '$2 $1'
                    )
                }
                else {
                    $vb365RetentionPeriod = $t
                }

                $vb365Encrypted = $false
                if($null -ne $repo.encryptionKeyId) {
                    $vb365Encrypted = $true
                }


                $vb365ObjectRepo = $repo.isObjectStorageRepository
                # Divide by 1TB for TB units (1TB = 1,099,511,627,776 bytes)
                $M365BackupSize = $repo.usedSpaceBytes / 1TB
                $vb365PrimaryRepo = $repo.isAvailableForBackupJob
                $vb365CopyRepo = $repo.isAvailableForCopyJob

                $ResourceInfo = [PSCustomObject]@{
                    'Company'               = $company.name
                    'ID'                    = $company.instanceUid
                    'Server'                = $resource.friendlyName
                    'Proxy'                 = $vb365Proxy
                    'Proxy Pool'            = $vb365ProxyPool
                    'Repository'            = $vb365Repo
                    'Retention Type'        = $vb365Retention
                    'Retention Period'      = $vb365RetentionPeriod
                    'Used Space (TB)'       = $M365BackupSize
                    'Encrypted'             = $vb365Encrypted
                    'Object Repository'     = $vb365ObjectRepo
                    'Primary Repository'    = $vb365PrimaryRepo
                    'Copy Repository'       = $vb365CopyRepo
                }
                $vb365Backups.Add($ResourceInfo)
            }
        }

        #get tenant usage
        $RIPUsage = 0
        if ($UsageByCompany.ContainsKey($company.instanceUid)) {

            $protectedUsers = $protectedTeams = $protectedSites = $M365BackupSize = $licensedUsers = 0
            $backedUpVMs = $replicatedVMs = 0
            $performanceTierUsage = $capacityTierUsage = $archiveTierUsage = $totalBackupUsage = 0
            $replicationStorageUsed = 0

            foreach($resource in $UsageByCompany[$company.instanceUid]) {
                foreach($c in $resource.counters){
                    if($counterMap.ContainsKey($c.type)){
                        $k = $counterMap[$c.type]
                        if($tbCounters.ContainsKey($k)){
                            $counterAcc[$k] = $c.value / 1TB
                        }
                        elseif($k -eq "backedUpVMs"){
                            $counterAcc[$k] += $c.value
                        }
                        else{
                            $counterAcc[$k] = $c.value
                        }
                    }
                }
            }

            if($counterAcc.performanceTierUsage -le 0 -and $counterAcc.capacityTierUsage -le 0 -and $counterAcc.archiveTierUsage -le 0) {
                $performanceTierUsage = $counterAcc.totalBackupUsage
            }

            if ($LicensedUsersByCompany.ContainsKey($company.instanceUid)) {
                $licensedUsers = 0
                foreach ($lic in $LicensedUsersByCompany[$company.instanceUid]) {
                    foreach ($srv in $lic.servers) {
                        foreach ($wl in $srv.workloads) {
                            $licensedUsers += $wl.usedCount
                        }
                    }
                }
            }

            if ($TenantsByCompany.ContainsKey($company.instanceUid)) {
                foreach ($tenant in $TenantsByCompany[$company.instanceUid]) {
                    $enabled = if ($tenant.isEnabled) { 'Yes' } else { 'No' }

                    $RIPEnabled = if ($tenant.isBackupProtectionEnabled) { 'Yes' } else { 'No' }
                    $RIPDays    = if ($tenant.isBackupProtectionEnabled) {
                        $tenant.backupProtectionPeriod
                    } else {
                        0
                    }

                    $ResourceInfo = [PSCustomObject]@{
                        'Company' = $company.Name
                        'Tenant Name' = $tenant.Name
                        'Description' = $tenant.description
                        'Enabled' = $enabled
                        'ID' = $company.instanceUid
                        # Insider protection information
                        'Insider Protection Enabled' = $RIPEnabled
                        'Insider Protection Days' = $RIPDays
                        'Insider Protection Usage' = $counterAcc.RIPUsage
                        # BaaS/DRaaS basic information
                        'Backed Up VMs' = $counterAcc.backedUpVMs
                        'Replicated VMs' = $counterAcc.replicatedVMs
                        'Performance Tier Usage (TB)' = $performanceTierUsage
                        'Capacity Tier Usage (TB)' = $capacityTierUsage
                        'Archive Tier Usage (TB)' = $archiveTierUsage
                        'Total Backup Storage Usage (TB)' = $totalBackupUsage
                        # Replication usage
                        'Replication Storage Used (TB)' = $replicationStorageUsed
                        # M365 usage
                        'Licensed Users' = $licensedUsers
                        'Protected Users' = $counterAcc.protectedUsers
                        'Protected Sites' = $counterAcc.protectedSites
                        'Protected Teams' = $counterAcc.protectedTeams
                        'M365 Backup Storage Usage (TB)' = $counterAcc.M365BackupSize
                    }
                    $UsageResources.Add($ResourceInfo)

                    # Get more detailed BaaS, DRaaS usage/inventory breakdowns
                    #get used backup resources
                    if($tenant.isBackupResourcesEnabled) {
                        Say $r.Name "Getting BaaS usage..."

                        if ($BaaSUsageByTenant.ContainsKey($tenant.instanceUid)) {
                            foreach ($usageItem in $BaaSUsageByTenant[$tenant.instanceUid]) {
                                $resource = $BaaSResourceByUid[$usageItem.backupResourceUid]

                                $CChost = $CCHostByUid[$resource.siteUid]

                                #get parent repository information
                                $repo = $BackupRepoByUid[$resource.repositoryUid]

                                $hotUsed = $capacityUsed = $archiveUsed = 0
                                $hotUsed = $usageItem.performanceTierUsage / 1TB
                                $capacityUsed = $usageItem.capacityTierUsage / 1TB
                                $archiveUsed = $usageItem.archiveTierUsage / 1TB
                                $totalUsed = $usageItem.usedStorageQuota / 1TB
                                if($hotUsed -le 0 -and $capacityUsed -le 0 -and $archiveUsed -le 0) {
                                    $hotUsed = $totalUsed
                                }

                                $ResourceInfo = [PSCustomObject]@{
                                    'Company' = $company.Name
                                    'Tenant Name' = $tenant.Name
                                    'Description' = $tenant.description
                                    'Enabled' = $enabled
                                    'ID' = $company.instanceUid
                                    'Last Active' = $tenant.lastActive
                                    'VBR Host' = $CChost.siteName
                                    'Repository' = $repo.name
                                    'Type' = $repo._embedded.type
                                    'Repository Host' = $repo._embedded.hostName
                                    'Hot Storage Used (TB)' = $hotUsed
                                    'Capacity Tier Used (TB)' = $capacityUsed
                                    'Archive Tier Used (TB)' = $archiveUsed
                                    'Total Storage Used (TB)' = $totalUsed
                                    'Quota (TB)' = $usageItem.storageQuota / 1TB
                                    'Percent Used' = $usageItem.usedStorageQuota / $usageItem.storageQuota * 100
                                }
                                $BaaSResources.Add($ResourceInfo)
                            }
                        }
                    }

                    #get used DRaaS resources
                    if($tenant.isNativeReplicationResourcesEnabled) {
                        Say $r.Name "Getting DRaaS usage..."

                        $RepUsage = $DraaSNativeByTenant[$tenant.instanceUid]

                        foreach($repResource in $RepUsage) {
                            $vCPUsConsumed = if ($repResource.vCPUsConsumed) { 
                                $repResource.vCPUsConsumed 
                            } else { 
                                0 
                            }

                            $CChost = $CCHostByUid[$repResource.siteUid]
                            
                            $ResourceInfo = [PSCustomObject]@{
                                'Company' = $company.Name
                                'Tenant Name' = $tenant.Name
                                'Description' = $tenant.description
                                'Enabled' = $enabled
                                'ID' = $company.instanceUid
                                'Last Active' = $tenant.lastActive
                                'VBR Host' = $CChost.siteName
                                'Replic vCPUs' = $vCPUsConsumed
                                'Replic Memory Used (GB)' = $repResource.memoryUsage / 1GB
                                'Replic Storage Used (TB)' = $repResource.storageUsage / 1TB
                            }

                            $DRaaSResources.Add($ResourceInfo)
                        }
                    }

                    #get used VCD DRaaS resources
                    if($tenant.isVcdReplicationResourcesEnabled) {
                        Say $r.Name "Getting VCD DRaaS usage..."

                        $VCDRepUsage = $DraaSVCDByTenant[$tenant.instanceUid]

                        foreach($repResource in $VCDRepUsage) {
                            $vCPUsConsumed = if ($repResource.vCPUsConsumed) { 
                                $repResource.vCPUsConsumed 
                            } else { 
                                0 
                            }

                            $CChost = $CCHostByUid[$repResource.siteUid]
                            
                            $ResourceInfo = [PSCustomObject]@{
                                'Company' = $company.Name
                                'Tenant Name' = $tenant.Name
                                'Description' = $tenant.description
                                'Enabled' = $enabled
                                'ID' = $company.instanceUid
                                'Last Active' = $tenant.lastActive
                                'VBR Host' = $CChost.siteName
                                'Replic vCPUs' = $vCPUsConsumed
                                'Replic Memory Used (GB)' = $repResource.memoryUsage / 1GB
                                'Replic Storage Used (TB)' = $repResource.storageUsage / 1TB
                            }

                            $DRaaSResources.Add($ResourceInfo)
                        }
                    }

                    $tenantUsesM365 = $TenantsUsingM365.ContainsKey($tenant.instanceUid)

                    #list tenants with no resources assigned
                    if(-not $tenantUsesM365 -and !$tenant.isBackupResourcesEnabled -and !$tenant.isNativeReplicationResourcesEnabled -and !$tenant.isVcdReplicationResourcesEnabled) {
                        $ResourceInfo = [PSCustomObject]@{
                            'Company' = $company.Name
                            'Tenant Name' = $tenant.Name
                            'Description' = $tenant.description
                            'Enabled' = $enabled
                            'ID' = $company.instanceUid
                            'Last Active' = $tenant.lastActive
                            'VBR Host' = $tenant.siteName
                        }

                        $NoResources.Add($ResourceInfo)
                    }
                }
            }
        }
        $count++
    }

    Say $r.Name "Saving data to XLSX..."
    #export usage information to CSV and save to downloads folder
    $Licenses | Export-Excel $r.FilePath -Autosize -TableName Licensing -WorksheetName Licensing
    $UsageResources | Export-Excel $r.FilePath -Autosize -TableName Usage -WorksheetName Usage
    $BaaSResources | Export-Excel $r.FilePath -Autosize -TableName BaaSResources -WorksheetName "BaaS Detailed"
    $DRaaSResources | Export-Excel $r.FilePath -Autosize -TableName DRaaSResources -WorksheetName "DRaaS Detailed"
    $vb365Backups | Export-Excel $r.FilePath -Autosize -TableName "VB365_Repos" -WorksheetName "VB365 Repository Information"
    $NoResources | Export-Excel $r.FilePath -Autosize -TableName "No_Resources" -WorksheetName "No Resources"
}

foreach($r in $regions) {
    Say $r.Name "Report available under $($r.FilePath)"
}

#tell user file location and exit when any key is pressed
Write-Host -NoNewLine 'Press any key to exit...';
$null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown');
