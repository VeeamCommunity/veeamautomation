Set-ExecutionPolicy Bypass -Scope Process -Force

#TODO: Code cleanup/optimization
#TODO: Debug S3-customers and 0 usage M365 customers
#TODO: Licensed users on M365 tab
#TODO: Fix "Keep Years" M365 retention string so it says "Forever"
#TODO: Add location column to tabs
# Ex: 
    #get all locations
    #$string = $regions[$i] + ": Getting list of all company locations..."
    #$moreRecords = $true
    #$locations = [System.Collections.ArrayList]::new()
    #$offset = 0
    #do {
    #    $sitesURL = $baseURL[$i] + 'organizations/{organization_uid}/locations?limit=500&offset=' + $offset
    #    $results = Invoke-RestMethod -Uri $sitesURL -Method GET -Headers $headers[$i]
    #    $sites.AddRange($results.data)
    #
    #    if($results.meta.pagingInfo.Count -ge 500) {
    #        $offset += 500
    #    }
    #    else {
    #        $moreRecords = $false
    #    }
    #} while($moreRecords)

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

$regions = @("US","NL")

#VSPC token and header
$USToken = '<API_TOKEN>'
$NLToken = '<API_TOKEN>'
$USHeaders = @{
    Authorization="Bearer $USToken"
    "x-client-version"="3.5.1"
}
$NLHeaders = @{
    Authorization="Bearer $NLToken"
    "x-client-version"="3.5.1"
}
$headers = @($USHeaders, $NLHeaders)


#VSPC base url
$BaseURL1 = "https://vspc01:1280/api/v3/"
$BaseURL2 = "https://vspc02:1280/api/v3/"
$baseURL = @($BaseURL2, $BaseURL1)

#File path for data export
$USFilePath = $env:USERPROFILE + "\Downloads\US_ConsumptionReport_" + (Get-Date -Format "yyyy-MM-dd_HH-mm-ss") + ".xlsx"
$NLFilePath = $env:USERPROFILE + "\Downloads\NL_ConsumptionReport_" + (Get-Date -Format "yyyy-MM-dd_HH-mm-ss") + ".xlsx"
$FilePath = @($USFilePath, $NLFilePath)

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

for($i=0; $i -lt $regions.Count; $i++) {
    $string =  $regions[$i]+ ": Getting active company list from VSPC..."
    Write-Host $string
    #get list of active VSPC companies
    $companies = [System.Collections.ArrayList]::new()
    $moreRecords = $true
    $offset = 0
    do {
        $companyURL = $baseURL[$i] + 'organizations/companies?filter=[{"property":"status","operation":"notEquals","collation":"ignorecase","value":"Deleted"}]&limit=500&offset=' + $offset
        $results = Invoke-RestMethod -Uri $companyURL -Method GET -Headers $headers[$i]
        $companies.AddRange($results.data)
        if($results.meta.pagingInfo.count -ge 500) {
            $offset += 500
        }
        else {
            $moreRecords = $false
        }
    } while($moreRecords)


    $string = $regions[$i] + ": Getting tenant information from VBR servers..."
    Write-Host $string
    #get list of tenants on hosts (to get VBR name)
    $tenants = [System.Collections.ArrayList]::new()
    $moreRecords = $true
    $offset = 0
    do {
        $tenantUrl = $baseURL[$i] + 'infrastructure/sites/tenants?limit=500&offset=' + $offset
        $results = Invoke-RestMethod -Uri $tenantURL -Method GET -Headers $headers[$i]
        $tenants.AddRange($results.data)
        if($results.meta.pagingInfo.count -ge 500) {
            $offset += 500
        }
        else {
            $moreRecords = $false
        }
    } while($moreRecords)

    #get VBR CC host information
    $CChosts = [System.Collections.ArrayList]::new()
    $moreRecords = $true
    $offset = 0
    do {
        $hostUrl = $baseURL[$i] + 'infrastructure/sites?limit=500&offset=' + $offset
        $results = Invoke-RestMethod -Uri $hostURL -Method GET -Headers $headers[$i]
        $CChosts.AddRange($results.data)
        if($results.meta.pagingInfo.count -ge 500) {
            $offset += 500
        }
        else {
            $moreRecords = $false
        }
    } while($moreRecords)

    #get license reporting info for all customers
    $string = $regions[$i] + ": Getting list of licenses in use..."
    Write-Host $string
    
    $licensingInfo = [System.Collections.ArrayList]::new()
    $moreRecords = $true
    $offset = 0
    do {
        $licensingURL = $baseURL[$i] + 'licensing/usage/organizations?limit=500&offset=' + $offset
        $results = Invoke-RestMethod -Uri $licensingURL -Method GET -Headers $headers[$i]
        $licensingInfo.AddRange($results.data)

        if($results.meta.pagingInfo.count -ge 500) {
            $offset += 500
        }
        else {
            $moreRecords = $false
        }
    } while($moreRecords)

    #get all sites
    $string = $regions[$i] + ": Getting list of all company sites..."
    Write-Host $string
    $sites = [System.Collections.ArrayList]::new()
    $moreRecords = $true
    $offset = 0
    do {
        $sitesURL = $baseURL[$i] + 'organizations/companies/sites?limit=500&offset=' + $offset
        $results = Invoke-RestMethod -Uri $sitesURL -Method GET -Headers $headers[$i]
        $sites.AddRange($results.data)

        if($results.meta.pagingInfo.Count -ge 500) {
            $offset += 500
        }
        else {
            $moreRecords = $false
        }
    } while($moreRecords)

    #get usage metrics for all companies (for things like Insider Protection)
    $string = $regions[$i] + ": Getting basic company usage information..."
    Write-Host $string
    $allResources = [System.Collections.ArrayList]::new()
    $moreRecords = $true
    $offset = 0
    do {
        $resourceUsageURL = $baseURL[$i] + 'organizations/companies/usage?limit=500&offset=' + $offset
        $results = Invoke-RestMethod -Uri $resourceUsageURL -Method GET -Headers $headers[$i]
        $allResources.AddRange($results.data)

        if($results.meta.pagingInfo.count -ge 500) {
            $offset += 500
        }
        else {
            $moreRecords = $false
        }
    } while($moreRecords)

    #get vb365 server list
    $string = $regions[$i] + ": Getting VB365 server list..."
    Write-Host $string
    $vb365Servers = [System.Collections.ArrayList]::new()
    $moreRecords = $true
    $offset = 0
    do {
        $vb365ServerURL = $baseURL[$i] + 'infrastructure/vb365Servers?limit=500&offset=' + $offset
        $results = Invoke-RestMethod -Uri $vb365ServerURL -Method GET -Headers $headers[$i]
        $vb365Servers.AddRange($results.data)

        if($results.meta.pagingInfo.count -ge 500) {
            $offset += 500
        }
        else {
            $moreRecords = $false
        }
    } while($moreRecords)

    #get vb365 proxies
    $string = $regions[$i] + ": Getting VB365 proxy list..."
    Write-Host $string
    $vb365Proxies = [System.Collections.ArrayList]::new()
    foreach($server in $vb365Servers) {
        $moreRecords = $true
        $offset = 0
        do {
            $vb365ProxyURL = $baseURL[$i] + 'infrastructure/vb365Servers/' + $server.instanceUid + '/backupProxies?limit=500&offset=' + $offset
            $results = Invoke-RestMethod -Uri $vb365ProxyURL -Method GET -Headers $headers[$i]
            $vb365Proxies.AddRange($results.data)

            if($results.meta.pagingInfo.count -ge 500) {
                $offset += 500
            }
            else {
                $moreRecords = $false
            }
        } while($moreRecords)
    }

    #get vb365 repositories
    $string = $regions[$i] + ": Getting VB365 repository list..."
    Write-Host $string
    $vb365Repos = [System.Collections.ArrayList]::new()
    foreach($server in $vb365Servers) {
        $moreRecords = $true
        $offset = 0
        do {
            $vb365RepoURL = $baseURL[$i] + 'infrastructure/vb365Servers/' + $server.instanceUid + '/backupRepositories?limit=500&offset=' + $offset
            $results = Invoke-RestMethod -Uri $vb365RepoURL -Method GET -Headers $headers[$i]
            $vb365Repos.AddRange($results.data)

            if($results.meta.pagingInfo.count -ge 500) {
                $offset += 500
            }
            else {
                $moreRecords = $false
            }
        } while($moreRecords)
    }

    #get vb365 resources
    $string = $regions[$i] + ": Getting VB365 resource information..."
    Write-Host $string
    $vb365Resources = [System.Collections.ArrayList]::new()
    $moreRecords = $true
    $offset = 0
    do {
        $vb365UsageURL = $baseURL[$i] + 'organizations/companies/hostedResources/vb365?limit=500&offset=' + $offset
        $results = Invoke-RestMethod -Uri $vb365UsageURL -Method GET -Headers $headers[$i]
        $vb365Resources.AddRange($results.data)

        if($results.meta.pagingInfo.count -ge 500) {
            $offset += 500
        }
        else {
            $moreRecords = $false
        }
    } while($moreRecords)

    #get M365 backup resources
    $string = $regions[$i] + ": Getting all VB365 backup resource information..."
    Write-Host $string
    $vb365BackupResourceList = [System.Collections.ArrayList]::new()
    $moreRecords = $true
    $offset = 0
    do {
        $vb365BackupUsageURL = $baseURL[$i] + 'organizations/companies/hostedResources/vb365/backupResources?limit=500&offset=' + $offset
        $results = Invoke-RestMethod -Uri $vb365BackupUsageURL -Method GET -Headers $headers[$i]
        $vb365BackupResourceList.AddRange($results.data)

        if($results.meta.pagingInfo.count -ge 500) {
            $offset += 500
        }
        else {
            $moreRecords = $false
        }
    } while($moreRecords)

    $string = $regions[$i] + ": Getting full resource usage information..."
    Write-Host $string
    #get backup, replication, and M365 information and save to array
    $BaaSResources = @()
    $DRaaSResources = @()
    $NoResources = @()
    $M365Resources = @()
    $vb365Backups = @()
    $Licenses = @()
    $ResourceInfo = ""
    $count = 1
    foreach($company in $companies) {
        $string = $regions[$i] + ": Collecting information for tenant (" + $count + "/" + $companies.Count + ") " + $company.name + " - " + $company.instanceUid + "..."
        Write-host $string

        #get license consumption
        foreach($license in $licensingInfo) {
            if($license.organizationUid -eq $company.instanceUid) {
                foreach($server in $license.servers) {
                    foreach($workload in $server.workloads) {
                        if($server.serverType -ne 'VB365' -and $server.serverType -ne 'CloudConnect') {
			                #Get who owns the license
			                $licenseOwnerUri = $baseURL[$i] + '/licensing/backupServers/' + $server.serverUid
			                $moreLicenseInfo = (Invoke-RestMethod -Method GET -Uri $licenseOwnerUri -Headers $headers[$i]).data
                            $licenseOwner = $moreLicenseInfo.company
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


                        $Licenses += $LicenseInfo
                    }
                }

                $licensingInfo.Remove($license)
                break
            }
        }

        # Process VB365 backup resources (join VB365 server, proxy, and repository info)
        $usingM365 = $false
        foreach ($backupResource in @($vb365BackupResourceList)) {
            if ($backupResource.companyUid -eq $company.instanceUid) {
                $usingM365 = $true
                foreach ($resource in @($vb365Resources)) {
                    if ($backupResource.vb365ResourceUid -eq $resource.instanceUid) {
                        $vb365Proxy = ""
                        $vb365ProxyPool = ""
                        $vb365Retention = ""
                        $vb365Repo = ""
                        $vb365ObjectRepo = $false
                        $vb365CopyRepo = $false
                        $vb365PrimaryRepo = $false

                        if ($null -ne $backupResource.proxyUid) {
                            foreach ($proxy in @($vb365Proxies)) {
                                if ($backupResource.proxyUid -eq $proxy.instanceUid) {
                                    $vb365Proxy = $proxy.hostName
                                    break
                                }
                            }
                        }

                        # Proxy pool logic
                        if ($null -ne $backupResource.proxyPoolUid) {
                            foreach ($proxyPool in @($vb365ProxyPools)) {
                                if ($backupResource.proxyPoolUid -eq $proxyPool.instanceUid) {
                                    $vb365ProxyPool = $proxyPool.name
                                    break
                                }
                            }
                        }

                        foreach ($repo in @($vb365Repos)) {
                            if ($backupResource.repositoryUid -eq $repo.instanceUid) {
                                $vb365Repo = $repo.name
                                $vb365Retention = $repo.retentionType
                                $vb365RetentionPeriodType = $repo.retentionPeriodType

                                if($vb365RetentionPeriodType -eq "Yearly" -or $vb365RetentionPeriodType -eq "Keep Years") {
                                    if($repo.yearlyRetentionPeriod -eq "Year1") {
                                        $vb365RetentionPeriod = $repo.yearlyRetentionPeriod.Replace("Year", "") + " Year"
                                    }
                                    else {
                                        $vb365RetentionPeriod = $repo.yearlyRetentionPeriod.Replace("Years", "") + " Years"
                                    }
                                }
                                elseif($vb365RetentionPeriodType -eq "Monthly") {
                                    if($repo.yearlyRetentionPeriod -eq "Month1") {
                                        $vb365RetentionPeriod = $repo.yearlyRetentionPeriod.Replace("Month", "") + " Month"
                                    }
                                    else {
                                        $vb365RetentionPeriod = $repo.yearlyRetentionPeriod.Replace("Months", "") + " Months"
                                    }
                                }
                                elseif($vb365RetentionPeriodType -eq "Daily") {
                                    if($repo.yearlyRetentionPeriod -eq "Day1") {
                                        $vb365RetentionPeriod = $repo.yearlyRetentionPeriod.Replace("Day", "") + " Day"
                                    }
                                    else {
                                        $vb365RetentionPeriod = $repo.yearlyRetentionPeriod.Replace("Day", "") + " Day"
                                    }
                                }
                                elseif($vb365RetentionPeriodType -eq "Keep Years") {
                                    $vb365RetentionPeriod = "Forever"
                                }
                                else {
                                    $vb365RetentionPeriod = $vb365RetentionPeriodType
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
                                break
                            }
                        }

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
                        $vb365Backups += $ResourceInfo
                        $vb365Resources.Remove($resource)
                        break
                    }
                }
                $vb365BackupResourceList.Remove($backupResource)
            }
        }

        #get M365 usage
        $RIPUsage = 0
        foreach($resource in $allResources) {
            if($resource.companyUid -eq $company.instanceUid) {
                $protectedUsers = $protectedTeams = $protectedSites = $M365BackupSize = 0
                foreach($counter in $resource.counters) {
                    if($counter.Type -eq "Vb365ProtectedUsers") {
                        $protectedUsers = $counter.value
                    }
                    elseif($counter.Type -eq "Vb365ProtectedTeams") {
                        $protectedTeams = $counter.value
                    }
                    elseif($counter.Type -eq "Vb365ProtectedSites") {
                        $protectedSites = $counter.value
                    }
                    elseif($counter.Type -eq "Vb365BackupSize") {
                        $M365BackupSize = $counter.value / 1TB
                    }
                    elseif($counter.type -eq "CloudInsiderProtectionBackupSize") {
                        $RIPUsage = $counter.value / 1TB
                    }
                }
                
                if($protectedUsers -gt 0 -or $M365BackupSize -gt 0) {
                    # Get the licensed user count for this company from the $Licenses collection
                    $licenseRecord = $Licenses | Where-Object { $_.ID -eq $company.instanceUid -and $_.Type -eq 'VB365' }
                    $licensedUsers = 0
                    if ($licenseRecord) {
                        $licensedUsers = $licenseRecord.Count
                    }

                    $ResourceInfo = [PSCustomObject]@{
                        'Company' = $company.Name
                        'ID' = $resource.companyUid
                        'Licensed Users' = $licensedUsers
                        'Protected Users' = $protectedUsers
                        'Protected Sites' = $protectedSites
                        'Protected Teams' = $protectedTeams
                        'Backup Size (TB)' = $M365BackupSize
                    }
                    $usingM365 = $true
                    $M365Resources += $ResourceInfo
                }

            }
        }

        foreach($site in $sites) {
            if($site.companyUid -eq $company.instanceUid) {
                $matchFound = $false

                foreach($tenant in $tenants) {
                    if($tenant.instanceUid -eq $site.cloudTenantUid) {
                        $matchFound = $true

                        $enabled = "No"
                        if($tenant.isEnabled) {
                            $enabled = "Yes"
                        }
                    
                        #get used BaaS space
                        $usageURL = $baseURL[$i] + 'organizations/companies/' + $company.instanceUid + '/sites/' + $site.siteUid + '/backupResources/usage'
                        $usage = (Invoke-RestMethod -Uri $usageURL -Method GET -Headers $headers[$i]).data

                        #temporarily not using due to VSPC bug
                        #if($tenant.isBackupResourcesEnabled) {
                        if($usage.Count -gt 0) {
                            $RIPEnabled = "No"
                            $RIPDays = 0
                            if($tenant.isBackupProtectionEnabled) {
                                $RIPEnabled = "Yes"
                                $RIPDays = $tenant.backupProtectionPeriod
                            }

                            if($usage.Count -gt 0) {
                                try {
                                    #get parent repository id
                                    $resourceURL = $baseURL[$i] + 'organizations/companies/' + $company.instanceUid + '/sites/' + $site.siteUid + '/backupResources/' + $usage.backupResourceUid
                                    $resource = (Invoke-RestMethod -Uri $resourceURL -Method GET -Headers $headers[$i]).data

                                    foreach($CChost in $CChosts) {
                                        if($tenant.backupServerUid -eq $CChost.siteUid) {
                                            #get parent repository information
                                            $repoURL = $baseURL[$i] + 'infrastructure/backupServers/' + $resource.siteUid + '/repositories/' + $resource.repositoryUid + '?expand=BackupRepositoryInfo'
                                            $repo = (Invoke-RestMethod -Uri $repoURL -Method GET -Headers $headers[$i]).data

                                            $hotUsed = $capacityUsed = $archiveUsed = 0
                                            $hotUsed = $usage.performanceTierUsage / 1TB
                                            $capacityUsed = $usage.capacityTierUsage / 1TB
                                            $archiveUsed = $usage.archiveTierUsage / 1TB
                                            $totalUsed = $usage.usedStorageQuota / 1TB
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
                                                'RIP Enabled' = $RIPEnabled
                                                'RIP Days' = $RIPDays
                                                'RIP Usage (TB)' = $RIPUsage
                                                'Hot Storage Used (TB)' = $hotUsed
                                                'Capacity Tier Used (TB)' = $capacityUsed
                                                'Archive Tier Used (TB)' = $archiveUsed
                                                'Total Storage Used (TB)' = $totalUsed
                                                'Quota (TB)' = $usage.storageQuota / 1TB
                                                'Percent Used' = $usage.usedStorageQuota / $usage.storageQuota * 100
                                            }
                                            $BaaSResources += $ResourceInfo
                                            break
                                        }
                                    }
                                }
                                catch {}
                            }
                        }

                        #get used DRaaS resources
                        if($tenant.isNativeReplicationResourcesEnabled) {
                            $RepUsageURL = $baseURL[$i] + 'organizations/companies/' + $company.instanceUid + '/sites/' + $site.siteUid + '/replicationResources/usage'
                            $RepUsage = (Invoke-RestMethod -Uri $RepUsageURL -Method GET -Headers $headers[$i]).data
                            if($RepUsage.Count -gt 0) {
                                if($null -eq $RepUsage.vCPUsConsumed) {
                                    $vCPUsConsumed = 0
                                }
                                else {
                                    $vCPUsConsumed = $RepUsage.vCPUsConsumed
                                }

                                foreach($CChost in $CChosts) {
                                    if($tenant.backupServerUid -eq $CChost.siteUid) {
                                        $ResourceInfo = [PSCustomObject]@{
                                            'Company' = $company.Name
                                            'Tenant Name' = $tenant.Name
                                            'Description' = $tenant.description
                                            'Enabled' = $enabled
                                            'ID' = $compnany.instanceUid
                                            'Last Active' = $tenant.lastActive
                                            'VBR Host' = $CChost.siteName
                                            'Replic vCPUs' = $vCPUsConsumed
                                            'Replic Memory Used (GB)' = $RepUsage.memoryUsage / 1024 / 1024 / 1024
                                            'Replic Storage Used (TB)' = $RepUsage.storageUsage / 1024 /1024 / 1024 / 1024
                                        }

                                        $DRaaSResources += $ResourceInfo
                                    }
                                }
                            }
                        }

                        #get used VCD DRaaS resources
                        if($tenant.isVcdReplicationResourcesEnabled) {
                            $VCDRepUsageURL = $baseURL[$i] + 'organizations/companies/' + $company.instanceUid + '/sites/' + $site.siteUid + '/vcdReplicationResources/usage'
                            $VCDRepUsage = (Invoke-RestMethod -Uri $VCDRepUsageURL -Method GET -Headers $headers[$i]).data
                            if($VCDRepUsage.Count -gt 0) {
                                if($null -eq $VCDRepUsage.vCPUsConsumed) {
                                    $vCPUsConsumed = 0
                                }
                                else {
                                    $vCPUsConsumed = $VCDRepUsage.vCPUsConsumed
                                }

                                foreach($CChost in $CChosts) {
                                    if($tenant.backupServerUid -eq $CChost.siteUid) {
                                        $ResourceInfo = [PSCustomObject]@{
                                            'Company' = $company.Name
                                            'Tenant Name' = $tenant.Name
                                            'Description' = $tenant.description
                                            'Enabled' = $enabled
                                            'ID' = $company.instanceUid
                                            'Last Active' = $tenant.lastActive
                                            'VBR Host' = $CChost.siteName
                                            'Replic vCPUs' = $vCPUsConsumed
                                            'Replic Memory Used (GB)' = $VCDRepUsage.memoryUsage / 1024 / 1024 / 1024
                                            'Replic Storage Used (TB)' = $VCDRepUsage.storageUsage / 1024 /1024 / 1024 / 1024
                                        }

                                        $DRaaSResources += $ResourceInfo
                                    }
                                }
                            }
                        }

                        #list tenants with no resources assigned
                        #disabled until VSPC bug is fixed
                        #if(!$usingM365 -and !$tenant.isBackupResourcesEnabled -and !$tenant.isNativeReplicationResourcesEnabled -and !$tenant.isVcdReplicationResourcesEnabled) {
                        if(!$usingM365 -and $usage.Count -le 0 -and !$tenant.isNativeReplicationResourcesEnabled -and !$tenant.isVcdReplicationResourcesEnabled) {
                            foreach($CChost in $CChosts) {
                                if($tenant.backupServerUid -eq $CChost.siteUid) {
                                    $ResourceInfo = [PSCustomObject]@{
                                        'Company' = $company.Name
                                        'Tenant Name' = $tenant.Name
                                        'Description' = $tenant.description
                                        'Enabled' = $enabled
                                        'ID' = $company.instanceUid
                                        'Last Active' = $tenant.lastActive
                                        'VBR Host' = $CChost.siteName
                                    }

                                    $NoResources += $ResourceInfo
                                }
                            }
                        }

                        $tenants.Remove($tenant)
                        break
                    }
                }

                if($matchFound) {
                    $sites.Remove($site)
                    break
                }
            }
        }

        $count++
    }

    Write-Host Saving data to XLSX...
    #export usage information to CSV and save to downloads folder
    $Licenses | Export-Excel $FilePath[$i] -Autosize -TableName Licensing -WorksheetName Licensing
    $BaaSResources | Export-Excel $FilePath[$i] -Autosize -TableName BaaSResources -WorksheetName BaaS
    $DRaaSResources | Export-Excel $FilePath[$i] -Autosize -TableName DRaaSResources -WorksheetName DRaaS
    $M365Resources | Export-Excel $FilePath[$i] -Autosize -TableName M365Resources -WorksheetName "M365 Protection Metrics"
    $vb365Backups | Export-Excel $FilePath[$i] -Autosize -TableName "VB365_Repos" -WorksheetName "VB365 Repository Information"
    $NoResources | Export-Excel $FilePath[$i] -Autosize -TableName "No_Resources" -WorksheetName "No Resources"
}

for($i=0; $i -lt $regions.Count; $i++) {
    $string = $regions[$i] + ": Report available under " + $FilePath[$i]
    Write-Host $string
}

#tell user file location and exit when any key is pressed
Write-Host -NoNewLine 'Press any key to exit...';
$null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown');
