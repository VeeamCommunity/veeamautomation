Set-ExecutionPolicy Bypass -Scope Process -Force

#TODO: Test backupResourcesEnabled boolean with v8
#TODO: Test new vb365 tab with VSPC v8

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

$regions = @("US", "NL")

#VSPC token and header
$Token1 = 'TOKEN_1_HERE'
$Token2 = 'TOKEN_2_HERE'
$Headers1 = @{
    Authorization="Bearer $Token1"
}
$Headers2 = @{
    Authorization="Bearer $Token2"
}
$headers = @($Headers1, $Headers2)


#VSPC base url
$BaseURL1 = "https://vspc01/api/v3/"
$BaseURL2 = "https://vspc02/api/v3/"
$baseURL = @($BaseURL1, $BaseURL2)

#File path for data export
$FilePath1 = $env:USERPROFILE + "\Downloads\" + $regions[0] + "_ConsumptionReport_" + (Get-Date -Format "yyyy-MM-dd_HH-mm-ss") + ".xlsx"
$FilePath2 = $env:USERPROFILE + "\Downloads\" + $regions[1] + "_ConsumptionReport_" + (Get-Date -Format "yyyy-MM-dd_HH-mm-ss") + ".xlsx"
$FilePath = @($FilePath1, $FilePath2)

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
    #NOTE: This API seems to act like a POST command with v7/API 3.3. Likely a bug. Needs to be tested against VSPC v8
    $string = $regions[$i] + ": Getting VB365 resource information..."
    Write-Host $string
    $vb365Resources = [System.Collections.ArrayList]::new()
    $moreRecords = $true
    $offset = 0
    do {
        $vb365UsageURL = $baseURL[$i] + 'organizations/companies/vb365Resources?limit=500&offset=' + $offset
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
    #NOTE: This API seems to act like a POST command with v7/API 3.3. Likely a bug. Needs to be tested against VSPC v8
    $string = $regions[$i] + ": Getting all VB365 backup resource information..."
    Write-Host $string
    $vb365BackupResourceList = [System.Collections.ArrayList]::new()
    $moreRecords = $true
    $offset = 0
    do {
        $vb365BackupUsageURL = $baseURL[$i] + 'organizations/companies/vb365Resources/backupResources?limit=500&offset=' + $offset
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

        #NOTE: This part needs testing with VSPC v8
        #get VB365 server name, repository name, retention period, and proxy name, repo type, and retention
        $usingM365 = $false
        foreach($backupResource in $vb36BackupResourceList) {
            if($backupResource.companyUid -eq $company.instanceUid) {
                $usingM365 = $true

                foreach($resource in $vb365Resources) {
                    if($backupResource.vb365ResourceUid -eq $resource.instanceUid) {
                        $vb365Host = $vb365Proxy = $vb365Retention = $vb365Repo = ""
                        $vb365ObjectRepo = $vb365CopyRepo = $vb365PrimaryRepo = $false

                        foreach($server in $vb365Servers) {
                            if($server.instanceUid -eq $resource.vb365ServerUid) {
                                $vb365Host = $resource.name
                                break
                            }
                        }

                        foreach($proxy in $vb365Proxies) {
                            if($backupResource.proxyUid -eq $proxy.instanceUid) {
                                $vb365Proxy = $proxy.hostName
                                break
                            }
                        }

                        foreach($repo in $vb365Repos) {
                            if($backupResource.repositoryUid -eq $repo.instanceUid) {
                                $vb365Repo = $repo.name
                                $vb365Retention = $repo.yearlyRetentionPeriod
                                $vb365ObjectRepo = $repo.isObjectStorageRepository
                                $M365BackupSize = $repo.usedSpaceBytes / 1024 / 1024 / 1024 / 1024
                                $vb365PrimaryRepo = $repo.isAvailableForBackupJob
                                $vb365CopyRepo = $repo.isAvailableForCopyJob
                                break
                            }
                        }
                        $ResourceInfo = [PSCustomObject]@{
                            'Company' = $company.name
                            'ID' = $company.instanceUid
                            'Server' = $vb365Host
                            'Proxy' = $vb365Proxy
                            'Repository' = $vb365Repo
                            'Retention' = $vb365Retention
                            'Is Object' = $vb365ObjectRepo
                            'Is Primary' = $vb365PrimaryRepo
                            'Is Copy' = $vb365CopyRepo
                        }
                        $vb365Backups += $ResourceInfo

                        $vb365Resources.Remove($resource)
                        break
                    }
                }

                $vb365BackupResourceList.Remove($backupResource)
                break;
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
                        $M365BackupSize = $counter.value
                    }
                    elseif($counter.type -eq "CloudInsiderProtectionBackupSize") {
                        $RIPUsage = $counter.value / 1024 / 1024 / 1024 / 1024
                    }
                }
                
                if($protectedUsers -gt 0 -or $M365BackupSize -gt 0) {
                    $ResourceInfo = [PSCustomObject]@{
                        'Company' = $company.Name
                        'ID' = $resource.companyUid
                        'Protected Users' = $protectedUsers
                        'Protected Sites' = $protectedSites
                        'Protected Teams' = $protectedTeams
                        'Backup Size (TB)' = $M365BackupSize
                    }
                    $usingM365 = $true
                    $M365Resources += $ResourceInfo
                }

                break
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
                                            $repoURL = $baseURL[$i] + 'infrastructure/backupServers/' + $resource.siteUid + '/repositories/' + $resource.repositoryUid
                                            $repo = (Invoke-RestMethod -Uri $repoURL -Method GET -Headers $headers[$i]).data

                                            $hotUsed = 0
                                            $archiveUsed = 0
                                            if($repo.type -eq "ScaleOut") {
                                                $hotUsed = $usage.perfomanceTierUsage / 1024 / 1024 / 1024 / 1024
                                                $archiveUsed = $usage.capacityTierUsage / 1024 / 1024 / 1024 / 1024
                                                $totalUsed = $hotUsed + $archiveUsed
                                            }
                                            else {
                                                $hotUsed = $usage.usedStorageQuota / 1024 / 1024 / 1024 / 1024
                                                $totalUsed = $usage.usedStorageQuota / 1024 / 1024 / 1024 / 1024
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
                                                'Type' = $repo.type
                                                'Repository Host' = $repo.hostName
                                                'RIP Enabled' = $RIPEnabled
                                                'RIP Days' = $RIPDays
                                                'RIP Usage (TB)' = $RIPUsage
                                                'Hot Storage Used (TB)' = $hotUsed
                                                'Archive Tier Used (TB)' = $archiveUsed
                                                'Total Storage Used (TB)' = $totalUsed
                                                'Quota (TB)' = $usage.storageQuota / 1024 / 1024 / 1024 / 1024
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

        #get license consumption
        foreach($license in $licensingInfo) {
            if($license.organizationUid -eq $company.instanceUid) {
                foreach($server in $license.servers) {
                    foreach($workload in $server.workloads) {
                        $LicenseInfo = [PSCustomObject]@{
                            'Company' = $company.Name
                            'ID' = $company.instanceUid
                            'Type' = $server.serverType
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

        $count++
    }



    Write-Host Saving data to XLSX...
    #export usage information to CSV and save to downloads folder
    $Licenses | Export-Excel $FilePath[$i] -Autosize -TableName Licensing -WorksheetName Licensing
    $BaaSResources | Export-Excel $FilePath[$i] -Autosize -TableName BaaSResources -WorksheetName BaaS
    $DRaaSResources | Export-Excel $FilePath[$i] -Autosize -TableName DRaaSResources -WorksheetName DRaaS
    $M365Resources | Export-Excel $FilePath[$i] -Autosize -TableName M365Resources -WorksheetName M365
    $M365Backups | Export-Excel $FilePath[$i] -Autosize -TableName "VB365_Repos" -WorksheetName "VB365 Repos"
    $NoResources | Export-Excel $FilePath[$i] -Autosize -TableName "No_Resources" -WorksheetName "No Resources"
}

for($i=0; $i -lt $regions.Count; $i++) {
    $string = $regions[$i] + ": Report available under " + $FilePath[$i]
    Write-Host $string
}

#tell user file location and exit when any key is pressed
Write-Host -NoNewLine 'Press any key to exit...';
$null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown');
