# VSPC_Tenant_Resources

Collect information on each tenant in VSPC.

# Description

This script provides an easy way of compiling backup, replication, M365, and license consumption, plus some basic resource assignment information, such as the VCC server and repository. This information is then saved to an XLSX file where it can be loaded and filtered, for easily identifying customers affected by specific component outages or for billing.

# Usage

1. Update the regions with the regions where each VSPC is located (ex. US, NL). NOTE: Each region should be unique.
2. Update tokens with valid API keys
3. Update the base URLs with the VSPC server name.
4. Add or remove entries in 1-3 for each VSPC server you wish to query. The script by default is configured for 2 servers.

# Related Links

* https://helpcenter.veeam.com/docs/vac/rest/reference/vspc-rest.html
