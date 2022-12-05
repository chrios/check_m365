# check_m365.py

## Usage
```
usage: check_m365.py [-h] -H TENANTDOMAIN -c CLIENTID -cs CLIENTSECRET [-cli] -t TENANTID -s SERVICES

options:
  -h, --help            show this help message and exit
  -H TENANTDOMAIN, --TenantDomain TENANTDOMAIN
                        Tenant Domain (xxx.onmicrosoft.com). This uses the $HOSTADDRESS$ macro usually, so make a host
                        with the name xxx.onmicrosoft.com and call the check on it.
  -c CLIENTID, --ClientID CLIENTID
                        Client ID (Application ID)
  -cs CLIENTSECRET, --ClientSecret CLIENTSECRET
                        Client Secret
  -cli, --UseCLI        Output CLI formatted output
  -t TENANTID, --TenantID TENANTID
                        Tenant ID (Directory ID)
  -s SERVICES, --Services SERVICES
                        Service list (comma separated). If empty, checks all services
```

This plugin is designed to be used with Nagios.

To use this application, first register a Azure Enterprise Application and grant it the following permissions on Microsoft Graph:

ServiceHealth.Read.All

Then, generate a client_secret, and use it in the script arguments, along with the tenant ID and application ID.

In Nagios, create a host with the name 'xxx.onmicrosoft.com', which is the tenant domain. Add a service using this script as the check_command.