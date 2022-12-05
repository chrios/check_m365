import requests
import argparse

valid_services = {
    'Exchange': 'Exchange Online',
    'OrgLiveID': 'Identity Service',
    'OSDPPlatform': 'Microsoft 365 suite',
    'Lync': 'Skype for Business',
    'SharePoint': 'SharePoint Online',
    'DynamicsCRM': 'Dynamics 365 Apps',
    'RMS': 'Azure Information Protection',
    'yammer': 'Yammer Enterprise',
    'MobileDeviceManagement': 'Mobile Device Management for Office 365',
    'Planner': 'Planner',
    'SwayEnterprise': 'Sway',
    'Intune': 'Microsoft Intune',
    'OneDriveForBusiness': 'OneDrive for Business',
    'microsoftteams': 'Microsoft Teams',
    'StaffHub': 'Microsoft StaffHub',
    'kaizalamessagingservices': 'Microsoft Kaizala',
    'Bookings': 'Microsoft Bookings',
    'officeonline': 'Office for the web',
    'DynamicsNAV': 'Dynamics 365 Business Central',
    'O365Client': 'Microsoft 365 Apps',
    'PowerApps': 'Power Apps',
    'PowerAppsM365': 'Power Apps in Microsoft 365',
    'MicrosoftFlow': 'Microsoft Power Automate',
    'MicrosoftFlowM365': 'Microsoft Power Automate in Microsoft 365',
    'Forms': 'Microsoft Forms',
    'Microsoft365Defender': 'Microsoft 365 Defender',
    'ProjectForTheWeb': 'Project for the web',
    'Stream': 'Microsoft Stream',
    'UniversalPrint': 'Universal Print',
    'Viva': 'Microsoft Viva',
    'cloudappsecurity': 'Microsoft Defender for Cloud Apps'
}

valid_service_ids = valid_services.keys()

all_args = argparse.ArgumentParser()

all_args.add_argument('-H', "--TenantDomain", required=True, help="Tenant Domain (xxx.onmicrosoft.com). This uses the $HOSTADDRESS$ macro usually, so make a host with the name xxx.onmicrosoft.com and call the check on it.")
all_args.add_argument('-c', "--ClientID", required=True, help="Client ID (Application ID)")
all_args.add_argument('-cs', "--ClientSecret", required=True, help="Client Secret")
all_args.add_argument('-cli', "--UseCLI", help="Output CLI formatted output", action='store_true')
all_args.add_argument('-t', "--TenantID", required=True, help="Tenant ID (Directory ID)")
all_args.add_argument('-s', "--Services", required=True, help="Service list (comma separated). If empty, checks all services")

args = vars(all_args.parse_args())

client_id = args['ClientID']
client_secret = args['ClientSecret']
tenant_id = args['TenantID']
tenant_domain = args['TenantDomain']
cli = args['UseCLI']

# construct allowed services list
query_services = list()
if ',' in args['Services']:
    # parse and convert to list
    query_services = args['Services'].split(',')
else:
    # append single entry to list
    query_services.append(args['Services'])

# make sure all items in allowed services list are in valid_services
for service in query_services:
    if not service in valid_service_ids:
        exit(f"Invalid service id: {service}. Please choose frome the list of valid service ids: {valid_service_ids}")

token_endpoint = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"

token_data = {
    "client_id": client_id,
    "scope": 'https://graph.microsoft.com/.default',
    "client_secret": client_secret,
    "grant_type": 'client_credentials'
}

token_response = requests.post(token_endpoint, token_data)

request_headers = {
    "Authorization": f"{token_response.json()['token_type']} {token_response.json()['access_token']}"
}

response = requests.get('https://graph.microsoft.com/v1.0/admin/serviceAnnouncement/healthOverviews', headers=request_headers)

api_services = response.json()['value']

degraded_services = list()

for api_service in api_services:
    # check that api_service is in query_services
    if api_service['id'] in query_services:
        if cli:
            print(f"Checking service {api_service['service']}")
        if (api_service['status']) != 'serviceOperational':
            if cli:
                print('  found issue!!')
            degraded_services.append(api_service['service'])

# construct return message
if len(degraded_services) > 0:
    print(f'WARNING: Found degraded services {degraded_services}')
else:
    print('OK: All services operational')