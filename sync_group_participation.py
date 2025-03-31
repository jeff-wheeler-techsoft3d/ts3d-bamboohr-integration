import json
from msal import ConfidentialClientApplication
import os
import base64
import requests
import pycountry

def get_country_code(text):
    countries = pycountry.countries
    for country in countries:
        if country.name.lower() in text.lower():
            return country.alpha_2
    return None

def check_and_add_to_group(ts3d_employee_group, employee_work_email, ts3d_groups):
    if ts3d_employee_group not in ts3d_groups:
        ts3d_groups[ts3d_employee_group] = []
    ts3d_groups[ts3d_employee_group].append(employee_work_email)
    

bamboo_report = 149

headers = {"Accept": "application/json", "content-type": "application/json"}

api_key = os.environ['bamboo_hr_key']

url = "https://" + api_key +":x@api.bamboohr.com/api/gateway.php/techsoft3d/v1/reports/" + str(bamboo_report) + "?format=json&fd=yes&onlyCurrent=true"


response = requests.get(url)

# JSON object containing distribution list association
ts3d_groups = {}

####################################
#Process Data Driven Employee Groups
####################################

# Process People Managers
bamboo_report = 144
url = "https://" + api_key +":x@api.bamboohr.com/api/gateway.php/techsoft3d/v1/reports/" + str(bamboo_report) + "?format=json&fd=yes&onlyCurrent=true"
response = requests.get(url)
if response.status_code == 200:
    if hasattr(response, "text") and response.text != '':
        json_raw = json.loads(response.text)
        if 'employees' in json_raw:
            for employee in json_raw['employees']:
                if employee['employmentHistoryStatus'] != 'Terminated' and employee['workEmail']:
                    # Update world Employee Groups
                    check_and_add_to_group("World", employee['workEmail'], ts3d_groups)

                    if employee['customDistributionListsTest'] != None:
                        for ts3d_group in employee['customDistributionListsTest'].split(','):
                            check_and_add_to_group(ts3d_group, employee['workEmail'], ts3d_groups)

                    if employee['customEmployeeGroupOverride'] != None:
                        for ts3d_group in employee['customEmployeeGroupOverride'].split(','):
                            check_and_add_to_group(ts3d_group, employee['workEmail'], ts3d_groups)
                            

                    #Update World.Country Employee Groups
                    if employee['country'] != None:
                        country_code = get_country_code(employee['country'])
                        if country_code != None:
                            check_and_add_to_group("Employees." + country_code, employee['workEmail'], ts3d_groups)                    

                        if 'location' in employee and employee['location'] and 'Remote' in employee['location']:
                            check_and_add_to_group("Employees.Remote", employee['workEmail'], ts3d_groups)

                    
                    if 'division' in employee:
                        if 'Toolkits' in employee['division']:
                            check_and_add_to_group("World.Toolkits", employee['workEmail'], ts3d_groups)
                        if 'Applications' in employee['division']:
                            check_and_add_to_group("World.Apps", employee['workEmail'], ts3d_groups)
                    
                    if employee['customExecutive']:
                        employee_by_exec_group = employee['customExecutive'].replace('Exec', '').replace(' ', '')
                        check_and_add_to_group(employee_by_exec_group, employee['workEmail'], ts3d_groups)
                        if employee['department']:
                            department = employee['department'].replace(' ', '')
                            if department != employee_by_exec_group:
                                employee_by_exec_department = employee_by_exec_group + "." + department
                                check_and_add_to_group(employee_by_exec_department, employee['workEmail'], ts3d_groups)

                    #Update People Managers Employee Groups
                    if employee["-44"] != None and employee["-44"] == "1":
                        check_and_add_to_group("People.Managers", employee['workEmail'], ts3d_groups)

                        if 'division' in employee:
                            if 'Toolkits' in employee['division']:
                                check_and_add_to_group("People.Managers.Toolkits", employee['workEmail'], ts3d_groups)
                            if 'Applications' in employee['division']:
                                check_and_add_to_group("People.Managers.Apps", employee['workEmail'], ts3d_groups)
                    


####################################
#Update Other Systems with Employee Groups
#Process distribution list from Bamboo HR into different systems.
####################################

for ts3d_group in ts3d_groups:
    group_description = "This is an auto generated group - " + ts3d_group + " - based on the Bamboo HR parameters"
    print(ts3d_group + "@techsoft3d.com employee group is described as follows: " + group_description + " with the users as follows: " + str(ts3d_groups[ts3d_group]))

    if ts3d_group == "OpsTech.Test":
        print("#######################################################################################################################################")
        print("#  TEST   #############################################################################################################################")
        print("#######################################################################################################################################")

        #############################################
        # Update  data in Office 365 from Bamboo HR.
        #############################################
        azure_app_client_id = os.environ['azure_app_client_id']
        azure_tenant_id = os.environ['azure_tenant_id']
        azure_client_secret = os.environ['azure_client_secret']

        client_id = azure_app_client_id
        client_secret = azure_client_secret
        tenant_id = azure_tenant_id
        authority = f"https://login.microsoftonline.com/{tenant_id}"
        scope = ["https://graph.microsoft.com/.default"]

        # Initialize the MSAL confidential client
        app = ConfidentialClientApplication(
            client_id,
            authority=authority,
            client_credential=client_secret
        )

        # Acquire a token
        token_response = app.acquire_token_for_client(scopes=scope)

        if "access_token" in token_response:
            ms_user_ids = {}
            access_token = token_response["access_token"]
            headers = {
                'Authorization': f'Bearer {access_token}',
                'Content-Type': 'application/json'
            }

            # Microsoft Graph API Base URL
            GRAPH_API_URL = "https://graph.microsoft.com/v1.0"

            # Helper function to make API requests
            def make_o365_request(method, endpoint, data=None):
                url = f"{GRAPH_API_URL}{endpoint}"
                headers = {
                    "Authorization": f"Bearer {access_token}",
                    "Content-Type": "application/json"
                }
                response = requests.request(method, url, headers=headers, json=data)
                if not response.ok:
                    print(f"ERROR: API request failed: {response.status_code} - {response.text}")
                return response.json() if response.text else {}

            # Check if a group exists by its email address
            def get_o365_group_by_email(mail_nickname):
                try:
                    response = make_o365_request("GET", f"/groups?$filter=mailNickname eq '{mail_nickname}'")
                    groups = response.get("value", [])
                    if groups:
                        return groups[0]  # Return the group object
                    return None
                except Exception as e:
                    print(f"ERROR: Failed to retrieve group: {e}")

            # Create a Microsoft 365 or Security Group
            def create_o365_group(group_email, display_name, mail_nickname, description):
                group_data = {
                    "displayName": display_name,
                    "mailNickname": mail_nickname,
                    "description": description,
                    "mailEnabled": True,  # True for M365, False for Security Groups
                    "securityEnabled": True,  # True for Security Groups
                    "groupTypes": ["Unified"],
                    "proxyAddresses": [
                        f"SMTP:{group_email}",
                        f"smtp:{mail_nickname}@hoops3d.onmicrosoft.com"  # The primary email address, "SMTP" in uppercase
                    ]
                }
                try:
                    response = make_o365_request("POST", "/groups", group_data)
                    print(f"O365 Group '{display_name}' created successfully.")
                    return response["id"]
                except Exception as e:
                    print(f"ERROR: Failed to create group '{display_name}': {e}")

            # Get current members of a group
            def get_o365_group_members(group_id):
                members = []
                endpoint = f"/groups/{group_id}/members"
                while endpoint:
                    response = make_o365_request("GET", endpoint)
                    members.extend(response.get("value", []))
                    endpoint = response.get("@odata.nextLink")  # Pagination support
                return [member["id"] for member in members]
            
            def delete_o365_group(group_id):
                try:
                    response = make_o365_request("DELETE", "/groups/" + str(group_id))
                    print(f"O365 Group '{group_id}' deleted successfully.")
                except Exception as e:
                    print(f"ERROR: Failed to delete group '{group_id}': {e}")
        

            # Add a user to a group
            def add_o365_user_to_group(group_id, user_id):
                member_ref = {
                    "@odata.id": f"{GRAPH_API_URL}/directoryObjects/{user_id}"
                }
                try:
                    make_o365_request("POST", f"/groups/{group_id}/members/$ref", member_ref)
                    print(f"User '{user_id}' added to group '{group_id}'.")
                except Exception as e:
                    print(f"ERROR: Failed to add user '{user_id}' to group '{group_id}': {e}")

            # Remove a user from a group
            def remove_o365_user_from_group(group_id, user_id):
                try:
                    make_o365_request("DELETE", f"/groups/{group_id}/members/{user_id}/$ref")
                    print(f"User '{user_id}' removed from group '{group_id}'.")
                except Exception as e:
                    print(f"ERROR: Failed to remove user '{user_id}' from group '{group_id}': {e}")

            # Map email addresses to Azure AD user IDs
            def get_o365_user_id_by_email(email):
                try:
                    response = make_o365_request("GET", f"/users?$filter=userPrincipalName eq '{email}'")
                    users = response.get("value", [])
                    if users:
                        return users[0]["id"]
                    print(f"User '{email}' not found.")
                    return None
                except Exception as e:
                    print(f"ERROR: Failed to retrieve user ID for '{email}': {e}")

            # Synchronize group membership with expected users
            def sync_o365_group_members(group_email, display_name, mail_nickname, description, expected_user_emails):
                # Check if the group exists, create if it doesn't
                group = get_o365_group_by_email(mail_nickname)
                if not group:
                    print(f"Group '{group_email}' does not exist. Creating...")
                    group_id = create_o365_group(group_email, display_name, mail_nickname, description)
                else:
                    group_id = group["id"]

                # Map expected emails to user IDs
                expected_user_ids = []
                for email in expected_user_emails:
                    user_id = get_o365_user_id_by_email(email)
                    if user_id:
                        expected_user_ids.append(user_id)

                # Get current members of the group
                current_members = get_o365_group_members(group_id)

                # Calculate users to add and remove
                to_add = set(expected_user_ids) - set(current_members)
                to_remove = set(current_members) - set(expected_user_ids)
                    
                # Update group
                if to_add:
                    print(f"Adding users to group {ts3d_group}@techsoft3d.com ({len(to_add)}): {to_add}")
                if to_remove:
                    print(f"Removing users from group {ts3d_group}@techsoft3d.com ({len(to_remove)}): {to_remove}")
                
                # Add missing users
                for user_id in to_add:
                    add_o365_user_to_group(group_id, user_id)

                # Remove extra users
                for user_id in to_remove:
                    remove_o365_user_from_group(group_id, user_id)

            sync_o365_group_members(ts3d_group + "@techsoft3d.com", ts3d_group, ts3d_group, group_description, ts3d_groups[ts3d_group])

        else:
            print("Error acquiring token:", token_response.get("error_description"))

        #############################################
        # Update data in Atlassian from Bamboo HR.
        #############################################

        # Variables
        BASE_URL = "https://techsoft3d.atlassian.net"  # Replace with your Atlassian domain
        API_EMAIL = os.environ['atlassian_user_id']
        API_TOKEN = os.environ['atlassian_token']

        # Helper function to make authenticated API requests
        def make_atlassian_request(method, endpoint, data=None):
            url = f"{BASE_URL}{endpoint}"
            headers = {
                "Authorization": "Basic " + base64.b64encode((API_EMAIL + ":" + API_TOKEN).encode('utf-8')).decode("utf-8"),
                "Content-Type": "application/json",
            }
            response = requests.request(method, url, headers=headers, json=data)
            if not response.ok:
                print(f"ERROR: API request failed: {response.status_code} - {response.text}")
            return response.json() if response.text else {}

        # Check if a group exists
        def get_atlassian_group(group_name):
            try:
                response = make_atlassian_request("GET", f"/rest/api/3/group/member?groupname={group_name}")
                return response
            except Exception as e:
                if "does not exist" in str(e):
                    return None
                print("ERROR: Unable to find atlassian group" + str(group_name))

        # Create a new group
        def create_atlassian_group(group_name):
            try:
                make_atlassian_request("POST", "/rest/api/3/group", {"name": group_name})
                print(f"Group '{group_name}' created successfully.")
            except Exception as e:
                print(f"Failed to create group '{group_name}': {e}")

        # Get current members of a group
        def get_atlassian_group_members(group_name):
            members = []
            start_at = 0
            while True:
                response = make_atlassian_request("GET", f"/rest/api/3/group/member?groupname={group_name}&startAt={start_at}")
                members.extend(response.get("values", []))
                if response.get("isLast", True):
                    break
                start_at += response.get("maxResults", 50)
            return [member["accountId"] for member in members]

        # Add a user to a group
        def add_atlassian_user_to_group(group_name, account_id):
            try:
                make_atlassian_request("POST", f"/rest/api/3/group/user?groupname={group_name}", {"accountId": account_id})
                print(f"User '{account_id}' added to group '{group_name}'.")
            except Exception as e:
                print(f"Failed to add user '{account_id}' to group '{group_name}': {e}")

        # Remove a user from a group
        def remove_atlassian_user_from_group(group_name, account_id):
            try:
                make_atlassian_request("DELETE", f"/rest/api/3/group/user?groupname={group_name}&accountId={account_id}")
                print(f"User '{account_id}' removed from group '{group_name}'.")
            except Exception as e:
                print(f"Failed to remove user '{account_id}' from group '{group_name}': {e}")

        # Map email addresses to account IDs
        def get_atlassian_user_account_id(email):
            try:
                response = make_atlassian_request("GET", f"/rest/api/3/user/search?query={email}")
                if response:
                    return response[0]["accountId"]
                else:
                    print(f"User '{email}' not found.")
                    return None
            except Exception as e:
                print(f"Failed to retrieve account ID for user '{email}': {e}")
                return None

        # Synchronize group membership with an expected list of emails
        def sync_atlassian_group_members(group_name, expected_user_emails):
            # Check if the group exists, create if it doesn't
            if not get_atlassian_group(group_name):
                print(f"Group '{group_name}' does not exist. Creating...")
                create_atlassian_group(group_name)

            # Map expected emails to account IDs
            expected_account_ids = []
            for email in expected_user_emails:
                account_id = get_atlassian_user_account_id(email)
                if account_id:
                    expected_account_ids.append(account_id)

            # Get current members of the group
            current_members = get_atlassian_group_members(group_name)

            # Calculate users to add and remove
            to_add = set(expected_account_ids) - set(current_members)
            to_remove = set(current_members) - set(expected_account_ids)

            # Update group
            if to_add:
                print(f"Adding users to group {ts3d_group}@techsoft3d.com ({len(to_add)}): {to_add}")
            if to_remove:
                print(f"Removing users from group {ts3d_group}@techsoft3d.com ({len(to_remove)}): {to_remove}")

            # Add missing users
            for account_id in to_add:
                add_atlassian_user_to_group(group_name, account_id)

            # Remove extra users
            for account_id in to_remove:
                remove_atlassian_user_from_group(group_name, account_id)

        sync_atlassian_group_members(ts3d_group, ts3d_groups[ts3d_group])

        #############################################
        # Update data in JumpCloud from Bamboo HR.
        #############################################

        # JumpCloud API base URL
        JC_API_URL = "https://console.jumpcloud.com/api"

        # Your JumpCloud API key
        API_KEY = os.environ['jumpcloud_token']

        # Headers for API requests
        HEADERS = {
            "Accept": "application/json",
            "Content-Type": "application/json",
            "x-api-key": API_KEY
        }

        def get_jumpcloud_user_group(group_name):
            url = "https://console.jumpcloud.com/api/v2/groups?filter=name:eq:" + group_name
            
            response = requests.get(url, headers=HEADERS)
            
            if response.status_code == 200:
                groups = response.json()
                return groups[0]
            else:
                print(f"Failed to retrieve group: {response.text}")
                return None

        def create_jumpcloud_user_group(group_name, description=""):
            """Create a new user group."""
            url = f"{JC_API_URL}/v2/usergroups"
            payload = {
                "name": group_name,
                "description": description,
                "type": "user_group"
            }
            response = requests.post(url, headers=HEADERS, json=payload)
            response.raise_for_status()
            return response.json()

        def get_jumpcloud_user_by_email(email):
            """Retrieve a user by email."""
            url = f"{JC_API_URL}/systemusers?filter=email:eq:{email}"
            response = requests.get(url, headers=HEADERS)
            response.raise_for_status()
            users = response.json().get("results", [])
            return users[0] if users else None

        def add_jumpcloud_user_to_group(user_id, group_id):
            """Add a user to a group."""
            url = f"{JC_API_URL}/v2/usergroups/{group_id}/members"
            payload = {
                "op": "add",
                "type": "user",
                "id": user_id
            }
            response = requests.post(url, headers=HEADERS, json=payload)
            response.raise_for_status()

        def remove_jumpcloud_user_from_group(user_id, group_id):
            """Remove a user from a group."""
            url = f"{JC_API_URL}/v2/usergroups/{group_id}/members"
            payload = {
                "op": "remove",
                "type": "user",
                "id": user_id
            }
            response = requests.post(url, headers=HEADERS, json=payload)
            response.raise_for_status()

        def get_jumpcloud_group_members(group_id):
            """Retrieve all members of a group."""
            url = f"{JC_API_URL}/v2/usergroups/{group_id}/members"
            response = requests.get(url, headers=HEADERS)
            response.raise_for_status()
            members = response.json()
            group_members = []
            for member in members:
                if 'to' in member:
                    member = member['to']
                if "type" in member and member["type"] == "user":
                    group_members.append(member["id"])
            
            return group_members

        def sync_jumpcloud_group_members(group_name, user_emails, description=""):
            """Synchronize group members with the specified list of user emails."""
            # Check if the group exists; create if it doesn't
            group = get_jumpcloud_user_group(group_name)
            if not group:
                print(f"Group '{group_name}' does not exist. Creating...")
                group = create_jumpcloud_user_group(group_name, description)
            group_id = group["id"]

            # Get current group members
            current_member_ids = set(get_jumpcloud_group_members(group_id))

            # Determine users to add and remove
            expected_member_ids = set()
            for email in user_emails:
                user = get_jumpcloud_user_by_email(email)
                if user:
                    expected_member_ids.add(user["id"])
                else:
                    print(f"User with email '{email}' not found.")

            to_add = expected_member_ids - current_member_ids
            to_remove = current_member_ids - expected_member_ids

            # Add users
            for user_id in to_add:
                add_jumpcloud_user_to_group(user_id, group_id)

            # Remove users
            for user_id in to_remove:
                remove_jumpcloud_user_from_group(user_id, group_id)

        sync_jumpcloud_group_members(ts3d_group, ts3d_groups[ts3d_group], group_description)
