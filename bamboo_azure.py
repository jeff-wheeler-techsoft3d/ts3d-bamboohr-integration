import json
from msal import ConfidentialClientApplication
import os
import requests


def lambda_handler(event, context):
    headers = {"Accept": "application/json", "content-type": "application/json"}
    
    api_key = os.environ['bamboo_hr_key']
    
    url = "https://" + api_key +":x@api.bamboohr.com/api/gateway.php/techsoft3d/v1/reports/135?format=json&fd=yes&onlyCurrent=true"
    
    
    response = requests.get(url)
    
    # JSON object containing user IDs and new titles
    users_to_update = []
    
    if response.status_code == 200:
        if hasattr(response, "text") and response.text != '':
            json_raw = json.loads(response.text)
            if 'employees' in json_raw:
                for employee in json_raw['employees']:
                    if employee['workEmail'] != None and employee['jobTitle'] != None and employee['country'] != None:
                        users_to_update.append({"email": employee['workEmail'], "title": employee['jobTitle'], "country": employee['country']})
                    else:
                        print(f"Skipping employee {employee['workEmail']} due to missing data.")
            else:
                print("No 'employees' key found in the response.")
    else:
        if hasattr(response, "text") and response.text != '':
            print(f"Error fetching data from Bamboo HR: {response.status_code} - {response.text}")
        else:
            print(f"Error fetching data from Bamboo HR with no Text: {response.status_code}")
    
                
    
    ########
    # Update data in Office 365 from Bamboo HR.
    ########
    
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
        print("Token acquired successfully.")
        print(f"Number of users to update: {len(users_to_update)}")
        access_token = token_response["access_token"]
        headers = {
            'Authorization': f'Bearer {access_token}',
            'Content-Type': 'application/json'
        }
    
        # Prepare the batch request
        batch_request = {
            "requests": [
                {
                    "id": str(index),
                    "method": "PATCH",
                    "url": f"/users/{user['email']}",
                    "headers": {
                        "Content-Type": "application/json"
                    },
                    "body": {
                        "jobTitle": user["title"]
                    }
                } for index, user in enumerate(users_to_update)
            ]
        }
    
        # Function to send a batch request
        def send_batch_request(batch):
            batch_request = {
                "requests": [
                    {
                        "id": str(index),
                        "method": "PATCH",
                        "url": f"/users/{user['email']}",
                        "headers": {"Content-Type": "application/json"},
                        "body": {"jobTitle": user["title"], "country": user["country"]}
                    } for index, user in enumerate(batch)
                ]
            }
            batch_endpoint = 'https://graph.microsoft.com/v1.0/$batch'
            response = requests.post(batch_endpoint, headers=headers, json=batch_request)
            
            if response.status_code == 200:
                if hasattr(response, "text") and response.text != '':
                    responses = json.loads(response.text)
                    if 'responses' in responses:
                        for resp in responses['responses']:
                            if 'status' in resp and resp['status'] > 399:
                                print("Error in processing: ", resp['body'])
            else:
                if hasattr(response, "text") and response.text != '':
                    print(f"Batch Response Error: {response.status} - {response.text}")
                else:
                    print(f"Batch Response Error with no Text: {response.status}")
    
        # Split users into batches of 20 and send each batch
        batch_size = 20
        for i in range(0, len(users_to_update), batch_size):
            batch = users_to_update[i:i+batch_size]
            send_batch_request(batch)
    else:
        print("Error acquiring token:", token_response.get("error_description"))

lambda_handler(None, None)