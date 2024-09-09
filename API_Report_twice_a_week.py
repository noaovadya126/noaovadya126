import requests
import pandas as pd
from datetime import datetime

def url_base(env: str):
    """
    Get the base url for the selected environment
    :param env: environment to log in to (dev/stg/prod/fda)
    :return: base url
    """
    return f"https://api.{env}.neteerahealthbridge.neteera.com/"

def login(env: str, username: str, password: str) -> str:
    """
    Login using username and password
    :param env: environment to log in to (dev/stg, prod, fda)
    :param username: username to log in with
    :param password: password to log in with
    :return: token if login was successful, empty string otherwise
    """
    LOGIN_URL = r"ums/v2/users/login"
    url = f"{url_base(env=env)}{LOGIN_URL}"

    headers = {
        "content-type": "application/json"
    }

    username_and_password = {
        "username": username,
        "password": password
    }

    resp = requests.post(url=url, headers=headers, json=username_and_password)

    if not resp.ok:
        print(f"Failed to login ({resp.status_code}) {resp.text} for {username} in {env} environment")
        return ""

    print(f"Logged in successfully to {username} in {env} environment")
    return resp.json().get("accessJwt", {}).get("token")

# Variables
username = "amatzia.mass+10@neteera.com"
password = "1qaz@WSX"
tenantId = "30e13fa2-c7bb-4d36-8476-0647891ca95d"

# Login
token = login(env="fda", username=username, password=password)

# API URLs
url = "https://api.fda.neteerahealthbridge.neteera.com/device/v2/devices/stats/continuously-disconnected/count?continuousDisconnectionSeconds=86400"
url2 = "https://api.fda.neteerahealthbridge.neteera.com/organization/v2/tenants/30e13fa2-c7bb-4d36-8476-0647891ca95d/sub-tenants?limit=10000000"

# Request headers
headers = {
  'authorization': f'Bearer {token} ',
}

# Fetching data
data_discon = requests.get(url, headers=headers).json()['data']
data_names = requests.get(url2, headers=headers).json()['data']

# Convert data to pandas DataFrame
df = pd.DataFrame(data_discon)

# Extract relevant data from the first data dict
tenant_name_data = {d['id']: d['name'] for d in data_names}

# Map names to the DataFrame
df['Tenant Name'] = df['tenantId'].map(tenant_name_data)

# Create the "Total" column by summing up all the other relevant columns
df['Total'] = df[['connectedAssigned', 'disconnectedAssigned', 'connectedUnassigned', 'disconnectedUnassigned']].sum(axis=1)

# Create the "Disconnected" column by summing 'disconnectedAssigned' and 'disconnectedUnassigned'
df['Disconnected'] = df[['disconnectedAssigned', 'disconnectedUnassigned']].sum(axis=1)

# Add the 'Group' column by extracting the first word from the 'Tenant Name' column
df['Group'] = df['Tenant Name'].apply(lambda x: x.split(' ', 1)[0] if ' ' in x else x)

# Update the 'Tenant Name' column by removing the first word (i.e., the 'Group')
df['Tenant Name'] = df['Tenant Name'].apply(lambda x: x.split(' ', 1)[1] if ' ' in x else '')

# Reorder columns according to the new requirements
df = df[['Group', 'Tenant Name', 'Total', 'Disconnected', 'connectedAssigned', 'disconnectedAssigned', 'connectedUnassigned', 'disconnectedUnassigned']]

# Save the DataFrame to an Excel file
current_time_str = datetime.now().strftime("%Y_%m_%d_%H_%M_%S")
df.to_excel(f'data_output_{current_time_str}.xlsx', index=False)

print("Data exported to Excel successfully.")
