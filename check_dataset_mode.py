import requests

TENANT_ID = "b4c8bca6-b827-45e8-88a1-b98cced69842"
CLIENT_ID = "d8064124-ba13-4b42-920c-e63f191c2824"
CLIENT_SECRET = "PLACEHOLDER_FILL_IN_ON_RENDER"
WORKSPACE_ID = "9de6373f-91cf-4a7b-a587-eb94dc4ab826"
DATASET_ID = "a8a9338a-13b2-45c6-baaa-96da36891704"

TOKEN_URL = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"


def get_access_token():
    response = requests.post(
        TOKEN_URL,
        data={
            "grant_type": "client_credentials",
            "client_id": CLIENT_ID,
            "client_secret": CLIENT_SECRET,
            "scope": "https://analysis.windows.net/powerbi/api/.default",
        },
    )
    response.raise_for_status()
    return response.json()["access_token"]


def main():
    token = get_access_token()
    print("Got access token OK.")
    url = f"https://api.powerbi.com/v1.0/myorg/groups/{WORKSPACE_ID}/datasets/{DATASET_ID}"
    response = requests.get(url, headers={"Authorization": f"Bearer {token}"})
    if response.status_code != 200:
        print(f"Request failed: {response.status_code}")
        print(response.text)
        return
    data = response.json()
    print("Dataset name:", data.get("name"))
    print("addRowsAPIEnabled:", data.get("addRowsAPIEnabled"))
    print("isRefreshable:", data.get("isRefreshable"))
    if data.get("addRowsAPIEnabled"):
        print("PUSH-CAPABLE")
    else:
        print("NOT push-capable")


if __name__ == "__main__":
    main()
