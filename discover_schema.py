import requests

TENANT_ID = "b4c8bca6-b827-45e8-88a1-b98cced69842"
CLIENT_ID = "d8064124-ba13-4b42-920c-e63f191c2824"
CLIENT_SECRET = "PLACEHOLDER"
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


def execute_dax(token, dax_query):
    url = f"https://api.powerbi.com/v1.0/myorg/groups/{WORKSPACE_ID}/datasets/{DATASET_ID}/executeQueries"
    body = {
        "queries": [{"query": dax_query}],
        "serializerSettings": {"includeNulls": True},
    }
    response = requests.post(
        url,
        headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
        json=body,
    )
    if response.status_code != 200:
        print(f"Query failed: {response.status_code}")
        print(response.text)
        return None
    return response.json()


def main():
    token = get_access_token()
    print("Got access token OK.\n")

    print("=== TABLES ===")
    result = execute_dax(token, "EVALUATE INFO.VIEW.TABLES()")
    if result:
        rows = result["results"][0]["tables"][0]["rows"]
        for row in rows:
            print(row)

    print("\n=== COLUMNS ===")
    result = execute_dax(token, "EVALUATE INFO.VIEW.COLUMNS()")
    if result:
        rows = result["results"][0]["tables"][0]["rows"]
        for row in rows:
            print(row)


if __name__ == "__main__":
    main()
