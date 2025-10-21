import json
import requests

# === CONFIGURATION ===
site_url = "https://tromkom.sharepoint.com/sites/O365-ProsjektervedArkiv-BK-plan"  # No trailing slash
site_pages_list = "Områdesider"  # Try "Pages" or check actual list name
auth_json = "auth.json"
output_file = "page_urls.txt"

# === STEP 1: Extract Cookies from auth.json ===
def get_auth_headers():
    with open(auth_json, "r") as f:
        state = json.load(f)
    cookies = {c["name"]: c["value"] for c in state["cookies"]}
    headers = {
        "Cookie": "; ".join([f"{k}={v}" for k, v in cookies.items()]),
        "Accept": "application/json"
    }
    return headers

# === STEP 2: Get All .aspx Page URLs from Site Pages ===
def get_sharepoint_page_urls():
    headers = get_auth_headers()
    endpoint = f"{site_url}/_api/web/lists/getbytitle('{site_pages_list}')/items?$select=FileLeafRef&$top=1000"
    response = requests.get(endpoint, headers=headers)

    if response.status_code != 200:
        print(f"Klarte ikke å hente sider: {response.status_code}")
        print("Feilmelding:")
        print(response.text)
        return []

    try:
        data = response.json()
    except json.JSONDecodeError:
        print("Klarte ikke å tolke json-fil. Feilmelding:")
        print(response.text)
        return []

    pages = data.get("value", [])
    urls = [
        f"{site_url}/SitePages/{page['FileLeafRef']}"
        for page in pages if page.get("FileLeafRef", "").endswith(".aspx")
    ]
    return urls

# === STEP 3: Save URLs to File ===
def save_urls(urls):
    with open(output_file, "w") as f:
        for url in urls:
            f.write(url + "\n")
    print(f"✓ Lagret {len(urls)} URL til {output_file}")

# === MAIN EXECUTION ===
if __name__ == "__main__":
    urls = get_sharepoint_page_urls()
    save_urls(urls)