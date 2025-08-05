from flask import Flask, request, jsonify
import requests
from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext

app = Flask(__name__)

# ---------------------
# 🔧 Zoho Configuration
# ---------------------
ZOHO_CLIENT_ID = "1000.33I8RMB265FK27LJWQJMYFM44Z6ZAC"
ZOHO_CLIENT_SECRET = "2470896664ea5a567864a3a813952009ea726afaad"
ZOHO_REFRESH_TOKEN = "1000.7c662bcbff88ef8f221b1eb3461484a9.a223501fa87a58c1049f1359723ec0d1"
ZOHO_API_BASE = "https://recruit.zoho.com"
REDIRECT_URI = "http://localhost"
MODULE = "Candidates"

# ---------------------
# 🔧 SharePoint Configuration
# ---------------------
SP_SITE_URL = "https://nahdetmisrcom.sharepoint.com/sites/ZOHO-Test/"
SP_USERNAME = "sp.zoho@nahdetmisr.com"
SP_PASSWORD = "Z0ho@SP0nline#123"
SP_LIST_NAME = "Recruitment tracker Test"

# ---------------------
# 🔐 Get Zoho Access Token
# ---------------------
def get_access_token():
    url = "https://accounts.zoho.com/oauth/v2/token"
    params = {
        "refresh_token": ZOHO_REFRESH_TOKEN,
        "client_id": ZOHO_CLIENT_ID,
        "client_secret": ZOHO_CLIENT_SECRET,
        "grant_type": "refresh_token",
        "redirect_uri": REDIRECT_URI,
    }
    response = requests.post(url, params=params)
    response.raise_for_status()
    data = response.json()
    print("🔑 Access Token Retrieved Successfully")
    return data["access_token"]

# ---------------------
# 📥 Fetch Candidates from Zoho
# ---------------------
def fetch_candidates(access_token):
    url = f"{ZOHO_API_BASE}/recruit/v2/{MODULE}"
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Accept": "application/json"
    }
    response = requests.get(url, headers=headers)
    print("Status Code:", response.status_code)
    if not response.ok:
        print("📩 Response Text:", response.text)
        response.raise_for_status()
    return response.json().get("data", [])

# ---------------------
# 🔎 Get Existing Titles from SharePoint
# ---------------------
def get_existing_titles(ctx, sp_list):
    existing_titles = set()
    items = sp_list.items
    ctx.load(items)
    ctx.execute_query()
    for item in items:
        existing_titles.add(item.properties.get("Title"))
    return existing_titles

# ---------------------
# 📤 Post to SharePoint
# ---------------------
def post_to_sharepoint(data):
    ctx_auth = AuthenticationContext(SP_SITE_URL)
    if ctx_auth.acquire_token_for_user(SP_USERNAME, SP_PASSWORD):
        ctx = ClientContext(SP_SITE_URL, ctx_auth)
        sp_list = ctx.web.lists.get_by_title(SP_LIST_NAME)

        print("📋 Checking existing records in SharePoint...")
        existing_titles = get_existing_titles(ctx, sp_list)

        for item in data:
            full_name = item.get("Full_Name", "NoName")
            if full_name in existing_titles:
                print(f"⚠️ Skipped duplicate: {full_name}")
                continue

            sp_item = {
                "Title": full_name,
                "EmpPosition": item.get("Posting_Title", "NoPosition"),
                "DepartmentName": item.get("Department_Name", {}).get("name", "NoDept"),
                "Mobile_Number": item.get("Mobile", "N/A"),
                "Email": item.get("Email", "N/A"),
                "JobID": item.get("Job_Opening_ID", "N/A"),
                "HiringDate": item.get("Date_Of_Hired", "N/A"),
                "Location": item.get("Location", "N/A"),
                "Gender": item.get("Gender", "N/A"),
                "Military_Service": item.get("Military_Service_Status", "N/A"),
                "ID_Number": item.get("ID_Number", "N/A"),
                "Hiring_Requester": item.get("Candidate_Owner", {}).get("name", "N/A"),
                "Hiring_Type": item.get("Reason_of_Requisition", "N/A"),
                "Replacement": item.get("Replacement", "N/A")
            }
            sp_list.add_item(sp_item)
            ctx.execute_query()
            print(f"✅ Added to SharePoint: {sp_item['Title']}")
    else:
        print("❌ SharePoint Authentication failed")

# ---------------------
# 🌐 API Endpoint
# ---------------------
@app.route("/zoho", methods=["POST"])
def handle_zoho_data():
    try:
        print("🔐 Fetching Zoho access token...")
        token = get_access_token()

        print("📦 Fetching candidates data from Zoho Recruit...")
        candidates = fetch_candidates(token)

        print(f"📤 Posting {len(candidates)} records to SharePoint...")
        post_to_sharepoint(candidates)

        return jsonify({"status": "success", "records_posted": len(candidates)}), 200

    except Exception as e:
        print("❌ Error:", str(e))
        return jsonify({"status": "error", "message": str(e)}), 500

# ---------------------
# 🚀 Run Flask App
# ---------------------
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
