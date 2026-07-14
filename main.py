import os
import shutil
from pathlib import Path
from uuid import uuid4
from urllib.parse import urlencode
import pandas as pd

from fastapi import FastAPI, UploadFile, File, Form, Request, BackgroundTasks
from fastapi.responses import HTMLResponse, FileResponse, RedirectResponse
from starlette.middleware.sessions import SessionMiddleware
from authlib.integrations.starlette_client import OAuth, OAuthError

from in_kind_donations_cleaner import clean_multiple_in_kind_donations
from irfas_cleaner import clean_irfas
from volunteer_cleaner import clean_volunteers
from expenditure_cleaner import clean_expenditure
from revenue_cleaner import clean_multiple_revenue_imports
from current_revenue import clean_multiple_current_revenue_imports
from hunger_prevention import clean_multiple_hunger_prevention_imports


app = FastAPI()

# ------------------------------------------------------------
# ENVIRONMENT / SECURITY SETTINGS
# ------------------------------------------------------------

# Render automatically sets RENDER=true when the app is running on Render.
IS_RENDER = os.getenv("RENDER") == "true"

SESSION_SECRET_KEY = os.getenv("SESSION_SECRET_KEY")

AUTH0_DOMAIN = os.getenv("AUTH0_DOMAIN")
AUTH0_CLIENT_ID = os.getenv("AUTH0_CLIENT_ID")
AUTH0_CLIENT_SECRET = os.getenv("AUTH0_CLIENT_SECRET")
AUTH0_CALLBACK_URL = os.getenv("AUTH0_CALLBACK_URL")

APPROVED_EMAILS_RAW = os.getenv("APPROVED_EMAILS", "")

APPROVED_EMAILS = {
    email.strip().lower()
    for email in APPROVED_EMAILS_RAW.split(",")
    if email.strip()
}

# If deployed on Render, do NOT allow the app to launch without real secret values.
if IS_RENDER:
    missing_settings = []

    if not SESSION_SECRET_KEY:
        missing_settings.append("SESSION_SECRET_KEY")

    if not AUTH0_DOMAIN:
        missing_settings.append("AUTH0_DOMAIN")

    if not AUTH0_CLIENT_ID:
        missing_settings.append("AUTH0_CLIENT_ID")

    if not AUTH0_CLIENT_SECRET:
        missing_settings.append("AUTH0_CLIENT_SECRET")

    if not AUTH0_CALLBACK_URL:
        missing_settings.append("AUTH0_CALLBACK_URL")

    if not APPROVED_EMAILS:
        missing_settings.append("APPROVED_EMAILS")

    if missing_settings:
        raise RuntimeError(
            "Missing required Render environment variables: "
            + ", ".join(missing_settings)
        )

# Local-only fallback values.
# These let the app start locally, but Auth0 login will only work locally
# if you also set local Auth0 environment variables.
if not SESSION_SECRET_KEY:
    SESSION_SECRET_KEY = "local-test-session-secret-do-not-use-in-production"


# Add signed cookie session support.
# On Render, cookies are marked HTTPS-only.
app.add_middleware(
    SessionMiddleware,
    secret_key=SESSION_SECRET_KEY,
    session_cookie="master_cleaner_session",
    max_age=12 * 60 * 60,   # 12 hours
    same_site="lax",
    https_only=IS_RENDER
)

# ------------------------------------------------------------
# AUTH0 SETUP
# ------------------------------------------------------------

oauth = OAuth()

if AUTH0_DOMAIN and AUTH0_CLIENT_ID and AUTH0_CLIENT_SECRET:
    oauth.register(
        "auth0",
        client_id=AUTH0_CLIENT_ID,
        client_secret=AUTH0_CLIENT_SECRET,
        client_kwargs={
            "scope": "openid profile email",
        },
        server_metadata_url=f"https://{AUTH0_DOMAIN}/.well-known/openid-configuration",
    )


# ------------------------------------------------------------
# FILE STORAGE
# ------------------------------------------------------------

BASE_DIR = Path(__file__).parent

UPLOAD_DIR = BASE_DIR / "uploads"
OUTPUT_DIR = BASE_DIR / "outputs"

UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)


# ------------------------------------------------------------
# HELPER FUNCTIONS
# ------------------------------------------------------------

ALLOWED_REPORT_TYPES = {
    "irfas",
    "volunteers",
    "revenue",
    "grants",
    "hunger_prevention",
    "expenditure",
    "in_kind_donations",
    "current_revenue"
}


def user_is_authenticated(request):
    return request.session.get("user") is not None

def email_is_approved(email):
    if not email:
        return False

    return email.strip().lower() in APPROVED_EMAILS


def safe_delete_file(file_path: Path):
    """
    Delete a file if it exists.
    This helps avoid keeping uploaded/exported documents around after download.
    """
    try:
        if file_path.exists():
            file_path.unlink()
    except Exception:
        pass


def login_page_html(error_message: str = "") -> str:
    """
    Returns the HTML for the access-key login page.
    """
    error_html = ""

    if error_message:
        error_html = f"""
        <p style="color: #b00020; font-weight: bold;">
            {error_message}
        </p>
        """

    return f"""
    <html>
        <head>
            <title>Master Cleaner - Employee Access</title>
        </head>

        <body style="font-family: Arial, sans-serif; max-width: 650px; margin: 60px auto; line-height: 1.5;">
            <h2>ICNA Relief Master Cleaner</h2>

            <p>
                Enter the employee access key to continue.
            </p>

            {error_html}

            <form action="/login" method="post">
                <label for="access_key">Access Key:</label>
                <br>
                <input
                    type="password"
                    name="access_key"
                    id="access_key"
                    required
                    style="width: 320px; padding: 8px; margin-top: 6px;"
                >

                <br><br>

                <button type="submit" style="padding: 10px 16px;">
                    Enter
                </button>
            </form>
        </body>
    </html>
    """


def cleaner_page_html() -> str:
    """
    Returns the HTML for the main upload/cleaning page.
    """
    return """
    <html>
        <head>
            <title>ICNA Relief Master Cleaner</title>
        </head>

        <body style="font-family: Arial, sans-serif; max-width: 750px; margin: 60px auto; line-height: 1.5;">
            <div style="display: flex; justify-content: space-between; align-items: center;">
                <h2>ICNA Relief Master Cleaner</h2>
                <a href="/logout">Logout</a>
            </div>

            <p>
                Upload Excel files, choose the report type, <strong>follow the instructions before uploading</strong>, and download the processed result.
            </p>
            
            </p>

            <hr style="margin: 24px 0;">

            <form action="/clean" method="post" enctype="multipart/form-data">

                <label for="report_type"><strong>Select report type:</strong></label>
                <br>
                <select 
                    name="report_type" 
                    id="report_type"
                    required
                    style="padding: 8px; width: 320px; margin-top: 6px;"
                >
                    <option value="current_revenue">Current Revenue</option>
                    <option value="expenditure">Expenditure</option>
                    <option value="hunger_prevention">Hunger Prevention</option>
                    <option value="in_kind_donations">In-Kind Donations</option>
                    <option value="irfas">IRFAS</option>
                    <option value="revenue">Revenue</option>
                    <option value="volunteers">Volunteers</option>
                </select>

                <div id="instructions_section" style="margin-top: 20px;">

                    <div id="instructions_current_revenue" class="cleaner-instructions" style="display: none;">
                        <p>
                            <strong>Special Instructions for Current Revenue Uploads:</strong>
                            <br><br>
                            Before you export from Salesforce, turn Subtotals and Grand Total <strong>OFF</strong>.
                            <br><br>
                            Before you upload, rename all files showing total revenue for a specific month in the form:
                            <strong>month_year</strong>
                            <br>
                            Examples: <strong>january_2026.xlsx, FEB_2026.xlsx</strong>
                            <br><br>
                            For grants, rename the file in the form: <strong>grants_year</strong>
                            <br>
                            Example: <strong>grants_2026.xlsx</strong>
                            <br><br>
                            The name of both revenue files and grant files are <strong>NOT</strong> case sensitive.
                        </p>
                    </div>

                    <div id="instructions_expenditure" class="cleaner-instructions" style="display: none;">
                        <p>
                            <strong>Special Instructions for Expenditure Uploads:</strong>
                            <br><br>
                            Make sure there is <strong>ONLY ONE</strong> Worksheet. Delete all other Worksheets.
                            <br><br>
                            Find the row in the Worksheet originally containing values with Parenthesis
                            <br>
                            <strong>Examples: (GEORGIA REGION), (ADMINISTRATION):</strong>
                            <br>
                            You will rename the values in the row. These values will fall in 4 categories:
                            <br><br>
                            1. Central Expenses
                            <br>
                            2. Office Expenses
                            <br>
                            3. Grants
                            <br>
                            4. Office Expenses for the whole state (if the State is CA, TX, FL, VA, MO)
                            <br><br>
                            1. If the column refers to a National Expense, make sure the column starts with "Central."
                            <br>
                            Example: <strong>Central Cobra</strong>
                            <br><br>
                            2. If the column refers to Office expenses, make sure the column is the exact name of the office.
                            <br><br>
                            2A. If the state does not have an office, then make sure it ends in "-Statewide" with no spaces.
                            <br>
                            Example: <strong>New Hampshire-Statewide</strong>
                            <br><br>
                            2B. If the state has an office, then make sure it ends with " Office"
                            <br>
                            Example: <strong>Chicago Office</strong>
                            <br><br>
                            3. If the column refers to a grant, then the column will be the City (if an office is present) or State followed by " Grants".
                            <br>
                            Examples: <strong>Atlanta Grants, New Mexico Grants</strong>
                            <br><br>
                            4. If the column refers to office expenses for an entire state and the state has multiple field offices, then use the abbreviation for the state followed by " Divide".
                            <br>
                            Example: <strong>CA Divide</strong>
                            <br><br>
                            <strong>NOTE: Column names ARE case senstitive.</strong>
                        </p>
                    </div>
                    
                    <div id="instructions_hunger_prevention" class="cleaner-instructions" style="display: none;">
                        <p>
                            <strong>Special Instructions for Hunger Prevention Uploads:</strong>
                            <br><br>
                            Before you upload, here are special instructions for each file before you export:
                            <br><br>
                            1. Registered Clients:
                            <br>
                            Make sure the columns include the following:
                            <br><br>
                            <strong>Office Location</strong>
                            <br>
                            <strong>Date Distributed</strong>
                            <br><br>
                            Before you export from Salesforce, turn Subtotals and Grand Total <strong>OFF</strong>.
                            <br><br>
                            2. Bulk:
                            <br>
                            Make sure the columns include the following:
                            <br><br>
                            <strong>ICNA Relief Office: Account Name</strong>
                            <br>
                            <strong>Date Distributed</strong>
                            <br>
                            <strong>Bulk Distribution Event: Bulk Distribution Name</strong>
                            <br>
                            <strong>ICNA Relief Office</strong>
                            <br>
                            <strong>Billing State/Province</strong>
                            <br>
                            <strong>ICNA Relief Office: Billing City</strong>
                            <br>
                            <strong>ICNA Relief Office: Billing Zip/Postal Code</strong>
                            <br><br>
                            Before you export from Salesforce, turn Subtotals and Grand Total <strong>OFF</strong>.
                            <br><br>
                            3. Cash:
                            <br>
                            Make sure the columns include <strong>ONLY</strong> the following and no other columns:
                            <br><br>
                            <strong>Primary Campaign Source</strong>
                            <br>
                            <strong>Payment Amount Recieved</strong>
                            <br>
                            <strong>Close Date</strong>
                            <br>
                            <strong>Mailing State/Province</strong>
                            <br>
                            <strong>Mailing City</strong>
                            <br>
                            <strong>Mailing Zip/Postal Code</strong>
                            <br><br>
                            Before you export from Salesforce, turn Subtotals and Grand Total <strong>OFF</strong>.
                            <br><br>
                            4. In-Kind:
                            <br>
                            No instructions. Get a copy from Power BI.
                            <br><br>
                            5. IRFAS:
                            <br>
                            No instructions. Get a copy from Power BI.
                            <br><br>
                            Before you upload, make sure you have the following files and follow these specific naming
                            instructions:
                            <br><br>
                            1. Registered Clients: <strong>reg_clients.xlsx</strong> or <strong>registered_clients.xlsx</strong>
                            <br>
                            2. Bulk: <strong>bulk.xlsx</strong>
                            <br>
                            3. Cash: <strong>cash.xlsx</strong>
                            <br>
                            4. In-Kind: <strong>in-kind.xlsx</strong>
                            <br>
                            5. IRFAS: <strong>irfas.xlsx</strong>
                            <br><br>
                            File names are <strong>NOT</strong> case sensitive.
                        </p>
                    </div>
                    
                    <div id="instructions_revenue" class="cleaner-instructions" style="display: none;">
                        <p>
                            <strong>Special Instructions for Revenue Uploads:</strong>
                            <br><br>
                            Before you export in Salesforce, <strong>ADD</strong> these two columns:
                            <br><br>
                            <strong>Contact: Mailing Zip/Postal Code</strong>
                            <br>
                            <strong>Primary Campaign Source: Campaign Name</strong>
                            <br><br>
                            Make sure to turn Subtotals and Grand Total <strong>OFF</strong>.
                        </p>
                    </div>

                    <div id="instructions_in_kind_donations" class="cleaner-instructions" style="display: none;">
                        <p>
                            <strong>Special Instructions for In-Kind Donations Uploads:</strong>
                            <br><br>
                            Before you export in Salesforce, make sure these are the <strong>ONLY</strong> columns:
                            <br><br>
                            <strong>ICNA Relief Office: Account Name</strong>
                            <br>
                            <strong>Program</strong>
                            <br>
                            <strong>InKind ID</strong>
                            <br>
                            <strong>Donor Person: Full Name</strong>
                            <br>
                            <strong>Donor Organization: Account Name</strong>
                            <br>
                            <strong>Total Value</strong>
                            <br>
                            <strong>Total Items</strong>
                            <br>
                            <strong>Total Quantity of Items</strong>
                            <br><br>
                            Make sure to turn Subtotals and Grand Total <strong>OFF</strong> before you export.
                            <br><br>
                            Upload one or more In-Kind Donations Excel files. The cleaner will combine them into one cleaned output.
                        </p>
                    </div>

                    <div id="instructions_irfas" class="cleaner-instructions" style="display: none;">
                        <p>
                            <strong>Special Instructions for IRFAS Uploads:</strong>
                            <br><br>
                            Upload the <strong>EXACT</strong> raw data excel file recieved through email
                        </p>
                    </div>

                    <div id="instructions_volunteers" class="cleaner-instructions" style="display: none;">
                        <p>
                            <strong>Special Instructions for Volunteers Uploads:</strong>
                            <br><br>
                            Upload the volunteer Excel file you want cleaned. <strong>All data should only be for a single year.</strong>
                            <br><br>
                            The base year (last year's volunteering data) will be stored on GitHub to get retention rates.
                            <br><br>
                            Example: <strong>Upload 2026 Volunteering Data.</strong> GitHub has the 2025 Volunteering data already.
                            <br>
                            After every year, the GitHub base year file will need to be changed.
                        </p>
                    </div>

                </div>

                <div 
                    id="expenditure_month_section" 
                    style="display: none; margin-top: 18px;"
                >
                    <label for="report_month"><strong>Expenditure Month:</strong></label>
                    <br>

                    <input
                        type="month"
                        name="report_month"
                        id="report_month"
                        style="padding: 8px; width: 320px; margin-top: 6px;"
                    >

                    <p style="font-size: 13px; color: #555; margin-top: 6px;">
                        Select the month and year this expenditure file is for.
                    </p>
                </div>

                <br>

                <label for="uploaded_files"><strong>Upload Excel file(s):</strong></label>
                <br>
                <input 
                    type="file" 
                    name="uploaded_files" 
                    id="uploaded_files" 
                    accept=".xlsx"
                    multiple
                    required
                    style="margin-top: 6px;"
                >

                <br><br>

                <button type="submit" style="padding: 10px 16px;">
                    Run Cleaner
                </button>

            </form>

            <script>
                const reportTypeSelect = document.getElementById("report_type");
                const expenditureMonthSection = document.getElementById("expenditure_month_section");
                const reportMonthInput = document.getElementById("report_month");

                function updateExpenditureMonthVisibility() {
                    if (reportTypeSelect.value === "expenditure") {
                        expenditureMonthSection.style.display = "block";
                        reportMonthInput.required = true;
                    } else {
                        expenditureMonthSection.style.display = "none";
                        reportMonthInput.required = false;
                        reportMonthInput.value = "";
                    }
                }

                function updateCleanerInstructions() {
                    const instructionBlocks = document.querySelectorAll(".cleaner-instructions");

                    instructionBlocks.forEach(function(block) {
                        block.style.display = "none";
                    });

                    const selectedReportType = reportTypeSelect.value;
                    const selectedInstructions = document.getElementById(
                        "instructions_" + selectedReportType
                    );

                    if (selectedInstructions) {
                        selectedInstructions.style.display = "block";
                    }
                }

                function updatePageForSelectedCleaner() {
                    updateExpenditureMonthVisibility();
                    updateCleanerInstructions();
                }

                reportTypeSelect.addEventListener("change", updatePageForSelectedCleaner);

                updatePageForSelectedCleaner();
            </script>
        </body>
    </html>
    """


def error_page_html(message: str) -> str:
    """
    Returns a simple error page.
    """
    return f"""
    <html>
        <head>
            <title>Master Cleaner - Error</title>
        </head>

        <body style="font-family: Arial, sans-serif; max-width: 700px; margin: 60px auto; line-height: 1.5;">
            <h2>Something went wrong</h2>

            <p style="color: #b00020; font-weight: bold;">
                {message}
            </p>

            <p>
                <a href="/">Return to Master Cleaner</a>
            </p>
        </body>
    </html>
    """


# ------------------------------------------------------------
# ACCESS KEY / LOGIN ROUTES
# ------------------------------------------------------------

# ------------------------------------------------------------
# AUTH0 LOGIN / LOGOUT ROUTES
# ------------------------------------------------------------

@app.get("/login")
async def login(request: Request):
    if user_is_authenticated(request):
        return RedirectResponse(url="/", status_code=303)

    if not AUTH0_CALLBACK_URL:
        return HTMLResponse(
            content=error_page_html("Auth0 is not configured correctly."),
            status_code=500
        )

    return await oauth.auth0.authorize_redirect(
        request,
        redirect_uri=AUTH0_CALLBACK_URL
    )


@app.get("/callback")
async def auth0_callback(request: Request):
    try:
        token = await oauth.auth0.authorize_access_token(request)
    except OAuthError:
        return HTMLResponse(
            content=error_page_html("Login failed. Please try again."),
            status_code=401
        )

    user_info = token.get("userinfo")

    if not user_info:
        return HTMLResponse(
            content=error_page_html("Could not retrieve user information from Auth0."),
            status_code=401
        )

    email = user_info.get("email", "").strip().lower()

    if not email_is_approved(email):
        request.session.clear()

        return HTMLResponse(
            content=error_page_html(
                "Your login worked, but your email is not approved to access this tool."
            ),
            status_code=403
        )

    request.session["user"] = {
        "email": email,
        "name": user_info.get("name", email)
    }

    return RedirectResponse(url="/", status_code=303)


@app.get("/logout")
def logout(request: Request):
    request.session.clear()

    if not AUTH0_DOMAIN or not AUTH0_CLIENT_ID or not AUTH0_CALLBACK_URL:
        return RedirectResponse(url="/login", status_code=303)

    base_url = AUTH0_CALLBACK_URL.replace("/callback", "")

    logout_params = {
        "returnTo": base_url,
        "client_id": AUTH0_CLIENT_ID
    }

    auth0_logout_url = (
        f"https://{AUTH0_DOMAIN}/v2/logout?"
        + urlencode(logout_params)
    )

    return RedirectResponse(url=auth0_logout_url, status_code=303)


# ------------------------------------------------------------
# MAIN WEBSITE ROUTE
# ------------------------------------------------------------

@app.get("/", response_class=HTMLResponse)
def home(request: Request):
    # Anyone who has not entered the access key gets sent to the login page.
    if not user_is_authenticated(request):
        return RedirectResponse(url="/login", status_code=303)

    return cleaner_page_html()


# ------------------------------------------------------------
# CLEANING ROUTE
# ------------------------------------------------------------

@app.post("/clean")
def clean_uploaded_file(
    request: Request,
    background_tasks: BackgroundTasks,
    report_type: str = Form(...),
    uploaded_files: list[UploadFile] = File(...),
    report_month: str = Form(None)

):
    # Block access if the user has not passed the access-key screen.
    if not user_is_authenticated(request):
        return RedirectResponse(url="/login", status_code=303)

    # Confirm the dropdown value is one we explicitly allow.
    if report_type not in ALLOWED_REPORT_TYPES:
        return HTMLResponse(
            content=error_page_html("Invalid report type selected."),
            status_code=400
        )

    # Make sure at least one file was uploaded.
    if not uploaded_files:
        return HTMLResponse(
            content=error_page_html("No files were uploaded."),
            status_code=400
        )

    # Create a unique run ID so this upload batch does not overlap with another user's upload.
    run_id = str(uuid4())

    input_paths = []
    original_filenames = []

    # Save every uploaded file.
    for file_number, uploaded_file in enumerate(uploaded_files, start=1):

        # Skip empty file inputs.
        if not uploaded_file.filename:
            continue

        # For now, only allow .xlsx files.
        if not uploaded_file.filename.lower().endswith(".xlsx"):
            return HTMLResponse(
                content=error_page_html("Only .xlsx Excel files are allowed."),
                status_code=400
            )

        # Get only the filename, not a full path.
        original_filename = Path(uploaded_file.filename).name
        original_filenames.append(original_filename)

        # Include file_number so duplicate filenames do not overwrite each other.
        input_path = UPLOAD_DIR / f"{run_id}_{file_number}_{original_filename}"

        # Save uploaded file into the uploads folder.
        with open(input_path, "wb") as buffer:
            shutil.copyfileobj(uploaded_file.file, buffer)

        input_paths.append(input_path)

    if not input_paths:
        return HTMLResponse(
            content=error_page_html("No valid files were uploaded."),
            status_code=400
        )

    # Create one output file for the whole batch.
    output_path = OUTPUT_DIR / f"{run_id}_{report_type}_cleaned.xlsx"

    # --------------------------------------------------------
    # ROUTE FILES TO THE CORRECT CLEANER
    # --------------------------------------------------------

    if report_type == "in_kind_donations":
        clean_multiple_in_kind_donations(input_paths, output_path)
    elif report_type == "irfas":
        clean_irfas(input_paths[0], output_path)
    elif report_type == "volunteers":
        clean_volunteers(input_paths[0], output_path)
    elif report_type == "expenditure":
        if not report_month:
            return HTMLResponse(
                content=error_page_html("Please select the expenditure month."),
                status_code=400
            )

        clean_expenditure(
            input_paths[0],
            output_path,
            report_month
        )
    elif report_type == "revenue":
        clean_multiple_revenue_imports(input_paths, output_path)
    elif report_type == "current_revenue":
        clean_multiple_current_revenue_imports(input_paths, original_filenames, output_path)
    elif report_type == "hunger_prevention":
        clean_multiple_hunger_prevention_imports(input_paths, original_filenames,output_path)
    else:
        # Temporary fallback for other report types.
        # These will eventually be replaced with their own multi-file cleaner functions.
        # For now, just copy the first uploaded file.
        shutil.copy(input_paths[0], output_path)

    # After the response is sent, delete all uploaded files and the output file.
    for input_path in input_paths:
        background_tasks.add_task(safe_delete_file, input_path)

    background_tasks.add_task(safe_delete_file, output_path)

    # Send the output file back to the browser as a download.
    return FileResponse(
        path=output_path,
        filename=f"{report_type}_cleaned.xlsx",
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        background=background_tasks
    )

# 1. Create a DataFrame
data = {
    'Name': ['Alice', 'Bob', 'Charlie'],
    'Age': [25, 30, 35],
    'City': ['New York', 'London', 'Paris']
}
df = pd.DataFrame(data)

# 2. Export to Excel
df.to_excel('output.xlsx', index=False)

# ------------------------------------------------------------
# LOCAL RUN COMMAND
# ------------------------------------------------------------

if __name__ == "__main__":
    import uvicorn

    uvicorn.run(
        app,
        host="0.0.0.0",
        port=int(os.environ.get("PORT", 8000))
    )