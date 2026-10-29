"""
Sync Performance Dashboard data from SharePoint Excel files → data.json on GitHub.

Reads all "Skill Will - *.xlsx" files from:
  tcapc.sharepoint.com/sites/performance
  └── Shared Documents/Skill Will/

Each file  = one SLT  (name parsed from filename)
Each sheet = one period (H1 2026, H2 2026, …) — "Legend" sheets skipped

Authenticates with Azure AD app (client credentials + Sites.Selected).
Outputs data.json at repo root for the dashboard to fetch.

GitHub Actions secrets required:
  AZURE_TENANT_ID
  AZURE_CLIENT_ID
  AZURE_CLIENT_SECRET
"""

import os
import json
import re
import sys
import msal
import requests
from datetime import date
from pathlib import Path

# ─── CONFIG ────────────────────────────────────────────────────────────────
SP_HOSTNAME  = "tcapc.sharepoint.com"
SP_SITE_PATH = "/sites/performance"
FOLDER_PATH  = "Shared Documents/Skill Will"   # under default drive root

GRAPH = "https://graph.microsoft.com/v1.0"

# Period sheet names must match this pattern: H1 2026, H2 2027, …
PERIOD_RE = re.compile(r"^H[12]\s+\d{4}$", re.IGNORECASE)


# ─── AUTH ──────────────────────────────────────────────────────────────────
def get_token():
    tenant = os.environ["AZURE_TENANT_ID"]
    app = msal.ConfidentialClientApplication(
        client_id=os.environ["AZURE_CLIENT_ID"],
        client_credential=os.environ["AZURE_CLIENT_SECRET"],
        authority=f"https://login.microsoftonline.com/{tenant}",
    )
    result = app.acquire_token_for_client(
        scopes=["https://graph.microsoft.com/.default"]
    )
    if "access_token" not in result:
        raise RuntimeError(
            f"Auth failed: {result.get('error')} - {result.get('error_description')}"
        )
    return result["access_token"]


# ─── GRAPH HELPERS ─────────────────────────────────────────────────────────
def gget(url, token):
    r = requests.get(
        url,
        headers={"Authorization": f"Bearer {token}"},
        timeout=30,
    )
    if r.status_code >= 400:
        raise RuntimeError(f"GET {url}\n→ {r.status_code}: {r.text[:400]}")
    return r.json()


# ─── SHAREPOINT HELPERS ────────────────────────────────────────────────────
def get_site_id(token):
    data = gget(f"{GRAPH}/sites/{SP_HOSTNAME}:{SP_SITE_PATH}", token)
    return data["id"]


def list_folder_files(token, site_id, folder_path):
    """Return list of Excel file items in the folder."""
    encoded = requests.utils.quote(folder_path)
    url = (
        f"{GRAPH}/sites/{site_id}/drive/root:/{encoded}:/children"
        f"?$select=id,name,parentReference&$filter=endswith(name,'.xlsx')"
    )
    data = gget(url, token)
    return data.get("value", [])


def list_worksheets(token, site_id, item_id):
    url = f"{GRAPH}/sites/{site_id}/drive/items/{item_id}/workbook/worksheets?$select=id,name"
    return gget(url, token).get("value", [])


def read_sheet(token, site_id, item_id, sheet_id):
    """Read usedRange values from a worksheet. Returns 2D list."""
    enc = requests.utils.quote(sheet_id)
    url = (
        f"{GRAPH}/sites/{site_id}/drive/items/{item_id}"
        f"/workbook/worksheets('{enc}')/usedRange?$select=values"
    )
    return gget(url, token).get("values", [])


# ─── DATA TRANSFORM ────────────────────────────────────────────────────────
def parse_slt_name(filename):
    """'Skill Will - Amy Nguyen.xlsx' → 'Amy Nguyen'"""
    name = re.sub(r"^Skill Will\s*-\s*", "", filename, flags=re.IGNORECASE)
    name = re.sub(r"\.xlsx$", "", name, flags=re.IGNORECASE)
    return name.strip()


def safe_float(val):
    try:
        return float(val) if val not in (None, "", " ") else None
    except (TypeError, ValueError):
        return None


def safe_str(val):
    if val is None:
        return ""
    s = str(val).strip()
    return s if s not in ("-", "N/A") else ""


def parse_date(val):
    """Convert Excel serial date or date string to ISO format string."""
    if val is None or val == "":
        return ""
    if isinstance(val, (int, float)):
        # Excel serial date (days since 1900-01-01, with leap year bug)
        try:
            from datetime import timedelta
            base = date(1899, 12, 30)
            return (base + timedelta(days=int(val))).isoformat()
        except Exception:
            return ""
    return str(val).strip()


def read_file_data(token, site_id, file_item, slt_name):
    """Read all period sheets from one SLT's Excel file. Returns list of rows."""
    item_id = file_item["id"]
    rows = []

    sheets = list_worksheets(token, site_id, item_id)
    for sheet in sheets:
        sname = sheet["name"].strip()

        # Skip non-period sheets
        if not PERIOD_RE.match(sname):
            print(f"    skip sheet: {sname}")
            continue

        period = sname
        values = read_sheet(token, site_id, item_id, sheet["id"])
        if len(values) < 2:
            print(f"    {period}: empty")
            continue

        headers = [str(h or "").strip() for h in values[0]]

        def col(name):
            try:
                return headers.index(name)
            except ValueError:
                return -1

        i_ukgid    = col("UKG_ID")
        i_full     = col("Full_Name")
        i_display  = col("Display_Name")
        i_dept     = col("Department")
        i_title    = col("Job_Title")
        i_sup      = col("Supervisor")
        i_email    = col("Email")
        i_hire     = col("Hire_Date")
        i_platform = col("Platform")
        i_skill    = col("Skill_Score")
        i_will     = col("Will_Score")
        i_pr       = col("PR_Score")
        i_color    = col("Color_Code")
        i_just     = col("Justification")

        period_rows = 0
        for row in values[1:]:
            def v(idx):
                return row[idx] if idx >= 0 and idx < len(row) else None

            ukgid = safe_str(v(i_ukgid))
            if not ukgid:
                continue

            rows.append({
                "UKG_ID"         : ukgid,
                "Full_Name"      : safe_str(v(i_full)) or safe_str(v(i_display)),
                "Display_Name"   : safe_str(v(i_display)) or safe_str(v(i_full)),
                "Department"     : safe_str(v(i_dept)),
                "Job_Title"      : safe_str(v(i_title)),
                "Supervisor"     : safe_str(v(i_sup)),
                "Email"          : safe_str(v(i_email)),
                "Hire_Date"      : parse_date(v(i_hire)),
                "Platform"       : safe_str(v(i_platform)),
                "SLT_Name"       : slt_name,
                "Period"         : period,
                "Skill_Score"    : safe_float(v(i_skill)),
                "Will_Score"     : safe_float(v(i_will)),
                "PR_Score"       : safe_float(v(i_pr)),
                "Color_Code"     : safe_str(v(i_color)),
                "Justification"  : safe_str(v(i_just)),
                # FR fields — placeholder until FR Excel added
                "Frustration"      : None,
                "Flight_Potential" : None,
                "FR_Score"         : None,
            })
            period_rows += 1

        print(f"    {period}: {period_rows} rows")

    return rows


# ─── MAIN ──────────────────────────────────────────────────────────────────
def main():
    print("→ Authenticating…")
    token = get_token()
    print("  ✅ Token acquired")

    print(f"→ Resolving site: {SP_HOSTNAME}{SP_SITE_PATH}")
    site_id = get_site_id(token)
    print(f"  ✅ Site ID: {site_id[:50]}…")

    print(f"→ Listing Excel files in '{FOLDER_PATH}'…")
    files = list_folder_files(token, site_id, FOLDER_PATH)
    print(f"  ✅ Found {len(files)} file(s): {[f['name'] for f in files]}")

    all_rows = []
    for f in files:
        slt_name = parse_slt_name(f["name"])
        print(f"  Reading: {f['name']} → SLT: {slt_name}")
        rows = read_file_data(token, site_id, f, slt_name)
        all_rows.extend(rows)
        print(f"    → {len(rows)} rows total")

    print(f"→ Total rows: {len(all_rows)}")

    out_path = Path(__file__).parent.parent / "data.json"
    out_path.write_text(
        json.dumps(all_rows, indent=2, ensure_ascii=False),
        encoding="utf-8",
    )
    print(f"→ Wrote {out_path} ({out_path.stat().st_size:,} bytes)")


if __name__ == "__main__":
    main()
