"""
Sync Performance Dashboard data from SharePoint Excel files → data.json on GitHub.

Reads "Skill Will - *.xlsx" from:
  tcapc.sharepoint.com/sites/performance
  └── Shared Documents/Performance Dashboard/Skill Will/

Reads "Flight Risk Matrix - *.xlsx" from:
  tcapc.sharepoint.com/sites/performance
  └── Shared Documents/Performance Dashboard/Flight Risk/

Each file  = one SLT  (name parsed from filename)
Each sheet = one period (H1 2026, H2 2026, …) — non-period sheets skipped

FR data is joined to SW rows by (UKG_ID, Period).

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
import msal
import requests
from datetime import date
from pathlib import Path

# ─── CONFIG ────────────────────────────────────────────────────────────────
SP_HOSTNAME  = "tcapc.sharepoint.com"
SP_SITE_PATH = "/sites/performance"
SW_FOLDER    = "Performance Dashboard/Skill Will"
FR_FOLDER    = "Performance Dashboard/Flight Risk"

GRAPH = "https://graph.microsoft.com/v1.0"

# Period sheet names must match: H1 2026, H2 2027, …
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
    encoded = requests.utils.quote(folder_path)
    url = f"{GRAPH}/sites/{site_id}/drive/root:/{encoded}:/children?$select=id,name"
    data = gget(url, token)
    return [f for f in data.get("value", []) if f["name"].lower().endswith(".xlsx")]


def list_worksheets(token, site_id, item_id):
    url = f"{GRAPH}/sites/{site_id}/drive/items/{item_id}/workbook/worksheets?$select=id,name"
    return gget(url, token).get("value", [])


def read_sheet(token, site_id, item_id, sheet_id):
    enc = requests.utils.quote(sheet_id)
    url = (
        f"{GRAPH}/sites/{site_id}/drive/items/{item_id}"
        f"/workbook/worksheets('{enc}')/usedRange?$select=values"
    )
    return gget(url, token).get("values", [])


# ─── DATA TRANSFORM ────────────────────────────────────────────────────────
def parse_slt_name(filename, prefix):
    """'Skill Will - Amy Nguyen.xlsx' → 'Amy Nguyen' (given prefix='Skill Will')"""
    name = re.sub(rf"^{re.escape(prefix)}\s*-\s*", "", filename, flags=re.IGNORECASE)
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
    if val is None or val == "":
        return ""
    if isinstance(val, (int, float)):
        try:
            from datetime import timedelta
            base = date(1899, 12, 30)
            return (base + timedelta(days=int(val))).isoformat()
        except Exception:
            return ""
    return str(val).strip()


def col_idx(headers, name):
    try:
        return headers.index(name)
    except ValueError:
        return -1


def v(row, idx):
    return row[idx] if 0 <= idx < len(row) else None


# ─── READ SKILL WILL ───────────────────────────────────────────────────────
def read_sw_file(token, site_id, file_item, slt_name):
    item_id = file_item["id"]
    rows = []
    for sheet in list_worksheets(token, site_id, item_id):
        sname = sheet["name"].strip()
        if not PERIOD_RE.match(sname):
            print(f"    skip sheet: {sname}")
            continue
        values = read_sheet(token, site_id, item_id, sheet["id"])
        if len(values) < 2:
            continue
        headers = [str(h or "").strip() for h in values[0]]
        i_ukgid    = col_idx(headers, "UKG_ID")
        i_full     = col_idx(headers, "Full_Name")
        i_display  = col_idx(headers, "Display_Name")
        i_dept     = col_idx(headers, "Department")
        i_title    = col_idx(headers, "Job_Title")
        i_sup      = col_idx(headers, "Supervisor")
        i_email    = col_idx(headers, "Email")
        i_hire     = col_idx(headers, "Hire_Date")
        i_platform = col_idx(headers, "Platform")
        i_skill    = col_idx(headers, "Skill_Score")
        i_will     = col_idx(headers, "Will_Score")
        i_pr       = col_idx(headers, "PR_Score")
        i_color    = col_idx(headers, "Color_Code")
        i_just     = col_idx(headers, "Justification")
        count = 0
        for row in values[1:]:
            ukgid = safe_str(v(row, i_ukgid))
            if not ukgid:
                continue
            rows.append({
                "UKG_ID"           : ukgid,
                "Full_Name"        : safe_str(v(row, i_full)) or safe_str(v(row, i_display)),
                "Display_Name"     : safe_str(v(row, i_display)) or safe_str(v(row, i_full)),
                "Department"       : safe_str(v(row, i_dept)),
                "Job_Title"        : safe_str(v(row, i_title)),
                "Supervisor"       : safe_str(v(row, i_sup)),
                "Email"            : safe_str(v(row, i_email)),
                "Hire_Date"        : parse_date(v(row, i_hire)),
                "Platform"         : safe_str(v(row, i_platform)),
                "SLT_Name"         : slt_name,
                "Period"           : sname,
                "Skill_Score"      : safe_float(v(row, i_skill)),
                "Will_Score"       : safe_float(v(row, i_will)),
                "PR_Score"         : safe_float(v(row, i_pr)),
                "Color_Code"       : safe_str(v(row, i_color)),
                "Justification"    : safe_str(v(row, i_just)),
                "Frustration"      : None,
                "Flight_Potential" : None,
                "FR_Score"         : None,
            })
            count += 1
        print(f"    {sname}: {count} rows")
    return rows


# ─── READ FLIGHT RISK ──────────────────────────────────────────────────────
def read_fr_files(token, site_id):
    """Return dict keyed by (ukgid, period) → {Frustration, Flight_Potential, FR_Score}"""
    lookup = {}
    try:
        files = list_folder_files(token, site_id, FR_FOLDER)
    except RuntimeError as e:
        print(f"  ⚠️  Flight Risk folder not accessible: {e}")
        return lookup

    print(f"  ✅ Found {len(files)} FR file(s): {[f['name'] for f in files]}")
    for file_item in files:
        slt_name = parse_slt_name(file_item["name"], "Flight Risk Matrix")
        print(f"  Reading FR: {file_item['name']} → SLT: {slt_name}")
        item_id = file_item["id"]
        for sheet in list_worksheets(token, site_id, item_id):
            sname = sheet["name"].strip()
            if not PERIOD_RE.match(sname):
                continue
            values = read_sheet(token, site_id, item_id, sheet["id"])
            if len(values) < 2:
                continue
            headers = [str(h or "").strip() for h in values[0]]
            i_ukgid  = col_idx(headers, "UKG_ID")
            i_frust  = col_idx(headers, "Frustration")
            i_flight = col_idx(headers, "Flight_Potential")
            i_frscore = col_idx(headers, "FR_Score")
            count = 0
            for row in values[1:]:
                ukgid = safe_str(v(row, i_ukgid))
                if not ukgid:
                    continue
                lookup[(ukgid, sname)] = {
                    "Frustration"      : safe_float(v(row, i_frust)),
                    "Flight_Potential" : safe_float(v(row, i_flight)),
                    "FR_Score"         : safe_float(v(row, i_frscore)),
                }
                count += 1
            print(f"    {sname}: {count} FR rows")
    return lookup


# ─── MAIN ──────────────────────────────────────────────────────────────────
def main():
    print("→ Authenticating…")
    token = get_token()
    print("  ✅ Token acquired")

    print(f"→ Resolving site: {SP_HOSTNAME}{SP_SITE_PATH}")
    site_id = get_site_id(token)
    print(f"  ✅ Site ID: {site_id[:50]}…")

    # ── Skill Will ──
    print(f"→ Listing Skill Will files in '{SW_FOLDER}'…")
    sw_files = list_folder_files(token, site_id, SW_FOLDER)
    print(f"  ✅ Found {len(sw_files)} file(s): {[f['name'] for f in sw_files]}")

    all_rows = []
    for f in sw_files:
        slt_name = parse_slt_name(f["name"], "Skill Will")
        print(f"  Reading SW: {f['name']} → SLT: {slt_name}")
        rows = read_sw_file(token, site_id, f, slt_name)
        all_rows.extend(rows)
        print(f"    → {len(rows)} rows total")

    # ── Flight Risk ──
    print(f"→ Reading Flight Risk files from '{FR_FOLDER}'…")
    fr_lookup = read_fr_files(token, site_id)
    print(f"  ✅ FR lookup: {len(fr_lookup)} (ukgid, period) entries")

    # ── Join FR into SW rows ──
    merged = 0
    for row in all_rows:
        key = (row["UKG_ID"], row["Period"])
        if key in fr_lookup:
            row.update(fr_lookup[key])
            merged += 1
    print(f"→ Merged FR data into {merged} / {len(all_rows)} rows")

    out_path = Path(__file__).parent.parent / "data.json"
    out_path.write_text(
        json.dumps(all_rows, indent=2, ensure_ascii=False),
        encoding="utf-8",
    )
    print(f"→ Wrote {out_path} ({out_path.stat().st_size:,} bytes)")


if __name__ == "__main__":
    main()
