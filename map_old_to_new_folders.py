"""
Map old SharePoint employee folders (organized by country) to new folder paths
(organized by pay schedule) using the BambooHR roster.

Old layout (discovered by listing SharePoint):
  {SHAREPOINT_FOLDER_PATH}/{Country}/.../{employee folder}/

New layout (from sync_employee_files_to_sharepoint.py):
  {SHAREPOINT_FOLDER_PATH}/{PaySchedule}/{Last, First}/

This script:
  1. Connects to SharePoint and lists the children of the root folder.
  2. Recursively explores each country directory to find leaf employee folders.
  3. Fetches the BambooHR roster (report 135) for employee names, countries,
     and pay schedules.
  4. Matches old folders to employees by name similarity.
  5. Outputs a JSON mapping of old_path -> new_path.

Required env vars (same as sync_employee_files_to_sharepoint.py):
  BAMBOO_HR_KEY
  AZURE_APP_CLIENT_ID
  AZURE_TENANT_ID
  AZURE_CLIENT_SECRET
  SHAREPOINT_HOSTNAME
  SHAREPOINT_SITE_PATH
  SHAREPOINT_DRIVE_NAME
  SHAREPOINT_FOLDER_PATH

Usage:
    python map_old_to_new_folders.py                         # human-readable console output for test group
    python map_old_to_new_folders.py --csv                   # writes folder_mapping_test.csv + employee_subfolders.csv
    python map_old_to_new_folders.py --json                  # machine-readable JSON to stdout
    python map_old_to_new_folders.py --all-users --csv       # writes full folder_mapping.csv
    python map_old_to_new_folders.py --csv --use-test-group  # explicit test-group mode
"""

from __future__ import annotations

import csv
import json
import os
import re
import sys
from urllib.parse import quote

import requests
from msal import ConfidentialClientApplication

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

# --- Constants ----------------------------------------------------------------

BAMBOO_COMPANY = "techsoft3d"
BAMBOO_BASE = f"https://api.bamboohr.com/api/gateway.php/{BAMBOO_COMPANY}/v1"
GRAPH = "https://graph.microsoft.com/v1.0"


# --- Test group helpers -------------------------------------------------------

def load_test_users() -> set[str]:
    """Load test user emails from test_users.json, return as lowercase set."""
    try:
        with open("test_users.json") as f:
            data = json.load(f)
            emails = data.get("test_emails", [])
            return {e.lower() for e in emails if e}
    except FileNotFoundError:
        print("ERROR: test_users.json not found", file=sys.stderr)
        return set()
    except json.JSONDecodeError as e:
        print(f"ERROR: test_users.json is invalid JSON: {e}", file=sys.stderr)
        return set()


def _mask_filename(filename: str) -> str:
    """Mask filename showing only first 4 and last 4 chars (for GitHub output)."""
    if not os.environ.get("GITHUB_ACTIONS"):
        return filename
    if len(filename) <= 8:
        return filename
    # Split on / to handle paths, mask each component
    parts = filename.split("/")
    masked_parts = []
    for part in parts:
        if len(part) <= 8:
            masked_parts.append(part)
        else:
            masked = part[:4] + "*" * (len(part) - 8) + part[-4:]
            masked_parts.append(masked)
    return "/".join(masked_parts)


# --- Auth helpers (shared with sync script) -----------------------------------

def bamboo_session(api_key: str) -> requests.Session:
    s = requests.Session()
    s.auth = (api_key, "x")
    s.headers.update({"Accept": "application/json"})
    return s


def graph_token() -> str:
    app = ConfidentialClientApplication(
        os.environ["AZURE_APP_CLIENT_ID"],
        authority=f"https://login.microsoftonline.com/{os.environ['AZURE_TENANT_ID']}",
        client_credential=os.environ["AZURE_CLIENT_SECRET"],
    )
    result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    if "access_token" not in result:
        raise RuntimeError(f"Failed to acquire Graph token: {result}")
    return result["access_token"]


def graph_session(token: str) -> requests.Session:
    s = requests.Session()
    s.headers.update({"Authorization": f"Bearer {token}", "Accept": "application/json"})
    return s


def resolve_site_id(session: requests.Session, hostname: str, site_path: str) -> str:
    site_path = "/" + site_path.strip("/")
    r = session.get(f"{GRAPH}/sites/{hostname}:{site_path}")
    r.raise_for_status()
    return r.json()["id"]


def resolve_drive_id(session: requests.Session, site_id: str, drive_name: str) -> str:
    r = session.get(f"{GRAPH}/sites/{site_id}/drives")
    r.raise_for_status()
    drives = r.json().get("value", [])
    for d in drives:
        if d.get("name") == drive_name:
            return d["id"]
    raise RuntimeError(
        f"Drive '{drive_name}' not found. Available: "
        + ", ".join(d.get("name", "?") for d in drives)
    )


# --- SharePoint listing -------------------------------------------------------

def list_children(session: requests.Session, drive_id: str, folder_path: str) -> list[dict]:
    """List immediate children (folders and files) of a SharePoint folder."""
    encoded = quote(folder_path.strip("/"))
    url = f"{GRAPH}/drives/{drive_id}/root:/{encoded}:/children"
    items = []
    while url:
        r = session.get(url)
        r.raise_for_status()
        data = r.json()
        items.extend(data.get("value", []))
        url = data.get("@odata.nextLink")
    return items


def _looks_like_employee_name(name: str) -> bool:
    """
    Return True if a folder name looks like a person's name rather than
    a category/container folder.

    Employee patterns:
      "Last, First"                  (US/UK style)
      "Last, First_DOH_..."          (US/UK with date of hire)
      "LASTNAME Firstname"           (European style, e.g. BRANDT Frode)
      "LASTNAME Firstname_DOH ..."   (European with date of hire)
      "Firstname Lastname"           (informal, e.g. Andrew Choi)

    Non-employee:
      "Interns, Apprentices, Temps", "Theorem Solutions, Inc",
      "1-Identity", "Employees", "Archive", "Templates",
      "_CONTRACTORS_and_TEMPS", "_CONFIDENTIAL Reports..."
    """
    # If it starts with a digit or underscore, it's a category/system folder
    if name and (name[0].isdigit() or name[0] == "_"):
        return False

    # Known non-employee folder names (case-insensitive)
    _NON_EMPLOYEE = {
        "employees", "archive", "templates", "contractor", "contractors",
        "organisation personnel file", "registre du personnel - staff register",
        "spinfire", "theorem", "tech soft 3d ltd", "industrial apps group",
        "toolkits group",
        # Countries / regions that could look like two-word names
        "czech republic", "united states", "united kingdom", "south korea",
        "new zealand", "costa rica", "puerto rico", "hong kong",
        "saudi arabia", "south africa", "sri lanka",
    }
    if name.lower().strip() in _NON_EMPLOYEE:
        return False

    # Contains DOH/DOG (date of hire) -> strong employee signal
    if re.search(r"\bDOH\b|_DOH|_DOG\b", name, re.I):
        return True

    # Comma-based detection (US/UK style: "Last, First")
    if "," in name:
        parts = name.split(",", 1)
        last_part = parts[0].strip()
        first_part = parts[1].strip()
        # Corporate suffixes
        if first_part.lower() in ("inc", "inc.", "llc", "ltd", "ltd."):
            return False
        # Numbered category that happens to have a comma
        if last_part and last_part[0].isdigit():
            return False
        # Multi-word list: "Interns, Apprentices, Temps"
        if name.count(",") >= 2:
            if " " in last_part and not re.search(r"DOH|DOG", name, re.I):
                return False
        # Category phrases: lowercase multi-word before comma
        words_before = last_part.split()
        if len(words_before) >= 2 and all(w[0].islower() for w in words_before if w):
            return False
        return True

    # European style: "LASTNAME Firstname" (first word all-caps, 2+ chars)
    # Also matches "LASTNAME Firstname_suffix"
    words = name.split()
    if len(words) >= 2:
        first_word = words[0]
        # All-uppercase first word of 2+ alpha chars -> likely a surname
        if len(first_word) >= 2 and first_word.isalpha() and first_word.isupper():
            return True
        # Also catch "Firstname Lastname" where both are capitalized names
        # (e.g. "Andrew Choi", "William Zu", "Francisco Cardoso_DOH...")
        if all(w.split("_")[0][0:1].isupper() for w in words[:2]):
            # But not if it looks like a title/category
            lower_name = name.lower()
            if any(kw in lower_name for kw in ("group", "file", "report",
                                                "template", "register", "organisation")):
                return False
            return True

    return False


def find_employee_folders(session: requests.Session, drive_id: str,
                          root_path: str, max_depth: int = 3
                          ) -> tuple[list[str], dict[str, list[str]]]:
    """
    Recursively discover employee-level folders under root_path.

    Heuristic: an "employee folder" is a folder whose name looks like a
    person's name (contains a comma in "Last, First" style).

    Folders that don't match are treated as intermediate directories (country,
    group, etc.) and are recursed into up to max_depth.

    Returns:
      - list of relative employee folder paths from root_path
      - dict mapping each employee folder rel path -> list of subfolder names inside it
    """
    employee_folders: list[str] = []
    employee_subfolders: dict[str, list[str]] = {}
    root_stripped = root_path.rstrip("/")

    def _walk(path: str, depth: int):
        if depth > max_depth:
            return
        print(f"  [depth={depth}] Listing: {path}", file=sys.stderr)
        children = list_children(session, drive_id, path)
        folders = [c for c in children if "folder" in c]
        print(f"           -> {len(folders)} subfolder(s)", file=sys.stderr)

        for folder in folders:
            name = folder["name"]
            child_path = f"{path}/{name}"

            if _looks_like_employee_name(name):
                rel = child_path[len(root_stripped):].strip("/")
                employee_folders.append(rel)
                # List subfolders inside this employee folder
                try:
                    emp_children = list_children(session, drive_id, child_path)
                    sub_dirs = [c["name"] for c in emp_children if "folder" in c]
                    employee_subfolders[rel] = sub_dirs
                except Exception:
                    employee_subfolders[rel] = []
            else:
                # Intermediate folder (country, group, etc.) - recurse
                _walk(child_path, depth + 1)

    _walk(root_stripped, 0)
    return employee_folders, employee_subfolders


# --- BambooHR roster ----------------------------------------------------------

def fetch_roster(session: requests.Session) -> list[dict]:
    r = session.get(
        f"{BAMBOO_BASE}/reports/135",
        params={"format": "json", "fd": "yes", "onlyCurrent": "true"},
    )
    r.raise_for_status()
    return r.json().get("employees", []) or []


# --- Matching -----------------------------------------------------------------

def normalize_name(name: str) -> str:
    """Normalize a name for fuzzy matching: lowercase, strip suffixes like _DOH_..., extra spaces."""
    name = name.lower().strip()
    # Remove _DOH... suffix and everything after it
    name = re.sub(r"[_\s]*doh[_\s].*$", "", name)
    # Remove _DOG... suffix
    name = re.sub(r"[_\s]*dog[_\s].*$", "", name)
    # Remove other trailing annotations like "_moved to India", "_Term..."
    name = re.sub(r"_moved\s.*$", "", name)
    name = re.sub(r"_term[_\s].*$", "", name, flags=re.I)
    # Remove any trailing underscores/spaces/hyphens
    name = name.strip(" _-")
    return name


def _parse_folder_to_last_first(folder_name: str) -> str | None:
    """
    Parse an employee folder name into normalized 'last, first' format.
    Handles both "Last, First..." and "LASTNAME Firstname..." styles.
    """
    normalized = normalize_name(folder_name)

    # If it contains a comma, it's already "last, first" style
    if "," in normalized:
        parts = normalized.split(",", 1)
        last = parts[0].strip()
        first = parts[1].strip().split()[0] if parts[1].strip() else ""
        return f"{last}, {first}"

    # European style: "LASTNAME Firstname" -> "lastname, firstname"
    # Also handles multi-word last names like "DE GALZAIN Isabelle"
    words = normalized.split()
    if len(words) >= 2:
        # Find where the uppercase surname ends and first name begins
        # In the normalized (lowercased) form, we rely on the original casing
        orig_words = folder_name.split()
        # Strip suffixes from last word
        orig_words_clean = []
        for w in orig_words:
            cleaned = re.sub(r"[_].*$", "", w).strip()
            if cleaned:
                orig_words_clean.append(cleaned)

        # Find the split: uppercase words are surname, first capitalized word is firstname
        surname_parts = []
        firstname_parts = []
        found_first = False
        for w in orig_words_clean:
            if not found_first and w.isupper() and w.isalpha():
                surname_parts.append(w)
            elif not found_first and w[0:1].isupper():
                # Could be start of first name or part of multi-word surname
                # If previous words were all caps, this is likely the first name
                if surname_parts:
                    found_first = True
                    firstname_parts.append(w)
                else:
                    # No all-caps surname yet - "Firstname Lastname" style
                    firstname_parts.append(w)
                    found_first = True
            elif found_first:
                firstname_parts.append(w)
            else:
                surname_parts.append(w)

        if surname_parts and firstname_parts:
            last = " ".join(surname_parts).lower()
            first = firstname_parts[0].lower()
            # Remove parenthetical maiden names like "(CISSE)"
            first = re.sub(r"\(.*?\)", "", first).strip()
            return f"{last}, {first}"

        # Fallback for "Firstname Lastname" style (e.g. "Andrew Choi")
        if len(words) == 2 and not surname_parts:
            return f"{words[1]}, {words[0]}"

    return normalized


def employee_folder_name(first: str, last: str) -> str:
    """Build the new-style folder name: 'Last, First'."""
    return f"{(last or '').strip()}, {(first or '').strip()}"


def build_mapping(old_folders: list[str], roster: list[dict], root: str) -> list[dict]:
    """
    Match each old folder to a BambooHR employee and compute the new path.

    Returns a list of dicts:
      {old_path, new_path, employee_email, match_confidence}
    """
    # Build lookup from normalized name -> employee info
    emp_by_name: dict[str, dict] = {}
    for emp in roster:
        first = (emp.get("firstName") or "").strip()
        last = (emp.get("lastName") or "").strip()
        if first and last:
            key = f"{last}, {first}".lower()
            emp_by_name[key] = emp

    # Manual overrides for folders that can't be auto-matched
    _MANUAL_OVERRIDES: dict[str, dict] = {
        # folder segment -> {new_path suffix, email, confidence, note}
        "PUGA DE BIASE Leonardo": {
            "new_path": f"{root.strip('/')}/Inactive Migration/Puga de Biase, Leonardo Jesus",
            "email": "", "confidence": "manual", "note": "No longer at Tech Soft 3D",
        },
        "OConnell, Maria_DOH_27.03.2021": {
            "new_path": f"{root.strip('/')}/Inactive Migration/O'Connell, Maria Anne",
            "email": "", "confidence": "manual", "note": "Name corrected: O'Connell",
        },
        "CHRISTOPHE Alexandra_DOH 1 SEPT 2025": {
            "new_path": f"{root.strip('/')}/France/Christophe-Argenvillier, Alexandra",
            "email": "", "confidence": "manual", "note": "Full name: Alexandra Christophe-Argenvillier",
        },
        "TURBE Yannick": {
            "new_path": f"{root.strip('/')}/Inactive Migration/Turbé, Yannick",
            "email": "", "confidence": "manual", "note": "No longer active",
        },
        "Krueger, Florian_DOH_09.01.2025": {
            "new_path": f"{root.strip('/')}/Inactive Migration/Krüger, Florian",
            "email": "", "confidence": "manual", "note": "Name corrected: Krüger",
        },
    }

    # Patterns indicating inactive/terminated employees
    _INACTIVE_PATTERNS = (
        "_CONTRACTORS_and_TEMPS/",
        "/Contractor/",
        "/Contractors/",
        "/Archive/",
    )
    _INACTIVE_SUFFIXES = re.compile(r"[_\s-]Term[_\s.]|_TERM[_\s]", re.I)

    mapping = []
    for old_rel in old_folders:
        # The employee folder is the last segment that looks like a name
        parts = old_rel.split("/")
        emp_segment = parts[-1]

        # Parse into "last, first" format for matching
        parsed = _parse_folder_to_last_first(emp_segment)
        match = emp_by_name.get(parsed) if parsed else None

        # Also try the simple normalized form
        if not match:
            normalized = normalize_name(emp_segment)
            match = emp_by_name.get(normalized)

        if match:
            pay_schedule = (match.get("paySchedule") or "Unknown").strip()
            new_name = employee_folder_name(match.get("firstName"), match.get("lastName"))
            new_path = f"{root.strip('/')}/{pay_schedule}/{new_name}"
            mapping.append({
                "old_path": f"{root.strip('/')}/{old_rel}",
                "new_path": new_path,
                "employee_email": match.get("workEmail", ""),
                "match_confidence": "exact",
            })
        else:
            # Try partial matching (last name only)
            partial_match = None
            search_name = parsed or normalize_name(emp_segment)
            for key, emp in emp_by_name.items():
                emp_last = (emp.get("lastName") or "").lower()
                if emp_last and len(emp_last) >= 3 and emp_last in search_name:
                    partial_match = emp
                    break

            if partial_match:
                pay_schedule = (partial_match.get("paySchedule") or "Unknown").strip()
                new_name = employee_folder_name(
                    partial_match.get("firstName"), partial_match.get("lastName")
                )
                new_path = f"{root.strip('/')}/{pay_schedule}/{new_name}"
                mapping.append({
                    "old_path": f"{root.strip('/')}/{old_rel}",
                    "new_path": new_path,
                    "employee_email": partial_match.get("workEmail", ""),
                    "match_confidence": "partial",
                })
            else:
                # Check manual overrides
                override = _MANUAL_OVERRIDES.get(emp_segment)
                if override:
                    mapping.append({
                        "old_path": f"{root.strip('/')}/{old_rel}",
                        "new_path": override["new_path"],
                        "employee_email": override["email"],
                        "match_confidence": override["confidence"],
                    })
                # Check if this is an inactive/terminated employee
                elif any(p in old_rel for p in _INACTIVE_PATTERNS) or \
                        _INACTIVE_SUFFIXES.search(emp_segment):
                    # Recommend Inactive Migration folder
                    # Try to extract a clean name for the destination
                    clean_parsed = _parse_folder_to_last_first(emp_segment)
                    if clean_parsed:
                        parts_name = clean_parsed.split(", ")
                        dest_name = f"{parts_name[0].title()}, {parts_name[1].title()}" \
                            if len(parts_name) == 2 else clean_parsed.title()
                    else:
                        dest_name = normalize_name(emp_segment).title()
                    new_path = f"{root.strip('/')}/Inactive Migration/{dest_name}"
                    mapping.append({
                        "old_path": f"{root.strip('/')}/{old_rel}",
                        "new_path": new_path,
                        "employee_email": None,
                        "match_confidence": "inactive",
                    })
                else:
                    mapping.append({
                        "old_path": f"{root.strip('/')}/{old_rel}",
                        "new_path": None,
                        "employee_email": None,
                        "match_confidence": "unmatched",
                    })

    return mapping


# --- Subfolder destination recommendations ------------------------------------

# Pattern-based rules for recommending where old subfolder contents should go.
# Each tuple: (pattern, recommended_destination, notes)
_SUBFOLDER_RULES: list[tuple[re.Pattern, str, str]] = [
    # --- English names ---
    (re.compile(r"^(1-?\s*ID|1-?\s*Identity)$", re.I),
     "Confidential", "ID and personal docs (passport, ID)"),
    (re.compile(r"^2-?\s*Employment\s*[Cc]ontract$", re.I),
     "Personnel/Employment Contracts", "Contract docs"),
    (re.compile(r"^3-?\s*Onboarding$", re.I),
     "Personnel/Onboarding", "Onboarding materials"),
    (re.compile(r"^(3-?\s*Benefits|4-?\s*Benefits|4-?\s*Health\s*affiliation).*$", re.I),
     "Confidential", "Benefits/health info is sensitive"),
    (re.compile(r"^(4-?\s*Sick\s*Leave|5-?\s*Sick\s*[Ll]eave)$", re.I),
     "Confidential", "Medical/leave info is sensitive"),
    (re.compile(r"^(5-?\s*Correspondence|6-?\s*Correspond[ae]nc[es]*)$", re.I),
     "Personnel/Correspondence", "General correspondence"),
    (re.compile(r"^(7-?\s*Degrees|7-?\s*Certificate[\s-]*degree|7-?\s*Diplome)$", re.I),
     "Personnel/Onboarding", "Degrees & certificates go into Onboarding"),
    (re.compile(r"^Confidential$", re.I),
     "Confidential", "Already in correct structure"),
    (re.compile(r"^Personnel$", re.I),
     "Personnel", "Already in correct structure"),
    # --- French names ---
    (re.compile(r"^1-?\s*Identit", re.I),
     "Confidential", "French: Identity docs (passport, ID)"),
    (re.compile(r"^2-?\s*Contrat\s*de\s*travail", re.I),
     "Personnel/Employment Contracts", "French: Employment contract"),
    (re.compile(r"^4-?\s*Affiliation", re.I),
     "Confidential", "French: Health/benefits affiliation"),
    (re.compile(r"^5-?\s*Arr[eê]ts?$", re.I),
     "Confidential", "French: Sick leave (arrêts)"),
    (re.compile(r"^7-?\s*Dipl[oô]mes?$", re.I),
     "Personnel/Onboarding", "French: Degrees/diplomas go into Onboarding"),
    # --- Legacy / Actify ---
    (re.compile(r"^Actify\s*files-", re.I),
     "Personnel/Legacy Actify Files", "Historical files from Actify acquisition"),
    # --- Compensation (sensitive) ---
    (re.compile(r"^(Paycheck|Payslip|Paychecks)", re.I),
     "Confidential", "Compensation records are sensitive"),
    (re.compile(r"^Salary\s", re.I),
     "Confidential", "Compensation arrangement is sensitive"),
    (re.compile(r"^Payroll", re.I),
     "Confidential", "Compensation/payroll records are sensitive"),
    (re.compile(r"^(Child\s*Voucher|Absences)", re.I),
     "Confidential", "Benefits/leave records are sensitive"),
    # --- Misc categories ---
    (re.compile(r"^(OFFBOARDING|Offboarding)$", re.I),
     "Personnel/Offboarding", "Offboarding/termination docs"),
    (re.compile(r"^(INTERNSHIP|Apprenticeship|Stage\s*\+)", re.I),
     "Personnel/Onboarding", "Intern/apprentice placement docs"),
    (re.compile(r"^(Transfer|Transfer GP)$", re.I),
     "Personnel/Employment Contracts", "Internal transfer documentation"),
    (re.compile(r"^GP\s*employment$", re.I),
     "Personnel/Employment Contracts", "Global Professional Employer contract"),
    (re.compile(r"^(Development\s*plan|fitness)$", re.I),
     "Personnel/Correspondence", "Performance/development docs"),
    (re.compile(r"^(PTO|8\s*-?\s*PTO|PTP)$", re.I),
     "Confidential", "Paid time off records"),
    (re.compile(r"^dossier\s*sans\s*titre$", re.I),
     "Personnel (review)", "French: Untitled folder - review manually"),
    (re.compile(r"^Interview\s", re.I),
     "Personnel/Correspondence", "Interview documentation"),
    # --- Exclusions ---
    (re.compile(r"[\s_-]DOH[\s_]|[\s_-]DOG[\s_]|[\s_-]DOH$|[\s_-]DOG$", re.I),
     "EXCLUDE", "Misidentified - this is an employee folder"),
    (re.compile(r"^[A-Z]{2,}\s+[A-Z][a-zà-ÿ]", re.I & 0),
     "EXCLUDE", "Misidentified - this is an employee folder (LASTNAME Firstname)"),
    (re.compile(r"^AMENDED_", re.I),
     "Personnel/Employment Contracts", "Signed contract amendment"),
    (re.compile(r"^Stage\s+(3eme|Lyc)", re.I),
     "EXCLUDE", "Intern program container folder"),
    (re.compile(r"contract$", re.I),
     "Personnel/Employment Contracts", "Contract documentation"),
    # --- Archive / historical ---
    (re.compile(r"^\d{4}\s*-\s*", re.I),
     "Personnel/Correspondence", "Historical/archive folder"),
    (re.compile(r"^Before\s", re.I),
     "Personnel/Employment Contracts", "Historical transfer documentation"),
]


def _recommend_destination(name: str) -> tuple[str, str]:
    """Return (recommended_destination, notes) for an old subfolder name."""
    for pattern, dest, notes in _SUBFOLDER_RULES:
        if pattern.search(name):
            return dest, notes
    return "Personnel (review)", "No automatic rule matched - review manually"


# --- Main ---------------------------------------------------------------------

def main(argv: list[str]) -> int:
    json_output = "--json" in argv
    csv_output = "--csv" in argv
    use_test_group = "--all-users" not in argv or "--use-test-group" in argv

    test_emails: set[str] = set()
    if use_test_group:
        test_emails = load_test_users()
        if not test_emails:
            return 2
        print(f"Test group mode: scoping to {len(test_emails)} user(s)", file=sys.stderr)

    required = [
        "BAMBOO_HR_KEY",
        "AZURE_APP_CLIENT_ID", "AZURE_TENANT_ID", "AZURE_CLIENT_SECRET",
        "SHAREPOINT_HOSTNAME", "SHAREPOINT_SITE_PATH",
        "SHAREPOINT_DRIVE_NAME", "SHAREPOINT_FOLDER_PATH",
    ]
    missing = [k for k in required if not os.environ.get(k)]
    if missing:
        print(f"ERROR: missing env vars: {', '.join(missing)}", file=sys.stderr)
        return 2

    # --- Connect to SharePoint ---
    print("Acquiring Graph token ...", file=sys.stderr)
    graph = graph_session(graph_token())

    hostname = os.environ["SHAREPOINT_HOSTNAME"]
    site_path = os.environ["SHAREPOINT_SITE_PATH"]
    drive_name = os.environ["SHAREPOINT_DRIVE_NAME"]
    root = os.environ["SHAREPOINT_FOLDER_PATH"]

    print(f"Resolving site {hostname}{site_path} ...", file=sys.stderr)
    site_id = resolve_site_id(graph, hostname, site_path)

    print(f"Resolving drive '{drive_name}' ...", file=sys.stderr)
    drive_id = resolve_drive_id(graph, site_id, drive_name)

    # --- Discover old folders ---
    print(f"Listing folders under: {root}", file=sys.stderr)
    old_folders, subfolder_map = find_employee_folders(graph, drive_id, root)
    print(f"  Found {len(old_folders)} employee folder(s)", file=sys.stderr)

    # --- Fetch BambooHR roster ---
    print("Fetching BambooHR roster ...", file=sys.stderr)
    bamboo = bamboo_session(os.environ["BAMBOO_HR_KEY"])
    roster = fetch_roster(bamboo)
    print(f"  {len(roster)} employee(s) in roster", file=sys.stderr)

    # --- Build mapping ---
    mapping = build_mapping(old_folders, roster, root)

    # Filter to test group if requested
    if use_test_group:
        mapping = [m for m in mapping if (m.get("employee_email") or "").lower() in test_emails]
        print(f"  Filtered to {len(mapping)} folder(s) matching test group", file=sys.stderr)

    # --- Collect unique subfolder names across all employee folders ---
    all_subfolders: set[str] = set()
    for subs in subfolder_map.values():
        all_subfolders.update(subs)

    # --- Output ---
    if csv_output:
        # --- Write mapping CSV ---
        mapping_file = "folder_mapping_test.csv" if use_test_group else "folder_mapping.csv"
        with open(mapping_file, "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow(["Old Path", "New Path", "Employee Email", "Match Confidence", "Subfolders"])
            for m in mapping:
                rel = m["old_path"][len(root.strip("/")) + 1:]  # relative from root
                subs = subfolder_map.get(rel, [])
                writer.writerow([
                    m["old_path"],
                    m["new_path"] or "",
                    m["employee_email"] or "",
                    m["match_confidence"],
                    "; ".join(subs),
                ])
        print(f"Wrote {mapping_file} ({len(mapping)} rows)", file=sys.stderr)

        # --- Write unique subfolders CSV ---
        subfolders_file = "employee_subfolders.csv"
        with open(subfolders_file, "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow(["Subfolder Name", "Occurrences", "Recommended Destination", "Notes"])
            # Count how many employee folders use each subfolder
            counts: dict[str, int] = {}
            for subs in subfolder_map.values():
                for s in subs:
                    counts[s] = counts.get(s, 0) + 1
            for name in sorted(counts, key=lambda k: (-counts[k], k)):
                dest, notes = _recommend_destination(name)
                writer.writerow([name, counts[name], dest, notes])
        print(f"Wrote {subfolders_file} ({len(counts)} unique subfolders)", file=sys.stderr)

    elif json_output:
        output = {
            "mapping": mapping,
            "subfolders": {
                "unique_names": sorted(all_subfolders),
                "per_employee": subfolder_map,
            },
        }
        print(json.dumps(output, indent=2))
    else:
        exact = [m for m in mapping if m["match_confidence"] == "exact"]
        partial = [m for m in mapping if m["match_confidence"] == "partial"]
        unmatched = [m for m in mapping if m["match_confidence"] == "unmatched"]

        print(f"\n{'='*80}")
        print(f"MAPPING RESULTS: {len(exact)} exact, {len(partial)} partial, "
              f"{len(unmatched)} unmatched")
        print(f"{'='*80}\n")

        if exact:
            print("--- EXACT MATCHES ---")
            for m in exact:
                print(f"  {_mask_filename(m['old_path'])}")
                print(f"    -> {_mask_filename(m['new_path'])}  ({m['employee_email']})")
                print()

        if partial:
            print("--- PARTIAL MATCHES (review these) ---")
            for m in partial:
                print(f"  {_mask_filename(m['old_path'])}")
                print(f"    -> {_mask_filename(m['new_path'])}  ({m['employee_email']})")
                print()

        if unmatched:
            print("--- UNMATCHED (no BambooHR employee found) ---")
            for m in unmatched:
                print(f"  {_mask_filename(m['old_path'])}")
                print()

        # --- Subfolder listing ---
        print(f"{'='*80}")
        print(f"SUBFOLDERS FOUND INSIDE EMPLOYEE FOLDERS ({len(all_subfolders)} unique)")
        print(f"{'='*80}\n")
        for name in sorted(all_subfolders):
            print(f"  {_mask_filename(name)}")
        print()

    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
