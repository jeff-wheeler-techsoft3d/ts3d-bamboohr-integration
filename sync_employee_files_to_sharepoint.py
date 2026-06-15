"""
End-to-end sync: BambooHR employee files -> SharePoint backup.

Layout produced in SharePoint:

  {SHAREPOINT_FOLDER_PATH}/{PaySchedule}/{Last, First}/Personnel/{BambooHR Category}/{filename}
  {SHAREPOINT_FOLDER_PATH}/{PaySchedule}/{Last, First}/Confidential/{filename}

Routing rule:
  BambooHR category name == "Confidential"  -> Confidential/  (flattened)
  everything else                            -> Personnel/<category>/

Employee folder name comes from BambooHR's firstName + lastName on report 135:
  "Last, First"

Idempotency:
  Before uploading, the script checks whether the SharePoint item already
  exists at the destination path; if it does and the size matches BambooHR's
  reported size, the upload is skipped.

Required env vars (see also .env):
  BAMBOO_HR_KEY
  AZURE_APP_CLIENT_ID
  AZURE_TENANT_ID
  AZURE_CLIENT_SECRET
  SHAREPOINT_HOSTNAME       e.g. hoops3d.sharepoint.com
  SHAREPOINT_SITE_PATH      e.g. /people-experience
  SHAREPOINT_DRIVE_NAME     e.g. "HR Only Documents"
  SHAREPOINT_FOLDER_PATH    e.g. "Personnel Files/Active"

Usage:
    python sync_employee_files_to_sharepoint.py                    # sync test group only
    python sync_employee_files_to_sharepoint.py EMAIL ...         # sync listed emails, limited to test group
    python sync_employee_files_to_sharepoint.py --dry-run         # plan only, no uploads
    python sync_employee_files_to_sharepoint.py --all-users       # disable test-group-only safeguard
    python sync_employee_files_to_sharepoint.py --use-test-group  # explicit test-group mode
"""

from __future__ import annotations

import json
import os
import re
import sys
from pathlib import Path
from typing import Iterable
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
SMALL_FILE_LIMIT = 4 * 1024 * 1024
CHUNK_SIZE = 5 * 1024 * 1024  # multiple of 320 KiB

CONFIDENTIAL_CATEGORY = "Confidential"

_PATH_SAFE_RE = re.compile(r'[<>:"/\\|?*\x00-\x1f]')


# --- Helpers ------------------------------------------------------------------

def _safe_segment(name: str) -> str:
    """Sanitize a single path segment for SharePoint."""
    s = (name or "").strip()
    s = _PATH_SAFE_RE.sub("_", s)
    # SharePoint also dislikes leading/trailing dots and spaces in segments.
    s = s.strip(" .")
    return s or "unnamed"


def employee_folder_name(first: str, last: str, email: str | None = None) -> str | None:
    """
    Build "Last, First" from BambooHR firstName/lastName. Falls back to
    parsing the email local-part when names are missing.
    """
    first = (first or "").strip()
    last = (last or "").strip()
    if first and last:
        return f"{last}, {first}"
    if email and "@" in email:
        local = email.split("@", 1)[0]
        parts = [p for p in local.split(".") if p]
        if len(parts) >= 2:
            f = parts[0].capitalize()
            l = " ".join(p.capitalize() for p in parts[1:])
            return f"{l}, {f}"
    return None


# --- BambooHR -----------------------------------------------------------------

def bamboo_session(api_key: str) -> requests.Session:
    s = requests.Session()
    s.auth = (api_key, "x")
    s.headers.update({"Accept": "application/json"})
    return s


def fetch_roster(session: requests.Session) -> list[dict]:
    """Pull active employees + country from report 135 (same one bamboo_azure.py uses)."""
    r = session.get(
        f"{BAMBOO_BASE}/reports/135",
        params={"format": "json", "fd": "yes", "onlyCurrent": "true"},
    )
    r.raise_for_status()
    return r.json().get("employees", []) or []


def list_employee_files(session: requests.Session, employee_id: str) -> dict:
    r = session.get(f"{BAMBOO_BASE}/employees/{employee_id}/files/view/")
    r.raise_for_status()
    return r.json()


def download_employee_file(session: requests.Session, employee_id: str,
                           file_id: str | int) -> bytes:
    """Return raw bytes for one employee file."""
    url = f"{BAMBOO_BASE}/employees/{employee_id}/files/{file_id}"
    r = session.get(url)
    r.raise_for_status()
    return r.content


# --- Microsoft Graph ----------------------------------------------------------

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


def get_item_by_path(session: requests.Session, drive_id: str, path: str) -> dict | None:
    """GET /drives/{drive-id}/root:/{path} -> item metadata or None on 404."""
    url = f"{GRAPH}/drives/{drive_id}/root:/{quote(path.strip('/'))}"
    r = session.get(url)
    if r.status_code == 404:
        return None
    r.raise_for_status()
    return r.json()


def upload_small(session: requests.Session, drive_id: str, dest_path: str,
                 data: bytes) -> dict:
    url = f"{GRAPH}/drives/{drive_id}/root:/{quote(dest_path.strip('/'))}:/content"
    r = session.put(url, data=data,
                    headers={"Content-Type": "application/octet-stream"})
    r.raise_for_status()
    return r.json()


def upload_large(session: requests.Session, drive_id: str, dest_path: str,
                 data: bytes) -> dict:
    name = dest_path.rsplit("/", 1)[-1]
    create_url = f"{GRAPH}/drives/{drive_id}/root:/{quote(dest_path.strip('/'))}:/createUploadSession"
    r = session.post(create_url, json={
        "item": {"@microsoft.graph.conflictBehavior": "replace", "name": name},
    })
    r.raise_for_status()
    upload_url = r.json()["uploadUrl"]

    size = len(data)
    sent = 0
    bare = requests.Session()
    last = None
    while sent < size:
        chunk = data[sent:sent + CHUNK_SIZE]
        end = sent + len(chunk) - 1
        resp = bare.put(upload_url, data=chunk, headers={
            "Content-Length": str(len(chunk)),
            "Content-Range": f"bytes {sent}-{end}/{size}",
        })
        if resp.status_code not in (200, 201, 202):
            raise RuntimeError(f"Chunk upload failed at {sent}-{end}: "
                               f"{resp.status_code} {resp.text}")
        sent += len(chunk)
        last = resp
    return last.json() if last is not None else {}


def upload_bytes(session: requests.Session, drive_id: str, dest_path: str,
                 data: bytes) -> dict:
    if len(data) <= SMALL_FILE_LIMIT:
        return upload_small(session, drive_id, dest_path, data)
    return upload_large(session, drive_id, dest_path, data)


# --- Path planning ------------------------------------------------------------

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


def plan_destination(root: str, pay_schedule: str, employee_folder: str,
                     category: str, filename: str) -> str:
    pay_schedule = _safe_segment(pay_schedule)
    employee_folder = _safe_segment(employee_folder)
    category = _safe_segment(category)
    filename = _safe_segment(filename)

    if category == CONFIDENTIAL_CATEGORY:
        # Flatten under Confidential/
        return f"{root.strip('/')}/{pay_schedule}/{employee_folder}/Confidential/{filename}"
    return f"{root.strip('/')}/{pay_schedule}/{employee_folder}/Personnel/{category}/{filename}"


# --- Sync ---------------------------------------------------------------------

def sync_employee(bamboo: requests.Session, graph: requests.Session,
                  drive_id: str, root: str, emp: dict,
                  dry_run: bool) -> tuple[int, int, int]:
    """Returns (uploaded, skipped, failed) counts for this employee."""
    email = emp.get("workEmail") or ""
    pay_schedule = emp.get("paySchedule") or ""
    emp_id = str(emp.get("id") or "")
    folder_name = employee_folder_name(emp.get("firstName"), emp.get("lastName"), email)

    if not (emp_id and email and pay_schedule and folder_name):
        print(f"  SKIP employee (missing fields): id={emp_id} email={email} "
              f"paySchedule={pay_schedule} folder={folder_name}")
        return (0, 0, 0)

    print(f"\n== {email}  (id={emp_id}, paySchedule={pay_schedule}, folder='{folder_name}') ==")

    try:
        listing = list_employee_files(bamboo, emp_id)
    except requests.HTTPError as e:
        print(f"  ERROR listing files: {e}")
        return (0, 0, 1)

    uploaded = skipped = failed = 0
    for category in listing.get("categories", []) or []:
        cat_name = category.get("name") or "Uncategorized"
        for f in category.get("files", []) or []:
            file_id = f.get("id")
            original = f.get("originalFileName") or f.get("name") or f"file_{file_id}"
            expected_size = f.get("size")
            dest = plan_destination(root, pay_schedule, folder_name, cat_name, original)

            # Skip if already present with matching size.
            try:
                existing = get_item_by_path(graph, drive_id, dest)
            except requests.HTTPError as e:
                print(f"  ERROR checking {_mask_filename(dest)}: {e}")
                failed += 1
                continue

            if existing is not None:
                actual = existing.get("size")
                size_ok = (not isinstance(expected_size, int)) or expected_size == 0 \
                    or actual == expected_size
                if size_ok:
                    print(f"  SKIP  {_mask_filename(dest)}  (exists, size={actual})")
                    skipped += 1
                    continue
                print(f"  RE-UP {_mask_filename(dest)}  (size mismatch: remote={actual}, expected={expected_size})")

            if dry_run:
                print(f"  PLAN  {_mask_filename(dest)}  (would upload {expected_size} bytes)")
                continue

            try:
                blob = download_employee_file(bamboo, emp_id, file_id)
            except requests.HTTPError as e:
                print(f"  FAIL  download fileId={file_id}: {e}")
                failed += 1
                continue

            try:
                upload_bytes(graph, drive_id, dest, blob)
                print(f"  OK    {_mask_filename(dest)}  ({len(blob)} bytes)")
                uploaded += 1
            except requests.HTTPError as e:
                print(f"  FAIL  upload {_mask_filename(dest)}: {e.response.status_code} {e.response.text[:200]}")
                failed += 1
            except Exception as e:
                print(f"  FAIL  upload {dest}: {e}")
                failed += 1

    return (uploaded, skipped, failed)


# --- Main ---------------------------------------------------------------------

def main(argv: list[str]) -> int:
    dry_run = "--dry-run" in argv
    use_test_group = "--all-users" not in argv or "--use-test-group" in argv
    only_emails = {a.lower() for a in argv if "@" in a}

    if use_test_group:
        test_emails = load_test_users()
        if not test_emails:
            return 2
        only_emails = only_emails & test_emails if only_emails else test_emails

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

    bamboo = bamboo_session(os.environ["BAMBOO_HR_KEY"])

    print("Acquiring Graph token ...")
    graph = graph_session(graph_token())

    print(f"Resolving SharePoint site {os.environ['SHAREPOINT_HOSTNAME']}"
          f"{os.environ['SHAREPOINT_SITE_PATH']} ...")
    site_id = resolve_site_id(graph, os.environ["SHAREPOINT_HOSTNAME"],
                              os.environ["SHAREPOINT_SITE_PATH"])
    print(f"  site id: {site_id}")

    print(f"Resolving drive '{os.environ['SHAREPOINT_DRIVE_NAME']}' ...")
    drive_id = resolve_drive_id(graph, site_id, os.environ["SHAREPOINT_DRIVE_NAME"])
    print(f"  drive id: {drive_id}")

    root = os.environ["SHAREPOINT_FOLDER_PATH"]
    print(f"Root: {root}")
    if use_test_group:
        print(f"TEST GROUP ONLY — scoped to {len(only_emails)} user(s) from test_users.json")
    if dry_run:
        print("DRY RUN — no uploads will be performed.")

    print("\nFetching BambooHR roster ...")
    roster = fetch_roster(bamboo)
    if only_emails:
        roster = [e for e in roster
                  if (e.get("workEmail") or "").lower() in only_emails]
    print(f"  {len(roster)} employee(s) to process")

    tot_up = tot_skip = tot_fail = 0
    for emp in roster:
        u, s, f = sync_employee(bamboo, graph, drive_id, root, emp, dry_run)
        tot_up += u
        tot_skip += s
        tot_fail += f

    print("\n=== Summary ===")
    print(f"  uploaded: {tot_up}")
    print(f"  skipped:  {tot_skip}")
    print(f"  failed:   {tot_fail}")
    return 0 if tot_fail == 0 else 1


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
