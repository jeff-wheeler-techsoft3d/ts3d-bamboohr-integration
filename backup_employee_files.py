"""
Prototype: List files belonging to a specific BambooHR employee so they can be
backed up to SharePoint in a future integration.

Flow:
  1. Look up the employee's BambooHR `id` by work email.
  2. Call the BambooHR "list employee files" endpoint to enumerate files
     (and categories) attached to that employee.
  3. Print a flat listing. (Downloading + SharePoint upload to be added later.)

BambooHR docs:
  - List employee files: GET /employees/{id}/files/view/
  - Download a file:     GET /employees/{id}/files/{fileId}

Auth: BambooHR uses HTTP Basic auth with the API key as the username and any
non-empty password (conventionally "x"). We use `requests` `auth=` to avoid
embedding the key in the URL.

Required env vars:
  - BAMBOO_HR_KEY

Usage:
  python backup_employee_files.py [employee_email]
  (defaults to jeff.wheeler@techsoft3d.com)
"""

import os
import re
import sys
import json
from pathlib import Path
import requests

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    # python-dotenv is optional; env vars can also be exported in the shell.
    pass

BAMBOO_COMPANY = "techsoft3d"
BAMBOO_BASE = f"https://api.bamboohr.com/api/gateway.php/{BAMBOO_COMPANY}/v1"
DEFAULT_EMAIL = "jeff.wheeler@techsoft3d.com"


def bamboo_session(api_key: str) -> requests.Session:
    s = requests.Session()
    s.auth = (api_key, "x")
    s.headers.update({"Accept": "application/json"})
    return s


def find_employee_id_by_email(session: requests.Session, email: str) -> str | None:
    """
    Use the employee directory to map an email to an employee id.
    Falls back to scanning the same report (135) the existing scripts use if
    the directory endpoint isn't enabled for the API key.
    """
    # Preferred: directory endpoint
    r = session.get(f"{BAMBOO_BASE}/employees/directory")
    if r.status_code == 200:
        data = r.json()
        for emp in data.get("employees", []):
            if (emp.get("workEmail") or "").lower() == email.lower():
                return str(emp["id"])

    # Fallback: report 135 (used elsewhere in this repo) — but it doesn't include id by default,
    # so request the report with id field included.
    r = session.get(
        f"{BAMBOO_BASE}/reports/135",
        params={"format": "json", "fd": "yes", "onlyCurrent": "true"},
    )
    if r.status_code == 200:
        data = r.json()
        for emp in data.get("employees", []):
            if (emp.get("workEmail") or "").lower() == email.lower():
                # Report rows include an "id" field.
                return str(emp.get("id"))
    return None


_SAFE_NAME_RE = re.compile(r'[^A-Za-z0-9._\- ]+')


def _safe(name: str) -> str:
    """Sanitize a path segment for the local filesystem."""
    name = (name or "").strip().replace("/", "_").replace("\\", "_")
    name = _SAFE_NAME_RE.sub("_", name)
    return name or "unnamed"


def download_employee_file(
    session: requests.Session,
    employee_id: str,
    file_id: str | int,
    dest_path: Path,
) -> Path:
    """
    GET /employees/{id}/files/{fileId} -> streams the binary to `dest_path`.
    If the response sets a Content-Disposition filename, we honor it instead
    of the caller-suggested filename (keeps the real extension).
    Returns the final path written.
    """
    url = f"{BAMBOO_BASE}/employees/{employee_id}/files/{file_id}"
    with session.get(url, stream=True) as r:
        r.raise_for_status()
        # Prefer server-provided filename when present.
        cd = r.headers.get("Content-Disposition", "")
        m = re.search(r'filename\*?=(?:UTF-8\'\')?"?([^";]+)"?', cd)
        if m:
            server_name = _safe(requests.utils.unquote(m.group(1)))
            if server_name:
                dest_path = dest_path.with_name(server_name)
        dest_path.parent.mkdir(parents=True, exist_ok=True)
        with open(dest_path, "wb") as fh:
            for chunk in r.iter_content(chunk_size=64 * 1024):
                if chunk:
                    fh.write(chunk)
    return dest_path


def list_employee_files(session: requests.Session, employee_id: str) -> dict:
    """
    GET /employees/{id}/files/view/
    Returns JSON with `categories`, each containing `files` metadata
    (id, name, originalFileName, size, dateCreated, createdBy, shareWithEmployee, etc.)
    """
    url = f"{BAMBOO_BASE}/employees/{employee_id}/files/view/"
    r = session.get(url)
    r.raise_for_status()
    return r.json()


def flatten_files(files_payload: dict) -> list[dict]:
    flat = []
    for category in files_payload.get("categories", []):
        cat_id = category.get("id")
        cat_name = category.get("name")
        for f in category.get("files", []) or []:
            flat.append(
                {
                    "categoryId": cat_id,
                    "category": cat_name,
                    "fileId": f.get("id"),
                    "name": f.get("name"),
                    "originalFileName": f.get("originalFileName"),
                    "size": f.get("size"),
                    "dateCreated": f.get("dateCreated"),
                    "createdBy": f.get("createdBy"),
                    "shareWithEmployee": f.get("shareWithEmployee"),
                }
            )
    return flat


def main() -> int:
    email = sys.argv[1] if len(sys.argv) > 1 else DEFAULT_EMAIL

    api_key = os.environ.get("BAMBOO_HR_KEY")
    if not api_key:
        print("ERROR: env var 'BAMBOO_HR_KEY' is not set.", file=sys.stderr)
        return 2

    session = bamboo_session(api_key)

    print(f"Looking up BambooHR employee id for {email} ...")
    emp_id = find_employee_id_by_email(session, email)
    if not emp_id:
        print(f"ERROR: could not find employee id for {email}", file=sys.stderr)
        return 1
    print(f"  -> employee id: {emp_id}")

    print(f"Listing files for employee {emp_id} ...")
    payload = list_employee_files(session, emp_id)

    files = flatten_files(payload)
    print(f"Found {len(files)} file(s) across {len(payload.get('categories', []))} category(ies).\n")

    for f in files:
        size = f["size"]
        size_str = f"{size} bytes" if isinstance(size, int) else str(size)
        print(
            f"[{f['category']} ({f['categoryId']})] "
            f"fileId={f['fileId']}  "
            f"name={f['originalFileName'] or f['name']}  "
            f"size={size_str}  "
            f"created={f['dateCreated']}"
        )

    # Download each file to ~/Downloads/bamboo_backup/<email>/<category>/<filename>
    download_root = Path.home() / "Downloads" / "bamboo_backup" / _safe(email)
    print(f"\nDownloading files to {download_root} ...")
    for f in files:
        suggested = f["originalFileName"] or f["name"] or f"file_{f['fileId']}"
        target = download_root / _safe(f["category"] or "Uncategorized") / _safe(suggested)

        # Skip if already downloaded. If BambooHR provided a size, also require
        # the local file size to match before treating it as a valid cache hit.
        if target.exists():
            expected = f.get("size")
            actual = target.stat().st_size
            size_ok = (not isinstance(expected, int)) or expected == 0 or actual == expected
            if size_ok:
                f["localPath"] = str(target)
                f["skipped"] = True
                print(f"  SKIP {target.relative_to(Path.home())} (already exists)")
                continue
            else:
                print(
                    f"  RE-DL {target.relative_to(Path.home())} "
                    f"(size mismatch: local={actual}, expected={expected})"
                )

        try:
            written = download_employee_file(session, emp_id, f["fileId"], target)
            f["localPath"] = str(written)
            print(f"  OK  {written.relative_to(Path.home())}")
        except requests.HTTPError as e:
            f["localPath"] = None
            f["downloadError"] = f"{e.response.status_code} {e.response.reason}"
            print(f"  FAIL fileId={f['fileId']}: {f['downloadError']}", file=sys.stderr)
        except Exception as e:
            f["localPath"] = None
            f["downloadError"] = str(e)
            print(f"  FAIL fileId={f['fileId']}: {e}", file=sys.stderr)

    # Also dump the raw JSON for reference / future SharePoint mapping work.
    out_path = download_root / "employee_files_listing.json"
    out_path.parent.mkdir(parents=True, exist_ok=True)
    with open(out_path, "w") as fh:
        json.dump({"employeeId": emp_id, "email": email, "raw": payload, "flat": files}, fh, indent=2)
    print(f"\nWrote raw + flattened listing to {out_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
