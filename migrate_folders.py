"""
Migrate employee folders from old SharePoint structure to new structure.

Reads folder_mapping.csv (produced by map_old_to_new_folders.py) and:
  1. Copies files from each old folder into the new destination folder.
  2. Moves the old folder to a "Migrated/" holding area for comparison.
  3. Moves inactive folders to "Not-Migrated - Inactive/".

Modes:
  --dry-run   Show what would happen without making any changes (default).
  --execute   Actually perform the migration.
  --full      Include the archive step (move old folders to Migrated/). By default, only copies files.

The script creates destination folders as needed and routes top-level legacy
employee subfolders into the new Personnel/Confidential structure using
employee_subfolders.csv. Nested structure beneath those mapped folders is
preserved.

When a destination file already exists with a different size, this script
skips that file by default to preserve destination metadata and avoid
destructive overwrite behavior during parallel migration/sync windows.

Required env vars (same as other scripts):
  AZURE_APP_CLIENT_ID
  AZURE_TENANT_ID
  AZURE_CLIENT_SECRET
  SHAREPOINT_HOSTNAME
  SHAREPOINT_SITE_PATH
  SHAREPOINT_DRIVE_NAME
  SHAREPOINT_FOLDER_PATH

Usage:
    python migrate_folders.py --dry-run                     # dry-run for test group (partial only)
    python migrate_folders.py --execute                     # execute for test group (partial only)
    python migrate_folders.py --execute --full             # execute with archive for test group
    python migrate_folders.py --all-users --dry-run        # dry-run full scope (partial only)
    python migrate_folders.py --all-users --execute        # execute full scope (partial only)
    python migrate_folders.py --all-users --execute --full # execute full scope with archive
    python migrate_folders.py --use-test-group             # explicit test-group mode
"""

from __future__ import annotations

import csv
import json
import os
import sys
import time
from urllib.parse import quote

import requests
from msal import ConfidentialClientApplication

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

# --- Constants ----------------------------------------------------------------

GRAPH = "https://graph.microsoft.com/v1.0"
CSV_FILE = "folder_mapping.csv"
SUBFOLDER_CSV_FILE = "employee_subfolders.csv"

# Folders within the root where migrated/inactive items go
MIGRATED_FOLDER = "Migrated"
INACTIVE_FOLDER = "Not-Migrated - Inactive"

# Throttle: pause between Graph API calls to avoid 429s
THROTTLE_SECONDS = 0.1


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


# --- Output helpers ----------------------------------------------------------

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


# --- Auth / Graph helpers -----------------------------------------------------

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


# --- Graph file operations ----------------------------------------------------

def get_item_by_path(session: requests.Session, drive_id: str, path: str) -> dict | None:
    """GET item metadata by path. Returns None on 404."""
    url = f"{GRAPH}/drives/{drive_id}/root:/{quote(path.strip('/'))}"
    r = session.get(url)
    if r.status_code == 404:
        return None
    r.raise_for_status()
    return r.json()


def ensure_folder(session: requests.Session, drive_id: str, path: str) -> dict:
    """Create a folder (and parents) if it doesn't exist. Returns item metadata."""
    existing = get_item_by_path(session, drive_id, path)
    if existing:
        return existing

    # Create parent first (recursive)
    parts = path.strip("/").rsplit("/", 1)
    if len(parts) == 2:
        parent_path, folder_name = parts
        parent = ensure_folder(session, drive_id, parent_path)
        parent_id = parent["id"]
    else:
        folder_name = parts[0]
        # Create at root
        parent_id = "root"

    url = f"{GRAPH}/drives/{drive_id}/items/{parent_id}/children"
    body = {
        "name": folder_name,
        "folder": {},
        "@microsoft.graph.conflictBehavior": "fail",
    }
    r = session.post(url, json=body)
    if r.status_code == 409:
        # Already exists (race condition), fetch it
        time.sleep(THROTTLE_SECONDS)
        return get_item_by_path(session, drive_id, path)
    r.raise_for_status()
    time.sleep(THROTTLE_SECONDS)
    return r.json()


def list_children(session: requests.Session, drive_id: str, item_id: str) -> list[dict]:
    """List all children of a folder by item ID."""
    url = f"{GRAPH}/drives/{drive_id}/items/{item_id}/children"
    items = []
    while url:
        r = session.get(url)
        r.raise_for_status()
        data = r.json()
        items.extend(data.get("value", []))
        url = data.get("@odata.nextLink")
    return items


def list_children_by_path(session: requests.Session, drive_id: str, path: str) -> list[dict]:
    """List all children of a folder by path."""
    encoded = quote(path.strip("/"))
    url = f"{GRAPH}/drives/{drive_id}/root:/{encoded}:/children"
    items = []
    while url:
        r = session.get(url)
        if r.status_code == 404:
            return []
        r.raise_for_status()
        data = r.json()
        items.extend(data.get("value", []))
        url = data.get("@odata.nextLink")
    return items


def copy_item(session: requests.Session, drive_id: str, item_id: str,
              dest_parent_id: str, new_name: str | None = None) -> str | None:
    """
    Start an async copy. Returns the monitor URL (or None if completed immediately).
    Graph copy is async — returns 202 with a Location header for monitoring.
    Conflict behavior is set to "replace" for idempotent reruns.
    """
    url = f"{GRAPH}/drives/{drive_id}/items/{item_id}/copy"
    body: dict = {
        "parentReference": {"driveId": drive_id, "id": dest_parent_id},
        "@microsoft.graph.conflictBehavior": "replace",
    }
    if new_name:
        body["name"] = new_name
    r = session.post(url, json=body)
    if r.status_code in (200, 201):
        return None  # completed synchronously
    if r.status_code == 202:
        return r.headers.get("Location")
    r.raise_for_status()
    return None


def wait_for_copy(monitor_url: str, timeout: int = 300) -> tuple[bool, str | None]:
    """Poll copy monitor URL until complete/failed/timeout. Returns (success, error_code)."""
    bare = requests.Session()
    deadline = time.time() + timeout
    while time.time() < deadline:
        time.sleep(2)
        r = bare.get(monitor_url)
        if r.status_code == 200:
            data = r.json()
            status = data.get("status", "")
            if status == "completed":
                return True, None
            if status == "failed":
                return False, (data.get("error") or {}).get("code")
        elif r.status_code in (303, 302):
            # Redirect to the new item — copy is done
            return True, None
    print(f"    COPY TIMEOUT after {timeout}s")
    return False, "timeout"


def move_item(session: requests.Session, drive_id: str, item_id: str,
              dest_parent_id: str, new_name: str | None = None) -> dict:
    """Move an item to a new parent folder (optionally rename)."""
    url = f"{GRAPH}/drives/{drive_id}/items/{item_id}"
    body: dict = {
        "parentReference": {"id": dest_parent_id},
    }
    if new_name:
        body["name"] = new_name
    r = session.patch(url, json=body)
    r.raise_for_status()
    return r.json()


# --- Recursive copy -----------------------------------------------------------

def count_folder_contents(session: requests.Session, drive_id: str,
                          source_item: dict) -> tuple[int, int]:
    """Return recursive (file_count, folder_count) for a source folder."""
    file_count = 0
    folder_count = 0

    for child in list_children(session, drive_id, source_item["id"]):
        if "folder" in child:
            folder_count += 1
            nested_files, nested_folders = count_folder_contents(session, drive_id, child)
            file_count += nested_files
            folder_count += nested_folders
        else:
            file_count += 1

    return file_count, folder_count


def load_subfolder_destinations(csv_path: str) -> dict[str, str]:
    """Load legacy subfolder -> recommended destination mappings from CSV."""
    mapping: dict[str, str] = {}
    with open(csv_path, newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for row in reader:
            name = (row.get("Subfolder Name") or "").strip()
            dest = (row.get("Recommended Destination") or "").strip()
            if name and dest:
                mapping[name.lower()] = dest
    return mapping


def resolve_legacy_subfolder_destination(employee_dest: str, subfolder_name: str,
                                         subfolder_destinations: dict[str, str]) -> str:
    """Map one legacy top-level subfolder into the target employee structure."""
    recommended = subfolder_destinations.get(subfolder_name.lower(), "")
    employee_dest = employee_dest.rstrip("/")

    if recommended == "Confidential":
        return f"{employee_dest}/Confidential"
    if recommended == "Personnel":
        return f"{employee_dest}/Personnel"
    if recommended.startswith("Personnel/"):
        return f"{employee_dest}/{recommended}"

    # Unknown, review-only, or exclude recommendations still get copied into a
    # review area so data is preserved while remaining clearly segregated.
    return f"{employee_dest}/Personnel/Review/{subfolder_name}"


def plan_folder_copy(session: requests.Session, drive_id: str,
                     source_item: dict, dest_path: str,
                     subfolder_destinations: dict[str, str] | None = None,
                     current_level: int = 0) -> tuple[int, int, int]:
    """Return recursive (would_copy, would_skip, subfolder_count) for one folder."""
    would_copy = 0
    would_skip = 0
    subfolder_count = 0

    for child in list_children(session, drive_id, source_item["id"]):
        name = child.get("name", "")
        if "folder" in child:
            subfolder_count += 1
            if current_level == 0 and subfolder_destinations is not None:
                child_dest = resolve_legacy_subfolder_destination(
                    dest_path, name, subfolder_destinations
                )
            else:
                child_dest = f"{dest_path}/{name}"
            nested_copy, nested_skip, nested_folders = plan_folder_copy(
                session, drive_id, child, child_dest, subfolder_destinations, current_level + 1
            )
            would_copy += nested_copy
            would_skip += nested_skip
            subfolder_count += nested_folders
        else:
            file_dest_path = dest_path
            if current_level == 0:
                file_dest_path = f"{dest_path.rstrip('/')}/Personnel/Review"
            file_size = child.get("size", 0)
            dest_file_path = f"{file_dest_path.rstrip('/')}/{name}"
            existing = get_item_by_path(session, drive_id, dest_file_path)
            if existing is not None and existing.get("size") == file_size:
                would_skip += 1
            else:
                would_copy += 1

    return would_copy, would_skip, subfolder_count

def copy_folder_contents(session: requests.Session, drive_id: str,
                         source_item: dict, source_path: str, dest_path: str,
                         dry_run: bool, indent: int = 2,
                         subfolder_destinations: dict[str, str] | None = None,
                         current_level: int = 0,
                         dest_index: dict[str, tuple[int, str]] | None = None) -> tuple[int, int]:
    """
    Recursively copy all contents from source folder into dest_path.
    Returns (files_copied, errors).
    """
    prefix = " " * indent
    children = list_children(session, drive_id, source_item["id"])
    copied = 0
    errors = 0
    if dest_index is None:
        dest_index = {}

    # Ensure destination exists
    if not dry_run:
        dest_folder = ensure_folder(session, drive_id, dest_path)
        dest_id = dest_folder["id"]

    for child in children:
        name = child.get("name", "")
        child_source = f"{source_path.rstrip('/')}/{name}"
        if "folder" in child:
            # Recurse into subfolder
            if current_level == 0 and subfolder_destinations is not None:
                child_dest = resolve_legacy_subfolder_destination(
                    dest_path, name, subfolder_destinations
                )
            else:
                child_dest = f"{dest_path}/{name}"
            if dry_run:
                would_copy, would_skip, folder_count = plan_folder_copy(
                    session, drive_id, child, child_dest, subfolder_destinations, current_level + 1
                )
                print(f"{prefix}[DIR]  {_mask_filename(child_source)}/")
                print(
                    f"{prefix}       -> {_mask_filename(child_dest)}/  "
                    f"({would_copy} would copy, {would_skip} would skip, {folder_count} subfolder(s))"
                )
            else:
                print(f"{prefix}[DIR]  {_mask_filename(name)}/")
            c, e = copy_folder_contents(
                session, drive_id, child, child_source, child_dest, dry_run, indent + 2,
                subfolder_destinations, current_level + 1, dest_index
            )
            copied += c
            errors += e
        else:
            # File — copy it
            file_dest_path = dest_path
            if current_level == 0:
                file_dest_path = f"{dest_path.rstrip('/')}/Personnel/Review"
            file_size = child.get("size", 0)
            dest_file_path = f"{file_dest_path.rstrip('/')}/{name}"
            prev = dest_index.get(dest_file_path)
            if prev is not None and prev[0] != file_size:
                print(
                    f"{prefix}  WARN  destination collision for {_mask_filename(name)} "
                    f"(source sizes: {prev[0]} vs {file_size})"
                )
            else:
                dest_index[dest_file_path] = (file_size, child_source)
            if not dry_run:
                size_str = _format_size(file_size)
                print(f"{prefix}[FILE] {_mask_filename(name)} ({size_str})")
            if not dry_run:
                try:
                    existing = get_item_by_path(session, drive_id, dest_file_path)
                    if existing is not None:
                        existing_size = existing.get("size")
                        if existing_size == file_size:
                            print(f"{prefix}  SKIP  {_mask_filename(name)} exists at destination ({size_str})")
                            continue
                        print(
                            f"{prefix}  SKIP  {_mask_filename(name)} size mismatch "
                            f"(source={file_size} bytes, dest={existing_size} bytes); "
                            f"preserving destination metadata"
                        )
                        continue
                    if file_dest_path != dest_path:
                        file_dest_folder = ensure_folder(session, drive_id, file_dest_path)
                        file_dest_id = file_dest_folder["id"]
                    else:
                        file_dest_id = dest_id
                    monitor = copy_item(session, drive_id, child["id"], file_dest_id, name)
                    if monitor:
                        success, failure_code = wait_for_copy(monitor)
                        if not success and failure_code == "nameAlreadyExists":
                            # Preserve destination metadata by skipping on conflicts.
                            existing_after = get_item_by_path(session, drive_id, dest_file_path)
                            existing_after_size = (
                                existing_after.get("size") if existing_after is not None else None
                            )
                            print(
                                f"{prefix}  SKIP  {_mask_filename(name)} conflict/nameAlreadyExists "
                                f"(source={file_size} bytes, dest={existing_after_size} bytes); "
                                f"preserving destination metadata"
                            )
                            success = True
                        if not success:
                            if failure_code:
                                print(
                                    f"{prefix}  ERROR copy failed for {_mask_filename(name)} "
                                    f"(code={failure_code})"
                                )
                            errors += 1
                            continue
                    copied += 1
                    time.sleep(THROTTLE_SECONDS)
                except requests.HTTPError as e:
                    print(f"{prefix}  ERROR: {e}")
                    errors += 1
            else:
                copied += 1

    return copied, errors


def _format_size(size: int) -> str:
    if size < 1024:
        return f"{size} B"
    if size < 1024 * 1024:
        return f"{size / 1024:.1f} KB"
    return f"{size / (1024 * 1024):.1f} MB"


# --- Migration plan -----------------------------------------------------------

def load_mapping(csv_path: str) -> list[dict]:
    """Load folder_mapping.csv into a list of dicts."""
    rows = []
    with open(csv_path, newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for row in reader:
            rows.append(row)
    return rows


def plan_migration(rows: list[dict], root: str) -> tuple[list[dict], list[dict], list[dict]]:
    """
    Categorize rows into:
      - active_migrations: folders with a new_path to copy to
      - inactive_moves: folders to move to Not-Migrated - Inactive
      - skipped: unmatched folders (no action)
    """
    active = []
    inactive = []
    skipped = []

    for row in rows:
        confidence = row.get("Match Confidence", "")
        new_path = row.get("New Path", "")
        old_path = row.get("Old Path", "")

        if confidence == "inactive":
            inactive.append(row)
        elif new_path and confidence in ("exact", "partial", "manual"):
            active.append(row)
        else:
            skipped.append(row)

    return active, inactive, skipped


# --- Main execution -----------------------------------------------------------

def migrate_active(session: requests.Session, drive_id: str, root: str,
                   active: list[dict], dry_run: bool,
                   subfolder_destinations: dict[str, str],
                   partial_migration: bool) -> tuple[int, int, int]:
    """
        For each active mapping:
            1. Copy folder contents to new path.
            2. Optionally move old folder to Migrated/ area.
    Returns (total_files_copied, folders_migrated, errors).
    """
    total_copied = 0
    total_migrated = 0
    total_errors = 0

    migrated_base = f"{root.strip('/')}/{MIGRATED_FOLDER}"

    for row in active:
        old_path = row["Old Path"]
        new_path = row["New Path"]
        confidence = row.get("Match Confidence", "")

        print(f"\n{'=' * 70}")
        print(f"MIGRATE: {_mask_filename(old_path)}")
        print(f"     TO: {_mask_filename(new_path)}")
        print(f"  TRUST: {_mask_filename(confidence)}")
        print(f"{'=' * 70}")

        # Resolve old folder
        source_item = get_item_by_path(session, drive_id, old_path)
        if not source_item:
            print(f"  ERROR: Source folder not found: {_mask_filename(old_path)}")
            total_errors += 1
            continue

        # Step 1: Copy contents to new location
        if dry_run:
            print(f"  Would copy contents from: {_mask_filename(old_path)}")
            print(f"                     into: {_mask_filename(new_path)} (using {SUBFOLDER_CSV_FILE})")
            copied, errors = copy_folder_contents(
                session, drive_id, source_item, old_path, new_path, dry_run=True,
                subfolder_destinations=subfolder_destinations
            )
            total_copied += copied
            total_errors += errors
        else:
            copied, errors = copy_folder_contents(
                session, drive_id, source_item, old_path, new_path, dry_run=False,
                subfolder_destinations=subfolder_destinations
            )
            total_copied += copied
            total_errors += errors

        # Step 2: Move old folder to Migrated/
        # Preserve relative path structure under Migrated/
        # e.g. "Personnel Files/Active/France/Employees/FOO Bar"
        #    -> "Personnel Files/Active/Migrated/France/Employees/FOO Bar"
        # We strip the root prefix and put it under Migrated
        rel_from_root = old_path[len(root):].strip("/") if old_path.startswith(root) else old_path
        migrated_dest = f"{migrated_base}/{rel_from_root}"
        migrated_parent = migrated_dest.rsplit("/", 1)[0]

        if partial_migration:
            print("  Partial migration: old folder remains in place (not moved to Migrated/)")
        else:
            if dry_run:
                print(f"  Would move old folder to: {_mask_filename(migrated_dest)}")
            else:
                parent_item = ensure_folder(session, drive_id, migrated_parent)
                move_item(session, drive_id, source_item["id"], parent_item["id"])
                print(f"  Moved old folder to: {_mask_filename(migrated_dest)}")

            total_migrated += 1

    return total_copied, total_migrated, total_errors


def migrate_inactive(session: requests.Session, drive_id: str, root: str,
                     inactive: list[dict], dry_run: bool) -> tuple[int, int]:
    """
    Move inactive/terminated folders to Not-Migrated - Inactive/.
    Returns (folders_moved, errors).
    """
    total_moved = 0
    total_errors = 0

    inactive_base = f"{root.strip('/')}/{INACTIVE_FOLDER}"

    for row in inactive:
        old_path = row["Old Path"]

        # Preserve structure under the inactive folder
        rel_from_root = old_path[len(root):].strip("/") if old_path.startswith(root) else old_path
        inactive_dest = f"{inactive_base}/{rel_from_root}"
        inactive_parent = inactive_dest.rsplit("/", 1)[0]

        print(f"\n  INACTIVE: {_mask_filename(old_path)}")
        print(f"   MOVE TO: {_mask_filename(inactive_dest)}")

        if dry_run:
            total_moved += 1
            continue

        source_item = get_item_by_path(session, drive_id, old_path)
        if not source_item:
            print(f"    ERROR: Source not found: {_mask_filename(old_path)}")
            total_errors += 1
            continue

        try:
            parent_item = ensure_folder(session, drive_id, inactive_parent)
            move_item(session, drive_id, source_item["id"], parent_item["id"])
            print(f"    Moved.")
            total_moved += 1
            time.sleep(THROTTLE_SECONDS)
        except requests.HTTPError as e:
            print(f"    ERROR: {e}")
            total_errors += 1

    return total_moved, total_errors


def main():
    if "--execute" in sys.argv:
        dry_run = False
    else:
        dry_run = True  # default to dry-run

    partial_migration = "--full" not in sys.argv  # default to partial, only full if --full
    use_test_group = "--all-users" not in sys.argv or "--use-test-group" in sys.argv
    test_emails: set[str] = set()
    if use_test_group:
        test_emails = load_test_users()
        if not test_emails:
            sys.exit(2)

    if dry_run:
        print("=" * 70)
        print("DRY RUN — no changes will be made")
        if not partial_migration:
            print("FULL MIGRATION — archive step will be included")
        else:
            print("PARTIAL MIGRATION (default) — archive step will be skipped")
        if use_test_group:
            print(f"TEST GROUP — scoped to {len(test_emails)} user(s) from test_users.json")
        print("=" * 70)
    else:
        print("=" * 70)
        print("EXECUTING MIGRATION — changes will be applied to SharePoint")
        if not partial_migration:
            print("FULL MIGRATION — archive step will be included")
        else:
            print("PARTIAL MIGRATION (default) — archive step will be skipped")
        if use_test_group:
            print(f"TEST GROUP — scoped to {len(test_emails)} user(s) from test_users.json")
        print("=" * 70)

    # Load mapping
    csv_file_to_use = "folder_mapping_test.csv" if use_test_group else CSV_FILE
    if not os.path.exists(csv_file_to_use):
        print(f"ERROR: {csv_file_to_use} not found.", file=sys.stderr)
        if use_test_group:
            print("       Run: python map_old_to_new_folders.py --csv", file=sys.stderr)
        else:
            print("       Run: python map_old_to_new_folders.py --all-users --csv", file=sys.stderr)
        sys.exit(1)

    rows = load_mapping(csv_file_to_use)
    if not os.path.exists(SUBFOLDER_CSV_FILE):
        print(f"ERROR: {SUBFOLDER_CSV_FILE} not found.", file=sys.stderr)
        print("       Run: python map_old_to_new_folders.py --csv", file=sys.stderr)
        sys.exit(1)
    subfolder_destinations = load_subfolder_destinations(SUBFOLDER_CSV_FILE)
    root = os.environ.get("SHAREPOINT_FOLDER_PATH", "Personnel Files/Active")

    # Categorize
    active, inactive, skipped = plan_migration(rows, root)

    print(f"\nMigration plan:")
    print(f"  Active migrations (copy + move to Migrated/):   {len(active)}")
    print(f"  Inactive (move to Not-Migrated - Inactive/):    {len(inactive)}")
    print(f"  Skipped (unmatched, no action):                 {len(skipped)}")

    if skipped:
        print(f"\n  Skipped folders (manual review needed):")
        for row in skipped:
            print(f"    - {row['Old Path']}")

    # Connect to SharePoint
    print("\nConnecting to SharePoint ...")
    token = graph_token()
    session = graph_session(token)
    hostname = os.environ["SHAREPOINT_HOSTNAME"]
    site_path = os.environ["SHAREPOINT_SITE_PATH"]
    drive_name = os.environ["SHAREPOINT_DRIVE_NAME"]
    site_id = resolve_site_id(session, hostname, site_path)
    drive_id = resolve_drive_id(session, site_id, drive_name)
    print(f"  Drive: {drive_name} ({drive_id})")

    # Execute active migrations
    if active:
        print(f"\n{'=' * 70}")
        print(f"ACTIVE MIGRATIONS ({len(active)} folders)")
        print(f"{'=' * 70}")
        files_copied, folders_migrated, errors = migrate_active(
                        session, drive_id, root, active, dry_run, subfolder_destinations,
                        partial_migration
        )
        print(f"\n  Summary: {files_copied} files copied, "
                            f"{folders_migrated} folders archived, {errors} errors")

    # Execute inactive moves
    if inactive:
        print(f"\n{'=' * 70}")
        print(f"INACTIVE MOVES ({len(inactive)} folders)")
        print(f"{'=' * 70}")
        moved, errors = migrate_inactive(session, drive_id, root, inactive, dry_run)
        print(f"\n  Summary: {moved} folders moved, {errors} errors")

    # Done
    mode_label = "DRY RUN complete" if dry_run else "MIGRATION complete"
    print(f"\n{'=' * 70}")
    print(f"{mode_label}")
    print(f"{'=' * 70}")


if __name__ == "__main__":
    main()
