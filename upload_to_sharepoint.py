"""
Prototype: Upload a single local file to a specific SharePoint folder via
Microsoft Graph (app-only auth, client credentials).

Target folder (decoded from the share URL):
  Host:   hoops3d.sharepoint.com
  Site:   /people-experience
  Drive:  "HR Only Documents"
  Folder: Personnel Files/Active/USA/Toolkits Group/
          Wheeler, Jeffrey_DOH_01.17.2022/Personnel

Required env vars:
  AZURE_APP_CLIENT_ID
  AZURE_TENANT_ID
  AZURE_CLIENT_SECRET
  SHAREPOINT_HOSTNAME       e.g. hoops3d.sharepoint.com
  SHAREPOINT_SITE_PATH      e.g. /people-experience
  SHAREPOINT_DRIVE_NAME     e.g. "HR Only Documents"
  SHAREPOINT_FOLDER_PATH    e.g. "Personnel Files/Active/.../Personnel"

Azure app must have Microsoft Graph application permission
`Sites.Selected` (preferred, scoped via POST /sites/{id}/permissions) or
`Sites.ReadWrite.All`, with admin consent granted.

Usage:
  python upload_to_sharepoint.py <local_file> [remote_filename]
"""

import os
import sys
from pathlib import Path
from urllib.parse import quote

import requests
from msal import ConfidentialClientApplication

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

GRAPH = "https://graph.microsoft.com/v1.0"
# Graph small-file PUT limit is 4 MB; use upload session above this.
SMALL_FILE_LIMIT = 4 * 1024 * 1024
CHUNK_SIZE = 5 * 1024 * 1024  # 5 MB, multiple of 320 KiB as Graph requires


def get_access_token() -> str:
    client_id = os.environ["AZURE_APP_CLIENT_ID"]
    tenant_id = os.environ["AZURE_TENANT_ID"]
    client_secret = os.environ["AZURE_CLIENT_SECRET"]
    app = ConfidentialClientApplication(
        client_id,
        authority=f"https://login.microsoftonline.com/{tenant_id}",
        client_credential=client_secret,
    )
    result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    if "access_token" not in result:
        raise RuntimeError(f"Failed to acquire token: {result}")
    return result["access_token"]


def graph_get(session: requests.Session, path: str) -> dict:
    r = session.get(f"{GRAPH}{path}")
    r.raise_for_status()
    return r.json()


def resolve_site_id(session: requests.Session, hostname: str, site_path: str) -> str:
    # GET /sites/{hostname}:{server-relative-path}
    site_path = "/" + site_path.strip("/")
    data = graph_get(session, f"/sites/{hostname}:{site_path}")
    return data["id"]


def resolve_drive_id(session: requests.Session, site_id: str, drive_name: str) -> str:
    data = graph_get(session, f"/sites/{site_id}/drives")
    for d in data.get("value", []):
        if d.get("name") == drive_name:
            return d["id"]
    available = ", ".join(d.get("name", "?") for d in data.get("value", []))
    raise RuntimeError(f"Drive '{drive_name}' not found. Available drives: {available}")


def upload_small(session: requests.Session, drive_id: str, folder_path: str,
                 local_path: Path, remote_name: str) -> dict:
    # PUT /drives/{drive-id}/root:/{folder-path}/{filename}:/content
    target = f"{folder_path.strip('/')}/{remote_name}"
    url = f"{GRAPH}/drives/{drive_id}/root:/{quote(target)}:/content"
    with open(local_path, "rb") as fh:
        r = session.put(
            url,
            data=fh.read(),
            headers={"Content-Type": "application/octet-stream"},
        )
    r.raise_for_status()
    return r.json()


def upload_large(session: requests.Session, drive_id: str, folder_path: str,
                 local_path: Path, remote_name: str) -> dict:
    target = f"{folder_path.strip('/')}/{remote_name}"
    create_url = f"{GRAPH}/drives/{drive_id}/root:/{quote(target)}:/createUploadSession"
    r = session.post(
        create_url,
        json={"item": {"@microsoft.graph.conflictBehavior": "replace", "name": remote_name}},
    )
    r.raise_for_status()
    upload_url = r.json()["uploadUrl"]

    size = local_path.stat().st_size
    sent = 0
    last_response = None
    # Upload session does NOT use the Graph auth header; the URL is pre-signed.
    bare = requests.Session()
    with open(local_path, "rb") as fh:
        while sent < size:
            chunk = fh.read(CHUNK_SIZE)
            if not chunk:
                break
            start = sent
            end = sent + len(chunk) - 1
            headers = {
                "Content-Length": str(len(chunk)),
                "Content-Range": f"bytes {start}-{end}/{size}",
            }
            resp = bare.put(upload_url, data=chunk, headers=headers)
            if resp.status_code not in (200, 201, 202):
                raise RuntimeError(
                    f"Chunk upload failed at bytes {start}-{end}: "
                    f"{resp.status_code} {resp.text}"
                )
            sent += len(chunk)
            last_response = resp
    return last_response.json() if last_response is not None else {}


def main() -> int:
    if len(sys.argv) < 2:
        print("Usage: python upload_to_sharepoint.py <local_file> [remote_filename]",
              file=sys.stderr)
        return 2
    local_path = Path(sys.argv[1]).expanduser().resolve()
    if not local_path.is_file():
        print(f"ERROR: local file not found: {local_path}", file=sys.stderr)
        return 1
    remote_name = sys.argv[2] if len(sys.argv) > 2 else local_path.name

    required = [
        "AZURE_APP_CLIENT_ID", "AZURE_TENANT_ID", "AZURE_CLIENT_SECRET",
        "SHAREPOINT_HOSTNAME", "SHAREPOINT_SITE_PATH",
        "SHAREPOINT_DRIVE_NAME", "SHAREPOINT_FOLDER_PATH",
    ]
    missing = [k for k in required if not os.environ.get(k)]
    if missing:
        print(f"ERROR: missing env vars: {', '.join(missing)}", file=sys.stderr)
        return 2

    hostname = os.environ["SHAREPOINT_HOSTNAME"]
    site_path = os.environ["SHAREPOINT_SITE_PATH"]
    drive_name = os.environ["SHAREPOINT_DRIVE_NAME"]
    folder_path = os.environ["SHAREPOINT_FOLDER_PATH"]

    print("Acquiring Graph token ...")
    token = get_access_token()
    session = requests.Session()
    session.headers.update({"Authorization": f"Bearer {token}", "Accept": "application/json"})

    print(f"Resolving site {hostname}{site_path} ...")
    site_id = resolve_site_id(session, hostname, site_path)
    print(f"  site id: {site_id}")

    print(f"Resolving drive '{drive_name}' ...")
    drive_id = resolve_drive_id(session, site_id, drive_name)
    print(f"  drive id: {drive_id}")

    size = local_path.stat().st_size
    print(f"Uploading {local_path} ({size} bytes) -> {folder_path}/{remote_name}")

    if size <= SMALL_FILE_LIMIT:
        item = upload_small(session, drive_id, folder_path, local_path, remote_name)
    else:
        item = upload_large(session, drive_id, folder_path, local_path, remote_name)

    print("Upload complete.")
    print(f"  webUrl: {item.get('webUrl')}")
    print(f"  id:     {item.get('id')}")
    print(f"  size:   {item.get('size')}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
