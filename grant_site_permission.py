"""
One-time helper: grant the BambooHR_Integration Azure app `write` permission
on a single SharePoint site, so it can use `Sites.Selected` to read/write
files in that site only.

Run this AS A SHAREPOINT/GLOBAL ADMIN. It uses MSAL's device-code flow so the
admin signs in with their own credentials in a browser; no admin secrets are
stored locally.

What it does:
  1. Admin signs in via device code (delegated auth, scope Sites.FullControl.All).
  2. Resolves the target site by hostname + server-relative path.
  3. POSTs to /sites/{site-id}/permissions to grant the app id `write` access.
  4. Prints existing grants so you can verify / re-run safely.

Env vars used (from .env):
  AZURE_APP_CLIENT_ID    -> the app being granted access
  AZURE_TENANT_ID        -> tenant
  SHAREPOINT_HOSTNAME    -> e.g. hoops3d.sharepoint.com
  SHAREPOINT_SITE_PATH   -> e.g. /people-experience

Optional:
  GRANT_APP_DISPLAY_NAME -> display name to store on the grant (defaults to
                            "BambooHR_Integration")

The MSAL "public client" used for the admin device-code login is the
well-known Microsoft Azure CLI client id (04b07795-...). No additional Azure
app registration is required for the admin's sign-in.
"""

from __future__ import annotations

import os
import sys
import json

import requests
from msal import PublicClientApplication

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

GRAPH = "https://graph.microsoft.com/v1.0"
# "Microsoft Graph Command Line Tools" public client id — preauthorized for
# Microsoft Graph in all tenants, so we don't have to register our own
# delegated app just for this one-time admin grant.
AZ_CLI_CLIENT_ID = "14d82eec-204b-4c2f-b7e8-296a70dab67e"


def admin_login(tenant_id: str) -> str:
    """Interactive device-code login. Returns an access token for Graph."""
    app = PublicClientApplication(
        AZ_CLI_CLIENT_ID,
        authority=f"https://login.microsoftonline.com/{tenant_id}",
    )
    # Sites.FullControl.All is needed to manage /sites/{id}/permissions.
    scopes = ["https://graph.microsoft.com/Sites.FullControl.All"]
    flow = app.initiate_device_flow(scopes=scopes)
    if "user_code" not in flow:
        raise RuntimeError(f"Failed to start device-code flow: {flow}")
    print("\n=== ADMIN SIGN-IN REQUIRED ===")
    print(flow["message"])
    print("(Waiting for you to complete sign-in in the browser...)\n")
    result = app.acquire_token_by_device_flow(flow)
    if "access_token" not in result:
        raise RuntimeError(f"Auth failed: {result}")
    return result["access_token"]


def graph_session(token: str) -> requests.Session:
    s = requests.Session()
    s.headers.update({
        "Authorization": f"Bearer {token}",
        "Accept": "application/json",
        "Content-Type": "application/json",
    })
    return s


def resolve_site_id(session: requests.Session, hostname: str, site_path: str) -> str:
    site_path = "/" + site_path.strip("/")
    r = session.get(f"{GRAPH}/sites/{hostname}:{site_path}")
    r.raise_for_status()
    return r.json()["id"]


def list_permissions(session: requests.Session, site_id: str) -> list[dict]:
    r = session.get(f"{GRAPH}/sites/{site_id}/permissions")
    if r.status_code >= 400:
        raise RuntimeError(f"list permissions failed: {r.status_code} {r.text}")
    return r.json().get("value", [])


def grant_app_permission(session: requests.Session, site_id: str,
                         app_client_id: str, display_name: str,
                         role: str = "write") -> dict:
    body = {
        "roles": [role],
        "grantedToIdentities": [{
            "application": {"id": app_client_id, "displayName": display_name}
        }],
    }
    r = session.post(f"{GRAPH}/sites/{site_id}/permissions", data=json.dumps(body))
    if r.status_code >= 400:
        raise RuntimeError(f"Grant failed: {r.status_code} {r.text}")
    return r.json()


def main() -> int:
    required = ["AZURE_APP_CLIENT_ID", "AZURE_TENANT_ID",
                "SHAREPOINT_HOSTNAME", "SHAREPOINT_SITE_PATH"]
    missing = [k for k in required if not os.environ.get(k)]
    if missing:
        print(f"ERROR: missing env vars: {', '.join(missing)}", file=sys.stderr)
        return 2

    app_id = os.environ["AZURE_APP_CLIENT_ID"]
    tenant_id = os.environ["AZURE_TENANT_ID"]
    hostname = os.environ["SHAREPOINT_HOSTNAME"]
    site_path = os.environ["SHAREPOINT_SITE_PATH"]
    display_name = os.environ.get("GRANT_APP_DISPLAY_NAME", "BambooHR_Integration")

    print(f"Granting '{display_name}' ({app_id}) write access on "
          f"{hostname}{site_path}")
    token = admin_login(tenant_id)
    session = graph_session(token)

    print("\nResolving site ...")
    site_id = resolve_site_id(session, hostname, site_path)
    print(f"  site id: {site_id}")

    print("\nCurrent permissions on this site:")
    try:
        existing = list_permissions(session, site_id)
    except RuntimeError as e:
        print(f"  (could not list: {e})")
        existing = []
    if not existing:
        print("  (none reported)")
    for p in existing:
        ids = p.get("grantedToIdentitiesV2") or p.get("grantedToIdentities") or []
        apps = [i.get("application", {}) for i in ids if "application" in i]
        for a in apps:
            print(f"  - permissionId={p.get('id')} roles={p.get('roles')} "
                  f"appId={a.get('id')} name={a.get('displayName')}")

    already = any(
        a.get("id") == app_id
        for p in existing
        for i in (p.get("grantedToIdentitiesV2") or p.get("grantedToIdentities") or [])
        for a in [i.get("application", {})]
    )
    if already:
        print("\nApp already has a permission entry on this site. Nothing to do.")
        return 0

    print("\nGranting write permission ...")
    granted = grant_app_permission(session, site_id, app_id, display_name, "write")
    print("  OK")
    print(f"  permissionId: {granted.get('id')}")
    print(f"  roles:        {granted.get('roles')}")
    print("\nDone. The BambooHR_Integration app can now read/write this site.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
