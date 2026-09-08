"""
Google helpers (common/google_utils.py).

Reuses the IntelliBI service-account approach (same key as
IntelliBI_Operations_Automation): one Credentials object scoped for BOTH Sheets
and Drive, so the same client can read the source workbook and upload the
generated report to Drive.

  get_services()            -> (sheets_service, drive_service)
  read_tab(sheets, id, tab) -> list[list] rows (values)
  upload_xlsx(drive, ...)   -> (file_id, webViewLink)
"""
from __future__ import annotations
import os
import paths
import config_loader as cfg

_READ_SCOPES = ["https://www.googleapis.com/auth/spreadsheets.readonly"]
_DRIVE_SCOPES = ["https://www.googleapis.com/auth/drive"]


def _service_account_file() -> str:
    name = cfg.get("google.service_account_file", "service_account.json")
    # An absolute path is honoured; otherwise resolve inside credentials/.
    if os.path.isabs(name):
        return name
    return str(paths.CREDENTIALS_DIR / name)


def get_services():
    """Return (sheets_service, drive_service).

    Sheets is read directly as the service account. Drive is authenticated by
    IMPERSONATING a real Workspace user (google.impersonate_user) via domain-wide
    delegation — because a service account has no storage quota of its own, so a
    file it tried to own would fail with 'storageQuotaExceeded'. Impersonation
    makes the uploaded report owned by that user (who has quota), the same
    approach the Follow-Up report uses.
    """
    from google.oauth2 import service_account
    from googleapiclient.discovery import build
    sa = _service_account_file()
    if not os.path.exists(sa):
        raise FileNotFoundError(
            f"Service-account key not found: {sa}\n"
            f"  Place the IntelliBI service_account.json in the credentials/ folder "
            f"(same key used by IntelliBI_Operations_Automation), and share the source "
            f"sheet (Viewer) with its client_email.")

    read_creds = service_account.Credentials.from_service_account_file(sa, scopes=_READ_SCOPES)
    sheets = build("sheets", "v4", credentials=read_creds, cache_discovery=False)

    drive_creds = service_account.Credentials.from_service_account_file(sa, scopes=_DRIVE_SCOPES)
    impersonate = cfg.get("google.impersonate_user", "")
    if impersonate:
        drive_creds = drive_creds.with_subject(impersonate)
    drive = build("drive", "v3", credentials=drive_creds, cache_discovery=False)
    return sheets, drive


def find_or_create_folder(drive, parent_id: str, name: str) -> str:
    """Return the ID of the sub-folder `name` under `parent_id`, creating it if it
    does not exist. Lets the report upload into the right per-type sub-folder while
    the operator only has to share the single parent folder with the service account."""
    safe = name.replace("'", "\\'")
    q = (f"name = '{safe}' and '{parent_id}' in parents and "
         f"mimeType = 'application/vnd.google-apps.folder' and trashed = false")
    found = drive.files().list(q=q, fields="files(id,name)",
                               supportsAllDrives=True,
                               includeItemsFromAllDrives=True).execute().get("files", [])
    if found:
        return found[0]["id"]
    meta = {"name": name, "parents": [parent_id],
            "mimeType": "application/vnd.google-apps.folder"}
    f = drive.files().create(body=meta, fields="id", supportsAllDrives=True).execute()
    return f.get("id")


def read_tab(sheets, spreadsheet_id: str, tab: str) -> list:
    """Return all rows of a tab as a list of lists (values only, un-padded)."""
    rng = f"'{tab}'"
    resp = sheets.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id, range=rng,
        valueRenderOption="UNFORMATTED_VALUE",
        dateTimeRenderOption="FORMATTED_STRING",
    ).execute()
    return resp.get("values", [])


def upload_xlsx(drive, local_path: str, name: str, folder_id: str,
                convert_to_sheet: bool = False):
    """Upload local_path to Drive folder_id. Returns (file_id, web_view_link).

    If convert_to_sheet is True the .xlsx is converted to a native Google Sheet;
    otherwise it is stored as an .xlsx file. Any existing file with the same name
    in the folder is updated in place so links stay stable across runs.
    """
    from googleapiclient.http import MediaFileUpload
    XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    target_mime = "application/vnd.google-apps.spreadsheet" if convert_to_sheet else XLSX

    q = (f"name = '{name.replace(chr(39), chr(92)+chr(39))}' and "
         f"'{folder_id}' in parents and trashed = false")
    existing = drive.files().list(q=q, fields="files(id,name)",
                                  supportsAllDrives=True,
                                  includeItemsFromAllDrives=True).execute().get("files", [])
    media = MediaFileUpload(local_path, mimetype=XLSX, resumable=False)
    if existing:
        fid = existing[0]["id"]
        f = drive.files().update(fileId=fid, media_body=media,
                                 supportsAllDrives=True,
                                 fields="id,webViewLink").execute()
    else:
        meta = {"name": name, "parents": [folder_id], "mimeType": target_mime}
        f = drive.files().create(body=meta, media_body=media,
                                 supportsAllDrives=True,
                                 fields="id,webViewLink").execute()
    return f.get("id"), f.get("webViewLink")
