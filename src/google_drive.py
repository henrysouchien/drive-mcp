"""
Google Drive API client for MCP server.
Extracted from drive_indexer.ipynb
"""

import io
import pickle
from pathlib import Path
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from pypdf import PdfReader

# Full access to Drive files
SCOPES = [
    'https://www.googleapis.com/auth/drive',
]

# Google Workspace MIME types that need export
GOOGLE_DOC_MIME = 'application/vnd.google-apps.document'
GOOGLE_SHEET_MIME = 'application/vnd.google-apps.spreadsheet'
GOOGLE_SLIDES_MIME = 'application/vnd.google-apps.presentation'

# Text-based MIME types we can read directly
TEXT_MIME_TYPES = [
    'text/plain',
    'text/markdown',
    'text/csv',
    'text/html',
    'application/json',
    'application/xml',
    'text/xml',
]

# Paths relative to this module's parent directory
BASE_DIR = Path(__file__).parent.parent
CREDENTIALS_FILE = BASE_DIR / 'drive_credentials.json'
TOKEN_FILE = BASE_DIR / 'token.pickle'

_cached_creds = None


def _get_missing_scopes(creds) -> list[str]:
    """Return required scopes that are missing from credentials."""
    granted = set()
    if getattr(creds, 'scopes', None):
        granted.update(creds.scopes)
    if getattr(creds, 'granted_scopes', None):
        granted.update(creds.granted_scopes)
    return [scope for scope in SCOPES if scope not in granted]


def _get_credentials():
    """Load, refresh, or create OAuth credentials with required scopes."""
    global _cached_creds

    creds = _cached_creds
    if creds is None and TOKEN_FILE.exists():
        with open(TOKEN_FILE, 'rb') as token:
            creds = pickle.load(token)

    missing_scopes = _get_missing_scopes(creds) if creds else []
    if missing_scopes:
        creds = None
        _cached_creds = None
        if TOKEN_FILE.exists():
            TOKEN_FILE.unlink()

    should_save_token = False
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
            should_save_token = True
        else:
            if not CREDENTIALS_FILE.exists():
                raise FileNotFoundError(
                    f"Credentials file not found at {CREDENTIALS_FILE}. "
                    "Please copy your drive_credentials.json to the drive-mcp folder."
                )
            flow = InstalledAppFlow.from_client_secrets_file(
                str(CREDENTIALS_FILE), SCOPES)
            creds = flow.run_local_server(port=0)
            should_save_token = True

    if should_save_token:
        with open(TOKEN_FILE, 'wb') as token:
            pickle.dump(creds, token)

    _cached_creds = creds
    return creds


def authenticate():
    """Authenticate with Google Drive API and return service object."""
    creds = _get_credentials()
    return build('drive', 'v3', credentials=creds)


def rename_file(service, file_id: str, new_name: str) -> dict:
    """Rename a file in Google Drive."""
    return service.files().update(
        fileId=file_id,
        body={'name': new_name},
        supportsAllDrives=True,
        fields='id, name'
    ).execute()


def move_file(service, file_id: str, new_parent_id: str) -> dict:
    """Move a file to a different folder in Google Drive."""
    file_info = service.files().get(
        fileId=file_id,
        fields='parents',
        supportsAllDrives=True
    ).execute()
    previous_parents = ','.join(file_info.get('parents', []))
    return service.files().update(
        fileId=file_id,
        addParents=new_parent_id,
        removeParents=previous_parents,
        supportsAllDrives=True,
        fields='id, name, parents'
    ).execute()


def get_folder_id(service, folder_name: str) -> str | None:
    """Get the ID of a folder by name."""
    query = f"name = '{folder_name}' and mimeType = 'application/vnd.google-apps.folder'"
    results = service.files().list(
        q=query,
        spaces='drive',
        fields="files(id, name)"
    ).execute()
    folders = results.get('files', [])
    if not folders:
        return None
    return folders[0]['id']


def list_files_in_folder(service, folder_id: str) -> list[dict]:
    """List all files in a folder (non-recursive)."""
    query = f"'{folder_id}' in parents"
    results = service.files().list(
        q=query,
        fields="files(id, name, mimeType, webViewLink, modifiedTime)"
    ).execute()
    return results.get('files', [])


def list_files_recursive(service, folder_id: str, path: str = "") -> list[dict]:
    """Recursively list all files in a folder and subfolders."""
    all_files = []
    query = f"'{folder_id}' in parents"
    results = service.files().list(
        q=query,
        fields="files(id, name, mimeType, webViewLink, modifiedTime)"
    ).execute()
    items = results.get('files', [])

    for item in items:
        current_path = f"{path}/{item['name']}" if path else item['name']
        if item['mimeType'] == 'application/vnd.google-apps.folder':
            # Recurse into subfolder
            all_files.extend(list_files_recursive(service, item['id'], current_path))
        else:
            all_files.append({
                'id': item['id'],
                'name': item['name'],
                'path': current_path,
                'mimeType': item['mimeType'],
                'webViewLink': item.get('webViewLink', ''),
                'modifiedTime': item.get('modifiedTime', '')
            })

    return all_files


def search_files(service, query: str, max_results: int = 20) -> list[dict]:
    """Search for files by name."""
    search_query = f"name contains '{query}'"
    results = service.files().list(
        q=search_query,
        spaces='drive',
        fields="files(id, name, mimeType, webViewLink, modifiedTime)",
        pageSize=max_results
    ).execute()
    return results.get('files', [])


def find_file_by_name(service, file_name: str) -> dict | None:
    """Find a file by name. Returns the first match."""
    query = f"name = '{file_name}'"
    results = service.files().list(
        q=query,
        spaces='drive',
        fields="files(id, name, mimeType)"
    ).execute()
    files = results.get('files', [])
    return files[0] if files else None


def read_file_content(service, file_id: str, mime_type: str, max_chars: int = 100000) -> str:
    """
    Read file content based on MIME type.

    - Google Docs → export as plain text
    - Google Sheets → export as CSV
    - PDFs → extract text
    - Text files → direct download
    """
    try:
        # Google Docs → export as plain text
        if mime_type == GOOGLE_DOC_MIME:
            request = service.files().export_media(
                fileId=file_id,
                mimeType='text/plain'
            )
            content = request.execute()
            text = content.decode('utf-8')

        # Google Sheets → export as CSV
        elif mime_type == GOOGLE_SHEET_MIME:
            request = service.files().export_media(
                fileId=file_id,
                mimeType='text/csv'
            )
            content = request.execute()
            text = content.decode('utf-8')

        # Google Slides → export as plain text
        elif mime_type == GOOGLE_SLIDES_MIME:
            request = service.files().export_media(
                fileId=file_id,
                mimeType='text/plain'
            )
            content = request.execute()
            text = content.decode('utf-8')

        # PDF → download and extract text
        elif mime_type == 'application/pdf':
            request = service.files().get_media(fileId=file_id)
            buffer = io.BytesIO()
            downloader = MediaIoBaseDownload(buffer, request)
            done = False
            while not done:
                _, done = downloader.next_chunk()
            buffer.seek(0)

            reader = PdfReader(buffer)
            text_parts = []
            for page in reader.pages:
                page_text = page.extract_text()
                if page_text:
                    text_parts.append(page_text)
            text = '\n'.join(text_parts)

        # Text-based files → direct download
        elif mime_type in TEXT_MIME_TYPES or mime_type.startswith('text/'):
            request = service.files().get_media(fileId=file_id)
            content = request.execute()
            text = content.decode('utf-8')

        else:
            return f"Cannot read file with MIME type: {mime_type}"

        # Truncate if too long
        if len(text) > max_chars:
            text = text[:max_chars] + f"\n\n... [truncated, {len(text)} total chars]"

        return text

    except Exception as e:
        return f"Error reading file: {str(e)}"


def read_file_by_name(service, file_name: str, max_chars: int = 100000) -> str:
    """Find a file by name and read its content."""
    file_info = find_file_by_name(service, file_name)
    if not file_info:
        return f"File not found: {file_name}"

    return read_file_content(service, file_info['id'], file_info['mimeType'], max_chars)


def read_file_by_id(service, file_id: str, max_chars: int = 100000) -> str:
    """Read a file by ID."""
    # First get the file metadata to know the MIME type
    file_info = service.files().get(fileId=file_id, fields='id, name, mimeType').execute()
    return read_file_content(service, file_id, file_info['mimeType'], max_chars)


# Quick test when run directly
if __name__ == "__main__":
    print("Testing Google Drive connection...")
    service = authenticate()
    print("✓ Authenticated successfully")

    # List root files
    results = service.files().list(pageSize=5, fields="files(id, name)").execute()
    files = results.get('files', [])
    print(f"✓ Found {len(files)} files in root")
    for f in files:
        print(f"  - {f['name']}")
