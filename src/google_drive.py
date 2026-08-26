"""
Google Drive API client for MCP server.
Extracted from drive_indexer.ipynb
"""

import io
import mimetypes
import os
import pickle
from pathlib import Path

from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload
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
GOOGLE_FOLDER_MIME = 'application/vnd.google-apps.folder'

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
CREDENTIALS_FILE = Path(
    os.environ.get("GOOGLE_CREDENTIALS_FILE") or BASE_DIR / 'drive_credentials.json'
)
TOKEN_FILE = Path(os.environ.get("GOOGLE_TOKEN_FILE") or BASE_DIR / 'token.pickle')

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


def set_file_parents(service, file_id: str, parent_ids: list[str]) -> dict:
    """Replace a file's parent folders."""
    file_info = service.files().get(
        fileId=file_id,
        fields='parents',
        supportsAllDrives=True
    ).execute()
    current_parents = ','.join(file_info.get('parents', []))
    kwargs = {
        'fileId': file_id,
        'removeParents': current_parents,
        'supportsAllDrives': True,
        'fields': 'id, name, parents',
    }
    if parent_ids:
        kwargs['addParents'] = ','.join(parent_ids)
    return service.files().update(**kwargs).execute()


def trash_file(service, file_id: str) -> dict:
    """Move a file or folder to trash."""
    return service.files().update(
        fileId=file_id,
        body={'trashed': True},
        supportsAllDrives=True,
        fields='id, name, trashed'
    ).execute()


def untrash_file(service, file_id: str) -> dict:
    """Restore a file or folder from trash."""
    return service.files().update(
        fileId=file_id,
        body={'trashed': False},
        supportsAllDrives=True,
        fields='id, name, trashed'
    ).execute()


def get_file_metadata(
    service,
    file_id: str,
    fields: str = 'id, name, mimeType, parents, webViewLink, trashed, size',
) -> dict | None:
    """Fetch file metadata by ID. Returns None when not found."""
    try:
        return service.files().get(
            fileId=file_id,
            fields=fields,
            supportsAllDrives=True,
        ).execute()
    except HttpError as exc:
        if getattr(exc, 'resp', None) is not None and exc.resp.status == 404:
            return None
        raise


def resolve_file(
    service,
    file_name: str | None = None,
    file_id: str | None = None,
    fields: str = 'id, name, mimeType, parents, webViewLink, trashed, size',
) -> dict | None:
    """Resolve a file by ID (preferred) or exact name."""
    if file_id:
        return get_file_metadata(service, file_id, fields=fields)
    if file_name:
        return find_file_by_name(service, file_name, fields=fields)
    raise ValueError("Provide file_name or file_id")


def resolve_folder(
    service,
    folder_name: str | None = None,
    folder_id: str | None = None,
) -> dict | None:
    """Resolve a folder by ID (preferred) or exact name."""
    if folder_id:
        info = get_file_metadata(
            service,
            folder_id,
            fields='id, name, mimeType, parents, webViewLink, trashed',
        )
        if not info:
            return None
        if info.get('mimeType') != GOOGLE_FOLDER_MIME:
            raise ValueError(f"ID is not a folder: {folder_id}")
        return info
    if folder_name:
        folder = find_file_by_name(
            service,
            folder_name,
            fields='id, name, mimeType, parents, webViewLink, trashed',
            folder_only=True,
        )
        return folder
    raise ValueError("Provide folder_name or folder_id")


def guess_mime_type(file_name: str, explicit: str | None = None) -> str:
    """Guess a MIME type from an explicit value or filename extension."""
    if explicit:
        return explicit
    guessed, _ = mimetypes.guess_type(file_name)
    return guessed or 'text/plain'


def write_file(
    service,
    file_name: str,
    content: str | bytes,
    *,
    parent_id: str | None = None,
    mime_type: str | None = None,
    file_id: str | None = None,
) -> dict:
    """Create or overwrite a Drive file with text/binary content."""
    resolved_mime = guess_mime_type(file_name, mime_type)
    data = content.encode('utf-8') if isinstance(content, str) else content
    media = MediaIoBaseUpload(
        io.BytesIO(data),
        mimetype=resolved_mime,
        resumable=False,
    )

    if file_id:
        kwargs = {
            'fileId': file_id,
            'media_body': media,
            'supportsAllDrives': True,
            'fields': 'id, name, mimeType, parents, webViewLink, size',
        }
        if file_name:
            kwargs['body'] = {'name': file_name}
        return service.files().update(**kwargs).execute()

    body: dict = {
        'name': file_name,
        'mimeType': resolved_mime,
    }
    if parent_id:
        body['parents'] = [parent_id]
    return service.files().create(
        body=body,
        media_body=media,
        supportsAllDrives=True,
        fields='id, name, mimeType, parents, webViewLink, size',
    ).execute()


def create_google_doc(
    service,
    title: str,
    *,
    content: str | None = None,
    parent_id: str | None = None,
) -> dict:
    """Create a Google Doc, optionally seeding content via text conversion."""
    body: dict = {
        'name': title,
        'mimeType': GOOGLE_DOC_MIME,
    }
    if parent_id:
        body['parents'] = [parent_id]

    if content is None:
        return service.files().create(
            body=body,
            supportsAllDrives=True,
            fields='id, name, mimeType, parents, webViewLink',
        ).execute()

    media = MediaIoBaseUpload(
        io.BytesIO(content.encode('utf-8')),
        mimetype='text/plain',
        resumable=False,
    )
    return service.files().create(
        body=body,
        media_body=media,
        supportsAllDrives=True,
        fields='id, name, mimeType, parents, webViewLink',
    ).execute()


def copy_file(
    service,
    file_id: str,
    *,
    new_name: str | None = None,
    parent_id: str | None = None,
) -> dict:
    """Copy a file, optionally renaming and/or placing under a new parent."""
    body: dict = {}
    if new_name:
        body['name'] = new_name
    if parent_id:
        body['parents'] = [parent_id]
    return service.files().copy(
        fileId=file_id,
        body=body,
        supportsAllDrives=True,
        fields='id, name, mimeType, parents, webViewLink',
    ).execute()


def share_file(
    service,
    file_id: str,
    email: str,
    role: str = 'reader',
    *,
    send_notification: bool = False,
) -> dict:
    """Share a file with a user email and role."""
    allowed_roles = {'reader', 'commenter', 'writer'}
    if role not in allowed_roles:
        raise ValueError(
            f"Invalid role '{role}'. Allowed: {', '.join(sorted(allowed_roles))}"
        )
    return service.permissions().create(
        fileId=file_id,
        body={
            'type': 'user',
            'role': role,
            'emailAddress': email,
        },
        sendNotificationEmail=send_notification,
        supportsAllDrives=True,
        fields='id, role, emailAddress, type',
    ).execute()


def delete_permission(service, file_id: str, permission_id: str) -> None:
    """Remove a sharing permission from a file."""
    service.permissions().delete(
        fileId=file_id,
        permissionId=permission_id,
        supportsAllDrives=True,
    ).execute()


def download_file_bytes(service, file_id: str, mime_type: str) -> tuple[bytes, str]:
    """
    Download file bytes.

    Google Workspace files are exported to a practical default format.
    Returns (bytes, suggested_extension_or_empty).
    """
    export_map = {
        GOOGLE_DOC_MIME: ('application/vnd.openxmlformats-officedocument.wordprocessingml.document', '.docx'),
        GOOGLE_SHEET_MIME: (
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            '.xlsx',
        ),
        GOOGLE_SLIDES_MIME: (
            'application/vnd.openxmlformats-officedocument.presentationml.presentation',
            '.pptx',
        ),
    }

    if mime_type in export_map:
        export_mime, extension = export_map[mime_type]
        content = service.files().export_media(
            fileId=file_id,
            mimeType=export_mime,
        ).execute()
        return content, extension

    request = service.files().get_media(fileId=file_id)
    buffer = io.BytesIO()
    downloader = MediaIoBaseDownload(buffer, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    return buffer.getvalue(), ''


def download_file_to_path(
    service,
    file_id: str,
    mime_type: str,
    local_path: str | Path,
    *,
    file_name: str | None = None,
) -> dict:
    """Download a Drive file to a local filesystem path."""
    content, export_ext = download_file_bytes(service, file_id, mime_type)
    path = Path(local_path).expanduser()
    if path.exists() and path.is_dir():
        base = file_name or file_id
        path = path / f"{base}{export_ext}"
    elif export_ext and path.suffix == '':
        path = path.with_suffix(export_ext)

    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(content)
    return {
        'local_path': str(path.resolve()),
        'bytes_written': len(content),
        'export_extension': export_ext or None,
    }


def list_shared_drives(service, max_results: int = 20) -> list[dict]:
    """List Shared Drives visible to the authenticated account."""
    results = service.drives().list(
        pageSize=max_results,
        fields='drives(id, name)',
    ).execute()
    return results.get('drives', [])


def find_child_item(
    service,
    name: str,
    parent_id: str | None = None,
    *,
    folder_only: bool = False,
) -> dict | None:
    """Find a non-trashed item by name under a parent (or My Drive root)."""
    query = (
        f"name = '{_escape_query(name)}' and trashed = false and "
        f"'{parent_id or 'root'}' in parents"
    )
    if folder_only:
        query += f" and mimeType = '{GOOGLE_FOLDER_MIME}'"
    results = service.files().list(
        q=query,
        supportsAllDrives=True,
        includeItemsFromAllDrives=True,
        fields="files(id, name, mimeType, parents, webViewLink)",
        pageSize=1,
    ).execute()
    items = results.get('files', [])
    return items[0] if items else None


def find_child_folder(
    service,
    folder_name: str,
    parent_id: str | None = None,
) -> dict | None:
    """Find a non-trashed folder by name under a parent (or My Drive root)."""
    return find_child_item(
        service,
        folder_name,
        parent_id,
        folder_only=True,
    )


def create_folder(
    service,
    folder_name: str,
    parent_id: str | None = None,
) -> dict:
    """Create a folder, optionally under a parent folder."""
    body: dict = {
        'name': folder_name,
        'mimeType': GOOGLE_FOLDER_MIME,
    }
    if parent_id:
        body['parents'] = [parent_id]
    return service.files().create(
        body=body,
        supportsAllDrives=True,
        fields='id, name, mimeType, parents, webViewLink',
    ).execute()


def get_parent_folder_name(service, file_info: dict) -> str | None:
    """Return the name of the file's first parent folder, if present."""
    parent_ids = file_info.get('parents') or []
    if not parent_ids:
        return None

    parent = service.files().get(
        fileId=parent_ids[0],
        fields='id, name',
        supportsAllDrives=True
    ).execute()
    return parent.get('name')


def _escape_query(value: str) -> str:
    """Escape single quotes for Drive API query strings."""
    return value.replace("\\", "\\\\").replace("'", "\\'")


def get_folder_id(service, folder_name: str) -> str | None:
    """Get the ID of a folder by name."""
    query = (
        f"name = '{_escape_query(folder_name)}' and "
        f"mimeType = '{GOOGLE_FOLDER_MIME}' and trashed = false"
    )
    results = service.files().list(
        q=query,
        spaces='drive',
        supportsAllDrives=True,
        includeItemsFromAllDrives=True,
        fields="files(id, name)"
    ).execute()
    folders = results.get('files', [])
    if not folders:
        return None
    return folders[0]['id']


def list_files_in_folder(service, folder_id: str) -> list[dict]:
    """List all files in a folder (non-recursive)."""
    query = f"'{folder_id}' in parents and trashed = false"
    results = service.files().list(
        q=query,
        supportsAllDrives=True,
        includeItemsFromAllDrives=True,
        fields="files(id, name, mimeType, webViewLink, modifiedTime)"
    ).execute()
    return results.get('files', [])


def list_files_recursive(service, folder_id: str, path: str = "") -> list[dict]:
    """Recursively list all files in a folder and subfolders."""
    all_files = []
    query = f"'{folder_id}' in parents and trashed = false"
    results = service.files().list(
        q=query,
        supportsAllDrives=True,
        includeItemsFromAllDrives=True,
        fields="files(id, name, mimeType, webViewLink, modifiedTime)"
    ).execute()
    items = results.get('files', [])

    for item in items:
        current_path = f"{path}/{item['name']}" if path else item['name']
        if item['mimeType'] == GOOGLE_FOLDER_MIME:
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
    """Search for files by name across My Drive and Shared Drives."""
    search_query = f"name contains '{_escape_query(query)}' and trashed = false"
    results = service.files().list(
        q=search_query,
        corpora='allDrives',
        includeItemsFromAllDrives=True,
        supportsAllDrives=True,
        fields="files(id, name, mimeType, webViewLink, modifiedTime)",
        pageSize=max_results
    ).execute()
    return results.get('files', [])


def find_file_by_name(
    service,
    file_name: str,
    *,
    fields: str = 'id, name, mimeType, parents',
    folder_only: bool = False,
) -> dict | None:
    """Find a non-trashed file by exact name. Returns the first match."""
    query = f"name = '{_escape_query(file_name)}' and trashed = false"
    if folder_only:
        query += f" and mimeType = '{GOOGLE_FOLDER_MIME}'"
    results = service.files().list(
        q=query,
        corpora='allDrives',
        includeItemsFromAllDrives=True,
        supportsAllDrives=True,
        fields=f"files({fields})",
        pageSize=1,
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
    file_info = get_file_metadata(service, file_id, fields='id, name, mimeType')
    if not file_info:
        return f"File not found: {file_id}"
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
