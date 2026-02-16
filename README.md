# drive-mcp

MCP (Model Context Protocol) server that provides Claude Code access to Google Drive and OneDrive.

## Overview

This MCP exposes 8 tools for interacting with cloud storage:

| Tool | Description |
|------|-------------|
| `gdrive_list_folder` | List files in a Google Drive folder by name |
| `gdrive_list_folder_recursive` | Recursively list all files in a folder and subfolders |
| `gdrive_search` | Search Google Drive by filename |
| `gdrive_read_file` | Read file contents (Google Docs, Sheets, PDFs, text files) |
| `onedrive_list_root` | List items in OneDrive root |
| `onedrive_list_folder` | List items in a OneDrive folder by path |
| `onedrive_search` | Search OneDrive by filename |
| `onedrive_read_file` | Read file contents (PDFs, text, docx, xlsx, pptx) |

## Project Structure

```
drive-mcp/
├── README.md
├── pyproject.toml              # Python package config
├── run_server.py               # Entry point for MCP server
├── drive_credentials.json      # Google OAuth client credentials
├── token.pickle                # Cached Google auth token (auto-generated)
├── onedrive_token_cache.json   # Cached OneDrive auth token (auto-generated)
├── venv/                       # Python virtual environment
└── src/
    ├── __init__.py
    ├── server.py               # MCP server definition & tool handlers
    ├── google_drive.py         # Google Drive API client
    └── onedrive.py             # OneDrive/Microsoft Graph API client
```

## How It Works

### MCP Server (`src/server.py`)

Uses the `FastMCP` class from the MCP SDK. Tools are defined with the `@mcp.tool()` decorator:

```python
from mcp.server.fastmcp import FastMCP

mcp = FastMCP("drive-mcp")

@mcp.tool()
def gdrive_search(query: str, max_results: int = 20) -> str:
    """Search for files in Google Drive by name."""
    # Implementation...
```

The server runs via stdio (standard input/output) - Claude Code launches it as a subprocess and communicates via JSON-RPC.

### Google Drive (`src/google_drive.py`)

- **Auth**: OAuth 2.0 via `google-auth-oauthlib`
- **API**: Google Drive API v3 via `googleapiclient`
- **Credentials**: `drive_credentials.json` (OAuth client ID from Google Cloud Console)
- **Token cache**: `token.pickle` (auto-refreshes when expired)

Key functions:
- `authenticate()` - Returns authenticated Drive service
- `get_folder_id(service, folder_name)` - Find folder ID by name
- `list_files_in_folder(service, folder_id)` - List files (non-recursive)
- `list_files_recursive(service, folder_id, path)` - List all files recursively
- `search_files(service, query, max_results)` - Search by filename
- `read_file_by_name(service, file_name, max_chars)` - Read file contents

**Supported file types for reading:**
- Google Docs → exported as plain text
- Google Sheets → exported as CSV
- Google Slides → exported as plain text
- PDFs → text extracted via pypdf
- Text files (.txt, .md, .csv, .json, etc.) → direct read

### OneDrive (`src/onedrive.py`)

- **Auth**: Device flow via `msal` (Microsoft Authentication Library)
- **API**: Microsoft Graph API
- **Credentials**: Azure App Registration (Client ID + Tenant ID hardcoded)
- **Token cache**: `onedrive_token_cache.json`

Key functions:
- `authenticate()` - Returns access token (prompts device flow if needed)
- `list_root_items()` - List OneDrive root
- `list_folder_by_path(folder_path)` - List folder by path string
- `list_folder_by_id(folder_id)` - List folder by ID
- `list_files_recursive(folder_id, path)` - List all files recursively
- `search_files(query, max_results)` - Search by filename
- `read_file_by_path(file_path, max_chars)` - Read file contents

**Supported file types for reading:**
- PDFs → text extracted via pypdf
- Text files (.txt, .md, .csv, .json, .py, .js, etc.) → direct read
- Word documents (.docx) → text extracted via python-docx
- Excel spreadsheets (.xlsx) → exported as CSV-like format via openpyxl
- PowerPoint presentations (.pptx) → text extracted via python-pptx

**Note:** Some xlsx files created with specialized tools may have compatibility issues with openpyxl.

## Setup

### Prerequisites

- Python 3.10+
- Google Cloud project with Drive API enabled
- Azure App Registration with Microsoft Graph permissions

### Installation

```bash
cd /Users/henrychien/Documents/Jupyter/drive-mcp
python3 -m venv venv
source venv/bin/activate
pip install mcp google-api-python-client google-auth-oauthlib msal requests
```

### Authentication

**Google Drive** (first time or when token expires):
```bash
source venv/bin/activate
python -c "from src import google_drive; google_drive.authenticate()"
# Opens browser for OAuth consent
```

**OneDrive** (first time or when token expires):
```bash
source venv/bin/activate
python -c "from src import onedrive; onedrive.authenticate()"
# Displays URL and code for device flow
```

### Claude Code Configuration

Add to `~/.claude.json`:

```json
{
  "mcpServers": {
    "drive-mcp": {
      "type": "stdio",
      "command": "/Users/henrychien/Documents/Jupyter/drive-mcp/venv/bin/python",
      "args": ["/Users/henrychien/Documents/Jupyter/drive-mcp/run_server.py"]
    }
  }
}
```

Then restart Claude Code.

## Development

### Testing individual modules

```bash
source venv/bin/activate

# Test Google Drive
python -c "from src import google_drive; svc = google_drive.authenticate(); print(svc.files().list(pageSize=5).execute())"

# Test OneDrive
python -c "from src import onedrive; print(onedrive.list_root_items())"

# Test MCP server loads correctly
python -c "from src.server import mcp; print([t.name for t in mcp._tool_manager._tools.values()])"
```

### Adding new tools

1. Add the API function to `google_drive.py` or `onedrive.py`
2. Add the tool wrapper in `server.py`:

```python
@mcp.tool()
def my_new_tool(param: str) -> str:
    """Description shown to Claude."""
    try:
        # Call your API function
        result = google_drive.my_new_function(param)
        return format_result(result)
    except Exception as e:
        return f"Error: {str(e)}"
```

3. Restart Claude Code to pick up changes

### Debugging

Run the server directly to see errors:
```bash
source venv/bin/activate
python run_server.py
```

The server expects JSON-RPC on stdin, so it will hang waiting for input - but any import/startup errors will show immediately.

## Credentials

### Google Drive

`drive_credentials.json` is an OAuth 2.0 Client ID from Google Cloud Console:
1. Go to https://console.cloud.google.com/
2. Create/select project → APIs & Services → Credentials
3. Create OAuth 2.0 Client ID (Desktop app)
4. Download JSON and save as `drive_credentials.json`

### OneDrive

Uses Azure App Registration with these settings:
- **Client ID**: `c7b60d92-d23f-474b-9708-fb8890be59e3`
- **Tenant ID**: `c57d3288-4e87-42e2-bd6c-fb6f632680c3`
- **Permissions**: `Files.Read.All`, `User.Read`
- **Auth flow**: Device code flow (no client secret needed)

To use a different Azure app, update the constants in `src/onedrive.py`.
