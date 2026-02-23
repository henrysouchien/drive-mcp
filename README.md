# drive-mcp

MCP server for Google Drive and OneDrive operations.

## Tools

| Tool | Description |
|------|-------------|
| `gdrive_list_folder` | List files in a Google Drive folder by name |
| `gdrive_list_folder_recursive` | Recursively list files in a folder and subfolders |
| `gdrive_search` | Search Google Drive by filename |
| `gdrive_read_file` | Read file contents from Google Drive |
| `gdrive_rename` | Rename a Google Drive file |
| `gdrive_move` | Move a Google Drive file to another folder |
| `onedrive_list_root` | List items in OneDrive root |
| `onedrive_list_folder` | List items in a OneDrive folder by path |
| `onedrive_search` | Search OneDrive by filename |
| `onedrive_read_file` | Read file contents from OneDrive |

## Setup

### Prerequisites
- Python 3.10+
- Google account with Drive API enabled
- Microsoft account with Graph API access

### Installation
```bash
git clone https://github.com/<your-user>/drive-mcp.git
cd drive-mcp
python3 -m venv venv
source venv/bin/activate
pip install -e .
```

### Authentication
- Google Drive credentials: place OAuth desktop client JSON at `drive_credentials.json`.
- OneDrive credentials: defaults in `src/onedrive.py` use a public device-flow app registration.
- Bootstrap auth once:
```bash
source venv/bin/activate
python -c "from src import google_drive; google_drive.authenticate()"
python -c "from src import onedrive; onedrive.authenticate()"
```

### Claude Code Configuration
Add to `~/.claude.json`:

```json
{
  "mcpServers": {
    "drive-mcp": {
      "type": "stdio",
      "command": "/path/to/drive-mcp/venv/bin/python",
      "args": ["/path/to/drive-mcp/run_server.py"]
    }
  }
}
```

## Development

```bash
source venv/bin/activate
python -c "from src.server import mcp; print([t.name for t in mcp._tool_manager._tools.values()])"
python run_server.py
```

## License
MIT
