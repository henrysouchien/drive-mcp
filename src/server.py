"""
MCP Server for Google Drive and OneDrive.
Provides tools for listing and searching files in both cloud storage services.
"""

import base64
import json
from dataclasses import dataclass
from typing import Any

from mcp.server.fastmcp import FastMCP

from . import google_drive
from . import onedrive

# Create the MCP server
mcp = FastMCP("drive-mcp")


@dataclass
class ToolError:
    error_class: str
    message: str
    names_correction: dict[str, Any] | None = None
    suggested_tool_calls: list[dict[str, Any]] | None = None

    def to_envelope(self) -> dict[str, Any]:
        return {
            "status": "error",
            "error_class": self.error_class,
            "message": self.message,
            "names_correction": self.names_correction or {},
            "suggested_tool_calls": self.suggested_tool_calls or [],
        }


def _exception_envelope(exc: Exception) -> dict[str, Any]:
    return ToolError(
        error_class=type(exc).__name__,
        message=str(exc),
        names_correction={},
        suggested_tool_calls=_suggested_discovery_calls(str(exc)),
    ).to_envelope()


def _not_found_envelope(kind: str, value: str, discovery_tool: str) -> dict[str, Any]:
    return ToolError(
        error_class=f"{kind.title()}NotFound",
        message=f"{kind} not found: {value}",
        names_correction={kind: f"Run {discovery_tool} and use an exact returned name."},
        suggested_tool_calls=[{"name": discovery_tool, "args": {"query": value}}],
    ).to_envelope()


def _suggested_discovery_calls(message: str) -> list[dict[str, Any]]:
    lowered = message.lower()
    if "onedrive" in lowered:
        return [{"name": "onedrive_start_reauth", "args": {}}]
    if "file" in lowered:
        return [{"name": "gdrive_search", "args": {"query": "file name", "max_results": 20}}]
    if "folder" in lowered:
        return [{"name": "gdrive_search", "args": {"query": "folder name", "max_results": 20}}]
    return []


def _make_restore_token(payload: dict[str, Any]) -> str:
    token_payload = {"version": 1, **payload}
    raw = json.dumps(token_payload, separators=(",", ":"), sort_keys=True).encode("utf-8")
    return base64.urlsafe_b64encode(raw).decode("ascii")


def _parse_restore_token(restore_token: str) -> dict[str, Any]:
    try:
        raw = base64.urlsafe_b64decode(restore_token.encode("ascii"))
        payload = json.loads(raw.decode("utf-8"))
    except Exception as exc:
        raise ValueError("Invalid restore_token") from exc
    if not isinstance(payload, dict) or payload.get("version") != 1:
        raise ValueError("Invalid restore_token")
    return payload


# =============================================================================
# GOOGLE DRIVE TOOLS
# =============================================================================

@mcp.tool()
def gdrive_list_folder(folder_name: str) -> str | dict:
    """
    List files in a Google Drive folder by name.

    Args:
        folder_name: Name of the folder to list (e.g., "Stock Investor Accelerator")
    """
    try:
        service = google_drive.authenticate()
        folder_id = google_drive.get_folder_id(service, folder_name)

        if not folder_id:
            return _not_found_envelope("folder", folder_name, "gdrive_search")

        files = google_drive.list_files_in_folder(service, folder_id)

        if not files:
            return f"Folder '{folder_name}' is empty."

        result = f"Files in '{folder_name}':\n\n"
        for f in files:
            icon = "📁" if f['mimeType'] == 'application/vnd.google-apps.folder' else "📄"
            result += f"{icon} {f['name']}\n"
            if f.get('modifiedTime'):
                result += f"   Modified: {f['modifiedTime']}\n"
            if f.get('webViewLink'):
                result += f"   Link: {f['webViewLink']}\n"

        return result
    except Exception as e:
        return _exception_envelope(e)


@mcp.tool()
def gdrive_list_folder_recursive(folder_name: str) -> str | dict:
    """
    Recursively list all files in a Google Drive folder and its subfolders.

    Args:
        folder_name: Name of the folder to list (e.g., "Stock Investor Accelerator")
    """
    try:
        service = google_drive.authenticate()
        folder_id = google_drive.get_folder_id(service, folder_name)

        if not folder_id:
            return _not_found_envelope("folder", folder_name, "gdrive_search")

        files = google_drive.list_files_recursive(service, folder_id)

        if not files:
            return f"Folder '{folder_name}' is empty."

        result = f"All files in '{folder_name}' ({len(files)} files):\n\n"
        for f in files:
            result += f"📄 {f['path']}\n"
            if f.get('modifiedTime'):
                result += f"   Modified: {f['modifiedTime']}\n"
            if f.get('webViewLink'):
                result += f"   Link: {f['webViewLink']}\n"

        return result
    except Exception as e:
        return _exception_envelope(e)


@mcp.tool()
def gdrive_search(query: str, max_results: int = 20) -> str | dict:
    """
    Search for files in Google Drive by name.

    Args:
        query: Search term to find in file names
        max_results: Maximum number of results to return (default: 20)
    """
    try:
        service = google_drive.authenticate()
        files = google_drive.search_files(service, query, max_results)

        if not files:
            return f"No files found matching '{query}'."

        result = f"Search results for '{query}' ({len(files)} files):\n\n"
        for f in files:
            icon = "📁" if f['mimeType'] == 'application/vnd.google-apps.folder' else "📄"
            result += f"{icon} {f['name']}\n"
            result += f"   ID: {f['id']}\n"
            if f.get('modifiedTime'):
                result += f"   Modified: {f['modifiedTime']}\n"
            if f.get('webViewLink'):
                result += f"   Link: {f['webViewLink']}\n"

        return result
    except Exception as e:
        return _exception_envelope(e)


@mcp.tool()
def gdrive_read_file(file_name: str, max_chars: int = 100000) -> str | dict:
    """
    Read the contents of a file from Google Drive.

    Supports:
    - Google Docs (exported as plain text)
    - Google Sheets (exported as CSV)
    - PDFs (text extracted)
    - Text files (.txt, .md, .csv, .json, etc.)

    Args:
        file_name: Name of the file to read (e.g., "My Document" or "report.pdf")
        max_chars: Maximum characters to return (default: 100000)
    """
    try:
        service = google_drive.authenticate()
        return google_drive.read_file_by_name(service, file_name, max_chars)
    except Exception as e:
        return _exception_envelope(e)


@mcp.tool()
def gdrive_rename(file_name: str, new_name: str, dry_run: bool = False) -> str | dict:
    """
    Rename a file in Google Drive.

    Discovery: run `gdrive_list_folder` first to obtain `file_name` values.

    Use this for: renaming files in place.
    Not for: moving a file to a different folder — see `gdrive_move`.

    Use dry_run=True to preview the change without committing.

    Args:
        file_name: Current name of the file to rename
        new_name: New name for the file
        dry_run: Preview the rename without committing it
    """
    try:
        service = google_drive.authenticate()
        file_info = google_drive.find_file_by_name(service, file_name)
        if not file_info:
            return _not_found_envelope("file", file_name, "gdrive_search")
        if dry_run:
            current_parent = google_drive.get_parent_folder_name(service, file_info)
            return {
                "dry_run": True,
                "would_rename": file_name,
                "to": new_name,
                "current_id": file_info['id'],
                "current_parent": current_parent,
            }
        result = google_drive.rename_file(service, file_info['id'], new_name)
        restore_token = _make_restore_token(
            {
                "operation": "gdrive_rename",
                "file_id": file_info["id"],
                "restore_name": file_info.get("name") or file_name,
                "current_name": result["name"],
            }
        )
        return {
            "status": "ok",
            "message": f"Renamed '{file_name}' to '{result['name']}'",
            "file_id": result["id"],
            "name": result["name"],
            "restore_token": restore_token,
            "undo_tool": "gdrive_undo",
        }
    except Exception as e:
        return _exception_envelope(e)


@mcp.tool()
def gdrive_move(file_name: str, destination_folder: str, dry_run: bool = False) -> str | dict:
    """
    Move a file to a different folder in Google Drive.

    Discovery: run `gdrive_list_folder` first to obtain `file_name` and `destination_folder` values.

    Use this for: moving files between folders.
    Not for: renaming a file in place — see `gdrive_rename`.

    Use dry_run=True to preview the change without committing.

    Args:
        file_name: Name of the file to move
        destination_folder: Name of the destination folder
        dry_run: Preview the move without committing it
    """
    try:
        service = google_drive.authenticate()
        file_info = google_drive.find_file_by_name(service, file_name)
        if not file_info:
            return _not_found_envelope("file", file_name, "gdrive_search")
        if dry_run:
            current_parent = google_drive.get_parent_folder_name(service, file_info)
            return {
                "dry_run": True,
                "would_move": file_name,
                "from": current_parent,
                "to": destination_folder,
                "current_id": file_info['id'],
            }
        folder_id = google_drive.get_folder_id(service, destination_folder)
        if not folder_id:
            return _not_found_envelope("folder", destination_folder, "gdrive_search")
        result = google_drive.move_file(service, file_info['id'], folder_id)
        restore_token = _make_restore_token(
            {
                "operation": "gdrive_move",
                "file_id": file_info["id"],
                "restore_parent_ids": file_info.get("parents") or [],
                "destination_parent_id": folder_id,
                "destination_folder": destination_folder,
                "file_name": result["name"],
            }
        )
        return {
            "status": "ok",
            "message": f"Moved '{result['name']}' to '{destination_folder}'",
            "file_id": result["id"],
            "name": result["name"],
            "parents": result.get("parents", []),
            "restore_token": restore_token,
            "undo_tool": "gdrive_undo",
        }
    except Exception as e:
        return _exception_envelope(e)


@mcp.tool()
def gdrive_undo(restore_token: str) -> dict:
    """
    Undo a prior gdrive_rename or gdrive_move using its restore_token.

    Args:
        restore_token: Token returned by gdrive_rename or gdrive_move after a committed change
    """
    try:
        payload = _parse_restore_token(restore_token)
        service = google_drive.authenticate()
        operation = payload.get("operation")

        if operation == "gdrive_rename":
            result = google_drive.rename_file(
                service,
                str(payload["file_id"]),
                str(payload["restore_name"]),
            )
            return {
                "status": "ok",
                "operation": operation,
                "message": f"Restored file name to '{result['name']}'",
                "file_id": result["id"],
                "name": result["name"],
            }

        if operation == "gdrive_move":
            result = google_drive.set_file_parents(
                service,
                str(payload["file_id"]),
                list(payload.get("restore_parent_ids") or []),
            )
            return {
                "status": "ok",
                "operation": operation,
                "message": f"Restored '{result['name']}' to its previous parent folder(s)",
                "file_id": result["id"],
                "name": result["name"],
                "parents": result.get("parents", []),
            }

        raise ValueError("restore_token operation is not supported by gdrive_undo")
    except Exception as e:
        return _exception_envelope(e)


# =============================================================================
# ONEDRIVE TOOLS
# =============================================================================

@mcp.tool()
def onedrive_list_root() -> str | dict:
    """
    List items in the OneDrive root folder.
    """
    try:
        items = onedrive.list_root_items()

        if not items:
            return "OneDrive root is empty."

        result = "OneDrive root contents:\n\n"
        for item in items:
            icon = "📁" if item.get("folder") else "📄"
            result += f"{icon} {item['name']}\n"
            if item.get('lastModifiedDateTime'):
                result += f"   Modified: {item['lastModifiedDateTime']}\n"
            if item.get('webUrl'):
                result += f"   Link: {item['webUrl']}\n"

        return result
    except Exception as e:
        return _exception_envelope(e)


@mcp.tool()
def onedrive_list_folder(folder_path: str) -> str | dict:
    """
    List items in a OneDrive folder by path.

    Args:
        folder_path: Path to the folder (e.g., "Documents/Projects" or "Stock Investor Accelerator")
    """
    try:
        items = onedrive.list_folder_by_path(folder_path)

        if not items:
            return f"Folder '{folder_path}' is empty or not found."

        result = f"Contents of '{folder_path}':\n\n"
        for item in items:
            icon = "📁" if item.get("folder") else "📄"
            result += f"{icon} {item['name']}\n"
            if item.get('lastModifiedDateTime'):
                result += f"   Modified: {item['lastModifiedDateTime']}\n"
            if item.get('webUrl'):
                result += f"   Link: {item['webUrl']}\n"

        return result
    except Exception as e:
        return _exception_envelope(e)


@mcp.tool()
def onedrive_search(query: str, max_results: int = 20) -> str | dict:
    """
    Search for files in OneDrive by name.

    Args:
        query: Search term to find in file names
        max_results: Maximum number of results to return (default: 20)
    """
    try:
        files = onedrive.search_files(query, max_results)

        if not files:
            return f"No files found matching '{query}'."

        result = f"Search results for '{query}' ({len(files)} files):\n\n"
        for f in files:
            icon = "📁" if f.get("folder") else "📄"
            result += f"{icon} {f['name']}\n"
            if f.get('lastModifiedDateTime'):
                result += f"   Modified: {f['lastModifiedDateTime']}\n"
            if f.get('webUrl'):
                result += f"   Link: {f['webUrl']}\n"

        return result
    except Exception as e:
        return _exception_envelope(e)


@mcp.tool()
def onedrive_read_file(file_path: str, max_chars: int = 100000) -> str | dict:
    """
    Read the contents of a file from OneDrive.

    Supports:
    - Text files (.txt, .md, .csv, .json, etc.)
    - PDFs (text extracted)
    - Word documents (.docx)
    - Excel spreadsheets (.xlsx) - exported as CSV-like format
    - PowerPoint presentations (.pptx)

    Args:
        file_path: Path to the file (e.g., "Documents/report.pdf" or "Stock Investor Accelerator/notes.txt")
        max_chars: Maximum characters to return (default: 100000)
    """
    try:
        return onedrive.read_file_by_path(file_path, max_chars)
    except Exception as e:
        return _exception_envelope(e)


@mcp.tool()
def onedrive_start_reauth() -> str | dict:
    """
    Start OneDrive re-authentication. Returns URL and code for user.
    """
    try:
        result = onedrive.start_reauth()
        return (
            "OneDrive re-authentication started.\n\n"
            f"1. Visit: {result['verification_uri']}\n"
            f"2. Enter code: {result['user_code']}\n"
            f"3. Code expires in: {result.get('expires_in')} seconds\n\n"
            "After completing the browser step, call onedrive_complete_reauth()."
        )
    except Exception as e:
        return _exception_envelope(e)


@mcp.tool()
def onedrive_complete_reauth() -> str | dict:
    """
    Check if OneDrive re-authentication completed. Call after user visits URL.
    """
    try:
        result = onedrive.poll_reauth()
        status = result.get("status")

        if status == "success":
            account = result.get("account")
            if account:
                return f"OneDrive re-authentication successful for {account}."
            return "OneDrive re-authentication successful."

        if status == "pending":
            description = result.get("error_description", "Authorization is still pending.")
            return (
                "OneDrive re-authentication is still pending.\n"
                f"{description}\n"
                "Call onedrive_complete_reauth() again in a few seconds."
            )

        description = result.get("error_description", result.get("error", "Unknown error"))
        return ToolError(
            error_class="OneDriveReauthFailed",
            message=f"OneDrive re-authentication failed: {description}",
            names_correction={},
            suggested_tool_calls=[{"name": "onedrive_start_reauth", "args": {}}],
        ).to_envelope()
    except Exception as e:
        return _exception_envelope(e)


# =============================================================================
# MAIN ENTRY POINT
# =============================================================================

def main():
    mcp.run()


if __name__ == "__main__":
    main()
