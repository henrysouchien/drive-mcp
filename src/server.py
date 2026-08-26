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
        names_correction={kind: f"Run {discovery_tool} and use an exact returned name or ID."},
        suggested_tool_calls=[{"name": discovery_tool, "args": {"query": value}}],
    ).to_envelope()


def _missing_ref_envelope(kind: str) -> dict[str, Any]:
    return ToolError(
        error_class="MissingReference",
        message=f"Provide {kind}_name or {kind}_id",
        names_correction={
            kind: f"Run gdrive_search and pass an exact {kind}_name or {kind}_id."
        },
        suggested_tool_calls=[{"name": "gdrive_search", "args": {"query": kind, "max_results": 20}}],
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


def _resolve_file_or_error(
    service,
    *,
    file_name: str | None = None,
    file_id: str | None = None,
) -> tuple[dict[str, Any] | None, dict[str, Any] | None]:
    if not file_name and not file_id:
        return None, _missing_ref_envelope("file")
    try:
        info = google_drive.resolve_file(service, file_name=file_name, file_id=file_id)
    except ValueError as exc:
        return None, _exception_envelope(exc)
    if not info:
        return None, _not_found_envelope("file", file_id or file_name or "", "gdrive_search")
    return info, None


def _resolve_folder_or_error(
    service,
    *,
    folder_name: str | None = None,
    folder_id: str | None = None,
    required: bool = True,
) -> tuple[dict[str, Any] | None, dict[str, Any] | None]:
    if not folder_name and not folder_id:
        if required:
            return None, _missing_ref_envelope("folder")
        return None, None
    try:
        info = google_drive.resolve_folder(
            service,
            folder_name=folder_name,
            folder_id=folder_id,
        )
    except ValueError as exc:
        return None, _exception_envelope(exc)
    if not info:
        return None, _not_found_envelope(
            "folder",
            folder_id or folder_name or "",
            "gdrive_search",
        )
    return info, None


def _format_file_lines(files: list[dict[str, Any]], *, use_path: bool = False) -> str:
    lines: list[str] = []
    for f in files:
        icon = "📁" if f.get("mimeType") == google_drive.GOOGLE_FOLDER_MIME else "📄"
        label = f.get("path") if use_path else f.get("name")
        lines.append(f"{icon} {label}")
        if f.get("id"):
            lines.append(f"   ID: {f['id']}")
        if f.get("modifiedTime"):
            lines.append(f"   Modified: {f['modifiedTime']}")
        if f.get("webViewLink"):
            lines.append(f"   Link: {f['webViewLink']}")
    return "\n".join(lines)


# =============================================================================
# GOOGLE DRIVE TOOLS
# =============================================================================

@mcp.tool()
def gdrive_list_folder(
    folder_name: str | None = None,
    folder_id: str | None = None,
) -> str | dict:
    """
    List files in a Google Drive folder by name or folder_id.

    Args:
        folder_name: Name of the folder to list (e.g., "Stock Investor Accelerator")
        folder_id: Optional stable folder ID from gdrive_search (preferred when known)

    Discovery: use gdrive_search first when the exact folder_name/folder_id is unknown.

    Sibling tools: use gdrive_list_folder_recursive to include subfolders and
    gdrive_read_file to read a discovered file.
    """
    try:
        service = google_drive.authenticate()
        folder, error = _resolve_folder_or_error(
            service,
            folder_name=folder_name,
            folder_id=folder_id,
        )
        if error:
            return error

        files = google_drive.list_files_in_folder(service, folder["id"])
        label = folder["name"]
        if not files:
            return f"Folder '{label}' is empty."

        return f"Files in '{label}':\n\n{_format_file_lines(files)}"
    except Exception as e:
        return _exception_envelope(e)


@mcp.tool()
def gdrive_list_folder_recursive(
    folder_name: str | None = None,
    folder_id: str | None = None,
) -> str | dict:
    """
    Recursively list all files in a Google Drive folder and its subfolders.

    Args:
        folder_name: Name of the folder to list (e.g., "Stock Investor Accelerator")
        folder_id: Optional stable folder ID from gdrive_search (preferred when known)

    Discovery: use gdrive_search first when the exact folder_name/folder_id is unknown.

    Sibling tools: use gdrive_list_folder for a shallow listing and
    gdrive_read_file to read a discovered file.
    """
    try:
        service = google_drive.authenticate()
        folder, error = _resolve_folder_or_error(
            service,
            folder_name=folder_name,
            folder_id=folder_id,
        )
        if error:
            return error

        files = google_drive.list_files_recursive(service, folder["id"])
        label = folder["name"]
        if not files:
            return f"Folder '{label}' is empty."

        return (
            f"All files in '{label}' ({len(files)} files):\n\n"
            f"{_format_file_lines(files, use_path=True)}"
        )
    except Exception as e:
        return _exception_envelope(e)


@mcp.tool()
def gdrive_search(query: str, max_results: int = 20) -> str | dict:
    """
    Search for files in Google Drive by name (includes Shared Drives).

    Args:
        query: Search term to find in file names
        max_results: Maximum number of results to return (default: 20)
    """
    try:
        service = google_drive.authenticate()
        files = google_drive.search_files(service, query, max_results)

        if not files:
            return f"No files found matching '{query}'."

        return (
            f"Search results for '{query}' ({len(files)} files):\n\n"
            f"{_format_file_lines(files)}"
        )
    except Exception as e:
        return _exception_envelope(e)


@mcp.tool()
def gdrive_list_shared_drives(max_results: int = 20) -> str | dict:
    """
    List Shared Drives visible to the authenticated Google account.

    Sibling tools: use gdrive_search / gdrive_list_folder with folder_id values
    discovered inside a Shared Drive.
    """
    try:
        service = google_drive.authenticate()
        drives = google_drive.list_shared_drives(service, max_results)
        if not drives:
            return "No Shared Drives found."

        lines = [f"Shared Drives ({len(drives)}):\n"]
        for drive in drives:
            lines.append(f"📁 {drive['name']}")
            lines.append(f"   ID: {drive['id']}")
        return "\n".join(lines)
    except Exception as e:
        return _exception_envelope(e)


@mcp.tool()
def gdrive_read_file(
    file_name: str | None = None,
    max_chars: int = 100000,
    file_id: str | None = None,
) -> str | dict:
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
        file_id: Optional stable file ID from gdrive_search (preferred when known)

    Discovery: use gdrive_search, gdrive_list_folder, or
    gdrive_list_folder_recursive first to find the exact file_name or file_id.
    """
    try:
        service = google_drive.authenticate()
        file_info, error = _resolve_file_or_error(
            service,
            file_name=file_name,
            file_id=file_id,
        )
        if error:
            return error
        return google_drive.read_file_content(
            service,
            file_info["id"],
            file_info["mimeType"],
            max_chars,
        )
    except Exception as e:
        return _exception_envelope(e)


@mcp.tool()
def gdrive_download_file(
    local_path: str,
    file_name: str | None = None,
    file_id: str | None = None,
    dry_run: bool = False,
) -> str | dict:
    """
    Download a Google Drive file to a local filesystem path.

    Google Docs/Sheets/Slides are exported to Office formats (.docx/.xlsx/.pptx).
    Other files are downloaded in their native binary form.

    Discovery: use gdrive_search or gdrive_list_folder first to obtain
    `file_name` / `file_id` values.

    Use this for: saving a Drive file locally for offline/binary use.
    Not for: reading text content in-chat — see `gdrive_read_file`.

    Args:
        local_path: Destination file path, or directory to download into
        file_name: Name of the file to download
        file_id: Optional stable file ID (preferred when known)
        dry_run: Preview the download without writing bytes
    """
    try:
        service = google_drive.authenticate()
        file_info, error = _resolve_file_or_error(
            service,
            file_name=file_name,
            file_id=file_id,
        )
        if error:
            return error

        if dry_run:
            return {
                "dry_run": True,
                "would_download": file_info["name"],
                "file_id": file_info["id"],
                "mime_type": file_info["mimeType"],
                "local_path": local_path,
            }

        result = google_drive.download_file_to_path(
            service,
            file_info["id"],
            file_info["mimeType"],
            local_path,
            file_name=file_info["name"],
        )
        return {
            "status": "ok",
            "message": f"Downloaded '{file_info['name']}' to '{result['local_path']}'",
            "file_id": file_info["id"],
            "name": file_info["name"],
            **result,
        }
    except Exception as e:
        return _exception_envelope(e)


@mcp.tool()
def gdrive_rename(
    file_name: str | None = None,
    new_name: str | None = None,
    dry_run: bool = False,
    file_id: str | None = None,
) -> str | dict:
    """
    Rename a file in Google Drive.

    Discovery: run `gdrive_list_folder` or `gdrive_search` first to obtain
    `file_name` / `file_id` values.

    Use this for: renaming files in place.
    Not for: moving a file to a different folder — see `gdrive_move`.

    Use dry_run=True to preview the change without committing.

    Args:
        file_name: Current name of the file to rename
        new_name: New name for the file
        dry_run: Preview the rename without committing it
        file_id: Optional stable file ID (preferred when known)
    """
    try:
        if not new_name:
            return ToolError(
                error_class="MissingArgument",
                message="new_name is required",
                names_correction={"new_name": "Provide the destination file name."},
                suggested_tool_calls=[],
            ).to_envelope()
        service = google_drive.authenticate()
        file_info, error = _resolve_file_or_error(
            service,
            file_name=file_name,
            file_id=file_id,
        )
        if error:
            return error
        current_name = file_info.get("name") or file_name or file_id
        if dry_run:
            current_parent = google_drive.get_parent_folder_name(service, file_info)
            return {
                "dry_run": True,
                "would_rename": current_name,
                "to": new_name,
                "current_id": file_info["id"],
                "current_parent": current_parent,
            }
        result = google_drive.rename_file(service, file_info["id"], new_name)
        restore_token = _make_restore_token(
            {
                "operation": "gdrive_rename",
                "file_id": file_info["id"],
                "restore_name": current_name,
                "current_name": result["name"],
            }
        )
        return {
            "status": "ok",
            "message": f"Renamed '{current_name}' to '{result['name']}'",
            "file_id": result["id"],
            "name": result["name"],
            "restore_token": restore_token,
            "undo_tool": "gdrive_undo",
        }
    except Exception as e:
        return _exception_envelope(e)


@mcp.tool()
def gdrive_move(
    file_name: str | None = None,
    destination_folder: str | None = None,
    dry_run: bool = False,
    file_id: str | None = None,
    destination_folder_id: str | None = None,
) -> str | dict:
    """
    Move a file to a different folder in Google Drive.

    Discovery: run `gdrive_list_folder` or `gdrive_search` first to obtain
    `file_name` / `file_id` and destination folder values.

    Use this for: moving files between folders.
    Not for: renaming a file in place — see `gdrive_rename`.

    Use dry_run=True to preview the change without committing.

    Args:
        file_name: Name of the file to move
        destination_folder: Name of the destination folder
        dry_run: Preview the move without committing it
        file_id: Optional stable file ID (preferred when known)
        destination_folder_id: Optional stable destination folder ID
    """
    try:
        service = google_drive.authenticate()
        file_info, error = _resolve_file_or_error(
            service,
            file_name=file_name,
            file_id=file_id,
        )
        if error:
            return error

        dest_label = destination_folder or destination_folder_id or ""
        if dry_run:
            current_parent = google_drive.get_parent_folder_name(service, file_info)
            return {
                "dry_run": True,
                "would_move": file_info.get("name") or file_name or file_id,
                "from": current_parent,
                "to": dest_label,
                "current_id": file_info["id"],
            }

        folder, folder_error = _resolve_folder_or_error(
            service,
            folder_name=destination_folder,
            folder_id=destination_folder_id,
        )
        if folder_error:
            return folder_error

        result = google_drive.move_file(service, file_info["id"], folder["id"])
        restore_token = _make_restore_token(
            {
                "operation": "gdrive_move",
                "file_id": file_info["id"],
                "restore_parent_ids": file_info.get("parents") or [],
                "destination_parent_id": folder["id"],
                "destination_folder": folder["name"],
                "file_name": result["name"],
            }
        )
        return {
            "status": "ok",
            "message": f"Moved '{result['name']}' to '{folder['name']}'",
            "file_id": result["id"],
            "name": result["name"],
            "parents": result.get("parents", []),
            "restore_token": restore_token,
            "undo_tool": "gdrive_undo",
        }
    except Exception as e:
        return _exception_envelope(e)


@mcp.tool()
def gdrive_create_folder(
    folder_name: str,
    parent_folder: str | None = None,
    parent_folder_id: str | None = None,
    exist_ok: bool = False,
    dry_run: bool = False,
) -> str | dict:
    """
    Create a folder in Google Drive.

    Discovery: run `gdrive_search` or `gdrive_list_folder` first to obtain
    parent folder name/ID values when nesting under an existing folder.

    Use this for: creating a new folder (optionally under a parent).
    Not for: creating a Google Doc — see `gdrive_create_doc`.
    Not for: renaming an existing folder — see `gdrive_rename`.
    Not for: moving an existing folder — see `gdrive_move`.

    Use dry_run=True to preview the change without committing.
    Use exist_ok=True to return an existing same-named child folder instead of erroring.

    Args:
        folder_name: Name of the folder to create
        parent_folder: Optional parent folder name (defaults to My Drive root)
        parent_folder_id: Optional stable parent folder ID
        exist_ok: If a matching folder already exists under the parent, return it
        dry_run: Preview the create without committing it
    """
    try:
        service = google_drive.authenticate()
        parent, parent_error = _resolve_folder_or_error(
            service,
            folder_name=parent_folder,
            folder_id=parent_folder_id,
            required=False,
        )
        if parent_error:
            return parent_error

        parent_id = parent["id"] if parent else None
        parent_label = parent["name"] if parent else "My Drive root"
        existing = google_drive.find_child_folder(service, folder_name, parent_id)

        if dry_run:
            return {
                "dry_run": True,
                "would_create": folder_name,
                "parent_folder": parent_label,
                "already_exists": bool(existing),
                "existing_id": existing["id"] if existing else None,
                "exist_ok": exist_ok,
            }

        if existing:
            if exist_ok:
                return {
                    "status": "ok",
                    "created": False,
                    "message": (
                        f"Folder '{folder_name}' already exists under '{parent_label}'"
                    ),
                    "file_id": existing["id"],
                    "name": existing["name"],
                    "parents": existing.get("parents", []),
                    "webViewLink": existing.get("webViewLink", ""),
                }
            return ToolError(
                error_class="FolderAlreadyExists",
                message=(
                    f"folder already exists under '{parent_label}': {folder_name}"
                ),
                names_correction={
                    "folder_name": (
                        "Choose a new name, or pass exist_ok=True to reuse "
                        "the existing folder."
                    )
                },
                suggested_tool_calls=[
                    {
                        "name": "gdrive_list_folder",
                        "args": {
                            "folder_name": parent_folder,
                            "folder_id": parent_folder_id or parent_id,
                        },
                    }
                ],
            ).to_envelope()

        result = google_drive.create_folder(service, folder_name, parent_id)
        restore_token = _make_restore_token(
            {
                "operation": "gdrive_create_folder",
                "file_id": result["id"],
                "folder_name": result["name"],
                "parent_folder": parent_label,
            }
        )
        return {
            "status": "ok",
            "created": True,
            "message": f"Created folder '{result['name']}' under '{parent_label}'",
            "file_id": result["id"],
            "name": result["name"],
            "parents": result.get("parents", []),
            "webViewLink": result.get("webViewLink", ""),
            "restore_token": restore_token,
            "undo_tool": "gdrive_undo",
        }
    except Exception as e:
        return _exception_envelope(e)


@mcp.tool()
def gdrive_write_file(
    file_name: str,
    content: str,
    parent_folder: str | None = None,
    parent_folder_id: str | None = None,
    mime_type: str | None = None,
    file_id: str | None = None,
    overwrite: bool = False,
    dry_run: bool = False,
) -> str | dict:
    """
    Create or overwrite a text/binary-capable file in Google Drive.

    Discovery: run `gdrive_search` or `gdrive_list_folder` first to obtain
    parent folder and optional existing `file_id` values.

    Use this for: writing markdown/text/json/csv files into Drive.
    Not for: creating a Google Doc — see `gdrive_create_doc`.
    Not for: spreadsheet edits — use gsheets-mcp.

    Use dry_run=True to preview the change without committing.
    Overwrites only when file_id is provided or overwrite=True finds a same-named
    file under the parent. Undo trashes newly created files; content overwrites
    are not undoable via restore_token.

    Args:
        file_name: Destination file name (e.g., "notes.md")
        content: File contents to write
        parent_folder: Optional parent folder name (defaults to My Drive root)
        parent_folder_id: Optional stable parent folder ID
        mime_type: Optional MIME type (guessed from file_name when omitted)
        file_id: Optional existing file ID to overwrite
        overwrite: If True, overwrite an existing same-named file under the parent
        dry_run: Preview the write without committing it
    """
    try:
        service = google_drive.authenticate()
        parent, parent_error = _resolve_folder_or_error(
            service,
            folder_name=parent_folder,
            folder_id=parent_folder_id,
            required=False,
        )
        if parent_error:
            return parent_error

        parent_id = parent["id"] if parent else None
        parent_label = parent["name"] if parent else "My Drive root"
        target_id = file_id

        if target_id:
            existing = google_drive.get_file_metadata(service, target_id)
            if not existing:
                return _not_found_envelope("file", target_id, "gdrive_search")
        elif overwrite:
            existing = google_drive.find_child_item(service, file_name, parent_id)
            if existing:
                target_id = existing["id"]

        resolved_mime = google_drive.guess_mime_type(file_name, mime_type)
        if dry_run:
            return {
                "dry_run": True,
                "would_write": file_name,
                "parent_folder": parent_label,
                "mime_type": resolved_mime,
                "bytes": len(content.encode("utf-8")),
                "overwrite": bool(target_id),
                "file_id": target_id,
            }

        result = google_drive.write_file(
            service,
            file_name,
            content,
            parent_id=parent_id,
            mime_type=resolved_mime,
            file_id=target_id,
        )
        created = target_id is None
        response: dict[str, Any] = {
            "status": "ok",
            "created": created,
            "message": (
                f"{'Created' if created else 'Updated'} '{result['name']}' "
                f"under '{parent_label}'"
            ),
            "file_id": result["id"],
            "name": result["name"],
            "mimeType": result.get("mimeType", resolved_mime),
            "parents": result.get("parents", []),
            "webViewLink": result.get("webViewLink", ""),
        }
        if created:
            response["restore_token"] = _make_restore_token(
                {
                    "operation": "gdrive_write_file",
                    "file_id": result["id"],
                    "file_name": result["name"],
                }
            )
            response["undo_tool"] = "gdrive_undo"
        return response
    except Exception as e:
        return _exception_envelope(e)


@mcp.tool()
def gdrive_create_doc(
    title: str,
    content: str | None = None,
    parent_folder: str | None = None,
    parent_folder_id: str | None = None,
    dry_run: bool = False,
) -> str | dict:
    """
    Create a Google Doc in Drive, optionally seeded with plain-text content.

    Discovery: run `gdrive_search` or `gdrive_list_folder` first to obtain
    parent folder name/ID values.

    Use this for: creating a Google Doc artifact.
    Not for: creating a folder — see `gdrive_create_folder`.
    Not for: writing a plain .md/.txt file — see `gdrive_write_file`.
    Not for: spreadsheets — use gsheets-mcp.

    Use dry_run=True to preview the change without committing.

    Args:
        title: Title of the Google Doc
        content: Optional plain-text content to seed into the Doc
        parent_folder: Optional parent folder name (defaults to My Drive root)
        parent_folder_id: Optional stable parent folder ID
        dry_run: Preview the create without committing it
    """
    try:
        service = google_drive.authenticate()
        parent, parent_error = _resolve_folder_or_error(
            service,
            folder_name=parent_folder,
            folder_id=parent_folder_id,
            required=False,
        )
        if parent_error:
            return parent_error

        parent_id = parent["id"] if parent else None
        parent_label = parent["name"] if parent else "My Drive root"
        if dry_run:
            return {
                "dry_run": True,
                "would_create_doc": title,
                "parent_folder": parent_label,
                "content_chars": len(content or ""),
            }

        result = google_drive.create_google_doc(
            service,
            title,
            content=content,
            parent_id=parent_id,
        )
        restore_token = _make_restore_token(
            {
                "operation": "gdrive_create_doc",
                "file_id": result["id"],
                "title": result["name"],
            }
        )
        return {
            "status": "ok",
            "created": True,
            "message": f"Created Google Doc '{result['name']}' under '{parent_label}'",
            "file_id": result["id"],
            "name": result["name"],
            "mimeType": result.get("mimeType"),
            "parents": result.get("parents", []),
            "webViewLink": result.get("webViewLink", ""),
            "restore_token": restore_token,
            "undo_tool": "gdrive_undo",
        }
    except Exception as e:
        return _exception_envelope(e)


@mcp.tool()
def gdrive_copy(
    file_name: str | None = None,
    new_name: str | None = None,
    destination_folder: str | None = None,
    destination_folder_id: str | None = None,
    file_id: str | None = None,
    dry_run: bool = False,
) -> str | dict:
    """
    Copy a Google Drive file, optionally renaming and/or placing under a new folder.

    Discovery: run `gdrive_search` or `gdrive_list_folder` first to obtain
    `file_name` / `file_id` and destination folder values.

    Use this for: duplicating a file/template.
    Not for: moving the original file — see `gdrive_move`.

    Use dry_run=True to preview the change without committing.

    Args:
        file_name: Name of the file to copy
        new_name: Optional name for the copy
        destination_folder: Optional destination folder name
        destination_folder_id: Optional stable destination folder ID
        file_id: Optional stable file ID (preferred when known)
        dry_run: Preview the copy without committing it
    """
    try:
        service = google_drive.authenticate()
        file_info, error = _resolve_file_or_error(
            service,
            file_name=file_name,
            file_id=file_id,
        )
        if error:
            return error

        dest = None
        if destination_folder or destination_folder_id:
            dest, dest_error = _resolve_folder_or_error(
                service,
                folder_name=destination_folder,
                folder_id=destination_folder_id,
            )
            if dest_error:
                return dest_error

        dest_label = dest["name"] if dest else "same parent"
        copy_name = new_name or f"{file_info['name']} (Copy)"
        if dry_run:
            return {
                "dry_run": True,
                "would_copy": file_info["name"],
                "to_name": copy_name,
                "destination_folder": dest_label,
                "current_id": file_info["id"],
            }

        result = google_drive.copy_file(
            service,
            file_info["id"],
            new_name=copy_name,
            parent_id=dest["id"] if dest else None,
        )
        restore_token = _make_restore_token(
            {
                "operation": "gdrive_copy",
                "file_id": result["id"],
                "source_file_id": file_info["id"],
                "file_name": result["name"],
            }
        )
        return {
            "status": "ok",
            "message": f"Copied '{file_info['name']}' to '{result['name']}'",
            "file_id": result["id"],
            "name": result["name"],
            "parents": result.get("parents", []),
            "webViewLink": result.get("webViewLink", ""),
            "restore_token": restore_token,
            "undo_tool": "gdrive_undo",
        }
    except Exception as e:
        return _exception_envelope(e)


@mcp.tool()
def gdrive_trash(
    file_name: str | None = None,
    file_id: str | None = None,
    dry_run: bool = False,
) -> str | dict:
    """
    Move a Google Drive file or folder to trash.

    Discovery: run `gdrive_search` or `gdrive_list_folder` first to obtain
    `file_name` / `file_id` values.

    Use this for: soft-deleting a file/folder (recoverable from trash).
    Not for: permanently deleting — not supported by this server.

    Use dry_run=True to preview the change without committing.

    Args:
        file_name: Name of the file/folder to trash
        file_id: Optional stable file ID (preferred when known)
        dry_run: Preview the trash without committing it
    """
    try:
        service = google_drive.authenticate()
        file_info, error = _resolve_file_or_error(
            service,
            file_name=file_name,
            file_id=file_id,
        )
        if error:
            return error

        if dry_run:
            return {
                "dry_run": True,
                "would_trash": file_info["name"],
                "current_id": file_info["id"],
                "mime_type": file_info.get("mimeType"),
            }

        result = google_drive.trash_file(service, file_info["id"])
        restore_token = _make_restore_token(
            {
                "operation": "gdrive_trash",
                "file_id": result["id"],
                "file_name": result["name"],
            }
        )
        return {
            "status": "ok",
            "message": f"Moved '{result['name']}' to trash",
            "file_id": result["id"],
            "name": result["name"],
            "trashed": result.get("trashed", True),
            "restore_token": restore_token,
            "undo_tool": "gdrive_undo",
        }
    except Exception as e:
        return _exception_envelope(e)


@mcp.tool()
def gdrive_share(
    email: str,
    role: str = "reader",
    file_name: str | None = None,
    file_id: str | None = None,
    send_notification: bool = False,
    dry_run: bool = False,
) -> str | dict:
    """
    Share a Google Drive file/folder with a user email.

    Discovery: run `gdrive_search` or `gdrive_list_folder` first to obtain
    `file_name` / `file_id` values.

    Use this for: granting reader/commenter/writer access to a collaborator.
    Not for: transferring ownership.

    Valid roles: reader, commenter, writer.
    Use dry_run=True to preview the change without committing.

    Args:
        email: Collaborator email address
        role: Access role — one of reader, commenter, writer
        file_name: Name of the file/folder to share
        file_id: Optional stable file ID (preferred when known)
        send_notification: Whether Google should email the collaborator
        dry_run: Preview the share without committing it
    """
    try:
        service = google_drive.authenticate()
        file_info, error = _resolve_file_or_error(
            service,
            file_name=file_name,
            file_id=file_id,
        )
        if error:
            return error

        if dry_run:
            return {
                "dry_run": True,
                "would_share": file_info["name"],
                "file_id": file_info["id"],
                "email": email,
                "role": role,
                "send_notification": send_notification,
            }

        permission = google_drive.share_file(
            service,
            file_info["id"],
            email,
            role=role,
            send_notification=send_notification,
        )
        restore_token = _make_restore_token(
            {
                "operation": "gdrive_share",
                "file_id": file_info["id"],
                "permission_id": permission["id"],
                "email": email,
                "role": role,
            }
        )
        return {
            "status": "ok",
            "message": f"Shared '{file_info['name']}' with {email} as {role}",
            "file_id": file_info["id"],
            "name": file_info["name"],
            "permission_id": permission["id"],
            "email": permission.get("emailAddress", email),
            "role": permission.get("role", role),
            "restore_token": restore_token,
            "undo_tool": "gdrive_undo",
        }
    except Exception as e:
        return _exception_envelope(e)


@mcp.tool()
def gdrive_undo(restore_token: str) -> dict:
    """
    Undo a prior Drive mutation using its restore_token.

    Supports: gdrive_rename, gdrive_move, gdrive_create_folder, gdrive_write_file
    (create only), gdrive_create_doc, gdrive_copy, gdrive_trash, gdrive_share.

    Args:
        restore_token: Token returned by a mutating gdrive_* tool after commit
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

        if operation in {
            "gdrive_create_folder",
            "gdrive_write_file",
            "gdrive_create_doc",
            "gdrive_copy",
        }:
            result = google_drive.trash_file(service, str(payload["file_id"]))
            return {
                "status": "ok",
                "operation": operation,
                "message": f"Moved created item '{result['name']}' to trash",
                "file_id": result["id"],
                "name": result["name"],
                "trashed": result.get("trashed", True),
            }

        if operation == "gdrive_trash":
            result = google_drive.untrash_file(service, str(payload["file_id"]))
            return {
                "status": "ok",
                "operation": operation,
                "message": f"Restored '{result['name']}' from trash",
                "file_id": result["id"],
                "name": result["name"],
                "trashed": result.get("trashed", False),
            }

        if operation == "gdrive_share":
            google_drive.delete_permission(
                service,
                str(payload["file_id"]),
                str(payload["permission_id"]),
            )
            return {
                "status": "ok",
                "operation": operation,
                "message": (
                    f"Removed {payload.get('role')} access for "
                    f"{payload.get('email')} from file {payload['file_id']}"
                ),
                "file_id": payload["file_id"],
                "permission_id": payload["permission_id"],
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

    Sibling tools: use onedrive_list_folder for nested folders and
    onedrive_read_file to read a discovered file.
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

    Discovery: use onedrive_list_root or onedrive_search first when the exact
    folder_path is unknown.

    Sibling tools: use onedrive_list_root for the top-level folder and
    onedrive_read_file to read a discovered file.
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

    Discovery: use onedrive_search, onedrive_list_root, or
    onedrive_list_folder first to find the exact file_path.
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
