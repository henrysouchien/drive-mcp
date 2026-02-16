"""
MCP Server for Google Drive and OneDrive.
Provides tools for listing and searching files in both cloud storage services.
"""

import json

from mcp.server.fastmcp import FastMCP

from . import google_drive
from . import onedrive

# Create the MCP server
mcp = FastMCP("drive-mcp")


# =============================================================================
# GOOGLE DRIVE TOOLS
# =============================================================================

@mcp.tool()
def gdrive_list_folder(folder_name: str) -> str:
    """
    List files in a Google Drive folder by name.

    Args:
        folder_name: Name of the folder to list (e.g., "Stock Investor Accelerator")
    """
    try:
        service = google_drive.authenticate()
        folder_id = google_drive.get_folder_id(service, folder_name)

        if not folder_id:
            return f"Folder '{folder_name}' not found in Google Drive."

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
        return f"Error: {str(e)}"


@mcp.tool()
def gdrive_list_folder_recursive(folder_name: str) -> str:
    """
    Recursively list all files in a Google Drive folder and its subfolders.

    Args:
        folder_name: Name of the folder to list (e.g., "Stock Investor Accelerator")
    """
    try:
        service = google_drive.authenticate()
        folder_id = google_drive.get_folder_id(service, folder_name)

        if not folder_id:
            return f"Folder '{folder_name}' not found in Google Drive."

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
        return f"Error: {str(e)}"


@mcp.tool()
def gdrive_search(query: str, max_results: int = 20) -> str:
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
            if f.get('modifiedTime'):
                result += f"   Modified: {f['modifiedTime']}\n"
            if f.get('webViewLink'):
                result += f"   Link: {f['webViewLink']}\n"

        return result
    except Exception as e:
        return f"Error: {str(e)}"


@mcp.tool()
def gdrive_read_file(file_name: str, max_chars: int = 100000) -> str:
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
        return f"Error: {str(e)}"


@mcp.tool()
def gdrive_rename(file_name: str, new_name: str) -> str:
    """
    Rename a file in Google Drive.

    Args:
        file_name: Current name of the file to rename
        new_name: New name for the file
    """
    try:
        service = google_drive.authenticate()
        file_info = google_drive.find_file_by_name(service, file_name)
        if not file_info:
            return f"File not found: {file_name}"
        result = google_drive.rename_file(service, file_info['id'], new_name)
        return f"Renamed '{file_name}' to '{result['name']}'"
    except Exception as e:
        return f"Error: {str(e)}"


@mcp.tool()
def gdrive_move(file_name: str, destination_folder: str) -> str:
    """
    Move a file to a different folder in Google Drive.

    Args:
        file_name: Name of the file to move
        destination_folder: Name of the destination folder
    """
    try:
        service = google_drive.authenticate()
        file_info = google_drive.find_file_by_name(service, file_name)
        if not file_info:
            return f"File not found: {file_name}"
        folder_id = google_drive.get_folder_id(service, destination_folder)
        if not folder_id:
            return f"Folder not found: {destination_folder}"
        result = google_drive.move_file(service, file_info['id'], folder_id)
        return f"Moved '{result['name']}' to '{destination_folder}'"
    except Exception as e:
        return f"Error: {str(e)}"


# =============================================================================
# GOOGLE SHEETS TOOLS
# =============================================================================

def _json_error(operation: str, error: Exception) -> str:
    """Return standardized JSON error payload for Sheets tools."""
    return json.dumps({
        "status": "error",
        "error": str(error),
        "operation": operation
    })


@mcp.tool()
def gsheet_list_tabs(spreadsheet: str) -> str:
    """
    List tabs in a Google Sheets spreadsheet.

    Args:
        spreadsheet: Spreadsheet name or spreadsheet ID
    """
    try:
        spreadsheet_id, title = google_drive.resolve_spreadsheet_id(spreadsheet)
        sheets_service = google_drive.get_sheets_service()
        tabs = google_drive.list_sheet_tabs(sheets_service, spreadsheet_id)
        return json.dumps({
            "status": "ok",
            "spreadsheet_id": spreadsheet_id,
            "title": title,
            "tabs": tabs
        })
    except Exception as e:
        return _json_error("gsheet_list_tabs", e)


@mcp.tool()
def gsheet_read_range(spreadsheet: str, cell_range: str) -> str:
    """
    Read a range from a Google Sheets spreadsheet.

    Args:
        spreadsheet: Spreadsheet name or spreadsheet ID
        cell_range: A1 notation range (e.g., "Sheet1!A1:C10")
    """
    try:
        spreadsheet_id, _ = google_drive.resolve_spreadsheet_id(spreadsheet)
        sheets_service = google_drive.get_sheets_service()
        values = google_drive.read_sheet_range(sheets_service, spreadsheet_id, cell_range)
        return json.dumps({
            "status": "ok",
            "spreadsheet_id": spreadsheet_id,
            "range": cell_range,
            "values": values
        })
    except Exception as e:
        return _json_error("gsheet_read_range", e)


@mcp.tool()
def gsheet_update_range(spreadsheet: str, cell_range: str, values: list[list]) -> str:
    """
    Update a range in a Google Sheets spreadsheet.

    Args:
        spreadsheet: Spreadsheet name or spreadsheet ID
        cell_range: A1 notation range (e.g., "Sheet1!A1:C2")
        values: 2D array of values to write
    """
    try:
        spreadsheet_id, _ = google_drive.resolve_spreadsheet_id(spreadsheet)
        sheets_service = google_drive.get_sheets_service()
        update_result = google_drive.update_sheet_range(
            sheets_service, spreadsheet_id, cell_range, values
        )
        return json.dumps({
            "status": "ok",
            "updatedRange": update_result.get("updatedRange", ""),
            "updatedCells": update_result.get("updatedCells", 0)
        })
    except Exception as e:
        return _json_error("gsheet_update_range", e)


@mcp.tool()
def gsheet_append_rows(spreadsheet: str, cell_range: str, values: list[list]) -> str:
    """
    Append rows to a range in a Google Sheets spreadsheet.

    Args:
        spreadsheet: Spreadsheet name or spreadsheet ID
        cell_range: A1 notation range (e.g., "Sheet1!A:C")
        values: 2D array of row values to append
    """
    try:
        spreadsheet_id, _ = google_drive.resolve_spreadsheet_id(spreadsheet)
        sheets_service = google_drive.get_sheets_service()
        append_result = google_drive.append_sheet_rows(
            sheets_service, spreadsheet_id, cell_range, values
        )
        return json.dumps({
            "status": "ok",
            "updatedRange": append_result.get("updatedRange", ""),
            "updatedCells": append_result.get("updatedCells", 0)
        })
    except Exception as e:
        return _json_error("gsheet_append_rows", e)


@mcp.tool()
def gsheet_create(title: str) -> str:
    """
    Create a new Google Sheets spreadsheet.

    Args:
        title: Title for the new spreadsheet
    """
    try:
        sheets_service = google_drive.get_sheets_service()
        spreadsheet_id, url = google_drive.create_spreadsheet(sheets_service, title)
        return json.dumps({
            "status": "ok",
            "spreadsheet_id": spreadsheet_id,
            "url": url
        })
    except Exception as e:
        return _json_error("gsheet_create", e)


# =============================================================================
# ONEDRIVE TOOLS
# =============================================================================

@mcp.tool()
def onedrive_list_root() -> str:
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
        return f"Error: {str(e)}"


@mcp.tool()
def onedrive_list_folder(folder_path: str) -> str:
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
        return f"Error: {str(e)}"


@mcp.tool()
def onedrive_search(query: str, max_results: int = 20) -> str:
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
        return f"Error: {str(e)}"


@mcp.tool()
def onedrive_read_file(file_path: str, max_chars: int = 100000) -> str:
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
        return f"Error: {str(e)}"


# =============================================================================
# MAIN ENTRY POINT
# =============================================================================

def main():
    mcp.run()


if __name__ == "__main__":
    main()
