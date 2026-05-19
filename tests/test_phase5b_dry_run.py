import importlib
import importlib.util
import inspect
import sys
import types
from unittest.mock import Mock, patch


def _server_module():
    if "pptx" not in sys.modules and importlib.util.find_spec("pptx") is None:
        pptx = types.ModuleType("pptx")
        pptx.Presentation = object
        sys.modules["pptx"] = pptx

    return importlib.import_module("src.server")


def _registered_tool_fn(server, tool_name):
    tool = server.mcp._tool_manager._tools[tool_name]
    return getattr(tool, "fn", None) or getattr(tool, "func", None)


def test_phase5b_dry_run_params_are_in_tool_signatures():
    server = _server_module()

    for tool_name in ("gdrive_rename", "gdrive_move"):
        fn = _registered_tool_fn(server, tool_name)
        dry_run = inspect.signature(fn).parameters["dry_run"]

        assert dry_run.default is False


def test_gdrive_rename_dry_run_returns_preview_without_mutation():
    server = _server_module()
    service = Mock(name="service")
    file_info = {
        "id": "file-123",
        "name": "Old Report",
        "mimeType": "text/plain",
        "parents": ["folder-1"],
    }

    with (
        patch.object(server.google_drive, "authenticate", return_value=service),
        patch.object(server.google_drive, "find_file_by_name", return_value=file_info) as find_file,
        patch.object(server.google_drive, "get_parent_folder_name", return_value="Current Folder") as parent_name,
        patch.object(server.google_drive, "rename_file") as rename_file,
    ):
        result = server.gdrive_rename("Old Report", "New Report", dry_run=True)

    assert result == {
        "dry_run": True,
        "would_rename": "Old Report",
        "to": "New Report",
        "current_id": "file-123",
        "current_parent": "Current Folder",
    }
    find_file.assert_called_once_with(service, "Old Report")
    parent_name.assert_called_once_with(service, file_info)
    rename_file.assert_not_called()
    service.files.assert_not_called()


def test_gdrive_move_dry_run_returns_preview_without_mutation():
    server = _server_module()
    service = Mock(name="service")
    file_info = {
        "id": "file-456",
        "name": "Planning Doc",
        "mimeType": "application/vnd.google-apps.document",
        "parents": ["folder-1"],
    }

    with (
        patch.object(server.google_drive, "authenticate", return_value=service),
        patch.object(server.google_drive, "find_file_by_name", return_value=file_info) as find_file,
        patch.object(server.google_drive, "get_parent_folder_name", return_value="Source Folder") as parent_name,
        patch.object(server.google_drive, "get_folder_id") as get_folder_id,
        patch.object(server.google_drive, "move_file") as move_file,
    ):
        result = server.gdrive_move("Planning Doc", "Archive", dry_run=True)

    assert result == {
        "dry_run": True,
        "would_move": "Planning Doc",
        "from": "Source Folder",
        "to": "Archive",
        "current_id": "file-456",
    }
    find_file.assert_called_once_with(service, "Planning Doc")
    parent_name.assert_called_once_with(service, file_info)
    get_folder_id.assert_not_called()
    move_file.assert_not_called()
    service.files.assert_not_called()
