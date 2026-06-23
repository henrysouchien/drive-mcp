import importlib
import importlib.util
import inspect
from pathlib import Path
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


def test_google_drive_credential_paths_can_be_overridden_by_env(monkeypatch):
    monkeypatch.setenv("GOOGLE_CREDENTIALS_FILE", "/tmp/drive-credentials.json")
    monkeypatch.setenv("GOOGLE_TOKEN_FILE", "/tmp/drive-token.pickle")
    google_drive = importlib.import_module("src.google_drive")
    google_drive = importlib.reload(google_drive)

    try:
        assert google_drive.CREDENTIALS_FILE == Path("/tmp/drive-credentials.json")
        assert google_drive.TOKEN_FILE == Path("/tmp/drive-token.pickle")
    finally:
        monkeypatch.delenv("GOOGLE_CREDENTIALS_FILE", raising=False)
        monkeypatch.delenv("GOOGLE_TOKEN_FILE", raising=False)
        importlib.reload(google_drive)


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


def test_gdrive_rename_not_found_returns_structured_error():
    server = _server_module()
    service = Mock(name="service")

    with (
        patch.object(server.google_drive, "authenticate", return_value=service),
        patch.object(server.google_drive, "find_file_by_name", return_value=None) as find_file,
        patch.object(server.google_drive, "rename_file") as rename_file,
    ):
        result = server.gdrive_rename("Missing Doc", "New Name")

    assert result == {
        "status": "error",
        "error_class": "FileNotFound",
        "message": "file not found: Missing Doc",
        "names_correction": {"file": "Run gdrive_search and use an exact returned name."},
        "suggested_tool_calls": [{"name": "gdrive_search", "args": {"query": "Missing Doc"}}],
    }
    find_file.assert_called_once_with(service, "Missing Doc")
    rename_file.assert_not_called()


def test_gdrive_rename_returns_restore_token_and_undo_restores_name():
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
        patch.object(server.google_drive, "find_file_by_name", return_value=file_info),
        patch.object(server.google_drive, "rename_file", return_value={"id": "file-123", "name": "New Report"}) as rename_file,
    ):
        result = server.gdrive_rename("Old Report", "New Report")

    assert result["status"] == "ok"
    assert result["undo_tool"] == "gdrive_undo"
    assert result["restore_token"]
    rename_file.assert_called_once_with(service, "file-123", "New Report")

    with (
        patch.object(server.google_drive, "authenticate", return_value=service),
        patch.object(server.google_drive, "rename_file", return_value={"id": "file-123", "name": "Old Report"}) as undo_rename,
    ):
        undo = server.gdrive_undo(result["restore_token"])

    assert undo["status"] == "ok"
    assert undo["operation"] == "gdrive_rename"
    assert undo["name"] == "Old Report"
    undo_rename.assert_called_once_with(service, "file-123", "Old Report")


def test_gdrive_move_returns_restore_token_and_undo_restores_parents():
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
        patch.object(server.google_drive, "find_file_by_name", return_value=file_info),
        patch.object(server.google_drive, "get_folder_id", return_value="folder-2"),
        patch.object(server.google_drive, "move_file", return_value={"id": "file-456", "name": "Planning Doc", "parents": ["folder-2"]}) as move_file,
    ):
        result = server.gdrive_move("Planning Doc", "Archive")

    assert result["status"] == "ok"
    assert result["undo_tool"] == "gdrive_undo"
    assert result["restore_token"]
    move_file.assert_called_once_with(service, "file-456", "folder-2")

    with (
        patch.object(server.google_drive, "authenticate", return_value=service),
        patch.object(server.google_drive, "set_file_parents", return_value={"id": "file-456", "name": "Planning Doc", "parents": ["folder-1"]}) as set_parents,
    ):
        undo = server.gdrive_undo(result["restore_token"])

    assert undo["status"] == "ok"
    assert undo["operation"] == "gdrive_move"
    assert undo["parents"] == ["folder-1"]
    set_parents.assert_called_once_with(service, "file-456", ["folder-1"])


def test_onedrive_error_returns_structured_envelope():
    server = _server_module()

    with patch.object(server.onedrive, "list_root_items", side_effect=RuntimeError("OneDrive not authenticated")):
        result = server.onedrive_list_root()

    assert result == {
        "status": "error",
        "error_class": "RuntimeError",
        "message": "OneDrive not authenticated",
        "names_correction": {},
        "suggested_tool_calls": [{"name": "onedrive_start_reauth", "args": {}}],
    }
