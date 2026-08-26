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

    for tool_name in (
        "gdrive_rename",
        "gdrive_move",
        "gdrive_create_folder",
        "gdrive_write_file",
        "gdrive_create_doc",
        "gdrive_copy",
        "gdrive_trash",
        "gdrive_share",
        "gdrive_download_file",
    ):
        fn = _registered_tool_fn(server, tool_name)
        dry_run = inspect.signature(fn).parameters["dry_run"]

        assert dry_run.default is False


def test_gdrive_search_surfaces_raw_file_id_for_consumers():
    server = _server_module()
    service = Mock(name="service")
    files = [
        {
            "id": "1hTgORVAftkCr9s-P9QqHOmyhH-kWEW4kj-w-UbGzKIk",
            "name": "Investment-newsletters",
            "mimeType": "application/vnd.google-apps.spreadsheet",
            "modifiedTime": "2026-01-28T18:17:29.653Z",
            "webViewLink": "https://docs.google.com/spreadsheets/d/example/edit",
        }
    ]

    with (
        patch.object(server.google_drive, "authenticate", return_value=service),
        patch.object(server.google_drive, "search_files", return_value=files) as search_files,
    ):
        result = server.gdrive_search("Investment-newsletters", max_results=7)

    assert "ID: 1hTgORVAftkCr9s-P9QqHOmyhH-kWEW4kj-w-UbGzKIk" in result
    assert result.index("ID:") < result.index("Modified:") < result.index("Link:")
    search_files.assert_called_once_with(service, "Investment-newsletters", 7)


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
        patch.object(server.google_drive, "resolve_file", return_value=file_info) as resolve_file,
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
    resolve_file.assert_called_once_with(service, file_name="Old Report", file_id=None)
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
        patch.object(server.google_drive, "resolve_file", return_value=file_info) as resolve_file,
        patch.object(server.google_drive, "get_parent_folder_name", return_value="Source Folder") as parent_name,
        patch.object(server.google_drive, "resolve_folder") as resolve_folder,
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
    resolve_file.assert_called_once_with(service, file_name="Planning Doc", file_id=None)
    parent_name.assert_called_once_with(service, file_info)
    resolve_folder.assert_not_called()
    move_file.assert_not_called()
    service.files.assert_not_called()


def test_gdrive_rename_not_found_returns_structured_error():
    server = _server_module()
    service = Mock(name="service")

    with (
        patch.object(server.google_drive, "authenticate", return_value=service),
        patch.object(server.google_drive, "resolve_file", return_value=None) as resolve_file,
        patch.object(server.google_drive, "rename_file") as rename_file,
    ):
        result = server.gdrive_rename("Missing Doc", "New Name")

    assert result == {
        "status": "error",
        "error_class": "FileNotFound",
        "message": "file not found: Missing Doc",
        "names_correction": {"file": "Run gdrive_search and use an exact returned name or ID."},
        "suggested_tool_calls": [{"name": "gdrive_search", "args": {"query": "Missing Doc"}}],
    }
    resolve_file.assert_called_once_with(service, file_name="Missing Doc", file_id=None)
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
        patch.object(server.google_drive, "resolve_file", return_value=file_info),
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
    dest = {
        "id": "folder-2",
        "name": "Archive",
        "mimeType": "application/vnd.google-apps.folder",
    }

    with (
        patch.object(server.google_drive, "authenticate", return_value=service),
        patch.object(server.google_drive, "resolve_file", return_value=file_info),
        patch.object(server.google_drive, "resolve_folder", return_value=dest),
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


def test_gdrive_create_folder_dry_run_returns_preview_without_mutation():
    server = _server_module()
    service = Mock(name="service")
    parent = {
        "id": "parent-1",
        "name": "Projects",
        "mimeType": "application/vnd.google-apps.folder",
    }

    with (
        patch.object(server.google_drive, "authenticate", return_value=service),
        patch.object(server.google_drive, "resolve_folder", return_value=parent) as resolve_folder,
        patch.object(server.google_drive, "find_child_folder", return_value=None) as find_child,
        patch.object(server.google_drive, "create_folder") as create_folder,
    ):
        result = server.gdrive_create_folder(
            "Research",
            parent_folder="Projects",
            dry_run=True,
        )

    assert result == {
        "dry_run": True,
        "would_create": "Research",
        "parent_folder": "Projects",
        "already_exists": False,
        "existing_id": None,
        "exist_ok": False,
    }
    resolve_folder.assert_called_once_with(
        service,
        folder_name="Projects",
        folder_id=None,
    )
    find_child.assert_called_once_with(service, "Research", "parent-1")
    create_folder.assert_not_called()


def test_gdrive_create_folder_returns_restore_token_and_undo_trashes():
    server = _server_module()
    service = Mock(name="service")
    parent = {
        "id": "parent-1",
        "name": "Projects",
        "mimeType": "application/vnd.google-apps.folder",
    }
    created = {
        "id": "folder-123",
        "name": "Research",
        "mimeType": "application/vnd.google-apps.folder",
        "parents": ["parent-1"],
        "webViewLink": "https://drive.google.com/drive/folders/folder-123",
    }

    with (
        patch.object(server.google_drive, "authenticate", return_value=service),
        patch.object(server.google_drive, "resolve_folder", return_value=parent),
        patch.object(server.google_drive, "find_child_folder", return_value=None),
        patch.object(server.google_drive, "create_folder", return_value=created) as create_folder,
    ):
        result = server.gdrive_create_folder("Research", parent_folder="Projects")

    assert result["status"] == "ok"
    assert result["created"] is True
    assert result["file_id"] == "folder-123"
    assert result["undo_tool"] == "gdrive_undo"
    assert result["restore_token"]
    create_folder.assert_called_once_with(service, "Research", "parent-1")

    with (
        patch.object(server.google_drive, "authenticate", return_value=service),
        patch.object(
            server.google_drive,
            "trash_file",
            return_value={"id": "folder-123", "name": "Research", "trashed": True},
        ) as trash_file,
    ):
        undo = server.gdrive_undo(result["restore_token"])

    assert undo["status"] == "ok"
    assert undo["operation"] == "gdrive_create_folder"
    assert undo["trashed"] is True
    trash_file.assert_called_once_with(service, "folder-123")


def test_gdrive_create_folder_exist_ok_reuses_existing_without_restore_token():
    server = _server_module()
    service = Mock(name="service")
    existing = {
        "id": "folder-existing",
        "name": "Research",
        "mimeType": "application/vnd.google-apps.folder",
        "parents": ["root"],
        "webViewLink": "https://drive.google.com/drive/folders/folder-existing",
    }

    with (
        patch.object(server.google_drive, "authenticate", return_value=service),
        patch.object(server.google_drive, "find_child_folder", return_value=existing),
        patch.object(server.google_drive, "create_folder") as create_folder,
    ):
        result = server.gdrive_create_folder("Research", exist_ok=True)

    assert result == {
        "status": "ok",
        "created": False,
        "message": "Folder 'Research' already exists under 'My Drive root'",
        "file_id": "folder-existing",
        "name": "Research",
        "parents": ["root"],
        "webViewLink": "https://drive.google.com/drive/folders/folder-existing",
    }
    create_folder.assert_not_called()


def test_gdrive_create_folder_errors_when_exists_without_exist_ok():
    server = _server_module()
    service = Mock(name="service")
    existing = {
        "id": "folder-existing",
        "name": "Research",
        "mimeType": "application/vnd.google-apps.folder",
        "parents": ["root"],
    }

    with (
        patch.object(server.google_drive, "authenticate", return_value=service),
        patch.object(server.google_drive, "find_child_folder", return_value=existing),
        patch.object(server.google_drive, "create_folder") as create_folder,
    ):
        result = server.gdrive_create_folder("Research")

    assert result["status"] == "error"
    assert result["error_class"] == "FolderAlreadyExists"
    create_folder.assert_not_called()


def test_gdrive_write_file_creates_with_restore_token():
    server = _server_module()
    service = Mock(name="service")
    created = {
        "id": "file-new",
        "name": "notes.md",
        "mimeType": "text/markdown",
        "parents": ["root"],
        "webViewLink": "https://drive.google.com/file/d/file-new/view",
    }

    with (
        patch.object(server.google_drive, "authenticate", return_value=service),
        patch.object(server.google_drive, "write_file", return_value=created) as write_file,
    ):
        result = server.gdrive_write_file("notes.md", "# hello")

    assert result["status"] == "ok"
    assert result["created"] is True
    assert result["restore_token"]
    write_file.assert_called_once()


def test_gdrive_trash_and_undo_restores():
    server = _server_module()
    service = Mock(name="service")
    file_info = {
        "id": "file-1",
        "name": "Old Notes",
        "mimeType": "text/plain",
        "parents": ["root"],
    }

    with (
        patch.object(server.google_drive, "authenticate", return_value=service),
        patch.object(server.google_drive, "resolve_file", return_value=file_info),
        patch.object(
            server.google_drive,
            "trash_file",
            return_value={"id": "file-1", "name": "Old Notes", "trashed": True},
        ),
    ):
        result = server.gdrive_trash(file_id="file-1")

    assert result["status"] == "ok"
    assert result["restore_token"]

    with (
        patch.object(server.google_drive, "authenticate", return_value=service),
        patch.object(
            server.google_drive,
            "untrash_file",
            return_value={"id": "file-1", "name": "Old Notes", "trashed": False},
        ) as untrash,
    ):
        undo = server.gdrive_undo(result["restore_token"])

    assert undo["status"] == "ok"
    assert undo["operation"] == "gdrive_trash"
    untrash.assert_called_once_with(service, "file-1")


def test_gdrive_share_and_undo_removes_permission():
    server = _server_module()
    service = Mock(name="service")
    file_info = {
        "id": "file-1",
        "name": "Shared Doc",
        "mimeType": "application/vnd.google-apps.document",
    }

    with (
        patch.object(server.google_drive, "authenticate", return_value=service),
        patch.object(server.google_drive, "resolve_file", return_value=file_info),
        patch.object(
            server.google_drive,
            "share_file",
            return_value={
                "id": "perm-1",
                "role": "reader",
                "emailAddress": "a@example.com",
                "type": "user",
            },
        ),
    ):
        result = server.gdrive_share("a@example.com", role="reader", file_id="file-1")

    assert result["status"] == "ok"
    assert result["permission_id"] == "perm-1"

    with (
        patch.object(server.google_drive, "authenticate", return_value=service),
        patch.object(server.google_drive, "delete_permission") as delete_permission,
    ):
        undo = server.gdrive_undo(result["restore_token"])

    assert undo["status"] == "ok"
    delete_permission.assert_called_once_with(service, "file-1", "perm-1")


def test_gdrive_copy_dry_run_and_create_doc_dry_run():
    server = _server_module()
    service = Mock(name="service")
    file_info = {
        "id": "file-1",
        "name": "Template",
        "mimeType": "text/plain",
        "parents": ["root"],
    }

    with (
        patch.object(server.google_drive, "authenticate", return_value=service),
        patch.object(server.google_drive, "resolve_file", return_value=file_info),
        patch.object(server.google_drive, "copy_file") as copy_file,
    ):
        copy_preview = server.gdrive_copy(file_id="file-1", new_name="Template 2", dry_run=True)

    assert copy_preview["dry_run"] is True
    assert copy_preview["to_name"] == "Template 2"
    copy_file.assert_not_called()

    with (
        patch.object(server.google_drive, "authenticate", return_value=service),
        patch.object(server.google_drive, "create_google_doc") as create_doc,
    ):
        doc_preview = server.gdrive_create_doc("Brief", content="hi", dry_run=True)

    assert doc_preview == {
        "dry_run": True,
        "would_create_doc": "Brief",
        "parent_folder": "My Drive root",
        "content_chars": 2,
    }
    create_doc.assert_not_called()


def test_gdrive_read_file_prefers_file_id():
    server = _server_module()
    service = Mock(name="service")
    file_info = {
        "id": "file-9",
        "name": "Report",
        "mimeType": "text/plain",
    }

    with (
        patch.object(server.google_drive, "authenticate", return_value=service),
        patch.object(server.google_drive, "resolve_file", return_value=file_info) as resolve_file,
        patch.object(server.google_drive, "read_file_content", return_value="body") as read_content,
    ):
        result = server.gdrive_read_file(file_id="file-9")

    assert result == "body"
    resolve_file.assert_called_once_with(service, file_name=None, file_id="file-9")
    read_content.assert_called_once_with(service, "file-9", "text/plain", 100000)


def test_gdrive_list_shared_drives_formats_ids():
    server = _server_module()
    service = Mock(name="service")

    with (
        patch.object(server.google_drive, "authenticate", return_value=service),
        patch.object(
            server.google_drive,
            "list_shared_drives",
            return_value=[{"id": "drive-1", "name": "Team Drive"}],
        ),
    ):
        result = server.gdrive_list_shared_drives()

    assert "Team Drive" in result
    assert "ID: drive-1" in result
