import ast
import importlib
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
SERVER_PATH = REPO_ROOT / "src" / "server.py"

TOOLS = {
    "gdrive_rename": ("gdrive_move", "gdrive_read_file", "gdrive_create_folder"),
    "gdrive_move": ("gdrive_rename", "gdrive_create_folder"),
    "gdrive_create_folder": ("gdrive_rename", "gdrive_move", "gdrive_create_doc"),
    "gdrive_write_file": ("gdrive_create_doc",),
    "gdrive_create_doc": ("gdrive_write_file", "gdrive_create_folder"),
    "gdrive_copy": ("gdrive_move",),
    "gdrive_trash": ("gdrive_search",),
    "gdrive_share": ("gdrive_search",),
    "gdrive_download_file": ("gdrive_read_file",),
}


def _ast_docstrings():
    tree = ast.parse(SERVER_PATH.read_text(encoding="utf-8"))
    return {
        node.name: ast.get_docstring(node) or ""
        for node in tree.body
        if isinstance(node, ast.FunctionDef) and node.name in TOOLS
    }


def _registered_tool_docstrings():
    server = importlib.import_module("src.server")
    registered_tools = server.mcp._tool_manager._tools

    docs = {}
    for tool_name in TOOLS:
        tool = registered_tools[tool_name]
        fn = getattr(tool, "fn", None) or getattr(tool, "func", None)
        docs[tool_name] = fn.__doc__ or ""
    return docs


def _tool_docstrings():
    try:
        return _registered_tool_docstrings()
    except Exception:
        return _ast_docstrings()


def test_phase5a_tools_have_discovery_contracts():
    docs = _tool_docstrings()

    for tool_name, doc in docs.items():
        assert "Discovery:" in doc, tool_name
        assert "gdrive_list_folder" in doc or "gdrive_search" in doc, tool_name


def test_phase5a_tools_have_sibling_disambiguation_contracts():
    docs = _tool_docstrings()

    for tool_name, siblings in TOOLS.items():
        doc = docs[tool_name]
        assert "Use this for:" in doc, tool_name
        assert "Not for:" in doc, tool_name
        assert any(sibling in doc for sibling in siblings), tool_name
