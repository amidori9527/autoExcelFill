from __future__ import annotations

from pathlib import Path
import sys


def application_directory() -> Path:
    """Return a writable directory beside the distributed application."""
    executable = Path(sys.executable).resolve()
    if (
        sys.platform == "darwin"
        and executable.parent.name == "MacOS"
        and executable.parent.parent.name == "Contents"
    ):
        return executable.parents[3]
    return executable.parent


def workspace_directory() -> Path:
    if getattr(sys, "frozen", False):
        return application_directory() / "workspace"
    return Path(__file__).resolve().parents[2] / "workspace"


def ensure_workspace_directories() -> Path:
    workspace = workspace_directory()
    (workspace / "diffOrders").mkdir(parents=True, exist_ok=True)
    return workspace


def bundled_resource(file_name: str) -> Path | None:
    bundle_root = getattr(sys, "_MEIPASS", None)
    if bundle_root is None:
        return None
    return Path(bundle_root) / file_name
