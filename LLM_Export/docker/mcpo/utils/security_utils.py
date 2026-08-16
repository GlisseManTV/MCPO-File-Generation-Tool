"""Security utilities for input validation."""
import os
from pathlib import PurePath


def safe_filename(filename: str) -> str:
    """
    Validate and sanitize a filename.
    Rejects path traversal attempts and returns the cleaned filename.

    Raises ValueError if the filename is malicious.
    """
    if not filename or not filename.strip():
        raise ValueError("Filename cannot be empty")

    if '..' in filename:
        raise ValueError("Filename contains path traversal characters")

    if filename.startswith('/') or filename.startswith('\\'):
        raise ValueError("Filename cannot be a path")

    path = PurePath(filename)
    basename = path.name

    if str(path) != basename:
        raise ValueError("Filename contains hidden path components")

    if not basename:
        raise ValueError("Filename is empty after normalization")

    return basename


def sanitize_path(base_dir: str, user_path: str) -> str:
    """
    Validate that user_path resolves inside base_dir.
    Prevents path traversal attacks.

    Raises ValueError if the resolved path escapes base_dir.
    """
    real_base = os.path.realpath(base_dir)
    resolved = os.path.realpath(os.path.join(base_dir, user_path))

    if not resolved.startswith(real_base + os.sep) and resolved != real_base:
        raise ValueError("Path traversal detected: path escapes base directory")

    return resolved