from fastapi import FastAPI, HTTPException
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse
import uvicorn
import os
from urllib.parse import unquote, quote
from pathlib import PurePath

# ---------------------------------------------------------------------------
# Security utilities (embedded to keep the file server self-contained)
# ---------------------------------------------------------------------------

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


EXPORT_DIR_ENV = os.getenv("FILE_EXPORT_DIR")
EXPORT_DIR = (EXPORT_DIR_ENV or r"/output").rstrip("/")
os.makedirs(EXPORT_DIR, exist_ok=True)

app = FastAPI()

@app.get("/files/{folder_name}/{filename}")
async def serve_file(folder_name: str, filename: str):
    # Decode percent-encoded characters from URL (Russian, accents, etc.)
    decoded_filename = unquote(filename)
    decoded_folder = unquote(folder_name)

    # Security: validate filenames to prevent path traversal
    try:
        safe_folder = safe_filename(decoded_folder)
        safe_file = safe_filename(decoded_filename)
    except ValueError as e:
        raise HTTPException(status_code=400, detail=f"Invalid path: {e}")

    # Build relative path and validate it stays within EXPORT_DIR
    relative_path = os.path.join(safe_folder, safe_file)
    try:
        file_path = sanitize_path(EXPORT_DIR, relative_path)
    except ValueError as e:
        raise HTTPException(status_code=400, detail=f"Invalid path: {e}")

    if not os.path.isfile(file_path):
        raise HTTPException(status_code=404, detail="File not found")

    # RFC 6266: ASCII fallback + UTF-8 filename*
    ascii_fallback = decoded_filename.encode("ascii", "ignore").decode("ascii") or "download"
    encoded_filename = quote(decoded_filename)
    content_disposition = (
        f'attachment; filename="{ascii_fallback}"; '
        f"filename*=UTF-8''{encoded_filename}"
    )

    return FileResponse(
        path=file_path,
        media_type='application/octet-stream',
        filename=decoded_filename,
        headers={"Content-Disposition": content_disposition}
    )

# Secure StaticFiles mount: use check_dir to validate access
class SecureStaticFiles(StaticFiles):
    """StaticFiles subclass that validates paths stay within the mounted directory."""
    async def get_response(self, path: str, scope):
        # Resolve the requested path and ensure it stays within directory
        full_path = os.path.realpath(os.path.join(self.config["directory"], *path.split("/")))
        base_dir = os.path.realpath(self.config["directory"])
        if not full_path.startswith(base_dir + os.sep) and full_path != base_dir:
            raise HTTPException(status_code=403, detail="Access denied: path outside allowed directory")
        return await super().get_response(path, scope)

app.mount("/files", SecureStaticFiles(directory=EXPORT_DIR), name="files")

if __name__ == "__main__":
    uvicorn.run(app, host="0.0.0.0", port=9003)