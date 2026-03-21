from fastapi import FastAPI, HTTPException
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse, RedirectResponse
import uvicorn
import os
import pathlib

from minio_storage import (
    _is_minio_enabled as _minio_enabled,
    presigned_url as minio_presigned_url,
)

EXPORT_DIR_ENV = os.getenv("FILE_EXPORT_DIR")
EXPORT_DIR = (EXPORT_DIR_ENV or r"C:\temp\output").rstrip("/")

os.makedirs(EXPORT_DIR, exist_ok=True)

app = FastAPI()

@app.get("/files/{folder_name}/{filename}")
async def serve_file(folder_name: str, filename: str):
    if _minio_enabled():
        object_key = f"{folder_name}/{filename}"
        try:
            url = minio_presigned_url(object_key)
            return RedirectResponse(url=url, status_code=302)
        except Exception:
            raise HTTPException(status_code=404, detail="File not found in MinIO")

    file_path = os.path.join(EXPORT_DIR, folder_name, filename)
    if not os.path.isfile(file_path):
        raise HTTPException(status_code=404, detail="File not found")
    return FileResponse(
        path=file_path,
        media_type='application/octet-stream',
        filename=filename,
        headers={"Content-Disposition": f"attachment; filename={filename}"}
    )

if not _minio_enabled():
    app.mount("/files", StaticFiles(directory=EXPORT_DIR), name="files")

if __name__ == "__main__":
    uvicorn.run(app, host="0.0.0.0", port=9003)