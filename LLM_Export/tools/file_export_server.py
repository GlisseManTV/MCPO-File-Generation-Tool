from fastapi import FastAPI, HTTPException
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse
import uvicorn
import os
import pathlib
from urllib.parse import unquote, quote

EXPORT_DIR_ENV = os.getenv("FILE_EXPORT_DIR")
EXPORT_DIR = (EXPORT_DIR_ENV or r"C:\temp\output").rstrip("/")

os.makedirs(EXPORT_DIR, exist_ok=True)

app = FastAPI()

@app.get("/files/{folder_name}/{filename}")
async def serve_file(folder_name: str, filename: str):
    # Decode the filename from URL-encoded format to match actual file on disk
    decoded_filename = unquote(filename)
    file_path = os.path.join(EXPORT_DIR, folder_name, decoded_filename)
    if not os.path.isfile(file_path):
        raise HTTPException(status_code=404, detail="File not found")
    ascii_fallback = decoded_filename.encode("ascii", "ignore").decode("ascii") or "download"
    encoded_filename = quote(decoded_filename)
    content_disposition = (
        f"attachment; filename=\"{ascii_fallback}\"; "
        f"filename*=UTF-8''{encoded_filename}"
    )
    return FileResponse(
        path=file_path,
        media_type='application/octet-stream',
        headers={
            # RFC5987 header is ASCII-only; quote() ensures safe encoding
            "Content-Disposition": f"attachment; filename*=UTF-8''{quote(decoded_filename)}"
        }
    )

app.mount("/files", StaticFiles(directory=EXPORT_DIR), name="files")

if __name__ == "__main__":
    uvicorn.run(app, host="0.0.0.0", port=9003)
