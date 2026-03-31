from fastapi import FastAPI, HTTPException
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse
import uvicorn
import os
import pathlib
from urllib.parse import unquote

EXPORT_DIR_ENV = os.getenv("FILE_EXPORT_DIR")
EXPORT_DIR = (EXPORT_DIR_ENV or r"/output").rstrip("/")
os.makedirs(EXPORT_DIR, exist_ok=True)

app = FastAPI()

@app.get("/files/{folder_name}/{filename}")
async def serve_file(folder_name: str, filename: str):
    # Decode percent-encoded characters from URL (Russian, accents, etc.)
    decoded_filename = unquote(filename)
    decoded_folder = unquote(folder_name)
    file_path = os.path.join(EXPORT_DIR, decoded_folder, decoded_filename)
    if not os.path.isfile(file_path):
        raise HTTPException(status_code=404, detail="File not found")
    return FileResponse(
        path=file_path,
        media_type='application/octet-stream',
        filename=decoded_filename, 
        headers={"Content-Disposition": f"attachment; filename={decoded_filename}"}
    )

app.mount("/files", StaticFiles(directory=EXPORT_DIR), name="files")

if __name__ == "__main__":
    uvicorn.run(app, host="0.0.0.0", port=9003)