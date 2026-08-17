import os
import uuid
import datetime
import shutil
import threading
import logging
import base64
import requests
from requests.auth import HTTPBasicAuth
from io import BytesIO
from pathlib import Path
from urllib.parse import quote

from .security_utils import safe_filename

EXPORT_DIR_ENV = os.getenv("FILE_EXPORT_DIR")
EXPORT_DIR = (EXPORT_DIR_ENV or r"/output").rstrip("/")
os.makedirs(EXPORT_DIR, exist_ok=True)

BASE_URL_ENV = os.getenv("FILE_EXPORT_BASE_URL")
BASE_URL = (BASE_URL_ENV or "http://localhost:9003/files").rstrip("/")

def _image_timeout() -> int:
    """Return the global image timeout (seconds). Falls back to LOCAL_SD_TIMEOUT for SD, then 60."""
    return int(os.getenv("IMAGE_TIMEOUT", os.getenv("LOCAL_SD_TIMEOUT", "60")))

def _public_url(folder_path: str, filename: str) -> str:
    """Build a stable public URL for a generated file with proper encoding.
    
    The filename is URL-encoded using utf-8 to support international characters.
    This prevents UnicodeEncodeError when the server or browser expects ASCII/latin-1.
    """
    folder = os.path.basename(folder_path).lstrip("/")
    name = filename.lstrip("/")
    # Encode filename to handle unicode characters (e.g., cyrillic, chinese, accented chars)
    # Using quote with safe='' ensures all special chars are encoded
    encoded_name = quote(name, safe='._-')
    return f"{BASE_URL}/{folder}/{encoded_name}"

def _generate_unique_folder() -> str:
    folder_name = f"export_{uuid.uuid4().hex[:10]}_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}"
    folder_path = os.path.join(EXPORT_DIR, folder_name)
    os.makedirs(folder_path, exist_ok=True)
    return folder_path

def _generate_filename(folder_path: str, ext: str, filename: str | None = None) -> tuple[str, str]:
    """
    Generate a non-colliding filename in folder_path for extension ext.
    Returns (filepath, filename).
    """
    if not filename:
        filename = f"export_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.{ext}"
    filename = safe_filename(filename)
    base, extension = os.path.splitext(filename)
    filepath = os.path.join(folder_path, filename)
    counter = 1
    while os.path.exists(filepath):
        filename = f"{base}_{counter}{extension}"
        filepath = os.path.join(folder_path, filename)
        counter += 1
    return filepath, filename

def _cleanup_files(folder_path: str, delay_minutes: int):
    def delete_files():
        import time
        time.sleep(delay_minutes * 60)
        try:
            shutil.rmtree(folder_path) 
        except Exception as e:
            import logging
            logging.error(f"Error deleting files : {e}")
    thread = threading.Thread(target=delete_files)
    thread.start()

def upload_file(file_path: str, filename: str, file_type: str, token: str) -> dict:
    """
    Upload a file to OpenWebUI server.
    """
    URL = os.getenv('OWUI_URL')
    url = f"{URL}/api/v1/files/"
    headers = {
        'Authorization': token,
        'Accept': 'application/json'
    }
    
    with open(file_path, 'rb') as f:
        files = {'file': f}
        response = requests.post(url, headers=headers, files=files)

    if response.status_code != 200:
        return {"error": {"message": f'Error uploading file: {response.status_code}'}}
    else:
        return {
            "file_path_download": f"[Download {filename}.{file_type}](/api/v1/files/{response.json()['id']}/content)"
        }

def download_file(file_id: str, token: str) -> BytesIO:
    """
    Download a file from OpenWebUI server.
    """
   
    URL = os.getenv('OWUI_URL')
    url = f"{URL}/api/v1/files/{file_id}/content"
    headers = {
        'Authorization': token,
        'Accept': 'application/json'
    }
    
    response = requests.get(url, headers=headers)
    
    if response.status_code != 200:
        return {"error": {"message": f'Error downloading the file: {response.status_code}'}}
    else:
        return BytesIO(response._content)

def search_image(query: str):
    """
    Search or generate an image based on env var IMAGE_SOURCE.
    Supports: unsplash, local_sd, pexels, openai, gemini.
    Returns a public URL or a local public URL (served via BASE_URL).
    """
    image_source = os.getenv("IMAGE_SOURCE", "unsplash").strip().lower()
    if image_source == "unsplash":
        return search_unsplash(query)
    elif image_source == "local_sd":
        return search_local_sd(query)
    elif image_source == "pexels":
        return search_pexels(query)
    elif image_source == "openai":
        return search_openai(query)
    elif image_source == "gemini":
        return search_gemini(query)
    logging.getLogger(__name__).warning(f"Unknown IMAGE_SOURCE '{image_source}'")
    return None

def search_unsplash(query: str) -> str | None:
    api_key = os.getenv("UNSPLASH_ACCESS_KEY")
    if not api_key:
        logging.getLogger(__name__).warning("UNSPLASH_ACCESS_KEY not set")
        return None
    url = "https://api.unsplash.com/search/photos"
    params = {"query": query, "per_page": 1, "orientation": "landscape"}
    headers = {"Authorization": f"Client-ID {api_key}"}
    try:
        resp = requests.get(url, params=params, headers=headers, timeout=_image_timeout())
        resp.raise_for_status()
        data = resp.json()
        if data.get("results"):
            return data["results"][0]["urls"]["regular"]
    except Exception as e:
        logging.getLogger(__name__).error(f"Unsplash error: {e}")
    return None

def search_pexels(query: str) -> str | None:
    api_key = os.getenv("PEXELS_ACCESS_KEY")
    if not api_key:
        logging.getLogger(__name__).warning("PEXELS_ACCESS_KEY not set")
        return None
    url = "https://api.pexels.com/v1/search"
    params = {"query": query, "per_page": 1, "orientation": "landscape"}
    headers = {"Authorization": api_key}
    try:
        resp = requests.get(url, params=params, headers=headers, timeout=_image_timeout())
        resp.raise_for_status()
        data = resp.json()
        if data.get("photos"):
            return data["photos"][0]["src"]["large"]
    except Exception as e:
        logging.getLogger(__name__).error(f"Pexels error: {e}")
    return None

def search_local_sd(query: str) -> str | None:
    """
    Generate image from a local Stable Diffusion API and store it under EXPORT_DIR
    to serve via BASE_URL.
    """
    log = logging.getLogger(__name__)
    SD_URL = os.getenv("LOCAL_SD_URL")
    SD_USERNAME = os.getenv("LOCAL_SD_USERNAME")
    SD_PASSWORD = os.getenv("LOCAL_SD_PASSWORD")
    DEFAULT_MODEL = os.getenv("LOCAL_SD_DEFAULT_MODEL", "sd_xl_base_1.0.safetensors")
    DEFAULT_STEPS = int(os.getenv("LOCAL_SD_STEPS", 20))
    DEFAULT_WIDTH = int(os.getenv("LOCAL_SD_WIDTH", 512))
    DEFAULT_HEIGHT = int(os.getenv("LOCAL_SD_HEIGHT", 512))
    DEFAULT_CFG_SCALE = float(os.getenv("LOCAL_SD_CFG_SCALE", 1.5))
    DEFAULT_SCHEDULER = os.getenv("LOCAL_SD_SCHEDULER", "Karras")
    DEFAULT_SAMPLE = os.getenv("LOCAL_SD_SAMPLE", "Euler a")
    # LOCAL_SD_TIMEOUT takes precedence over IMAGE_TIMEOUT for SD-specific control
    DEFAULT_SD_TIMEOUT = int(os.getenv("LOCAL_SD_TIMEOUT", str(_image_timeout())))

    if not SD_URL:
        log.warning("LOCAL_SD_URL is not defined.")
        return None

    payload = {
        "prompt": query.strip(),
        "steps": DEFAULT_STEPS,
        "width": DEFAULT_WIDTH,
        "height": DEFAULT_HEIGHT,
        "cfg_scale": DEFAULT_CFG_SCALE,
        "sampler_name": DEFAULT_SAMPLE,
        "scheduler": DEFAULT_SCHEDULER,
        "enable_hr": False,
        "hr_upscaler": "Latent",
        "seed": -1,
        "override_settings": {"sd_model_checkpoint": DEFAULT_MODEL}
    }

    try:
        url = f"{SD_URL}/sdapi/v1/txt2img"
        response = requests.post(
            url, json=payload, headers={"Content-Type": "application/json"},
            auth=HTTPBasicAuth(SD_USERNAME, SD_PASSWORD) if SD_USERNAME or SD_PASSWORD else None,
            timeout=DEFAULT_SD_TIMEOUT
        )
        response.raise_for_status()
        data = response.json()
        images = data.get("images", [])
        if not images:
            log.warning(f"No image generated for: '{query}'")
            return None

        image_b64 = images[0]
        image_data = base64.b64decode(image_b64)
        folder_path = _generate_unique_folder()
        filename = f"{query.replace(' ', '_')}.png"
        filepath = os.path.join(folder_path, filename)
        os.makedirs(os.path.dirname(filepath), exist_ok=True)
        with open(filepath, "wb") as f:
            f.write(image_data)
        return _public_url(folder_path, filename)
    except Exception as e:
        log.error(f"Local SD generation error: {e}")
    return None

def search_gemini(query: str) -> str | None:
    """Generate an image using Google Gemini Image Generation API.
    
    Uses the new google.genai SDK (client.interactions.create),
    falls back to raw HTTP if SDK unavailable or call fails.
    """
    log = logging.getLogger(__name__)
    api_key = os.getenv("GOOGLE_GEMINI_API_KEY")
    if not api_key:
        log.warning("GOOGLE_GEMINI_API_KEY not set")
        return None

    api_base = os.getenv("GOOGLE_GEMINI_API_BASE", "https://generativelanguage.googleapis.com/v1beta")
    model = os.getenv("GEMINI_MODEL", "gemini-3.1-flash-image")
    image_b64 = None

    # --- Try new SDK first ---
    try:
        from google import genai
        client = genai.Client(api_key=api_key)
        interaction = client.interactions.create(
            model=model,
            input=query.strip(),
        )
        image_b64 = interaction.output_image.data
        log.info("Gemini image generated via SDK (google.genai)")
    except Exception as e:
        log.warning("Google.genai SDK failed, falling back to raw HTTP: %s", e)

    # --- Raw HTTP fallback ---
    if not image_b64:
        try:
            url = f"{api_base.rstrip('/')}/models/{model}:generateContent?key={api_key}"
            payload = {
                "contents": [{"parts": [{"text": query.strip()}]}],
                "generationConfig": {"responseMimeType": "image/png"}
            }
            headers = {"Content-Type": "application/json"}
            resp = requests.post(
                url, json=payload, headers=headers, timeout=_image_timeout()
            )
            resp.raise_for_status()
            data = resp.json()
            inline = data["candidates"][0]["content"]["parts"][0].get("inlineData", {})
            image_b64 = base64.b64decode(inline.get("data", ""))
            log.info("Gemini image generated via HTTP fallback")
        except Exception as e:
            log.error("Gemini HTTP fallback error: %s", e)

    # --- Save and return ---
    if image_b64:
        folder_path = _generate_unique_folder()
        filename = f"gemini_{uuid.uuid4().hex[:8]}.png"
        filepath = os.path.join(folder_path, filename)
        with open(filepath, "wb") as f:
            f.write(image_b64)
        return _public_url(folder_path, filename)

    return None


def search_openai(query: str) -> str | None:
    """
    Generate an image using OpenAI DALL-E API (dall-e-3).
    Returns a local public URL (served via BASE_URL) by saving the image locally.
    """
    log = logging.getLogger(__name__)
    api_key = os.getenv("OPENAI_API_KEY")
    if not api_key:
        log.warning("OPENAI_API_KEY not set")
        return None

    api_base = os.getenv("OPENAI_API_BASE", "https://api.openai.com/v1")
    model = os.getenv("OPENAI_DALLE_MODEL", "gpt-image-1.5")
    size = os.getenv("OPENAI_IMAGE_SIZE", "auto")
    quality = os.getenv("OPENAI_IMAGE_QUALITY", "auto")
    output_format = os.getenv("OPENAI_IMAGE_OUTPUT_FORMAT", "png")

    image_b64 = None

    # Try SDK first (handles import + API call)
    try:
        import openai
        client = openai.OpenAI(api_key=api_key, base_url=api_base, timeout=_image_timeout())
        # Only add response_format for DALL-E models, not for gpt-image
        gen_kwargs = {
            "model": model,
            "prompt": query.strip(),
            "size": size,
            "n": 1,
            "quality": quality,
        }
        if "dall-e" in model.lower():
            gen_kwargs["response_format"] = "b64_json"
        response = client.images.generate(**gen_kwargs)
        image_b64 = response.data[0].b64_json
    except Exception as e:
        log.warning("OpenAI SDK failed, falling back to raw HTTP: %s", e)

    # Raw HTTP fallback (used when SDK fails or is not installed)
    if not image_b64:
        try:
            url = f"{api_base.rstrip('/')}/images/generations"
            headers = {
                "Authorization": f"Bearer {api_key}",
                "Content-Type": "application/json",
            }
            # Only add response_format for DALL-E models, not for gpt-image
            payload = {
                "model": model,
                "prompt": query.strip(),
                "size": size,
                "n": 1,
                "quality": quality,
            }
            if "dall-e" in model.lower():
                payload["response_format"] = "b64_json"

            resp = requests.post(url, json=payload, headers=headers, timeout=_image_timeout())
            resp.raise_for_status()
            data = resp.json()
            if data.get("data"):
                image_b64 = data["data"][0]["b64_json"]
        except Exception as e:
            log.error("OpenAI DALL-E error (HTTP fallback): %s", e)

    # Save and return the image
    if image_b64:
        image_data = base64.b64decode(image_b64)
        folder_path = _generate_unique_folder()
        filename = f"openai_{uuid.uuid4().hex[:8]}.{output_format}"
        filepath = os.path.join(folder_path, filename)
        with open(filepath, "wb") as f:
            f.write(image_data)
        return _public_url(folder_path, filename)

    return None


def _create_csv(data: list[list[str]] | list[str], filename: str | None = None, folder_path: str | None = None) -> dict:
    """
    Create a CSV file under folder_path and return {'url','path'}.
    """
    import csv
    if folder_path is None:
        folder_path = _generate_unique_folder()
    if filename:
        filename = safe_filename(filename)
        filepath = os.path.join(folder_path, filename)
        os.makedirs(os.path.dirname(filepath), exist_ok=True)
        fname = filename
    else:
        filepath, fname = _generate_filename(folder_path, "csv")
    with open(filepath, "w", newline="", encoding="utf-8") as f:
        if isinstance(data, list) and data and isinstance(data[0], list):
            csv.writer(f).writerows(data)
        elif isinstance(data, list):
            csv.writer(f).writerow(data)
        else:
            csv.writer(f).writerow([str(data)])
    return {"url": _public_url(folder_path, fname), "path": filepath}

def _create_raw_file(content: str, filename: str | None = None, folder_path: str | None = None) -> dict:
    """
    Create a raw text-like file and return {'url','path'}.
    """
    if folder_path is None:
        folder_path = _generate_unique_folder()
    if filename:
        filename = safe_filename(filename)
        filepath = os.path.join(folder_path, filename)
        os.makedirs(os.path.dirname(filepath), exist_ok=True)
        fname = filename
    else:
        filepath, fname = _generate_filename(folder_path, "txt")
    if fname.lower().endswith(".xml") and isinstance(content, str) and not content.strip().startswith("<?xml"):
        content = f'<?xml version="1.0" encoding="UTF-8"?>\n{content}'
    with open(filepath, "w", encoding="utf-8") as f:
        f.write(content or "")
    return {"url": _public_url(folder_path, fname), "path": filepath}
