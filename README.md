# MCPO-File-Generation-Tool – Export Files Directly from Open WebUI

A lightweight, MCPO-integrated tool that lets you **generate and export real files** (PDF, Excel, PowerPoint, ZIP, etc.) directly from Open WebUI — just like ChatGPT or Claude.

✅ Supports both **Python** and **Docker**  
✅ Fully configurable  
✅ Ready for production workflows  
✅ Open source & MIT licensed

---

🚀 **Create and export files easily from Open WebUI!**

This tool allows seamless file generation and export directly from your Open WebUI environment using Python and FastAPI.

## Multi files

https://github.com/user-attachments/assets/41dadef9-7981-4439-bf5f-3b82fcbaff04


## Single archive

https://github.com/user-attachments/assets/1e70a977-62f1-498c-895c-7db135ded95b


# 🚀 Quick Start

## Best practices here: [Best_Practices.md](https://github.com/GlisseManTV/MCPO-File-Generation-Tool/blob/master/Best_Practices.md)
## Prompt examples here: [Prompt_Examples.md](https://github.com/GlisseManTV/MCPO-File-Generation-Tool/blob/master/Prompt_Examples.md)

### 🔧 For Python Users

1. Clone the repo:
   ```bash
   git clone https://github.com/GlisseManTV/MCPO-File-Generation-Tool.git
   ```

2. Update env variables in `config.json`:
  These ones only concerns the MCPO part

   - `PYTHONPATH`: Path to your `LLM_Export` folder (e.g., `C:\temp\LLM_Export`) <=== MANDATORY no default value
   - `FILE_EXPORT_BASE_URL`: URL of your file export server (default is `http://localhost:9003/files`)
   - `FILE_EXPORT_DIR`: Directory where files will be saved (must match the server's export directory) (default is `PYTHONPATH\output`)
   - `PERSISTENT_FILES`: Set to `true` to keep files after download, `false` to delete after delay (default is false)
   - `FILES_DELAY`: Delay in minut to wait before checking for new files (default is 60)
   - `UNSPLASH_ACCESS_KEY`: Your Unsplash API key (no default value, not mandatory but advised) see [here](https://unsplash.com/documentation#creating-a-developer-account)
   - `IMAGE_SOURCE`: "unsplash" to use Unsplash for image generation or "local_sd" to use your local Stable Diffusion instance (default is "unsplash")
   - `LOCAL_SD_URL`: URL of your local Stable Diffusion instance (if using local_sd) (no default value, mandatory if local_sd is used above)
   - `LOCAL_SD_USERNAME`: Username of your local Stable Diffusion instance (if any) (no default value, not mandatory)
   - `LOCAL_SD_PASSWORD`: Password of your local Stable Diffusion instance (if any) (no default value, not mandatory)
   - `LOCAL_SD_DEFAULT_MODEL`: Default model to use (if any) (default `sd_xl_base_1.0.safetensors`, not mandatory)
   - `LOCAL_SD_STEPS`: Number of steps to use (default 20, not mandatory)
   - `LOCAL_SD_WIDTH`: Width of the image to generate (default 512, not mandatory)
   - `LOCAL_SD_HEIGHT`: Height of the image to generate (default 512, not mandatory)
   - `LOCAL_SD_CFG_SCALE`: CFG scale to use (default 1.5, not mandatory)
   - `LOCAL_SD_SCHEDULER`: Scheduler to use (default `Karras`, not mandatory)
   - `LOCAL_SD_SAMPLE`: Sampler to use (default `Euler a`, not mandatory)
   
3. Install dependencies:
   ```bash
   pip install openpyxl reportlab py7zr fastapi uvicorn python-multipart mcp
   ```

4. Run the file server:
   ```bat
   set FILE_EXPORT_DIR=C:\temp\LLM_Export\output
   start "File Export Server" python "YourPATH/LLM_Export/tools/file_export_server.py"
   ```

5. Use it in Open WebUI — your AI can now generate and export files in real time!

---

### PYTHON EXAMPLE
This file only concerns the MCPO part, you need to run the file server separately as shown above
This is an example of a minimal `config.json` for MCPO to enable file export but you can add other (or to other) MCP servers as needed.

```config.json
{
  "mcpServers": {
		"file_export": {
			"command": "python",
			"args": [
				"-m",
				"tools.file_export_mcp"
			],
			"env": {
				"PYTHONPATH": "C:\\temp\\LLM_Export", <==== HERE set the path to your LLM_Export folder (this one is Mandatory)
				"FILE_EXPORT_BASE_URL": "http://localhost:9003/files", <==== HERE set the URL of your file export server
				"FILE_EXPORT_DIR": "C:\\temp\\LLM_Export\\output", <==== HERE set the directory where files will be saved (must match the server's export directory)
				"PERSISTENT_FILES": "false", <==== HERE set to true to keep files after download, false to delete after delay
				"FILES_DELAY": "60", <==== HERE set the delay in minut to wait before checking for new files
                "UNSPLASH_ACCESS_KEY":"", <== Your Unsplash API key (no default value, not mandatory but advised) see [here](https://unsplash.com/documentation#creating-a-developer-account)
				"IMAGE_SOURCE": "local_sd", <==== HERE set to "unsplash" to use Unsplash for image generation or "local_sd" to use your local Stable Diffusion instance>
				"LOCAL_SD_URL": "http://localhost:7860", <==== HERE set to the URL of your local Stable Diffusion instance>
                "LOCAL_SD_USERNAME": "local_user", <==== HERE set to the username of your local Stable Diffusion instance (if any)>
                "LOCAL_SD_PASSWORD": "local_password", <==== HERE set to the password of your local Stable Diffusion instance (if any)>
                "LOCAL_SD_DEFAULT_MODEL": "sd_xl_base_1.0.safetensors", <==== HERE set to the default model to use (if any)>
                "LOCAL_SD_STEPS": "20", <==== HERE set to the number of steps to use (if any)>
                "LOCAL_SD_WIDTH": "512", <==== HERE set to the width of the image to generate (if any)>
                "LOCAL_SD_HEIGHT": "512", <==== HERE set to the height of the image to generate (if any)>
                "LOCAL_SD_CFG_SCALE": "1.5", <==== HERE set to the CFG scale to use (if any)>
                "LOCAL_SD_SCHEDULER": "Karras", <==== HERE set to the scheduler to use (if any)>
                "LOCAL_SD_SAMPLE": "Euler a" <==== HERE set to the sampler to use (if any)>
			},
			"disabled": false,
			"autoApprove": []
		}
}

```

---

## 🐳 For Docker User (Recommended)

Use 
```
docker pull ghcr.io/glissemantv/owui-file-export-server:latest
docker pull ghcr.io/glissemantv/owui-mcpo:latest
```
	

### 🛠️ DOCKER ENV VARIABLES

For OWUI-MCPO
   - `MCPO_API_KEY`: Your MCPO API key (no default value, not mandatory but advised)
   - `FILE_EXPORT_BASE_URL`: URL of your file export server (default is `http://localhost:9003/files`)
   - `FILE_EXPORT_DIR`: Directory where files will be saved (must match the server's export directory) (default is `/output`) path must be mounted as a volume
   - `PERSISTENT_FILES`: Set to `true` to keep files after download, `false` to delete after delay (default is `false`)
   - `FILES_DELAY`: Delay in minut to wait before checking for new files (default is 60)
   - `UNSPLASH_ACCESS_KEY`: Your Unsplash API key (no default value, not mandatory but advised) see [here](https://unsplash.com/documentation#creating-a-developer-account)
   - `IMAGE_SOURCE`: "unsplash" to use Unsplash  or "local_sd" to use your local Stable Diffusion instance (default is "unsplash")
   - `LOCAL_SD_URL`: URL of your local Stable Diffusion instance (if using local_sd) (no default value, mandatory if local_sd is used above)
   - `LOCAL_SD_USERNAME`: Username of your local Stable Diffusion instance (if any) (no default value, not mandatory)
   - `LOCAL_SD_PASSWORD`: Password of your local Stable Diffusion instance (if any) (no default value, not mandatory)
   - `LOCAL_SD_DEFAULT_MODEL`: Default model to use (if any) (default `sd_xl_base_1.0.safetensors`, not mandatory)
   - `LOCAL_SD_STEPS`: Number of steps to use (default 20, not mandatory)
   - `LOCAL_SD_WIDTH`: Width of the image to generate (default 512, not mandatory)
   - `LOCAL_SD_HEIGHT`: Height of the image to generate (default 512, not mandatory)
   - `LOCAL_SD_CFG_SCALE`: CFG scale to use (default 1.5, not mandatory)
   - `LOCAL_SD_SCHEDULER`: Scheduler to use (default `Karras`, not mandatory)
   - `LOCAL_SD_SAMPLE`: Sampler to use (default `Euler a`, not mandatory)

For OWUI-FILE-EXPORT-SERVER
   - `FILE_EXPORT_DIR`: Directory where files will be saved (must match the MCPO's export directory) (default is `/output`) path must be mounted as a volume

> ✅ This ensures MCPO can correctly reach the file export server.
> ❌ If not set, file export will fail with a 404 or connection error.

---

### DOCKER EXAMPLE

Here is an example of a docker run script file to run both the file export server and the MCPO server:
```
docker run -d --name file-export-server --network host -e FILE_EXPORT_DIR=/data/output -p 9003:9003 -v /path/to/your/export/folder:/data/output ghcr.io/glissemantv/owui-file-export-server:latest
docker run -d --name owui-mcpo --network host -e FILE_EXPORT_BASE_URL=http://192.168.0.100:9003/files -e FILE_EXPORT_DIR=/output -e MCPO_API_KEY=top-secret -e PERSISTENT_FILES=True -e FILES_DELAY=1 -e -e LOG_LEVEL=INFO -e UNSPLASH_ACCESS_KEY=top-secret -p 8000:8000 -v /path/to/your/export/folder:/output ghcr.io/glissemantv/owui-mcpo:latest

```

Here is an example of a `docker-compose.yaml` file to run both the file export server and the MCPO server:
```yaml
services:
  file-export-server:
    image: ghcr.io/glissemantv/owui-file-export-server:latest
    container_name: file-export-server
    environment:
      - FILE_EXPORT_DIR=/output
    ports:
      - "9003:9003"
    volumes:
      - /your/export-data:/output

  owui-mcpo:
    image: ghcr.io/glissemantv/owui-mcpo:latest
    container_name: owui-mcpo
    environment:
      - FILE_EXPORT_BASE_URL=http://file-export-server:9003/files
      - FILE_EXPORT_DIR=/output
      - MCPO_API_KEY=top-secret
      - PERSISTENT_FILES=true
      - FILES_DELAY=1
      - LOG_LEVEL=INFO
      - UNSPLASH_ACCESS_KEY=top-secret
      - IMAGE_SOURCE=local_sd
      - LOCAL_SD_URL=http://localhost:7860
      - LOCAL_SD_USERNAME=local_user
      - LOCAL_SD_PASSWORD=local_password
      - LOCAL_SD_DEFAULT_MODEL=sd_xl_base_1.0.safetensors
      - LOCAL_SD_STEPS=20
      - LOCAL_SD_WIDTH=512
      - LOCAL_SD_HEIGHT=512
      - LOCAL_SD_CFG_SCALE=1.5
      - LOCAL_SD_SCHEDULER=Karras
      - LOCAL_SD_SAMPLE=Euler a
    ports:
      - "8000:8000"
    restart: unless-stopped
    volumes:
      - /your/export-data:/output
    depends_on:
      - file-export-server
```
---

## 📌 Document Review Support (v0.7.0+)

A major enhancement in **v0.7.0** introduces **native document revision support** for Word, Excel, and PowerPoint files — enabling AI-powered editing directly within the chat.

### ✅ Supported Formats & Methods

| Format       | Revision Method                     | Notes |
|--------------|-------------------------------------|-------|
| **.docx**    | Native Word comments                | Fully compatible with Microsoft Word and other editors |
| **.xlsx**    | Legacy cell notes (comments)        | Preserves original metadata and structure |
| **.pptx**    | Top-of-slide TextBox (simulated)    | Visual feedback without breaking layout |

> 💡 **How it works**:
> 1. Receive file metadata in the chat  
> 2. Tool detects document type  
> 3. AI analyzes full context and generates edits  
> 4. Changes are applied and returned via download URL  
> 5. Temporary file is automatically deleted  

✅ **No manual file handling required** — everything is seamless.

---

### 🛠️ Required Setup for Document Review

To enable revision, you must:

1. **Add the `Files Metadata Injector Function`**  
   - Available here: [Files Metadata Injector • Open WebUI Community](https://openwebui.com/f/gaetan/files_metadata_injector)  
   - Or place it in your tool's `functions/` directory (Docker path: `/rootPath/functions`)

2. **Set two new environment variables**:
   ```env
   OWUI_URL=https://myAI.myDomain.com
   JWT_SECRET=your_api_key_or_jwt_token
   ```
   - `OWUI_URL`: Your Open WebUI instance URL (e.g., `http://localhost:3000` or `https://openwebui.yourdomain.com`)
   - `JWT_SECRET`: Generated from **User Settings > API Keys** in Open WebUI

---

## 📌 Document Review Support (v0.7.0+)

A major enhancement in **v0.7.0** introduces **native document revision support** for Word, Excel, and PowerPoint files — enabling AI-powered editing directly within the chat.

### ✅ Supported Formats & Methods

| Format       | Revision Method                     | Notes |
|--------------|-------------------------------------|-------|
| **.docx**    | Native Word comments                | Fully compatible with Microsoft Word and other editors |
| **.xlsx**    | Legacy cell notes (comments)        | Preserves original metadata and structure |
| **.pptx**    | Top-of-slide TextBox (simulated)    | Visual feedback without breaking layout |

> 💡 **How it works**:
> 1. Receive file metadata in the chat  
> 2. Tool detects document type  
> 3. AI analyzes full context and generates edits  
> 4. Changes are applied and returned via download URL  
> 5. Temporary file is automatically deleted  

✅ **No manual file handling required** — everything is seamless.

---

### 🛠️ Required Setup for Document Review

To enable revision, you must:

1. **Add the `Files Metadata Injector Function`**  
   - Available here: [Files Metadata Injector • Open WebUI Community](https://openwebui.com/f/gaetan/files_metadata_injector)  
   - Or place it in your tool's `functions/` directory (Docker path: `/rootPath/functions`)

2. **Set two new environment variables**:
   ```env
   OWUI_URL=https://myAI.myDomain.com
   JWT_SECRET=your_api_key_or_jwt_token
   ```
   - `OWUI_URL`: Your Open WebUI instance URL (e.g., `http://localhost:3000` or `https://openwebui.yourdomain.com`)
   - `JWT_SECRET`: Generated from **User Settings > API Keys** in Open WebUI

---

## 📦 Supported File Types

- ✅ `.xlsx` (Excel)
- ✅ `.pdf` (PDF)
- ✅ `.csv` (CSV)
- ✅ `.pptx` (PowerPoint)
- ✅ `.docx` (Word)
- ✅ `.zip`n `tar.gz` and `.7z` (Archives)
- ✅ Any other file type 

---

## 📂 Project Structure

```
MCPO-File-Generation-Tool/
├── LLM_Export/
│   ├── tools/
│   │   ├── file_export_server.py
│   │   └── file_export_mcp.py
│   └── ...
├── docker/
│   ├── file_server/
│   │   ├── Dockerfile.server
│   │   ├── file_server_compose.yaml
│   │   └── file_export_server.py
│   ├── mcpo/
│   │   ├── Dockerfile
│   │   ├── requirements.txt
│   │   ├── config.json
│   │   ├── MCPO_server_compose.yaml
│   │   └──tools/
│   │       └── file_export_mcp.py
│   └── docker-compose.yaml
└── README.md
```

---

## 📌 Notes

- File output paths must match between `file_server` and `MCPO`
- Always use **absolute paths** for volume mounts
  
⚠️Some users are experiencing trouble with the MCPO server, please use this fix⚠️
```config.json
{
  "mcpServers": {
		"file_export": {
			"command": "python", <==== HERE change "python" to "python3", "python3.11" or "python3.12"
			"args": [
				"-m",
				"tools.file_export_mcp"
			],
			"env": {
				"PYTHONPATH": "C:\\temp\\LLM_Export" <==== HERE set the path to your LLM_Export folder (this one is Mandatory)
			},
			"disabled": false,
			"autoApprove": []
		}
}

```
---

## 🌟 Why This Matters

This tool turns Open WebUI into a **true productivity engine** — where AI doesn’t just chat, but **delivers usable, downloadable files**.

---

## 📄 License

MIT License – Feel free to use, modify, and distribute.

---

📬 **Need help?** Open an issue or start a discussion on GitHub! 

---

## 🌟 Credits

A big thank you to the contributors and open-source projects that made this work possible:

- **tjbck** for creating [**Open WebUI**](https://github.com/open-webui/open-webui) and [**mcpo**](https://github.com/open-webui/mcpo), foundational pillars of this integration.

- [**modelcontextprotocol/servers**](https://github.com/modelcontextprotocol/servers) for high-quality tools and architectural inspiration that guided the development of MCP servers and file generation workflows.

-  [**gentoorax**](https://chrislaw.me/) for close collaboration, technical rigor, and invaluable contributions to the quality and stability of this project.

Thank you to everyone for your passion, expertise, and dedication to the open-source community. 🙌

---


---

# 🚀 Quick Start for Development Versions

## Using development versions of libraries is at your own risk. Always test in a safe environment first.

Use 
```
docker pull ghcr.io/glissemantv/owui-file-export-server:dev-latest
docker pull ghcr.io/glissemantv/owui-mcpo:dev-latest
```

### 🛠️ DOCKER ENV VARIABLES

For OWUI-MCPO
   - `MCPO_API_KEY`: Your MCPO API key (no default value, not mandatory but advised)
   - `FILE_EXPORT_BASE_URL`: URL of your file export server (default is `http://localhost:9003/files`)
   - `FILE_EXPORT_DIR`: Directory where files will be saved (must match the server's export directory) (default is `/output`) path must be mounted as a volume
   - `PERSISTENT_FILES`: Set to `true` to keep files after download, `false` to delete after delay (default is `false`)
   - `FILES_DELAY`: Delay in minut to wait before checking for new files (default is 60)
   - `UNSPLASH_ACCESS_KEY`: Your Unsplash API key (no default value, not mandatory but advised) see [here](https://unsplash.com/documentation#creating-a-developer-account)
   - `IMAGE_SOURCE`: "unsplash" to use Unsplash  or "local_sd" to use your local Stable Diffusion instance (default is "unsplash")
   - `LOCAL_SD_URL`: URL of your local Stable Diffusion instance (if using local_sd) (no default value, mandatory if local_sd is used above)
   - `LOCAL_SD_USERNAME`: Username of your local Stable Diffusion instance (if any) (no default value, not mandatory)
   - `LOCAL_SD_PASSWORD`: Password of your local Stable Diffusion instance (if any) (no default value, not mandatory)
   - `LOCAL_SD_DEFAULT_MODEL`: Default model to use (if any) (default `sd_xl_base_1.0.safetensors`, not mandatory)
   - `LOCAL_SD_STEPS`: Number of steps to use (default 20, not mandatory)
   - `LOCAL_SD_WIDTH`: Width of the image to generate (default 512, not mandatory)
   - `LOCAL_SD_HEIGHT`: Height of the image to generate (default 512, not mandatory)
   - `LOCAL_SD_CFG_SCALE`: CFG scale to use (default 1.5, not mandatory)
   - `LOCAL_SD_SCHEDULER`: Scheduler to use (default `Karras`, not mandatory)
   - `LOCAL_SD_SAMPLE`: Sampler to use (default `Euler a`, not mandatory)
  
For OWUI-FILE-EXPORT-SERVER
   - `FILE_EXPORT_DIR`: Directory where files will be saved (must match the MCPO's export directory) (default is `/output`) path must be mounted as a volume

> ✅ This ensures MCPO can correctly reach the file export server.
> ❌ If not set, file export will fail with a 404 or connection error.

---

### DOCKER EXAMPLE


Here is an example of a docker run script file to run both the file export server and the MCPO server:
```
docker run -d --name file-export-server --network host -e FILE_EXPORT_DIR=/data/output -p 9003:9003 -v /path/to/your/export/folder:/data/output ghcr.io/glissemantv/owui-file-export-server:dev-latest
docker run -d --name owui-mcpo --network host -e FILE_EXPORT_BASE_URL=http://192.168.0.100:9003/files -e FILE_EXPORT_DIR=/output -e MCPO_API_KEY=top-secret -e PERSISTENT_FILES=True -e FILES_DELAY=1 -e LOG_LEVEL=DEBUG -e UNSPLASH_ACCESS_KEY=top-secret -p 8000:8000 -v /path/to/your/export/folder:/output ghcr.io/glissemantv/owui-mcpo:dev-latest
```

Here is an example of a `docker-compose.yaml` file to run both the file export server and the MCPO server:
```yaml
services:
  file-export-server:
    image: ghcr.io/glissemantv/owui-file-export-server:dev-latest
    container_name: file-export-server
    environment:
      - FILE_EXPORT_DIR=/output
    ports:
      - "9003:9003"
    volumes:
      - /your/export-data:/output

  owui-mcpo:
    image: ghcr.io/glissemantv/owui-mcpo:dev-latest
    container_name: owui-mcpo
    environment:
      - FILE_EXPORT_BASE_URL=http://file-export-server:9003/files
      - FILE_EXPORT_DIR=/output
      - MCPO_API_KEY=top-secret
      - PERSISTENT_FILES=true
      - FILES_DELAY=1
      - LOG_LEVEL=DEBUG
      - UNSPLASH_ACCESS_KEY=top-secret
      - IMAGE_SOURCE=local_sd
      - LOCAL_SD_URL=http://localhost:7860
      - LOCAL_SD_USERNAME=local_user
      - LOCAL_SD_PASSWORD=local_password
      - LOCAL_SD_DEFAULT_MODEL=sd_xl_base_1.0.safetensors
      - LOCAL_SD_STEPS=20
      - LOCAL_SD_WIDTH=512
      - LOCAL_SD_HEIGHT=512
      - LOCAL_SD_CFG_SCALE=1.5
      - LOCAL_SD_SCHEDULER=Karras
      - LOCAL_SD_SAMPLE=Euler a
    ports:
      - "8000:8000"
    restart: unless-stopped
    volumes:
      - /your/export-data:/output
    depends_on:
      - file-export-server
```
---
