# MCPO-File-Generation-Tool Architecture

## Overview

This project is an MCP (Model Context Protocol) server for file generation from Open WebUI via two transport modes: stdio (integrated MCPO) and SSE/HTTP (enterprise mode).

---

## Main Architecture

```
┌─────────────────┐      ┌──────────────────────┐      ┌─────────────────┐
│  Open WebUI     │◄────►│  MCP Server          │◄────►│ File Export     │
│  (Client AI)    │      │  (MCPO or SSE/HTTP)  │      │  Server         │
└─────────────────┘      └──────────────────────┘      │  (Port 9003)    │
                                                        └─────────────────┘
```

### Two Transport Modes

| Mode | Transport | Usage | Default Port |
|------|-----------|-------|---------------|
| **MCPO** | stdio (Python module) | Beginners, simple setups | Configurable |
| **SSE/HTTP** | SSE or HTTP Streamable | Enterprise | 9004 |

---

## Key Components (sse_http)

### Entry Point: `tools/server.py` (v1.0.0-dev1)

MCP Server with the following tools:

| Tool | Description | Formats |
|------|-------------|---------|
| `create_file` | Create a single file | PDF, DOCX, PPTX, XLSX, CSV, TXT |
| `generate_and_archive` | Create and archive (ZIP/7Z/TAR.GZ) | Multi-format |
| `full_context_document` | Analyze document structure | DOCX, XLSX, PPTX |
| `edit_document` | Modify existing document | DOCX, XLSX, PPTX |
| `review_document` | Add comments/reviews | DOCX, XLSX, PPTX |

### Processing Modules (`utils/`)

- **file_treatment.py**: Common functions (public URLs, unique folder/file generation, OpenWebUI upload/download, image search, automatic cleanup)
- **docx_treatment.py**: Word creation/editing with templates and formatting preservation
- **xlsx_treatment.py**: Excel creation with auto-sizing columns and native comments
- **pptx_treatment.py**: PowerPoint creation with intelligent image positioning
- **pdf_treatment.py**: HTML→PDF rendering via ReportLab with enhanced Markdown support

---

## Communication Protocol (SSE/HTTP)

### Endpoints

```
GET  /sse        → Server-Sent Events stream (endpoint, ping, error events)
POST /messages   → JSON-RPC 2.0 messages for tool calls
GET  /health     → Health check
```

### Example JSON-RPC Message

```json
{
  "jsonrpc": "2.0",
  "id": 1,
  "method": "tools/call",
  "params": {
    "name": "create_file",
    "arguments": {"data": {...}}
  }
}
```

---

## Configuration

### Critical Environment Variables

```bash
# Transport
MODE=sse|http          # Transport mode
MCP_HTTP_PORT=9004     # MCP port
MCP_HTTP_HOST=0.0.0.0  # Listening host

# Files
FILE_EXPORT_DIR=/output              # Export directory (mount volume)
FILE_EXPORT_BASE_URL=http://server:9003/files
PERSISTENT_FILES=true|false
FILES_DELAY=60                       # Cleanup delay in minutes

# Images
IMAGE_SOURCE=unsplash|pexels|local_sd
UNSPLASH_ACCESS_KEY=your-key

# OpenWebUI
OWUI_URL=http://localhost:8000
JWT_SECRET=token
```

---

## Coding Convention

**IMPORTANT**: All code interactions must be in **ENGLISH**:
- Comments, logs, error messages
- Docstrings and tool descriptions
- Variables and functions (naming convention)

### File Structure
1. Standard Python imports
2. Third-party library imports
3. Local imports (`module.parent`)
4. Global constants
5. Classes/functions