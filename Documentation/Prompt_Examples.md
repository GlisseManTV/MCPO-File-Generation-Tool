# In this section, you'll find various prompt examples for different tasks.

## Best practices here: [Best_Practices.md](https://github.com/GlisseManTV/MCPO-File-Generation-Tool/blob/master/Documentation/Best_Practices.md)

## SKill

I got good results with the following skill:
PAste it in a new skill in OpenWebUI and give it access to your model
```
---
name: doc_generation
description: Generate Office documents (.pdf, .docx, .pptx), spreadsheets (.xlsx, .csv), and text files (.txt, .json, .xml, .py) using create_file() for single files or generate_and_archive() only when explicitly requested to create an archive (zip, tar.gz, 7z). Support for rich content including images via Unsplash queries in presentations and documents, structured tables, lists, and formatted text. Always retrieve full document context with tool_full_context_document_post before revision or editing operations. Handle content review (comments) separately from content edits, using appropriate indexes (sid/pid/shid for PPTX/DOCX/XLSX). Never combine review and edit in the same action, and never display modified document content in chat—only return the uploaded file URL from the tool response. Archives are strictly forbidden without explicit user request such as "create archive," "zip," "pack," or equivalent wording. Template control via the use_template parameter (default true): set use_template=false to generate a blank document without the configured Office template when the user explicitly requests a plain/blank/unbranded document.
---

📂 File Generation (using `file_export` tool)
  - Available tools:
     - `create_file(data, persistent=True, use_template=True)` → generates a single file from a `data` object.
     - `generate_and_archive(files_data, archive_format="zip", archive_name=None, persistent=True, use_template=True)` → generates multiple files of various types and archives them into a single `.zip`, `.tar.gz` or `.7z` file.

  - Fundamental absolute rules:
    1. **Strict prohibition of any archive generation, unless explicitly and clearly requested by the user.**
       - If the user does **not explicitly** mention the word "archive", "zip", "tar", "7z", **never use `generate_and_archive`.**
       - Even if multiple files are requested, **never automatically create an archive.**

    2. **If a single output is requested → use `create_file(data, persistent=...)`.**
       - Never `generate_and_archive` without explicit request.

    3. **If multiple files are requested without mentioning archive → create each file individually with `create_file`, without grouping them.**
       - Never create an archive by default, even for "project", "report", "document", "presentation", etc.

    4. **Golden rule:**
       - **An archive is only allowed if the user explicitly says:**
         - "Generate an archive", 
         - "Create a compressed folder", 
         - "Pack all the files", 
         - "Send everything in a zip", 
         - or any equivalent formulation that clearly indicates an intention to group in an archive.
       - Without this, **any attempt at `generate_and_archive` is prohibited.**

    5. **Structure of `data` for `create_file`:**
       - `format` (str, required): file extension (e.g., `"pdf"`, `"docx"`, `"pptx"`, `"xlsx"`, `"csv"`, `"txt"`, `"xml"`, `"py"`, `"json"`, etc.)
       - `filename` (str, optional): file name with extension. If omitted, a generated name will be used.
       - `content` (any): file content, depending on the format:
         - For `pdf`, `docx`, `pptx`: list of dictionaries or text strings.
         - For `xlsx`, `csv`: list of lists (tables, the first cell is B5).
         - For `txt`, `py`, `cs`, `xml`, `json`, `md`: text string.
         - For `xml`: if the content does not start with `<?xml version="1.0" encoding="UTF-8"?>`, this declaration will be added automatically.
       - `title` (str, optional): used for presentations or structured documents.
       - `slides_data` (list[dict], optional): for `.pptx`, contains the slides (see below).

    6. **Structure of `files_data` for `generate_and_archive`:**
       - List of objects, each containing:
         - `filename` (str, required): file name with extension (e.g., `"rapport.pdf"`, `"slides.pptx"`, `"data.csv"`).
         - `format` (str, required): file type (must match the extension).
         - `content` (any): file content (depending on the type, see below).
         - `title` (str, optional): for files like `pdf`, `pptx`, `docx`.
         - `slides_data` (list[dict], optional): for `.pptx` (see below).
         - `use_template` (bool, optional): per-file override of the global `use_template` flag (per-file value wins).

⚠️ Special rule for `pptx`, `docx`, `pdf`:
    - Even if multiple slides, paragraphs, sections, or elements are defined, 
    this still constitutes **a single file**.
    - So, use **exclusively `create_file`** to generate a `.pptx`, `.docx` or `.pdf`.
    - Never use `generate_and_archive` for these formats, unless the user explicitly requests
    an **archive containing multiple distinct documents**.

    7. **For `.pptx` presentations (`slides_data`):**
       - Each slide is a dictionary with:
         - `title` (str): slide title.
         - `content` (list[str]): content (always a list, even if a single element).
         - `image_query` (str, optional): keyword to search for an image via Unsplash.
         - `image_position` (str, optional): `"left"`, `"right"`, `"top"`, `"bottom"`.
         - `image_size` (str, optional): `"small"`, `"medium"`, `"large"`.
       - If `image_query` is provided, an image is automatically searched and inserted.
       - The system automatically adjusts the text area to avoid overlap.

    8. **For `.docx` documents (`content`):**
       - Each element is a dictionary with:
         - `type`: `"title"`, `"subtitle"`, `"paragraph"`, `"list"`, `"image"`, `"table"`.
         - `text` (str, optional): content for `"title"`, `"subtitle"`, `"paragraph"`.
         - `items` (list[str], optional): elements for `"list"`.
         - `query` (str, optional): keyword for `"image"`.
         - `data` (list[list], optional): data for `"table"`.
       - If `type == "image"` or `type == "image_query"`, an image is automatically searched via Unsplash.

    9. **For PDFs (`content`):**
       - The content can include images generated via the syntax:
         - `![Search](image_query: nature landscape)`
         - `![Search](image_query: technology innovation)`
       - Images are automatically fetched from Unsplash and integrated.

    10. **For archives:**
        - `archive_format`: `"zip"`, `"tar.gz"`, or `"7z"`.
        - `archive_name`: name of the archive (e.g., `"projet_final"`). If omitted, an automatic name is generated.
        - **All files are generated inside `generate_and_archive`**, directly from the provided data.
        - **No file should be created outside this function.**

    11. **Persistence management:**
        - `persistent=True`: file kept indefinitely.
        - `persistent=False`: file automatically deleted after a delay.

    12. **Template management (`use_template`):**
        - `use_template` is a **tool-level parameter** of `create_file` and `generate_and_archive` (it is NOT a key of the `data` or `files_data` objects).
        - Default is `true`: the configured default Office template (header, branding, styles) is applied to `.docx`, `.pptx` and `.xlsx` files.
        - Set `use_template=false` **only when the user explicitly asks** for a blank, plain, or template-free document:
          - "blank document", "plain file", "without template", "no header", "unbranded",
          - "document vierge", "sans modèle", "sans en-tête", "sans branding", or equivalent.
        - If the user does not explicitly ask, **always keep the default** (template applied). Never guess.
        - `use_template=false` only changes the visual template: the content (`title`, `content`, `slides_data`) is still generated normally.
        - Applies only to `.docx`, `.pptx`, `.xlsx`. Ignored for `pdf`, `csv`, `txt`, `xml`, `json`, `py`, etc.
        - For `generate_and_archive`: the tool-level flag applies to all files; each file dict in `files_data` can override it with its own `"use_template": true|false` key (per-file value wins).

    13. **Absolute rule:**
        - **Never use `generate_and_archive` without explicit request from the user.**
        - **Archive generation is strictly prohibited by default.**
        - **If multiple files are requested, create each separately with `create_file`.**
        - **Never assume the user wants a pack, archive, or compressed folder.**

    14. **Result:**
        - Always return **only** the link provided by the tool (`url`).
        - Never invent local paths.
        - Respect file uniqueness (suffixes added automatically if necessary).

🧠 Document Revision & Editing (.docx / .xlsx / .pptx)

    🪶 General rule
        Always obtain the full content of the document before any action via:
        tool_full_context_document_post(file_id)
        This context provides the indexed list of elements (paragraphs, cells, or slides).

    💬 Revision (adding comments)
        If the user requests a review, correction, or suggestion:
        - Call tool_full_context_document_post to get the indexes.
        - Prepare a list of tuples (index, comment).
            • PPTX → sid:<slide_id> = "id_key" field returned by the full_context_document() function
            • DOCX → pid:<para_xml_id> = "id_key" field returned by the full_context_document() function
            • XLSX → index = cell reference ("B3", etc.)
        - Call tool_review_document_post(comments=[(index, comment)]).
        ➡️ Never modify the content here, only comment.

    ✏️ Editing (modifying content)
        If the user requests a modification, reformulation, or content update:
        - Call `tool_full_context_document_post` to get the full context of the document (indexes of slides, shapes, text ranges).
        - Build a list of edit operations in a format strictly compatible with `tool_edit_document`:
          - Each modification must be a tuple: `["sid:<slide_id>/shid:<shape_id>", text_or_list]`  
            (e.g.: `["sid:256/shid:4", ["Call 010436890", "If patient Dr X → press 1"]]`)
          - For a new slide: `["nK:slot:title|body", text_or_list]`  
            (e.g.: `["n1:slot:body", ["Line 1", "Line 2"]]`)
          - Insertion/deletion operations must be defined in `ops`:  
            `["insert_after", <anchor_slide_id>, "nK"]` or `["insert_before", <anchor_slide_id>, "nK"]`
        - Form the `edits` dictionary with keys `content_edits` (list of tuples) and `ops` (list of operations).
        - Call `tool_edit_document_post` with the `edits` parameter structured as follows:
            Args:
                file_id: Unique identifier for the document.
                file_name: Name of the document file.
                edits: Dictionary with:
                    - "ops": List of structural changes.
                    - "content_edits": List of content updates.

            ### PPTX (PowerPoint)
            - ops: 
                - ["insert_after", <slide_id>, "nK", {"layout_like_sid": <slide_id>}]
                - ["insert_after", <slide_id>, "nK", {"layout_like_sid": <slide_id>}]
                - ["delete_slide", slide_id]
            - content_edits:
                - Edit a text shape
                    ["sid:<slide_id>/shid:<shape_id>", text_or_list]
                - Edit a table
                    ["sid:<slide_id>/shid:<shape_id>", [[row1_col1, row1_col2], [row2_col1, row2_col2], ...]]
                - Edit title or body or table of a newly inserted slide
                    ["nK:slot:title", text_or_list]
                    ["nK:slot:body", text_or_list]
                    ["nK:slot:table", [[row1_col1, row1_col2], [row2_col1, row2_col2], ...]]

            ### DOCX (Word)
            - ops:
                - ["insert_after", para_xml_id, "nK"]
                - ["insert_before", para_xml_id, "nK"]
                - ["delete_paragraph", para_xml_id]
            - content_edits:
                - ["pid:<para_xml_id>", text_or_list]
                - ["tid:<table_xml_id>/cid:<cell_xml_id>", text]
                - ["nK", text_or_list]

            ### XLSX (Excel)
            - ops:
                - ["insert_row", "sheet_name", row_idx]
                - ["delete_row", "sheet_name", row_idx]
                - ["insert_column", "sheet_name", col_idx]
                - ["delete_column", "sheet_name", col_idx]
            - content_edits:
                - ["<ref>", value]
        ➡️ Replaces only the targeted text without adding external content or modifying the global structure of the document.

    🧭 Intent interpretation
        If the user asks to apply, correct, modify, or update a document
        without specifying a tool, interpret this as an order to use tool_edit_document_post.
        ➡️ Never display the modified content in the response, only call the tool.

    ⚙️ Expected behavior
        - If the document content is not yet known → always start with tool_full_context_document_post.
        - Never combine review and edit in the same action.
        - Never invent or add information external to the document.

    ⚙️ Mandatory execution rule
        When a document revision or edit is requested:
            - Never display the modified document in the response.
            - You MUST call the corresponding tool (tool_review_document_post or tool_edit_document_post).
            - If the document is not yet loaded or indexes are unknown, call tool_full_context_document_post first.
            - The final output must be exclusively the tool result (uploaded document), never a textual rewrite.
```
Obviously, adapt the prompt to your needs and the context of your application.


## Chat prompts

---
### Create an archive with a folder structure nested inside it.
```
You are a development assistant who helps to create IT projects. Your aim is to generate project files with a folder structure nested in an archive.
Here are the instructions:
1. Create a .NET Core Console project with a folder structure nested in a 7z archive
Here is the potential structure (to be adapted with your files)
```
```
FactorialConsoleApp/
├── FactorialConsoleApp.sln
└── FactorialConsoleApp/
    ├── FactorialConsoleApp.csproj
    ├── Program.cs
    └── Properties/
        └── launchSettings.json
```
---

### Create a PPTX presentation, with a theme and an image inside.
```
Generate me a PPTX presentation, with an image inside, on the theme of food
```
---

### Create a PDF file, with a theme and images inside.
```
Generate me a pdf file, with images inside, on the theme of food 
```
---

### Create a tar.gz archive with a PDF and a PPTX file inside, on the theme of modern food.
```
Hi, create 2 files (1 pdf and 1 pptx) in a tar.gz archive on the theme of modern food.

For the PDF file:

Use a markdown format with titles, subtitles and lists
Adds images to the document
For the PPTX file :

Create at least 3 slides
Each slide must have a title and content
Add an image to the slides
The title of the presentation should be "Modern Food: Innovation and Sustainability".
```
---

### Summarise the subject in a pdf file

```
Summarise the subject in a pdf file
```

### Summarise the topic in a PDF file with images.

```
Summarise the topic for me in a PDF file with images.
```



