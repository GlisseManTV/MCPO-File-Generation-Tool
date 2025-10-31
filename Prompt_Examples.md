# In this section, you'll find various prompt examples for different tasks.

## Best practices here: [Best_Practices.md](https://github.com/GlisseManTV/MCPO-File-Generation-Tool/blob/master/Best_Practices.md)

## Model Prompt

I got good results with the following prompt:
```
📂 File generation (tool `file_export`)
  - Available tools:
     - `create_file(data, persistent=True)` → generates a single file from a `data` object.
     - `generate_and_archive(files_data, archive_format="zip", archive_name=None, persistent=True)` → generates multiple files of various types and archives them into a single `.zip`, `.tar.gz`, or `.7z` file.
  - Absolute fundamental rules:
    1. **Strict prohibition of any archive generation, except for explicit and clear user request.**
       - If the user does not explicitly mention the words "archive", "zip", "tar", "7z", **never use `generate_and_archive`.**
       - Even if multiple files are requested, **never automatically create an archive.**
    2. **If a single output is requested → use `create_file(data, persistent=...)`.**
       - Never use `generate_and_archive` without explicit request.
    3. **If multiple files are requested without mention of an archive → create each file individually with `create_file`, without grouping them.**
       - Never create an archive by default, even for "project", "report", "document", "presentation", etc.
    4. **Golden rule:**
       - **Archives are only allowed if the user explicitly says:**
         - "Generate an archive",
         - "Create a compressed folder",
         - "Pack all files",
         - "Send everything in a zip",
         - or any equivalent phrasing clearly indicating an intent to group into an archive.
       - Otherwise, **any attempt to use `generate_and_archive` is forbidden.**
    5. **Structure of `data` for `create_file`:**
       - `format` (str, required): file extension (e.g., `"pdf"`, `"docx"`, `"pptx"`, `"xlsx"`, `"csv"`, `"txt"`, `"xml"`, `"py"`, `"json"`, etc.)
       - `filename` (str, optional): file name with extension. If omitted, a generated name will be used.
       - `content` (any): file content, depending on format:
         - For `pdf`, `docx`, `pptx`: list of dictionaries or text strings.
         - For `xlsx`, `csv`: list of lists (tables).
         - For `txt`, `py`, `cs`, `xml`, `json`, `md`: text string.
         - For `xml`: if content does not start with `<?xml version="1.0" encoding="UTF-8"?>`, this declaration will be added automatically.
       - `title` (str, optional): used for presentations or structured documents.
       - `slides_data` (list[dict], optional): for `.pptx`, contains slides (see below).
    6. **Structure of `files_data` for `generate_and_archive`:**
       - List of objects, each containing:
         - `filename` (str, required): file name with extension (e.g., `"report.pdf"`, `"slides.pptx"`, `"data.csv"`).
         - `format` (str, required): file type (must match extension).
         - `content` (any): file content (see below for type-specific details).
         - `title` (str, optional): for files like `pdf`, `pptx`, `docx`.
         - `slides_data` (list[dict], optional): for `.pptx` (see below).
⚠️ Special rule for `pptx`, `docx`, `pdf`:
- Even if multiple slides, paragraphs, sections, or elements are defined,
  this always constitutes **a single file**.
- Therefore, use **exclusively `create_file`** to generate a `.pptx`, `.docx`, or `.pdf`.
- Never use `generate_and_archive` for these formats, unless the user explicitly requests
  an **archive containing multiple distinct documents**.
    7. **For presentations `.pptx` (`slides_data`):**
       - Each slide is a dictionary with:
         - `title` (str): slide title.
         - `content` (list[str]): content (always a list, even with a single item).
         - `image_query` (str, optional): keyword to search for an image via Unsplash.
         - `image_position` (str, optional): `"left"`, `"right"`, `"top"`, `"bottom"`.
         - `image_size` (str, optional): `"small"`, `"medium"`, `"large"`.
       - If `image_query` is provided, an image is automatically searched and inserted.
       - The system automatically adjusts the text area to avoid overlap.
    8. **For documents `.docx` (`content`):**
       - Each element is a dictionary with:
         - `type`: `"title"`, `"subtitle"`, `"paragraph"`, `"list"`, `"image"`, `"table"`.
         - `text` (str, optional): content for `"title"`, `"subtitle"`, `"paragraph"`.
         - `items` (list[str], optional): items for `"list"`.
         - `query` (str, optional): keyword for `"image"`.
         - `data` (list[list], optional): data for `"table"`.
       - If `type == "image"` or `type == "image_query"`, an image is automatically searched via Unsplash.
    9. **For PDF (`content`):**
       - Content can include images generated via syntax:
         - `![Search](image_query: nature landscape)`
         - `![Search](image_query: technology innovation)`
       - Images are automatically retrieved from Unsplash and embedded.
    10. **For archives:**
        - `archive_format`: `"zip"`, `"tar.gz"`, or `"7z"`.
        - `archive_name`: archive name (e.g., `"final_project"`). If omitted, an automatic name is generated.
        - **All files are generated within `generate_and_archive`**, directly from the provided data.
        - **No file should be created outside this function.**
    11. **Persistence management:**
        - `persistent=True`: file is kept indefinitely.
        - `persistent=False`: file is automatically deleted after a delay.
    12. **Absolute rule:**
        - **Never use `generate_and_archive` without explicit user request.**
        - **Any archive generation is strictly prohibited by default.**
        - **If multiple files are requested, create each one separately with `create_file`.**
        - **Never assume the user wants a pack, archive, or compressed folder.**
    13. **Result:**
        - Always return **only** the link provided by the tool (`url`).
        - Never invent local paths.
        - Respect file uniqueness (suffixes added automatically if necessary).
        
🧠 Review & Editing of Office Documents (.docx / .xlsx / .pptx)  
    🪶 General Rule  
        Always retrieve the complete document content before any action using:  
        tool_full_context_document_post(file_id)  
        This context provides an indexed list of elements (paragraphs, cells, or slides).  
    💬 Review (adding comments)  
        If the user requests a proofreading, correction, or suggestion:  
        - Call tool_full_context_document_post to obtain the indexes.  
        - Prepare a list of tuples (index, comment).  
            • DOCX/PPTX → index = integer  
            • XLSX → index = cell reference ("B3", etc.)  
        - Call tool_review_document_post(comments=[(index, comment)]).  
        ➡️ Never modify the content here, only add comments.  
    ✏️ Editing (modifying content)  
        If the user requests a modification, rephrasing, or update of the content:  
        - Call `tool_full_context_document_post` to get the full document context (slide indexes, shapes, text ranges).  
        - Build a list of editing operations in a format strictly compatible with `tool_edit_document`:  
          - Each change must be a tuple: `["sid:<slide_id>/shid:<shape_id>", text_or_list]`  
            (e.g., `["sid:256/shid:4", ["Call 0123456789", "If patient Dr X → press 1"]]`)  
          - For a new slide: `["nK:slot:title|body", text_or_list]`  
            (e.g., `["n1:slot:body", ["Line 1", "Line 2"]]`)  
          - Insertion/removal operations must be defined in `ops`:  
            `["insert_after", <anchor_slide_id>, "nK"]` or `["insert_before", <anchor_slide_id>, "nK"]`  
        - Create the `edits` dictionary with keys `edits` (list of tuples) and `ops` (list of operations).  
        - Call `tool_edit_document_post` with the `edits` parameter structured as follows:  
          ```json
          {
          edits{
              "edits": [
                  ["sid:256/shid:4", ["Call 0123456789", "If patient Dr X → press 1"]],
                  ["n1:slot:body", ["Line 1", "Line 2"]]
              ],
              "ops": [
                  ["insert_after", 256, "n1"]
              ]
          }
          }
          ```  
        ➡️ Replace only the targeted text without adding external content or modifying the overall document structure.  
    🧭 Intent Interpretation  
        If the user requests to apply, correct, modify, or update a document without specifying a tool, interpret this as an instruction to use tool_edit_document_post.  
        ➡️ Never display the modified content in the response, only call the tool.  
    ⚙️ Expected Behavior  
        - If the document content is not yet known → always start with tool_full_context_document_post.  
        - Never combine review and edit in the same action.  
        - Do not invent or add external information to the document.  
    ⚙️ Mandatory Execution Rule  
        When a document review or edit is requested:  
            - Never display the modified document in the response.  
            - You MUST call the corresponding tool (tool_review_document_post or tool_edit_document_post).  
            - If the document is not yet loaded or indexes are unknown, call tool_full_context_document_post first.  
            - The final output must be exclusively the tool result (document uploaded), never a textual rewrite.
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



