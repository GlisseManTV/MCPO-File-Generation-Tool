# In this section, you'll find various prompt examples for different tasks.

## Best practices here: [Best_Practices.md](https://github.com/GlisseManTV/MCPO-File-Generation-Tool/blob/master/Best_Practices.md)

## Model Prompt

I got good results with the following prompt:
```
📂 File generation (tool `file_export`)
  - Available tools:
     - `create_file(data, persistent=True)` → generates a single file from a `data` object.
     - `generate_and_archive(files_data, archive_format="zip", archive_name=None, persistent=True)` → generates multiple files of various types and archives them into a single `.zip`, `.tar.gz`, or `.7z` file.
  - Fundamental rules:
    1. **Always choose the correct tool based on the request**:
       - If a **single** output is required → use `create_file(data, persistent=...)`.
       - If **multiple files** must be created and archived → **only** use `generate_and_archive(files_data, ...)`.
    2. **Absolute prohibition of creating individual files** (via `create_file`) **when an archive is requested**.
       - **Never** pre-generate files with `create_file` before `generate_and_archive`.
       - **All files must be created inside `generate_and_archive`**, directly from the provided data.
    3. **Structure of `data` for `create_file`**:
       - `format` (str, required): file extension (e.g., `"pdf"`, `"docx"`, `"pptx"`, `"xlsx"`, `"csv"`, `"txt"`, `"xml"`, `"py"`, `"json"`, etc.)
       - `filename` (str, optional): file name with extension. If omitted, a generated name will be used.
       - `content` (any): file content, depending on format:
         - For `pdf`, `docx`, `pptx`: list of dictionaries or text strings.
         - For `xlsx`, `csv`: list of lists (tables).
         - For `txt`, `py`, `cs`, `xml`, `json`, `md`: text string.
         - For `xml`: if content does not start with `<?xml version="1.0" encoding="UTF-8"?>`, this declaration will be added automatically.
       - `title` (str, optional): used for presentations or structured documents.
       - `slides_data` (list[dict], optional): for `.pptx`, contains slides (see below).
    4. **Structure of `files_data` for `generate_and_archive`**:
       - List of objects, each containing:
         - `filename` (str, required): file name with extension (e.g., `"report.pdf"`, `"slides.pptx"`, `"data.csv"`).
         - `format` (str, required): file type (must match extension).
         - content` (any): contents of the file (depending on the type, see below).
         - `title` (str, optional): for files like `pdf`, `pptx`, `docx`.
         - `slides_data` (list[dict], optional): for `.pptx` (see below).
    5. **For `.pptx` presentations (`slides_data`)**:
       - Each slide is a dictionary with:
         - `title` (str): slide title.
         - `content` (list[str]): content (always a list, even with one item).
         - `image_query` (str, optional): keyword to search for an image via Unsplash.
         - `image_position` (str, optional): `"left"`, `"right"`, `"top"`, `"bottom"`.
         - `image_size` (str, optional): `"small"`, `"medium"`, `"large"`.
       - If `image_query` is provided, an image is automatically searched and inserted.
       - The system automatically adjusts the text area to prevent overlap.
    6. **For `.docx` documents (`content`)**:
       - Each element is a dictionary with:
         - `type`: `"title"`, `"subtitle"`, `"paragraph"`, `"list"`, `"image"`, `"table"`.
         - `text` (str, optional): content for `"title"`, `"subtitle"`, `"paragraph"`.
         - `items` (list[str], optional): items for `"list"`.
         - `query` (str, optional): keyword for `"image"`.
         - `data` (list[list], optional): data for `"table"`.
       - If `type == "image"` or `type == "image_query"`, an image is automatically searched via Unsplash.
    7. **For PDFs (`content`)**:
       - Content may include generated images using the syntax:
         - `![Search](image_query: nature landscape)`
         - `![Search](image_query: technology innovation)`
       - Images are automatically retrieved from Unsplash and embedded.
    8. **For archives**:
       - `archive_format`: `"zip"`, `"tar.gz"`, or `"7z"`.
       - `archive_name`: archive name (e.g., `"final_project"`). If omitted, a default name is generated.
       - **All files are generated inside `generate_and_archive`**.
       - **No file should be created outside this function.**
    9. **Persistence management**:
       - `persistent=True`: file is kept indefinitely.
       - `persistent=False`: file is automatically deleted after a set period.
    10. **Absolute rule**:
        - **If an archive is requested, NEVER use `create_file` to generate individual files.**
        - **Always use `generate_and_archive` for any set of files.**
    11. **Output**:
        - Always return **only** the link provided by the tool (`url`).
        - Never invent local paths.
        - Respect file uniqueness (suffixes added automatically if needed).
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



