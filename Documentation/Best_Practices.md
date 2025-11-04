# 📂 File Generation & Archive Management System

🚀 A powerful, automated system to create, organize, and package files — from simple documents to full project archives — all in one go.

Whether you're building a project, generating reports, or archiving assets, this tool handles everything with precision, consistency, and elegance. Perfect for developers, content creators, and teams who value clean, reusable workflows.

---

## 🛠️ Available Tools

| Tool | Purpose | Format |
|------|--------|--------|
| `create_excel` | **Generate Excel files from arrays** | `.xlsx` |
| `create_csv` | **Create CSV files from structured data** | `.csv` |
| `create_pdf` | **Turn content into polished PDFs** | `.pdf` |
| `create_word` | **Build Word documents with rich formatting** | `.docx` |
| `create_file` | **Output any text-based file** (`.py`, `.html`, `.json`, `.xml`, `.md`, etc.) | Any |
| `create_presentation` | **Build dynamic presentations** | `.pptx` |
| `generate_and_archive` | **Pack multiple files into a single `.zip`, `.tar.gz`, or `.7z` archive** | `.zip`, `.tar.gz`, `.7z` |
| `edit_document` | **Edit existing files in-place** | `.docx`, `.xlsx`, `.pptx` |
| `review_document` | **Review and suggest improvements to existing files** | `.docx`, `.xlsx`, `.pptx`|

> 🔥 **Pro Tip**: Always use `generate_and_archive` when bundling multiple files — never mix individual creators!

---

## 📁 Example: .NET Console Project Structure

```bash
FactorialConsoleApp/
├── FactorialConsoleApp.sln
└── FactorialConsoleApp/
    ├── FactorialConsoleApp.csproj
    ├── Program.cs
    └── Properties/
        └── launchSettings.json
```

Clean, scalable, and ready for deployment. 🚀

---

## 🎨 Create a Presentation with Images

> 🎯 *"Generate me a PPTX presentation, with an image inside, on the theme of food."*

✅ Automatically:
- Creates 3+ slides with titles and content  
- Inserts relevant images from Unsplash  
- Applies a cohesive design theme  
- Delivers a polished `.pptx` file — ready to present! 🍽️✨

---

## 📄 Create a PDF with Visuals & Structure

> 📝 *"Generate me a PDF file, with images inside, on the theme of food."*

✅ Automatically:
- Uses Markdown formatting (headings, lists, sections)  
- Embeds high-quality images from Unsplash  
- Outputs a professional, readable document — ideal for reports or portfolios 📊🖼️

---

## 📦 Create a `tar.gz` Archive with Multiple Files

> 📌 *"Hi, create 2 files (1 PDF and 1 PPTX) in a tar.gz archive on the theme of modern food."*

### ✅ PDF Requirements:
- Markdown with titles, subtitles, and bullet points  
- Images (e.g., `image_query: modern food innovation`)  

### ✅ PPTX Requirements:
- At least 3 slides  
- Each with a title and descriptive content  
- An image on each slide  
- Title: **"Modern Food: Innovation and Sustainability"**

📦 The result? A single, compressed `.tar.gz` file — perfect for sharing or deployment. 🔗

---

## 📚 Summarise a Topic in a PDF

> 📝 *"Summarise the subject in a PDF file."*  
> 📝 *"Summarise the topic for me in a PDF file with images."*

✅ Automatically:
- Extracts key points  
- Structures content clearly  
- Adds visuals where relevant  
- Delivers a professional, ready-to-use document in minutes 📄✨

---

## ✅ Key Rules & Best Practices

- 🚫 **Never use** `create_file`, `create_csv`, etc. **before** `generate_and_archive` — it’s strictly forbidden.
- ✅ Always use `generate_and_archive` for **any** multi-file packaging.
- 🖼️ Image queries like `![Search](image_query: nature landscape)` are fetched automatically from Unsplash.
- 🔐 File persistence: `persistent=True` (keeps files forever) or `persistent=False` (auto-delete).
- 🔗 The **only** valid download link is provided by the tool — no fake paths!

---

## 🌟 Why This Tool?

- ✅ **Automated & Reliable** – No manual file handling.
- ✅ **Consistent & Scalable** – Works the same for 1 file or 100.
- ✅ **Visual & Professional** – Outputs polished, presentation-ready files.

---

## 📌 Get Started

Just define your files, choose the right tool, and let the system do the rest.  
Your project, your rules — all in one powerful workflow. 🚀

## Prompt examples here: [Prompt_Examples.md](https://github.com/GlisseManTV/MCPO-File-Generation-Tool/blob/master/Prompt_Examples.md)

---

📌 **Made with ❤️ for developers, creators, and teams who want to work smarter — not harder.**