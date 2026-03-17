import os
import re
import logging
import requests
import markdown2
from io import BytesIO
from bs4 import BeautifulSoup, NavigableString
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, ListFlowable, ListItem, Image as ReportLabImage
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.enums import TA_LEFT
from reportlab.lib.units import mm
import emoji

from utils.file_treatment import (
    _generate_unique_folder,
    _generate_filename,
    _public_url,
    search_image,
)

log = logging.getLogger(__name__)

# Styles
styles = getSampleStyleSheet()
styles.add(ParagraphStyle(
    name="CustomHeading1",
    parent=styles["Heading1"],
    textColor=colors.HexColor("#0A1F44"),
    fontSize=18,
    spaceAfter=16,
    spaceBefore=12,
    alignment=TA_LEFT
))
styles.add(ParagraphStyle(
    name="CustomHeading2",
    parent=styles["Heading2"],
    textColor=colors.HexColor("#1C3F77"),
    fontSize=14,
    spaceAfter=12,
    spaceBefore=10,
    alignment=TA_LEFT
))
styles.add(ParagraphStyle(
    name="CustomHeading3",
    parent=styles["Heading3"],
    textColor=colors.HexColor("#3A6FB0"),
    fontSize=12,
    spaceAfter=10,
    spaceBefore=8,
    alignment=TA_LEFT
))
styles.add(ParagraphStyle(
    name="CustomNormal",
    parent=styles["Normal"],
    fontSize=11,
    leading=14,
    alignment=TA_LEFT
))
styles.add(ParagraphStyle(
    name="CustomListItem",
    parent=styles["Normal"],
    fontSize=11,
    leading=14,
    alignment=TA_LEFT
))
styles.add(ParagraphStyle(
    name="CustomCode",
    parent=styles["Code"],
    fontSize=10,
    leading=12,
    fontName="Courier",
    backColor=colors.HexColor("#F5F5F5"),
    borderColor=colors.HexColor("#CCCCCC"),
    borderWidth=1,
    leftIndent=10,
    rightIndent=10,
    topPadding=5,
    bottomPadding=5
))

def render_text_with_emojis(text: str) -> str:
    if not text:
        return ""
    try:
        return emoji.emojize(text, language="alias")
    except Exception as e:
        log.error(f"Error in emoji conversion: {e}")
        return text

def process_list_items(ul_or_ol_element, is_ordered=False):
    items = []
    for li in ul_or_ol_element.find_all('li', recursive=False):
        li_text_parts = []
        for content in li.contents:
            if isinstance(content, NavigableString):
                li_text_parts.append(str(content))
            elif getattr(content, "name", None) not in ['ul', 'ol']:
                li_text_parts.append(content.get_text())
        li_text = ''.join(li_text_parts).strip()
        list_item_paragraph = None
        if li_text:
            rendered_text = render_text_with_emojis(li_text)
            list_item_paragraph = Paragraph(rendered_text, styles["CustomListItem"])
        sub_lists = li.find_all(['ul', 'ol'], recursive=False)
        sub_flowables = []
        if list_item_paragraph:
            sub_flowables.append(list_item_paragraph)
        for sub_list in sub_lists:
            is_sub_ordered = sub_list.name == 'ol'
            nested_items = process_list_items(sub_list, is_sub_ordered)
            if nested_items:
                nested_list_flowable = ListFlowable(
                    nested_items,
                    bulletType='1' if is_sub_ordered else 'bullet',
                    leftIndent=10 * mm,
                    bulletIndent=5 * mm,
                    spaceBefore=2,
                    spaceAfter=2
                )
                sub_flowables.append(nested_list_flowable)
        if sub_flowables:
            items.append(ListItem(sub_flowables))
    return items

def render_html_elements(soup):
    log.debug("Starting render_html_elements...")
    story = []
    element_count = 0
    for elem in soup.children:
        element_count += 1
        if isinstance(elem, NavigableString):
            text = str(elem).strip()
            if text:
                story.append(Paragraph(render_text_with_emojis(text), styles["CustomNormal"]))
                story.append(Spacer(1, 6))
        elif hasattr(elem, 'name'):
            tag_name = elem.name
            if tag_name == "h1":
                text = render_text_with_emojis(elem.get_text().strip())
                story.append(Paragraph(text, styles["CustomHeading1"]))
                story.append(Spacer(1, 10))
            elif tag_name == "h2":
                text = render_text_with_emojis(elem.get_text().strip())
                story.append(Paragraph(text, styles["CustomHeading2"]))
                story.append(Spacer(1, 8))
            elif tag_name == "h3":
                text = render_text_with_emojis(elem.get_text().strip())
                story.append(Paragraph(text, styles["CustomHeading3"]))
                story.append(Spacer(1, 6))
            elif tag_name == "p":
                imgs = elem.find_all("img")
                if imgs:
                    for img_tag in imgs:
                        src = img_tag.get("src")
                        alt = img_tag.get("alt", "[Image]")
                        try:
                            if src and src.startswith("http"):
                                response = requests.get(src)
                                response.raise_for_status()
                                img_data = BytesIO(response.content)
                                img = ReportLabImage(img_data, width=200, height=150)
                            else:
                                img = ReportLabImage(src, width=200, height=150)
                            story.append(img)
                            story.append(Spacer(1, 10))
                        except Exception as e:
                            log.error(f"Error loading image {src}: {e}")
                            story.append(Paragraph(f"[Image: {alt}]", styles["CustomNormal"]))
                            story.append(Spacer(1, 6))
                else:
                    text = render_text_with_emojis(elem.get_text().strip())
                    if text:
                        story.append(Paragraph(text, styles["CustomNormal"]))
                        story.append(Spacer(1, 6))
            elif tag_name in ["ul", "ol"]:
                is_ordered = tag_name == "ol"
                items = process_list_items(elem, is_ordered)
                if items:
                    story.append(ListFlowable(
                        items,
                        bulletType='1' if is_ordered else 'bullet',
                        leftIndent=10 * mm,
                        bulletIndent=5 * mm,
                        spaceBefore=6,
                        spaceAfter=10
                    ))
            elif tag_name == "blockquote":
                text = render_text_with_emojis(elem.get_text().strip())
                if text:
                    story.append(Paragraph(f"{text}", styles["CustomNormal"]))
                    story.append(Spacer(1, 8))
            elif tag_name in ["code", "pre"]:
                text = elem.get_text().strip()
                if text:
                    story.append(Paragraph(text, styles["CustomCode"]))
                    story.append(Spacer(1, 6 if tag_name == "code" else 8))
            elif tag_name == "img":
                src = elem.get("src")
                alt = elem.get("alt", "[Image]")
                if src is not None:
                    try:
                        if src.startswith("image_query:"):
                            query = src.replace("image_query:", "").strip()
                            image_url = search_image(query)
                            if image_url:
                                response = requests.get(image_url)
                                response.raise_for_status()
                                img_data = BytesIO(response.content)
                                img = ReportLabImage(img_data, width=200, height=150)
                                story.append(img)
                                story.append(Spacer(1, 10))
                            else:
                                story.append(Paragraph(f"[Image non trouvée pour: {query}]", styles["CustomNormal"]))
                                story.append(Spacer(1, 6))
                        elif src.startswith("http"):
                            response = requests.get(src)
                            response.raise_for_status()
                            img_data = BytesIO(response.content)
                            img = ReportLabImage(img_data, width=200, height=150)
                            story.append(img)
                            story.append(Spacer(1, 10))
                        else:
                            if os.path.exists(src):
                                img = ReportLabImage(src, width=200, height=150)
                                story.append(img)
                                story.append(Spacer(1, 10))
                            else:
                                story.append(Paragraph(f"[Image locale non trouvée: {src}]", styles["CustomNormal"]))
                                story.append(Spacer(1, 6))
                    except requests.exceptions.RequestException as e:
                        log.error(f"Network error loading image {src}: {e}")
                        story.append(Paragraph(f"[Image (erreur réseau): {alt}]", styles["CustomNormal"]))
                        story.append(Spacer(1, 6))
                    except Exception as e:
                        log.error(f"Error processing image {src}: {e}", exc_info=True)
                        story.append(Paragraph(f"[Image: {alt}]", styles["CustomNormal"]))
                        story.append(Spacer(1, 6))
                else:
                    story.append(Paragraph(f"[Image: {alt} (source manquante)]", styles["CustomNormal"]))
                    story.append(Spacer(1, 6))
            elif tag_name == "br":
                story.append(Spacer(1, 6))
            else:
                text = elem.get_text().strip()
                if text:
                    story.append(Paragraph(render_text_with_emojis(text), styles["CustomNormal"]))
                    story.append(Spacer(1, 6))
    return story

def create_pdf(text: str | list[str], filename: str, folder_path: str | None = None) -> dict:
    """
    Build a PDF from markdown or structured list. Supports image_query: syntax.
    Returns: {"url": public_url, "path": local_path}
    """
    log.debug("Creating PDF file")
    if folder_path is None:
        folder_path = _generate_unique_folder()
    if filename:
        filepath = os.path.join(folder_path, filename)
        os.makedirs(os.path.dirname(filepath), exist_ok=True)
        fname = filename
    else:
        filepath, fname = _generate_filename(folder_path, "pdf")

    md_parts = []
    if isinstance(text, list):
        for item in text:
            if isinstance(item, str):
                md_parts.append(item)
            elif isinstance(item, dict):
                t = item.get("type")
                if t == "title":
                    md_parts.append(f"# {item.get('text','')}")
                elif t == "subtitle":
                    md_parts.append(f"## {item.get('text','')}")
                elif t == "paragraph":
                    md_parts.append(item.get("text",""))
                elif t == "list":
                    md_parts.append("\n".join([f"- {x}" for x in item.get("items",[])]))
                elif t in ("image","image_query"):
                    query = item.get("query","")
                    if query:
                        md_parts.append(f"![Image](image_query: {query})")
    else:
        md_parts = [str(text or "")]
        
    md_text = "\n\n".join(md_parts)

    def replace_image_query(match):
        query = match.group(1).strip()
        image_url = search_image(query)
        return f'\n\n<img src="{image_url}" alt="Image: {query}" />\n\n' if image_url else ""

    md_text = re.sub(r'!\[[^\]]*\]\(\s*image_query:\s*([^)]+)\)', replace_image_query, md_text)
    html = markdown2.markdown(md_text, extras=['fenced-code-blocks','tables','break-on-newline','cuddled-lists'])
    soup = BeautifulSoup(html, "html.parser")
    story = render_html_elements(soup) or [Paragraph("Empty Content", styles["CustomNormal"])]

    doc = SimpleDocTemplate(filepath, topMargin=72, bottomMargin=72, leftMargin=72, rightMargin=72)
    try:
        doc.build(story)
    except Exception as e:
        log.error(f"Error building PDF {fname}: {e}", exc_info=True)
        doc2 = SimpleDocTemplate(filepath)
        doc2.build([Paragraph("Error in PDF generation", styles["CustomNormal"])])

    return {"url": _public_url(folder_path, fname), "path": filepath}
