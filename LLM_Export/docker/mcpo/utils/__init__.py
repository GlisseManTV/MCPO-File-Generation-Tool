"""
Initialize the utils package.
Expose the public API for document creation/edition helpers and file operations.
"""

# Document generators
from .pptx_treatment import (
    create_presentation,
    _add_table_from_matrix,
    _set_table_from_matrix,
    _set_text_with_runs,
    shape_by_id,
    ensure_slot_textbox,
    _layout_has,
    _pick_layout_for_slots,
    _collect_needs,
    _body_placeholder_bounds,
    _get_pptx_namespaces,
    _add_native_pptx_comment_zip,
)
from .docx_treatment import (
    create_word,
    _apply_text_to_paragraph,
    _apply_run_formatting,
    _extract_paragraph_style_info,
    _snapshot_runs,
    _apply_font,
)
from .xlsx_treatment import (
    create_excel,
    add_auto_sized_review_comment,
)
from .pdf_treatment import create_pdf

# File operations, storage, and search
from .file_treatment import (
    _create_csv,
    _create_raw_file,
    upload_file,
    download_file,
    search_image,
    search_local_sd,
    search_unsplash,
    search_pexels,
    _public_url,
    _generate_unique_folder,
    _generate_filename,
    _cleanup_files,
)

__all__ = [
    # Main creation APIs
    "create_presentation",
    "create_word",
    "create_excel",
    "create_pdf",

    # File ops + helpers
    "_create_csv",
    "_create_raw_file",
    "upload_file",
    "download_file",
    "search_image",
    "search_local_sd",
    "search_unsplash",
    "search_pexels",
    "_public_url",
    "_generate_unique_folder",
    "_generate_filename",
    "_cleanup_files",

    # PPTX helpers
    "_add_table_from_matrix",
    "_set_table_from_matrix",
    "_set_text_with_runs",
    "shape_by_id",
    "ensure_slot_textbox",
    "_layout_has",
    "_pick_layout_for_slots",
    "_collect_needs",
    "_body_placeholder_bounds",
    "_get_pptx_namespaces",
    "_add_native_pptx_comment_zip",

    # DOCX helpers
    "_apply_text_to_paragraph",
    "_apply_run_formatting",
    "_extract_paragraph_style_info",
    "_snapshot_runs",
    "_apply_font",

    # XLSX helper
    "add_auto_sized_review_comment",
]
