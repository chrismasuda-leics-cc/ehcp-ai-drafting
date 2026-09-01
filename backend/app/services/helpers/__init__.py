"""
EHCP Agent Framework — Internal Helpers

Re-exports from implementation modules for convenient imports.
"""

# --- Reader / Extractor / Validator helpers ---
from app.services.helpers.reader_helpers import (
    build_document_text,
    extract_document,
    load_text_file,
    safe_parse_json,
    run_llm_extraction,
    start_token_tracking,
    get_token_usage,
)

# --- Quality Checker helpers ---
from app.services.helpers.validation_helpers import (
    recheck_incorrect_fields,
    RECHECK_RULES,
    VALIDATOR_PROMPT_FILE,
    compute_completeness,
    CRITICAL_FIELDS_BY_DOC_TYPE,
    EXPECTED_SECTIONS_BY_DOC_TYPE,
)

# --- Writer helper ---
from app.services.helpers.template_filler import fill_template_streamlit

# --- Writer Validator helpers ---
from app.services.helpers.writer_validation import (
    ValidationReport,
    validate_docx,
    load_json_file,
    validate_personal_json,
    validate_education_json,
    validate_socialcare_json,
    validate_health_json,
    load_workbook_safe,
    validate_workbook_headers,
    validate_personal_mapping,
    validate_section_mapping,
    compare_docx_outputs,
    build_json_report,
    EDUCATION_DETAILS_SHEET,
    SOCIAL_CARE_DETAILS_SHEET,
    HEALTH_DETAILS_SHEET,
)
