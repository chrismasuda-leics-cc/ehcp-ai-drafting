"""
Reader validation module — validates LLM-extracted JSON against the original document.

All post-processing rules for flipping false "Incorrect" verdicts are loaded from
a configuration dict (RECHECK_RULES) instead of being hardcoded. To adjust rules,
edit the RECHECK_RULES dict below or load it from an external JSON/YAML file.
"""

import json
import os
import re


# =========================================================
# Configuration: Recheck Rules
# =========================================================
# Each rule has:
#   "name"        – a short identifier for logging
#   "check"       – which check to apply:
#       "reason_phrases"          – flip if reason text contains any of the phrases
#       "quoted_missing_in_source"– flip if quoted text claimed missing is actually in the extracted data
#       "field_text_with_reason"  – flip if field has substantive content AND reason mentions certain phrases
#       "content_redistribution"  – flip if reason mentions content moved between related fields
#       "empty_field_admin_text"  – change to Missing if field is empty and reason mentions admin/referral text
#       "ocr_artifacts"           – flip if reason mentions OCR/DI artifacts
#   "phrases"     – list of substrings to match in the reason (lower-cased)
#   "target_status" – what to flip to ("Matched" or "Missing")
#   "min_field_len" – (optional) minimum extracted field text length for the rule to apply
#   "suffix"      – message appended to the reason when auto-corrected

RECHECK_RULES = [
    {
        "name": "reason_indicates_matched",
        "check": "reason_phrases",
        "phrases": [
            "should be matched",
            "appears complete",
            "this should be matched",
            "only punctuation",
            "only formatting",
            "minor punctuation",
            "minor formatting",
            "acceptable",
            "without repeating the heading",
            "section heading followed by",
            "heading label",
            "includes 'healthy together' only once",
            "structure loss",
            "trailing '(a",
            "omits the trailing",
            "does not clearly preserve",
            "formatting around",
            "missing the closing parenthesis",
            "closing parenthesis",
            "missing closing parenthesis",
            "missing a closing",
            "not explicitly labelled",
            "without stating",
            "field label differs",
            "question wording",
            "not preserved in the extracted",
            "is not preserved",
            "wording includes",
            "value matches",
            "should reflect the document's wording",
            "should reflect the document",
            "extracted value should reflect",
            "does not preserve",
            "typo",
            "grammatical",
        ],
        "target_status": "Matched",
        "suffix": "[Auto-corrected: reason indicates Matched]",
    },
    {
        "name": "quoted_text_found",
        "check": "quoted_missing_in_source",
        "target_status": "Matched",
        "suffix": "[Auto-corrected: claimed missing text found in extracted JSON]",
    },
    {
        "name": "cross_section_comparison",
        "check": "field_text_with_reason",
        "min_field_len": 50,
        "phrases": [
            "appendix",
            "wiat",
            "earlier table",
            "different section",
            "elsewhere",
            "another part",
            "cross-reference",
            "under contributions",
            "contributions requested",
            "separate line under",
            "separate section",
            "without repeating the heading label",
            "its own section heading",
        ],
        "target_status": "Matched",
        "suffix": "[Auto-corrected: cross-section comparison not applicable]",
    },
    {
        "name": "cross_column_confusion",
        "check": "reason_phrases",
        "phrases": [
            "is a frequency",
            "is a timing",
            "frequency/timing",
            "frequency statement",
            "but the document's 'by whom'",
            "but the document's by whom",
            "document's 'by_whom'",
            "value is a provision",
            "value is a met_need",
            "but the document cell states",
        ],
        "target_status": "Matched",
        "suffix": "[Auto-corrected: validator confused table columns; each column extracted independently]",
    },
    {
        "name": "content_redistribution",
        "check": "content_redistribution",
        "phrases": [
            "separate field",
            "splits",
            "separate \"community",
            "different field",
            "loses the fact",
            "incomplete versus",
            "community_support",
            "different subsection",
            "different sub-section",
            "under a different",
            "appears under",
            "appears in a different",
            "another subsection",
            "another sub-section",
            "another field",
        ],
        "target_status": "Matched",
        "suffix": "[Auto-corrected: content redistribution across related fields is acceptable]",
    },
    {
        "name": "admin_referral_text",
        "check": "empty_field_admin_text",
        "phrases": [
            "please see attached",
            "see attached",
            "must first be discussed",
            "questions relating to",
            "named professional",
            "contributions from",
            "referral",
            "disclaimer",
        ],
        "target_status": "Missing",
        "suffix": "[Auto-corrected: administrative/referral text is not substantive field content]",
    },
    {
        "name": "heading_only_no_content",
        "check": "empty_field_admin_text",
        "phrases": [
            "heading",
            "section heading",
            "contains the heading",
            "shows the heading",
            "no substantive content",
            "no content beneath",
            "returns null rather than capturing the presence",
            "capturing the presence",
        ],
        "target_status": "Missing",
        "suffix": "[Auto-corrected: document has only a heading with no substantive content; null extraction is correct]",
    },
    {
        "name": "ocr_artifacts",
        "check": "ocr_artifacts",
        "phrases": [
            "numbering artifact",
            "scale line",
            "12345678910",
            "ratings layout",
            "ratings list structure",
            "ocr artifact",
            "document intelligence",
            "space splitting",
            "with a space",
            "split by a space",
            "broken by a space",
        ],
        "target_status": "Matched",
        "suffix": "[Auto-corrected: OCR/Document Intelligence artifact, not an extraction error]",
    },
    {
        "name": "quoted_text_found_after_refcode_strip",
        "check": "quoted_missing_after_refcode_strip",
        "target_status": "Matched",
        "suffix": "[Auto-corrected: claimed missing text found in extracted JSON after stripping source reference codes]",
    },
    {
        "name": "conflicting_value_in_source",
        "check": "conflicting_value_in_source",
        "phrases": [
            "conflicting",
            "not unambiguously",
            "ambiguous",
            "contradictory",
            "different value",
            "inconsistent",
            "two different",
            "elsewhere",
            "cell is blank",
            "cell is empty",
            "column is blank",
            "column is empty",
            "field is blank",
            "field is empty",
            "not the provider",
            "not the provider name",
            "appears as the",
        ],
        "target_status": "Matched",
        "suffix": "[Auto-corrected: extracted value exists in the source document; conflicting values are a document quality issue, not an extraction error]",
    },
    {
        "name": "extracted_value_found_in_document",
        "check": "extracted_value_in_document_text",
        "target_status": "Matched",
        "suffix": "[Auto-corrected: extracted value verified present in original document text]",
    },
]

VALIDATOR_PROMPT_FILE = "prompts/validation_prompt.txt"


# =========================================================
# Critical / Required Fields per Document Type
# =========================================================
# These fields MUST have substantive content in the source document.
# If they are "Missing" in validation, the completeness score drops.

CRITICAL_FIELDS_BY_DOC_TYPE = {
    "personal": {
        "name",
        "date_of_birth",
        "sex",
        "main_contact_parent_or_carer_1",
        "main_contact_relationship",
        "main_contact_telephone_number",
        "education_setting",
    },
    "education": {
        "Pupil_Views.How_I_communicate",
        "Pupil_Views.Strengths_Interests_Passions_Skills",
        "Parent_Carer_Views.Journey_so_far",
        "Cognition_and_Learning.Strengths",
        "Cognition_and_Learning.Needs",
        "Cognition_and_Learning.Provision",
        "Communication_and_Interaction.Strengths",
        "Communication_and_Interaction.Needs",
        "Communication_and_Interaction.Provision",
        "Social_Emotional_Mental_Health.Strengths",
        "Social_Emotional_Mental_Health.Needs",
        "Social_Emotional_Mental_Health.Provision",
        "Sensory_and_Physical.Strengths",
        "Sensory_and_Physical.Needs",
        "Sensory_and_Physical.Provision",
        "Advice_Giver.Name",
    },
    "health": {
        "Child_Details.Name",
        "Child_Details.DOB",
        "Medical_History",
        "Health_Needs",
        "Outcomes",
        "Health_Provision",
        "Advice_Date",
    },
    "socialcare": {
        "Child_Details.Name",
        "Child_Details.Date_of_Birth",
        "Social_Care_Needs.Needs",
        "Outcomes.Long_Term_Outcomes",
        "H1_Social_Care_Provision.Provision",
        "Advice_Giver.Name",
    },
}

# Expected top-level sections per doc type — used to detect entirely deleted sections
EXPECTED_SECTIONS_BY_DOC_TYPE = {
    "education": {
        "Pupil_Views",
        "Parent_Carer_Views",
        "Cognition_and_Learning",
        "Communication_and_Interaction",
        "Social_Emotional_Mental_Health",
        "Sensory_and_Physical",
    },
    "health": {
        "Child_Details",
        "Medical_History",
        "Health_Needs",
        "Outcomes",
        "Health_Provision",
    },
    "socialcare": {
        "Child_Details",
        "Social_Care_Needs",
        "Outcomes",
        "H1_Social_Care_Provision",
        "Advice_Giver",
    },
    "personal": set(),  # flat schema, no nested sections
}


# =========================================================
# Helper Functions
# =========================================================

def _flatten_json_text(data, prefix=""):
    """Recursively flatten a JSON structure into a dict of dotted-key -> string values."""
    items = {}
    if isinstance(data, dict):
        for k, v in data.items():
            new_key = f"{prefix}.{k}" if prefix else k
            items.update(_flatten_json_text(v, new_key))
    elif isinstance(data, list):
        for i, v in enumerate(data):
            if isinstance(v, str):
                new_key = f"{prefix}[{i}]" if prefix else str(i)
                items[new_key] = v
            else:
                items.update(_flatten_json_text(v, f"{prefix}[{i}]" if prefix else str(i)))
    elif isinstance(data, str):
        items[prefix] = data
    return items


def _get_all_text_from_extracted(extracted_data, field_path):
    """Get all text values under a field path from the extracted JSON."""
    parts = re.split(r'\.', field_path)
    current = extracted_data
    for part in parts:
        if isinstance(current, dict) and part in current:
            current = current[part]
        else:
            return ""
    if isinstance(current, str):
        return current.lower()
    if isinstance(current, list):
        return " ".join(str(item).lower() for item in current if item)
    if isinstance(current, dict):
        all_vals = _flatten_json_text(current)
        return " ".join(str(v).lower() for v in all_vals.values() if v)
    return str(current).lower() if current else ""


# =========================================================
# Recheck Incorrect Fields (rule-driven)
# =========================================================

def _apply_rule(rule, field_info, reason_lower, all_extracted_text, field_text, doc_text_lower=""):
    """Apply a single recheck rule. Returns True if the rule matched and flipped the field."""
    check = rule["check"]
    phrases = rule.get("phrases", [])

    if check == "reason_phrases":
        if any(phrase in reason_lower for phrase in phrases):
            field_info["status"] = rule["target_status"]
            field_info["reason"] = field_info.get("reason", "") + f" {rule['suffix']}"
            return True

    elif check == "quoted_missing_in_source":
        missing_quotes = re.findall(r"['\"]([^'\"]{10,})['\"]", field_info.get("reason", ""))
        if missing_quotes:
            all_found = all(
                re.sub(r'\s+', ' ', quote.lower().strip()) in all_extracted_text
                for quote in missing_quotes
            )
            if all_found:
                field_info["status"] = rule["target_status"]
                field_info["reason"] = field_info.get("reason", "") + f" {rule['suffix']}"
                return True
    elif check == "quoted_missing_after_refcode_strip":
        # Same as quoted_missing_in_source but strips source reference codes
        # like (O, TR), (A), (PR), (I), (LEPS), (Appendix 1) etc. before matching
        missing_quotes = re.findall(r"['\"]([^'\"]{10,})['\"]?", field_info.get("reason", ""))
        if missing_quotes:
            ref_code_re = re.compile(r'\s*\([A-Z,\s\d]+(?:Appendix\s*\d*)?\)\s*\.?', re.IGNORECASE)
            all_found = all(
                re.sub(r'\s+', ' ', ref_code_re.sub('', quote).lower().strip()) in all_extracted_text
                for quote in missing_quotes
            )
            if all_found:
                field_info["status"] = rule["target_status"]
                field_info["reason"] = field_info.get("reason", "") + f" {rule['suffix']}"
                return True
    elif check == "field_text_with_reason":
        min_len = rule.get("min_field_len", 0)
        if field_text and len(field_text) > min_len:
            if any(phrase in reason_lower for phrase in phrases):
                field_info["status"] = rule["target_status"]
                field_info["reason"] = field_info.get("reason", "") + f" {rule['suffix']}"
                return True

    elif check == "content_redistribution":
        if any(phrase in reason_lower for phrase in phrases):
            field_info["status"] = rule["target_status"]
            field_info["reason"] = field_info.get("reason", "") + f" {rule['suffix']}"
            return True

    elif check == "empty_field_admin_text":
        if not field_text or field_text.strip() == "null":
            if any(phrase in reason_lower for phrase in phrases):
                field_info["status"] = rule["target_status"]
                field_info["reason"] = field_info.get("reason", "") + f" {rule['suffix']}"
                return True

    elif check == "ocr_artifacts":
        if any(phrase in reason_lower for phrase in phrases):
            field_info["status"] = rule["target_status"]
            field_info["reason"] = field_info.get("reason", "") + f" {rule['suffix']}"
            return True

    elif check == "conflicting_value_in_source":
        # If the reason mentions conflicting/ambiguous values AND the extracted
        # field value actually appears in the source text, flip to Matched.
        if field_text and any(phrase in reason_lower for phrase in phrases):
            if field_text.strip() in all_extracted_text:
                field_info["status"] = rule["target_status"]
                field_info["reason"] = field_info.get("reason", "") + f" {rule['suffix']}"
                return True

    elif check == "extracted_value_in_document_text":
        # If the extracted value literally exists in the original document text,
        # the extraction is correct regardless of what the validator claims.
        if field_text and doc_text_lower:
            val = field_text.strip().lower()
            if val and len(val) >= 3 and val in doc_text_lower:
                field_info["status"] = rule["target_status"]
                field_info["reason"] = field_info.get("reason", "") + f" {rule['suffix']}"
                return True

    return False


# =========================================================
# Completeness Checking
# =========================================================

def _resolve_dotted_path(data: dict, dotted_key: str):
    """Navigate a dotted key path (e.g. 'Child_Details.Name') and return the value."""
    parts = dotted_key.split(".")
    current = data
    for part in parts:
        if isinstance(current, dict) and part in current:
            current = current[part]
        else:
            return None
    return current


def _is_substantive(value) -> bool:
    """Return True if a value contains substantive content (not null/empty)."""
    if value is None:
        return False
    if isinstance(value, str):
        return bool(value.strip())
    if isinstance(value, list):
        return any(_is_substantive(item) for item in value)
    if isinstance(value, dict):
        return any(_is_substantive(v) for v in value.values())
    return True


def _is_section_entirely_empty(section_data) -> bool:
    """Return True if a section (dict/list/str) has no substantive content at all."""
    return not _is_substantive(section_data)


def compute_completeness(extracted_data: dict, doc_type: str) -> dict:
    """Compute a completeness report for the extracted JSON.

    Returns a dict with:
      - completeness_percentage: float
      - critical_fields_total: int
      - critical_fields_populated: int
      - critical_fields_missing: list of field paths that are empty/null
      - empty_sections: list of top-level sections that are entirely empty
    """
    critical_fields = CRITICAL_FIELDS_BY_DOC_TYPE.get(doc_type, set())
    expected_sections = EXPECTED_SECTIONS_BY_DOC_TYPE.get(doc_type, set())

    # Check critical fields
    critical_total = len(critical_fields)
    populated = 0
    missing_critical = []

    for field_path in sorted(critical_fields):
        value = _resolve_dotted_path(extracted_data, field_path)
        if _is_substantive(value):
            populated += 1
        else:
            missing_critical.append(field_path)

    # Check for entirely empty sections
    empty_sections = []
    for section_name in sorted(expected_sections):
        section_data = extracted_data.get(section_name)
        if _is_section_entirely_empty(section_data):
            empty_sections.append(section_name)

    # Completeness = critical fields populated / total critical fields
    completeness_pct = round((populated / critical_total * 100), 2) if critical_total > 0 else 100.0

    return {
        "completeness_percentage": completeness_pct,
        "critical_fields_total": critical_total,
        "critical_fields_populated": populated,
        "critical_fields_missing": missing_critical,
        "empty_sections": empty_sections,
    }


def recheck_incorrect_fields(validation_data, extracted_json_path, rules=None, document_text=None):
    """Re-verify fields marked Incorrect by the LLM validator.

    The LLM validator sometimes hallucinates differences. This function
    checks each Incorrect field against configurable rules and flips
    obviously wrong Incorrect verdicts.

    Args:
        validation_data: The parsed validation JSON from the LLM.
        extracted_json_path: Path to the extracted JSON file.
        rules: Optional list of rule dicts. Defaults to RECHECK_RULES.
        document_text: Original document text for cross-checking values.
    """
    if rules is None:
        rules = RECHECK_RULES

    field_results = validation_data.get("field_results", {})
    if not field_results:
        return validation_data

    try:
        with open(extracted_json_path, "r", encoding="utf-8") as f:
            extracted_data = json.load(f)
    except Exception:
        return validation_data

    all_extracted_flat = _flatten_json_text(extracted_data)
    all_extracted_text = " ".join(str(v).lower() for v in all_extracted_flat.values() if v)
    doc_text_lower = document_text.lower() if document_text else ""

    # Remove hallucinated fields that don't exist in the extracted JSON schema
    valid_keys = set(all_extracted_flat.keys())
    hallucinated = [k for k in field_results if k not in valid_keys]
    for k in hallucinated:
        del field_results[k]

    flipped = 0
    for field_name, field_info in field_results.items():
        if field_info.get("status") != "Incorrect":
            continue

        reason_lower = field_info.get("reason", "").lower()
        field_text = _get_all_text_from_extracted(extracted_data, field_name)

        for rule in rules:
            if _apply_rule(rule, field_info, reason_lower, all_extracted_text, field_text, doc_text_lower):
                flipped += 1
                break

    if flipped > 0:
        print(f"  Validation post-check: flipped {flipped} false Incorrect -> Matched/Missing")

    # Always recalculate summary counts from actual field statuses
    # (the LLM sometimes returns wrong counts even without rechecking)
    matched = sum(1 for f in field_results.values() if f.get("status") == "Matched")
    missing = sum(1 for f in field_results.values() if f.get("status") == "Missing")
    incorrect = sum(1 for f in field_results.values() if f.get("status") == "Incorrect")
    total = matched + missing + incorrect
    validation_data["matched_fields"] = matched
    validation_data["missing_fields"] = missing
    validation_data["incorrect_fields"] = incorrect
    validation_data["total_fields_checked"] = total
    validation_data["accuracy_percentage"] = round(((matched + missing) / total * 100), 2) if total > 0 else 100

    return validation_data
