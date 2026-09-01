"""
EHCP Agent Framework — Agent Definitions

Uses Microsoft Agent Framework (agent-framework) with OpenAIChatCompletionClient
for Azure OpenAI. Each agent has @tool-decorated functions and is created via
the Agent class.

Reader Pipeline agents:
    1. DocumentReaderAgent  — DI text extraction  → writes doctext file
    2. ExtractorAgent       — LLM JSON extraction → writes output JSON
    3. ValidatorAgent       — LLM validation      → writes validation JSON
    4. QualityCheckerAgent  — Rule-based recheck  → overwrites validation JSON

Writer Pipeline agents:
    5. TemplateWriterAgent      — DOCX template fill → writes filled DOCX
    6. WriterValidatorAgent     — Deterministic validation → writes report JSON
"""

import os
from dotenv import load_dotenv

from agent_framework import Agent, tool
from agent_framework.openai import OpenAIChatCompletionClient
from azure.identity.aio import DefaultAzureCredential

from app.services.helpers import (
    build_document_text,
    extract_document,
    load_text_file,
    safe_parse_json,
    run_llm_extraction,
    recheck_incorrect_fields,
    RECHECK_RULES,
    VALIDATOR_PROMPT_FILE,
    compute_completeness,
    fill_template_streamlit,
)
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

import json

load_dotenv()

from app.settings import (
    USE_MANAGED_IDENTITY,
    AZURE_OPENAI_ENDPOINT,
    AZURE_OPENAI_API_KEY,
    AZURE_OPENAI_DEPLOYMENT,
    AZURE_OPENAI_API_VERSION,
)


# =========================================================
# Azure OpenAI Chat Client (shared)
# =========================================================

def _create_client() -> OpenAIChatCompletionClient:
    """Create an OpenAIChatCompletionClient for Azure OpenAI."""
    if USE_MANAGED_IDENTITY:
        return OpenAIChatCompletionClient(
            model=AZURE_OPENAI_DEPLOYMENT,
            azure_endpoint=AZURE_OPENAI_ENDPOINT,
            credential=DefaultAzureCredential(),
            api_version=AZURE_OPENAI_API_VERSION,
        )
    return OpenAIChatCompletionClient(
        model=AZURE_OPENAI_DEPLOYMENT,
        azure_endpoint=AZURE_OPENAI_ENDPOINT,
        api_key=AZURE_OPENAI_API_KEY,
        api_version=AZURE_OPENAI_API_VERSION,
    )


# =========================================================
# Tool Functions
# =========================================================

@tool(approval_mode="never_require")
def read_document(file_path: str, output_text_path: str) -> str:
    """Extract text from a DOCX/PDF document using Azure Document Intelligence.
    Saves extracted text to output_text_path on disk."""
    print(f"  [ReaderAgent] Reading document: {file_path}")
    doc_text = build_document_text(file_path)
    os.makedirs(os.path.dirname(output_text_path) or ".", exist_ok=True)
    with open(output_text_path, "w", encoding="utf-8") as f:
        f.write(doc_text)
    print(f"  [ReaderAgent] Extracted {len(doc_text)} chars → {output_text_path}")
    return f"Document text extracted ({len(doc_text)} characters) and saved to {output_text_path}"


@tool(approval_mode="never_require")
async def extract_to_json(
    doctext_path: str, prompt_file: str, schema_file: str, output_json_path: str
) -> str:
    """Extract structured JSON from document text using LLM with the given
    prompt and schema. Reads text from doctext_path, writes JSON to output_json_path."""
    print(f"  [ExtractorAgent] Extracting: prompt={prompt_file}, schema={schema_file}")
    with open(doctext_path, "r", encoding="utf-8") as f:
        doc_text = f.read()
    data = await extract_document(doc_text, prompt_file, schema_file)
    os.makedirs(os.path.dirname(output_json_path) or ".", exist_ok=True)
    with open(output_json_path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=4, ensure_ascii=False)
    n_keys = len(data) if isinstance(data, dict) else 0
    print(f"  [ExtractorAgent] Extracted {n_keys} top-level keys → {output_json_path}")
    return f"Extracted JSON ({n_keys} sections) saved to {output_json_path}"


@tool(approval_mode="never_require")
async def validate_extraction(
    doctext_path: str, extracted_json_path: str, validation_output_path: str
) -> str:
    """Validate extracted JSON against original document text using LLM.
    Reads document text from doctext_path and extracted JSON from extracted_json_path.
    Writes validation report to validation_output_path."""
    print(f"  [ValidatorAgent] Validating {extracted_json_path}...")
    with open(doctext_path, "r", encoding="utf-8") as f:
        doc_text = f.read()
    with open(extracted_json_path, "r", encoding="utf-8") as f:
        extracted_json_str = f.read()
    prompt_template = load_text_file(VALIDATOR_PROMPT_FILE)
    full_prompt = prompt_template.replace(
        "{extracted_json}", extracted_json_str
    ).replace("{doc_text}", doc_text)
    result_text, _usage = await run_llm_extraction(full_prompt)
    validation_data = safe_parse_json(result_text)
    os.makedirs(os.path.dirname(validation_output_path) or ".", exist_ok=True)
    with open(validation_output_path, "w", encoding="utf-8") as f:
        json.dump(validation_data, f, indent=4, ensure_ascii=False)
    acc = validation_data.get("accuracy_percentage", "N/A")
    print(f"  [ValidatorAgent] Accuracy: {acc}% → {validation_output_path}")
    return f"Validation complete ({acc}% accuracy), saved to {validation_output_path}"


def _infer_doc_type(extracted_data: dict) -> str:
    """Infer the document type from characteristic keys in the extracted JSON."""
    keys = set(extracted_data.keys())
    if "Pupil_Views" in keys or "Cognition_and_Learning" in keys:
        return "education"
    if "Medical_History" in keys or "Health_Provision" in keys or "Health_Needs" in keys:
        return "health"
    if "H1_Social_Care_Provision" in keys or "Social_Care_Needs" in keys:
        return "socialcare"
    return "personal"


@tool(approval_mode="never_require")
def recheck_validation(
    validation_json_path: str, extracted_json_path: str, doctext_path: str,
) -> str:
    """Apply rule-based post-processing to validation report.
    Reads validation JSON from validation_json_path and extracted JSON from
    extracted_json_path. Reads document text from doctext_path.
    Overwrites the validation file with corrected results."""
    print(f"  [QualityCheckerAgent] Rechecking {validation_json_path}...")
    with open(validation_json_path, "r", encoding="utf-8") as f:
        validation_data = json.load(f)
    with open(doctext_path, "r", encoding="utf-8") as f:
        doc_text = f.read()
    corrected = recheck_incorrect_fields(
        validation_data, extracted_json_path,
        rules=RECHECK_RULES, document_text=doc_text,
    )
    try:
        with open(extracted_json_path, "r", encoding="utf-8") as f:
            extracted_data = json.load(f)
        doc_type = _infer_doc_type(extracted_data)
        completeness = compute_completeness(extracted_data, doc_type)
        corrected["completeness_percentage"] = completeness["completeness_percentage"]
        corrected["critical_fields_total"] = completeness["critical_fields_total"]
        corrected["critical_fields_populated"] = completeness["critical_fields_populated"]
        corrected["critical_fields_missing"] = completeness["critical_fields_missing"]
        corrected["empty_sections"] = completeness["empty_sections"]
        print(f"  [QualityCheckerAgent] Completeness: {completeness['completeness_percentage']}% "
              f"({completeness['critical_fields_populated']}/{completeness['critical_fields_total']} critical fields)")
    except Exception as e:
        print(f"  [QualityCheckerAgent] Completeness check failed: {e}")
    with open(validation_json_path, "w", encoding="utf-8") as f:
        json.dump(corrected, f, indent=4, ensure_ascii=False)
    acc = corrected.get("accuracy_percentage", "N/A")
    comp = corrected.get("completeness_percentage", "N/A")
    print(f"  [QualityCheckerAgent] Final accuracy: {acc}%, completeness: {comp}% → {validation_json_path}")
    return f"PIPELINE_COMPLETE: Quality check done, final accuracy {acc}%, completeness {comp}%"


@tool(approval_mode="never_require")
def fill_template(
    template_docx: str, output_docx: str,
    personal_json_path: str, education_json_path: str,
    health_json_path: str, socialcare_json_path: str,
    mapping_workbook: str,
) -> str:
    """Fill EHCP DOCX template with extracted JSON data using the mapping workbook.
    Writes the filled DOCX to output_docx path."""
    print(f"  [WriterAgent] Filling template: {template_docx}")
    result_path = fill_template_streamlit(
        template_docx=template_docx,
        output_docx=output_docx,
        personal_json_path=personal_json_path,
        education_json_path=education_json_path,
        health_json_path=health_json_path,
        socialcare_json_path=socialcare_json_path,
        mapping_workbook=mapping_workbook,
    )
    print(f"  [WriterAgent] Template filled → {result_path}")
    return f"Template filled and saved to {result_path}"


@tool(approval_mode="never_require")
def validate_writer_output(
    filled_docx_path: str,
    personal_json: str, education_json: str,
    health_json: str, socialcare_json: str,
    mapping_workbook: str, expected_output_docx: str,
    report_output_path: str,
) -> str:
    """Validate the filled EHCP DOCX against extracted JSONs, mapping workbook,
    and expected output. Writes validation report to report_output_path."""
    print(f"  [WriterValidatorAgent] Validating: {filled_docx_path}")
    report = ValidationReport()
    validate_docx(filled_docx_path, report)
    personal_data = (
        load_json_file(personal_json, report, "Personal details JSON", required=False)
        if personal_json else None
    )
    education_data = (
        load_json_file(education_json, report, "Education advice JSON", required=False)
        if education_json else None
    )
    socialcare_data = (
        load_json_file(socialcare_json, report, "Social care advice JSON", required=False)
        if socialcare_json else None
    )
    health_data = (
        load_json_file(health_json, report, "Health advice JSON", required=False)
        if health_json else None
    )
    validate_personal_json(personal_data, report)
    validate_education_json(education_data, report)
    validate_socialcare_json(socialcare_data, report)
    validate_health_json(health_data, report)
    if os.path.exists(mapping_workbook):
        workbook = load_workbook_safe(mapping_workbook, report)
        if workbook is not None:
            validate_workbook_headers(workbook, report)
            validate_personal_mapping(workbook, personal_data, report)
            validate_section_mapping(workbook, EDUCATION_DETAILS_SHEET, education_data, report)
            validate_section_mapping(workbook, SOCIAL_CARE_DETAILS_SHEET, socialcare_data, report)
            validate_section_mapping(workbook, HEALTH_DETAILS_SHEET, health_data, report)
    if expected_output_docx and os.path.exists(expected_output_docx):
        compare_docx_outputs(filled_docx_path, expected_output_docx, report)
    report_data = build_json_report(report)
    os.makedirs(os.path.dirname(report_output_path) or ".", exist_ok=True)
    with open(report_output_path, "w", encoding="utf-8") as f:
        json.dump(report_data, f, indent=2, ensure_ascii=False)
    total_checks = report_data.get("summary", {}).get("total_checks", 0)
    print(f"  [WriterValidatorAgent] {total_checks} checks → {report_output_path}")
    return f"PIPELINE_COMPLETE: Writer validation done ({total_checks} checks), saved to {report_output_path}"


# =========================================================
# Agent Creation Functions
# =========================================================

READER_INSTRUCTIONS = (
    "You are the DocumentReaderAgent in an EHCP document processing pipeline.\n"
    "When you receive a task message, look for the 'file_path' and 'doctext_path' values.\n"
    "Call your read_document tool with those exact paths to extract text from the document.\n"
    "Always call the tool — never try to answer without using it."
)

EXTRACTOR_INSTRUCTIONS = (
    "You are the ExtractorAgent in an EHCP document processing pipeline.\n"
    "When it is your turn, look for 'doctext_path', 'prompt_file', 'schema_file', "
    "and 'output_json_path' in the conversation.\n"
    "Call your extract_to_json tool with those exact paths.\n"
    "Always call the tool — never try to answer without using it."
)

VALIDATOR_INSTRUCTIONS = (
    "You are the ValidatorAgent in an EHCP document processing pipeline.\n"
    "When it is your turn, look for 'doctext_path', 'output_json_path', "
    "and 'validation_json_path' in the conversation.\n"
    "Call your validate_extraction tool with those exact paths.\n"
    "Always call the tool — never try to answer without using it."
)

QUALITY_CHECKER_INSTRUCTIONS = (
    "You are the QualityCheckerAgent in an EHCP document processing pipeline.\n"
    "When it is your turn, look for 'validation_json_path', 'output_json_path', "
    "and 'doctext_path' in the conversation.\n"
    "Call your recheck_validation tool with those exact paths.\n"
    "Always call the tool — never try to answer without using it.\n"
    "After the tool returns, respond with the word PIPELINE_COMPLETE followed by "
    "the final accuracy percentage."
)

WRITER_INSTRUCTIONS = (
    "You are the TemplateWriterAgent in an EHCP writer pipeline.\n"
    "When you receive a task message, look for 'template_docx', 'output_docx', "
    "'personal_json', 'education_json', 'health_json', 'socialcare_json', "
    "and 'mapping_workbook' in the message.\n"
    "Call your fill_template tool with those exact paths.\n"
    "Always call the tool — never try to answer without using it."
)

WRITER_VALIDATOR_INSTRUCTIONS = (
    "You are the WriterValidatorAgent in an EHCP writer pipeline.\n"
    "When it is your turn, look for the filled DOCX path, JSON paths, "
    "'mapping_workbook', and 'expected_output_docx' in the conversation.\n"
    "Call your validate_writer_output tool with those exact paths.\n"
    "Always call the tool — never try to answer without using it.\n"
    "After the tool returns, respond with the word PIPELINE_COMPLETE followed by "
    "the total number of checks."
)


def create_reader_agent() -> Agent:
    return Agent(
        client=_create_client(),
        name="DocumentReaderAgent",
        instructions=READER_INSTRUCTIONS,
        tools=[read_document],
    )


def create_extractor_agent() -> Agent:
    return Agent(
        client=_create_client(),
        name="ExtractorAgent",
        instructions=EXTRACTOR_INSTRUCTIONS,
        tools=[extract_to_json],
    )


def create_validator_agent() -> Agent:
    return Agent(
        client=_create_client(),
        name="ValidatorAgent",
        instructions=VALIDATOR_INSTRUCTIONS,
        tools=[validate_extraction],
    )


def create_quality_checker_agent() -> Agent:
    return Agent(
        client=_create_client(),
        name="QualityCheckerAgent",
        instructions=QUALITY_CHECKER_INSTRUCTIONS,
        tools=[recheck_validation],
    )


def create_writer_agent() -> Agent:
    return Agent(
        client=_create_client(),
        name="TemplateWriterAgent",
        instructions=WRITER_INSTRUCTIONS,
        tools=[fill_template],
    )


def create_writer_validator_agent() -> Agent:
    return Agent(
        client=_create_client(),
        name="WriterValidatorAgent",
        instructions=WRITER_VALIDATOR_INSTRUCTIONS,
        tools=[validate_writer_output],
    )
