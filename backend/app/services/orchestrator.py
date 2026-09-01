"""
EHCP Agent Framework — Orchestrator

Calls @tool functions sequentially for a deterministic pipeline.
Agent objects are defined in agents.py for the Agent Framework,
but the orchestrator invokes the tool functions directly for efficiency
(no extra LLM round-trips for tool selection).

Reader Pipeline (per file, parallel across files):
    read_document → extract_to_json → validate_extraction → recheck_validation

Writer Pipeline (single pass):
    fill_template → validate_writer_output
"""

import os
import json
import re
import asyncio
import time
from typing import Callable, Dict, List, Optional

from app.services.agents import (
    read_document,
    extract_to_json,
    validate_extraction,
    recheck_validation,
    fill_template,
    validate_writer_output,
)
from app.services.helpers import start_token_tracking, get_token_usage
from app.services.blob_storage import upload_file_to_blob, is_blob_storage_enabled


# =========================================================
# Reader Pipeline Orchestrator
# =========================================================

class EHCPAgentOrchestrator:
    """Orchestrates the reader pipeline by calling tool functions sequentially.

    Each file is processed through 4 steps:
        read_document → extract_to_json → validate_extraction → recheck_validation
    """

    async def process_single_file(
        self,
        file_path: str,
        prompt_file: str,
        schema_file: str,
        output_file: str,
        validation_output_file: str,
        progress_callback: Optional[Callable] = None,
    ) -> Dict:
        """Process one EHCP document through the 4-step reader pipeline."""

        # Start token tracking for this file
        start_token_tracking()

        def _notify(event):
            if progress_callback:
                progress_callback(event)

        base = os.path.splitext(os.path.basename(file_path))[0]
        # Place doctext in the same directory as output_file (session-scoped)
        output_dir = os.path.dirname(os.path.abspath(output_file))
        os.makedirs(output_dir, exist_ok=True)
        doctext_path = os.path.join(output_dir, f"{base}_doctext.txt")
        file_path = os.path.abspath(file_path)
        prompt_file = os.path.abspath(prompt_file)
        schema_file = os.path.abspath(schema_file)
        output_file = os.path.abspath(output_file)
        validation_output_file = os.path.abspath(validation_output_file)

        _notify({"file_path": file_path, "stage": "agent_pipeline", "event": "start",
                 "agent": "Pipeline"})
        pipeline_start = time.perf_counter()

        # Step 1: Read document (sync tool)
        _notify({"file_path": file_path, "stage": "agent_reader", "event": "start",
                 "agent": "DocumentReaderAgent"})
        read_document(file_path=file_path, output_text_path=doctext_path)
        elapsed = time.perf_counter() - pipeline_start
        print(f"  [DocumentReaderAgent] done ({elapsed:.1f}s)")
        _notify({"file_path": file_path, "stage": "agent_reader", "event": "done",
                 "agent": "DocumentReaderAgent", "elapsed_seconds": elapsed})

        # Step 2: Extract JSON (async tool)
        _notify({"file_path": file_path, "stage": "agent_extractor", "event": "start",
                 "agent": "ExtractorAgent"})
        await extract_to_json(
            doctext_path=doctext_path,
            prompt_file=prompt_file,
            schema_file=schema_file,
            output_json_path=output_file,
        )
        elapsed = time.perf_counter() - pipeline_start
        print(f"  [ExtractorAgent] done ({elapsed:.1f}s)")
        _notify({"file_path": file_path, "stage": "agent_extractor", "event": "done",
                 "agent": "ExtractorAgent", "elapsed_seconds": elapsed})

        # Step 3: Validate extraction (async tool)
        _notify({"file_path": file_path, "stage": "agent_validator", "event": "start",
                 "agent": "ValidatorAgent"})
        await validate_extraction(
            doctext_path=doctext_path,
            extracted_json_path=output_file,
            validation_output_path=validation_output_file,
        )
        elapsed = time.perf_counter() - pipeline_start
        print(f"  [ValidatorAgent] done ({elapsed:.1f}s)")
        _notify({"file_path": file_path, "stage": "agent_validator", "event": "done",
                 "agent": "ValidatorAgent", "elapsed_seconds": elapsed})

        # Step 4: Quality check (sync tool)
        _notify({"file_path": file_path, "stage": "agent_quality", "event": "start",
                 "agent": "QualityCheckerAgent"})
        recheck_validation(
            validation_json_path=validation_output_file,
            extracted_json_path=output_file,
            doctext_path=doctext_path,
        )
        elapsed = time.perf_counter() - pipeline_start
        print(f"  [QualityCheckerAgent] done ({elapsed:.1f}s)")
        _notify({"file_path": file_path, "stage": "agent_quality", "event": "done",
                 "agent": "QualityCheckerAgent", "elapsed_seconds": elapsed})

        pipeline_elapsed = time.perf_counter() - pipeline_start
        _notify({"file_path": file_path, "stage": "agent_pipeline", "event": "done",
                 "agent": "Pipeline", "elapsed_seconds": pipeline_elapsed})

        # Read results from disk (tools wrote them during execution)
        result = {
            "output_file": output_file,
            "validation_file": validation_output_file,
            "document_text": None,
            "extracted_data": None,
            "validation_data": None,
            "token_usage": get_token_usage(),
        }

        if os.path.exists(doctext_path):
            with open(doctext_path, "r", encoding="utf-8") as f:
                result["document_text"] = f.read()

        if os.path.exists(output_file):
            with open(output_file, "r", encoding="utf-8") as f:
                result["extracted_data"] = json.load(f)

        if os.path.exists(validation_output_file):
            with open(validation_output_file, "r", encoding="utf-8") as f:
                result["validation_data"] = json.load(f)

        # Upload to blob storage (session-scoped key derived from the path)
        if is_blob_storage_enabled():
            for fpath in [output_file, validation_output_file]:
                if os.path.exists(fpath):
                    try:
                        upload_file_to_blob(fpath)
                        print(
                            f"  [Orchestrator] Uploaded {os.path.basename(fpath)} to blob")
                    except Exception as blob_exc:
                        print(
                            f"  [Orchestrator] Blob upload failed: {blob_exc}")

        return result

    async def process_batch(
        self,
        file_configs: List[Dict],
        progress_callback: Optional[Callable] = None,
    ) -> tuple:
        """Process multiple files in parallel."""
        document_texts = {}

        async def _process_one(config):
            result = await self.process_single_file(
                file_path=config["input_docx"],
                prompt_file=config["prompt_file"],
                schema_file=config["schema_file"],
                output_file=config["output_file"],
                validation_output_file=config["validation_output_file"],
                progress_callback=progress_callback,
            )
            document_texts[config["input_docx"]] = result["document_text"]
            return result

        tasks = [_process_one(cfg) for cfg in file_configs]
        results = await asyncio.gather(*tasks, return_exceptions=True)
        return results, document_texts


# =========================================================
# Convenience wrapper for Streamlit
# =========================================================

async def run_maf_pipeline(
    file_configs: List[Dict], progress_callback: Optional[Callable] = None
) -> tuple:
    """Entry point for the reader pipeline."""
    orchestrator = EHCPAgentOrchestrator()
    return await orchestrator.process_batch(file_configs, progress_callback)


# =========================================================
# Writer Pipeline Orchestrator
# =========================================================

class EHCPWriterOrchestrator:
    """Orchestrates the writer pipeline by calling tool functions sequentially.

    Two steps:
        fill_template → validate_writer_output
    """

    async def run_writer_pipeline(
        self,
        template_docx: str,
        output_docx: Optional[str],
        json_paths: Dict[str, Optional[str]],
        mapping_workbook: str,
        expected_output_docx: str,
        progress_callback: Optional[Callable] = None,
        output_dir_override: Optional[str] = None,
        report_dir_override: Optional[str] = None,
    ) -> Dict:
        """Run the writer + validation pipeline."""

        def _notify(event):
            if progress_callback:
                progress_callback(event)

        report_dir = os.path.abspath(
            report_dir_override) if report_dir_override else os.path.abspath("temp")
        os.makedirs(report_dir, exist_ok=True)
        report_output_path = os.path.join(
            report_dir, "writer_validation_report.json")
        template_docx = os.path.abspath(template_docx)
        mapping_workbook = os.path.abspath(mapping_workbook)
        expected_output_docx = os.path.abspath(expected_output_docx)

        # Resolve the output DOCX path
        if output_docx:
            filled_docx_hint = os.path.abspath(output_docx)
        elif output_dir_override:
            os.makedirs(output_dir_override, exist_ok=True)
            _pjson = json_paths.get("personal") or ""
            _child_name = "OUTPUT"
            if _pjson and os.path.exists(_pjson):
                try:
                    with open(_pjson, "r", encoding="utf-8") as _f:
                        _pdata = json.load(_f)
                    _child_name = (str(_pdata.get("name") or _pdata.get(
                        "preferred_name") or "")).strip()
                    if not _child_name:
                        _child_name = "OUTPUT"
                except Exception:
                    pass
            _safe = re.sub(r'[<>:"/\\|?*]+', " ", _child_name)
            _safe = re.sub(r"\s+", " ", _safe).strip() or "OUTPUT"
            filled_docx_hint = os.path.abspath(os.path.join(
                output_dir_override, f"{_safe} Draft EHCP.docx"))
        else:
            filled_docx_hint = ""

        _notify({"stage": "agent_writer_pipeline", "event": "start",
                 "agent": "Pipeline"})
        pipeline_start = time.perf_counter()

        # Step 1: Fill template (sync tool)
        _notify({"stage": "agent_writer", "event": "start",
                 "agent": "TemplateWriterAgent"})
        writer_result = fill_template(
            template_docx=template_docx,
            output_docx=filled_docx_hint,
            personal_json_path=json_paths.get("personal") or "",
            education_json_path=json_paths.get("education") or "",
            health_json_path=json_paths.get("health") or "",
            socialcare_json_path=json_paths.get("socialcare") or "",
            mapping_workbook=mapping_workbook,
        )
        elapsed = time.perf_counter() - pipeline_start
        print(
            f"  [TemplateWriterAgent] done ({elapsed:.1f}s): {writer_result[:120]}")
        _notify({"stage": "agent_writer", "event": "done",
                 "agent": "TemplateWriterAgent", "elapsed_seconds": elapsed})

        # Capture the actual output path from the tool's return string
        actual_filled_path = None
        if "saved to " in writer_result:
            raw_path = writer_result.split("saved to ", 1)[-1].strip()
            docx_idx = raw_path.lower().find(".docx")
            if docx_idx != -1:
                raw_path = raw_path[: docx_idx + 5]
            if os.path.exists(raw_path):
                actual_filled_path = raw_path

        # Step 2: Validate writer output (sync tool)
        filled_docx = actual_filled_path or filled_docx_hint

        _notify({"stage": "agent_writer_validator", "event": "start",
                 "agent": "WriterValidatorAgent"})
        validate_writer_output(
            filled_docx_path=filled_docx,
            personal_json=json_paths.get("personal") or "",
            education_json=json_paths.get("education") or "",
            health_json=json_paths.get("health") or "",
            socialcare_json=json_paths.get("socialcare") or "",
            mapping_workbook=mapping_workbook,
            expected_output_docx=expected_output_docx,
            report_output_path=report_output_path,
        )
        elapsed = time.perf_counter() - pipeline_start
        print(f"  [WriterValidatorAgent] done ({elapsed:.1f}s)")
        _notify({"stage": "agent_writer_validator", "event": "done",
                 "agent": "WriterValidatorAgent", "elapsed_seconds": elapsed})

        pipeline_elapsed = time.perf_counter() - pipeline_start
        _notify({"stage": "agent_writer_pipeline", "event": "done",
                 "agent": "Pipeline", "elapsed_seconds": pipeline_elapsed})

        # Fallback: if path capture failed, find the newest .docx in output/
        if not filled_docx or not os.path.exists(filled_docx):
            search_dir = output_dir_override or os.path.join(
                os.path.dirname(template_docx), "output")
            if os.path.isdir(search_dir):
                docx_files = [
                    os.path.join(search_dir, f)
                    for f in os.listdir(search_dir)
                    if f.lower().endswith(".docx")
                ]
                if docx_files:
                    filled_docx = max(docx_files, key=os.path.getmtime)
                    print(
                        f"  [Orchestrator] Fallback: found {os.path.basename(filled_docx)}")

        result = {
            "filled_docx_path": filled_docx if filled_docx and os.path.exists(filled_docx) else None,
            "validation_report": None,
        }

        if os.path.exists(report_output_path):
            with open(report_output_path, "r", encoding="utf-8") as f:
                result["validation_report"] = json.load(f)

        # Blob upload
        if is_blob_storage_enabled():
            if result["filled_docx_path"] and os.path.exists(result["filled_docx_path"]):
                try:
                    docx_name = os.path.basename(result["filled_docx_path"])
                    upload_file_to_blob(result["filled_docx_path"])
                    print(f"  [Orchestrator] Uploaded {docx_name} to blob")
                except Exception as blob_exc:
                    print(
                        f"  [Orchestrator] Blob upload failed for DOCX: {blob_exc}")

            if os.path.exists(report_output_path):
                try:
                    upload_file_to_blob(report_output_path)
                    print(
                        f"  [Orchestrator] Uploaded writer_validation_report.json to blob")
                except Exception as blob_exc:
                    print(f"  [Orchestrator] Blob upload failed: {blob_exc}")

        return result


async def run_maf_writer_pipeline(
    template_docx: str,
    output_docx: Optional[str],
    json_paths: Dict[str, Optional[str]],
    mapping_workbook: str,
    expected_output_docx: str,
    progress_callback: Optional[Callable] = None,
    output_dir_override: Optional[str] = None,
    report_dir_override: Optional[str] = None,
) -> Dict:
    """Entry point for the writer pipeline."""
    orchestrator = EHCPWriterOrchestrator()
    return await orchestrator.run_writer_pipeline(
        template_docx=template_docx,
        output_docx=output_docx,
        json_paths=json_paths,
        mapping_workbook=mapping_workbook,
        expected_output_docx=expected_output_docx,
        progress_callback=progress_callback,
        output_dir_override=output_dir_override,
        report_dir_override=report_dir_override,
    )
