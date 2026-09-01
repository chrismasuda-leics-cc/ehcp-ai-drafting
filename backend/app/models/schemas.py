from pydantic import BaseModel
from typing import Dict, List, Optional, Any


class FileConfig(BaseModel):
    filename: str
    doc_type: str  # "Personal Details", "Education Advice", "Health Advice", "Social Care Advice"


class AnalyzeRequest(BaseModel):
    files: List[FileConfig]
    job_id: Optional[str] = None


class AnalyzeResult(BaseModel):
    filename: str
    output_file: str
    validation_file: str
    success: bool
    error: Optional[str] = None


class AnalyzeResponse(BaseModel):
    results: List[AnalyzeResult]


class WriteRequest(BaseModel):
    json_paths: Dict[str, Optional[str]]
    job_id: Optional[str] = None


class WriteResponse(BaseModel):
    filled_docx_path: Optional[str] = None
    validation_report: Optional[Dict[str, Any]] = None
