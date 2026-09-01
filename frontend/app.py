import time
import uuid
import base64
import streamlit as st
import os
import json
import requests
from urllib.parse import quote
from auth import login_ui, get_auth_headers, logout, AUTH_ENABLED
from dotenv import load_dotenv

load_dotenv()

# -----------------------------------------------------
# Backend API URL
# -----------------------------------------------------

ENV = os.getenv("ENV", "")
BACKEND_URL = os.getenv("BACKEND_URL", "http://localhost:8000")
API_BASE = f"{BACKEND_URL}/api"

# Debug mode: show JSON/validation download buttons (set via env var or query param ?debug=1)
DEBUG_MODE = os.getenv("DEBUG_MODE", "false").lower() == "true"

# -----------------------------------------------------
# Session ID — isolates each user's files on the backend
# -----------------------------------------------------

if "session_id" not in st.session_state:
    st.session_state.session_id = f"user_{uuid.uuid4().hex[:12]}"

# Job ID — one stable id per case so repeated uploads map to a single job
# record (prevents duplicate job records under rapid/concurrent uploads).
# Cleared with the rest of the session state on "Start New Case".
if "job_id" not in st.session_state:
    st.session_state.job_id = uuid.uuid4().hex

SESSION_HEADERS = {"X-Session-ID": st.session_state.session_id}


def _get_headers() -> dict:
    """Merge session headers with auth headers and user info."""
    headers = {**SESSION_HEADERS}
    headers.update(get_auth_headers())
    # Always pass user info for audit logging
    user_name = st.session_state.get("user_name", "")
    user_email = st.session_state.get("user_email", "")
    if user_name:
        headers["X-User-Name"] = user_name
    if user_email:
        headers["X-User-Email"] = user_email
    return headers


def _log_activity(action: str, details: dict = None):
    """Send an audit log entry to the backend."""
    try:
        requests.post(
            f"{API_BASE}/log-activity",
            json={"action": action, "details": details or {}},
            headers=_get_headers(),
            timeout=5,
        )
    except Exception:
        pass


def _auto_download(data: bytes, filename: str, mime: str):
    """Trigger an automatic file download using st.html (no iframe sandbox)."""
    b64 = base64.b64encode(data).decode()
    html = f"""
    <script>
    (function() {{
        var a = document.createElement('a');
        a.href = 'data:{mime};base64,{b64}';
        a.download = '{filename}';
        a.style.display = 'none';
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
    }})();
    </script>
    """
    st.html(html)


def _get_certification_footer() -> str:
    """Return the certification footer text with user details and timestamp."""
    from datetime import datetime, timezone
    user_name = st.session_state.get("certify_user_name", "")
    timestamp = datetime.now(timezone.utc).strftime("%d/%m/%Y, %H:%M:%S")
    lines = [
        "",
        "─" * 40,
        f"Reviewed by: {user_name}",
        f"Reviewed at: {timestamp}",
        "Origin: EHCP Document Processor",
    ]
    return "\n".join(lines)


@st.dialog("Certify Before Downloading")
def _certify_download_dialog(download_key: str):
    """Show certification form before allowing download."""
    user_name = st.session_state.get("user_name", "Unknown User")
    st.markdown(f"**User:** {user_name}")

    # Only show checkbox for final EHCP document download
    if download_key == "download_docx":
        confirmed = st.checkbox(
            "I understand I am responsible for verifying this AI-generated content before official use.",
            value=False,
        )
    else:
        confirmed = True

    col1, col2 = st.columns(2)
    with col1:
        cancel_clicked = st.button("Cancel", key="cert_dialog_cancel")
    with col2:
        confirm_clicked = st.button(
            "Confirm", key="cert_dialog_ok",
            disabled=not confirmed, type="primary",
        )

    if cancel_clicked:
        st.rerun()
    if confirm_clicked and confirmed:
        from datetime import datetime, timezone
        cert_timestamp = datetime.now(
            timezone.utc).strftime("%d/%m/%Y, %H:%M:%S")
        st.session_state.certify_user_name = user_name
        st.session_state[f"certified_{download_key}"] = True
        _log_activity("certify_download", {
            "download_key": download_key,
            "user_name": user_name,
            "certified_at": cert_timestamp,
            "confirmation": "User confirmed responsibility for verifying AI-generated content",
            "origin": "EHCP Document Processor",
        })
        st.rerun()


# -----------------------------------------------------
# Page Config
# -----------------------------------------------------

st.set_page_config(
    page_title="EHCP Document Processor",
    layout="wide"
)

# Override DEBUG_MODE from query params (?debug=1)
if st.query_params.get("debug") == "1":
    DEBUG_MODE = True

# -----------------------------------------------------
# Global CSS
# -----------------------------------------------------

st.markdown("""
<style>
/* Constrain main content width */
.block-container {
    max-width: 900px;
    padding-left: 2rem;
    padding-right: 2rem;
    padding-top: 4rem;
}

/* Align sidebar top with main content */
[data-testid="stSidebar"] > div:first-child {
    padding-top: 2rem;
}

/* Sidebar width */
[data-testid="stSidebar"] {
    display: none !important;
}

/* Top-right user bar */
.user-bar {
    position: fixed;
    top: 10px;
    right: 20px;
    z-index: 999;
    display: flex;
    align-items: center;
    gap: 12px;
    font-family: 'Segoe UI', Arial, sans-serif;
    font-size: 14px;
    color: #333;
}
.user-bar .user-name {
    font-weight: 600;
}

/* Smaller file uploader - hide native file list, show custom one */
[data-testid="stFileUploaderDropzone"] {
    max-width: 500px;
}
/* Hide ALL native file list items and pagination */
[data-testid="stFileUploader"] ul,
[data-testid="stFileUploader"] [data-testid="stFileUploaderFile"],
[data-testid="stFileUploader"] small,
[data-testid="stFileUploader"] nav,
[data-testid="stFileUploader"] [role="navigation"],
[data-testid="stFileUploaderDropzone"] ul,
[data-testid="stFileUploaderDropzone"] li,
[data-testid="stFileUploaderDropzone"] small,
[data-testid="stFileUploaderDropzone"] ~ div[data-testid],
[data-testid="stFileUploaderDropzone"] + div {
    display: none !important;
}
}

/* Hide link anchors on headings */
[data-testid="stMarkdownContainer"] a.anchor-link,
[data-testid="stMarkdownContainer"] a[href^="#"],
h1 a, h2 a, h3 a, h4 a, h5 a,
[data-testid="stHeadingWithActionElements"] [data-testid="stHeaderActionElements"],
.stMarkdown a[href^="#"],
[data-testid="stHeader"] a,
.css-1629p8f a,
a.streamlit-header-anchor,
[kind="header"] a,
[data-testid="StyledLinkIconContainer"],
[data-testid="stHeaderActionElements"],
button[kind="headerAction"],
.st-emotion-cache-1aehpvj,
[class*="anchor"],
[class*="link-icon"] {
    display: none !important;
    visibility: hidden !important;
    width: 0 !important;
    height: 0 !important;
    overflow: hidden !important;
    position: absolute !important;
    pointer-events: none !important;
}

/* Grey toggle switch - override Streamlit theme color */
[data-testid="stToggle"] * {
    --primaryColor: #888 !important;
}
[data-testid="stToggle"] [aria-checked="true"] {
    background-color: #888 !important;
    background: #888 !important;
    border-color: #888 !important;
}
/* Align toggle vertically with heading */
[data-testid="stToggle"] {
    padding-top: 12px;
}

/* Remove extra gap between banner and upload section */
.block-container [data-testid="stVerticalBlock"] > div:nth-child(3) {
    margin-top: -0.5rem;
}

/* Gray primary button */
button[kind="primary"] {
    background-color: #6c757d !important;
    border-color: #6c757d !important;
    color: white !important;
}
button[kind="primary"]:hover {
    background-color: #5a6268 !important;
    border-color: #545b62 !important;
}

/* Green thick checkmark on completed status */
[data-testid="stStatusWidget"] svg {
    color: #1B8A2A !important;
    stroke-width: 3 !important;
    width: 26px !important;
    height: 26px !important;
}

/* Completed status label text bold green */
[data-testid="stStatusWidget"] [data-testid="stMarkdownContainer"] p {
    color: #1B8A2A !important;
    font-weight: 700 !important;
    font-size: 15px !important;
}

/* Hide Streamlit toolbar (record, print, stop, deploy, menu) */
[data-testid="stToolbar"] {
    display: none !important;
}
#MainMenu {
    display: none !important;
}
header[data-testid="stHeader"] {
    display: none !important;
}
[data-testid="stDecoration"] {
    display: none !important;
}

/* Fixed user bar top-right */
.user-bar-fixed {
    position: fixed;
    top: 8px;
    right: 20px;
    z-index: 1000;
    font-family: 'Segoe UI', Arial, sans-serif;
    font-size: 13px;
    color: #333;
    background: rgba(255,255,255,0.95);
    padding: 6px 14px;
    border-radius: 8px;
    box-shadow: 0 1px 4px rgba(0,0,0,0.08);
    white-space: nowrap;
}
</style>
""", unsafe_allow_html=True)

# -----------------------------------------------------
# Authentication Gate
# -----------------------------------------------------

if not login_ui():
    st.stop()

# Handle sign-out via query param (triggered by link in user bar)
if st.query_params.get("signout") == "1":
    st.query_params.clear()
    logout()

# -----------------------------------------------------
# Top-right user bar (fixed position)
# -----------------------------------------------------

# Build user bar HTML
_user_bar_parts = []
if AUTH_ENABLED and "user_name" in st.session_state:
    _user_bar_parts.append(
        f'<span style="font-weight:600;">👤 {st.session_state.user_name}</span>')
_user_bar_parts.append(
    f'<span style="font-size:11px;color:#888;">🔑 {st.session_state.session_id}</span>')
if AUTH_ENABLED and "user_name" in st.session_state:
    _user_bar_parts.append(
        '<a href="?signout=1" style="font-size:12px;color:#666;text-decoration:none;border:1px solid #ccc;'
        'padding:2px 8px;border-radius:4px;margin-left:4px;" '
        'onmouseover="this.style.background=\'#f0f0f0\'" '
        'onmouseout="this.style.background=\'transparent\'">Sign out</a>')

st.markdown(
    '<div class="user-bar-fixed">' +
    ' &nbsp;|&nbsp; '.join(_user_bar_parts) + '</div>',
    unsafe_allow_html=True,
)

st.markdown("")

# -----------------------------------------------------
# Banner
# -----------------------------------------------------

st.markdown(f"""
<div style="
background-color:#236c34;
padding:6px 24px;
border-radius:12px;
max-width:600px;
margin-top:30px;
overflow:hidden;
">

<div style="
color:white;
font-size:18px;
font-weight:700;
margin-bottom:0px;
margin-top:0px;
font-family: Arial, Helvetica, sans-serif;
">
Draft EHCP {ENV}
</div>

<p style="
color:white;
font-size:12px;
margin-top:0px;
margin-bottom:0px;
font-family: Arial, Helvetica, sans-serif;
">
Leicestershire County Council
</p>

</div>
""", unsafe_allow_html=True)

st.write("")

# -----------------------------------------------------
# Document type definitions
# -----------------------------------------------------

DOC_TYPE_OPTIONS = ["Personal Details", "Education Advice",
                    "Health Advice", "Social Care Advice"]

DOC_TYPE_KEY_MAP = {
    "Personal Details": "personal",
    "Education Advice": "education",
    "Health Advice": "health",
    "Social Care Advice": "socialcare",
}


def auto_detect_doc_type(filename: str) -> str:
    """Best-effort guess of document type from filename."""
    name_lower = filename.lower()
    if any(kw in name_lower for kw in ("education", "leps")):
        return "Education Advice"
    if "health" in name_lower:
        return "Health Advice"
    if any(kw in name_lower for kw in ("social", "_sc ", "_sc_", "_sc.")):
        return "Social Care Advice"
    stem = os.path.splitext(name_lower)[0]
    if stem.endswith("_sc") or "_sc " in name_lower or " sc " in name_lower or " sc_" in name_lower:
        return "Social Care Advice"
    return "Personal Details"


def classify_file_by_type(doc_type: str) -> str:
    # Accept both display names ("Education Advice") and key names ("education")
    if doc_type in DOC_TYPE_KEY_MAP:
        return DOC_TYPE_KEY_MAP[doc_type]
    if doc_type in DOC_TYPE_KEY_MAP.values():
        return doc_type
    return "personal"


def count_json_fields(data, prefix=""):
    """Recursively count leaf fields in extracted JSON and classify as filled or empty."""
    filled = []
    empty = []
    if isinstance(data, dict):
        for k, v in data.items():
            path = f"{prefix}.{k}" if prefix else k
            if isinstance(v, dict):
                f, e = count_json_fields(v, path)
                filled.extend(f)
                empty.extend(e)
            elif isinstance(v, list):
                if v and any(item for item in v if item):
                    filled.append(path)
                else:
                    empty.append(path)
            elif v is not None and str(v).strip():
                filled.append(path)
            else:
                empty.append(path)
    return filled, empty


@st.cache_data(ttl=3600)
def fetch_mapping_fields():
    """Fetch the output-document field definitions from the mapping Excel via backend."""
    try:
        resp = requests.get(f"{API_BASE}/mapping-fields", timeout=10)
        resp.raise_for_status()
        return resp.json()
    except Exception:
        return None


def _resolve_json_value(data, json_path):
    """Look up a dotted json_path like 'Social_Care_Needs.Strengths' in the JSON data."""
    parts = json_path.split(".")
    current = data
    for part in parts:
        if isinstance(current, dict) and part in current:
            current = current[part]
        else:
            return None
    return current


def count_mapping_fields(output_data, mapping_fields):
    """Count filled/empty fields based on mapping definitions (what appears in output doc).
    Returns (filled_list, empty_list) of json_path strings."""
    filled = []
    empty = []
    seen = set()
    for field in mapping_fields:
        jp = field["json_path"]
        if jp in seen:
            continue  # deduplicate (e.g. 'name' appears twice in Personal)
        seen.add(jp)
        val = _resolve_json_value(output_data, jp)
        # Check if value is meaningfully filled
        if val is None:
            empty.append(jp)
        elif isinstance(val, list):
            if val and any(item for item in val if item):
                filled.append(jp)
            else:
                empty.append(jp)
        elif isinstance(val, dict):
            # A dict value is filled if it has any non-empty leaf values
            leaf_filled, _ = count_json_fields(val)
            if leaf_filled:
                filled.append(jp)
            else:
                empty.append(jp)
        elif str(val).strip():
            filled.append(jp)
        else:
            empty.append(jp)
    return filled, empty

# -----------------------------------------------------
# Session State Initialization
# -----------------------------------------------------


if "extracted_files" not in st.session_state:
    st.session_state.extracted_files = {}

if "view_output" not in st.session_state:
    st.session_state.view_output = None

if "view_mode" not in st.session_state:
    st.session_state.view_mode = None

if "processing_active" not in st.session_state:
    st.session_state.processing_active = False

if "filled_docx_path" not in st.session_state:
    st.session_state.filled_docx_path = None

if "writing_active" not in st.session_state:
    st.session_state.writing_active = False

if "writer_validation_report" not in st.session_state:
    st.session_state.writer_validation_report = None

if "file_types" not in st.session_state:
    st.session_state.file_types = {}

if "uploaded_file_names" not in st.session_state:
    st.session_state.uploaded_file_names = set()

if "upload_errors" not in st.session_state:
    st.session_state.upload_errors = {}

# Cache for downloaded JSON results
if "json_cache" not in st.session_state:
    st.session_state.json_cache = {}

# Cache for filled DOCX binary content (survives reruns)
if "filled_docx_bytes" not in st.session_state:
    st.session_state.filled_docx_bytes = None

# Child name tracking: {filename: child_name}
if "child_names" not in st.session_state:
    st.session_state.child_names = {}

# Filenames flagged as belonging to a different child
if "mismatched_files" not in st.session_state:
    st.session_state.mismatched_files = set()


# -----------------------------------------------------
# Helper: child name mismatch detection
# -----------------------------------------------------

def _names_match(name_a: str, name_b: str) -> bool:
    """True only if both names are exactly the same (case-insensitive)."""
    return name_a.strip().lower() == name_b.strip().lower()


def _check_child_name_mismatches():
    """Compare child names across all uploaded files.  The Personal Details
    document is treated as the primary child name; other files with a different
    name are flagged.  Files with no detected name are NOT flagged."""
    names = st.session_state.child_names  # {filename: name}
    if len(names) < 2:
        st.session_state.mismatched_files = set()
        return

    # Use Personal Details as the primary name source
    file_types = st.session_state.get("file_types", {})
    primary_name = None
    for fname, cname in names.items():
        dt = file_types.get(fname, auto_detect_doc_type(fname))
        if classify_file_by_type(dt) == "personal":
            primary_name = cname
            break

    # Fallback to most common name if Personal Details not found
    if not primary_name:
        from collections import Counter
        name_counts = Counter(names.values())
        primary_name = name_counts.most_common(1)[0][0]

    mismatched = set()
    for fname, cname in names.items():
        if not _names_match(cname, primary_name):
            mismatched.add(fname)

    st.session_state.mismatched_files = mismatched


# -----------------------------------------------------
# Upload Section
# -----------------------------------------------------

# Track when uploader changes to defer backend uploads to next rerun
if "pending_upload" not in st.session_state:
    st.session_state.pending_upload = False

# File store: holds file data independently of the Streamlit widget
# {filename: {"name": str, "data": bytes, "size": int}}
if "file_store" not in st.session_state:
    st.session_state.file_store = {}
if "removed_files" not in st.session_state:
    st.session_state.removed_files = set()


def _on_uploader_change():
    """Callback: mark that files changed so backend upload happens on next
    rerun (after Streamlit has fully received all files from the browser)."""
    st.session_state.pending_upload = True
    st.session_state.pop("cleaned_removed_files", None)
    st.session_state.pop("_mismatch_override", None)
    st.session_state.pop("_mismatch_rejected", None)
    # Clear removed_files so re-browsed files can be added fresh
    st.session_state.removed_files = set()
    # Log the browse/file-select action to audit log
    try:
        requests.post(
            f"{API_BASE}/log-browse",
            json=[],
            headers=_get_headers(),
            timeout=5,
        )
    except Exception:
        pass


class StoredFile:
    """Lightweight wrapper to mimic UploadedFile interface using stored bytes."""

    def __init__(self, name, data, size):
        self.name = name
        self._data = data
        self.size = size
        self._pos = 0

    def seek(self, pos):
        self._pos = pos

    def read(self):
        self._pos = len(self._data)
        return self._data


st.markdown("<p style='font-weight:700;font-size:1.25rem;margin-bottom:0.5rem;'>Upload EHCP Documents</p>",
            unsafe_allow_html=True)
widget_files = st.file_uploader(
    "",
    type=["docx", "pdf"],
    accept_multiple_files=True,
    key=f"ehcp_file_uploader_{st.session_state.get('uploader_key', 0)}",
    on_change=_on_uploader_change,
)

# Always merge widget files into file_store (source of truth)
# Don't gate on pending_upload — files may arrive across multiple reruns
if widget_files:
    for f in widget_files:
        if f.name not in st.session_state.file_store:
            f.seek(0)
            st.session_state.file_store[f.name] = {
                "name": f.name,
                "data": f.read(),
                "size": f.size,
            }

# Build uploaded_files list from file_store (source of truth)
uploaded_files = [StoredFile(v["name"], v["data"], v["size"])
                  for v in st.session_state.file_store.values()]

# Force Streamlit file uploader to show all files (expand pagination)
if uploaded_files and len(uploaded_files) > 3:
    pass  # Native list hidden via CSS; custom list rendered below

# Custom file list showing ALL files with inline X buttons
if uploaded_files:
    # Hide delete buttons once analysis has started or completed
    allow_delete = not st.session_state.get(
        "extracted_files") and not st.session_state.get("processing_active")

    for f in uploaded_files:
        if allow_delete:
            col_info, col_del = st.columns([9, 1])
            with col_info:
                st.markdown(
                    f"📄 **{f.name}** &nbsp; <span style='color:#888;font-size:12px;'>{f.size / 1024:.1f} KB</span>", unsafe_allow_html=True)
            with col_del:
                if st.button("✕", key=f"remove_{f.name}"):
                    st.session_state.file_store.pop(f.name, None)
                    st.session_state.extracted_files.pop(f.name, None)
                    st.session_state.json_cache.pop(f.name, None)
                    st.session_state.child_names.pop(f.name, None)
                    st.session_state.mismatched_files.discard(f.name)
                    st.session_state.uploaded_file_names.discard(f.name)
                    st.session_state.upload_errors.pop(f.name, None)
                    st.session_state.file_types.pop(f.name, None)
                    # Reset widget to clear deleted file from its state
                    st.session_state.uploader_key = st.session_state.get(
                        "uploader_key", 0) + 1
                    try:
                        requests.delete(f"{API_BASE}/files/{f.name}",
                                        headers=_get_headers(), timeout=5)
                    except Exception:
                        pass
                    st.rerun()
        else:
            st.markdown(
                f"📄 **{f.name}** &nbsp; <span style='color:#888;font-size:12px;'>{f.size / 1024:.1f} KB</span>", unsafe_allow_html=True)

# -----------------------------------------------------
# Sync session state with current file list
# -----------------------------------------------------

if uploaded_files:
    current_names = set(f.name for f in uploaded_files)

    # --- Delete removed/deselected files from backend + blob ---
    removed_extracted = [
        name for name in st.session_state.extracted_files if name not in current_names]
    for name in removed_extracted:
        del st.session_state.extracted_files[name]
        st.session_state.json_cache.pop(name, None)
        st.session_state.child_names.pop(name, None)
        st.session_state.mismatched_files.discard(name)
        try:
            requests.delete(f"{API_BASE}/files/{name}",
                            headers=_get_headers(), timeout=5)
        except Exception:
            pass

    removed_uploaded = st.session_state.uploaded_file_names - current_names
    for name in removed_uploaded:
        st.session_state.child_names.pop(name, None)
        st.session_state.mismatched_files.discard(name)
        try:
            requests.delete(f"{API_BASE}/files/{name}",
                            headers=_get_headers(), timeout=5)
        except Exception:
            pass
    st.session_state.uploaded_file_names -= removed_uploaded

    # Recheck child name mismatches after removals
    if removed_extracted or removed_uploaded:
        _check_child_name_mismatches()

    # --- Upload newly added files to backend ---
    # Only process when the uploader callback has fired (deferred upload),
    # so we don't block the Streamlit rerun cycle while files are still
    # being received from the browser (which causes the AxiosError 400).
    newly_added = current_names - st.session_state.uploaded_file_names

    active_upload_errors = {
        name: error for name, error in st.session_state.upload_errors.items()
        if name in newly_added
    }
    if active_upload_errors and not st.session_state.get("pending_upload"):
        failed_names = ", ".join(sorted(active_upload_errors))
        st.error(f"Upload failed for: {failed_names}. Please retry.")
        if st.button("Retry failed uploads", key="retry_failed_uploads"):
            st.session_state.pending_upload = True
            st.rerun()

    # Always clear pending_upload to prevent it from interfering with file removal
    if st.session_state.get("pending_upload") and not newly_added:
        st.session_state.pending_upload = False

    if newly_added and st.session_state.get("pending_upload"):
        st.session_state.pending_upload = False
        upload_msg = st.info("Uploading input documents to storage account...")
        upload_failed = False

        for f in uploaded_files:
            if f.name in newly_added:
                f.seek(0)
                file_bytes = f.read()

                files = {
                    "files": (f.name, file_bytes, "application/octet-stream")
                }

                try:
                    upload_headers = _get_headers()
                    # Pass job_id to accumulate documents in same job record
                    if st.session_state.get("job_id"):
                        upload_headers["X-Job-Id"] = st.session_state["job_id"]

                    resp = requests.post(
                        f"{API_BASE}/upload",
                        files=files,
                        headers=upload_headers,
                        timeout=120,
                    )
                    resp.raise_for_status()

                    upload_data = resp.json()
                    # Store job_id from upload response
                    if upload_data.get("job_id"):
                        st.session_state.job_id = upload_data["job_id"]

                    for item in upload_data.get("uploaded", []):
                        st.session_state.uploaded_file_names.add(
                            item["filename"])
                        # Store content-detected doc type from backend
                        if item.get("detected_type"):
                            st.session_state.file_types[item["filename"]
                                                        ] = item["detected_type"]
                        # Store child name from backend
                        if item.get("child_name"):
                            st.session_state.child_names[item["filename"]
                                                         ] = item["child_name"]

                    if f.name not in st.session_state.uploaded_file_names:
                        raise RuntimeError(
                            "The backend did not confirm the upload")
                    st.session_state.upload_errors.pop(f.name, None)

                except Exception as e:
                    upload_failed = True
                    st.session_state.upload_errors[f.name] = str(e)
                    st.error(f"Upload failed for {f.name}: {e}")

        if upload_failed:
            upload_msg.error("One or more documents failed to upload")
        else:
            upload_msg.success("Documents Uploaded")

        # --- Check for child name mismatches across uploaded files ---
        _check_child_name_mismatches()
        print(f"[DEBUG] child_names: {st.session_state.child_names}")
        print(f"[DEBUG] mismatched_files: {st.session_state.mismatched_files}")

    if st.session_state.view_output:
        view_base = os.path.basename(st.session_state.view_output)
        if not any(view_base.startswith(os.path.splitext(n)[0]) for n in current_names):
            st.session_state.view_output = None
            st.session_state.view_mode = None
else:
    # All files removed — clean up blob storage via backend
    for name in list(st.session_state.uploaded_file_names):
        try:
            requests.delete(f"{API_BASE}/files/{name}",
                            headers=_get_headers(), timeout=5)
        except Exception:
            pass
    st.session_state.uploaded_file_names = set()
    st.session_state.upload_errors = {}
    st.session_state.extracted_files = {}
    st.session_state.json_cache = {}
    st.session_state.view_output = None
    st.session_state.view_mode = None
    st.session_state.processing_active = False
    st.session_state.filled_docx_path = None
    st.session_state.filled_docx_bytes = None
    st.session_state.writing_active = False
    st.session_state.writer_validation_report = None
    st.session_state.file_types = {}
    st.session_state.child_names = {}
    st.session_state.mismatched_files = set()
    st.session_state.file_store = {}

# -----------------------------------------------------
# Helper: fetch JSON from backend
# -----------------------------------------------------


def fetch_json_from_backend(filename: str) -> dict | None:
    """Download a JSON result file from the backend API."""
    if filename in st.session_state.json_cache:
        return st.session_state.json_cache[filename]
    try:
        encoded_name = quote(filename)
        resp = requests.get(
            f"{API_BASE}/results/{encoded_name}", headers=_get_headers(), timeout=30)
        if resp.status_code == 200:
            data = resp.json()
            st.session_state.json_cache[filename] = data
            return data
    except Exception:
        pass
    return None

# -----------------------------------------------------
# Show results: Dropdown + action buttons
# -----------------------------------------------------


if st.session_state.extracted_files:

    col_label, col_toggle, col_hint = st.columns([1.2, 0.4, 2.5])
    with col_label:
        st.markdown(
            "<span style='font-weight:700;font-size:1.25rem;padding-top:4px;display:inline-block;'>Extracted Files</span>", unsafe_allow_html=True)
    with col_toggle:
        debug_on = st.toggle(" ", value=st.session_state.get(
            "debug_toggle", False), key="debug_toggle")
        DEBUG_MODE = debug_on
    with col_hint:
        if not DEBUG_MODE:
            st.markdown("<span style='font-style:italic;color:#888;padding-top:6px;display:inline-block;'>(Enable toggle to view/download JSON)</span>", unsafe_allow_html=True)

    file_names = list(st.session_state.extracted_files.keys())

    if DEBUG_MODE:
        selected_file = st.selectbox(
            "Select a file", file_names, key="file_selector")
    else:
        selected_file = None

    if selected_file:
        output_filename = st.session_state.extracted_files[selected_file]

        if DEBUG_MODE:
            col1, col2 = st.columns([1, 1])

            with col1:
                if st.button("View JSON", key="view_json_btn"):
                    _log_activity("view_json", {"filename": output_filename})
                    st.session_state.view_output = output_filename
                    st.session_state.view_mode = "json"

            with col2:
                data = fetch_json_from_backend(output_filename)
                if data:
                    st.download_button(
                        label="Download JSON",
                        data=json.dumps(data, indent=4),
                        file_name=output_filename,
                        mime="application/json",
                        key="download_btn_ready",
                    )

    # -------------------------------------------------
    # Reader Validation Report (auto-displayed)
    # -------------------------------------------------

    # Map uploaded files to their doc type category
    DOC_TYPE_DISPLAY = {
        "personal": "Personal Details",
        "education": "Education Advice",
        "health": "Health Advice",
        "socialcare": "Social Care Advice",
    }
    ALL_READER_AREAS = ["Personal Details", "Education Advice",
                        "Health Advice", "Social Care Advice"]

    reader_val_data = {}
    all_reader_validation = {}
    for file_name, out_filename in st.session_state.extracted_files.items():
        val_filename = out_filename.replace("_output.json", "_validation.json")
        vdata = fetch_json_from_backend(val_filename)
        if vdata:
            section = os.path.splitext(file_name)[0]
            reader_val_data[section] = {
                "accuracy": vdata.get("accuracy_percentage", 0),
                "completeness": vdata.get("completeness_percentage", None),
                "total_fields": vdata.get("total_fields_checked", 0),
                "matched_fields": vdata.get("matched_fields", 0),
                "missing_fields": vdata.get("missing_fields", 0),
                "incorrect_fields": vdata.get("incorrect_fields", 0),
                "critical_fields_total": vdata.get("critical_fields_total", 0),
                "critical_fields_populated": vdata.get("critical_fields_populated", 0),
                "critical_fields_missing": vdata.get("critical_fields_missing", []),
                "empty_sections": vdata.get("empty_sections", []),
            }
            all_reader_validation[section] = vdata

    # Determine which doc types are uploaded and which are missing
    uploaded_doc_types = set()
    for file_name in st.session_state.extracted_files:
        detected = auto_detect_doc_type(file_name)
        doc_key = classify_file_by_type(
            st.session_state.file_types.get(file_name, detected))
        uploaded_doc_types.add(
            DOC_TYPE_DISPLAY.get(doc_key, "Personal Details"))

    missing_reader_types = [
        a for a in ALL_READER_AREAS if a not in uploaded_doc_types]

    if reader_val_data or missing_reader_types:
        # --- Enrich reader download JSON with field-level summary ---
        area_names = list(reader_val_data.keys())
        for name in area_names:
            matching_file = None
            for file_name, out_filename in st.session_state.extracted_files.items():
                if os.path.splitext(file_name)[0] == name:
                    matching_file = out_filename
                    break
            if not matching_file:
                continue
            output_data = fetch_json_from_backend(matching_file)
            if not output_data:
                continue
            filled_fields, empty_fields = count_json_fields(output_data)
            all_reader_validation[name]["field_summary"] = {
                "total_fields": len(filled_fields) + len(empty_fields),
                "filled_count": len(filled_fields),
                "empty_count": len(empty_fields),
                "filled_fields": filled_fields,
                "empty_fields": empty_fields,
                "empty_field_note": "These fields were not found in the source document or were left blank",
            }

        if DEBUG_MODE:
            st.download_button(
                label="Download Reader Validation Report",
                data=json.dumps(all_reader_validation,
                                indent=4, ensure_ascii=False),
                file_name="reader_validation_report.json",
                mime="application/json",
                key="download_reader_val_ready",
            )

    # -------------------------------------------------
    # Write to EHCP Template button (manual trigger)
    # -------------------------------------------------

    if not st.session_state.filled_docx_path:
        st.divider()
        if not st.session_state.writing_active:
            if st.button("Create Draft EHCP", type="primary", key="write_btn"):
                st.session_state.writing_active = True
                st.session_state.writer_validation_report = None
                st.session_state.filled_docx_bytes = None
                st.rerun()

        if st.session_state.writing_active:

            with st.status("Running Writer Agent Pipeline...", expanded=True) as status:
                st.write("Running agents: TemplateWriter → WriterValidator...")

                json_paths = {
                    "personal": None,
                    "education": None,
                    "health": None,
                    "socialcare": None,
                }

                for file_name, out_filename in st.session_state.extracted_files.items():
                    file_type = classify_file_by_type(
                        st.session_state.file_types.get(
                            file_name, auto_detect_doc_type(file_name))
                    )
                    json_paths[file_type] = out_filename

                try:
                    resp = requests.post(
                        f"{API_BASE}/write-ehcp",
                        json={"json_paths": json_paths,
                              "job_id": st.session_state.get("job_id")},
                        headers=_get_headers(),
                        timeout=300,
                    )
                    resp.raise_for_status()
                    writer_result = resp.json()

                    st.session_state.filled_docx_path = writer_result.get(
                        "filled_docx_path")
                    st.session_state.writer_validation_report = writer_result.get(
                        "validation_report")

                except Exception as e:
                    st.error(f"Writer pipeline error: {e}")

                status.update(
                    label="Writer Agent Pipeline complete", state="complete")

            st.session_state.writing_active = False
            st.rerun()

    # Show filled DOCX download if available
    if st.session_state.filled_docx_path:

        # Use cached DOCX bytes if available, otherwise fetch from backend
        if st.session_state.filled_docx_bytes is None:
            try:
                encoded_name = quote(st.session_state.filled_docx_path)
                resp = requests.get(
                    f"{API_BASE}/results/{encoded_name}",
                    headers=_get_headers(),
                    timeout=30,
                )
                if resp.status_code == 200:
                    st.session_state.filled_docx_bytes = resp.content
                else:
                    st.warning(
                        f"Could not fetch DOCX (status {resp.status_code}).")
            except Exception as e:
                st.warning(
                    f"Could not fetch the filled DOCX from backend: {e}")

        # -------------------------------------------------
        # Writer Validation Report (auto-displayed)
        # -------------------------------------------------

        # Prepare writer validation download data (no table displayed)
        all_writer_areas = [
            "Personal Details", "Education Advice", "Health Advice", "Social Care Advice"]

        if st.session_state.writer_validation_report:
            report_data = st.session_state.writer_validation_report

            # Enrich writer download JSON with field summary per area
            writer_field_summary = {}
            mapping_data_for_download = fetch_mapping_fields()
            for file_name, out_filename in st.session_state.extracted_files.items():
                doc_type = st.session_state.file_types.get(
                    file_name, auto_detect_doc_type(file_name))
                doc_key = classify_file_by_type(doc_type)
                area_label = DOC_TYPE_DISPLAY.get(doc_key, "Personal Details")
                output_data = fetch_json_from_backend(out_filename)
                if output_data:
                    if mapping_data_for_download and doc_key in mapping_data_for_download:
                        filled_fields, empty_fields = count_mapping_fields(
                            output_data, mapping_data_for_download[doc_key])
                    else:
                        filled_fields, empty_fields = count_json_fields(
                            output_data)
                    writer_field_summary[area_label] = {
                        "source_file": file_name,
                        "total_fields": len(filled_fields) + len(empty_fields),
                        "filled_count": len(filled_fields),
                        "empty_count": len(empty_fields),
                        "filled_fields": filled_fields,
                        "empty_fields": empty_fields,
                        "empty_field_note": "These fields were not found in the source document or were left blank",
                    }
            for area_w in all_writer_areas:
                if area_w not in writer_field_summary:
                    writer_field_summary[area_w] = {"status": "Not uploaded"}

            writer_download_data = dict(report_data)
            writer_download_data["field_summary"] = writer_field_summary

            if DEBUG_MODE:
                st.download_button(
                    label="Download Writer Validation Report",
                    data=json.dumps(writer_download_data,
                                    indent=4, ensure_ascii=False),
                    file_name="writer_validation_report.json",
                    mime="application/json",
                    key="download_writer_val_ready",
                )

        # -------------------------------------------------
        # Validator Agent: EHCP Output Field Summary
        # -------------------------------------------------

        st.divider()
        st.markdown("<div style='font-weight:700;font-size:1.5rem;margin:0;padding:0;'>Validator Agent: EHCP Output Field Summary</div>", unsafe_allow_html=True)

        # Fetch mapping fields from backend (defines what appears in the output EHCP doc)
        mapping_fields_data = fetch_mapping_fields()

        # --- Collect per-area stats for accuracy table ---
        area_summary_data = {}

        for area in all_writer_areas:
            area_key = {v: k for k, v in DOC_TYPE_DISPLAY.items()}.get(area)
            source_file = None
            out_filename = None
            val_filename = None
            for file_name, of in st.session_state.extracted_files.items():
                detected = auto_detect_doc_type(file_name)
                ftype = st.session_state.file_types.get(file_name, detected)
                fkey = classify_file_by_type(ftype)
                if fkey == area_key:
                    source_file = file_name
                    out_filename = of
                    val_filename = of.replace(
                        "_output.json", "_validation.json")
                    break

            if not source_file:
                area_summary_data[area] = None
                continue

            output_data = fetch_json_from_backend(
                out_filename) if out_filename else None
            filled_fields, empty_fields = ([], [])
            if output_data:
                # Use mapping-based counting if available, fall back to JSON traversal
                if mapping_fields_data and area_key in mapping_fields_data:
                    filled_fields, empty_fields = count_mapping_fields(
                        output_data, mapping_fields_data[area_key])
                else:
                    filled_fields, empty_fields = count_json_fields(
                        output_data)

            total = len(filled_fields) + len(empty_fields)
            filled = len(filled_fields)
            missing = len(empty_fields)
            accuracy = round((filled / total * 100)) if total > 0 else 0

            area_summary_data[area] = {
                "source_file": source_file,
                "filled_fields": filled_fields,
                "empty_fields": empty_fields,
                "total": total,
                "filled": filled,
                "missing": missing,
                "accuracy": accuracy,
            }

        # --- Pivoted accuracy summary table ---
        # Columns = doc types, Rows = metrics (Accuracy, Total, Filled, Missing)
        areas_with_data = [
            a for a in all_writer_areas if area_summary_data.get(a) is not None]

        acc_table = []
        acc_table.append(
            '<table style="border-collapse:collapse; font-family:Arial, sans-serif; font-size:14px; margin-top:10px; width:100%;">')
        # Header row: empty corner + doc type columns
        acc_table.append(
            '<thead><tr style="background-color:#236c34; color:white;">')
        acc_table.append(
            '<th style="padding:10px 14px; text-align:left; border:1px solid #ddd;"></th>')
        for area in areas_with_data:
            acc_table.append(
                f'<th style="padding:10px 14px; text-align:center; border:1px solid #ddd;">{area}</th>')
        acc_table.append('</tr></thead><tbody>')

        # Source File row
        acc_table.append('<tr style="background-color:#f9f9f9;">')
        acc_table.append(
            '<td style="padding:8px 14px; border:1px solid #ddd; font-weight:600;">Source File</td>')
        for area in areas_with_data:
            info = area_summary_data[area]
            acc_table.append(
                f'<td style="padding:8px 14px; border:1px solid #ddd; text-align:center; font-size:12px;">{info["source_file"]}</td>')
        acc_table.append('</tr>')

        # Total Fields row
        acc_table.append('<tr style="background-color:#fff;">')
        acc_table.append(
            '<td style="padding:8px 14px; border:1px solid #ddd; font-weight:600;">Total Fields</td>')
        for area in areas_with_data:
            info = area_summary_data[area]
            acc_table.append(
                f'<td style="padding:8px 14px; border:1px solid #ddd; text-align:center;">{info["total"]}</td>')
        acc_table.append('</tr>')

        # Filled row
        acc_table.append('<tr style="background-color:#f9f9f9;">')
        acc_table.append(
            '<td style="padding:8px 14px; border:1px solid #ddd; font-weight:600;">Filled</td>')
        for area in areas_with_data:
            info = area_summary_data[area]
            acc_table.append(
                f'<td style="padding:8px 14px; border:1px solid #ddd; text-align:center; color:#1B8A2A; font-weight:600;">{info["filled"]}</td>')
        acc_table.append('</tr>')

        # Missing row
        acc_table.append('<tr style="background-color:#fff;">')
        acc_table.append(
            '<td style="padding:8px 14px; border:1px solid #ddd; font-weight:600;">Missing</td>')
        for area in areas_with_data:
            info = area_summary_data[area]
            m_color = "#D32F2F" if info["missing"] > 0 else "#1B8A2A"
            acc_table.append(
                f'<td style="padding:8px 14px; border:1px solid #ddd; text-align:center; color:{m_color}; font-weight:600;">{info["missing"]}</td>')
        acc_table.append('</tr>')

        # Accuracy row
        acc_table.append('<tr style="background-color:#f9f9f9;">')
        acc_table.append(
            '<td style="padding:8px 14px; border:1px solid #ddd; font-weight:600;">Completed</td>')
        for area in areas_with_data:
            info = area_summary_data[area]
            acc_val = info["accuracy"]
            acc_color = "#1B8A2A" if acc_val >= 90 else (
                "#E6A817" if acc_val >= 70 else "#D32F2F")
            acc_table.append(
                f'<td style="padding:8px 14px; border:1px solid #ddd; text-align:center; font-weight:700; font-size:16px; color:{acc_color};">{acc_val}%</td>')
        acc_table.append('</tr>')

        acc_table.append('</tbody></table>')
        acc_table.append(
            '<p style="font-size:12px; color:#888; margin-top:4px;">Completed = Filled / Total.</p>')
        st.markdown(''.join(acc_table), unsafe_allow_html=True)

        st.write("")

        # --- Per-area detail cards ---
        for area in all_writer_areas:
            info = area_summary_data.get(area)
            if info is None:
                continue

            source_file = info["source_file"]
            filled_fields = info["filled_fields"]
            empty_fields = info["empty_fields"]
            filled = info["filled"]
            total = info["total"]
            missing = info["missing"]
            accuracy = info["accuracy"]

            pct_color = "#1B8A2A" if accuracy >= 90 else (
                "#E6A817" if accuracy >= 70 else "#D32F2F")
            bar_pct = round(filled / total * 100) if total > 0 else 0
            bar_bg = "#e0e0e0"

            card_html = (
                f'<div style="background:#fff; border:1px solid #ddd; border-radius:8px; '
                f'padding:16px 20px; margin-bottom:14px;">'
                f'<div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:6px;">'
                f'<span style="font-weight:700; font-size:16px;">{area}</span>'
                f'<span style="font-size:13px; color:#666;">Source: <b>{source_file}</b></span>'
                f'</div>'
                f'<div style="background:{bar_bg}; border-radius:6px; height:22px; width:100%; margin-bottom:8px; overflow:hidden;">'
                f'<div style="background:{pct_color}; height:100%; width:{bar_pct}%; border-radius:6px; '
                f'transition:width 0.3s;"></div>'
                f'</div>'
                f'<div style="display:flex; gap:24px; font-size:14px; color:#444;">'
                f'<span><b>{filled}</b> / {total} fields filled</span>'
                f'<span style="color:{pct_color}; font-weight:700;">{accuracy}% completed</span>'
            )
            if missing > 0:
                card_html += f'<span style="color:#D32F2F;"><b>{missing}</b> missing</span>'
            card_html += '</div>'

            if empty_fields:
                card_html += (
                    '<div style="margin-top:10px; padding-top:8px; border-top:1px solid #eee;">'
                    '<span style="font-size:13px; font-weight:600; color:#D32F2F;">Missing fields:</span>'
                    '<ul style="margin:4px 0 0 18px; padding:0; font-size:13px; color:#555;">'
                )
                for ef in empty_fields:
                    display_name = ef.replace("_", " ").replace(".", " → ")
                    card_html += f'<li>{display_name}</li>'
                card_html += '</ul></div>'

            card_html += '</div>'
            st.markdown(card_html, unsafe_allow_html=True)

        # -------------------------------------------------
        # Final EHCP Document: Download + Start New Case
        # -------------------------------------------------

        st.divider()
        st.markdown(
            "<div style='font-weight:700;font-size:1.5rem;margin:0 0 1rem 0;padding:0;'>Final EHCP Document</div>", unsafe_allow_html=True)

        if st.session_state.filled_docx_bytes:
            col_dl, col_new = st.columns([1, 1])
            with col_dl:
                if not st.session_state.get("certified_download_docx"):
                    if st.button("Download Draft EHCP", key="download_filled_docx_btn"):
                        _certify_download_dialog("download_docx")
                else:
                    st.success("You are verified, you can download")
                    st.download_button(
                        label="⬇ Download Draft EHCP",
                        data=st.session_state.filled_docx_bytes,
                        file_name=st.session_state.filled_docx_path,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        key="download_filled_docx_ready",
                    )
            with col_new:
                if st.button("Start New Case", key="start_new_case_btn"):
                    st.session_state.confirm_new_case = True

        # Confirmation dialog for Start New Case
        if st.session_state.get("confirm_new_case"):
            st.warning(
                "**Are you sure you want to start a new case?** All current data will be cleared.")
            col_yes, col_no = st.columns([1, 1])
            with col_yes:
                if st.button("Yes, start new case", key="confirm_new_yes"):
                    # Increment uploader key to force fresh widget
                    new_uploader_key = st.session_state.get(
                        "uploader_key", 0) + 1
                    # Clear all session state except auth
                    for key in list(st.session_state.keys()):
                        if key not in ("user_name", "user_email", "access_token"):
                            del st.session_state[key]
                    st.session_state.uploader_key = new_uploader_key
                    st.rerun()
            with col_no:
                if st.button("Cancel", key="confirm_new_no"):
                    st.session_state.confirm_new_case = False
                    st.rerun()

# -----------------------------------------------------
# Show JSON or Validation when clicked
# -----------------------------------------------------

if st.session_state.view_output:
    if st.session_state.view_mode == "json":
        st.divider()
        st.write(f"### Viewing: {st.session_state.view_output}")

        output_data = fetch_json_from_backend(st.session_state.view_output)
        if output_data:
            st.json(output_data)
        else:
            st.warning("Could not fetch JSON from backend.")

# -----------------------------------------------------
# Analyze & Process: Upload to backend + trigger pipeline
# -----------------------------------------------------

if uploaded_files:

    # Exclude mismatched files from pending ONLY if user has NOT overridden
    mismatched = st.session_state.mismatched_files
    if st.session_state.get("_mismatch_override"):
        # User clicked "Yes, proceed with all files" — include everything
        pending = [f for f in uploaded_files
                   if f.name not in st.session_state.extracted_files]
    else:
        pending = [f for f in uploaded_files
                   if f.name not in st.session_state.extracted_files
                   and f.name not in mismatched]

    # --- Detect and skip duplicate doc types ---
    # Only keep the first file per doc type; flag extras as duplicates
    duplicate_files = set()
    if pending:
        seen_doc_types = {}
        # Also consider already-extracted files' doc types as "taken"
        for fn in st.session_state.extracted_files:
            dt = st.session_state.file_types.get(fn, auto_detect_doc_type(fn))
            dk = classify_file_by_type(dt)
            seen_doc_types[dk] = fn

        for f in pending:
            dt = st.session_state.file_types.get(
                f.name, auto_detect_doc_type(f.name))
            dk = classify_file_by_type(dt)
            if dk in seen_doc_types:
                duplicate_files.add(f.name)
            else:
                seen_doc_types[dk] = f.name

        pending = [f for f in pending if f.name not in duplicate_files]

    # --- Show child name mismatch warnings (only before extraction, not after writer) ---
    block_analyse = False  # Track if we should disable the button

    if mismatched and pending and not st.session_state.filled_docx_path:
        # Get the primary name from Personal Details
        file_types = st.session_state.get("file_types", {})
        primary_child = None
        for fname, cname in st.session_state.child_names.items():
            dt = file_types.get(fname, auto_detect_doc_type(fname))
            if classify_file_by_type(dt) == "personal":
                primary_child = cname
                break
        if not primary_child:
            from collections import Counter
            name_counts = Counter(st.session_state.child_names.values())
            primary_child = name_counts.most_common(
                1)[0][0] if name_counts else "Unknown"

        # Build message showing correct and mismatched files
        all_uploaded_names = set(f.name for f in uploaded_files)
        correct_files = [fn for fn in sorted(
            all_uploaded_names) if fn not in mismatched]
        correct_lines = []
        for fn in correct_files:
            detected = st.session_state.child_names.get(fn)
            if detected:
                correct_lines.append(f"- ✅ **{fn}** → {detected}")
            else:
                correct_lines.append(
                    f"- ✅ **{fn}** *(name not detected — will be processed)*")

        mismatch_lines = []
        mismatched_file_list = []
        for mf in sorted(mismatched):
            other_name = st.session_state.child_names.get(mf, "Unknown")
            mismatch_lines.append(f"- ❌ **{mf}** → {other_name}")
            mismatched_file_list.append(mf)

        # Check if user already confirmed the mismatch override
        if not st.session_state.get("_mismatch_override"):
            st.warning(
                f"⚠️ **The child name in Personal Details is: {primary_child}**\n\n"
                "**Matching files:**\n\n"
                + "\n".join(correct_lines)
                + "\n\n**Different name detected:**\n\n"
                + "\n".join(mismatch_lines)
                + "\n\nAre you sure you want to proceed?"
            )
            col_yes, col_no = st.columns(2)
            with col_yes:
                if st.button("Yes, proceed with all files", key="mismatch_proceed"):
                    st.session_state._mismatch_override = True
                    st.rerun()
            with col_no:
                if st.button("No, I'll fix the uploads", key="mismatch_reject"):
                    st.session_state._mismatch_rejected = True
                    st.rerun()

            if st.session_state.get("_mismatch_rejected"):
                mismatched_names = ", ".join(mismatched_file_list)
                st.error(
                    f"Please check and re-upload the correct document(s): **{mismatched_names}**"
                )
                st.session_state.pop("_mismatch_rejected", None)

            block_analyse = True

    # --- Check for missing required document types ---
    all_file_types = set()
    for f in uploaded_files:
        dt = st.session_state.file_types.get(
            f.name, auto_detect_doc_type(f.name))
        all_file_types.add(classify_file_by_type(dt))

    required_types = {"personal"}
    missing_types = required_types - all_file_types
    if missing_types and not st.session_state.filled_docx_path:
        st.error(
            "🚫 **Missing required document: Personal Details.**\n\n"
            "A Personal Details document is required to create a Draft EHCP. "
            "Please upload it before proceeding."
        )
        block_analyse = True

    # --- Check if more than 4 files uploaded ---
    if len(uploaded_files) > 4 and not st.session_state.filled_docx_path:
        st.error(
            "🚫 **Too many files uploaded.** "
            "Only 4 documents are allowed (Personal Details, Education Advice, Health Advice, Social Care Advice). "
            "Please remove the extra file(s) before proceeding."
        )
        block_analyse = True

    # --- Check for unrecognized/wrong file types ---
    valid_types = set(DOC_TYPE_KEY_MAP.keys()) | set(DOC_TYPE_KEY_MAP.values())
    wrong_files = []
    for f in uploaded_files:
        dt = st.session_state.file_types.get(
            f.name, auto_detect_doc_type(f.name))
        if dt not in valid_types:
            wrong_files.append(f.name)
    if wrong_files and not st.session_state.filled_docx_path:
        wrong_list = ", ".join(sorted(wrong_files))
        st.error(
            f"🚫 **Unrecognized file(s): {wrong_list}**\n\n"
            "Only Personal Details, Education Advice, Health Advice, and Social Care Advice documents are accepted. "
            "Please remove the incorrect file(s) before proceeding."
        )
        block_analyse = True

    # --- Show duplicate doc type warnings ---
    if duplicate_files and pending and not st.session_state.filled_docx_path:
        dup_lines = []
        for fn in sorted(duplicate_files):
            dt = st.session_state.file_types.get(fn, auto_detect_doc_type(fn))
            dup_lines.append(
                f"- **{fn}** → {dt} *(duplicate — will be skipped)*")
        st.warning(
            "**Duplicate document type detected.**\n\n"
            + "\n".join(dup_lines)
            + "\n\nOnly the first file per document type will be processed."
        )

    if pending and not st.session_state.filled_docx_path and not st.session_state.extracted_files:

        if not st.session_state.processing_active:
            if block_analyse:
                st.button("Read and Analyse", type="primary", disabled=True)
            else:
                if st.button("Read and Analyse", type="primary"):
                    st.session_state.processing_active = True
                    st.rerun()

        if st.session_state.processing_active:
            # Stop pipeline if validations now fail (e.g. files changed)
            if block_analyse:
                st.session_state.processing_active = False
                st.rerun()

            total = len(pending)
            st.write(
                f"**Processing {total} file{'s' if total > 1 else ''} in parallel...**")

            # =============================================
            # Step 1: Upload any files not yet uploaded
            # =============================================

            not_yet_uploaded = [
                f for f in pending if f.name not in st.session_state.uploaded_file_names]
            if not_yet_uploaded:
                with st.status("Uploading files to backend...", expanded=True) as status:
                    for file in not_yet_uploaded:
                        file.seek(0)
                        st.write(f"Uploading {file.name}...")
                        files = {
                            "files": (file.name, file, "application/octet-stream")}
                        try:
                            upload_headers = _get_headers()
                            if st.session_state.get("job_id"):
                                upload_headers["X-Job-Id"] = st.session_state["job_id"]
                            resp = requests.post(
                                f"{API_BASE}/upload", files=files, headers=upload_headers, timeout=120)
                            resp.raise_for_status()
                            upload_data = resp.json()
                            # Store job_id from upload response
                            if upload_data.get("job_id"):
                                st.session_state.job_id = upload_data["job_id"]
                            for item in upload_data.get("uploaded", []):
                                st.write(
                                    f"Uploaded {item['filename']} (detected: {item['detected_type']})")
                                st.session_state.uploaded_file_names.add(
                                    item["filename"])
                                # Store content-detected doc type from backend
                                if item.get("detected_type"):
                                    st.session_state.file_types[item["filename"]
                                                                ] = item["detected_type"]
                                # Store child name from backend
                                if item.get("child_name"):
                                    st.session_state.child_names[item["filename"]
                                                                 ] = item["child_name"]
                        except Exception as e:
                            st.error(f"Upload failed for {file.name}: {e}")
                            st.session_state.processing_active = False
                            st.stop()
                    status.update(label="All files uploaded", state="complete")
            # =============================================
            # Step 2: Run MAF Pipeline via backend API (with streaming progress)
            # =============================================

            with st.status("Running Multi-Agent Pipeline...", expanded=True) as status:
                st.write(
                    "Running agents: DocumentReader → Extractor → Validator → QualityChecker...")

                maf_progress_placeholder = st.empty()
                maf_progress_lines = []

                AGENT_STAGE_LABELS = {
                    "agent_reader": "DocumentReaderAgent",
                    "agent_extractor": "ExtractorAgent",
                    "agent_validator": "ValidatorAgent",
                    "agent_quality": "QualityCheckerAgent",
                }

                file_configs = []
                for file in pending:
                    doc_type = st.session_state.file_types.get(
                        file.name, auto_detect_doc_type(file.name)
                    )
                    file_configs.append({
                        "filename": file.name,
                        "doc_type": doc_type,
                    })

                pipeline_started_at = time.perf_counter()
                analyze_result = None

                try:
                    resp = requests.post(
                        f"{API_BASE}/analyze-stream",
                        json={"files": file_configs,
                              "job_id": st.session_state.get("job_id")},
                        headers=_get_headers(),
                        timeout=600,
                        stream=True,
                    )
                    resp.raise_for_status()

                    for line in resp.iter_lines(decode_unicode=True):
                        if line and line.startswith("data: "):
                            event = json.loads(line[6:])

                            # Final result message
                            if event.get("type") == "complete":
                                analyze_result = {"results": event["results"]}
                                continue
                            if event.get("type") == "error":
                                st.error(f"Pipeline error: {event['error']}")
                                st.session_state.processing_active = False
                                st.stop()

                            # Progress event
                            file_label = os.path.basename(
                                event.get("file_path") or "")
                            stage = event.get("stage", "")
                            action = event.get("event")
                            elapsed = event.get("elapsed_seconds")
                            agent_name = event.get(
                                "agent") or AGENT_STAGE_LABELS.get(stage, stage)

                            if action == "start":
                                maf_progress_lines.append(
                                    f"{file_label}: [{agent_name}] started")
                            elif action == "done":
                                maf_progress_lines.append(
                                    f"{file_label}: [{agent_name}] finished in {elapsed:.1f}s")
                            elif action == "error":
                                maf_progress_lines.append(
                                    f"{file_label}: [{agent_name}] failed: {event.get('error')}")

                            maf_progress_placeholder.markdown(
                                "\n".join(
                                    f"- {line}" for line in maf_progress_lines) if maf_progress_lines else ""
                            )

                except Exception as e:
                    st.error(f"Pipeline error: {e}")
                    st.session_state.processing_active = False
                    st.stop()

                # Fallback: if streaming didn't return results, call regular endpoint
                if analyze_result is None:
                    try:
                        resp = requests.post(
                            f"{API_BASE}/analyze",
                            json={"files": file_configs},
                            headers=_get_headers(),
                            timeout=600,
                        )
                        resp.raise_for_status()
                        analyze_result = resp.json()
                    except Exception as e:
                        st.error(f"Pipeline error: {e}")
                        st.session_state.processing_active = False
                        st.stop()

                elapsed = time.perf_counter() - pipeline_started_at
                maf_progress_lines.append(
                    f"Multi-Agent Pipeline finished in {elapsed:.1f}s")
                maf_progress_placeholder.markdown(
                    "\n".join(
                        f"- {l}" for l in maf_progress_lines) if maf_progress_lines else ""
                )

                for result in analyze_result.get("results", []):
                    if result["success"]:
                        st.write(f"  {result['filename']} — Done")
                    else:
                        st.write(
                            f"  {result['filename']} — Failed: {result.get('error', 'unknown')}")

                status.update(
                    label="Multi-Agent Pipeline complete", state="complete")

            # =============================================
            # Store results in session state
            # =============================================

            for result in analyze_result.get("results", []):
                if result["success"]:
                    st.session_state.extracted_files[result["filename"]
                                                     ] = result["output_file"]
                else:
                    st.error(
                        f"Error processing {result['filename']}: {result.get('error')}")

            st.success(f"All {total} files processed by Multi-Agent Pipeline!")

            st.session_state.processing_active = False
            st.rerun()

    else:
        st.session_state.processing_active = False
