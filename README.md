# EHCP AI Drafting

An Azure-based, multi-agent solution that turns the professional advice documents used in the
**Education, Health and Care Plan (EHCP)** process into a **draft EHCP document**.

Case officers upload the source advice files (Personal Details, Education Advice, Health Advice,
Social Care Advice) in DOCX or PDF format. The solution extracts the text, structures it into JSON
against per-section schemas, validates and quality-checks the extraction, and then fills the
statutory EHCP DOCX template using a spreadsheet-driven field mapping. The result is a draft EHCP
that a human reviews and finalises — the system is a drafting assistant, **not** a decision maker.

---
**NOTE on Location** :- This solution uses swedencentral as foundry resource region and Global standard deployment type which means data wil be stored in swedencentral but can be processed outside EU. Please choose the location carefully depending on your data residency requirements.

**NOTE on Deployment scripts** :- Please try running solution and deployment scripts only in a sandbox/test env.**(Not for Production use)**. We are working on making solution hardening and deployment as simple as possible and will notify you once the repo get updates.

## Table of contents

- [Key capabilities](#key-capabilities)
- [Technical architecture](#technical-architecture)
- [Azure components used](#azure-components-used)
- [Repository layout](#repository-layout)
- [How the pipelines work](#how-the-pipelines-work)
- [Configuration reference](#configuration-reference)
- [Running locally](#running-locally)
- [Deploying to Azure](#deploying-to-azure)
- [API reference](#api-reference)
- [Customising the solution](#customising-the-solution)
- [Test cases](#test-cases)
- [Security, privacy and responsible AI](#security-privacy-and-responsible-ai)
- [Troubleshooting](#troubleshooting)

---

## Key capabilities

| Capability | Description |
|---|---|
| Multi-format ingest | DOCX and PDF advice documents |
| Automatic document typing | File name heuristics plus content-marker scoring classify each upload as Personal / Education / Health / Social Care advice |
| Structured extraction | Azure OpenAI extracts JSON that conforms to a per-section JSON schema |
| LLM + rule-based validation | An LLM validator scores extraction accuracy; a deterministic quality checker re-checks fields and computes completeness over critical fields |
| Template writing | The EHCP DOCX template is filled from the extracted JSON using an Excel mapping workbook |
| Writer validation | Deterministic checks compare the filled DOCX against the JSONs, the mapping workbook and (optionally) an expected output |
| Live progress | Server-sent events stream per-agent progress to the UI |
| Session isolation | Every browser session gets its own temp/output directory on the backend; downloads are scoped to the session |
| Auditability | Per-action and per-job records written to Azure Cosmos DB; token usage tracked per run |
| Enterprise auth | Optional Microsoft Entra ID sign-in (MSAL auth-code flow in the UI, JWT validation in the API) |
| Keyless operation | Optional managed identity for Azure OpenAI, Document Intelligence, Blob Storage and Cosmos DB |

---

## Technical architecture

```
                     ┌──────────────────────────────────────────────┐
                     │            Microsoft Entra ID                │
                     │  (frontend app reg + backend API app reg)    │
                     └───────────────┬──────────────────────────────┘
                                     │ OAuth2 auth-code flow / JWT
                                     ▼
  Browser ──HTTPS──►  ┌──────────────────────────┐   internal HTTPS   ┌───────────────────────────┐
                      │  Frontend Container App  │ ─────────────────► │  Backend Container App    │
                      │  Streamlit (port 8501)   │  Authorization +   │  FastAPI + Uvicorn (8000) │
                      │  external ingress        │   X-Session-ID     │  internal ingress         │
                      └──────────────────────────┘                    └─────────┬─────────────────┘
                                                                                │
                        ┌───────────────────────────────────────────────────────┼───────────────────────────────┐
                        │                        │                    │                     │                   │
                        ▼                        ▼                    ▼                     ▼                   ▼
              ┌───────────────────┐   ┌────────────────────┐  ┌───────────────┐   ┌──────────────────┐  ┌──────────────┐
              │ Azure OpenAI      │   │ Azure AI Document  │  │ Azure Blob    │   │ Azure Cosmos DB  │  │ Azure        │
              │ (chat deployment) │   │ Intelligence       │  │ Storage       │   │ (activity-logs,  │  │ Container    │
              │ extraction +      │   │ prebuilt-layout    │  │ uploads +     │   │  job-logs)       │  │ Registry     │
              │ validation        │   │ OCR / layout       │  │ outputs       │   │ audit trail      │  │ images       │
              └───────────────────┘   └────────────────────┘  └───────────────┘   └──────────────────┘  └──────────────┘
```

### Application layers

1. **Presentation — Streamlit (`frontend/`)**
   Single-page app (`app.py`) for upload, document-type confirmation, live analysis progress,
   accuracy/completeness dashboards and draft download. `auth.py` implements the MSAL confidential
   client authorization-code flow. Streamlit runtime settings live in `.streamlit/config.toml`.

2. **API — FastAPI (`backend/main.py`, `backend/app/routers/pipeline.py`)**
   All routes are exposed under the `/api` prefix. CORS is enabled, requests carry an
   `X-Session-ID` header, and every route depends on `get_current_user` for authentication and
   audit attribution.

3. **Orchestration (`backend/app/services/orchestrator.py`)**
   `EHCPAgentOrchestrator` runs the reader pipeline per file (files are processed in parallel with
   `asyncio`) and the writer pipeline once per case. For determinism and cost control the
   orchestrator invokes the agents' tool functions directly rather than letting the model choose
   tools, while the agent definitions remain available for agent-driven execution.

4. **Agents (`backend/app/services/agents.py`)**
   Built on the **Microsoft Agent Framework (MAF)** — `agent_framework.Agent`, `@tool` functions
   and `agent_framework.openai.OpenAIChatCompletionClient` bound to an Azure OpenAI deployment.

5. **Helpers (`backend/app/services/helpers/`)**
   - `reader_helpers.py` — Document Intelligence and PyMuPDF/mammoth text extraction, LLM
     extraction calls, token tracking
   - `validation_helpers.py` — re-check rules and completeness scoring
   - `template_filler.py` — DOCX template population driven by `ehcp_mapping.xlsx`
   - `writer_validation.py` — deterministic post-write validation and JSON report building

6. **Platform services** — `blob_storage.py` (durable file storage), `audit_logger.py`
   (per-action logs), `job_logger.py` (one consolidated record per case), `settings.py`
   (configuration and credential helpers), `auth.py` (Entra ID JWT validation).

### Agents

| # | Agent | Tool | Purpose |
|---|---|---|---|
| 1 | `DocumentReaderAgent` | `read_document` | Extract raw text/layout from DOCX or PDF |
| 2 | `ExtractorAgent` | `extract_to_json` | Produce schema-conformant JSON with Azure OpenAI |
| 3 | `ValidatorAgent` | `validate_extraction` | LLM comparison of JSON against source text, yields an accuracy percentage |
| 4 | `QualityCheckerAgent` | `recheck_validation` | Rule-based correction of false negatives plus completeness scoring |
| 5 | `TemplateWriterAgent` | `fill_template` | Fill the EHCP DOCX template from the four JSONs via the mapping workbook |
| 6 | `WriterValidatorAgent` | `validate_writer_output` | Deterministic validation of the filled DOCX and mapping coverage |

---

## Azure components used

| Azure service | Role in the solution | Where it is configured |
|---|---|---|
| **Azure OpenAI Service** | Chat-completion deployment (e.g. `gpt-4o`) used for structured extraction and LLM validation. Accessed through MAF's `OpenAIChatCompletionClient` and the `openai.AzureOpenAI` SDK. | `AZURE_OPENAI_*` in `backend/app/settings.py` |
| **Microsoft Agent Framework (MAF)** | The `agent-framework` Python package that defines agents, tools and the chat client abstraction used by every pipeline stage. | `backend/app/services/agents.py`, `backend/requirements.txt` |
| **Azure AI Document Intelligence** | `prebuilt-layout` model for OCR and layout-aware text/table extraction from scanned or complex PDFs and DOCX files. | `AZURE_DOCUMENT_INTELLIGENCE_*` |
| **Azure Blob Storage** | Optional durable store for uploaded source files and generated outputs so container replicas remain stateless and restart-safe. | `AZURE_STORAGE_*` |
| **Azure Cosmos DB (NoSQL)** | Audit trail. `activity-logs` container records individual user actions; `job-logs` records one document per case covering upload → analyse → create EHCP, including token usage and completeness. | `COSMOS_DB_*`, `AUDIT_LOG_ENABLED` |
| **Microsoft Entra ID** | Sign-in for the Streamlit app (MSAL confidential client) and JWT bearer validation for the FastAPI backend, using separate frontend and backend app registrations. | `ENTRA_*`, `AUTH_ENABLED` |
| **Azure Container Registry (ACR)** | Stores the `ehcp-backend` and `ehcp-frontend` container images. | `build-push.ps1` |
| **Azure Container Apps (ACA)** | Hosts both containers in one managed environment: frontend with external ingress, backend with internal-only ingress; secrets injected as Container App secrets. | `deploy-aca.ps1` |
| **Managed Identity** | Optional keyless access to Azure OpenAI, Document Intelligence, Blob Storage and Cosmos DB via `DefaultAzureCredential`. Supports user-assigned identities through `AZURE_CLIENT_ID`. | `USE_MANAGED_IDENTITY` |

---

## Repository layout

```
.
├── backend/
│   ├── main.py                       # FastAPI application entry point
│   ├── Dockerfile                    # Python 3.11-slim backend image
│   ├── requirements.txt
│   ├── .env.example                  # Backend configuration template
│   ├── EHCP_LCC_Template.docx        # Statutory EHCP output template
│   ├── ehcp_mapping.xlsx             # JSON field → template placeholder mapping
│   ├── prompts/                      # Per-section extraction + validation prompts
│   ├── schemas/                      # Per-section JSON schemas
│   └── app/
│       ├── settings.py               # Env config + credential helpers
│       ├── auth.py                   # Entra ID JWT validation
│       ├── dependencies.py           # Ensures temp/ and output/ exist
│       ├── models/schemas.py         # Pydantic request/response models
│       ├── routers/pipeline.py       # All /api routes
│       └── services/
│           ├── agents.py             # MAF agents + tools
│           ├── orchestrator.py       # Reader and writer pipelines
│           ├── blob_storage.py       # Azure Blob Storage helpers
│           ├── audit_logger.py       # Cosmos DB action logging
│           ├── job_logger.py         # Cosmos DB job-level logging
│           └── helpers/              # Extraction, validation, template filling
├── frontend/
│   ├── app.py                        # Streamlit UI
│   ├── auth.py                       # MSAL auth-code flow
│   ├── Dockerfile                    # Streamlit image
│   ├── requirements.txt
│   └── .streamlit/config.toml
├── Test Cases/                       # Sample inputs, blank templates, expected outputs
├── build-push.ps1                    # Build + push both images to ACR
└── deploy-aca.ps1                    # Create/update both Azure Container Apps
```

---

## How the pipelines work

### Reader pipeline (per uploaded file, run in parallel)

1. **Upload** — `POST /api/upload` saves files into a session-scoped temp directory, optionally
   mirrors them to Blob Storage, auto-detects the document type (file name heuristics first, then
   content-marker scoring on the first ~5000 characters) and records the upload in the job log.
2. **Read** — text and layout are extracted with Azure AI Document Intelligence
   (`prebuilt-layout`), with PyMuPDF/mammoth used for direct text extraction where appropriate.
   The extracted text is written to `<name>_doctext.txt`.
3. **Extract** — the per-type prompt (`prompts/*.txt`) and JSON schema (`schemas/*.json`) are sent
   to Azure OpenAI; the structured result is written to `<name>_output.json`.
4. **Validate** — `prompts/validation_prompt.txt` asks the model to compare the JSON with the
   source text and return field-level correctness plus an `accuracy_percentage`.
5. **Quality check** — deterministic re-check rules correct known false negatives, and
   completeness is computed over the critical fields for that document type. The validation JSON is
   overwritten with the final accuracy, completeness and missing-field details.

Progress for every stage is streamed to the UI through `POST /api/analyze-stream` (SSE).

### Writer pipeline (once per case)

1. **Fill template** — `fill_template` loads `EHCP_LCC_Template.docx` and the `ehcp_mapping.xlsx`
   workbook, resolves each mapped JSON path from the four section JSONs, and writes the completed
   DOCX to the session output directory.
2. **Validate output** — `validate_writer_output` re-opens the DOCX, validates each section JSON,
   checks the mapping workbook headers and per-sheet mappings, optionally compares against an
   expected output document, and emits a JSON validation report with a check summary.
3. **Download** — the draft is retrieved with `GET /api/results/{filename}`, restricted to the
   requesting session.

---

## Configuration reference

### Backend (`backend/.env`, template in `backend/.env.example`)

| Variable | Default | Description |
|---|---|---|
| `USE_MANAGED_IDENTITY` | `false` | Use `DefaultAzureCredential` instead of keys for all Azure data services |
| `AZURE_CLIENT_ID` | – | Client ID of a **user-assigned** managed identity (optional) |
| `AZURE_OPENAI_ENDPOINT` | – | e.g. `https://<name>.openai.azure.com/` |
| `AZURE_OPENAI_API_KEY` | – | Only when `USE_MANAGED_IDENTITY=false` |
| `AZURE_OPENAI_DEPLOYMENT` | – | Chat deployment name, e.g. `gpt-4o` |
| `AZURE_OPENAI_API_VERSION` | – | e.g. `2025-04-01-preview` |
| `MODEL_TEMPERATURE` | `0` | Deterministic extraction is recommended |
| `MODEL_MAX_TOKENS` | `30` | Token cap for the short auxiliary completion used to infer the child's name; extraction and validation calls are not capped by this value |
| `AZURE_DOCUMENT_INTELLIGENCE_ENDPOINT` | – | Document Intelligence endpoint |
| `AZURE_DOCUMENT_INTELLIGENCE_KEY` | – | Only when `USE_MANAGED_IDENTITY=false` |
| `AZURE_STORAGE_CONNECTION_STRING` | – | Blob access with keys (optional feature) |
| `AZURE_STORAGE_ACCOUNT_URL` | – | Required for blob access with managed identity |
| `AZURE_STORAGE_CONTAINER_NAME` | `ehcp-outputs` | Blob container name |
| `AUDIT_LOG_ENABLED` | `false` | Enable Cosmos DB audit and job logging |
| `COSMOS_DB_ENDPOINT` | – | e.g. `https://<account>.documents.azure.com:443/` |
| `COSMOS_DB_KEY` | – | Only when `USE_MANAGED_IDENTITY=false` |
| `COSMOS_DB_DATABASE` | `ehcp-audit` | Database name |
| `COSMOS_DB_CONTAINER` | `activity-logs` | Per-action audit container |
| `COSMOS_DB_JOB_CONTAINER` | `job-logs` | Per-case job record container |
| `AUTH_ENABLED` | `false` | Enforce Entra ID JWT validation on the API |
| `ENTRA_TENANT_ID` | – | Directory (tenant) ID |
| `ENTRA_CLIENT_ID` | – | Backend API app registration client ID |
| `BACKEND_HOST` / `BACKEND_PORT` / `BACKEND_WORKERS` | `0.0.0.0` / `8000` / `4` | Uvicorn settings |

### Frontend (`frontend/.env`)

| Variable | Default | Description |
|---|---|---|
| `BACKEND_URL` | `http://localhost:8000` | Backend base URL (injected automatically by `deploy-aca.ps1`) |
| `ENV` | – | Free-form environment label shown in the UI |
| `DEBUG_MODE` | `false` | Expose intermediate JSON/validation downloads |
| `AUTH_ENABLED` | `false` | Enable MSAL sign-in |
| `ENTRA_TENANT_ID` | – | Directory (tenant) ID |
| `ENTRA_FRONTEND_CLIENT_ID` | falls back to `ENTRA_CLIENT_ID` | Frontend app registration |
| `ENTRA_BACKEND_CLIENT_ID` | falls back to `ENTRA_CLIENT_ID` | Backend API app registration (audience) |
| `ENTRA_CLIENT_SECRET` | – | Frontend app registration secret |
| `ENTRA_SCOPE` | `api://<backend-client-id>/user_impersonation` | Scope requested for the backend API |
| `ENTRA_REDIRECT_URI` | `FRONTEND_URL` or `http://localhost:8501` | Must match the app registration redirect URI |

> Never commit `.env` files. `.gitignore` already excludes them, and `deploy-aca.ps1` injects
> values at runtime rather than baking them into images.

---

## Running locally

### Prerequisites

- Python 3.11
- Docker (optional, for container parity)
- Azure CLI (`az`) if you plan to deploy
- Provisioned Azure OpenAI and Document Intelligence resources (Blob Storage, Cosmos DB and
  Entra ID are optional)

### 1. Backend

```bash
cd backend
python -m venv .venv
source .venv/bin/activate        # Windows: .venv\Scripts\Activate.ps1
pip install -r requirements.txt
cp .env.example .env             # then fill in your endpoints and keys
uvicorn main:app --reload --port 8000
```

Open http://localhost:8000/docs for the interactive OpenAPI documentation and
http://localhost:8000/api/health for a health check.

### 2. Frontend

```bash
cd frontend
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
# create .env with at least: BACKEND_URL=http://localhost:8000
streamlit run app.py
```

The UI is served at http://localhost:8501.

### 3. Run with Docker

```bash
docker build -t ehcp-backend ./backend
docker run --env-file backend/.env -p 8000:8000 ehcp-backend

docker build -t ehcp-frontend ./frontend
docker run --env-file frontend/.env -e BACKEND_URL=http://host.docker.internal:8000 -p 8501:8501 ehcp-frontend
```

---
## API reference

All routes are prefixed with `/api`. Requests should include an `X-Session-ID` header (and an
`Authorization` header carrying the Entra ID access credential when `AUTH_ENABLED=true`).

| Method | Path | Purpose |
|---|---|---|
| `GET` | `/health` | Liveness probe |
| `GET` | `/doc-types` | Supported document types |
| `GET` | `/mapping-fields` | Output field definitions from the mapping workbook |
| `POST` | `/upload` | Upload advice documents; returns detected document types |
| `POST` | `/analyze` | Run the reader pipeline and return results |
| `POST` | `/analyze-stream` | Run the reader pipeline with SSE progress events |
| `POST` | `/write-ehcp` | Fill the EHCP template and validate the output |
| `GET` | `/results/{filename}` | Download a session-scoped result file |
| `GET` | `/download/{filepath}` | Download a file by relative path |
| `DELETE` | `/files/{filename}` | Remove an uploaded file from the session |
| `POST` | `/log-browse`, `/log-activity` | Client-side audit events |

Interactive documentation is available at `/docs` on the backend.

---

## Customising the solution

- **Different EHCP template** — replace `backend/EHCP_LCC_Template.docx` and update the
  placeholder references in `backend/ehcp_mapping.xlsx`. The mapping workbook holds one sheet per
  section, mapping JSON paths to template fields, so most layout changes need no code edits.
- **New or changed fields** — update the relevant `backend/schemas/*_schema.json` and the matching
  `backend/prompts/*_prompt.txt`, then add the field to the mapping workbook.
- **New document type** — add a prompt, a schema, an entry in `DOC_TYPE_MAP`, content markers in
  `_CONTENT_MARKERS` (both in `backend/app/routers/pipeline.py`), and a mapping sheet.
- **Model choice** — change `AZURE_OPENAI_DEPLOYMENT`; keep `MODEL_TEMPERATURE=0` for reproducible
  extraction.
- **Validation strictness** — tune the re-check rules and critical-field lists in
  `backend/app/services/helpers/validation_helpers.py`.

---

## Test cases

`Test Cases/` contains fully synthetic material for end-to-end verification:

- `Simple Case Inputs/` and `Complex Case Inputs/` — sample advice documents (DOCX and PDF)
- `Blank Templates/` — the empty advice forms and the EHCP output template
- `Simple Case Output/` and `Complex Case Output/` — reference draft EHCP outputs

Use them to validate a new deployment and to benchmark accuracy/completeness after prompt, schema
or model changes. No real personal data is included.

---

## Security, privacy and responsible AI

- **Human in the loop.** Output is a *draft*. A qualified professional must review, edit and
  approve every plan before it is issued.
- **Sensitive data.** Inputs contain special-category personal data about children. Deploy into a
  tenant and region that meet your organisation's data-residency and DPIA requirements, restrict
  network access, and set retention/lifecycle policies on the blob container and Cosmos DB
  containers.
- **Secrets.** No secrets are committed. Keys are passed as Container App secrets, and managed
  identity removes them entirely. Prefer `USE_MANAGED_IDENTITY=true` in production.
- **Authentication.** Enable `AUTH_ENABLED=true` in any non-local environment so the API validates
  Entra ID JWTs (signature, audience, issuer and expiry).
- **Network isolation.** The backend is deployed with internal-only ingress; only the frontend is
  publicly reachable. Consider tightening the backend CORS policy in `backend/main.py` from `*` to
  the frontend origin.
- **Session isolation.** Uploads and outputs are written to per-session directories, and result
  downloads are restricted to the owning session.
- **Auditability.** With `AUDIT_LOG_ENABLED=true`, every user action and a consolidated per-case
  job record (including token usage, accuracy and completeness) are persisted to Cosmos DB.

---

## Troubleshooting

| Symptom | Likely cause / fix |
|---|---|
| `401 Invalid token audience` | Backend `ENTRA_CLIENT_ID` does not match the audience of the token; check `ENTRA_SCOPE` on the frontend |
| `AZURE_STORAGE_ACCOUNT_URL is not set` | Managed identity is on but the blob account URL is missing |
| Empty or partial extraction | Check the Document Intelligence endpoint/key and that the file is a supported DOCX/PDF; inspect the `_doctext.txt` artefact |
| Low completeness scores | The source advice document is genuinely missing critical fields — review `critical_fields_missing` in the validation JSON |
| Rate-limit / 429 errors from Azure OpenAI | Increase the deployment's TPM quota; files are processed in parallel |
| Frontend cannot reach backend | Confirm `BACKEND_URL`; in ACA the backend uses internal ingress and is only reachable from within the environment |
| Sign-in redirect loop | The deployed frontend URL must be registered as a redirect URI and set via `ENTRA_REDIRECT_URI` |
