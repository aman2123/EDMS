# EDMS Project — Handover Document

**Prepared by:** Ruchika Jha
**Date:** September 2026
**Project:** Enterprise Document Management System (EDMS) — Document Split, Categorize, Extract & SharePoint Upload Platform

---

## 1. Overview

This platform ingests scanned multi-page PDF batches (Delivery Notes, Journal Vouchers, Invoices, Supporting Documents), automatically classifies each page, extracts structured fields (using OCR + Azure AI), and uploads categorized documents to SharePoint with auto-populated metadata columns.

**Tech stack:**
- Backend: Python, FastAPI, SQLAlchemy, PostgreSQL
- OCR/AI: Tesseract (local), Azure Document Intelligence (cloud — `prebuilt-invoice` and `prebuilt-read` models)
- Cloud integration: Microsoft Graph API (SharePoint uploads), Azure AD (OAuth 2.0 client-credentials auth)
- Frontend: HTML/JS single-page "Document Split & Categorize" interface, served by the FastAPI backend

---

## 2. System Architecture

```
Upload PDF → Split into pages → Classify each page (JV / Invoice / DN / Supporting Document)
    → Group into blocks → Extract fields per block (Tesseract or Azure, depending on category)
    → Save extraction records to PostgreSQL → Upload selected blocks to SharePoint (DN library)
    → Auto-populate BoxCode / FileCode / DocNumber metadata columns on each uploaded file
```

Key backend files:
- `app/api/routes/documents.py` — main API routes: split, reassign, extract, download, SharePoint upload
- `app/services/classification/document_classifier.py` — page classification logic (JV/Invoice/DN/Supporting)
- `app/services/extraction/` — per-category field extractors (`invoice_extractor.py`, `jv_extractor.py`, `dn_extractor.py`)
- `app/services/ocr/azure_ocr_service.py` — wrapper around Azure Document Intelligence
- `app/services/integrations/sharepoint_service.py` — Microsoft Graph SharePoint upload + metadata setting
- `app/database/connection.py` — SQLAlchemy engine/session setup, reads `DATABASE_URL` from `.env`
- `app/utils/filename_parser.py` — parses BoxCode/FileCode from the original uploaded filename

---

## 3. Environment Configuration (`.env`)

Location: project root (`EDMS/.env`)

```
AZURE_DOCINTEL_ENDPOINT=<Azure Document Intelligence resource endpoint>
AZURE_DOCINTEL_KEY=<Azure Document Intelligence API key>
DATABASE_URL=postgresql://<user>:<url-encoded-password>@localhost:5432/idp_platform
SHAREPOINT_TENANT_ID=<Azure AD Directory (tenant) ID>
SHAREPOINT_CLIENT_ID=<Azure AD Application (client) ID>
SHAREPOINT_CLIENT_SECRET=<Azure AD client secret value>
SHAREPOINT_HOSTNAME=alumetalllc.sharepoint.com
SHAREPOINT_SITE_PATH=/sites/FinanceDept
SHAREPOINT_LIBRARY_NAME=DN
```

**Important gotcha:** if the database password contains special characters (e.g. `@`), it must be URL-encoded in `DATABASE_URL` (e.g. `@` → `%40`). A double-`@` password needs double-encoding (`%40%40`). This caused a multi-hour debugging session — always verify with:
```powershell
python -c "from sqlalchemy.engine import make_url; u = make_url('<paste DATABASE_URL here>'); print(u.username, u.password, u.host)"
```

---

## 4. Azure AD / SharePoint Authentication Setup

This was the single largest point of friction during the project and is the most important section to get right for whoever inherits this.

### 4.1 App Registration
- App name: **EDMS SharePoint Integration**
- Registered in Azure AD (Entra ID) under the company's Microsoft 365 tenant
- Uses **OAuth 2.0 client-credentials flow** (app-only auth, no user sign-in) via the `msal` Python library

### 4.2 The critical gotcha: permission must be under Microsoft Graph, NOT the legacy SharePoint API

When adding API permissions in Azure AD (App registrations → API permissions → Add a permission), there are **two different resources** that both offer a permission called `Sites.ReadWrite.All`:
- ❌ **"SharePoint"** (legacy API) — looks identical, admin-consents fine, shows a green checkmark, but **tokens requested for Microsoft Graph will not carry this permission**. This caused silent 401 errors (`spException`) for weeks despite the permission showing as "Granted" in the portal.
- ✅ **"Microsoft Graph"** — this is the one that must be used, since all upload code calls `https://graph.microsoft.com/v1.0/...`

**How to verify this is set correctly:** on the API permissions page, the `Sites.ReadWrite.All` entry must appear nested under **"Microsoft Graph (n)"** in the permissions table — not under a separate "SharePoint (n)" group.

**How to verify a token actually has the permission** (useful diagnostic if uploads start failing again):
```python
import jwt as pyjwt
decoded = pyjwt.decode(access_token, options={"verify_signature": False})
print(decoded.get("roles"))  # should include 'Sites.ReadWrite.All'
```
If `roles` is `None` or empty, the permission is attached to the wrong resource — go back to API permissions and fix it there, not in code.

### 4.3 Admin consent
After adding/changing the permission, someone with tenant admin rights must click **"Grant admin consent for [tenant]"** on the API permissions page. This is a separate step from adding the permission — both are required.

### 4.4 Client secret rotation
Client secrets expire (a date is set at creation). To rotate: Azure AD → App registrations → EDMS SharePoint Integration → Certificates & secrets → New client secret → copy the value immediately (shown once) → update `SHAREPOINT_CLIENT_SECRET` in `.env`.

---

## 5. SharePoint Upload Logic (`sharepoint_service.py`)

- Targets the **DN document library specifically** — this is a separate Graph "drive" from the site's default "Documents" library. The code resolves the site to a `site_id`, then resolves the DN library by name to its own `drive_id` via `_get_drive_id()`. Uploading through the wrong drive ID silently lands files in the wrong library (this happened early on — files went to `Documents/DN/` instead of the actual `DN` library).
- Files are uploaded via `PUT` to `/drives/{drive_id}/root:/{path}/{filename}:/content`. Graph auto-creates any folder in the path that doesn't yet exist (no explicit "create folder" call needed).
- **Metadata (BoxCode, FileCode, DocNumber)** is set in a second call after upload, via `PATCH` to `/drives/{drive_id}/items/{item_id}/listItem/fields`. The internal SharePoint column names (confirmed via List Settings → click column → check `Field=` in the URL) are:
  - `BoxCode`
  - `FileCode`
  - `DocNumber`
- **BoxCode / FileCode** are parsed from the original batch filename (format: `{BoxCode}_{FileCode}_{...}.pdf`) via `app/utils/filename_parser.py` — constant for the whole batch.
- **DocNumber** is OCR'd per-block from the first page of each DN block, since it's a handwritten number that varies per document within a batch — see Section 6.

**Known open item flagged for change:** currently each upload batch creates a new folder named after a random 8-character UUID (`batch_id`) inside the DN library (e.g. `DN/7f4645c5/filename.pdf`). This was flagged by [manager] as needing to change to place files **directly under DN** with no per-batch subfolder. Not yet implemented — whoever picks this up should update the `upload_url` path construction in `upload_files_to_sharepoint()` to drop the `batch_id` folder segment, and decide on a filename-collision strategy (two batches could produce identically-named block files).

---

## 6. Handwritten DocNumber Extraction

- `extract_doc_number_from_image()` in `dn_extractor.py` reads the handwritten DocNumber (7-8 digit number, e.g. `16001917`) off a DN page image.
- **First attempt:** Azure's `prebuilt-invoice` model (via `analyze_invoice()`). This model is tuned for invoice layouts and **reliably returns empty content for DN/Receipt Voucher pages** — confirmed via debug logging.
- **Fallback:** Azure's `prebuilt-read` model — general-purpose OCR/handwriting recognition, no layout assumptions. This is the one that actually succeeds for DN pages.
- A regex (`\b\d{7,8}\b`) then extracts the DocNumber-shaped token from the returned text.

**Known inefficiency flagged for optimization:** since `prebuilt-invoice` is known to always fail on DN pages, calling it first wastes an Azure-billed page-analysis on every single DocNumber extraction (Azure bills per page analyzed, not per token). For DN-category documents specifically, the code should skip `prebuilt-invoice` and call `prebuilt-read` directly — cutting Azure cost roughly in half for this operation at scale (relevant given the ~200,000-file volume discussed for future scaling).

**Known accuracy limitation:** there is currently no confidence-score check on the OCR'd DocNumber — whatever Azure returns as its best guess is used directly. Azure's Read API does return per-word confidence scores that are not currently being surfaced. A reasonable next step would be flagging low-confidence DocNumbers for manual review, mirroring the existing `needs_review` pattern already used for low-confidence page classification.

---

## 7. Database

- PostgreSQL, accessed via SQLAlchemy (`app/database/connection.py`)
- Table: `document_records` (via `DocumentRecord` model) — stores extraction results (batch_id, page_number, category, invoice_no, vendor, amount, invoice_date, status)
- Initialize/recreate schema: `python -m app.database.init_db`
- Direct DB access: `psql` (path may not be on PATH by default — full path was `C:\Program Files\PostgreSQL\18\bin\psql.exe`, requires the PowerShell call operator `&` when using the full quoted path)

---

## 8. Known Issues / Flagged for Future Work

1. **SharePoint folder structure** — change from per-batch subfolders to flat placement directly under the DN library (see Section 5).
2. **Azure cost optimization** — skip the redundant `prebuilt-invoice` call for DN-category documents; call `prebuilt-read` directly (see Section 6).
3. **No confidence-based review flag on OCR'd DocNumber** — worth adding, given handwriting OCR is inherently imperfect.
4. **Passport extractor** — not yet implemented; `dispatch_extraction()` currently returns raw OCR text as a placeholder for the "Passport" category rather than structured fields.
5. **Login/auth system** — a login page with encrypted (hashed) password storage and a super-admin user-creation flow was requested but not yet implemented as of this handover. If in progress, check for a partial implementation before starting from scratch.
6. **`dispatch_extraction()` pattern** — this function in `documents.py` routes extraction by category; if a new document category is added in the future, make sure a corresponding branch is added here (a previous bug involved this function being referenced before it was defined at all — worth being cautious about incomplete refactors here).

---

## 9. Quick Reference — Common Commands

```powershell
# Activate virtual environment
& C:\Users\lenovo\Desktop\EDMS\.venv\Scripts\Activate.ps1

# Run the server
uvicorn app.main:app --reload

# Recreate database tables
python -m app.database.init_db

# Test SharePoint/Graph token + permissions directly
python -m app.services.integrations.test_token
```

---

## 10. Contacts

- Manager / handover coordinator: [fill in name]
- SharePoint/Azure AD admin: Vijay Lodha (has tenant admin rights, needed for any future permission changes or client secret rotation)
