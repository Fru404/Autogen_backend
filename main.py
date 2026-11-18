import io
import os
import shutil
import subprocess
import tempfile
import zipfile
import threading
import time
from pathlib import Path

from fastapi import (
    FastAPI,
    File,
    Form,
    HTTPException,
    UploadFile
)

from fastapi.middleware.cors import CORSMiddleware

from fastapi.responses import StreamingResponse

from docx import Document

from document_processor import (
    get_placeholders,
    read_excel,
    generate_document
)


# ============================================================
# CONFIGURATION
# ============================================================

# Master switch for PDF generation.
#
# True:
#     PDF generation can be requested from the UI.
#
# False:
#     PDF generation is disabled completely.
#
GENERATE_PDF = True


# ============================================================
# PDF CONVERSION
# ============================================================

def find_libreoffice():
    """
    Find the LibreOffice executable.

    The environment variable LIBREOFFICE_PATH can be used to
    explicitly specify the executable path.

    Examples:
        Windows:
            C:\\Program Files\\LibreOffice\\program\\soffice.exe

        Linux:
            /usr/bin/soffice
    """

    configured_path = os.getenv("LIBREOFFICE_PATH")

    if configured_path:
        configured = Path(configured_path)

        if configured.exists():
            return str(configured)

        # It may be an executable available on PATH.
        found = shutil.which(configured_path)

        if found:
            return found

    # Standard PATH lookup.
    found = shutil.which("soffice")

    if found:
        return found

    found = shutil.which("libreoffice")

    if found:
        return found

    # Common Windows installation locations.
    windows_paths = [
        Path(os.environ.get("PROGRAMFILES", ""))
        / "LibreOffice"
        / "program"
        / "soffice.exe",

        Path(os.environ.get("PROGRAMFILES(X86)", ""))
        / "LibreOffice"
        / "program"
        / "soffice.exe",
    ]

    for path in windows_paths:
        if path.exists():
            return str(path)

    return None


def convert_docx_to_pdf(
    docx_path,
    output_folder
):
    """
    Convert one DOCX file to PDF using LibreOffice.

    The PDF is written into output_folder and the resulting
    PDF path is returned.
    """

    docx_path = Path(docx_path)
    output_folder = Path(output_folder)

    output_folder.mkdir(
        parents=True,
        exist_ok=True
    )

    soffice = find_libreoffice()

    if not soffice:
        raise RuntimeError(
            "PDF generation was requested, but LibreOffice "
            "could not be found. Install LibreOffice and make "
            "sure 'soffice' is available, or set the "
            "LIBREOFFICE_PATH environment variable."
        )

    # Give each conversion its own LibreOffice profile.
    # This prevents profile locking when multiple requests
    # are processed by the API.
    profile_dir = Path(
        tempfile.mkdtemp(
            prefix="autogen-libreoffice-profile-"
        )
    )

    try:

        command = [
            soffice,
            "--headless",
            "--convert-to",
            "pdf",
            "--outdir",
            str(output_folder),
            f"-env:UserInstallation={profile_dir.as_uri()}",
            str(docx_path),
        ]

        result = subprocess.run(
            command,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            timeout=120,
        )

        if result.returncode != 0:
            details = (
                result.stderr.strip()
                or result.stdout.strip()
                or "Unknown LibreOffice error."
            )

            raise RuntimeError(
                f"LibreOffice PDF conversion failed for "
                f"{docx_path.name}: {details}"
            )

        pdf_path = (
            output_folder /
            f"{docx_path.stem}.pdf"
        )

        if not pdf_path.exists():
            raise RuntimeError(
                f"LibreOffice completed without creating "
                f"the expected PDF for {docx_path.name}."
            )

        return pdf_path

    except subprocess.TimeoutExpired:
        raise RuntimeError(
            f"PDF conversion timed out for {docx_path.name}."
        )

    finally:

        shutil.rmtree(
            profile_dir,
            ignore_errors=True
        )


# ============================================================
# FASTAPI
# ============================================================

app = FastAPI(
    title="Document Generator API",
    version="1.1.0"
)


# ============================================================
# GENERATION PROGRESS STORE
# ============================================================

# Progress is kept in memory for the lifetime of this API process.
# The frontend polls this state while the /generate request returns
# the ZIP directly when processing is complete.
_generation_progress = {}
_generation_progress_lock = threading.Lock()
_PROGRESS_TTL_SECONDS = 3600


def set_generation_progress(
    request_id,
    *,
    status="processing",
    percent=0,
    message="Preparing generation...",
    completed=0,
    total=0,
):
    now = time.time()
    with _generation_progress_lock:
        # Discard old entries periodically so this dictionary stays bounded.
        expired = [
            key for key, value in _generation_progress.items()
            if now - value.get("updated_at", now) > _PROGRESS_TTL_SECONDS
        ]
        for key in expired:
            _generation_progress.pop(key, None)

        _generation_progress[request_id] = {
            "request_id": request_id,
            "status": status,
            "percent": max(0, min(100, int(percent))),
            "message": message,
            "completed": int(completed),
            "total": int(total),
            "updated_at": now,
        }


def get_generation_progress(request_id):
    with _generation_progress_lock:
        progress = _generation_progress.get(request_id)
        if progress is None:
            return None
        # Do not expose the internal timestamp.
        return {
            key: value for key, value in progress.items()
            if key != "updated_at"
        }


# ============================================================
# CORS
# ============================================================

app.add_middleware(
    CORSMiddleware,

    allow_origins=[
        "*",
    ],

    allow_credentials=True,

    allow_methods=[
        "*"
    ],

    allow_headers=[
        "*"
    ],
)


# ============================================================
# HEALTH CHECK
# ============================================================

@app.get("/health")
def health():

    return {
        "status": "ok",
        "service": "document-generator",
        "pdf_enabled": GENERATE_PDF
    }


# ============================================================
# INSPECT FILES
# ============================================================

@app.post("/inspect")
async def inspect_files(
    template: UploadFile = File(...),
    excel: UploadFile = File(...)
):

    # --------------------------------------------------------
    # Validate extensions
    # --------------------------------------------------------

    template_name = (
        template.filename or ""
    ).lower()

    excel_name = (
        excel.filename or ""
    ).lower()

    if not template_name.endswith(
        ".docx"
    ):

        raise HTTPException(
            status_code=400,
            detail="Template must be a .docx file."
        )

    if not excel_name.endswith(
        (".xlsx", ".xls")
    ):

        raise HTTPException(
            status_code=400,
            detail="Excel file must be .xlsx or .xls."
        )

    temp_dir = tempfile.mkdtemp()

    try:

        template_path = (
            Path(temp_dir)
            / "template.docx"
        )

        excel_path = (
            Path(temp_dir)
            / "clients.xlsx"
        )

        # Save uploaded files
        with open(
            template_path,
            "wb"
        ) as file:

            shutil.copyfileobj(
                template.file,
                file
            )

        with open(
            excel_path,
            "wb"
        ) as file:

            shutil.copyfileobj(
                excel.file,
                file
            )

        # ----------------------------------------------------
        # Read Word template
        # ----------------------------------------------------

        document = Document(
            str(template_path)
        )

        placeholders = get_placeholders(
            document
        )

        # ----------------------------------------------------
        # Read Excel
        # ----------------------------------------------------

        columns, records = read_excel(
            excel_path
        )

        # ----------------------------------------------------
        # Determine missing fields
        # ----------------------------------------------------

        missing_fields = [
            field
            for field in placeholders
            if field not in columns
        ]

        # ----------------------------------------------------
        # Extra Excel fields
        # ----------------------------------------------------

        extra_columns = [
            column
            for column in columns
            if column not in placeholders
        ]

        return {
            "template": {
                "filename": template.filename,
                "fields": placeholders
            },

            "excel": {
                "filename": excel.filename,
                "columns": columns,
                "row_count": len(records),
                "rows": records
            },

            "validation": {
                "missing_fields": missing_fields,
                "extra_columns": extra_columns,
                "ready": len(missing_fields) == 0
            },

            "settings": {
                "pdf_enabled": GENERATE_PDF
            }
        }

    except Exception as error:

        raise HTTPException(
            status_code=400,
            detail=str(error)
        )

    finally:

        shutil.rmtree(
            temp_dir,
            ignore_errors=True
        )


# ============================================================
# GENERATION PROGRESS
# ============================================================

@app.get("/generate/progress/{request_id}")
def read_generation_progress(request_id: str):
    progress = get_generation_progress(request_id)
    if progress is None:
        raise HTTPException(
            status_code=404,
            detail="Generation progress was not found."
        )
    return progress


# ============================================================
# GENERATE DOCUMENTS
# ============================================================

@app.post("/generate")
def generate_documents(
    template: UploadFile = File(...),
    data_json: str | None = Form(None),
    rows: str | None = Form(None),
    generate_pdf: bool = Form(False),
    request_id: str = Form(...)
):

    # --------------------------------------------------------
    # Validate template
    # --------------------------------------------------------

    template_name = (
        template.filename or ""
    ).lower()

    if not template_name.endswith(
        ".docx"
    ):

        raise HTTPException(
            status_code=400,
            detail="Template must be a .docx file."
        )

    # --------------------------------------------------------
    # Parse JSON
    # --------------------------------------------------------

    import json

    # The original API uses "data_json".
    # The current Next.js UI may send "rows".
    # Accept both so the frontend and backend remain compatible.
    payload = data_json if data_json is not None else rows

    if payload is None:
        raise HTTPException(
            status_code=400,
            detail="No client data was supplied."
        )

    try:

        records = json.loads(payload)

    except json.JSONDecodeError:

        raise HTTPException(
            status_code=400,
            detail="Invalid client data JSON."
        )

    if not isinstance(
        records,
        list
    ):

        raise HTTPException(
            status_code=400,
            detail="data_json must contain a list of records."
        )

    # --------------------------------------------------------
    # PDF permission
    # --------------------------------------------------------

    if generate_pdf and not GENERATE_PDF:

        raise HTTPException(
            status_code=400,
            detail=(
                "PDF generation is disabled "
                "by the server configuration."
            )
        )

    if not request_id or len(request_id) > 128:
        raise HTTPException(
            status_code=400,
            detail="A valid request_id is required for progress tracking."
        )

    set_generation_progress(
        request_id,
        status="processing",
        percent=0,
        message="Preparing your receipts...",
        completed=0,
        total=len(records),
    )

    # --------------------------------------------------------
    # Temporary workspace
    # --------------------------------------------------------

    temp_dir = tempfile.mkdtemp()

    try:

        template_path = (
            Path(temp_dir)
            / "template.docx"
        )

        output_dir = (
            Path(temp_dir)
            / "receipts"
        )

        output_dir.mkdir(
            parents=True,
            exist_ok=True
        )

        # ----------------------------------------------------
        # Save template
        # ----------------------------------------------------

        with open(
            template_path,
            "wb"
        ) as file:

            shutil.copyfileobj(
                template.file,
                file
            )

        # ----------------------------------------------------
        # Generate documents
        # ----------------------------------------------------

        generated_files = []
        total_records = len(records)
        # Count each concrete unit of work: one DOCX generation,
        # one PDF conversion when enabled, and one ZIP file addition.
        total_work_units = total_records * (4 if generate_pdf else 2)
        completed_work_units = 0

        # ----------------------------------------------------
        # Generate DOCX receipts
        # ----------------------------------------------------

        for index, record in enumerate(records, start=1):
            if not isinstance(record, dict):
                raise HTTPException(
                    status_code=400,
                    detail=f"Row {index} is invalid."
                )

            output_file = generate_document(
                template_file=str(template_path),
                output_folder=str(output_dir),
                data=record,
                row_number=index
            )

            generated_files.append(output_file)
            completed_work_units += 1
            percent = round(100 * completed_work_units / total_work_units)
            set_generation_progress(
                request_id,
                percent=percent,
                message=f"Generating receipt {index} of {total_records}...",
                completed=completed_work_units,
                total=total_work_units,
            )

        # ----------------------------------------------------
        # PDF conversion
        # ----------------------------------------------------

        if generate_pdf:
            docx_files = list(generated_files)
            for index, docx_file in enumerate(docx_files, start=1):
                try:
                    pdf_file = convert_docx_to_pdf(
                        docx_path=docx_file,
                        output_folder=output_dir
                    )
                except Exception as error:
                    raise HTTPException(
                        status_code=500,
                        detail=(
                            f"PDF generation failed for "
                            f"{Path(docx_file).name}: {error}"
                        )
                    )

                generated_files.append(pdf_file)
                completed_work_units += 1
                percent = round(100 * completed_work_units / total_work_units)
                set_generation_progress(
                    request_id,
                    percent=percent,
                    message=f"Converting PDF {index} of {len(docx_files)}...",
                    completed=completed_work_units,
                    total=total_work_units,
                )

        # ----------------------------------------------------
        # Create ZIP and report each file added
        # ----------------------------------------------------

        zip_buffer = io.BytesIO()
        total_files_to_zip = len(generated_files)

        with zipfile.ZipFile(
            zip_buffer,
            mode="w",
            compression=zipfile.ZIP_DEFLATED
        ) as zip_file:
            for index, file_path in enumerate(generated_files, start=1):
                zip_file.write(file_path, arcname=file_path.name)
                completed_work_units += 1
                percent = round(100 * completed_work_units / total_work_units)
                set_generation_progress(
                    request_id,
                    percent=percent,
                    message=f"Adding file {index} of {total_files_to_zip} to ZIP...",
                    completed=completed_work_units,
                    total=total_work_units,
                )

        zip_buffer.seek(0)
        set_generation_progress(
            request_id,
            status="completed",
            percent=100,
            message="ZIP archive ready.",
            completed=total_work_units,
            total=total_work_units,
        )

        # ----------------------------------------------------
        # Return ZIP
        # ----------------------------------------------------

        return StreamingResponse(
            zip_buffer,
            media_type="application/zip",
            headers={
                "Content-Disposition":
                    'attachment; filename="Autogen_Receipts.zip"'
            }
        )

    except HTTPException as error:
        set_generation_progress(
            request_id,
            status="failed",
            percent=0,
            message=str(error.detail),
            completed=0,
            total=len(records),
        )
        raise

    except Exception as error:
        set_generation_progress(
            request_id,
            status="failed",
            percent=0,
            message=str(error),
            completed=0,
            total=len(records),
        )
        raise HTTPException(
            status_code=500,
            detail=str(error)
        )

    finally:

        # Cleanup after response processing.
        #
        # For a production deployment we may change this
        # slightly depending on the hosting platform.

        shutil.rmtree(
            temp_dir,
            ignore_errors=True
        )