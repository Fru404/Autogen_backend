import io
import os
import shutil
import tempfile
import zipfile
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
GENERATE_PDF = False


# ============================================================
# FASTAPI
# ============================================================

app = FastAPI(
    title="Document Generator API",
    version="1.0.0"
)


# ============================================================
# CORS
# ============================================================

app.add_middleware(
    CORSMiddleware,

    allow_origins=[
        "http://localhost:3000",
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
# GENERATE DOCUMENTS
# ============================================================

@app.post("/generate")
async def generate_documents(
    template: UploadFile = File(...),
    data_json: str = Form(...),
    generate_pdf: bool = Form(False)
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

    try:

        records = json.loads(
            data_json
        )

    except json.JSONDecodeError:

        raise HTTPException(
            status_code=400,
            detail="Invalid data_json."
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

        for index, record in enumerate(
            records,
            start=1
        ):

            if not isinstance(
                record,
                dict
            ):

                raise HTTPException(
                    status_code=400,
                    detail=(
                        f"Row {index} is invalid."
                    )
                )

            output_file = generate_document(
                template_file=str(
                    template_path
                ),
                output_folder=str(
                    output_dir
                ),
                data=record,
                row_number=index
            )

            generated_files.append(
                output_file
            )

        # ----------------------------------------------------
        # PDF placeholder
        # ----------------------------------------------------
        #
        # We intentionally do not convert PDF yet.
        #
        # Once PDF conversion is added, this section will
        # create the PDF next to the DOCX file.
        #

        if generate_pdf:

            # PDF conversion will be implemented here.
            #
            # For now:
            #
            # raise an error instead of silently pretending
            # that PDFs were generated.

            raise HTTPException(
                status_code=501,
                detail=(
                    "PDF generation is enabled but "
                    "the PDF conversion engine has not "
                    "yet been configured."
                )
            )

        # ----------------------------------------------------
        # Create ZIP
        # ----------------------------------------------------

        zip_buffer = io.BytesIO()

        with zipfile.ZipFile(
            zip_buffer,
            mode="w",
            compression=zipfile.ZIP_DEFLATED
        ) as zip_file:

            for file_path in generated_files:

                zip_file.write(
                    file_path,
                    arcname=file_path.name
                )

        zip_buffer.seek(0)

        # ----------------------------------------------------
        # Return ZIP
        # ----------------------------------------------------

        return StreamingResponse(
            zip_buffer,
            media_type="application/zip",
            headers={
                "Content-Disposition":
                    'attachment; filename="receipts.zip"'
            }
        )

    except HTTPException:
        raise

    except Exception as error:

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