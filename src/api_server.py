import os
import shutil
import tempfile
from fastapi import FastAPI, UploadFile, File, HTTPException, BackgroundTasks
from fastapi.responses import FileResponse
from typing import List, Optional

# Add project root to sys.path for direct execution of api_server.py
import sys
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from src.core_converter import convert_word_to_pdf, convert_pdf_to_word, merge_pdfs_core

app = FastAPI(
    title="Word/PDF Converter API",
    description="API for converting Word to PDF, PDF to Word, and merging PDFs.",
    version="1.0.0",
)

def _cleanup_temp_dir(temp_dir: str):
    if os.path.exists(temp_dir):
        shutil.rmtree(temp_dir)

@app.get("/health")
async def health_check():
    return {"status": "ok"}

@app.post("/convert/word-to-pdf")
async def convert_word_to_pdf_api(background_tasks: BackgroundTasks, file: UploadFile = File(...)):
    if not file.filename.lower().endswith((".docx", ".doc")):
        raise HTTPException(status_code=400, detail="Invalid file type. Only .docx and .doc are supported.")

    temp_dir = tempfile.mkdtemp()
    background_tasks.add_task(_cleanup_temp_dir, temp_dir)

    input_path = os.path.join(temp_dir, file.filename)
    output_filename = os.path.splitext(file.filename)[0] + ".pdf"
    output_path = os.path.join(temp_dir, output_filename)

    # Save uploaded file to temporary directory
    with open(input_path, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)

    success = convert_word_to_pdf(input_path, output_path)

    if not success or not os.path.exists(output_path) or os.path.getsize(output_path) == 0:
        raise HTTPException(status_code=500, detail="Word to PDF conversion failed.")

    return FileResponse(path=output_path, filename=output_filename, media_type="application/pdf")

@app.post("/convert/pdf-to-word")
async def convert_pdf_to_word_api(background_tasks: BackgroundTasks, file: UploadFile = File(...)):
    if not file.filename.lower().endswith(".pdf"):
        raise HTTPException(status_code=400, detail="Invalid file type. Only .pdf is supported.")

    temp_dir = None
    try:
        temp_dir = tempfile.mkdtemp()
        background_tasks.add_task(_cleanup_temp_dir, temp_dir)
        input_path = os.path.join(temp_dir, file.filename)
        output_filename = os.path.splitext(file.filename)[0] + ".docx"
        output_path = os.path.join(temp_dir, output_filename)

        # Save uploaded file to temporary directory
        with open(input_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

        success = convert_pdf_to_word(input_path, output_path)

        if not success or not os.path.exists(output_path) or os.path.getsize(output_path) == 0:
            raise HTTPException(status_code=500, detail="PDF to Word conversion failed.")

        return FileResponse(path=output_path, filename=output_filename, media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    finally:
        pass # Cleanup is handled by BackgroundTasks

@app.post("/merge/pdfs")
async def merge_pdfs_api(
    background_tasks: BackgroundTasks,
    files: List[UploadFile] = File(...),
    generate_toc: bool = False,
    add_page_numbers: bool = False,
    page_number_format: str = "數字",
    start_page_number: int = 1
):
    if not files:
        raise HTTPException(status_code=400, detail="No files provided for merging.")

    file_list_for_core = []
    temp_input_paths = []
    temp_dir = None

    try:
        temp_dir = tempfile.mkdtemp()
        background_tasks.add_task(_cleanup_temp_dir, temp_dir)

        for i, file in enumerate(files):
            # Validate file type for merging (Word or PDF)
            if not file.filename.lower().endswith((".docx", ".doc", ".pdf")):
                raise HTTPException(status_code=400, detail=f"Invalid file type for merging: {file.filename}. Only .docx, .doc, and .pdf are supported.")
            
            input_path = os.path.join(temp_dir, f"input_{i}_{file.filename}")
            temp_input_paths.append(input_path)
            with open(input_path, "wb") as buffer:
                shutil.copyfileobj(file.file, buffer)
            
            file_list_for_core.append({'path': input_path, 'title': os.path.splitext(file.filename)[0]})

        output_filename = "merged_document.pdf"
        output_path = os.path.join(temp_dir, output_filename)

        options = {
            'generate_toc': generate_toc,
            'add_page_numbers': add_page_numbers,
            'page_number_format': page_number_format,
            'start_page_number': start_page_number
        }

        # Non-interactive password callback for API
        def api_password_callback(filename):
            raise HTTPException(status_code=400, detail=f"Encrypted PDF '{filename}' encountered. API does not support interactive password input.")

        try:
            merge_pdfs_core(
                file_list=file_list_for_core,
                output_file=output_path,
                options=options,
                temp_dir=temp_dir,
                password_callback=api_password_callback
            )
        except HTTPException as e:
            raise e # Re-raise HTTPExceptions from callback
        except Exception as e:
            raise HTTPException(status_code=500, detail=f"PDF merging failed: {e}")

        if not os.path.exists(output_path) or os.path.getsize(output_path) == 0:
            raise HTTPException(status_code=500, detail="PDF merging failed: Output file is empty or not created.")

        return FileResponse(path=output_path, filename=output_filename, media_type="application/pdf")
    finally:
        pass # Cleanup is handled by BackgroundTasks
