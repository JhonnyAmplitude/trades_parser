from fastapi import FastAPI, File, UploadFile, HTTPException
from fastapi.responses import JSONResponse
import uvicorn
import shutil
import os
import tempfile

from full_statement import parse_full_statement

app = FastAPI(
    title="Financial Statement Parser API",
    description="API for parsing financial statements including cash operations, forex trades, and stock/bond trades.",
    version="1.0.0"
)

@app.post("/parse-statement")
async def parse_statement(file: UploadFile = File(...)):
    """
    Upload an Excel file (.xls or .xlsx) of a brokerage statement.
    Returns a JSON with account metadata and a unified list of operations.
    """
    # Validate file extension
    filename = file.filename
    if not filename.lower().endswith(('.xls', '.xlsx')):
        raise HTTPException(status_code=400, detail="Unsupported file type. Please upload .xls or .xlsx")

    # Save uploaded file to a temporary location
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(filename)[1]) as tmp:
            shutil.copyfileobj(file.file, tmp)
            tmp_path = tmp.name
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Failed to save uploaded file: {e}")
    finally:
        file.file.close()

    # Parse the statement
    try:
        result = parse_full_statement(tmp_path)
    except Exception as e:
        # Clean up temp file
        os.remove(tmp_path)
        raise HTTPException(status_code=500, detail=f"Error parsing statement: {e}")

    # Clean up temp file
    os.remove(tmp_path)

    return JSONResponse(content=result)


