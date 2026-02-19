import os
import shutil
import uuid
from typing import List, Optional
from fastapi import FastAPI, UploadFile, File, Form, Request
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.templating import Jinja2Templates
from fastapi.staticfiles import StaticFiles

import matcher
from excel_io import get_sheet_names, read_header_file

app = FastAPI(title="ExcelMatcher Web")

# Setup templates and static (if needed)
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
templates = Jinja2Templates(directory=os.path.join(BASE_DIR, "templates"))

# Use /tmp for writable directories on Vercel
IS_VERCEL = "VERCEL" in os.environ
UPLOAD_DIR = "/tmp/uploads" if IS_VERCEL else "uploads"
OUTPUT_DIR = "/tmp/outputs" if IS_VERCEL else "outputs"

os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

@app.get("/", response_class=HTMLResponse)
async def read_root(request: Request, pw: Optional[str] = None):
    # Simple password protection for developer access
    # In a real app, use environment variables. For now, we'll use a simple query toggle or hardcoded check.
    # The user asked for "developer me only use"
    ACCESS_PASSWORD = "admin" # USER can change this later
    
    if pw != ACCESS_PASSWORD:
        return HTMLResponse("<html><body><form method='get'>Password: <input type='password' name='pw'><input type='submit'></form></body></html>")
        
    return templates.TemplateResponse("index.html", {"request": request, "pw": pw})

@app.post("/inspect")
async def inspect_file(file: UploadFile = File(...), sheet_index: int = Form(0), pw: Optional[str] = Form(None)):
    if pw != "admin": return {"error": "Invalid password"}
    temp_id = str(uuid.uuid4())
    temp_path = os.path.join(UPLOAD_DIR, f"temp_{temp_id}_{file.filename}")
    
    try:
        with open(temp_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)
            
        sheets = get_sheet_names(temp_path)
        if not sheets:
            return {"sheets": [], "columns": [], "error": "No sheets found"}
            
        # Validate sheet index
        target_sheet = sheets[sheet_index] if 0 <= sheet_index < len(sheets) else sheets[0]
        
        # Read headers (assume row 1)
        columns = read_header_file(temp_path, target_sheet, 1)
        
        return {
            "sheets": sheets,
            "columns": columns,
            "filename": file.filename
        }
    except Exception as e:
        return {"error": str(e)}
    finally:
        # Clean up temp file? 
        # For a prototype, maybe keep it or delete. 
        # If we delete, user has to re-upload for actual processing.
        # Let's delete to keep clean, browser will re-upload on final submit.
        if os.path.exists(temp_path):
            os.remove(temp_path)

@app.post("/upload")
async def process_match(
    base_file: UploadFile = File(...),
    target_file: UploadFile = File(...),
    base_sheet: int = Form(0),
    target_sheet: int = Form(0),
    key_cols: str = Form(...),
    take_cols: str = Form(...),
    pw: Optional[str] = Form(None),
):
    if pw != "admin": return {"error": "Invalid password"}
    # 1. Save Uploaded Files
    session_id = str(uuid.uuid4())
    upload_session_dir = os.path.join(UPLOAD_DIR, session_id)
    os.makedirs(upload_session_dir, exist_ok=True)
    
    base_path = os.path.join(upload_session_dir, base_file.filename)
    target_path = os.path.join(upload_session_dir, target_file.filename)
    
    with open(base_path, "wb") as buffer:
        shutil.copyfileobj(base_file.file, buffer)
    with open(target_path, "wb") as buffer:
        shutil.copyfileobj(target_file.file, buffer)
        
    # 2. Prepare Config
    base_config = {
        "type": "file",
        "path": os.path.abspath(base_path),
        "sheet": base_sheet, # Use selected sheet index
        "header": 1
    }
    target_config = {
        "type": "file",
        "path": os.path.abspath(target_path),
        "sheet": target_sheet, # Use selected sheet index
        "header": 1
    }
    
    # Parse columns (comma separated)
    keys = [k.strip() for k in key_cols.split(",") if k.strip()]
    takes = [t.strip() for t in take_cols.split(",") if t.strip()]
    
    # Output Directory
    out_dir = os.path.join(OUTPUT_DIR, session_id)
    
    # 3. Run Matcher
    def progress_callback(msg, val):
        print(f"[Progress] {msg} {val}%")
        
    options = {
        "fuzzy": False,
        "match_only": False
    }
    
    try:
        out_path, summary, _ = matcher.match_universal(
            base_config, target_config, keys, takes, out_dir, options, progress=progress_callback
        )
        
        # 4. Return Result
        filename = os.path.basename(out_path)
        return FileResponse(out_path, filename=filename, media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        
    except Exception as e:
        import traceback
        traceback.print_exc()
        return {"error": str(e)}

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
