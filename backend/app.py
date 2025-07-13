from fastapi import FastAPI, File, UploadFile, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse
import os
import tempfile
import shutil
from typing import List, Optional
import json
from datetime import datetime

from ppt_parser import PPTParser
from excel_parser import ExcelParser
from matcher import NumberMatcher
from report_generator import ReportGenerator

app = FastAPI(title="Deck Auditor API", version="1.0.0")

# Configure CORS
app.add_middleware(
    CORSMiddleware,
    allow_origins=["http://localhost:3000"],  # React dev server
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Global storage for current session
current_session = {
    "ppt_data": None,
    "excel_data": None,
    "audit_results": None,
    "session_id": None
}

@app.post("/upload-files")
async def upload_files(
    ppt_file: UploadFile = File(...),
    excel_files: List[UploadFile] = File(...)
):
    """Upload PPT and Excel files for auditing"""
    try:
        session_id = f"session_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
        current_session["session_id"] = session_id
        
        # Create temporary directory
        temp_dir = tempfile.mkdtemp()
        
        # Save PPT file
        ppt_path = os.path.join(temp_dir, ppt_file.filename)
        with open(ppt_path, "wb") as buffer:
            shutil.copyfileobj(ppt_file.file, buffer)
        
        # Save Excel files
        excel_paths = []
        for excel_file in excel_files:
            excel_path = os.path.join(temp_dir, excel_file.filename)
            with open(excel_path, "wb") as buffer:
                shutil.copyfileobj(excel_file.file, buffer)
            excel_paths.append(excel_path)
        
        # Parse files
        ppt_parser = PPTParser()
        excel_parser = ExcelParser()
        
        print(f"Parsing PPT: {ppt_path}")
        ppt_data = ppt_parser.parse_presentation(ppt_path)
        current_session["ppt_data"] = ppt_data
        
        print(f"Parsing Excel files: {excel_paths}")
        excel_data = {}
        for excel_path in excel_paths:
            filename = os.path.basename(excel_path)
            excel_data[filename] = excel_parser.parse_workbook(excel_path)
        current_session["excel_data"] = excel_data
        
        # Clean up temp files
        shutil.rmtree(temp_dir)
        
        return {
            "status": "success",
            "session_id": session_id,
            "ppt_slides": len(ppt_data),
            "excel_files": list(excel_data.keys()),
            "message": "Files uploaded and parsed successfully"
        }
        
    except Exception as e:
        print(f"Error in upload_files: {str(e)}")
        raise HTTPException(status_code=500, detail=f"Error processing files: {str(e)}")

@app.post("/audit")
async def run_audit():
    """Run the audit process on uploaded files"""
    try:
        if not current_session["ppt_data"] or not current_session["excel_data"]:
            raise HTTPException(status_code=400, detail="No files uploaded")
        
        matcher = NumberMatcher()
        audit_results = matcher.match_numbers(
            current_session["ppt_data"], 
            current_session["excel_data"]
        )
        
        current_session["audit_results"] = audit_results
        
        return {
            "status": "success",
            "audit_results": audit_results,
            "total_numbers_found": len(audit_results),
            "matches": len([r for r in audit_results if r["status"] == "Match"]),
            "mismatches": len([r for r in audit_results if r["status"] == "Mismatch"]),
            "untraceable": len([r for r in audit_results if r["status"] == "Untraceable"])
        }
        
    except Exception as e:
        print(f"Error in run_audit: {str(e)}")
        raise HTTPException(status_code=500, detail=f"Error running audit: {str(e)}")

@app.get("/audit-results")
async def get_audit_results():
    """Get current audit results"""
    if not current_session["audit_results"]:
        raise HTTPException(status_code=400, detail="No audit results available")
    
    return {
        "status": "success",
        "results": current_session["audit_results"]
    }

@app.get("/download-report/{format}")
async def download_report(format: str):
    """Download audit report in specified format (pdf/csv)"""
    try:
        if not current_session["audit_results"]:
            raise HTTPException(status_code=400, detail="No audit results available")
        
        report_gen = ReportGenerator()
        
        if format.lower() == "pdf":
            file_path = report_gen.generate_pdf_report(current_session["audit_results"])
            return FileResponse(
                file_path, 
                media_type="application/pdf",
                filename=f"audit_report_{current_session['session_id']}.pdf"
            )
        elif format.lower() == "csv":
            file_path = report_gen.generate_csv_report(current_session["audit_results"])
            return FileResponse(
                file_path,
                media_type="text/csv", 
                filename=f"audit_report_{current_session['session_id']}.csv"
            )
        else:
            raise HTTPException(status_code=400, detail="Invalid format. Use 'pdf' or 'csv'")
            
    except Exception as e:
        print(f"Error in download_report: {str(e)}")
        raise HTTPException(status_code=500, detail=f"Error generating report: {str(e)}")

@app.get("/health")
async def health_check():
    """Health check endpoint"""
    return {"status": "healthy", "timestamp": datetime.now().isoformat()}

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)