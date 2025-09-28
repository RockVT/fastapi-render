from fastapi import FastAPI, UploadFile, Form
from fastapi.responses import JSONResponse

app = FastAPI()

@app.post("/process_gsheet")
async def process_gsheet(pdf: UploadFile, gsheet_url: str = Form(...)):
    content = await pdf.read()
    size_kb = round(len(content) / 1024, 2)

    return JSONResponse({
        "message": "API nhận thành công!",
        "pdf_filename": pdf.filename,
        "pdf_size_kb": size_kb,
        "gsheet_url": gsheet_url
    })
