from fastapi import FastAPI, HTTPException
from fastapi.responses import FileResponse
from pydantic import BaseModel
from typing import List, Optional
from docx import Document
import re
from datetime import datetime

app = FastAPI(title="Analyst-X DOCX Export API")

class Section(BaseModel):
    number: int
    title: str
    type: str
    content: Optional[List[str]] = None
    columns: Optional[List[str]] = None
    rows: Optional[List[List[str]]] = None

class Metadata(BaseModel):
    company_name: str
    date: str
    time: str
    company_type: str
    research_goal: str
    language: str
    comparison_company: str
    confidence: str
    confidence_rationale: str

class Report(BaseModel):
    metadata: Metadata
    sections: List[Section]

class ExportRequest(BaseModel):
    file_name: Optional[str] = None
    report: Report

def safe_filename(name: str) -> str:
    name = re.sub(r"[^\w\-.]+", "_", name)
    return name.strip("_")

def add_table(doc: Document, columns: List[str], rows: List[List[str]]):
    table = doc.add_table(rows=1, cols=len(columns))
    table.style = "Table Grid"

    header_cells = table.rows[0].cells
    for i, col in enumerate(columns):
        header_cells[i].text = str(col)

    for row in rows:
        cells = table.add_row().cells
        for i, value in enumerate(row):
            cells[i].text = str(value) if value else ""

def validate_report(report: Report):
    section_numbers = {s.number for s in report.sections}
    missing = [i for i in range(1, 16) if i not in section_numbers]
    if missing:
        raise HTTPException(status_code=400, detail=f"Missing sections: {missing}")

    for section in report.sections:
        if section.type == "table":
            if not section.columns or not section.rows:
                raise HTTPException(status_code=400, detail=f"Section {section.number} missing table data")

            if "Source" in section.columns:
                source_index = section.columns.index("Source")
                for i, row in enumerate(section.rows):
                    if len(row) <= source_index or not str(row[source_index]).strip():
                        raise HTTPException(
                            status_code=400,
                            detail=f"Missing Source in section {section.number}, row {i+1}"
                        )

@app.post("/generate-docx")
def generate_docx(payload: ExportRequest):
    report = payload.report
    validate_report(report)

    meta = report.metadata

    time_val = meta.time if meta.time and meta.time != "—" else datetime.now().strftime("%H-%M")
    time_val = time_val.replace(":", "-")

    filename = payload.file_name or f"{meta.company_name}_{meta.date}_{time_val}.docx"
    filename = safe_filename(filename)

    doc = Document()

    doc.add_heading(meta.company_name, 1)
    doc.add_heading("Market & Competitive Intelligence Report", 2)

    doc.add_paragraph(f"Date: {meta.date}")
    doc.add_paragraph(f"Time: {meta.time}")
    doc.add_paragraph(f"Company Type: {meta.company_type}")
    doc.add_paragraph(f"Research Goal: {meta.research_goal}")
    doc.add_paragraph(f"Language: {meta.language}")
    doc.add_paragraph(f"Comparison Company: {meta.comparison_company}")

    doc.add_paragraph("")
    doc.add_paragraph(f"Confidence: {meta.confidence}")
    doc.add_paragraph(f"Confidence rationale: {meta.confidence_rationale}")

    for section in sorted(report.sections, key=lambda s: s.number):
        doc.add_heading(f"{section.number}. {section.title}", 2)

        if section.type == "table":
            add_table(doc, section.columns, section.rows)
        else:
            for line in section.content or []:
                doc.add_paragraph(line)

    path = f"/tmp/{filename}"
    doc.save(path)

    return FileResponse(path, filename=filename)
