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
    name = re.sub(r"[^\w\-.]+", "_", name, flags=re.UNICODE)
    return name.strip("_")


def add_table(doc: Document, columns: List[str], rows: List[List[str]]):
    table = doc.add_table(rows=1, cols=len(columns))
    table.style = "Table Grid"

    header_cells = table.rows[0].cells
    for i, col in enumerate(columns):
        header_cells[i].text = str(col)

    for row in rows:
        cells = table.add_row().cells
        for i in range(len(columns)):
            value = row[i] if i < len(row) else ""
            cells[i].text = str(value) if value is not None else ""

    return table


def validate_report(report: Report):
    section_numbers = {section.number for section in report.sections}
    missing_sections = [i for i in range(1, 16) if i not in section_numbers]

    if missing_sections:
        raise HTTPException(
            status_code=400,
            detail=f"Missing sections: {missing_sections}"
        )

    for section in report.sections:
        if section.type not in ["paragraphs", "table"]:
            raise HTTPException(
                status_code=400,
                detail=f"Invalid section type in section {section.number}: {section.type}"
            )

        if section.type == "paragraphs":
            if section.content is None:
                raise HTTPException(
                    status_code=400,
                    detail=f"Section {section.number} is missing content"
                )

        if section.type == "table":
            if not section.columns:
                raise HTTPException(
                    status_code=400,
                    detail=f"Section {section.number} is missing table columns"
                )

            if not section.rows:
                raise HTTPException(
                    status_code=400,
                    detail=f"Section {section.number} is missing table rows"
                )

            # Source validation:
            # To support every language, the API assumes the Source column is the LAST column.
            # This avoids dependency on the word "Source" being translated.
            source_index = len(section.columns) - 1

            for row_i, row in enumerate(section.rows):
                if len(row) <= source_index:
                    raise HTTPException(
                        status_code=400,
                        detail=f"Row {row_i + 1} in section {section.number} has fewer cells than columns"
                    )

                source_value = str(row[source_index]).strip()

                if not source_value:
                    raise HTTPException(
                        status_code=400,
                        detail=f"Empty Source in section {section.number}, row {row_i + 1}"
                    )


@app.get("/")
def root():
    return {
        "status": "ok",
        "message": "Analyst-X DOCX Export API is running",
        "endpoint": "/generate-docx"
    }


@app.post("/generate-docx")
def generate_docx(payload: ExportRequest):
    report = payload.report
    validate_report(report)

    meta = report.metadata

    report_time = meta.time if meta.time and meta.time.strip() not in ["—", "-", ""] else datetime.now().strftime("%H-%M")
    report_time = report_time.replace(":", "-")

    if payload.file_name:
        file_name = payload.file_name
    else:
        file_name = f"{meta.company_name}_{meta.date}_{report_time}.docx"

    file_name = safe_filename(file_name)

    if not file_name.endswith(".docx"):
        file_name += ".docx"

    output_path = f"/tmp/{file_name}"

    doc = Document()

    doc.add_heading(meta.company_name, level=1)
    doc.add_heading("Market & Competitive Intelligence Report", level=2)

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
        doc.add_heading(f"{section.number}. {section.title}", level=2)

        if section.type == "table":
            add_table(doc, section.columns or [], section.rows or [])
        else:
            for paragraph in section.content or []:
                doc.add_paragraph(str(paragraph))

    doc.save(output_path)

    return FileResponse(
        output_path,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        filename=file_name
    )
