#!/usr/bin/env python3
"""Audit cleaned Markdown files against their source documents."""

from __future__ import annotations

import csv
import json
import re
import sys
import zipfile
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Iterable
from xml.etree import ElementTree as ET

import pandas as pd
import pypdfium2 as pdfium


ROOT = Path(".")
FILES_DIR = ROOT / "files"
MD_DIR = ROOT / "processed_md"
IMAGES_DIR = ROOT / "files_images"
REPORTS_DIR = ROOT / "reports"

SPACE_RE = re.compile(r"\s+")
MD_TABLE_RE = re.compile(r"^\s*\|.*\|\s*$", re.M)
BAD_EXCEL_TOKEN_RE = re.compile(r"\b(?:nan|unnamed:\s*\d+)\b", re.I)
AUTO_NUMBER_RE = re.compile(r"^\s*(?:#{1,6}\s*)?\d+(?:\.\d+)*\.\s+\S")
WORD_NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}


@dataclass
class Finding:
    severity: str
    message: str
    evidence: list[str] = field(default_factory=list)


@dataclass
class FileAudit:
    source: str
    markdown: str | None
    kind: str
    status: str = "ok"
    metrics: dict[str, object] = field(default_factory=dict)
    findings: list[Finding] = field(default_factory=list)

    def add(self, severity: str, message: str, evidence: Iterable[str] = ()) -> None:
        self.findings.append(Finding(severity, message, list(evidence)))
        order = {"ok": 0, "info": 1, "minor": 2, "major": 3, "critical": 4}
        if order[severity] > order[self.status]:
            self.status = severity


def canonical(text: str) -> str:
    text = text.replace("\u00a0", " ")
    text = text.replace("–", "-").replace("—", "-")
    text = text.replace("“", '"').replace("”", '"').replace("’", "'")
    return SPACE_RE.sub("", text).casefold()


def readable(text: str, limit: int = 180) -> str:
    text = SPACE_RE.sub(" ", text).strip()
    return text if len(text) <= limit else text[: limit - 3] + "..."


def significant(text: str) -> bool:
    stripped = SPACE_RE.sub(" ", text).strip()
    if len(stripped) < 6:
        return False
    return bool(re.search(r"[A-Za-zÀ-ỹ0-9]", stripped))


def block_variants(block: str) -> list[str]:
    variants = [block]
    if "\n" in block or "\r" in block:
        parts = [part.strip() for part in re.split(r"\r\n|\r|\n", block) if part.strip()]
        variants.extend(parts)
        variants.append("<br>".join(parts))
        variants.append(" ".join(parts))
    return variants


def coverage(blocks: list[str], md_text: str) -> tuple[int, int, list[str]]:
    unique_blocks = []
    seen = set()
    for block in blocks:
        block = SPACE_RE.sub(" ", block).strip()
        if not significant(block):
            continue
        key = canonical(" ".join(block_variants(block)))
        if key and key not in seen:
            unique_blocks.append(block)
            seen.add(key)

    md_canon = canonical(md_text)
    missing = [
        block
        for block in unique_blocks
        if not any(canonical(variant) in md_canon for variant in block_variants(block))
    ]
    return len(unique_blocks) - len(missing), len(unique_blocks), missing


def markdown_table_stats(md_text: str) -> dict[str, int]:
    rows = MD_TABLE_RE.findall(md_text)
    return {
        "table_rows": len(rows),
        "table_separator_rows": sum(1 for row in rows if re.search(r"\|\s*:?-{3,}:?\s*(?=\|)", row)),
    }


def docx_blocks(path: Path) -> list[str]:
    blocks: list[str] = []
    with zipfile.ZipFile(path) as zf:
        with zf.open("word/document.xml") as handle:
            root = ET.parse(handle).getroot()

    for child in root.findall(".//w:body/*", WORD_NS):
        if child.tag.endswith("}p"):
            texts = [node.text or "" for node in child.findall(".//w:t", WORD_NS)]
            text = "".join(texts).strip()
            if text:
                blocks.append(text)
        elif child.tag.endswith("}tbl"):
            for row in child.findall(".//w:tr", WORD_NS):
                cells = []
                for cell in row.findall("./w:tc", WORD_NS):
                    cell_text = " ".join(
                        "".join(node.text or "" for node in para.findall(".//w:t", WORD_NS)).strip()
                        for para in cell.findall(".//w:p", WORD_NS)
                    )
                    cells.append(SPACE_RE.sub(" ", cell_text).strip())
                row_text = " | ".join(cell for cell in cells if cell)
                if row_text:
                    blocks.append(row_text)
    return blocks


def excel_cells(path: Path) -> tuple[list[str], dict[str, object]]:
    sheets = pd.read_excel(path, sheet_name=None, header=None, dtype=object)
    cells: list[str] = []
    sheet_metrics: dict[str, object] = {}
    for sheet_name, df in sheets.items():
        non_empty = 0
        for value in df.to_numpy().flatten():
            if pd.isna(value):
                continue
            text = str(value).strip()
            if not text:
                continue
            non_empty += 1
            cells.append(text)
        sheet_metrics[str(sheet_name)] = {"shape": list(df.shape), "non_empty_cells": non_empty}
    return cells, {"sheets": sheet_metrics, "total_non_empty_cells": len(cells)}


def csv_cells(path: Path) -> tuple[list[str], dict[str, object]]:
    cells: list[str] = []
    rows = 0
    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.reader(handle)
        for row in reader:
            rows += 1
            for value in row:
                value = value.strip()
                if value:
                    cells.append(value)
    return cells, {"rows": rows, "total_non_empty_cells": len(cells)}


def pdf_pages_text(path: Path) -> tuple[list[str], dict[str, object]]:
    doc = pdfium.PdfDocument(str(path))
    pages: list[str] = []
    for page in doc:
        textpage = page.get_textpage()
        pages.append(textpage.get_text_range())
    return pages, {"pages": len(pages), "chars": sum(len(page) for page in pages)}


def audit_docx(src: Path, md_path: Path, md_text: str) -> FileAudit:
    audit = FileAudit(src.name, md_path.name, "docx")
    blocks = docx_blocks(src)
    found, total, missing = coverage(blocks, md_text)
    audit.metrics.update({"source_blocks": total, "covered_blocks": found, "coverage": round(found / total, 3) if total else 1})
    audit.metrics.update(markdown_table_stats(md_text))

    if total and found / total < 0.98:
        audit.add("major", "Thiếu hoặc biến đổi đáng kể nội dung từ DOCX.", map(readable, missing[:12]))

    bad_numbering = [
        line.strip()
        for line in md_text.splitlines()
        if AUTO_NUMBER_RE.match(line) and not line.strip().startswith("**")
    ]
    if bad_numbering:
        audit.add(
            "major",
            "DOCX còn dòng numbering Markdown tự động; quy định yêu cầu số thứ tự cứng dạng in đậm.",
            bad_numbering[:12],
        )
    return audit


def audit_excel(src: Path, md_path: Path, md_text: str) -> FileAudit:
    audit = FileAudit(src.name, md_path.name, src.suffix.lower().lstrip("."))
    cells, metrics = excel_cells(src)
    found, total, missing = coverage(cells, md_text)
    audit.metrics.update(metrics)
    audit.metrics.update({"covered_cells": found, "coverage": round(found / total, 3) if total else 1})
    audit.metrics.update(markdown_table_stats(md_text))

    bad_tokens = sorted(set(match.group(0) for match in BAD_EXCEL_TOKEN_RE.finditer(md_text)))
    if bad_tokens:
        audit.add("major", "Markdown còn token rác từ Excel.", bad_tokens)

    newline_cells = [cell for cell in cells if "\n" in cell or "\r" in cell]
    missing_br = []
    for cell in newline_cells:
        expected = re.sub(r"\r\n|\r|\n", "<br>", cell.strip())
        if canonical(expected) not in canonical(md_text):
            missing_br.append(readable(cell))
    if missing_br:
        audit.add("minor", "Một số ô Excel có xuống dòng nhưng chưa thấy dạng <br> tương ứng trong Markdown.", missing_br[:12])

    if total and found / total < 0.95:
        audit.add("major", "Thiếu hoặc biến đổi đáng kể nội dung ô từ bảng Excel.", map(readable, missing[:20]))

    if audit.metrics.get("table_rows", 0) == 0 and total > 10:
        audit.add("major", "Không phát hiện bảng Markdown dù file nguồn là bảng tính.", [])
    return audit


def audit_csv(src: Path, md_path: Path, md_text: str) -> FileAudit:
    audit = FileAudit(src.name, md_path.name, "csv")
    cells, metrics = csv_cells(src)
    found, total, missing = coverage(cells, md_text)
    audit.metrics.update(metrics)
    audit.metrics.update({"covered_cells": found, "coverage": round(found / total, 3) if total else 1})
    audit.metrics.update(markdown_table_stats(md_text))
    if total and found / total < 0.98:
        audit.add("major", "Thiếu hoặc biến đổi đáng kể nội dung từ CSV.", map(readable, missing[:20]))
    if audit.metrics.get("table_rows", 0) == 0 and total > 5:
        audit.add("major", "Không phát hiện bảng Markdown dù file nguồn là CSV.", [])
    return audit


def audit_pdf(src: Path, md_path: Path, md_text: str) -> FileAudit:
    audit = FileAudit(src.name, md_path.name, "pdf")
    pages, metrics = pdf_pages_text(src)
    found, total, missing = coverage(pages, md_text)
    image_dir = IMAGES_DIR / src.stem
    image_count = len(list(image_dir.glob("page_*.png"))) if image_dir.exists() else 0
    audit.metrics.update(metrics)
    audit.metrics.update(
        {
            "covered_pages_text": found,
            "page_text_coverage": round(found / total, 3) if total else 1,
            "rendered_images": image_count,
        }
    )
    audit.metrics.update(markdown_table_stats(md_text))

    if image_count != metrics["pages"]:
        audit.add("major", "Số ảnh render PDF không khớp số trang.", [f"pages={metrics['pages']}", f"images={image_count}"])

    if total and found / total < 0.8:
        audit.add(
            "major",
            "Text coverage PDF thấp; cần đối chiếu ảnh trang vì có thể thiếu bảng/cột/nội dung.",
            [readable(item) for item in missing[:6]],
        )
    elif missing:
        audit.add("minor", "Một số text page PDF không khớp nguyên văn sau làm sạch.", [readable(item) for item in missing[:4]])
    return audit


def audit_one(src: Path) -> FileAudit:
    md_path = MD_DIR / f"{src.stem}.md"
    if not md_path.exists():
        audit = FileAudit(src.name, None, src.suffix.lower().lstrip("."))
        audit.add("critical", "Không tìm thấy file Markdown tương ứng.", [])
        return audit

    md_text = md_path.read_text(encoding="utf-8")
    suffix = src.suffix.lower()
    if suffix == ".docx":
        return audit_docx(src, md_path, md_text)
    if suffix in {".xlsx", ".xls"}:
        return audit_excel(src, md_path, md_text)
    if suffix == ".csv":
        return audit_csv(src, md_path, md_text)
    if suffix == ".pdf":
        return audit_pdf(src, md_path, md_text)

    audit = FileAudit(src.name, md_path.name, suffix.lstrip("."))
    audit.add("minor", f"Chưa có bộ kiểm tra chuyên biệt cho định dạng {suffix}.", [])
    return audit


def to_dict(audit: FileAudit) -> dict[str, object]:
    return {
        "source": audit.source,
        "markdown": audit.markdown,
        "kind": audit.kind,
        "status": audit.status,
        "metrics": audit.metrics,
        "findings": [finding.__dict__ for finding in audit.findings],
    }


def write_markdown(audits: list[FileAudit], path: Path) -> None:
    summary = {}
    for audit in audits:
        summary[audit.status] = summary.get(audit.status, 0) + 1

    lines = [
        "# Báo cáo kiểm toán processed_md",
        "",
        f"- Thời điểm: {datetime.now().isoformat(timespec='seconds')}",
        f"- Tổng file gốc: {len(audits)}",
        f"- Tóm tắt trạng thái: {summary}",
        "",
        "## Kết quả chi tiết",
        "",
    ]

    for audit in audits:
        lines.extend(
            [
                f"### {audit.source}",
                "",
                f"- Markdown: `{audit.markdown or 'MISSING'}`",
                f"- Loại: `{audit.kind}`",
                f"- Trạng thái: `{audit.status}`",
                f"- Metrics: `{json.dumps(audit.metrics, ensure_ascii=False)}`",
            ]
        )
        if audit.findings:
            lines.append("- Findings:")
            for finding in audit.findings:
                lines.append(f"  - **{finding.severity}**: {finding.message}")
                for item in finding.evidence:
                    lines.append(f"    - {item}")
        else:
            lines.append("- Findings: không phát hiện lỗi bằng kiểm tra tự động.")
        lines.append("")

    path.write_text("\n".join(lines), encoding="utf-8")


def main() -> int:
    if not FILES_DIR.exists() or not MD_DIR.exists():
        print("Thiếu thư mục files/ hoặc processed_md/.", file=sys.stderr)
        return 2

    REPORTS_DIR.mkdir(exist_ok=True)
    sources = sorted(path for path in FILES_DIR.iterdir() if path.is_file())
    audits = [audit_one(path) for path in sources]

    json_path = REPORTS_DIR / "processed_md_audit.json"
    md_path = REPORTS_DIR / "processed_md_audit.md"
    json_path.write_text(json.dumps([to_dict(audit) for audit in audits], ensure_ascii=False, indent=2), encoding="utf-8")
    write_markdown(audits, md_path)

    worst = max(["ok", *[audit.status for audit in audits]], key={"ok": 0, "info": 1, "minor": 2, "major": 3, "critical": 4}.get)
    print(f"Wrote {md_path} and {json_path}. Worst status: {worst}")
    return 1 if worst in {"major", "critical"} else 0


if __name__ == "__main__":
    raise SystemExit(main())
