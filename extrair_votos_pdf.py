#!/usr/bin/env python3
import argparse
import base64
import json
import re
import sys
import unicodedata
from pathlib import Path
from typing import Any

from openpyxl import Workbook
from pypdf import PdfReader

COLUMNS = [
    "NOME DO ARQUIVO",
    "PROMOTORIA DE JUSTICA",
    "NUMERO DO PROCEDIMENTO",
    "TIPO DO PROCEDIMENTO",
    "OBJETO",
    "RELATOR",
    "EMENTA",
    "CASO EM EXAME",
    "CONCLUSAO DO RELATOR",
    "NOME DO PROMOTOR DE JUSTICA QUE ARQUIVOU",
    "DATA DE HOMOLOGACAO",
    "LINK CURTO PARA O ARQUIVO",
]

TYPE_PATTERNS = [
    r"\b(Inqu[eé]rito\s+Civil)\b",
    r"\b(Procedimento\s+Preparat[oó]rio)\b",
    r"\b(Not[ií]cia\s+de\s+Fato)\b",
    r"\b(Procedimento\s+Administrativo|PA)\b",
]

DATE_PATTERNS = [
    r"DATA\s+DE\s+HOMOLOGA[CÇ][AÃ]O\s*[:\-]\s*([0-3]?\d/[01]?\d/\d{4})",
    r"HOMOLOGA[CÇ][AÃ]O\s+EM\s*[:\-]\s*([0-3]?\d/[01]?\d/\d{4})",
]


def to_single_line(value: str) -> str:
    value = re.sub(r"\s+", " ", value or "").strip()
    return value


def extract_first(patterns: list[str], text: str, flags: int = 0) -> str:
    for pattern in patterns:
        match = re.search(pattern, text, flags)
        if match:
            return to_single_line(match.group(1))
    return ""


def read_pdf_text(path: Path) -> str:
    reader = PdfReader(str(path))
    chunks = []
    for page in reader.pages:
        chunks.append(page.extract_text() or "")
    text = "\n".join(chunks)
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n[ \t]+", "\n", text)
    return text.strip()


def normalize_token(value: str) -> str:
    base = unicodedata.normalize("NFD", value)
    base = "".join(ch for ch in base if unicodedata.category(ch) != "Mn")
    return re.sub(r"\s+", " ", base).strip().upper()


def non_empty_lines(text: str) -> list[str]:
    return [to_single_line(line) for line in text.split("\n") if to_single_line(line)]


def find_line_index(lines: list[str], startswith_options: list[str]) -> int:
    options = [normalize_token(v) for v in startswith_options]
    for idx, line in enumerate(lines):
        norm = normalize_token(line)
        for option in options:
            if norm.startswith(option):
                return idx
    return -1


def paragraph_after_heading(lines: list[str], headings: list[str]) -> str:
    idx = find_line_index(lines, headings)
    if idx < 0:
        return ""
    if idx + 1 >= len(lines):
        return ""
    return lines[idx + 1]


def extract_line_after_colon(lines: list[str], headings: list[str]) -> str:
    idx = find_line_index(lines, headings)
    if idx < 0:
        return ""
    line = lines[idx]
    if ":" in line:
        return to_single_line(line.split(":", 1)[1])
    if idx + 1 < len(lines):
        return lines[idx + 1]
    return ""


def extract_objeto(lines: list[str]) -> str:
    if len(lines) < 3:
        return ""
    start = 2
    end = find_line_index(lines, ["RELATOR", "RELATORA"])
    if end < 0:
        end = len(lines)
    if end <= start:
        return ""
    return to_single_line(" ".join(lines[start:end]))


def extract_numero_procedimento(lines: list[str], text: str) -> str:
    if len(lines) >= 2:
        second_line = lines[1]
        proc = re.search(r"\d{5}\.\d{3}\.\d{3}-\d{4}", second_line)
        if proc:
            return proc.group(0)
        any_number = re.search(r"\d[\d.\-]{6,}", second_line)
        if any_number:
            return any_number.group(0)
    proc = re.search(r"\d{5}\.\d{3}\.\d{3}-\d{4}", text)
    return proc.group(0) if proc else ""


def extract_promotor_nome(text: str) -> str:
    patterns = [
        r"Promotor(?:a)?\s+de\s+Justi[cç]a\s*[:\-–]\s*([A-ZÁÉÍÓÚÂÊÔÃÕÇ][^\n,;.]{3,})",
        r"([A-ZÁÉÍÓÚÂÊÔÃÕÇ][A-Za-zÁÉÍÓÚÂÊÔÃÕÇáéíóúâêôãõç ]{3,})\s*,?\s+Promotor(?:a)?\s+de\s+Justi[cç]a",
    ]
    for pattern in patterns:
        match = re.search(pattern, text, flags=re.IGNORECASE)
        if match:
            return to_single_line(match.group(1))
    return ""


def extract_block_after_heading(text: str, heading: str, next_headings: list[str]) -> str:
    heading_rx = re.escape(heading).replace("\\ ", r"\s+")
    next_union = "|".join(
        re.escape(v).replace("\\ ", r"\s+").replace("\\.", r"\.?")
        for v in next_headings
    )
    pattern = rf"{heading_rx}\s*[:\-]?\s*(.+?)(?:\n(?:{next_union})\b|$)"
    match = re.search(pattern, text, flags=re.IGNORECASE | re.DOTALL)
    return to_single_line(match.group(1)) if match else ""


def extract_record(pdf_path: Path, file_id: str = "", file_name: str = "") -> dict[str, str]:
    text = read_pdf_text(pdf_path)
    lines = non_empty_lines(text)
    record = {col: "" for col in COLUMNS}

    record["NOME DO ARQUIVO"] = to_single_line(file_name) if file_name else pdf_path.name
    record["PROMOTORIA DE JUSTICA"] = lines[0] if lines else ""
    record["NUMERO DO PROCEDIMENTO"] = extract_numero_procedimento(lines, text)

    tipo = extract_first(TYPE_PATTERNS, text, flags=re.IGNORECASE)
    record["TIPO DO PROCEDIMENTO"] = tipo

    record["OBJETO"] = extract_objeto(lines)
    record["RELATOR"] = extract_line_after_colon(lines, ["RELATOR", "RELATORA"])
    record["EMENTA"] = extract_line_after_colon(lines, ["EMENTA"])
    if not record["EMENTA"]:
        record["EMENTA"] = extract_block_after_heading(
            text,
            "EMENTA",
            ["CASO EM EXAME", "CASO EXAMINADO", "CONCLUSAO DO RELATOR", "CONCLUSAO", "RELATOR"],
        )

    record["CASO EM EXAME"] = paragraph_after_heading(lines, ["CASO EM EXAME", "CASO EXAMINADO"])
    record["CONCLUSAO DO RELATOR"] = paragraph_after_heading(
        lines, ["CONCLUSAO DO RELATOR", "CONCLUSÃO DO RELATOR"]
    )

    record["NOME DO PROMOTOR DE JUSTICA QUE ARQUIVOU"] = extract_promotor_nome(text)
    record["DATA DE HOMOLOGACAO"] = extract_first(DATE_PATTERNS, text, flags=re.IGNORECASE)

    if file_id:
        record["LINK CURTO PARA O ARQUIVO"] = f"https://drive.google.com/file/d/{file_id}/view"

    return record


def autosize_columns(ws: Any) -> None:
    for column_cells in ws.columns:
        max_len = 0
        col_letter = column_cells[0].column_letter
        for cell in column_cells:
            value = "" if cell.value is None else str(cell.value)
            max_len = max(max_len, len(value))
        ws.column_dimensions[col_letter].width = min(max(max_len + 2, 16), 80)


def build_workbook(rows: list[dict[str, Any]], output_path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Votos"
    ws.append(COLUMNS)

    for row in rows:
        ws.append([to_single_line(str(row.get(col, ""))) for col in COLUMNS])

    autosize_columns(ws)
    wb.save(output_path)


def cmd_extract(args: argparse.Namespace) -> int:
    pdf_path = Path(args.pdf).expanduser().resolve()
    if not pdf_path.exists():
        print(f"Arquivo PDF nao encontrado: {pdf_path}", file=sys.stderr)
        return 2

    record = extract_record(
        pdf_path=pdf_path,
        file_id=args.file_id or "",
        file_name=args.file_name or "",
    )
    print(json.dumps(record, ensure_ascii=False))
    return 0


def cmd_build(args: argparse.Namespace) -> int:
    try:
        rows_json = base64.b64decode(args.rows_b64.encode("utf-8")).decode("utf-8")
        rows = json.loads(rows_json)
    except Exception as exc:
        print(f"Falha ao decodificar rows_b64: {exc}", file=sys.stderr)
        return 2

    if not isinstance(rows, list):
        print("rows_b64 deve conter uma lista JSON", file=sys.stderr)
        return 2

    output = Path(args.output).expanduser().resolve()
    output.parent.mkdir(parents=True, exist_ok=True)
    build_workbook(rows=rows, output_path=output)
    print(str(output))
    return 0


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Extrai campos de votos em PDF e gera planilha.")
    subparsers = parser.add_subparsers(dest="cmd", required=True)

    extract_parser = subparsers.add_parser("extract", help="Extrai campos de um PDF")
    extract_parser.add_argument("--pdf", required=True, help="Caminho do PDF")
    extract_parser.add_argument("--file-id", required=False, help="ID do arquivo no Google Drive")
    extract_parser.add_argument("--file-name", required=False, help="Nome original do arquivo PDF")

    build_parser = subparsers.add_parser("build", help="Gera XLSX a partir de registros")
    build_parser.add_argument("--rows-b64", required=True, help="Lista JSON em base64")
    build_parser.add_argument("--output", required=True, help="Caminho de saida do XLSX")

    return parser.parse_args()


def main() -> int:
    args = parse_args()
    if args.cmd == "extract":
        return cmd_extract(args)
    if args.cmd == "build":
        return cmd_build(args)
    return 2


if __name__ == "__main__":
    raise SystemExit(main())
