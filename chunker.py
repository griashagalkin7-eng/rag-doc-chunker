from pathlib import Path
import re
import hashlib
import json
import argparse
from typing import List, Dict



# =========================
# НАСТРОЙКИ
# =========================

BASE_DIR = Path(__file__).resolve().parent
INPUT_DIR = BASE_DIR / "input_docs"
OUTPUT_DIR = BASE_DIR / "output_md_chunks"

CHUNK_SIZE = 1800          # примерный размер чанка в символах
CHUNK_OVERLAP = 250        # пересечение между чанками
MIN_CHUNK_SIZE = 300       # слишком маленькие чанки будут склеиваться

SUPPORTED_EXTENSIONS = {
    ".txt",
    ".md",
    ".pdf",
    ".docx",
    ".html",
    ".htm",
}


# =========================
# ЧТЕНИЕ ФАЙЛОВ
# =========================

def read_txt_or_md(file_path: Path) -> str:
    return file_path.read_text(encoding="utf-8", errors="strict")


def read_pdf(file_path: Path) -> str:
    from pypdf import PdfReader

    reader = PdfReader(str(file_path))
    pages = []

    for page_num, page in enumerate(reader.pages, start=1):
        text = page.extract_text() or ""
        if text.strip():
            pages.append(f"\n\n[Страница {page_num}]\n{text}")

    return "\n".join(pages)


def read_docx(file_path: Path) -> str:
    from docx import Document
    from docx.table import Table
    from docx.text.paragraph import Paragraph

    doc = Document(str(file_path))
    blocks = []
    for element in doc.element.body:
        if element.tag.endswith("}p"):
            text = Paragraph(element, doc).text.strip()
            if text:
                blocks.append(text)
        elif element.tag.endswith("}tbl"):
            table = Table(element, doc)
            rows = [" | ".join(cell.text.strip() for cell in row.cells)
                    for row in table.rows]
            blocks.append("\n".join(rows))
    return "\n\n".join(blocks)


def read_html(file_path: Path) -> str:
    from bs4 import BeautifulSoup

    html = file_path.read_text(encoding="utf-8", errors="strict")
    soup = BeautifulSoup(html, "html.parser")

    for tag in soup(["script", "style", "nav", "footer", "header"]):
        tag.decompose()

    text = soup.get_text(separator="\n")
    return text


def extract_text(file_path: Path) -> str:
    ext = file_path.suffix.lower()

    if ext in [".txt", ".md"]:
        return read_txt_or_md(file_path)

    if ext == ".pdf":
        return read_pdf(file_path)

    if ext == ".docx":
        return read_docx(file_path)

    if ext in [".html", ".htm"]:
        return read_html(file_path)

    raise ValueError(f"Формат не поддерживается: {ext}")


# =========================
# ОЧИСТКА ТЕКСТА
# =========================

def clean_text(text: str) -> str:
    text = text.replace("\r\n", "\n").replace("\r", "\n")

    # Preserve Markdown indentation and code blocks.
    return text.strip("\n")


# =========================
# ЧАНКИНГ
# =========================

def split_into_paragraphs(text: str) -> List[str]:
    paragraphs = re.split(r"\n\s*\n", text)
    return [p.strip() for p in paragraphs if p.strip()]


def chunk_text(
    text: str,
    chunk_size: int = CHUNK_SIZE,
    overlap: int = CHUNK_OVERLAP,
    min_chunk_size: int = MIN_CHUNK_SIZE,
) -> List[str]:
    validate_chunk_settings(chunk_size, overlap, min_chunk_size)

    # Reserve room for the overlap and its paragraph separator.
    content_size = chunk_size - overlap - 2 if overlap else chunk_size
    paragraphs = split_into_paragraphs(text)

    chunks = []
    current_chunk = ""

    for paragraph in paragraphs:
        if len(current_chunk) + len(paragraph) + 2 <= content_size:
            current_chunk += "\n\n" + paragraph if current_chunk else paragraph
        else:
            if current_chunk:
                chunks.append(current_chunk.strip())

            # Если абзац сам по себе слишком большой — режем его по предложениям
            if len(paragraph) > content_size:
                sentence_chunks = split_large_paragraph(paragraph, content_size)
                chunks.extend(sentence_chunks)
                current_chunk = ""
            else:
                current_chunk = paragraph

    if current_chunk:
        chunks.append(current_chunk.strip())

    chunks = merge_small_chunks(chunks, min(min_chunk_size, content_size), content_size)
    chunks = add_overlap(chunks, overlap, chunk_size)

    return chunks


def validate_chunk_settings(chunk_size: int, overlap: int, min_chunk_size: int) -> None:
    if chunk_size <= 0:
        raise ValueError("chunk_size должен быть больше 0")

    if overlap < 0:
        raise ValueError("overlap не может быть отрицательным")

    if overlap >= chunk_size:
        raise ValueError("overlap должен быть меньше chunk_size")

    if overlap and chunk_size - overlap - 2 < 1:
        raise ValueError("chunk_size должен оставлять место для overlap и разделителя")

    if min_chunk_size < 0:
        raise ValueError("min_chunk_size не может быть отрицательным")

    if min_chunk_size > chunk_size:
        raise ValueError("min_chunk_size должен быть не больше chunk_size")


def split_large_paragraph(text: str, chunk_size: int) -> List[str]:
    sentences = re.split(r"(?<=[.!?…])\s+", text)
    chunks = []
    current = ""

    for sentence in sentences:
        if len(sentence) > chunk_size:
            if current:
                chunks.append(current.strip())
                current = ""

            chunks.extend(split_long_text(sentence, chunk_size))
            continue

        if len(current) + len(sentence) + 1 <= chunk_size:
            current += " " + sentence if current else sentence
        else:
            if current:
                chunks.append(current.strip())
            current = sentence

    if current:
        chunks.append(current.strip())

    return chunks


def split_long_text(text: str, chunk_size: int) -> List[str]:
    """Режет длинный текст без знаков конца предложения, сохраняя слова где возможно."""
    parts = []
    remaining = text.strip()

    while len(remaining) > chunk_size:
        split_at = remaining.rfind(" ", 0, chunk_size + 1)
        if split_at < chunk_size // 2:
            split_at = chunk_size

        parts.append(remaining[:split_at].strip())
        remaining = remaining[split_at:].strip()

    if remaining:
        parts.append(remaining)

    return parts


def merge_small_chunks(
    chunks: List[str],
    min_chunk_size: int,
    chunk_size: int,
) -> List[str]:
    if not chunks:
        return []

    merged = []
    buffer = ""

    for chunk in chunks:
        if len(chunk) < min_chunk_size:
            candidate = buffer + "\n\n" + chunk if buffer else chunk
            if buffer and len(candidate) > chunk_size:
                merged.append(buffer.strip())
                buffer = chunk
            else:
                buffer = candidate
        else:
            if buffer:
                combined = buffer + "\n\n" + chunk
                if len(combined) <= chunk_size:
                    merged.append(combined.strip())
                else:
                    merged.append(buffer.strip())
                    merged.append(chunk.strip())
                buffer = ""
            else:
                merged.append(chunk.strip())

    if buffer:
        if merged:
            combined = merged[-1] + "\n\n" + buffer
            if len(combined) <= chunk_size:
                merged[-1] = combined.strip()
            else:
                merged.append(buffer.strip())
        else:
            merged.append(buffer.strip())

    return merged


def add_overlap(chunks: List[str], overlap: int, chunk_size: int) -> List[str]:
    if overlap <= 0 or len(chunks) <= 1:
        return chunks

    result = []

    for i, chunk in enumerate(chunks):
        if i == 0:
            result.append(chunk)
            continue

        previous_chunk = chunks[i - 1]
        prefix_budget = max(0, chunk_size - len(chunk) - 2)
        budget = min(overlap, prefix_budget)
        overlap_text = previous_chunk[-budget:] if budget else ""
        # Avoid a partial first word when a boundary exists in the tail.
        if budget and len(previous_chunk) > budget and not previous_chunk[-budget - 1].isspace():
            boundary = re.search(r"\s+", overlap_text)
            if boundary:
                overlap_text = overlap_text[boundary.end():]
        overlap_text = overlap_text.strip()
        combined = f"{overlap_text}\n\n{chunk}" if overlap_text else chunk

        result.append(combined.strip())

    return result


# =========================
# СОХРАНЕНИЕ В MARKDOWN
# =========================

def make_chunk_id(source_file: str, chunk_number: int, chunk_text: str) -> str:
    raw = json.dumps([source_file, chunk_number, chunk_text], ensure_ascii=False)
    return hashlib.sha256(raw.encode("utf-8")).hexdigest()[:20]


def build_markdown(
    source_path: Path,
    chunks: List[str],
    input_dir: Path = INPUT_DIR,
) -> str:
    source_name = source_path.resolve().relative_to(input_dir.resolve()).as_posix()
    md_parts = []

    md_parts.append(f"# RAG chunks: {source_path.name}")
    md_parts.append("")
    md_parts.append("## Metadata")
    md_parts.append("")
    md_parts.append(f"- source_file: `{source_path.name}`")
    md_parts.append(f"- source_extension: `{source_path.suffix.lower()}`")
    md_parts.append(f"- total_chunks: `{len(chunks)}`")
    md_parts.append(f"- chunk_size: `{CHUNK_SIZE}`")
    md_parts.append(f"- chunk_overlap: `{CHUNK_OVERLAP}`")
    md_parts.append("")

    for index, chunk in enumerate(chunks, start=1):
        chunk_id = make_chunk_id(source_name, index, chunk)

        md_parts.append("---")
        md_parts.append("")
        md_parts.append(f"## Chunk {index}")
        md_parts.append("")
        md_parts.append("```yaml")
        md_parts.append(f"chunk_id: {chunk_id}")
        md_parts.append(f"source_file: {json.dumps(source_name, ensure_ascii=False)}")
        md_parts.append(f"chunk_index: {index}")
        md_parts.append(f"total_chunks: {len(chunks)}")
        md_parts.append(f"characters: {len(chunk)}")
        md_parts.append("```")
        md_parts.append("")
        md_parts.append(chunk)
        md_parts.append("")

    return "\n".join(md_parts)


def save_markdown(source_path: Path, markdown_text: str, input_dir: Path = INPUT_DIR, output_dir: Path = OUTPUT_DIR) -> Path:
    output_dir.mkdir(parents=True, exist_ok=True)

    output_file = output_dir / f"{make_output_stem(source_path, input_dir)}_chunks.md"
    output_file.write_text(markdown_text, encoding="utf-8")

    return output_file


def make_output_stem(source_path: Path, input_dir: Path = INPUT_DIR) -> str:
    try:
        relative = source_path.resolve().relative_to(input_dir.resolve())
    except ValueError:
        relative = Path(source_path.name)

    digest = hashlib.sha256(relative.as_posix().encode("utf-8")).hexdigest()[:16]
    safe_name = re.sub(r"[^\w.-]+", "_", relative.name)[:100]
    return f"{safe_name}_{digest}"


# =========================
# ОСНОВНОЙ ПРОЦЕСС
# =========================

def process_file(file_path: Path, input_dir: Path = INPUT_DIR, output_dir: Path = OUTPUT_DIR) -> Dict:
    raw_text = extract_text(file_path)
    cleaned = clean_text(raw_text)

    if not cleaned:
        return {
            "file": str(file_path),
            "status": "skipped",
            "reason": "empty_text",
        }

    chunks = chunk_text(cleaned)
    markdown_text = build_markdown(file_path, chunks, input_dir)
    output_file = save_markdown(file_path, markdown_text, input_dir, output_dir)

    return {
        "file": str(file_path),
        "status": "processed",
        "chunks": len(chunks),
        "output": str(output_file),
    }


def process_folder(input_dir: Path = INPUT_DIR, output_dir: Path = OUTPUT_DIR) -> List[Dict]:
    if not input_dir.exists():
        raise FileNotFoundError(f"Входная папка не найдена: {input_dir}")

    results = []

    for file_path in sorted(input_dir.rglob("*")):
        if output_dir.resolve() in file_path.resolve().parents:
            continue
        if not file_path.is_file():
            continue

        if file_path.suffix.lower() not in SUPPORTED_EXTENSIONS:
            results.append({
                "file": str(file_path),
                "status": "skipped",
                "reason": "unsupported_format",
            })
            continue

        try:
            result = process_file(file_path, input_dir, output_dir)
            results.append(result)
        except Exception as e:
            results.append({
                "file": str(file_path),
                "status": "error",
                "reason": str(e),
            })

    return results


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Split documents into Markdown chunks for RAG.")
    parser.add_argument("--input", type=Path, default=INPUT_DIR)
    parser.add_argument("--output", type=Path, default=OUTPUT_DIR)
    args = parser.parse_args()
    if args.input.resolve() == args.output.resolve():
        parser.error("Input and output directories must differ")
    if not args.input.is_dir():
        parser.error(f"Input directory not found: {args.input}")
    results = process_folder(args.input, args.output)

    print("\n=== Результат обработки ===\n")

    for item in results:
        if item["status"] == "processed":
            print(f"[OK] {item['file']}")
            print(f"     chunks: {item['chunks']}")
            print(f"     output: {item['output']}")
        elif item["status"] == "skipped":
            print(f"[SKIP] {item['file']}")
            print(f"       reason: {item['reason']}")
        else:
            print(f"[ERROR] {item['file']}")
            print(f"        reason: {item['reason']}")

    raise SystemExit(1 if any(item["status"] == "error" for item in results) else 0)
