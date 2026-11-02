from pathlib import Path
import os
import zipfile
import tempfile
import fitz
from docx import Document
from pptx import Presentation

OLD = "Anant" + " Guhey"
NEW = "Aditya Rohira"
ROOT = Path(".")

TEXT_EXTENSIONS = {
    ".md", ".txt", ".csv", ".json", ".yml", ".yaml", ".toml",
    ".py", ".js", ".jsx", ".ts", ".tsx", ".html", ".css", ".xml",
    ".sql", ".r", ".ipynb"
}

def replace_paragraph(paragraph):
    before = "".join(run.text for run in paragraph.runs)
    if OLD not in before:
        return False
    after = before.replace(OLD, NEW)
    if paragraph.runs:
        paragraph.runs[0].text = after
        for run in paragraph.runs[1:]:
            run.text = ""
    else:
        paragraph.add_run(after)
    return True

def process_docx(path):
    doc = Document(path)
    changed = False
    for paragraph in doc.paragraphs:
        changed |= replace_paragraph(paragraph)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    changed |= replace_paragraph(paragraph)
    for section in doc.sections:
        for part in (section.header, section.footer, section.first_page_header,
                     section.first_page_footer, section.even_page_header,
                     section.even_page_footer):
            for paragraph in part.paragraphs:
                changed |= replace_paragraph(paragraph)
            for table in part.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for paragraph in cell.paragraphs:
                            changed |= replace_paragraph(paragraph)
    if changed:
        doc.save(path)
    return changed

def process_pptx(path):
    prs = Presentation(path)
    changed = False

    def visit_shapes(shapes):
        nonlocal changed
        for shape in shapes:
            if shape.has_text_frame:
                for paragraph in shape.text_frame.paragraphs:
                    changed |= replace_paragraph(paragraph)
            if shape.has_table:
                for row in shape.table.rows:
                    for cell in row.cells:
                        for paragraph in cell.text_frame.paragraphs:
                            changed |= replace_paragraph(paragraph)
            if shape.shape_type == 6:
                visit_shapes(shape.shapes)

    for slide in prs.slides:
        visit_shapes(slide.shapes)
    if changed:
        prs.save(path)
    return changed

def process_pdf(path):
    doc = fitz.open(path)
    changed = False
    for page in doc:
        hits = page.search_for(OLD)
        if not hits:
            continue
        for rect in hits:
            page.add_redact_annot(rect, fill=(1, 1, 1))
        page.apply_redactions()
        for rect in hits:
            size = min(10.0, max(6.0, rect.height * 0.78))
            inserted = False
            while size >= 5.0:
                spare = page.insert_textbox(
                    rect, NEW, fontname="helv", fontsize=size,
                    color=(0, 0, 0), overlay=True
                )
                if spare >= 0:
                    inserted = True
                    break
                size -= 0.5
            if not inserted:
                print(f"WARNING: replacement may not fit in {path}, page {page.number + 1}")
        changed = True
    if changed:
        temp = str(path) + ".tmp"
        doc.save(temp, garbage=4, deflate=True)
        doc.close()
        os.replace(temp, path)
    else:
        doc.close()
    return changed

def process_twbx(path):
    changed = False
    temp_path = str(path) + ".tmp"
    with zipfile.ZipFile(path, "r") as source, zipfile.ZipFile(temp_path, "w", zipfile.ZIP_DEFLATED) as target:
        for item in source.infolist():
            data = source.read(item.filename)
            if OLD.encode("utf-8") in data:
                data = data.replace(OLD.encode("utf-8"), NEW.encode("utf-8"))
                changed = True
            target.writestr(item, data)
    if changed:
        os.replace(temp_path, path)
    else:
        os.remove(temp_path)
    return changed

def process_text(path):
    try:
        original = path.read_text(encoding="utf-8")
    except (UnicodeDecodeError, OSError):
        return False
    updated = original.replace(OLD, NEW)
    if updated != original:
        path.write_text(updated, encoding="utf-8")
        return True
    return False

def main():
    changed_files = []
    for path in list(ROOT.rglob("*")):
        if not path.is_file() or ".git" in path.parts:
            continue
        suffix = path.suffix.lower()
        try:
            changed = False
            if suffix == ".docx":
                changed = process_docx(path)
            elif suffix == ".pptx":
                changed = process_pptx(path)
            elif suffix == ".pdf":
                changed = process_pdf(path)
            elif suffix == ".twbx":
                changed = process_twbx(path)
            elif suffix in TEXT_EXTENSIONS:
                changed = process_text(path)
            if changed:
                changed_files.append(str(path))
        except Exception as exc:
            print(f"ERROR processing {path}: {exc}")
            raise

    # Rename repository files whose names contain the old author name.
    for path in sorted(ROOT.rglob("*"), key=lambda p: len(p.parts), reverse=True):
        if path.is_file() and OLD.lower() in path.name.lower():
            new_name = path.name.replace(OLD, NEW)
            if new_name == path.name:
                new_name = path.name.replace("Anant_Guhey", "Aditya_Rohira")
            destination = path.with_name(new_name)
            path.rename(destination)
            changed_files.append(f"{path} -> {destination}")

    print("Updated files:")
    for name in changed_files:
        print(f" - {name}")
    print(f"Total changed files: {len(changed_files)}")

if __name__ == "__main__":
    main()
