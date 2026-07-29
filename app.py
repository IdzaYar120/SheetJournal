"""
Flask application for automated Google Sheets Academic Journal creation.
Phase 1: File upload, parsing, and preview (Excel, CSV, Word, Google Doc).
Phase 2: Google Sheets integration — create and share journals.
"""

import json
import logging
import os
import re
import tempfile
from pathlib import Path

import pandas as pd
from flask import Flask, flash, redirect, render_template, request, session, url_for
from werkzeug.utils import secure_filename

from google_sheets import create_academic_journal, CREDENTIALS_PATH

logger = logging.getLogger(__name__)

app = Flask(__name__)
app.secret_key = os.urandom(32)

ALLOWED_EXTENSIONS = {"xlsx", "xls", "csv", "docx"}


def allowed_file(filename: str) -> bool:
    """Check if the uploaded file has an allowed extension."""
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


def extract_google_doc_id(url: str) -> str | None:
    """Extract Document ID from a Google Docs URL."""
    match = re.search(r"/document/d/([a-zA-Z0-9-_]+)", url)
    if match:
        return match.group(1)
    return None


def parse_docx_file(filepath: str) -> dict:
    """
    Parse an uploaded Word (.docx) file containing group list table.
    Expects a table containing student names (under a PІБ column) and disciplines.
    """
    try:
        import docx
    except ImportError:
        return {"error": "Бібліотеку 'python-docx' не встановлено на сервері."}

    try:
        doc = docx.Document(filepath)
        
        # 1. Search for Group Name in paragraphs and tables
        group_name = None
        for p in doc.paragraphs:
            text = p.text.strip()
            if not text:
                continue
            text_lower = text.lower()
            if any(kw in text_lower for kw in ["група", "group", "назва"]):
                if ":" in text:
                    group_name = text.split(":", 1)[1].strip()
                    break
                else:
                    words = text.split()
                    for idx, w in enumerate(words):
                        if any(kw in w.lower() for kw in ["група", "group"]):
                            if idx + 1 < len(words):
                                group_name = words[idx + 1].strip(" -:")
                                break
                    if group_name:
                        break

        # 2. Locate students/disciplines table
        students_table = None
        student_col_idx = None
        header_row_idx = 0
        student_col_keywords = ["піб", "прізвище", "студент", "ім'я", "name"]

        for table in doc.tables:
            for r_idx in range(min(len(table.rows), 3)):
                row_cells = [cell.text.strip().lower() for cell in table.rows[r_idx].cells]
                for idx, cell_text in enumerate(row_cells):
                    if any(kw in cell_text for kw in student_col_keywords):
                        students_table = table
                        student_col_idx = idx
                        header_row_idx = r_idx
                        break
                if students_table:
                    break
            if students_table:
                break

        if not students_table:
            return {"error": "Не знайдено таблицю зі списком студентів та дисциплін у документі Word."}

        # Retrieve group name from table if not found in paragraphs
        if not group_name:
            for r_idx in range(header_row_idx):
                cells = [c.text.strip() for c in students_table.rows[r_idx].cells]
                for idx, val in enumerate(cells):
                    if any(kw in val.lower() for kw in ["група", "group"]):
                        if idx + 1 < len(cells):
                            group_name = cells[idx + 1].strip(" -:")
                            break
                        
        if not group_name:
            group_name = "Не визначено"

        # 4. Parse headers
        header_cells = [c.text.strip() for c in students_table.rows[header_row_idx].cells]
        
        # 5. Extract Students and Disciplines
        students = []
        discipline_cols = []
        skip_keywords = ["№", "#", "номер", "num", header_cells[student_col_idx].lower()]

        for idx, col_name in enumerate(header_cells):
            if col_name.lower().strip() not in skip_keywords and idx != student_col_idx:
                discipline_cols.append((idx, col_name))

        disciplines_data = {col_idx: [] for col_idx, _ in discipline_cols}

        # Parse student rows
        for r_idx in range(header_row_idx + 1, len(students_table.rows)):
            row = students_table.rows[r_idx]
            cells = [c.text.strip() for c in row.cells]
            
            if len(cells) <= student_col_idx or not cells[student_col_idx]:
                continue
                
            students.append(cells[student_col_idx])
            
            for col_idx, _ in discipline_cols:
                val = cells[col_idx] if col_idx < len(cells) else ""
                disciplines_data[col_idx].append(val)

        # Build disciplines metadata
        disciplines = []
        for col_idx, col_name in discipline_cols:
            count = None
            for val in disciplines_data[col_idx]:
                if val.isdigit():
                    count = int(val)
                    break
            
            col_lower = col_name.lower().strip()
            is_cp = False
            for kw in ["курсов", "course project", "coursework", "course work"]:
                if kw in col_lower:
                    is_cp = True
                    break
            if not is_cp:
                tokens = col_lower.replace("(", " ").replace(")", " ").replace(".", " ").replace(",", " ").split()
                if any(t in tokens for t in ["кп", "кр", "kp", "kr"]):
                    is_cp = True

            if is_cp:
                control_type = "course_project"
            else:
                is_exam = any(kw in col_lower for kw in ["екзамен", "екз", "exam"])
                control_type = "exam" if is_exam else "credit"

            disciplines.append({
                "name": col_name,
                "class_count": count,
                "control_type": control_type
            })

        if not students:
            return {"error": "У таблиці документа Word не знайдено жодного студента."}

        return {
            "group_name": group_name,
            "students": students,
            "disciplines": disciplines,
            "error": None
        }
    except Exception as exc:
        return {"error": f"Помилка при зчитуванні Word-файлу: {exc}"}


def parse_uploaded_file(filepath: str) -> dict:
    """Parse an uploaded Excel, CSV, or Word file."""
    ext = Path(filepath).suffix.lower()
    
    if ext == ".docx":
        return parse_docx_file(filepath)

    try:
        if ext == ".csv":
            raw_df = pd.read_csv(filepath, header=None, dtype=str)
        else:
            raw_df = pd.read_excel(filepath, header=None, dtype=str, engine="openpyxl")

        if raw_df.empty:
            return {"error": "Файл порожній. Завантажте файл із даними."}

        group_name = None
        header_row_idx = 0

        first_row_values = raw_df.iloc[0].dropna().tolist()
        for col_idx in range(len(first_row_values)):
            cell = str(first_row_values[col_idx]).strip().lower()
            if any(kw in cell for kw in ["група", "group", "назва"]):
                if col_idx + 1 < len(first_row_values):
                    group_name = str(first_row_values[col_idx + 1]).strip()
                break

        if group_name is None and len(first_row_values) == 2:
            group_name = str(first_row_values[1]).strip()

        if group_name is None:
            group_name = "Не визначено"

        student_col_keywords = ["піб", "прізвище", "студент", "ім'я", "name"]
        header_found = False

        for row_idx in range(raw_df.shape[0]):
            row_cells = raw_df.iloc[row_idx].dropna().astype(str).str.strip().str.lower()
            for cell in row_cells:
                if any(kw in cell for kw in student_col_keywords):
                    header_row_idx = row_idx
                    header_found = True
                    break
            if header_found:
                break

        if not header_found:
            header_row_idx = 1 if raw_df.shape[0] > 1 else 0

        if ext == ".csv":
            df = pd.read_csv(filepath, header=header_row_idx, dtype=str)
        else:
            df = pd.read_excel(filepath, header=header_row_idx, dtype=str, engine="openpyxl")

        df.columns = df.columns.astype(str).str.strip()

        student_col = None
        for col in df.columns:
            if any(kw in col.lower() for kw in student_col_keywords):
                student_col = col
                break

        if student_col is None:
            student_col = df.columns[1] if len(df.columns) > 1 else df.columns[0]

        students = (
            df[student_col]
            .dropna()
            .astype(str)
            .str.strip()
            .loc[lambda s: s != ""]
            .tolist()
        )

        skip_keywords = ["№", "#", "номер", "num", student_col.lower()]
        discipline_cols = [
            col
            for col in df.columns
            if col.lower().strip() not in skip_keywords and col != student_col
        ]

        disciplines: list[dict] = []
        for col in discipline_cols:
            count = None
            series = pd.to_numeric(df[col], errors="coerce").dropna()
            if not series.empty:
                count = int(series.iloc[0])

            col_lower = col.lower().strip()
            is_cp = False
            for kw in ["курсов", "course project", "coursework", "course work"]:
                if kw in col_lower:
                    is_cp = True
                    break
            if not is_cp:
                tokens = col_lower.replace("(", " ").replace(")", " ").replace(".", " ").replace(",", " ").split()
                if any(t in tokens for t in ["кп", "кр", "kp", "kr"]):
                    is_cp = True

            if is_cp:
                control_type = "course_project"
            else:
                is_exam = any(kw in col_lower for kw in ["екзамен", "екз", "exam"])
                control_type = "exam" if is_exam else "credit"

            disciplines.append({
                "name": col,
                "class_count": count,
                "control_type": control_type
            })

        return {
            "group_name": group_name,
            "students": students,
            "disciplines": disciplines,
            "error": None,
        }

    except Exception as exc:
        return {"error": f"Помилка під час обробки файлу: {exc}"}


# ==========================================================================
# Routes
# ==========================================================================

@app.route("/", methods=["GET"])
def upload_page():
    """Render the file-upload page, displaying the service email if configured."""
    service_email = ""
    if CREDENTIALS_PATH.exists():
        try:
            with open(CREDENTIALS_PATH, "r", encoding="utf-8") as f:
                data = json.load(f)
                service_email = data.get("client_email", "")
        except Exception:
            pass
    return render_template("upload.html", service_email=service_email)


@app.route("/upload", methods=["POST"])
def handle_upload():
    """Receive the uploaded file, parse it, and redirect to preview."""
    if "file" not in request.files:
        flash("Файл не було надіслано.", "error")
        return redirect(url_for("upload_page"))

    file = request.files["file"]
    if file.filename == "" or file.filename is None:
        flash("Файл не обрано.", "error")
        return redirect(url_for("upload_page"))

    if not allowed_file(file.filename):
        flash("Недопустимий формат файлу. Дозволені: .xlsx, .xls, .csv, .docx", "error")
        return redirect(url_for("upload_page"))

    filename = secure_filename(file.filename)
    tmp_dir = tempfile.mkdtemp()
    filepath = os.path.join(tmp_dir, filename)
    file.save(filepath)

    result = parse_uploaded_file(filepath)

    try:
        os.remove(filepath)
        os.rmdir(tmp_dir)
    except OSError:
        pass

    if result.get("error"):
        flash(result["error"], "error")
        return redirect(url_for("upload_page"))

    session["parsed_data"] = json.dumps(result, ensure_ascii=False)

    return render_template(
        "preview.html",
        group_name=result["group_name"],
        students=result["students"],
        disciplines=result["disciplines"],
    )


@app.route("/import-gdoc", methods=["POST"])
def handle_gdoc_import():
    """Import and parse a Google Doc URL as a .docx file."""
    doc_url = request.form.get("doc_url", "").strip()
    if not doc_url:
        flash("Введіть посилання на Google Doc.", "error")
        return redirect(url_for("upload_page"))

    doc_id = extract_google_doc_id(doc_url)
    if not doc_id:
        flash("Неправильний формат посилання на Google Doc. Будь ласка, введіть коректне посилання.", "error")
        return redirect(url_for("upload_page"))

    tmp_dir = tempfile.mkdtemp()
    filename = f"{doc_id}.docx"
    filepath = os.path.join(tmp_dir, filename)

    try:
        from google.oauth2 import service_account
        import google.auth.transport.requests
        import requests

        scopes = ["https://www.googleapis.com/auth/drive.readonly"]
        if not CREDENTIALS_PATH.exists():
            flash("Файл credentials.json не знайдено. Перевірте конфігурацію сервера.", "error")
            return redirect(url_for("upload_page"))

        creds = service_account.Credentials.from_service_account_file(
            str(CREDENTIALS_PATH), scopes=scopes
        )
        auth_request = google.auth.transport.requests.Request()
        creds.refresh(auth_request)

        headers = {"Authorization": f"Bearer {creds.token}"}
        export_url = f"https://www.googleapis.com/drive/v3/files/{doc_id}/export?mimeType=application/vnd.openxmlformats-officedocument.wordprocessingml.document"

        response = requests.get(export_url, headers=headers)
        if response.status_code != 200:
            logger.error("Failed to export Google Doc: %s", response.text)
            flash(
                "Не вдалося завантажити Google Doc. Переконайтеся, що ви надали доступ для перегляду "
                "сервісному акаунту або зробили документ доступним за посиланням.",
                "error"
            )
            return redirect(url_for("upload_page"))

        with open(filepath, "wb") as f:
            f.write(response.content)

        result = parse_docx_file(filepath)

    except Exception as exc:
        logger.exception("Google Doc import failed")
        flash(f"Помилка при імпорті з Google Doc: {exc}", "error")
        return redirect(url_for("upload_page"))
    finally:
        try:
            if os.path.exists(filepath):
                os.remove(filepath)
            if os.path.exists(tmp_dir):
                os.rmdir(tmp_dir)
        except OSError:
            pass

    if result.get("error"):
        flash(result["error"], "error")
        return redirect(url_for("upload_page"))

    session["parsed_data"] = json.dumps(result, ensure_ascii=False)

    return render_template(
        "preview.html",
        group_name=result["group_name"],
        students=result["students"],
        disciplines=result["disciplines"],
    )


@app.route("/create-journal", methods=["POST"])
def create_journal():
    """Create a Google Sheets journal from previously parsed data."""
    raw = session.get("parsed_data")
    if not raw:
        flash("Дані не знайдено. Будь ласка, завантажте файл або Google Doc спочатку.", "error")
        return redirect(url_for("upload_page"))

    data = json.loads(raw)
    share_email = request.form.get("email", "").strip() or None
    folder_id = request.form.get("folder_id", "").strip() or None

    for idx, d in enumerate(data["disciplines"]):
        t_email = request.form.get(f"teacher_email_{idx}", "").strip() or None
        d["teacher_email"] = t_email

    try:
        result = create_academic_journal(
            group_name=data["group_name"],
            students=data["students"],
            disciplines=data["disciplines"],
            share_email=share_email,
            folder_id=folder_id,
        )
    except FileNotFoundError as exc:
        flash(str(exc), "error")
        return redirect(url_for("upload_page"))
    except Exception as exc:
        logger.exception("Failed to create Google Sheets journal")
        flash(f"Помилка при створенні Google Sheets: {exc}", "error")
        return redirect(url_for("upload_page"))

    session.pop("parsed_data", None)

    return render_template(
        "success.html",
        title=result["title"],
        spreadsheet_url=result["spreadsheet_url"],
        tab_count=len(data["disciplines"]) or 1,
        shared_with=share_email,
    )


if __name__ == "__main__":
    app.run(debug=True, port=5000)
