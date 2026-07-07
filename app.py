"""
Flask application for automated Google Sheets Academic Journal creation.
Phase 1: File upload, parsing, and preview.
Phase 2: Google Sheets integration — create and share journals.
"""

import json
import logging
import os
import tempfile
from pathlib import Path

import pandas as pd
from flask import Flask, flash, redirect, render_template, request, session, url_for
from werkzeug.utils import secure_filename

from google_sheets import create_academic_journal

logger = logging.getLogger(__name__)

app = Flask(__name__)
app.secret_key = os.urandom(32)

ALLOWED_EXTENSIONS = {"xlsx", "xls", "csv"}


def allowed_file(filename: str) -> bool:
    """Check if the uploaded file has an allowed extension."""
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


def parse_uploaded_file(filepath: str) -> dict:
    """
    Parse an uploaded Excel or CSV file and extract academic journal data.

    Expected file structure:
        Row 0: "Назва групи" (or similar) in col A, actual group name in col B.
        Row 1: Empty separator row (optional — skipped automatically).
        Row 2+: Header row with "№", "ПІБ" (student name column),
                 followed by discipline names as remaining column headers.
        Subsequent rows: student data. The discipline columns contain
                         the number of classes (кількість пар) for each.

    Returns a dict with keys: group_name, students, disciplines, error.
    """
    ext = Path(filepath).suffix.lower()

    try:
        # ------------------------------------------------------------------
        # 1. Read the raw file into a DataFrame (no header inference).
        # ------------------------------------------------------------------
        if ext == ".csv":
            raw_df = pd.read_csv(filepath, header=None, dtype=str)
        else:
            raw_df = pd.read_excel(filepath, header=None, dtype=str, engine="openpyxl")

        if raw_df.empty:
            return {"error": "Файл порожній. Завантажте файл із даними."}

        # ------------------------------------------------------------------
        # 2. Extract the Group Name from the first row.
        # ------------------------------------------------------------------
        group_name = None
        header_row_idx = 0  # will be updated once we find the header

        first_row_values = raw_df.iloc[0].dropna().tolist()
        # Strategy: look for a cell whose neighbour (or itself) holds
        # a keyword indicating the group name row.
        for col_idx in range(len(first_row_values)):
            cell = str(first_row_values[col_idx]).strip().lower()
            if any(kw in cell for kw in ["група", "group", "назва"]):
                # The group name is the next non-empty cell in the same row.
                if col_idx + 1 < len(first_row_values):
                    group_name = str(first_row_values[col_idx + 1]).strip()
                break

        # Fallback: if the first row has exactly two cells and no keyword
        # matched, treat the second cell as the group name.
        if group_name is None and len(first_row_values) == 2:
            group_name = str(first_row_values[1]).strip()

        if group_name is None:
            group_name = "Не визначено"

        # ------------------------------------------------------------------
        # 3. Locate the header row (contains "ПІБ" or "Прізвище" or "Студент").
        # ------------------------------------------------------------------
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
            # If no keyword found, assume the second row (index 1) is the header.
            header_row_idx = 1 if raw_df.shape[0] > 1 else 0

        # ------------------------------------------------------------------
        # 4. Re-read with the detected header row.
        # ------------------------------------------------------------------
        if ext == ".csv":
            df = pd.read_csv(filepath, header=header_row_idx, dtype=str)
        else:
            df = pd.read_excel(
                filepath, header=header_row_idx, dtype=str, engine="openpyxl"
            )

        df.columns = df.columns.astype(str).str.strip()

        # ------------------------------------------------------------------
        # 5. Identify the student-name column.
        # ------------------------------------------------------------------
        student_col = None
        for col in df.columns:
            if any(kw in col.lower() for kw in student_col_keywords):
                student_col = col
                break

        if student_col is None:
            # Heuristic: pick the second column (first is usually "№").
            student_col = df.columns[1] if len(df.columns) > 1 else df.columns[0]

        # ------------------------------------------------------------------
        # 6. Extract the student list (drop NaN / empty rows).
        # ------------------------------------------------------------------
        students = (
            df[student_col]
            .dropna()
            .astype(str)
            .str.strip()
            .loc[lambda s: s != ""]
            .tolist()
        )

        # ------------------------------------------------------------------
        # 7. Identify discipline columns and their class counts.
        #    Everything after the student column (excluding "№"-like cols)
        #    is treated as a discipline.
        # ------------------------------------------------------------------
        skip_keywords = ["№", "#", "номер", "num", student_col.lower()]
        discipline_cols = [
            col
            for col in df.columns
            if col.lower().strip() not in skip_keywords and col != student_col
        ]

        disciplines: list[dict] = []
        for col in discipline_cols:
            # Try to parse the class count from the first non-empty value,
            # or from the column header itself if it contains a number suffix.
            count = None
            series = pd.to_numeric(df[col], errors="coerce").dropna()
            if not series.empty:
                # Use the first student's value as the representative count.
                count = int(series.iloc[0])
            disciplines.append({"name": col, "class_count": count})

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
    """Render the file-upload page."""
    return render_template("upload.html")


@app.route("/upload", methods=["POST"])
def handle_upload():
    """Receive the uploaded file, parse it, and redirect to the preview."""
    if "file" not in request.files:
        flash("Файл не було надіслано.", "error")
        return redirect(url_for("upload_page"))

    file = request.files["file"]
    if file.filename == "" or file.filename is None:
        flash("Файл не обрано.", "error")
        return redirect(url_for("upload_page"))

    if not allowed_file(file.filename):
        flash("Недопустимий формат файлу. Дозволені: .xlsx, .xls, .csv", "error")
        return redirect(url_for("upload_page"))

    # Save to a temporary location for parsing.
    filename = secure_filename(file.filename)
    tmp_dir = tempfile.mkdtemp()
    filepath = os.path.join(tmp_dir, filename)
    file.save(filepath)

    result = parse_uploaded_file(filepath)

    # Clean up the temp file.
    try:
        os.remove(filepath)
        os.rmdir(tmp_dir)
    except OSError:
        pass

    if result["error"]:
        flash(result["error"], "error")
        return redirect(url_for("upload_page"))

    # Store parsed data in session so it survives the redirect to /create-journal.
    session["parsed_data"] = json.dumps(result, ensure_ascii=False)

    return render_template(
        "preview.html",
        group_name=result["group_name"],
        students=result["students"],
        disciplines=result["disciplines"],
    )


@app.route("/create-journal", methods=["POST"])
def create_journal():
    """Create a Google Sheets journal from the previously parsed data."""
    raw = session.get("parsed_data")
    if not raw:
        flash("Дані не знайдено. Будь ласка, завантажте файл спочатку.", "error")
        return redirect(url_for("upload_page"))

    data = json.loads(raw)
    share_email = request.form.get("email", "").strip() or None
    folder_id = request.form.get("folder_id", "").strip() or None

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

    # Clear session data after successful creation.
    session.pop("parsed_data", None)

    return render_template(
        "success.html",
        title=result["title"],
        spreadsheet_url=result["spreadsheet_url"],
        tab_count=len(data["disciplines"]) or 1,
        shared_with=share_email,
    )


# ==========================================================================
# Entry point
# ==========================================================================

if __name__ == "__main__":
    app.run(debug=True, port=5000)
