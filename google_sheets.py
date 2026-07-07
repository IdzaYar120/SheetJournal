"""
Google Sheets integration for Academic Journal creation.

Uses a Service Account to create and format Google Spreadsheets.
Expects a credentials JSON file at the path specified by CREDENTIALS_PATH.
"""

import logging
from pathlib import Path

import gspread
from gspread.utils import rowcol_to_a1

logger = logging.getLogger(__name__)

CREDENTIALS_PATH = Path(__file__).parent / "credentials.json"


def _get_client() -> gspread.Client:
    """Authenticate and return a gspread client using Service Account credentials."""
    if not CREDENTIALS_PATH.exists():
        raise FileNotFoundError(
            f"Файл credentials.json не знайдено за шляхом: {CREDENTIALS_PATH}\n"
            "Завантажте JSON-ключ сервісного акаунту з Google Cloud Console."
        )
    return gspread.service_account(filename=str(CREDENTIALS_PATH))


def _col_letter(col_index: int) -> str:
    """Convert a 1-based column index to a spreadsheet column letter (A, B, …, Z, AA, …)."""
    result = ""
    while col_index > 0:
        col_index, remainder = divmod(col_index - 1, 26)
        result = chr(65 + remainder) + result
    return result


def create_academic_journal(
    group_name: str,
    students: list[str],
    disciplines: list[dict],
    share_email: str | None = None,
    folder_id: str | None = None,
) -> dict:
    """
    Create a fully-formatted Google Sheets academic journal.

    Args:
        group_name:  Name of the student group (e.g. "КН-21").
        students:    List of student full names.
        disciplines: List of dicts with keys "name" (str) and "class_count" (int|None).
        share_email: Optional Google email to share the spreadsheet with (editor role).
        folder_id:   Optional Google Drive folder ID to place the spreadsheet in.

    Returns:
        A dict with "spreadsheet_id", "spreadsheet_url", and "title".
    """
    client = _get_client()

    title = f"Журнал — {group_name}"
    spreadsheet = client.create(title)

    # Move to specific folder if provided.
    if folder_id:
        try:
            client.move_spreadsheet(spreadsheet.id, folder_id)
        except Exception as exc:
            logger.warning("Could not move spreadsheet to folder %s: %s", folder_id, exc)

    # Share with the user's personal Google account.
    if share_email:
        spreadsheet.share(share_email, perm_type="user", role="writer")
        # Also make it accessible via link for convenience.
        spreadsheet.share("", perm_type="anyone", role="reader")

    # ---------------------------------------------------------------
    # Create one worksheet (tab) per discipline.
    # ---------------------------------------------------------------
    existing_sheet = spreadsheet.sheet1  # Default "Sheet1" — will be renamed or deleted.
    created_first = False

    for disc in disciplines:
        disc_name = disc["name"]
        class_count = disc.get("class_count") or 0

        if not created_first:
            # Rename the default sheet instead of creating a new one.
            worksheet = existing_sheet
            worksheet.update_title(disc_name)
            created_first = True
        else:
            worksheet = spreadsheet.add_worksheet(title=disc_name, rows=1, cols=1)

        _populate_discipline_sheet(worksheet, students, disc_name, class_count, group_name)

    # If no disciplines were provided, set up a generic sheet.
    if not created_first:
        existing_sheet.update_title(group_name)
        _populate_discipline_sheet(existing_sheet, students, group_name, 0, group_name)

    return {
        "spreadsheet_id": spreadsheet.id,
        "spreadsheet_url": spreadsheet.url,
        "title": title,
    }


def _populate_discipline_sheet(
    worksheet: gspread.Worksheet,
    students: list[str],
    discipline_name: str,
    class_count: int,
    group_name: str,
) -> None:
    """
    Fill and format a single discipline worksheet.

    Layout:
        Row 1: Group name (merged across header) + discipline name
        Row 2: Headers — "№" | "ПІБ студента" | 1 | 2 | 3 | … | N
        Row 3+: Student data rows.
    """
    # Ensure minimum class columns.
    class_count = max(class_count, 1)
    total_cols = 2 + class_count  # "№" + "ПІБ" + class columns
    total_rows = 2 + len(students)  # 2 header rows + student rows

    # Resize the worksheet to fit the data.
    worksheet.resize(rows=max(total_rows, 20), cols=total_cols)

    # ----- Build all cell values in a 2D list -----
    rows_data: list[list] = []

    # Row 1: Group + discipline label (will be merged visually).
    row1 = [f"{group_name} — {discipline_name}"] + [""] * (total_cols - 1)
    rows_data.append(row1)

    # Row 2: Column headers.
    header = ["№", "ПІБ студента"] + [str(i) for i in range(1, class_count + 1)]
    rows_data.append(header)

    # Student rows.
    for idx, name in enumerate(students, start=1):
        row = [idx, name] + [""] * class_count
        rows_data.append(row)

    # Write everything in a single batch call.
    end_cell = rowcol_to_a1(len(rows_data), total_cols)
    worksheet.update(f"A1:{end_cell}", rows_data, value_input_option="RAW")

    # ----- Formatting -----
    # We use batch_format to minimise API calls.
    last_col_letter = _col_letter(total_cols)
    formats = []

    # Row 1: Title row — bold, larger font, merged, background colour.
    formats.append({
        "range": f"A1:{last_col_letter}1",
        "format": {
            "textFormat": {"bold": True, "fontSize": 13},
            "horizontalAlignment": "CENTER",
            "verticalAlignment": "MIDDLE",
            "backgroundColor": {"red": 0.24, "green": 0.31, "blue": 0.71},  # Indigo-ish
            "textFormat": {"bold": True, "fontSize": 13, "foregroundColor": {"red": 1, "green": 1, "blue": 1}},
        },
    })

    # Merge row 1 across all columns.
    worksheet.merge_cells(f"A1:{last_col_letter}1", merge_type="MERGE_ALL")

    # Row 2: Header row — bold, centred, light background.
    formats.append({
        "range": f"A2:{last_col_letter}2",
        "format": {
            "textFormat": {"bold": True, "fontSize": 10},
            "horizontalAlignment": "CENTER",
            "verticalAlignment": "MIDDLE",
            "backgroundColor": {"red": 0.9, "green": 0.92, "blue": 0.98},
        },
    })

    # Student name column (B) — left-aligned.
    last_student_row = 2 + len(students)
    formats.append({
        "range": f"B3:B{last_student_row}",
        "format": {
            "horizontalAlignment": "LEFT",
            "verticalAlignment": "MIDDLE",
        },
    })

    # Number column (A) + class columns (C…) — centred.
    formats.append({
        "range": f"A3:A{last_student_row}",
        "format": {"horizontalAlignment": "CENTER", "verticalAlignment": "MIDDLE"},
    })
    if class_count > 0:
        class_start = _col_letter(3)
        class_end = _col_letter(total_cols)
        formats.append({
            "range": f"{class_start}3:{class_end}{last_student_row}",
            "format": {"horizontalAlignment": "CENTER", "verticalAlignment": "MIDDLE"},
        })

    # Apply all formats.
    worksheet.batch_format(formats)

    # Freeze the first 2 rows and first 2 columns (№ + ПІБ).
    worksheet.freeze(rows=2, cols=2)

    # Set column widths: "№" narrow, "ПІБ" wide, class cols narrow.
    requests = [
        {
            "updateDimensionProperties": {
                "range": {
                    "sheetId": worksheet.id,
                    "dimension": "COLUMNS",
                    "startIndex": 0,
                    "endIndex": 1,
                },
                "properties": {"pixelSize": 40},
                "fields": "pixelSize",
            }
        },
        {
            "updateDimensionProperties": {
                "range": {
                    "sheetId": worksheet.id,
                    "dimension": "COLUMNS",
                    "startIndex": 1,
                    "endIndex": 2,
                },
                "properties": {"pixelSize": 280},
                "fields": "pixelSize",
            }
        },
    ]
    if class_count > 0:
        requests.append({
            "updateDimensionProperties": {
                "range": {
                    "sheetId": worksheet.id,
                    "dimension": "COLUMNS",
                    "startIndex": 2,
                    "endIndex": total_cols,
                },
                "properties": {"pixelSize": 35},
                "fields": "pixelSize",
            }
        })

    # Add borders to the entire data area.
    border_style = {"style": "SOLID", "color": {"red": 0.7, "green": 0.7, "blue": 0.7}}
    requests.append({
        "updateBorders": {
            "range": {
                "sheetId": worksheet.id,
                "startRowIndex": 1,  # From header row (0-indexed)
                "endRowIndex": last_student_row,
                "startColumnIndex": 0,
                "endColumnIndex": total_cols,
            },
            "top": border_style,
            "bottom": border_style,
            "left": border_style,
            "right": border_style,
            "innerHorizontal": border_style,
            "innerVertical": border_style,
        }
    })

    worksheet.spreadsheet.batch_update({"requests": requests})
