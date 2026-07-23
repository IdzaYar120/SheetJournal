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

    # Share with each teacher
    teacher_emails = set(d.get("teacher_email") for d in disciplines if d.get("teacher_email"))
    for t_email in teacher_emails:
        if t_email != share_email:
            try:
                spreadsheet.share(t_email, perm_type="user", role="writer")
            except Exception as exc:
                logger.warning("Could not share spreadsheet with teacher %s: %s", t_email, exc)

    # ---------------------------------------------------------------
    # ---------------------------------------------------------------
    # Create one worksheet (tab) per discipline / course project.
    # ---------------------------------------------------------------
    existing_sheet = spreadsheet.sheet1  # Default "Sheet1" — will be renamed or deleted.
    created_first = False

    for disc in disciplines:
        disc_name = disc["name"]
        class_count = disc.get("class_count") or 0
        control_type = disc.get("control_type", "credit")
        teacher_email = disc.get("teacher_email")

        # Determine tab title: course projects get КП prefix
        if control_type == "course_project":
            tab_title = f"КП — {disc_name}"
        else:
            tab_title = disc_name

        if not created_first:
            # Rename the default sheet instead of creating a new one.
            worksheet = existing_sheet
            worksheet.update_title(tab_title)
            created_first = True
        else:
            worksheet = spreadsheet.add_worksheet(title=tab_title, rows=1, cols=1)

        if control_type == "course_project":
            _populate_course_project_sheet(worksheet, students, disc_name, group_name, teacher_email)
        else:
            _populate_discipline_sheet(worksheet, students, disc_name, class_count, group_name, control_type, teacher_email)

    # If no disciplines were provided, set up a generic sheet.
    if not created_first:
        existing_sheet.update_title(group_name)
        _populate_discipline_sheet(existing_sheet, students, group_name, 0, group_name, "credit", None)

    # Create milestone sheets РК1 and РК2, semester sheet, and general summary report sheet
    if students:
        actual_disciplines = disciplines if disciplines else [{"name": group_name, "class_count": 0, "control_type": "credit"}]
        
        # РК1 and РК2 only include non-course-project disciplines
        milestone_disciplines = [d for d in actual_disciplines if d.get("control_type", "credit") != "course_project"]
        if not milestone_disciplines:
            milestone_disciplines = [{"name": group_name, "class_count": 0, "control_type": "credit"}]
            
        _create_milestone_sheet(spreadsheet, students, milestone_disciplines, 1, group_name)
        _create_milestone_sheet(spreadsheet, students, milestone_disciplines, 2, group_name)
        _create_semester_sheet(spreadsheet, students, actual_disciplines, group_name)
        _create_summary_report_sheet(spreadsheet, students, actual_disciplines, group_name)

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
    control_type: str = "credit",
    teacher_email: str | None = None,
) -> None:
    """
    Fill and format a single discipline worksheet.

    Layout:
        Row 1: Group name (merged across header) + discipline name
        Row 2: Headers — "№" | "ПІБ студента" | 1 | 2 | 3 | … | N [ | Екзамен (макс. 40) ]
        Row 3+: Student data rows.
    """
    # Ensure minimum class columns.
    class_count = max(class_count, 1)
    is_exam = (control_type == "exam")
    total_cols = 2 + class_count
    if is_exam:
        total_cols += 1

    total_rows = 2 + len(students)  # 2 header rows + student rows

    # Resize the worksheet to fit the data.
    worksheet.resize(rows=max(total_rows, 20), cols=total_cols)

    # ----- Build all cell values in a 2D list -----
    rows_data: list[list] = []

    # Row 1: Group + discipline label (will be merged visually).
    label = f"{group_name} — {discipline_name}"
    if is_exam:
        label += " (Екзамен)"
    if teacher_email:
        label += f" (Викладач: {teacher_email})"
    row1 = [label] + [""] * (total_cols - 1)
    rows_data.append(row1)

    # Row 2: Column headers.
    header = ["№", "ПІБ студента"] + [str(i) for i in range(1, class_count + 1)]
    if is_exam:
        header.append("Екзамен (макс. 40)")
    rows_data.append(header)

    # Student rows.
    for idx, name in enumerate(students, start=1):
        row = [idx, name] + [""] * class_count
        if is_exam:
            row.append("")
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
        end_class_idx = 2 + class_count
        requests.append({
            "updateDimensionProperties": {
                "range": {
                    "sheetId": worksheet.id,
                    "dimension": "COLUMNS",
                    "startIndex": 2,
                    "endIndex": end_class_idx,
                },
                "properties": {"pixelSize": 35},
                "fields": "pixelSize",
            }
        })
        if is_exam:
            requests.append({
                "updateDimensionProperties": {
                    "range": {
                        "sheetId": worksheet.id,
                        "dimension": "COLUMNS",
                        "startIndex": end_class_idx,
                        "endIndex": end_class_idx + 1,
                    },
                    "properties": {"pixelSize": 90},
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


def _create_milestone_sheet(
    spreadsheet: gspread.Spreadsheet,
    students: list[str],
    disciplines: list[dict],
    milestone: int,  # 1 for РК1, 2 for РК2
    group_name: str,
) -> gspread.Worksheet:
    """
    Create and format a milestone control worksheet (РК1 or РК2).

    Contains the student list and auto-calculating AVERAGE formulas for each
    discipline based on the first or second half of class columns.
    """
    title = f"РК{milestone}"
    total_cols = 2 + len(disciplines)
    total_rows = 2 + len(students)

    # Create the worksheet
    worksheet = spreadsheet.add_worksheet(title=title, rows=max(total_rows, 20), cols=max(total_cols, 5))

    # Build data
    rows_data: list[list] = []

    # Row 1: Title
    title_text = f"Рубіжний контроль {milestone} (РК{milestone}) — {group_name}"
    rows_data.append([title_text] + [""] * (total_cols - 1))

    # Row 2: Headers
    headers = ["№", "ПІБ студента"] + [d["name"] for d in disciplines]
    rows_data.append(headers)

    # Determine the first sheet to link names from
    first_sheet_name = disciplines[0]["name"] if disciplines else group_name

    # Row 3+: Student rows
    for idx in range(1, len(students) + 1):
        sheet_row = 2 + idx
        # Dynamically reference student's name from the first sheet
        student_ref = f"='{first_sheet_name}'!B{sheet_row}"

        row = [idx, student_ref]

        # Add formulas for each discipline
        for d in disciplines:
            d_name = d["name"]
            num_classes = d.get("class_count") or 0
            num_classes = max(num_classes, 1)
            half = (num_classes + 1) // 2

            if milestone == 1:
                # First half of classes (Columns C to 2+half)
                col_start = _col_letter(3)
                col_end = _col_letter(2 + half)
            else:
                # Second half of classes
                if num_classes <= 1:
                    col_start = _col_letter(3)
                    col_end = _col_letter(3)
                else:
                    col_start = _col_letter(3 + half)
                    col_end = _col_letter(2 + num_classes)

            # We use IFERROR and AVERAGE to avoid division by zero when cells are empty.
            # Using commas for arguments since the API expects US/English formulas.
            formula = f"=IFERROR(AVERAGE('{d_name}'!{col_start}{sheet_row}:{col_end}{sheet_row}), \"\")"
            row.append(formula)

        rows_data.append(row)

    # Write to sheet
    end_cell = rowcol_to_a1(len(rows_data), total_cols)
    worksheet.update(f"A1:{end_cell}", rows_data, value_input_option="USER_ENTERED")

    # Formatting
    last_col_letter = _col_letter(total_cols)
    formats = []

    # Row 1 format: Bold, white text, different color depending on milestone
    if milestone == 1:
        bg_color = {"red": 0.44, "green": 0.16, "blue": 0.39}  # Purple
    else:
        bg_color = {"red": 0.1, "green": 0.45, "blue": 0.45}  # Dark Teal

    formats.append({
        "range": f"A1:{last_col_letter}1",
        "format": {
            "textFormat": {"bold": True, "fontSize": 13, "foregroundColor": {"red": 1, "green": 1, "blue": 1}},
            "horizontalAlignment": "CENTER",
            "verticalAlignment": "MIDDLE",
            "backgroundColor": bg_color,
        },
    })

    worksheet.merge_cells(f"A1:{last_col_letter}1", merge_type="MERGE_ALL")

    # Row 2 format: Header
    formats.append({
        "range": f"A2:{last_col_letter}2",
        "format": {
            "textFormat": {"bold": True, "fontSize": 10},
            "horizontalAlignment": "CENTER",
            "verticalAlignment": "MIDDLE",
            "backgroundColor": {"red": 0.9, "green": 0.92, "blue": 0.98},
        },
    })

    # Student name column
    formats.append({
        "range": f"B3:B{total_rows}",
        "format": {
            "horizontalAlignment": "LEFT",
            "verticalAlignment": "MIDDLE",
        },
    })

    # Other columns: centered
    formats.append({
        "range": f"A3:A{total_rows}",
        "format": {"horizontalAlignment": "CENTER", "verticalAlignment": "MIDDLE"},
    })

    if len(disciplines) > 0:
        disc_start = _col_letter(3)
        formats.append({
            "range": f"{disc_start}3:{last_col_letter}{total_rows}",
            "format": {"horizontalAlignment": "CENTER", "verticalAlignment": "MIDDLE"},
        })

    worksheet.batch_format(formats)
    worksheet.freeze(rows=2, cols=2)

    # Column widths
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
    if len(disciplines) > 0:
        requests.append({
            "updateDimensionProperties": {
                "range": {
                    "sheetId": worksheet.id,
                    "dimension": "COLUMNS",
                    "startIndex": 2,
                    "endIndex": total_cols,
                },
                "properties": {"pixelSize": 120},
                "fields": "pixelSize",
            }
        })

    # Borders
    border_style = {"style": "SOLID", "color": {"red": 0.7, "green": 0.7, "blue": 0.7}}
    requests.append({
        "updateBorders": {
            "range": {
                "sheetId": worksheet.id,
                "startRowIndex": 1,
                "endRowIndex": total_rows,
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
    return worksheet


def _create_semester_sheet(
    spreadsheet: gspread.Spreadsheet,
    students: list[str],
    disciplines: list[dict],
    group_name: str,
) -> gspread.Worksheet:
    """
    Create and format a semester summary worksheet ("Семестр").

    Contains the student list and auto-calculating formulas that average the
    РК1 and РК2 grades for each discipline. For exam disciplines, it handles
    a 60/40 split (60% semester average + 40% exam grade).
    """
    title = "Семестр"

    # Calculate total columns based on control type of each discipline
    total_cols = 2
    for d in disciplines:
        if d.get("control_type", "credit") == "exam":
            total_cols += 3
        else:
            total_cols += 1

    total_rows = 2 + len(students)

    # Create the worksheet
    worksheet = spreadsheet.add_worksheet(title=title, rows=max(total_rows, 20), cols=max(total_cols, 5))

    # Build data
    rows_data: list[list] = []

    # Row 1: Title
    title_text = f"Семестровий контроль (РК1 + РК2) — {group_name}"
    rows_data.append([title_text] + [""] * (total_cols - 1))

    # Row 2: Headers
    headers = ["№", "ПІБ студента"]
    for d in disciplines:
        d_name = d["name"]
        ctype = d.get("control_type", "credit")
        if ctype == "exam":
            headers.extend([
                f"{d_name} (Сем. 60%)",
                f"{d_name} (Екз. 40%)",
                f"{d_name} (Всього)"
            ])
        elif ctype == "course_project":
            headers.append(f"{d_name} (КП)")
        else:
            headers.append(f"{d_name} (Залік)")
    rows_data.append(headers)

    # Row 3+: Student rows
    for idx in range(1, len(students) + 1):
        sheet_row = 2 + idx
        # Dynamically reference student's name from РК1 sheet
        student_ref = f"='РК1'!B{sheet_row}"

        row = [idx, student_ref]

        rk_disc_counter = 0
        # Add formulas for each discipline
        for d in disciplines:
            d_name = d["name"]
            control_type = d.get("control_type", "credit")

            if control_type == "exam":
                rk_col_letter = _col_letter(3 + rk_disc_counter)
                rk_disc_counter += 1

                # 1. Semester score (60%): average of РК1 and РК2 * 0.6
                formula_sem = f"=IFERROR(AVERAGE('РК1'!{rk_col_letter}{sheet_row}, 'РК2'!{rk_col_letter}{sheet_row}) * 0.6, \"\")"
                row.append(formula_sem)

                # 2. Exam score (40%): link to individual discipline sheet's last column (Exam column)
                class_count = d.get("class_count") or 0
                class_count = max(class_count, 1)
                # Exam column is at index 3 + class_count (since №, ПІБ + classes cols + exam col)
                exam_col_letter = _col_letter(3 + class_count)
                formula_exam = f"='{d_name}'!{exam_col_letter}{sheet_row}"
                row.append(formula_exam)

                # 3. Total (Sum of Semester 60% and Exam 40%)
                c1_idx = len(row) - 1  # 1-based index of Semester 60% col
                c2_idx = len(row)      # 1-based index of Exam 40% col
                c1_letter = _col_letter(c1_idx)
                c2_letter = _col_letter(c2_idx)
                formula_total = f"=IFERROR(SUM({c1_letter}{sheet_row}, {c2_letter}{sheet_row}), \"\")"
                row.append(formula_total)
            elif control_type == "course_project":
                # Link to 'КП — {d_name}'!G{sheet_row} (Column G is "Середній бал")
                formula_kp = f"='КП — {d_name}'!G{sheet_row}"
                row.append(formula_kp)
            else:
                # Credit (100%): average of РК1 and РК2
                rk_col_letter = _col_letter(3 + rk_disc_counter)
                rk_disc_counter += 1
                formula_credit = f"=IFERROR(AVERAGE('РК1'!{rk_col_letter}{sheet_row}, 'РК2'!{rk_col_letter}{sheet_row}), \"\")"
                row.append(formula_credit)

        rows_data.append(row)

    # Write to sheet
    end_cell = rowcol_to_a1(len(rows_data), total_cols)
    worksheet.update(f"A1:{end_cell}", rows_data, value_input_option="USER_ENTERED")

    # Formatting
    last_col_letter = _col_letter(total_cols)
    formats = []

    # Row 1 format: Bold, white text, dark blue/indigo color
    bg_color = {"red": 0.08, "green": 0.18, "blue": 0.36}  # Dark Navy Blue

    formats.append({
        "range": f"A1:{last_col_letter}1",
        "format": {
            "textFormat": {"bold": True, "fontSize": 13, "foregroundColor": {"red": 1, "green": 1, "blue": 1}},
            "horizontalAlignment": "CENTER",
            "verticalAlignment": "MIDDLE",
            "backgroundColor": bg_color,
        },
    })

    worksheet.merge_cells(f"A1:{last_col_letter}1", merge_type="MERGE_ALL")

    # Row 2 format: Header
    formats.append({
        "range": f"A2:{last_col_letter}2",
        "format": {
            "textFormat": {"bold": True, "fontSize": 10},
            "horizontalAlignment": "CENTER",
            "verticalAlignment": "MIDDLE",
            "backgroundColor": {"red": 0.9, "green": 0.92, "blue": 0.98},
        },
    })

    # Student name column
    formats.append({
        "range": f"B3:B{total_rows}",
        "format": {
            "horizontalAlignment": "LEFT",
            "verticalAlignment": "MIDDLE",
        },
    })

    # Other columns: centered
    formats.append({
        "range": f"A3:A{total_rows}",
        "format": {"horizontalAlignment": "CENTER", "verticalAlignment": "MIDDLE"},
    })

    if total_cols > 2:
        disc_start = _col_letter(3)
        formats.append({
            "range": f"{disc_start}3:{last_col_letter}{total_rows}",
            "format": {"horizontalAlignment": "CENTER", "verticalAlignment": "MIDDLE"},
        })

    worksheet.batch_format(formats)
    worksheet.freeze(rows=2, cols=2)

    # Column widths
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

    col_idx = 2
    for d in disciplines:
        if d.get("control_type", "credit") == "exam":
            requests.append({
                "updateDimensionProperties": {
                    "range": {
                        "sheetId": worksheet.id,
                        "dimension": "COLUMNS",
                        "startIndex": col_idx,
                        "endIndex": col_idx + 1,
                    },
                    "properties": {"pixelSize": 110},
                    "fields": "pixelSize",
                }
            })
            requests.append({
                "updateDimensionProperties": {
                    "range": {
                        "sheetId": worksheet.id,
                        "dimension": "COLUMNS",
                        "startIndex": col_idx + 1,
                        "endIndex": col_idx + 2,
                    },
                    "properties": {"pixelSize": 90},
                    "fields": "pixelSize",
                }
            })
            requests.append({
                "updateDimensionProperties": {
                    "range": {
                        "sheetId": worksheet.id,
                        "dimension": "COLUMNS",
                        "startIndex": col_idx + 2,
                        "endIndex": col_idx + 3,
                    },
                    "properties": {"pixelSize": 90},
                    "fields": "pixelSize",
                }
            })
            col_idx += 3
        else:
            requests.append({
                "updateDimensionProperties": {
                    "range": {
                        "sheetId": worksheet.id,
                        "dimension": "COLUMNS",
                        "startIndex": col_idx,
                        "endIndex": col_idx + 1,
                    },
                    "properties": {"pixelSize": 120},
                    "fields": "pixelSize",
                }
            })
            col_idx += 1

    # Borders
    border_style = {"style": "SOLID", "color": {"red": 0.7, "green": 0.7, "blue": 0.7}}
    requests.append({
        "updateBorders": {
            "range": {
                "sheetId": worksheet.id,
                "startRowIndex": 1,
                "endRowIndex": total_rows,
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
    return worksheet


def _populate_course_project_sheet(
    worksheet: gspread.Worksheet,
    students: list[str],
    discipline_name: str,
    group_name: str,
    teacher_email: str | None = None,
) -> None:
    """
    Fill and format a course project (КП/КР) worksheet.

    Layout:
        Row 1: Group name + "Курсовий проект з дисципліни:" + discipline name
        Row 2: Headers — "№" | "ПІБ студента" | "Тема курсового проекту / роботи"
                         | "Член комісії 1" | "Член комісії 2" | "Член комісії 3" | "Середній бал"
        Row 3+: Student rows (referencing first sheet name if appropriate).
    """
    total_cols = 7  # №, ПІБ, Тема, 3 members, Average
    total_rows = 2 + len(students)

    worksheet.resize(rows=max(total_rows, 20), cols=total_cols)

    rows_data: list[list] = []

    # Row 1: Title
    title_text = f"{group_name} — Захист курсового проекту з дисципліни: {discipline_name}"
    if teacher_email:
        title_text += f" (Викладач: {teacher_email})"
    rows_data.append([title_text] + [""] * (total_cols - 1))

    # Row 2: Headers
    headers = [
        "№",
        "ПІБ студента",
        "Тема курсового проекту / роботи",
        "Член комісії 1",
        "Член комісії 2",
        "Член комісії 3",
        "Середній бал"
    ]
    rows_data.append(headers)

    for idx, name in enumerate(students, start=1):
        sheet_row = 2 + idx
        # Formula for average: AVERAGE(D{row}:F{row})
        formula_avg = f"=IFERROR(AVERAGE(D{sheet_row}:F{sheet_row}), \"\")"
        row = [
            idx,
            name,
            "",  # Theme (empty, to be filled by teacher)
            "",  # Member 1
            "",  # Member 2
            "",  # Member 3
            formula_avg
        ]
        rows_data.append(row)

    # Write data
    end_cell = rowcol_to_a1(len(rows_data), total_cols)
    worksheet.update(f"A1:{end_cell}", rows_data, value_input_option="USER_ENTERED")

    # Formatting
    last_col_letter = _col_letter(total_cols)
    formats = []

    # Row 1 format: Bold, white text, rust/brown background color
    bg_color = {"red": 0.58, "green": 0.27, "blue": 0.21}  # Rust / terracotta

    formats.append({
        "range": f"A1:{last_col_letter}1",
        "format": {
            "textFormat": {"bold": True, "fontSize": 12, "foregroundColor": {"red": 1, "green": 1, "blue": 1}},
            "horizontalAlignment": "CENTER",
            "verticalAlignment": "MIDDLE",
            "backgroundColor": bg_color,
        },
    })

    worksheet.merge_cells(f"A1:{last_col_letter}1", merge_type="MERGE_ALL")

    # Row 2 format: Header
    formats.append({
        "range": f"A2:{last_col_letter}2",
        "format": {
            "textFormat": {"bold": True, "fontSize": 10},
            "horizontalAlignment": "CENTER",
            "verticalAlignment": "MIDDLE",
            "backgroundColor": {"red": 0.9, "green": 0.92, "blue": 0.98},
        },
    })

    # Columns A, D, E, F, G - centered alignment
    formats.append({
        "range": f"A3:A{total_rows}",
        "format": {"horizontalAlignment": "CENTER", "verticalAlignment": "MIDDLE"},
    })
    formats.append({
        "range": f"D3:G{total_rows}",
        "format": {"horizontalAlignment": "CENTER", "verticalAlignment": "MIDDLE"},
    })

    # Columns B, C - left alignment
    formats.append({
        "range": f"B3:C{total_rows}",
        "format": {"horizontalAlignment": "LEFT", "verticalAlignment": "MIDDLE"},
    })

    worksheet.batch_format(formats)
    worksheet.freeze(rows=2, cols=2)

    # Column widths
    requests = [
        {"updateDimensionProperties": {"range": {"sheetId": worksheet.id, "dimension": "COLUMNS", "startIndex": 0, "endIndex": 1}, "properties": {"pixelSize": 40}, "fields": "pixelSize"}},
        {"updateDimensionProperties": {"range": {"sheetId": worksheet.id, "dimension": "COLUMNS", "startIndex": 1, "endIndex": 2}, "properties": {"pixelSize": 280}, "fields": "pixelSize"}},
        {"updateDimensionProperties": {"range": {"sheetId": worksheet.id, "dimension": "COLUMNS", "startIndex": 2, "endIndex": 3}, "properties": {"pixelSize": 300}, "fields": "pixelSize"}},  # Topic
        {"updateDimensionProperties": {"range": {"sheetId": worksheet.id, "dimension": "COLUMNS", "startIndex": 3, "endIndex": 4}, "properties": {"pixelSize": 120}, "fields": "pixelSize"}},  # Member 1
        {"updateDimensionProperties": {"range": {"sheetId": worksheet.id, "dimension": "COLUMNS", "startIndex": 4, "endIndex": 5}, "properties": {"pixelSize": 120}, "fields": "pixelSize"}},  # Member 2
        {"updateDimensionProperties": {"range": {"sheetId": worksheet.id, "dimension": "COLUMNS", "startIndex": 5, "endIndex": 6}, "properties": {"pixelSize": 120}, "fields": "pixelSize"}},  # Member 3
        {"updateDimensionProperties": {"range": {"sheetId": worksheet.id, "dimension": "COLUMNS", "startIndex": 6, "endIndex": 7}, "properties": {"pixelSize": 100}, "fields": "pixelSize"}},  # Average
    ]

    # Borders
    border_style = {"style": "SOLID", "color": {"red": 0.7, "green": 0.7, "blue": 0.7}}
    requests.append({
        "updateBorders": {
            "range": {
                "sheetId": worksheet.id,
                "startRowIndex": 1,
                "endRowIndex": total_rows,
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


def _create_summary_report_sheet(
    spreadsheet: gspread.Spreadsheet,
    students: list[str],
    disciplines: list[dict],
    group_name: str,
) -> gspread.Worksheet:
    """
    Create and format a general semester summary report worksheet ("Загальна звітність").

    Contains student list, final grades for every discipline (linked from 'Семестр'),
    calculated GPA, academic status, and group average summary row at the bottom.
    """
    title = "Загальна звітність"
    total_cols = 2 + len(disciplines) + 2  # №, ПІБ, [disciplines...], Середній бал, Результат
    total_rows = 3 + len(students)  # 2 header rows + student rows + 1 group summary row

    worksheet = spreadsheet.add_worksheet(title=title, rows=max(total_rows, 20), cols=max(total_cols, 6))

    # Pre-calculate the column letter in 'Семестр' sheet for each discipline's final grade
    semester_final_col_letters = []
    sem_col_idx = 3
    for d in disciplines:
        ctype = d.get("control_type", "credit")
        if ctype == "exam":
            # Exam has 3 cols in 'Семестр': Semic 60%, Exam 40%, Total
            final_col = sem_col_idx + 2
            sem_col_idx += 3
        else:
            # Credit and Course Project have 1 col in 'Семестр'
            final_col = sem_col_idx
            sem_col_idx += 1
        semester_final_col_letters.append(_col_letter(final_col))

    rows_data: list[list] = []

    # Row 1: Title
    title_text = f"Загальна звітність успішності за семестр — {group_name}"
    rows_data.append([title_text] + [""] * (total_cols - 1))

    # Row 2: Headers
    headers = ["№", "ПІБ студента"]
    for d in disciplines:
        d_name = d["name"]
        ctype = d.get("control_type", "credit")
        if ctype == "exam":
            headers.append(f"{d_name} (Екз)")
        elif ctype == "course_project":
            headers.append(f"{d_name} (КП)")
        else:
            headers.append(f"{d_name} (Залік)")
    headers.extend(["Середній бал", "Результат"])
    rows_data.append(headers)

    # Row 3 to 2+len(students): Student rows
    disc_start_col = "C"
    disc_end_col = _col_letter(2 + len(disciplines))
    gpa_col = _col_letter(2 + len(disciplines) + 1)
    status_col = _col_letter(2 + len(disciplines) + 2)

    for idx in range(1, len(students) + 1):
        sheet_row = 2 + idx
        student_ref = f"='РК1'!B{sheet_row}"
        row = [idx, student_ref]

        for sem_col_let in semester_final_col_letters:
            row.append(f"='Семестр'!{sem_col_let}{sheet_row}")

        # GPA formula
        formula_gpa = f"=IFERROR(AVERAGE({disc_start_col}{sheet_row}:{disc_end_col}{sheet_row}), \"\")"
        row.append(formula_gpa)

        # Result status formula
        formula_status = f"=IF({disc_start_col}{sheet_row}=\"\", \"\", IF(MIN({disc_start_col}{sheet_row}:{disc_end_col}{sheet_row})>=60, \"Атестовано\", \"Заборгованість\"))"
        row.append(formula_status)

        rows_data.append(row)

    # Bottom summary row (Group Averages)
    last_student_row = 2 + len(students)
    summary_row = ["", "Середній бал групи:"]

    for d_i in range(len(disciplines)):
        c_let = _col_letter(3 + d_i)
        summary_row.append(f"=IFERROR(AVERAGE({c_let}3:{c_let}{last_student_row}), \"\")")

    # Group average GPA
    summary_row.append(f"=IFERROR(AVERAGE({gpa_col}3:{gpa_col}{last_student_row}), \"\")")
    # Group pass count
    summary_row.append(f"=IFERROR(COUNTIF({status_col}3:{status_col}{last_student_row}, \"Атестовано\"), \"\")")

    rows_data.append(summary_row)

    # Write data
    end_cell = rowcol_to_a1(len(rows_data), total_cols)
    worksheet.update(f"A1:{end_cell}", rows_data, value_input_option="USER_ENTERED")

    # Formatting
    last_col_letter = _col_letter(total_cols)
    formats = []

    # Row 1 format: Bold, white text, forest/emerald green
    bg_color = {"red": 0.08, "green": 0.32, "blue": 0.22}  # Forest Emerald

    formats.append({
        "range": f"A1:{last_col_letter}1",
        "format": {
            "textFormat": {"bold": True, "fontSize": 13, "foregroundColor": {"red": 1, "green": 1, "blue": 1}},
            "horizontalAlignment": "CENTER",
            "verticalAlignment": "MIDDLE",
            "backgroundColor": bg_color,
        },
    })

    worksheet.merge_cells(f"A1:{last_col_letter}1", merge_type="MERGE_ALL")

    # Row 2 format: Header
    formats.append({
        "range": f"A2:{last_col_letter}2",
        "format": {
            "textFormat": {"bold": True, "fontSize": 10},
            "horizontalAlignment": "CENTER",
            "verticalAlignment": "MIDDLE",
            "backgroundColor": {"red": 0.9, "green": 0.92, "blue": 0.98},
        },
    })

    # Student name column (B)
    formats.append({
        "range": f"B3:B{total_rows}",
        "format": {
            "horizontalAlignment": "LEFT",
            "verticalAlignment": "MIDDLE",
        },
    })

    # Other columns: centered
    formats.append({
        "range": f"A3:A{total_rows}",
        "format": {"horizontalAlignment": "CENTER", "verticalAlignment": "MIDDLE"},
    })

    if total_cols > 2:
        disc_start = _col_letter(3)
        formats.append({
            "range": f"{disc_start}3:{last_col_letter}{total_rows}",
            "format": {"horizontalAlignment": "CENTER", "verticalAlignment": "MIDDLE"},
        })

    # Summary row format (bottom row): bold text, light mint background
    summary_row_idx = total_rows
    formats.append({
        "range": f"A{summary_row_idx}:{last_col_letter}{summary_row_idx}",
        "format": {
            "textFormat": {"bold": True, "fontSize": 10},
            "backgroundColor": {"red": 0.88, "green": 0.95, "blue": 0.90},
            "horizontalAlignment": "CENTER",
            "verticalAlignment": "MIDDLE",
        },
    })

    worksheet.batch_format(formats)
    worksheet.freeze(rows=2, cols=2)

    # Column widths
    requests = [
        {"updateDimensionProperties": {"range": {"sheetId": worksheet.id, "dimension": "COLUMNS", "startIndex": 0, "endIndex": 1}, "properties": {"pixelSize": 40}, "fields": "pixelSize"}},
        {"updateDimensionProperties": {"range": {"sheetId": worksheet.id, "dimension": "COLUMNS", "startIndex": 1, "endIndex": 2}, "properties": {"pixelSize": 280}, "fields": "pixelSize"}},
    ]

    col_idx = 2
    for _ in disciplines:
        requests.append({
            "updateDimensionProperties": {
                "range": {"sheetId": worksheet.id, "dimension": "COLUMNS", "startIndex": col_idx, "endIndex": col_idx + 1},
                "properties": {"pixelSize": 140},
                "fields": "pixelSize",
            }
        })
        col_idx += 1

    # GPA column width
    requests.append({
        "updateDimensionProperties": {
            "range": {"sheetId": worksheet.id, "dimension": "COLUMNS", "startIndex": col_idx, "endIndex": col_idx + 1},
            "properties": {"pixelSize": 120},
            "fields": "pixelSize",
        }
    })
    col_idx += 1

    # Status column width
    requests.append({
        "updateDimensionProperties": {
            "range": {"sheetId": worksheet.id, "dimension": "COLUMNS", "startIndex": col_idx, "endIndex": col_idx + 1},
            "properties": {"pixelSize": 130},
            "fields": "pixelSize",
        }
    })

    # Borders
    border_style = {"style": "SOLID", "color": {"red": 0.7, "green": 0.7, "blue": 0.7}}
    requests.append({
        "updateBorders": {
            "range": {
                "sheetId": worksheet.id,
                "startRowIndex": 1,
                "endRowIndex": total_rows,
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
    return worksheet




