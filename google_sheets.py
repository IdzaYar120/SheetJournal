"""
Google Sheets integration for Academic Journal creation aligned with Reference.

Authenticates, in order of preference:
  1. Application Default Credentials from `gcloud auth application-default
     login` — uses Google's own pre-verified CLI OAuth client, so there is no
     consent-screen setup, no test-user allowlist, and no "unverified app"
     warning. This is the recommended, low-friction path.
  2. Whatever is found at CREDENTIALS_PATH — a service-account key, or an
     OAuth "Desktop app" Client ID (needed because Google now disables
     service-account key creation by default on new Cloud projects).
"""

import json
import logging
import math
import datetime
from pathlib import Path

import gspread
from gspread.utils import rowcol_to_a1

logger = logging.getLogger(__name__)

CREDENTIALS_PATH = Path(__file__).parent / "credentials.json"
AUTHORIZED_USER_PATH = Path(__file__).parent / "authorized_user.json"

OAUTH_SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]


def _credentials_kind() -> str:
    """Classify CREDENTIALS_PATH as 'service_account', 'oauth_client', 'missing', or 'unknown'."""
    if not CREDENTIALS_PATH.exists():
        return "missing"
    try:
        with open(CREDENTIALS_PATH, "r", encoding="utf-8") as f:
            data = json.load(f)
    except Exception:
        return "unknown"
    if data.get("type") == "service_account":
        return "service_account"
    if "installed" in data or "web" in data:
        return "oauth_client"
    return "unknown"


def _try_application_default_credentials():
    """Try `gcloud`-issued Application Default Credentials, if present."""
    try:
        import google.auth
        creds, _project = google.auth.default(scopes=OAUTH_SCOPES)
        return creds
    except Exception:
        return None


def _get_client() -> gspread.Client:
    """Authenticate and return a gspread client.

    Tries, in order:
      - Application Default Credentials (from `gcloud auth application-default
        login`) — no credentials.json needed at all.
      - A service-account key at CREDENTIALS_PATH — non-interactive.
      - An OAuth Client ID (Desktop app) at CREDENTIALS_PATH — opens a
        one-time browser login and caches the token in AUTHORIZED_USER_PATH.
    """
    adc = _try_application_default_credentials()
    if adc is not None:
        return gspread.authorize(adc)

    kind = _credentials_kind()
    if kind == "missing":
        raise FileNotFoundError(
            "Немає жодного способу авторизації в Google. Або виконайте "
            "'gcloud auth application-default login --scopes=...' (див. INSTRUCTION.md), "
            f"або покладіть credentials.json (ключ сервісного акаунту чи OAuth Client ID) за шляхом: {CREDENTIALS_PATH}"
        )
    if kind == "service_account":
        return gspread.service_account(filename=str(CREDENTIALS_PATH))
    if kind == "oauth_client":
        return gspread.oauth(
            scopes=OAUTH_SCOPES,
            credentials_filename=str(CREDENTIALS_PATH),
            authorized_user_filename=str(AUTHORIZED_USER_PATH),
        )
    raise RuntimeError(
        "Файл credentials.json має нерозпізнаний формат — це має бути або ключ "
        "сервісного акаунту (поле \"type\": \"service_account\"), або OAuth Client ID "
        "(Desktop app, поле \"installed\") у форматі JSON."
    )


def get_drive_credentials():
    """Return valid Google credentials (either kind) for direct REST calls,
    e.g. exporting a Google Doc via the Drive API from app.py."""
    import google.auth.transport.requests

    creds = _get_client().auth
    if not creds.valid:
        creds.refresh(google.auth.transport.requests.Request())
    return creds


def get_service_account_email() -> str | None:
    """Return the service-account email for display, or None when using OAuth
    (where the signed-in Google account itself is the actor, not a robot email)."""
    if _credentials_kind() != "service_account":
        return None
    try:
        with open(CREDENTIALS_PATH, "r", encoding="utf-8") as f:
            return json.load(f).get("client_email")
    except Exception:
        return None


def _col_letter(col_index: int) -> str:
    """Convert a 1-based column index to a spreadsheet column letter (A, B, …, Z, AA, …)."""
    result = ""
    while col_index > 0:
        col_index, remainder = divmod(col_index - 1, 26)
        result = chr(65 + remainder) + result
    return result


def get_short_name(name: str) -> str:
    """Generate a short Cyrillic abbreviation or truncated string for a discipline name."""
    clean = name.replace("Курсовий проект з", "").replace("Курсова робота з", "")
    clean = clean.replace("Курсовий проект", "").replace("Курсова робота", "")
    clean = clean.replace(",", " ").replace("-", " ").strip()
    
    stop_words = {"та", "і", "в", "на", "з", "за", "для", "по", "у", "й", "о"}
    words = [w for w in clean.split() if w.lower() not in stop_words]
    
    if not words:
        return "Дисц"
        
    if len(words) == 1:
        w = words[0]
        return w[:3].capitalize()
    
    acronym = ""
    for w in words:
        if w:
            acronym += w[0].upper()
    return acronym[:5]


def get_first_v_tab(disciplines: list[dict]) -> str:
    """Name of the first В_* report tab, the one other report tabs link names from."""
    for d in disciplines:
        if d.get("control_type") != "course_project":
            return f"В_{get_short_name(d['name'])}"
    return f"В_{get_short_name(disciplines[0]['name'])}_КР"


def get_sk_cols_for_discipline(d: dict, disciplines: list[dict]) -> dict:
    """Determine the column letters for a discipline in the СК worksheet."""
    col = 3
    for x in disciplines:
        if x is d:
            if x.get("control_type") == "course_project":
                return {
                    "def_points": _col_letter(col),
                    "def_pct": _col_letter(col + 1),
                    "cnt_points": _col_letter(col + 2),
                    "cnt_pct": _col_letter(col + 3)
                }
            else:
                return {
                    "points": _col_letter(col),
                    "pct": _col_letter(col + 1)
                }
        if x.get("control_type") == "course_project":
            col += 4
        else:
            col += 2
    return {}


def _populate_pc_sheet(
    worksheet: gspread.Worksheet,
    students: list[str],
    discipline_name: str,
    class_count: int,
    group_name: str,
    control_type: str = "credit",
    teacher_email: str | None = None,
    first_pc_tab_name: str | None = None,
) -> None:
    """Fill and format a single ПК_{ShortName} worksheet."""
    class_count = max(class_count, 1)
    total_cols = 2 + class_count
    total_rows = 4 + len(students) + 1
    
    worksheet.resize(rows=max(total_rows, 20), cols=total_cols)
    rows_data: list[list] = []
    
    # Row 1: Title Card
    label = f"{group_name} — {discipline_name}"
    if control_type == "exam":
        label += " (Екзамен)"
    if teacher_email:
        label += f" (Викладач: {teacher_email})"
    row1 = [label] + [""] * (total_cols - 1)
    rows_data.append(row1)
    
    # Row 2: Date row
    start_date = datetime.date(2026, 2, 2)
    dates = []
    for col_idx in range(class_count):
        d = start_date + datetime.timedelta(days=7 * col_idx)
        dates.append(d.strftime("%Y-%m-%d"))
    row2 = ["", ""] + dates
    rows_data.append(row2)
    
    # Row 3: Class type
    if "фізичн" in discipline_name.lower() or "іноземн" in discipline_name.lower():
        types = ["п"] * class_count
    else:
        types = []
        for i in range(class_count):
            types.append("л" if i % 2 == 0 else "п")
    row3 = ["вид занять", ""] + types
    rows_data.append(row3)
    
    # Row 4: Max points
    target_total = 80.0 if control_type == "credit" else 50.0
    num_practicals = sum(1 for t in types if t == "п")
    if num_practicals > 0:
        val = round(target_total / num_practicals, 1)
    else:
        val = 10.0
    max_points = [val if t == "п" else "" for t in types]
    row4 = ["бали", ""] + max_points
    rows_data.append(row4)
    
    # Student rows (Row 5+)
    for idx, name in enumerate(students):
        row_num = 5 + idx
        no_cell = f"{idx + 1}.0"
        
        # Link name to first pc sheet
        if first_pc_tab_name:
            name_cell = f"='{first_pc_tab_name}'!B{row_num}"
        else:
            name_cell = name
            
        row = [no_cell, name_cell] + [""] * class_count
        rows_data.append(row)
        
    # Serial numbers row at the bottom
    serial_row = ["", ""] + [f"{i + 1}.0" for i in range(class_count)]
    rows_data.append(serial_row)
    
    # Write everything
    end_cell = rowcol_to_a1(len(rows_data), total_cols)
    worksheet.update(f"A1:{end_cell}", rows_data, value_input_option="USER_ENTERED")
    
    # Styling and borders
    last_col_letter = _col_letter(total_cols)
    formats = []
    
    # Row 1 Title Card
    formats.append({
        "range": f"A1:{last_col_letter}1",
        "format": {
            "textFormat": {"bold": True, "fontSize": 12, "foregroundColor": {"red": 1.0, "green": 1.0, "blue": 1.0}},
            "horizontalAlignment": "CENTER",
            "verticalAlignment": "MIDDLE",
            "backgroundColor": {"red": 0.15, "green": 0.27, "blue": 0.49},
        }
    })
    worksheet.merge_cells(f"A1:{last_col_letter}1", merge_type="MERGE_ALL")
    
    # Headers A2:last_col4
    formats.append({
        "range": f"A2:{last_col_letter}4",
        "format": {
            "textFormat": {"bold": True, "fontSize": 9},
            "horizontalAlignment": "CENTER",
            "verticalAlignment": "MIDDLE",
            "backgroundColor": {"red": 0.93, "green": 0.94, "blue": 0.96},
        }
    })
    
    # Student Names B5:B{last_student_row}
    formats.append({
        "range": f"B5:B{4 + len(students)}",
        "format": {
            "horizontalAlignment": "LEFT",
            "verticalAlignment": "MIDDLE",
        }
    })
    
    # Centered indices and grades
    formats.append({
        "range": f"A5:A{4 + len(students)}",
        "format": {"horizontalAlignment": "CENTER", "verticalAlignment": "MIDDLE"},
    })
    if class_count > 0:
        class_start = _col_letter(3)
        formats.append({
            "range": f"{class_start}5:{last_col_letter}{5 + len(students)}",
            "format": {"horizontalAlignment": "CENTER", "verticalAlignment": "MIDDLE"},
        })
        
    worksheet.batch_format(formats)
    worksheet.freeze(rows=4, cols=2)
    
    # Columns dimension sizing
    requests = [
        {"updateDimensionProperties": {"range": {"sheetId": worksheet.id, "dimension": "COLUMNS", "startIndex": 0, "endIndex": 1}, "properties": {"pixelSize": 45}, "fields": "pixelSize"}},
        {"updateDimensionProperties": {"range": {"sheetId": worksheet.id, "dimension": "COLUMNS", "startIndex": 1, "endIndex": 2}, "properties": {"pixelSize": 260}, "fields": "pixelSize"}}
    ]
    if class_count > 0:
        requests.append({
            "updateDimensionProperties": {
                "range": {"sheetId": worksheet.id, "dimension": "COLUMNS", "startIndex": 2, "endIndex": 2 + class_count},
                "properties": {"pixelSize": 35},
                "fields": "pixelSize"
            }
        })
        
    border_style = {"style": "SOLID", "color": {"red": 0.7, "green": 0.7, "blue": 0.7}}
    requests.append({
        "updateBorders": {
            "range": {
                "sheetId": worksheet.id,
                "startRowIndex": 0,
                "endRowIndex": total_rows,
                "startColumnIndex": 0,
                "endColumnIndex": total_cols,
            },
            "top": border_style, "bottom": border_style, "left": border_style, "right": border_style,
            "innerHorizontal": border_style, "innerVertical": border_style,
        }
    })
    worksheet.spreadsheet.batch_update({"requests": requests})


def _populate_kr_sheets(
    ws_content: gspread.Worksheet,
    ws_defense: gspread.Worksheet,
    students: list[str],
    discipline_name: str,
    group_name: str,
    teacher_email: str | None = None,
    first_pc_tab_name: str | None = None,
) -> None:
    """Populate content and defense worksheets for course projects."""
    num_students = len(students)
    total_rows = 2 + 3 * num_students + 2
    
    # Content sheet
    ws_content.resize(rows=max(total_rows, 35), cols=9)
    c_title = ws_content.title
    
    row1_c = [
        f"{group_name} — Курсовий проект з дисципліни: {discipline_name} (Зміст)" + (f" (Викладач: {teacher_email})" if teacher_email else ""),
        "", "", "", "", "", "", "", ""
    ]
    row2_c = [
        "", "Підготовка курсової роботи (Максимальні бали)",
        "5.0", "5.0", "10.0", "10.0", "15.0", "15.0", "=SUM(C2:H2)"
    ]
    rows_c = [row1_c, row2_c]
    
    # Block 1
    for idx, name in enumerate(students):
        row_num = 3 + idx
        name_cell = f"='{first_pc_tab_name}'!B{row_num + 2}" if first_pc_tab_name else name
        rows_c.append([
            f"{idx+1}.0", name_cell, "", "", "", "", "", "", f"=SUM(C{row_num}:H{row_num})"
        ])
        
    # Spacer 1
    rows_c.append(["", "Комісія Член 2 (Максимальні бали)", "5.0", "5.0", "10.0", "10.0", "15.0", "15.0", f"=SUM(C{3+num_students}:H{3+num_students})"])
    
    # Block 2
    for idx in range(num_students):
        row_num = 3 + num_students + 1 + idx
        rows_c.append([
            f"{idx+1}.0", f"=B{3+idx}", "", "", "", "", "", "", f"=SUM(C{row_num}:H{row_num})"
        ])
        
    # Spacer 2
    rows_c.append(["", "Комісія Член 3 (Максимальні бали)", "5.0", "5.0", "10.0", "10.0", "15.0", "15.0", f"=SUM(C{3+2*num_students+1}:H{3+2*num_students+1})"])
    
    # Block 3
    for idx in range(num_students):
        row_num = 3 + 2 * num_students + 2 + idx
        rows_c.append([
            f"{idx+1}.0", f"=B{3+num_students+1+idx}", "", "", "", "", "", "", f"=SUM(C{row_num}:H{row_num})"
        ])
        
    ws_content.update(f"A1:I{len(rows_c)}", rows_c, value_input_option="USER_ENTERED")
    
    # Defense sheet
    ws_defense.resize(rows=max(total_rows, 35), cols=6)
    
    row1_d = [
        f"{group_name} — Курсовий проект з дисципліни: {discipline_name} (Захист)" + (f" (Викладач: {teacher_email})" if teacher_email else ""),
        "", "", "", "", ""
    ]
    row2_d = [
        "", "Захист курсової роботи (Максимальні бали)",
        "15.0", "10.0", "15.0", "=SUM(C2:E2)"
    ]
    rows_d = [row1_d, row2_d]
    
    # Block 1
    for idx, name in enumerate(students):
        row_num = 3 + idx
        rows_d.append([
            f"{idx+1}.0", f"='{c_title}'!B{row_num}", "", "", "", f"=SUM(C{row_num}:E{row_num})"
        ])
        
    # Spacer 1
    rows_d.append(["", "Комісія Член 2 (Максимальні бали)", "15.0", "10.0", "15.0", f"=SUM(C{3+num_students}:E{3+num_students})"])
    
    # Block 2
    for idx in range(num_students):
        row_num = 3 + num_students + 1 + idx
        rows_d.append([
            f"{idx+1}.0", f"=B{3+idx}", "", "", "", f"=SUM(C{row_num}:E{row_num})"
        ])
        
    # Spacer 2
    rows_d.append(["", "Комісія Член 3 (Максимальні бали)", "15.0", "10.0", "15.0", f"=SUM(C{3+2*num_students+1}:E{3+2*num_students+1})"])
    
    # Block 3
    for idx in range(num_students):
        row_num = 3 + 2 * num_students + 2 + idx
        rows_d.append([
            f"{idx+1}.0", f"=B{3+num_students+1+idx}", "", "", "", f"=SUM(C{row_num}:E{row_num})"
        ])
        
    ws_defense.update(f"A1:F{len(rows_d)}", rows_d, value_input_option="USER_ENTERED")
    
    # Shared layout formatting
    for ws, cols_count in [(ws_content, 9), (ws_defense, 6)]:
        last_col = _col_letter(cols_count)
        requests = [
            {
                "updateCells": {
                    "range": {"sheetId": ws.id, "startRowIndex": 0, "endRowIndex": 1, "startColumnIndex": 0, "endColumnIndex": cols_count},
                    "rows": [{
                        "values": [{
                            "userEnteredFormat": {
                                "textFormat": {"bold": True, "fontSize": 12, "foregroundColor": {"red": 1.0, "green": 1.0, "blue": 1.0}},
                                "horizontalAlignment": "CENTER",
                                "verticalAlignment": "MIDDLE",
                                "backgroundColor": {"red": 0.45, "green": 0.15, "blue": 0.20},
                            }
                        } for _ in range(cols_count)]
                    }],
                    "fields": "userEnteredFormat(textFormat,horizontalAlignment,verticalAlignment,backgroundColor)"
                }
            },
            {
                "updateCells": {
                    "range": {"sheetId": ws.id, "startRowIndex": 1, "endRowIndex": 2, "startColumnIndex": 0, "endColumnIndex": cols_count},
                    "rows": [{
                        "values": [{
                            "userEnteredFormat": {
                                "textFormat": {"bold": True, "fontSize": 9},
                                "horizontalAlignment": "CENTER",
                                "verticalAlignment": "MIDDLE",
                                "backgroundColor": {"red": 0.95, "green": 0.9, "blue": 0.9},
                            }
                        } for _ in range(cols_count)]
                    }],
                    "fields": "userEnteredFormat(textFormat,horizontalAlignment,verticalAlignment,backgroundColor)"
                }
            },
            {
                "updateCells": {
                    "range": {"sheetId": ws.id, "startRowIndex": 2 + num_students, "endRowIndex": 2 + num_students + 1, "startColumnIndex": 0, "endColumnIndex": cols_count},
                    "rows": [{
                        "values": [{
                            "userEnteredFormat": {
                                "textFormat": {"bold": True, "fontSize": 9},
                                "horizontalAlignment": "CENTER",
                                "verticalAlignment": "MIDDLE",
                                "backgroundColor": {"red": 0.95, "green": 0.9, "blue": 0.9},
                            }
                        } for _ in range(cols_count)]
                    }],
                    "fields": "userEnteredFormat(textFormat,horizontalAlignment,verticalAlignment,backgroundColor)"
                }
            },
            {
                "updateCells": {
                    "range": {"sheetId": ws.id, "startRowIndex": 2 + 2 * num_students + 1, "endRowIndex": 2 + 2 * num_students + 2, "startColumnIndex": 0, "endColumnIndex": cols_count},
                    "rows": [{
                        "values": [{
                            "userEnteredFormat": {
                                "textFormat": {"bold": True, "fontSize": 9},
                                "horizontalAlignment": "CENTER",
                                "verticalAlignment": "MIDDLE",
                                "backgroundColor": {"red": 0.95, "green": 0.9, "blue": 0.9},
                            }
                        } for _ in range(cols_count)]
                    }],
                    "fields": "userEnteredFormat(textFormat,horizontalAlignment,verticalAlignment,backgroundColor)"
                }
            },
            {"updateDimensionProperties": {"range": {"sheetId": ws.id, "dimension": "COLUMNS", "startIndex": 0, "endIndex": 1}, "properties": {"pixelSize": 45}, "fields": "pixelSize"}},
            {"updateDimensionProperties": {"range": {"sheetId": ws.id, "dimension": "COLUMNS", "startIndex": 1, "endIndex": 2}, "properties": {"pixelSize": 260}, "fields": "pixelSize"}}
        ]
        
        ws.merge_cells(f"A1:{last_col}1", merge_type="MERGE_ALL")
        border_style = {"style": "SOLID", "color": {"red": 0.7, "green": 0.7, "blue": 0.7}}
        requests.append({
            "updateBorders": {
                "range": {
                    "sheetId": ws.id,
                    "startRowIndex": 0,
                    "endRowIndex": total_rows,
                    "startColumnIndex": 0,
                    "endColumnIndex": cols_count,
                },
                "top": border_style, "bottom": border_style, "left": border_style, "right": border_style,
                "innerHorizontal": border_style, "innerVertical": border_style,
            }
        })
        ws.spreadsheet.batch_update({"requests": requests})
        ws.freeze(rows=2, cols=2)


def _create_milestone_sheet(
    spreadsheet: gspread.Spreadsheet,
    students: list[str],
    disciplines: list[dict],
    milestone: int,
    group_name: str,
    first_pc_tab_name: str | None = None,
) -> gspread.Worksheet:
    """Create and format РК1/РК2 sheets."""
    title = f"РК{milestone}"
    milestone_disciplines = [d for d in disciplines if d.get("control_type", "credit") != "course_project"]
    if not milestone_disciplines:
        milestone_disciplines = [{"name": group_name, "class_count": 0, "control_type": "credit"}]
        
    num_students = len(students)
    total_cols = 2 + len(milestone_disciplines) * 3
    total_rows = 2 + num_students
    
    worksheet = spreadsheet.add_worksheet(title=title, rows=max(total_rows, 20), cols=max(total_cols, 8))
    rows_data: list[list] = []
    
    # Row 1: Header names of subjects
    row1 = ["№ з/п", "ПІП"]
    for d in milestone_disciplines:
        short = get_short_name(d["name"])
        pc_tab = f"ПК_{short}"
        row1.extend([f"='{pc_tab}'!C1", "", ""])
    rows_data.append(row1)
    
    # Row 2: Max points sum, %, ECTS/National label
    row2 = ["", ""]
    for d in milestone_disciplines:
        short = get_short_name(d["name"])
        pc_tab = f"ПК_{short}"
        class_count = max(d.get("class_count") or 0, 1)
        middle = math.ceil(class_count / 2)
        
        if milestone == 1:
            start_col = "C"
            end_col = _col_letter(2 + middle)
        else:
            start_col = _col_letter(3 + middle)
            end_col = _col_letter(2 + class_count)
            
        sum_formula = f"=SUM('{pc_tab}'!{start_col}4:{end_col}4)"
        row2.extend([sum_formula, "%", "Національна / ECTS"])
    rows_data.append(row2)
    
    # Row 3+: Student rows
    for idx in range(num_students):
        row_num = 3 + idx
        student_no = f"{idx+1}.0"
        name_cell = f"='{first_pc_tab_name}'!B{row_num + 2}" if first_pc_tab_name else students[idx]
        row = [student_no, name_cell]
        
        col_idx = 3
        for d in milestone_disciplines:
            short = get_short_name(d["name"])
            pc_tab = f"ПК_{short}"
            class_count = max(d.get("class_count") or 0, 1)
            middle = math.ceil(class_count / 2)
            
            if milestone == 1:
                start_col = "C"
                end_col = _col_letter(2 + middle)
            else:
                start_col = _col_letter(3 + middle)
                end_col = _col_letter(2 + class_count)
                
            pts_letter = _col_letter(col_idx)
            pct_letter = _col_letter(col_idx + 1)
            
            pts_formula = f"=SUM('{pc_tab}'!{start_col}{row_num+2}:{end_col}{row_num+2})"
            pct_formula = f"={pts_letter}{row_num}/{pts_letter}$2*100"
            ects_formula = (
                f'=IF(AND({pct_letter}{row_num}<=100,88<={pct_letter}{row_num}),"5A",'
                f'IF(AND({pct_letter}{row_num}<88,81<={pct_letter}{row_num}),"4B",'
                f'IF(AND({pct_letter}{row_num}<81,74<={pct_letter}{row_num}),"4C",'
                f'IF(AND({pct_letter}{row_num}<74,67<={pct_letter}{row_num}),"3D",'
                f'IF(AND({pct_letter}{row_num}<67,60<={pct_letter}{row_num}),"3E",'
                f'IF(AND({pct_letter}{row_num}<60,35<={pct_letter}{row_num}),"2FX",'
                f'IF(AND({pct_letter}{row_num}<35,0<={pct_letter}{row_num}),"2F","")))))))'
            )
            row.extend([pts_formula, pct_formula, ects_formula])
            col_idx += 3
            
        rows_data.append(row)
        
    end_cell = rowcol_to_a1(len(rows_data), total_cols)
    worksheet.update(f"A1:{end_cell}", rows_data, value_input_option="USER_ENTERED")
    
    # Styling and widths
    requests = [
        {
            "updateCells": {
                "range": {"sheetId": worksheet.id, "startRowIndex": 0, "endRowIndex": 2, "startColumnIndex": 0, "endColumnIndex": total_cols},
                "rows": [{
                    "values": [{
                        "userEnteredFormat": {
                            "textFormat": {"bold": True, "fontSize": 9},
                            "horizontalAlignment": "CENTER",
                            "verticalAlignment": "MIDDLE",
                            "backgroundColor": {"red": 0.88, "green": 0.92, "blue": 0.95},
                        }
                    } for _ in range(total_cols)]
                } for _ in range(2)],
                "fields": "userEnteredFormat(textFormat,horizontalAlignment,verticalAlignment,backgroundColor)"
            }
        },
        {"updateDimensionProperties": {"range": {"sheetId": worksheet.id, "dimension": "COLUMNS", "startIndex": 0, "endIndex": 1}, "properties": {"pixelSize": 45}, "fields": "pixelSize"}},
        {"updateDimensionProperties": {"range": {"sheetId": worksheet.id, "dimension": "COLUMNS", "startIndex": 1, "endIndex": 2}, "properties": {"pixelSize": 260}, "fields": "pixelSize"}}
    ]
    
    for i in range(len(milestone_disciplines)):
        sub_col = 3 + i * 3
        start_letter = _col_letter(sub_col)
        end_letter = _col_letter(sub_col + 2)
        worksheet.merge_cells(f"{start_letter}1:{end_letter}1", merge_type="MERGE_ALL")
        
    border_style = {"style": "SOLID", "color": {"red": 0.7, "green": 0.7, "blue": 0.7}}
    requests.append({
        "updateBorders": {
            "range": {
                "sheetId": worksheet.id,
                "startRowIndex": 0,
                "endRowIndex": total_rows,
                "startColumnIndex": 0,
                "endColumnIndex": total_cols,
            },
            "top": border_style, "bottom": border_style, "left": border_style, "right": border_style,
            "innerHorizontal": border_style, "innerVertical": border_style,
        }
    })
    worksheet.spreadsheet.batch_update({"requests": requests})
    worksheet.freeze(rows=2, cols=2)
    return worksheet


def _create_helper_sheet(
    spreadsheet: gspread.Spreadsheet,
    students: list[str],
    disciplines: list[dict],
    title: str,
    ref_sheet: str | None = None,
    first_pc_tab_name: str | None = None,
) -> gspread.Worksheet:
    """Create helper sheets like ДОД_РК1, ДОД_РК2, ІНДЗ, ДОД_ІНДЗ."""
    milestone_disciplines = [d for d in disciplines if d.get("control_type", "credit") != "course_project"]
    if not milestone_disciplines:
        milestone_disciplines = [{"name": spreadsheet.title, "class_count": 0, "control_type": "credit"}]
        
    num_students = len(students)
    total_cols = 2 + len(milestone_disciplines) * 2
    total_rows = 2 + num_students
    
    worksheet = spreadsheet.add_worksheet(title=title, rows=max(total_rows, 20), cols=max(total_cols, 6))
    rows_data: list[list] = []
    
    # Row 1: Header names of subjects
    row1 = ["№ з/п", "ПІП"]
    for d in milestone_disciplines:
        short = get_short_name(d["name"])
        pc_tab = f"ПК_{short}"
        row1.extend([f"='{pc_tab}'!C1", ""])
    rows_data.append(row1)
    
    # Row 2: Labels
    row2 = ["", ""]
    for i, d in enumerate(milestone_disciplines):
        max_val = "10.0" if title == "ІНДЗ" else ""
        row2.extend([max_val, "%"])
    rows_data.append(row2)
    
    # Row 3+: Student rows
    for idx in range(num_students):
        row_num = 3 + idx
        student_no = f"{idx+1}.0"
        name_cell = f"='{first_pc_tab_name}'!B{row_num + 2}" if first_pc_tab_name else students[idx]
        row = [student_no, name_cell]
        
        for i, d in enumerate(milestone_disciplines):
            col_idx = 3 + i * 2
            pts_letter = _col_letter(col_idx)
            
            if title == "ІНДЗ":
                pct_formula = f"={pts_letter}{row_num}/${pts_letter}$2*100"
            elif title == "ДОД_ІНДЗ":
                pct_formula = f"={pts_letter}{row_num}/'ІНДЗ'!${pts_letter}$2*100"
            else:
                ref_pts_col = _col_letter(3 + i * 3)
                pct_formula = f"={pts_letter}{row_num}/'{ref_sheet}'!${ref_pts_col}$2*100"
                
            row.extend(["", pct_formula])
        rows_data.append(row)
        
    end_cell = rowcol_to_a1(len(rows_data), total_cols)
    worksheet.update(f"A1:{end_cell}", rows_data, value_input_option="USER_ENTERED")
    
    # Styling and border requests
    requests = [
        {
            "updateCells": {
                "range": {"sheetId": worksheet.id, "startRowIndex": 0, "endRowIndex": 2, "startColumnIndex": 0, "endColumnIndex": total_cols},
                "rows": [{
                    "values": [{
                        "userEnteredFormat": {
                            "textFormat": {"bold": True, "fontSize": 9},
                            "horizontalAlignment": "CENTER",
                            "verticalAlignment": "MIDDLE",
                            "backgroundColor": {"red": 0.9, "green": 0.94, "blue": 0.92},
                        }
                    } for _ in range(total_cols)]
                } for _ in range(2)],
                "fields": "userEnteredFormat(textFormat,horizontalAlignment,verticalAlignment,backgroundColor)"
            }
        },
        {"updateDimensionProperties": {"range": {"sheetId": worksheet.id, "dimension": "COLUMNS", "startIndex": 0, "endIndex": 1}, "properties": {"pixelSize": 45}, "fields": "pixelSize"}},
        {"updateDimensionProperties": {"range": {"sheetId": worksheet.id, "dimension": "COLUMNS", "startIndex": 1, "endIndex": 2}, "properties": {"pixelSize": 260}, "fields": "pixelSize"}}
    ]
    
    for i in range(len(milestone_disciplines)):
        sub_col = 3 + i * 2
        start_letter = _col_letter(sub_col)
        end_letter = _col_letter(sub_col + 1)
        worksheet.merge_cells(f"{start_letter}1:{end_letter}1", merge_type="MERGE_ALL")
        
    border_style = {"style": "SOLID", "color": {"red": 0.7, "green": 0.7, "blue": 0.7}}
    requests.append({
        "updateBorders": {
            "range": {
                "sheetId": worksheet.id,
                "startRowIndex": 0,
                "endRowIndex": total_rows,
                "startColumnIndex": 0,
                "endColumnIndex": total_cols,
            },
            "top": border_style, "bottom": border_style, "left": border_style, "right": border_style,
            "innerHorizontal": border_style, "innerVertical": border_style,
        }
    })
    
    worksheet.spreadsheet.batch_update({"requests": requests})
    worksheet.freeze(rows=2, cols=2)
    return worksheet


def _create_sk_sheet(
    spreadsheet: gspread.Spreadsheet,
    students: list[str],
    disciplines: list[dict],
    group_name: str,
    first_pc_tab_name: str | None = None,
) -> gspread.Worksheet:
    """Create the Semester Control worksheet ("СК")."""
    num_students = len(students)
    
    total_cols = 2
    for d in disciplines:
        if d.get("control_type", "credit") == "course_project":
            total_cols += 4
        else:
            total_cols += 2
            
    worksheet = spreadsheet.add_worksheet(title="СК", rows=max(2 + num_students, 20), cols=max(total_cols, 8))
    rows_data: list[list] = []
    
    # Row 1: Header names of subjects
    row1 = ["№ з/п", "ПІП"]
    for d in disciplines:
        short = get_short_name(d["name"])
        if d.get("control_type", "credit") == "course_project":
            row1.extend([f"{d['name']} (курсова робота)", "", "", ""])
        else:
            pc_tab = f"ПК_{short}"
            row1.extend([f"='{pc_tab}'!C1", ""])
    rows_data.append(row1)
    
    # Row 2: Max points formula or label, and %
    row2 = ["", ""]
    milestone_disciplines = [d for d in disciplines if d.get("control_type", "credit") != "course_project"]
    
    for d in disciplines:
        if d.get("control_type", "credit") == "course_project":
            row2.extend(["40.0", "%", "60.0", "%"])
        else:
            m_idx = milestone_disciplines.index(d)
            rk_col = _col_letter(3 + m_idx * 3)
            dod_col = _col_letter(3 + m_idx * 2)
            indz_col = _col_letter(3 + m_idx * 2)
            
            max_formula = (
                f"='РК1'!{rk_col}2+'РК2'!{rk_col}2+"
                f"'ДОД_РК1'!{dod_col}2+'ДОД_РК2'!{dod_col}2+"
                f"'ІНДЗ'!{indz_col}2+'ДОД_ІНДЗ'!{indz_col}2"
            )
            row2.extend([max_formula, "%"])
    rows_data.append(row2)
    
    # Row 3+: Student rows
    for idx in range(num_students):
        row_num = 3 + idx
        student_no = f"{idx+1}.0"
        name_cell = f"='{first_pc_tab_name}'!B{row_num + 2}" if first_pc_tab_name else students[idx]
        row = [student_no, name_cell]
        
        for d in disciplines:
            if d.get("control_type", "credit") == "course_project":
                short = get_short_name(d["name"])
                r1 = row_num
                r2 = row_num + num_students + 1
                r3 = row_num + 2 * num_students + 2
                
                def_formula = f"=AVERAGE('КР_{short}_Захист'!F{r1},'КР_{short}_Захист'!F{r2},'КР_{short}_Захист'!F{r3})"
                def_pct = f"={_col_letter(len(row)+1)}{row_num}/40.0*100"
                cnt_formula = f"=AVERAGE('КР_{short}_Зміст'!I{r1},'КР_{short}_Зміст'!I{r2},'КР_{short}_Зміст'!I{r3})"
                cnt_pct = f"={_col_letter(len(row)+3)}{row_num}/60.0*100"
                row.extend([def_formula, def_pct, cnt_formula, cnt_pct])
            else:
                m_idx = milestone_disciplines.index(d)
                rk_col = _col_letter(3 + m_idx * 3)
                dod_col = _col_letter(3 + m_idx * 2)
                indz_col = _col_letter(3 + m_idx * 2)
                
                pts_formula = (
                    f"='РК1'!{rk_col}{row_num}+'ДОД_РК1'!{dod_col}{row_num}+"
                    f"'РК2'!{rk_col}{row_num}+'ДОД_РК2'!{dod_col}{row_num}+"
                    f"'ІНДЗ'!{indz_col}{row_num}+'ДОД_ІНДЗ'!{indz_col}{row_num}"
                )
                pts_letter = _col_letter(len(row) + 1)
                pct_formula = f"={pts_letter}{row_num}/{pts_letter}$2*100"
                row.extend([pts_formula, pct_formula])
                
        rows_data.append(row)
        
    end_cell = rowcol_to_a1(len(rows_data), total_cols)
    worksheet.update(f"A1:{end_cell}", rows_data, value_input_option="USER_ENTERED")
    
    # Styling and borders
    requests = [
        {
            "updateCells": {
                "range": {"sheetId": worksheet.id, "startRowIndex": 0, "endRowIndex": 2, "startColumnIndex": 0, "endColumnIndex": total_cols},
                "rows": [{
                    "values": [{
                        "userEnteredFormat": {
                            "textFormat": {"bold": True, "fontSize": 9},
                            "horizontalAlignment": "CENTER",
                            "verticalAlignment": "MIDDLE",
                            "backgroundColor": {"red": 0.85, "green": 0.9, "blue": 0.88},
                        }
                    } for _ in range(total_cols)]
                } for _ in range(2)],
                "fields": "userEnteredFormat(textFormat,horizontalAlignment,verticalAlignment,backgroundColor)"
            }
        },
        {"updateDimensionProperties": {"range": {"sheetId": worksheet.id, "dimension": "COLUMNS", "startIndex": 0, "endIndex": 1}, "properties": {"pixelSize": 45}, "fields": "pixelSize"}},
        {"updateDimensionProperties": {"range": {"sheetId": worksheet.id, "dimension": "COLUMNS", "startIndex": 1, "endIndex": 2}, "properties": {"pixelSize": 260}, "fields": "pixelSize"}}
    ]
    
    curr_col = 3
    for d in disciplines:
        if d.get("control_type", "credit") == "course_project":
            start_letter = _col_letter(curr_col)
            end_letter = _col_letter(curr_col + 3)
            worksheet.merge_cells(f"{start_letter}1:{end_letter}1", merge_type="MERGE_ALL")
            curr_col += 4
        else:
            start_letter = _col_letter(curr_col)
            end_letter = _col_letter(curr_col + 1)
            worksheet.merge_cells(f"{start_letter}1:{end_letter}1", merge_type="MERGE_ALL")
            curr_col += 2
            
    border_style = {"style": "SOLID", "color": {"red": 0.7, "green": 0.7, "blue": 0.7}}
    requests.append({
        "updateBorders": {
            "range": {
                "sheetId": worksheet.id,
                "startRowIndex": 0,
                "endRowIndex": 2 + num_students,
                "startColumnIndex": 0,
                "endColumnIndex": total_cols,
            },
            "top": border_style, "bottom": border_style, "left": border_style, "right": border_style,
            "innerHorizontal": border_style, "innerVertical": border_style,
        }
    })
    
    worksheet.spreadsheet.batch_update({"requests": requests})
    worksheet.freeze(rows=2, cols=2)
    return worksheet


def _create_v_report_sheet(
    worksheet: gspread.Worksheet,
    students: list[str],
    discipline: dict,
    disciplines: list[dict],
    group_name: str,
    first_pc_tab_name: str | None = None,
) -> gspread.Worksheet:
    """Create a single official report sheet "В_{ShortName}" for a discipline."""
    short = get_short_name(discipline["name"])
    title = worksheet.title
    control_type = discipline.get("control_type", "credit")
    
    num_students = len(students)
    total_cols = 11
    total_rows = 13 + num_students + 15
    
    worksheet.resize(rows=max(total_rows, 40), cols=total_cols)
    rows_data: list[list] = []
    
    # Metadata Header rows
    rows_data.append(["Університет економіки і підприємництва"] + [""] * 10)
    rows_data.append([""] * 11)
    
    first_v_tab = get_first_v_tab(disciplines)
    rows_data.append(["2025 / 2026 навчальний рік", "", "", "", "", "", "", "Група", f"='{first_v_tab}'!I3" if first_pc_tab_name and title != first_v_tab else group_name, "", ""])
    rows_data.append(["", "", "", "", "", "", "", "Курс", f"='{first_v_tab}'!I4" if first_pc_tab_name and title != first_v_tab else "I", "", ""])
    rows_data.append(["Заліково-екзаменаційна відомість № ", "", "", "", "", "", "", "", "", "", ""])
    
    pc_tab = f"ПК_{short}"
    rows_data.append(["з дисципліни", "", f"='{pc_tab}'!C1" if control_type != "course_project" else f"='КР_{short}_Зміст'!B1", "", "", "", "", "", "", "", ""])
    rows_data.append([""] * 11)
    rows_data.append(["за 2 навчальний семестр", "", "", "", "", "", "", "Екзамен / Залік", "", "", ""])
    rows_data.append(["Екзаменатор", "", "", "", "", "", "", "", "", "", ""])
    rows_data.append(["Викладачі, які ведуть практичні, семінарські і лабораторні роботи", "", "", "", "", "", "", "", "", "", ""])
    rows_data.append([""] * 11)
    
    # Table headers
    if control_type == "course_project":
        header_row12 = ["№ з/п", "Прізвище, ініціали студента", "Номер індивід. навч. плану", "Оцінка курсової роботи", "", "", "", "", "", "Підпис керівника та членів комісії", ""]
        header_row13 = ["", "", "", "Національна", "", "ECTS", "", "Бали", "", "", ""]
    else:
        header_row12 = ["№ з/п", "Прізвище, ініціали студента", "Номер індивід. навч. плану", "Залікова оцінка", "", "", "Підпис   викла- дача", "Екзаменаційна оцінка", "", "", "Підпис екзаме- натора"]
        header_row13 = ["", "", "", "Національна", "ECTS", "Бали", "", "Національна", "ECTS", "Бали", ""]
        
    rows_data.append(header_row12)
    rows_data.append(header_row13)
    
    # Get column letters from СК for this discipline
    sk_cols = get_sk_cols_for_discipline(discipline, disciplines)
    
    # Student rows (Row 14+)
    for idx in range(num_students):
        row_num = 14 + idx
        sk_row = 3 + idx
        
        first_report_sheet = get_first_v_tab(disciplines)
        if title == first_report_sheet:
            name_formula = (
                f'=IFERROR(CONCATENATE(LEFT(\'СК\'!B{sk_row},(FIND(" ",\'СК\'!B{sk_row})-1))," ",'
                f'MID(\'СК\'!B{sk_row},FIND(" ",\'СК\'!B{sk_row})+1,1),". ",'
                f'MID(\'СК\'!B{sk_row},FIND(" ",\'СК\'!B{sk_row},FIND(" ",\'СК\'!B{sk_row})+1)+1,1),"."),\'СК\'!B{sk_row})'
            )
        else:
            name_formula = f"='{first_report_sheet}'!B{row_num}"
            
        no_cell = f"='{first_report_sheet}'!A{row_num}" if title != first_report_sheet else f"{idx+1}.0"
        record_book_no = f"='{first_report_sheet}'!C{row_num}" if title != first_report_sheet else f"{4000+idx+1}"
        
        if control_type == "course_project":
            pts_formula = f"='СК'!{sk_cols['def_points']}{sk_row}+'СК'!{sk_cols['cnt_points']}{sk_row}"
            nat_formula = (
                f'=IF(AND(H{row_num}<=100,88<=H{row_num}),"Відмінно",'
                f'IF(AND(H{row_num}<88,81<=H{row_num}),"Добре",'
                f'IF(AND(H{row_num}<81,74<=H{row_num}),"Добре",'
                f'IF(AND(H{row_num}<74,67<=H{row_num}),"Задовільно",'
                f'IF(AND(H{row_num}<67,60<=H{row_num}),"Задовільно",'
                f'IF(AND(H{row_num}<60,35<=H{row_num}),"Незадовільно",'
                f'IF(AND(H{row_num}<35,0<=H{row_num}),"Незадовільно","")))))))'
            )
            ects_formula = (
                f'=IF(AND(H{row_num}<=100,88<=H{row_num}),"A",'
                f'IF(AND(H{row_num}<88,81<=H{row_num}),"B",'
                f'IF(AND(H{row_num}<81,74<=H{row_num}),"C",'
                f'IF(AND(H{row_num}<74,67<=H{row_num}),"D",'
                f'IF(AND(H{row_num}<67,60<=H{row_num}),"E",'
                f'IF(AND(H{row_num}<60,35<=H{row_num}),"FX",'
                f'IF(AND(H{row_num}<35,0<=H{row_num}),"F","")))))))'
            )
            student_row = [no_cell, name_formula, record_book_no, nat_formula, "", ects_formula, "", pts_formula, "", "", ""]
            
        elif control_type == "exam":
            pts_formula = f"='СК'!{sk_cols['def_points']}{sk_row}+'СК'!{sk_cols['cnt_points']}{sk_row}"
            nat_formula = (
                f'=IF(AND(J{row_num}<=100,88<=J{row_num}),"Відмінно",'
                f'IF(AND(J{row_num}<88,81<=J{row_num}),"Добре",'
                f'IF(AND(J{row_num}<81,74<=J{row_num}),"Добре",'
                f'IF(AND(J{row_num}<74,67<=J{row_num}),"Задовільно",'
                f'IF(AND(J{row_num}<67,60<=J{row_num}),"Задовільно",'
                f'IF(AND(J{row_num}<60,35<=J{row_num}),"Незадовільно",'
                f'IF(AND(J{row_num}<35,0<=J{row_num}),"Незадовільно","")))))))'
            )
            ects_formula = (
                f'=IF(AND(J{row_num}<=100,88<=J{row_num}),"A",'
                f'IF(AND(J{row_num}<88,81<=J{row_num}),"B",'
                f'IF(AND(J{row_num}<81,74<=J{row_num}),"C",'
                f'IF(AND(J{row_num}<74,67<=J{row_num}),"D",'
                f'IF(AND(J{row_num}<67,60<=J{row_num}),"E",'
                f'IF(AND(J{row_num}<60,35<=J{row_num}),"FX",'
                f'IF(AND(J{row_num}<35,0<=J{row_num}),"F","")))))))'
            )
            student_row = [no_cell, name_formula, record_book_no, "", "", "", "", nat_formula, ects_formula, pts_formula, ""]
            
        else:  # credit
            pts_formula = f"='СК'!{sk_cols['points']}{sk_row}"
            nat_formula = (
                f'=IF(AND(F{row_num}<=100,60<=F{row_num}),"Зараховано",'
                f'IF(AND(F{row_num}<60,35<=F{row_num}),"Не зараховано",'
                f'IF(AND(F{row_num}<35,0<=F{row_num}),"Не зараховано","")))'
            )
            ects_formula = (
                f'=IF(AND(F{row_num}<=100,88<=F{row_num}),"A",'
                f'IF(AND(F{row_num}<88,81<=F{row_num}),"B",'
                f'IF(AND(F{row_num}<81,74<=F{row_num}),"C",'
                f'IF(AND(F{row_num}<74,67<=F{row_num}),"D",'
                f'IF(AND(F{row_num}<67,60<=F{row_num}),"E",'
                f'IF(AND(F{row_num}<60,35<=F{row_num}),"FX",'
                f'IF(AND(F{row_num}<35,0<=F{row_num}),"F","")))))))'
            )
            student_row = [no_cell, name_formula, record_book_no, nat_formula, ects_formula, pts_formula, "", "", "", "", ""]
            
        rows_data.append(student_row)
        
    # Bottom Stats
    last_student_row = 13 + num_students
    rows_data.append([""] * 11)
    
    r_grp_size = last_student_row + 2
    r_present = last_student_row + 3
    r_excellent = last_student_row + 4
    
    ects_col = "F" if control_type == "course_project" else ("I" if control_type == "exam" else "E")
    
    rows_data.append(["", "Студентів в групі", f"=A{last_student_row}", "", "", "", "Не з’явилось", "", "", f"=C{r_grp_size}-J{r_present}", ""])
    rows_data.append(["", "", "", "", "", "", "Було присутніх", "", "", f"=SUM(J{r_excellent}:J{r_excellent+6})", ""])
    rows_data.append(["", "Не допущено", "0.0", "", "", "", "Відмінно / А", "", "", f'=COUNTIF(${ects_col}$14:${ects_col}${last_student_row},"A")', ""])
    rows_data.append(["", "", "", "", "", "", "Добре / B", "", "", f'=COUNTIF(${ects_col}$14:${ects_col}${last_student_row},"B")', ""])
    rows_data.append(["", "", "", "", "", "", "Добре / C", "", "", f'=COUNTIF(${ects_col}$14:${ects_col}${last_student_row},"C")', ""])
    rows_data.append(["", "", "", "", "", "", "Задовільно / D", "", "", f'=COUNTIF(${ects_col}$14:${ects_col}${last_student_row},"D")', ""])
    rows_data.append(["", "", "", "", "", "", "Задовільно / E", "", "", f'=COUNTIF(${ects_col}$14:${ects_col}${last_student_row},"E")', ""])
    rows_data.append(["", "", "", "", "", "", "Незадовільно / FX", "", "", f'=COUNTIF(${ects_col}$14:${ects_col}${last_student_row},"FX")', ""])
    rows_data.append(["", "", "", "", "", "", "Незадовільно / F", "", "", f'=COUNTIF(${ects_col}$14:${ects_col}${last_student_row},"F")', ""])
    
    end_cell = rowcol_to_a1(len(rows_data), total_cols)
    worksheet.update(f"A1:{end_cell}", rows_data, value_input_option="USER_ENTERED")
    
    # Formatting
    requests = [
        {
            "updateCells": {
                "range": {"sheetId": worksheet.id, "startRowIndex": 11, "endRowIndex": 13, "startColumnIndex": 0, "endColumnIndex": total_cols},
                "rows": [{
                    "values": [{
                        "userEnteredFormat": {
                            "textFormat": {"bold": True, "fontSize": 9},
                            "horizontalAlignment": "CENTER",
                            "verticalAlignment": "MIDDLE",
                            "backgroundColor": {"red": 0.94, "green": 0.94, "blue": 0.94},
                        }
                    } for _ in range(total_cols)]
                } for _ in range(2)],
                "fields": "userEnteredFormat(textFormat,horizontalAlignment,verticalAlignment,backgroundColor)"
            }
        },
        {"updateDimensionProperties": {"range": {"sheetId": worksheet.id, "dimension": "COLUMNS", "startIndex": 0, "endIndex": 1}, "properties": {"pixelSize": 45}, "fields": "pixelSize"}},
        {"updateDimensionProperties": {"range": {"sheetId": worksheet.id, "dimension": "COLUMNS", "startIndex": 1, "endIndex": 2}, "properties": {"pixelSize": 240}, "fields": "pixelSize"}},
        {"updateDimensionProperties": {"range": {"sheetId": worksheet.id, "dimension": "COLUMNS", "startIndex": 2, "endIndex": 3}, "properties": {"pixelSize": 100}, "fields": "pixelSize"}},
    ]
    
    if control_type == "course_project":
        worksheet.merge_cells("D12:E12", merge_type="MERGE_ALL")
        worksheet.merge_cells("F12:G12", merge_type="MERGE_ALL")
        worksheet.merge_cells("H12:I12", merge_type="MERGE_ALL")
    else:
        worksheet.merge_cells("D12:F12", merge_type="MERGE_ALL")
        worksheet.merge_cells("H12:J12", merge_type="MERGE_ALL")
        
    border_style = {"style": "SOLID", "color": {"red": 0.7, "green": 0.7, "blue": 0.7}}
    requests.append({
        "updateBorders": {
            "range": {
                "sheetId": worksheet.id,
                "startRowIndex": 11,
                "endRowIndex": last_student_row,
                "startColumnIndex": 0,
                "endColumnIndex": total_cols,
            },
            "top": border_style, "bottom": border_style, "left": border_style, "right": border_style,
            "innerHorizontal": border_style, "innerVertical": border_style,
        }
    })
    worksheet.spreadsheet.batch_update({"requests": requests})
    worksheet.freeze(rows=13, cols=2)
    return worksheet


def _create_consolidated_dashboard(
    worksheet: gspread.Worksheet,
    students: list[str],
    disciplines: list[dict],
    group_name: str,
    first_pc_tab_name: str | None = None,
) -> gspread.Worksheet:
    """Create the final consolidated dashboard dashboard "Відомості"."""
    num_students = len(students)
    total_cols = 3 + len(disciplines) * 3
    total_rows = 10 + num_students
    
    worksheet.resize(rows=max(total_rows, 30), cols=total_cols)
    rows_data: list[list] = []
    
    # Metadata Header rows
    rows_data.append(["Університет економіки і підприємництва"] + [""] * (total_cols - 1))
    rows_data.append([""] * total_cols)
    rows_data.append(["", "", "", "", "", "", "", "ХІД"] + [""] * (total_cols - 8))
    rows_data.append(["", "складання заліків та екзаменів сесії", "", "", "", "", "", ""] + [""] * (total_cols - 8))
    rows_data.append(["", "2025 / 2026 навчальний рік", "", "", "", "", "", ""] + [""] * (total_cols - 8))
    
    first_report_sheet = get_first_v_tab(disciplines)
    rows_data.append(["", f"='{first_report_sheet}'!A8", "", "", "", "", "група", "", "", "", "", "", group_name] + [""] * (total_cols - 13))
    rows_data.append([""] * total_cols)
    
    # Row 8: Table Header
    row8 = ["№ з/п", "Прізвище, ініціали студента", "Номер індивід. навч. плану", "Заліки, Екзамени"] + [""] * (total_cols - 4)
    rows_data.append(row8)
    
    # Row 9: Control types (З, Е, КП)
    row9 = ["", "", ""]
    for d in disciplines:
        ctype = d.get("control_type", "credit")
        label = "КП" if ctype == "course_project" else ("Е" if ctype == "exam" else "З")
        row9.extend([label, "", ""])
    rows_data.append(row9)
    
    # Row 10: Subject names
    row10 = ["", "", ""]
    for d in disciplines:
        row10.extend([d["name"], "", ""])
    rows_data.append(row10)
    
    # Student rows (Row 11+)
    for idx in range(num_students):
        row_num = 11 + idx
        v_row = 14 + idx
        
        no_cell = f"='{first_report_sheet}'!A{v_row}"
        name_cell = f"='{first_report_sheet}'!B{v_row}"
        record_book = f"='{first_report_sheet}'!C{v_row}"
        row = [no_cell, name_cell, record_book]
        
        for i, d in enumerate(disciplines):
            short = get_short_name(d["name"])
            v_tab = f"В_{short}" if d.get("control_type") != "course_project" else f"В_{short}_КР"
            ctype = d.get("control_type", "credit")
            
            if ctype == "course_project":
                ects_col = "F"
                score_col = "H"
            elif ctype == "exam":
                ects_col = "I"
                score_col = "J"
            else:
                ects_col = "E"
                score_col = "F"
                
            ects_dash_col = _col_letter(3 + i * 3 + 2)
            
            nat_formula = (
                f'=IFS({ects_dash_col}{row_num}="A",5,{ects_dash_col}{row_num}="B",4,'
                f'{ects_dash_col}{row_num}="C",4,{ects_dash_col}{row_num}="D",3,'
                f'{ects_dash_col}{row_num}="E",3,{ects_dash_col}{row_num}="F"," ",'
                f'{ects_dash_col}{row_num}="FX"," ")'
            )
            ects_formula = f"='{v_tab}'!{ects_col}{v_row}"
            score_formula = f"='{v_tab}'!{score_col}{v_row}"
            row.extend([nat_formula, ects_formula, score_formula])
            
        rows_data.append(row)
        
    end_cell = rowcol_to_a1(len(rows_data), total_cols)
    worksheet.update(f"A1:{end_cell}", rows_data, value_input_option="USER_ENTERED")
    
    # Styling and dimensions
    requests = [
        {
            "updateCells": {
                "range": {"sheetId": worksheet.id, "startRowIndex": 7, "endRowIndex": 10, "startColumnIndex": 0, "endColumnIndex": total_cols},
                "rows": [{
                    "values": [{
                        "userEnteredFormat": {
                            "textFormat": {"bold": True, "fontSize": 9},
                            "horizontalAlignment": "CENTER",
                            "verticalAlignment": "MIDDLE",
                            "backgroundColor": {"red": 0.85, "green": 0.88, "blue": 0.93},
                        }
                    } for _ in range(total_cols)]
                } for _ in range(3)],
                "fields": "userEnteredFormat(textFormat,horizontalAlignment,verticalAlignment,backgroundColor)"
            }
        },
        {"updateDimensionProperties": {"range": {"sheetId": worksheet.id, "dimension": "COLUMNS", "startIndex": 0, "endIndex": 1}, "properties": {"pixelSize": 45}, "fields": "pixelSize"}},
        {"updateDimensionProperties": {"range": {"sheetId": worksheet.id, "dimension": "COLUMNS", "startIndex": 1, "endIndex": 2}, "properties": {"pixelSize": 240}, "fields": "pixelSize"}},
        {"updateDimensionProperties": {"range": {"sheetId": worksheet.id, "dimension": "COLUMNS", "startIndex": 2, "endIndex": 3}, "properties": {"pixelSize": 100}, "fields": "pixelSize"}},
    ]
    
    worksheet.merge_cells(f"D8:{_col_letter(total_cols)}8", merge_type="MERGE_ALL")
    
    for i in range(len(disciplines)):
        sub_col = 4 + i * 3
        start_letter = _col_letter(sub_col)
        end_letter = _col_letter(sub_col + 2)
        worksheet.merge_cells(f"{start_letter}9:{end_letter}9", merge_type="MERGE_ALL")
        worksheet.merge_cells(f"{start_letter}10:{end_letter}10", merge_type="MERGE_ALL")
        
    border_style = {"style": "SOLID", "color": {"red": 0.7, "green": 0.7, "blue": 0.7}}
    requests.append({
        "updateBorders": {
            "range": {
                "sheetId": worksheet.id,
                "startRowIndex": 7,
                "endRowIndex": total_rows,
                "startColumnIndex": 0,
                "endColumnIndex": total_cols,
            },
            "top": border_style, "bottom": border_style, "left": border_style, "right": border_style,
            "innerHorizontal": border_style, "innerVertical": border_style,
        }
    })
    worksheet.spreadsheet.batch_update({"requests": requests})
    worksheet.freeze(rows=10, cols=2)
    return worksheet


def create_academic_journal(
    group_name: str,
    students: list[str],
    disciplines: list[dict],
    share_email: str | None = None,
    folder_id: str | None = None,
) -> dict:
    """Create a fully-formatted Google Sheets academic journal matching the reference workbook structure."""
    client = _get_client()
    
    # Create the Spreadsheet
    title = f"Журнал_{group_name}"
    spreadsheet = client.create(title)
    
    # Move to folder if specified
    if folder_id:
        try:
            client.move_spreadsheet(spreadsheet.id, folder_id)
        except Exception as exc:
            logger.warning("Could not move spreadsheet to folder %s: %s", folder_id, exc)
            
    # Sharing
    if share_email:
        spreadsheet.share(share_email, perm_type="user", role="writer")
        spreadsheet.share("", perm_type="anyone", role="reader")
        
    # Share with each teacher
    teacher_emails = set(d.get("teacher_email") for d in disciplines if d.get("teacher_email"))
    for t_email in teacher_emails:
        if t_email != share_email:
            try:
                spreadsheet.share(t_email, perm_type="user", role="writer")
            except Exception as exc:
                logger.warning("Could not share spreadsheet with teacher %s: %s", t_email, exc)
                
    # -------------------------------------------------------------
    # Pass 1: Create all worksheets with correct names and resize
    # -------------------------------------------------------------
    existing_sheet = spreadsheet.sheet1
    created_first = False
    first_pc_tab_name = None
    
    sheets_map = {}
    
    # A. ПК_* and КР_* sheets
    for disc in disciplines:
        disc_name = disc["name"]
        ctype = disc.get("control_type", "credit")
        short = get_short_name(disc_name)
        
        if ctype == "course_project":
            c_title = f"КР_{short}_Зміст"
            d_title = f"КР_{short}_Захист"
            
            # Content
            if not created_first:
                existing_sheet.update_title(c_title)
                sheets_map[c_title] = existing_sheet
                created_first = True
                first_pc_tab_name = c_title
            else:
                sheets_map[c_title] = spreadsheet.add_worksheet(title=c_title, rows=1, cols=1)
                
            # Defense
            sheets_map[d_title] = spreadsheet.add_worksheet(title=d_title, rows=1, cols=1)
        else:
            pc_title = f"ПК_{short}"
            if not created_first:
                existing_sheet.update_title(pc_title)
                sheets_map[pc_title] = existing_sheet
                created_first = True
                first_pc_tab_name = pc_title
            else:
                sheets_map[pc_title] = spreadsheet.add_worksheet(title=pc_title, rows=1, cols=1)
                
    # B. Milestone & helper sheets
    sheets_map["РК1"] = spreadsheet.add_worksheet(title="РК1", rows=1, cols=1)
    sheets_map["ДОД_РК1"] = spreadsheet.add_worksheet(title="ДОД_РК1", rows=1, cols=1)
    sheets_map["РК2"] = spreadsheet.add_worksheet(title="РК2", rows=1, cols=1)
    sheets_map["ДОД_РК2"] = spreadsheet.add_worksheet(title="ДОД_РК2", rows=1, cols=1)
    sheets_map["ІНДЗ"] = spreadsheet.add_worksheet(title="ІНДЗ", rows=1, cols=1)
    sheets_map["ДОД_ІНДЗ"] = spreadsheet.add_worksheet(title="ДОД_ІНДЗ", rows=1, cols=1)
    
    # C. СК sheet
    sheets_map["СК"] = spreadsheet.add_worksheet(title="СК", rows=1, cols=1)
    
    # D. В_* report sheets
    for disc in disciplines:
        disc_name = disc["name"]
        short = get_short_name(disc_name)
        ctype = disc.get("control_type", "credit")
        v_title = f"В_{short}" if ctype != "course_project" else f"В_{short}_КР"
        sheets_map[v_title] = spreadsheet.add_worksheet(title=v_title, rows=1, cols=1)
        
    # E. Consolidated Dashboard sheet
    sheets_map["Відомості"] = spreadsheet.add_worksheet(title="Відомості", rows=1, cols=1)
    
    # -------------------------------------------------------------
    # Pass 2: Populate and format worksheets in order
    # -------------------------------------------------------------
    # A. ПК_* and КР_* sheets
    for disc in disciplines:
        disc_name = disc["name"]
        ctype = disc.get("control_type", "credit")
        short = get_short_name(disc_name)
        teacher_email = disc.get("teacher_email")
        
        if ctype == "course_project":
            c_title = f"КР_{short}_Зміст"
            d_title = f"КР_{short}_Захист"
            _populate_kr_sheets(
                ws_content=sheets_map[c_title],
                ws_defense=sheets_map[d_title],
                students=students,
                discipline_name=disc_name,
                group_name=group_name,
                teacher_email=teacher_email,
                first_pc_tab_name=first_pc_tab_name,
            )
        else:
            pc_title = f"ПК_{short}"
            class_count = disc.get("class_count") or 0
            _populate_pc_sheet(
                worksheet=sheets_map[pc_title],
                students=students,
                discipline_name=disc_name,
                class_count=class_count,
                group_name=group_name,
                control_type=ctype,
                teacher_email=teacher_email,
                first_pc_tab_name=None if pc_title == first_pc_tab_name else first_pc_tab_name
            )
            
    # B. Milestone & helper sheets
    _create_milestone_sheet(spreadsheet, students, disciplines, 1, group_name, first_pc_tab_name)
    _create_helper_sheet(spreadsheet, students, disciplines, "ДОД_РК1", "РК1", first_pc_tab_name)
    _create_milestone_sheet(spreadsheet, students, disciplines, 2, group_name, first_pc_tab_name)
    _create_helper_sheet(spreadsheet, students, disciplines, "ДОД_РК2", "РК2", first_pc_tab_name)
    _create_helper_sheet(spreadsheet, students, disciplines, "ІНДЗ", None, first_pc_tab_name)
    _create_helper_sheet(spreadsheet, students, disciplines, "ДОД_ІНДЗ", "ІНДЗ", first_pc_tab_name)
    
    # C. СК sheet
    _create_sk_sheet(spreadsheet, students, disciplines, group_name, first_pc_tab_name)
    
    # D. В_* report sheets
    for disc in disciplines:
        disc_name = disc["name"]
        short = get_short_name(disc_name)
        ctype = disc.get("control_type", "credit")
        v_title = f"В_{short}" if ctype != "course_project" else f"В_{short}_КР"
        _create_v_report_sheet(
            worksheet=sheets_map[v_title],
            students=students,
            discipline=disc,
            disciplines=disciplines,
            group_name=group_name,
            first_pc_tab_name=first_pc_tab_name
        )
        
    # E. Consolidated Dashboard sheet
    _create_consolidated_dashboard(sheets_map["Відомості"], students, disciplines, group_name, first_pc_tab_name)
    
    return {
        "spreadsheet_id": spreadsheet.id,
        "spreadsheet_url": spreadsheet.url,
        "title": title,
    }
