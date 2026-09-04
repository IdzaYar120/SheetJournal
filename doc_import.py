"""
Import pipeline for the university's native document pair:
  - "Список студентів" (.doc) — roster of ALL groups/specialties for 3rd & 4th year.
  - "РНП" (.doc) — Робочий навчальний план (work curriculum) for ONE specialty,
    covering 4 semesters (two per course year).

Both are legacy Word 97-2003 binary files (.doc), which python-docx cannot open.
Text is extracted via the `antiword` utility, which ships with Git for Windows
(mingw64/bin/antiword.exe) — a machine that can run this project's git repo
already has it available.
"""

import datetime
import os
import re
import shutil
import subprocess
from pathlib import Path

# ==========================================================================
# antiword discovery & text extraction
# ==========================================================================

def _find_antiword() -> tuple[str, str] | None:
    """Locate antiword.exe and its mapping-file directory.

    Looks on PATH first, then falls back to a Git for Windows installation
    (antiword ships inside Git's bundled mingw64 toolchain).
    """
    exe = shutil.which("antiword")
    if exe:
        home = Path(exe).resolve().parent.parent / "share" / "antiword"
        return exe, str(home)

    git_exe = shutil.which("git")
    if git_exe:
        git_root = Path(git_exe).resolve().parent.parent  # .../Git
        candidate = git_root / "mingw64" / "bin" / "antiword.exe"
        if candidate.exists():
            home = git_root / "mingw64" / "share" / "antiword"
            return str(candidate), str(home)

    return None


def doc_to_text(filepath: str) -> str:
    """Extract plain UTF-8 text from a legacy .doc file using antiword."""
    found = _find_antiword()
    if not found:
        raise RuntimeError(
            "Не вдалося знайти утиліту 'antiword', потрібну для читання .doc файлів. "
            "Вона зазвичай встановлена разом з Git for Windows "
            "(шлях: <Git>\\mingw64\\bin\\antiword.exe). "
            "Перевірте встановлення Git, або встановіть antiword окремо і додайте його до PATH."
        )
    exe, home = found
    env = os.environ.copy()
    env["ANTIWORDHOME"] = home
    try:
        result = subprocess.run(
            [exe, "-m", "UTF-8.txt", filepath],
            capture_output=True,
            env=env,
            timeout=60,
        )
    except FileNotFoundError as exc:
        raise RuntimeError(f"Не вдалося запустити antiword: {exc}") from exc

    if result.returncode != 0:
        stderr = result.stderr.decode("utf-8", "ignore")
        raise RuntimeError(f"antiword не зміг обробити файл: {stderr}")

    return result.stdout.decode("utf-8", "ignore")


# ==========================================================================
# Student roster parsing ("Список студентів")
# ==========================================================================

_GROUP_HEADER_RE = re.compile(r"^([А-ЯІЇЄҐ]{2,6})-(\d)(\d)$")
_STUDENT_LINE_RE = re.compile(r"^\s*\d+\s*\.\s*(.+?)\s*$")


def parse_group_name(group_name: str) -> tuple[str, int, int] | None:
    """Split a group name like 'МН-31' into (specialty_code, course, group_seq)."""
    m = _GROUP_HEADER_RE.match(group_name.strip())
    if not m:
        return None
    specialty, course_digit, seq_digit = m.groups()
    return specialty, int(course_digit), int(seq_digit)


def parse_student_roster(filepath: str) -> dict:
    """Parse the multi-group, multi-specialty roster document.

    Returns {"groups": {group_name: [student_names...]}, "error": None|str}.
    """
    ext = Path(filepath).suffix.lower()
    try:
        text = doc_to_text(filepath) if ext == ".doc" else _docx_paragraph_text(filepath)
    except Exception as exc:
        return {"groups": {}, "error": str(exc)}

    groups: dict[str, list[str]] = {}
    current_group = None

    for line in text.splitlines():
        stripped = line.strip()
        if not stripped:
            continue

        if parse_group_name(stripped):
            current_group = stripped
            groups.setdefault(current_group, [])
            continue

        m = _STUDENT_LINE_RE.match(stripped)
        if m and current_group:
            name = m.group(1).strip()
            if name:
                groups[current_group].append(name)

    groups = {g: names for g, names in groups.items() if names}

    if not groups:
        return {"groups": {}, "error": "У документі не знайдено жодної групи зі студентами."}

    return {"groups": groups, "error": None}


def _docx_paragraph_text(filepath: str) -> str:
    """Fallback plain-paragraph text extraction for a .docx roster (no table)."""
    import docx

    doc = docx.Document(filepath)
    return "\n".join(p.text for p in doc.paragraphs)


# ==========================================================================
# RNP (curriculum plan) parsing
# ==========================================================================

_SEMESTER_HEADER_RE = re.compile(
    r"на\s+(\d+)-й\s+семестр.*?тижнів\s*\(([А-ЯІЇЄҐ]{2,6})\)",
    re.DOTALL,
)
_ROW_START_RE = re.compile(r"\|\s*\d+\s*\.\s*\|")

_EXCLUDED_NAME_KEYWORDS = ("виробнича практика", "публічний захист кваліфікаційної", "переддипломна практика")


def _clean_cell(cell: str) -> str:
    return re.sub(r"\s+", " ", cell).strip()


def _split_row(row_text: str) -> list[str]:
    return [_clean_cell(c) for c in row_text.split("|")]


def _parse_semester_block(block_text: str) -> list[dict]:
    """Parse one semester's discipline table rows out of its raw antiword block."""
    joined = re.sub(r"\s+", " ", block_text)

    # Isolate the discipline rows: after the last header cell, before the totals row.
    header_end = joined.rfind("Підсумковий контроль")
    totals_start = joined.find("Разом", header_end if header_end != -1 else 0)
    if header_end == -1 or totals_start == -1 or totals_start <= header_end:
        return []

    rows_region = joined[header_end + len("Підсумковий контроль"):totals_start]

    matches = list(_ROW_START_RE.finditer(rows_region))
    raw_rows = []
    for i, m in enumerate(matches):
        end = matches[i + 1].start() if i + 1 < len(matches) else len(rows_region)
        raw_rows.append(rows_region[m.end():end])

    parsed_rows = []
    for raw in raw_rows:
        cells = _split_row(raw)
        # Expected layout: Шифр, Назва, ЗагальнийОбсяг, ФормаКонтролю, ОбсягРоботи,
        #                  Кредити, АудВсього, Лекції, Практичні, Лабораторні,
        #                  Індивідуальні, ПідсумкКонтроль, Самостійна, (тижневе)
        if len(cells) < 9:
            continue
        name = cells[1]
        if not name:
            continue
        control_form = cells[3]
        try:
            credits = float(cells[5].replace(",", ".")) if cells[5] else None
        except ValueError:
            credits = None

        def _num(idx: int) -> float:
            try:
                return float(cells[idx].replace(",", ".")) if idx < len(cells) and cells[idx] else 0.0
            except ValueError:
                return 0.0

        lectures = _num(7)
        practicals = _num(8)
        labs = _num(9)

        parsed_rows.append({
            "shifr": cells[0],
            "name": name,
            "control_form": control_form,
            "credits": credits,
            "aud_hours": lectures + practicals + labs,
        })

    return parsed_rows


def _control_form_to_type(form: str) -> str | None:
    """Map RNP control-form abbreviation to the app's control_type. None = exclude."""
    f = form.strip().lower()
    if not f:
        return None
    if f in ("зп",):
        return None  # practicum credit — not a classroom subject
    if f in ("е/у", "е", "екзамен"):
        return "exam"
    if f in ("дз",):
        return "exam"  # graded (differentiated) credit — needs National/ECTS scale like exam
    if f in ("з", "залік"):
        return "credit"
    return None


def _is_excluded_name(name: str) -> bool:
    lower = name.lower()
    if "**" in name or "факультатив" in lower:
        return True
    return any(kw in lower for kw in _EXCLUDED_NAME_KEYWORDS)


def parse_rnp(filepath: str) -> dict:
    """Parse the RNP document into {specialty_code, semesters: {N: [discipline,...]}}.

    Each discipline dict: {name, control_type, class_count, credits, teacher_email}.
    A discipline with both an exam/credit AND a course-project component is
    represented as two consecutive entries (one "exam"/"credit", one
    "course_project") sharing the same base name.
    """
    ext = Path(filepath).suffix.lower()
    try:
        text = doc_to_text(filepath) if ext == ".doc" else _docx_paragraph_text(filepath)
    except Exception as exc:
        return {"specialty_code": None, "semesters": {}, "error": str(exc)}

    joined = re.sub(r"\s+", " ", text)
    headers = list(re.finditer(
        r"на\s+(\d+)-й\s+семестр", joined
    ))
    if not headers:
        return {"specialty_code": None, "semesters": {}, "error": "Не знайдено семестрових блоків у РНП."}

    specialty_code = None
    semesters: dict[int, list[dict]] = {}

    for i, h in enumerate(headers):
        block_start = h.start()
        block_end = headers[i + 1].start() if i + 1 < len(headers) else len(joined)
        block = joined[block_start:block_end]

        sem_num = int(h.group(1))

        code_match = re.search(r"тижнів\s*\(([А-ЯІЇЄҐ]{2,6})\)", block)
        if code_match:
            specialty_code = code_match.group(1)

        raw_rows = _parse_semester_block(block)

        disciplines: list[dict] = []
        pending_course_project = {}  # base_name -> credits, to merge onto parent

        for row in raw_rows:
            name = row["name"]

            cp_match = re.match(r"^(.*?)\s*\(курсов[аи]\s+(?:робот[аи]|проект[аи]?)\)\s*$", name, re.IGNORECASE)
            if cp_match and not row["control_form"]:
                base_name = cp_match.group(1).strip()
                pending_course_project[base_name] = row
                continue

            if _is_excluded_name(name):
                continue

            ctype = _control_form_to_type(row["control_form"])
            if ctype is None:
                continue

            class_count = max(1, round(row["aud_hours"] / 2)) if row["aud_hours"] else None
            disciplines.append({
                "name": name,
                "control_type": ctype,
                "class_count": class_count,
                "credits": row["credits"],
                "shifr": row["shifr"],
                "teacher_email": None,
            })

        # Attach course-project components right after their parent discipline.
        for base_name, cp_row in pending_course_project.items():
            for idx, d in enumerate(disciplines):
                if d["name"].strip().lower() == base_name.strip().lower():
                    disciplines.insert(idx + 1, {
                        "name": d["name"],
                        "control_type": "course_project",
                        "class_count": None,
                        "credits": cp_row["credits"],
                        "shifr": cp_row["shifr"],
                        "teacher_email": None,
                    })
                    break

        if disciplines:
            semesters[sem_num] = disciplines

    if not semesters:
        return {"specialty_code": specialty_code, "semesters": {}, "error": "Не вдалося розпізнати дисципліни в РНП."}

    return {"specialty_code": specialty_code, "semesters": semesters, "error": None}


# ==========================================================================
# Semester auto-detection
# ==========================================================================

def current_semester_for_course(course: int, today: datetime.date | None = None) -> int:
    """Map a course year (1-4) to the RNP semester number currently in progress.

    Autumn half (Sep–Jan, plus the Jul/Aug break defaulting to the upcoming
    autumn term) is the odd semester of the course year; Feb–Jun is the even one.
    """
    today = today or datetime.date.today()
    is_autumn = not (2 <= today.month <= 6)
    return (course - 1) * 2 + (1 if is_autumn else 2)


# ==========================================================================
# Matching roster groups to an RNP specialty
# ==========================================================================

def match_groups_to_rnp(roster_groups: dict, rnp: dict, today: datetime.date | None = None) -> dict:
    """Build a per-group journal-ready dataset for every roster group whose
    specialty prefix matches the RNP's specialty code.

    Returns {"matched": {group_name: {"students":.., "semester":.., "disciplines":..}},
             "skipped": {group_name: reason}}.
    """
    specialty_code = rnp.get("specialty_code")
    matched = {}
    skipped = {}

    for group_name, students in roster_groups.items():
        parsed = parse_group_name(group_name)
        if not parsed:
            skipped[group_name] = "Не вдалося розпізнати формат назви групи."
            continue
        prefix, course, _seq = parsed
        if prefix != specialty_code:
            continue  # belongs to a different specialty — not this RNP's concern

        semester = current_semester_for_course(course, today)
        disciplines = rnp["semesters"].get(semester)
        if not disciplines:
            skipped[group_name] = f"У РНП немає даних для {semester}-го семестру."
            continue

        import copy
        matched[group_name] = {
            "students": students,
            "semester": semester,
            "course": course,
            "disciplines": copy.deepcopy(disciplines),
        }

    return {"matched": matched, "skipped": skipped}
