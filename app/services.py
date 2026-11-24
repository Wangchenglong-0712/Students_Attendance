"""业务逻辑：导入、检索、缺勤登记、统计、导出。"""
import io
import json
import re
import sqlite3
from datetime import datetime

from openpyxl import Workbook, load_workbook

PERIODS = ["1-2", "3-4", "5-6", "7-8", "9-10", "11-12"]
PERIOD_ORDER = {p: i for i, p in enumerate(PERIODS)}
WEEKDAYS = ["星期一", "星期二", "星期三", "星期四", "星期五", "星期六", "星期日"]
WEEKDAY_ORDER = {w: i for i, w in enumerate(WEEKDAYS)}
REQUIRED_HEADERS = ["姓名", "学号", "专业", "班级", "年级"]
HEADER_FIELD = {"姓名": "name", "学号": "student_id", "专业": "major",
                "班级": "class_name", "年级": "grade"}
VALID_TYPES = ("absent", "leave")
TYPE_LABEL = {"absent": "旷课", "leave": "请假"}


class ImportError_(Exception):
    """信息表导入失败（用户可读的原因）。"""


# ---------- 导入 ----------

def import_students(conn: sqlite3.Connection, data: bytes) -> dict:
    """解析 xlsx 并整体覆盖学生名单。返回 {'total': n} 或抛出 ImportError_。"""
    try:
        wb = load_workbook(io.BytesIO(data), read_only=True, data_only=True)
    except Exception:
        raise ImportError_("文件无法解析，请确认是有效的 xlsx 表格")
    try:
        ws = wb.active
        rows = ws.iter_rows(values_only=True)
        header_row = next(rows, None)
        if header_row is None:
            raise ImportError_("表格为空")
        headers = [str(h).strip() if h is not None else "" for h in header_row]
        missing = [h for h in REQUIRED_HEADERS if h not in headers]
        if missing:
            raise ImportError_("缺少必需表头：" + "、".join(missing))
        idx = {h: i for i, h in enumerate(headers) if h}
        students = []
        seen = set()
        for row in rows:
            if row is None:
                continue
            sid = row[idx["学号"]]
            name = row[idx["姓名"]]
            if sid is None and name is None:
                continue
            sid = str(sid).strip() if sid is not None else ""
            name = str(name).strip() if name is not None else ""
            if not sid or not name:
                raise ImportError_("存在学号或姓名为空的数据行，请检查表格")
            if sid in seen:
                raise ImportError_(f"学号重复：{sid}")
            seen.add(sid)
            record = {"student_id": sid, "name": name}
            for header, field in HEADER_FIELD.items():
                if field in ("student_id", "name"):
                    continue
                value = row[idx[header]]
                record[field] = str(value).strip() if value is not None else ""
            extra = {h: row[i] for h, i in idx.items()
                     if h not in HEADER_FIELD and i < len(row) and row[i] is not None}
            record["extra_json"] = json.dumps(extra, ensure_ascii=False, default=str)
            students.append(record)
    finally:
        wb.close()

    if not students:
        raise ImportError_("表格中没有数据行")

    # 单表覆盖：清空旧名单；学号仍在新表中的学生保留缺勤记录，其余记录一并删除
    conn.execute("DELETE FROM students")
    conn.executemany(
        "INSERT INTO students (student_id, name, major, class_name, grade, extra_json)"
        " VALUES (:student_id, :name, :major, :class_name, :grade, :extra_json)",
        students)
    conn.execute(
        "DELETE FROM absence_records WHERE student_id NOT IN (SELECT student_id FROM students)")
    conn.commit()
    return {"total": len(students)}


def count_students(conn: sqlite3.Connection) -> int:
    return conn.execute("SELECT COUNT(*) FROM students").fetchone()[0]


def list_students(conn: sqlite3.Connection) -> list:
    return conn.execute("SELECT * FROM students ORDER BY student_id").fetchall()


# ---------- 检索 ----------

def search_students(conn: sqlite3.Connection, query: str) -> list:
    q = query.strip()
    if not q:
        return []
    rows = conn.execute(
        "SELECT * FROM students WHERE student_id = ? OR name LIKE ? OR student_id LIKE ?"
        " ORDER BY student_id", (q, f"%{q}%", f"%{q}%")).fetchall()
    # 精确学号、精确姓名排前
    exact = [r for r in rows if r["student_id"] == q or r["name"] == q]
    rest = [r for r in rows if r not in exact]
    return exact + rest


def get_student(conn: sqlite3.Connection, student_id: str):
    return conn.execute("SELECT * FROM students WHERE student_id = ?",
                        (student_id,)).fetchone()


# ---------- 缺勤登记 ----------

def validate_input(week, weekday, periods, rtype):
    try:
        week = int(week)
    except (TypeError, ValueError):
        raise ValueError("周次必须是正整数")
    if week < 1:
        raise ValueError("周次必须是正整数")
    if weekday not in WEEKDAYS:
        raise ValueError("星期必须是星期一至星期日")
    if not periods:
        raise ValueError("请至少选择一个节次")
    bad = [p for p in periods if p not in PERIODS]
    if bad:
        raise ValueError("非法节次：" + "、".join(bad))
    if rtype not in VALID_TYPES:
        raise ValueError("缺勤类型必须是旷课或请假")
    return week, weekday, list(dict.fromkeys(periods)), rtype


def add_records(conn: sqlite3.Connection, student_id: str, week, weekday, periods,
                rtype: str, note: str = "") -> dict:
    """登记缺勤，返回 {'created': n, 'skipped': m}（同周同星期同节次已存在则跳过）。"""
    week, weekday, periods, rtype = validate_input(week, weekday, periods, rtype)
    now = datetime.now().isoformat(timespec="seconds")
    created = skipped = 0
    for p in periods:
        cur = conn.execute(
            "INSERT OR IGNORE INTO absence_records"
            " (student_id, week, weekday, period, type, note, created_at)"
            " VALUES (?, ?, ?, ?, ?, ?, ?)",
            (student_id, week, weekday, p, rtype, note.strip(), now))
        if cur.rowcount:
            created += 1
        else:
            skipped += 1
    conn.commit()
    return {"created": created, "skipped": skipped}


def list_records(conn: sqlite3.Connection, student_id: str) -> list:
    rows = conn.execute(
        "SELECT * FROM absence_records WHERE student_id = ?",
        (student_id,)).fetchall()
    return sorted(rows, key=lambda r: (-r["week"],
                                       WEEKDAY_ORDER.get(r["weekday"], 99),
                                       PERIOD_ORDER.get(r["period"], 99)))


def update_record_type(conn: sqlite3.Connection, record_id: int, rtype: str) -> bool:
    if rtype not in VALID_TYPES:
        raise ValueError("缺勤类型必须是旷课或请假")
    cur = conn.execute("UPDATE absence_records SET type = ? WHERE id = ?",
                       (rtype, record_id))
    conn.commit()
    return cur.rowcount > 0


def delete_record(conn: sqlite3.Connection, record_id: int) -> bool:
    cur = conn.execute("DELETE FROM absence_records WHERE id = ?", (record_id,))
    conn.commit()
    return cur.rowcount > 0


# ---------- 统计 ----------

def weekly_summary(conn: sqlite3.Connection, student_id: str) -> dict:
    """{week: {'leave': 学时, 'absent': 学时}}，每个节次对 = 2 学时。"""
    rows = conn.execute(
        "SELECT week, type, COUNT(*) AS n FROM absence_records"
        " WHERE student_id = ? GROUP BY week, type", (student_id,)).fetchall()
    summary: dict = {}
    for r in rows:
        summary.setdefault(r["week"], {"leave": 0, "absent": 0})
        summary[r["week"]][r["type"]] = r["n"] * 2
    return summary


# ---------- 导出 ----------

def all_weeks(conn: sqlite3.Connection) -> list:
    """所有存在缺勤记录的周次（升序）。"""
    return [r[0] for r in conn.execute(
        "SELECT DISTINCT week FROM absence_records ORDER BY week").fetchall()]


def _class_label(student) -> str:
    """班级列：专业 + 年级后两位 + 班级号，如 计算机科学与技术262。
    若班级列本身已含年级前缀（如“262”已含“26”），则直接用班级列编号，避免重复拼接。"""
    major = student["major"] or ""
    grade_digits = re.sub(r"\D", "", student["grade"] or "")
    grade_part = grade_digits[-2:] if grade_digits else ""
    cls = student["class_name"] or ""
    m = re.search(r"\d+", cls)
    cls_no = m.group() if m else cls
    if grade_part and cls_no.startswith(grade_part):
        return f"{major}{cls_no}"
    return f"{major}{grade_part}{cls_no}"


def _major_rank(major: str) -> int:
    """专业排序：计科-软工-网工-智科-中外，其余排最后。"""
    if "中外" in major:
        return 4
    if "计算机" in major or "计科" in major:
        return 0
    if "软件" in major:
        return 1
    if "网络" in major:
        return 2
    if "智能" in major:
        return 3
    return 5


def _export_sort_key(student):
    """年级升序（2024→2026）→ 专业顺序 → 班号 → 学号。"""
    grade_digits = re.sub(r"\D", "", student["grade"] or "")
    if grade_digits:
        g = int(grade_digits)
        grade = 2000 + g if g < 100 else g
    else:
        grade = 0
    grade_part = grade_digits[-2:] if grade_digits else ""
    cls_digits = re.sub(r"\D", "", student["class_name"] or "")
    # 班号可能含年级前缀（如“262”），先剥离再比较
    if grade_part and cls_digits.startswith(grade_part) \
            and len(cls_digits) > len(grade_part):
        cls_digits = cls_digits[len(grade_part):]
    cls_no = int(cls_digits) if cls_digits.isdigit() else 0
    return (grade, _major_rank(student["major"] or ""), cls_no,
            student["student_id"])


def _detail_text(conn: sqlite3.Connection, weeks: list, with_week_prefix: bool) -> dict:
    """{学号: 缺勤明细文字}。同星期节次用“，”连接，星期之间两个空格分隔；
    导出多周时每段前加“第N周 ”前缀。"""
    marks = {"leave": "假", "absent": "旷"}
    grouped: dict = {}
    placeholders = ",".join("?" * len(weeks))
    rows = conn.execute(
        "SELECT student_id, week, weekday, period, type FROM absence_records"
        f" WHERE week IN ({placeholders}) ORDER BY week", list(weeks)).fetchall()
    for r in rows:
        grouped.setdefault(r["student_id"], {}).setdefault(r["week"], {}).setdefault(
            r["weekday"], []).append(
            (PERIOD_ORDER.get(r["period"], 99), f"{r['period']}（{marks[r['type']]}）"))

    details: dict = {}
    for sid, by_week in grouped.items():
        week_parts = []
        for w in sorted(by_week):
            day_parts = []
            for wd, entries in sorted(by_week[w].items(),
                                      key=lambda kv: WEEKDAY_ORDER.get(kv[0], 99)):
                text = "，".join(e for _, e in sorted(entries))
                day_parts.append(f"{wd}：{text}" if wd else text)
            week_text = "  ".join(day_parts)
            if with_week_prefix:
                week_text = f"第{w}周 {week_text}"
            week_parts.append(week_text)
        details[sid] = "  ".join(week_parts)
    return details


def export_report(conn: sqlite3.Connection, week=None) -> bytes:
    """统计报表：表头固定，只导出有缺勤记录的学生；
    week 为空统计全学期并汇总明细，否则仅统计该周。"""
    if week is None:
        weeks = all_weeks(conn)
        week_filter, params = "", []
        with_week_prefix = True
    else:
        weeks = [int(week)]
        week_filter = " WHERE week = ?"
        params = [int(week)]
        with_week_prefix = False
    if not weeks:
        weeks = [0]  # 全无记录时仅用于占位，保证输出表头
    stats: dict = {}
    for r in conn.execute(
            "SELECT student_id, week, type, COUNT(*) AS n FROM absence_records"
            + week_filter + " GROUP BY student_id, week, type", params).fetchall():
        stats.setdefault(r["student_id"], {}).setdefault(
            r["week"], {"leave": 0, "absent": 0})[r["type"]] = r["n"] * 2
    # 只保留有缺勤记录的学生，按 年级 → 专业（计科-软工-网工-智科-中外）→ 班号 排序
    students = sorted((s for s in list_students(conn) if s["student_id"] in stats),
                      key=_export_sort_key)
    details = _detail_text(conn, weeks, with_week_prefix) if stats else {}

    wb = Workbook()
    ws = wb.active
    ws.title = "缺勤统计"
    ws.append(["序号", "学号", "姓名", "班级", "旷课（合计）", "请假（合计）",
               "课时（缺勤记录）", "备注"])

    for i, s in enumerate(students, start=1):
        absent_total = leave_total = 0
        for w in weeks:
            cell = stats.get(s["student_id"], {}).get(w, {"leave": 0, "absent": 0})
            absent_total += cell["absent"]
            leave_total += cell["leave"]
        ws.append([i, s["student_id"], s["name"], _class_label(s),
                   absent_total, leave_total,
                   details.get(s["student_id"], ""), ""])

    out = io.BytesIO()
    wb.save(out)
    wb.close()
    return out.getvalue()
