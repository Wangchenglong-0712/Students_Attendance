"""业务逻辑：导入、检索、缺勤登记、统计、导出。"""
import io
import json
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


def export_report(conn: sqlite3.Connection, week=None) -> bytes:
    """统计报表：只导出有缺勤记录的学生；week 为空导出全部周次，否则仅该周。"""
    if week is None:
        weeks = all_weeks(conn)
        week_filter, params = "", []
    else:
        weeks = [int(week)]
        week_filter = " WHERE week = ?"
        params = [int(week)]
    stats: dict = {}
    for r in conn.execute(
            "SELECT student_id, week, type, COUNT(*) AS n FROM absence_records"
            + week_filter + " GROUP BY student_id, week, type", params).fetchall():
        stats.setdefault(r["student_id"], {}).setdefault(
            r["week"], {"leave": 0, "absent": 0})[r["type"]] = r["n"] * 2
    # 只保留有缺勤记录的学生
    students = [s for s in list_students(conn) if s["student_id"] in stats]

    # 指定周次时生成缺勤明细，如“星期五：1-2（假）；星期三：5-6（旷）”
    details: dict = {}
    if week is not None:
        for r in conn.execute(
                "SELECT student_id, weekday, period, type FROM absence_records"
                " WHERE week = ?", (int(week),)).fetchall():
            mark = "假" if r["type"] == "leave" else "旷"
            entry = f"{r['weekday']}：{r['period']}（{mark}）" if r["weekday"] \
                else f"{r['period']}（{mark}）"
            details.setdefault(r["student_id"], []).append(
                (WEEKDAY_ORDER.get(r["weekday"], 99),
                 PERIOD_ORDER.get(r["period"], 99), entry))
        for sid in details:
            details[sid] = "；".join(e for _, _, e in sorted(details[sid]))

    wb = Workbook()
    ws = wb.active
    ws.title = "缺勤统计"
    header = ["姓名", "学号", "专业", "班级", "年级"]
    for w in weeks:
        header += [f"第{w}周 请假(学时)", f"第{w}周 旷课(学时)"]
    header += ["请假合计(学时)", "旷课合计(学时)"]
    if week is not None:
        header.append("缺勤明细")
    ws.append(header)

    for s in students:
        row = [s["name"], s["student_id"], s["major"] or "", s["class_name"] or "",
               s["grade"] or ""]
        leave_total = absent_total = 0
        for w in weeks:
            cell = stats.get(s["student_id"], {}).get(w, {"leave": 0, "absent": 0})
            row += [cell["leave"], cell["absent"]]
            leave_total += cell["leave"]
            absent_total += cell["absent"]
        row += [leave_total, absent_total]
        if week is not None:
            row.append(details.get(s["student_id"], ""))
        ws.append(row)

    out = io.BytesIO()
    wb.save(out)
    wb.close()
    return out.getvalue()
