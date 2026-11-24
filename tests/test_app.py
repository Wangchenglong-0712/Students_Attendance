"""接口测试：导入、检索、登记、统计、导出。数据库位于 pytest 临时目录。"""
import io
from urllib.parse import unquote_plus

from fastapi.testclient import TestClient
from openpyxl import Workbook, load_workbook

from app.main import create_app


def loc(resp) -> str:
    """重定向地址（已解码，便于断言中文提示）。"""
    return unquote_plus(resp.headers["location"])

HEADER = ["姓名", "学号", "专业", "班级", "年级"]
ROSTER = [
    HEADER,
    ["张三", "2023010101", "计算机", "2班", "2023"],
    ["李四", "2023010102", "计算机", "2班", "2023"],
    ["张三", "2023010201", "软件", "1班", "2023"],  # 与第一名重名
]


def make_xlsx(rows) -> bytes:
    wb = Workbook()
    ws = wb.active
    for row in rows:
        ws.append(row)
    out = io.BytesIO()
    wb.save(out)
    wb.close()
    return out.getvalue()


def make_client(tmp_path):
    app = create_app(tmp_path / "test.db")
    return TestClient(app)


def import_roster(client, rows=ROSTER):
    resp = client.post("/import", files={"file": ("roster.xlsx", make_xlsx(rows))},
                       follow_redirects=False)
    assert resp.status_code == 303, resp.text
    return resp


# ---------- 导入 ----------

def test_index_empty(tmp_path):
    client = make_client(tmp_path)
    resp = client.get("/")
    assert resp.status_code == 200
    assert "尚未导入信息表" in resp.text


def test_import_success_and_list(tmp_path):
    client = make_client(tmp_path)
    resp = import_roster(client)
    assert "导入成功，共 3 名学生" in loc(resp)
    home = client.get("/")
    assert "2023010101" in home.text and "张三" in home.text


def test_import_missing_header(tmp_path):
    client = make_client(tmp_path)
    resp = client.post("/import",
                       files={"file": ("bad.xlsx", make_xlsx([["姓名", "学号"], ["张三", "1"]]))},
                       follow_redirects=False)
    assert "缺少必需表头" in loc(resp)


def test_import_duplicate_student_id(tmp_path):
    client = make_client(tmp_path)
    rows = [HEADER, ["张三", "1"], ["李四", "1"]]
    resp = client.post("/import", files={"file": ("dup.xlsx", make_xlsx(rows))},
                       follow_redirects=False)
    assert "学号重复" in loc(resp)


def test_import_invalid_file(tmp_path):
    client = make_client(tmp_path)
    resp = client.post("/import", files={"file": ("x.xlsx", b"not an excel")},
                       follow_redirects=False)
    assert "无法解析" in loc(resp)


def test_reimport_replaces_roster(tmp_path):
    client = make_client(tmp_path)
    import_roster(client)
    new_roster = [HEADER, ["王五", "2023010101", "计算机", "2班", "2023"]]  # 沿用学号
    import_roster(client, new_roster)
    home = client.get("/")
    assert "王五" in home.text and "李四" not in home.text


# ---------- 检索 ----------

def test_search_by_name_redirects(tmp_path):
    client = make_client(tmp_path)
    import_roster(client)
    resp = client.get("/search", params={"q": "李四"}, follow_redirects=False)
    assert resp.status_code == 303
    assert loc(resp).startswith("/student/2023010102")


def test_search_by_student_id_redirects(tmp_path):
    client = make_client(tmp_path)
    import_roster(client)
    resp = client.get("/search", params={"q": "2023010102"}, follow_redirects=False)
    assert loc(resp) == "/student/2023010102"


def test_search_duplicate_name_lists_candidates(tmp_path):
    client = make_client(tmp_path)
    import_roster(client)
    resp = client.get("/search", params={"q": "张三"})
    assert resp.status_code == 200
    assert "找到 2 名匹配" in resp.text
    assert "2023010101" in resp.text and "2023010201" in resp.text


def test_search_no_match(tmp_path):
    client = make_client(tmp_path)
    import_roster(client)
    resp = client.get("/search", params={"q": "不存在"}, follow_redirects=False)
    assert "未找到" in loc(resp)


def test_student_detail_not_found(tmp_path):
    client = make_client(tmp_path)
    resp = client.get("/student/999", follow_redirects=False)
    assert resp.status_code == 303
    assert "该学生不存在" in loc(resp)


# ---------- 登记与统计 ----------

def add(client, sid, week, periods, rtype, note="", weekday="星期一"):
    return client.post(f"/student/{sid}/records",
                       data={"week": week, "weekday": weekday, "rtype": rtype,
                             "note": note, "period": periods},
                       follow_redirects=False)


def test_add_records_and_weekly_totals(tmp_path):
    client = make_client(tmp_path)
    import_roster(client)
    resp = add(client, "2023010101", 5, ["1-2", "3-4"], "leave", "病假")
    assert "登记成功 2 条" in loc(resp)
    add(client, "2023010101", 5, ["5-6"], "absent")
    page = client.get("/student/2023010101", params={"week": 5})
    # 每节次对 = 2 学时
    assert "请假 <b class=\"leave\">4 学时</b>" in page.text
    assert "旷课 <b class=\"absent\">2 学时</b>" in page.text
    assert page.text.count("第 5 周") >= 2  # 顶部合计 + 记录分组


def test_duplicate_period_skipped(tmp_path):
    client = make_client(tmp_path)
    import_roster(client)
    add(client, "2023010101", 5, ["1-2"], "leave")
    resp = add(client, "2023010101", 5, ["1-2"], "absent")
    assert "跳过已存在的 1 条" in loc(resp)
    # 类型不被覆盖，仍为请假
    page = client.get("/student/2023010101")
    assert page.text.count("请假") >= 1


def test_weeks_are_independent(tmp_path):
    client = make_client(tmp_path)
    import_roster(client)
    add(client, "2023010101", 3, ["1-2"], "absent")
    add(client, "2023010101", 4, ["1-2"], "leave")
    page3 = client.get("/student/2023010101", params={"week": 3})
    assert "请假 <b class=\"leave\">0 学时</b>" in page3.text
    assert "旷课 <b class=\"absent\">2 学时</b>" in page3.text


def test_invalid_period_rejected(tmp_path):
    client = make_client(tmp_path)
    import_roster(client)
    resp = add(client, "2023010101", 5, ["2-3"], "leave")
    assert "非法节次" in loc(resp)


def test_invalid_week_rejected(tmp_path):
    client = make_client(tmp_path)
    import_roster(client)
    resp = add(client, "2023010101", "abc", ["1-2"], "leave")
    assert "周次必须是正整数" in loc(resp)
    resp = add(client, "2023010101", 0, ["1-2"], "leave")
    assert "周次必须是正整数" in loc(resp)


def test_no_period_rejected(tmp_path):
    client = make_client(tmp_path)
    import_roster(client)
    resp = add(client, "2023010101", 5, [], "leave")
    assert "请至少选择一个节次" in loc(resp)


def test_update_record_type(tmp_path):
    client = make_client(tmp_path)
    import_roster(client)
    add(client, "2023010101", 5, ["1-2"], "leave")
    page = client.get("/student/2023010101")
    # 找到记录 id：通过再次登记前先解析
    import re
    m = re.search(r"/records/(\d+)/type", page.text)
    assert m, "未找到类型修改按钮"
    rid = m.group(1)
    resp = client.post(f"/records/{rid}/type",
                       data={"student_id": "2023010101", "rtype": "absent"},
                       follow_redirects=False)
    assert "类型已更新" in loc(resp)
    page2 = client.get("/student/2023010101", params={"week": 5})
    assert "请假 <b class=\"leave\">0 学时</b>" in page2.text
    assert "旷课 <b class=\"absent\">2 学时</b>" in page2.text


def test_delete_record(tmp_path):
    client = make_client(tmp_path)
    import_roster(client)
    add(client, "2023010101", 5, ["1-2"], "leave")
    page = client.get("/student/2023010101")
    import re
    m = re.search(r"/records/(\d+)/delete", page.text)
    assert m
    resp = client.post(f"/records/{m.group(1)}/delete",
                       data={"student_id": "2023010101"}, follow_redirects=False)
    assert "记录已删除" in loc(resp)
    page2 = client.get("/student/2023010101")
    assert "暂无缺勤记录" in page2.text


# ---------- 导出 ----------

def test_export_report(tmp_path):
    client = make_client(tmp_path)
    import_roster(client)
    add(client, "2023010101", 5, ["1-2", "3-4"], "leave")
    add(client, "2023010101", 5, ["5-6"], "absent")
    add(client, "2023010102", 5, ["11-12"], "absent")
    resp = client.get("/export")
    assert resp.status_code == 200
    assert "spreadsheetml" in resp.headers["content-type"]
    assert "attachment" in resp.headers["content-disposition"]

    wb = load_workbook(io.BytesIO(resp.content))
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    header = rows[0]
    assert list(header) == ["序号", "学号", "姓名", "班级", "旷课（合计）",
                            "请假（合计）", "课时（缺勤记录）", "备注"]

    by_sid = {r[1]: r for r in rows[1:]}
    zhang = by_sid["2023010101"]
    assert zhang[0] == 1
    assert zhang[3] == "计算机232"          # 专业+年级后两位+班级号
    assert zhang[4] == 2 and zhang[5] == 4  # 旷课 2 学时、请假 4 学时
    # 全学期导出时明细带“第N周 ”前缀
    assert zhang[6] == "第5周 星期一：1-2（假），3-4（假），5-6（旷）"
    assert zhang[7] in ("", None)           # 备注留空
    li = by_sid["2023010102"]
    assert li[4] == 2 and li[5] == 0
    assert li[6] == "第5周 星期一：11-12（旷）"
    # 无缺勤记录的学生不出现在报表中
    assert "2023010201" not in by_sid
    wb.close()


def test_export_specific_week(tmp_path):
    client = make_client(tmp_path)
    import_roster(client)
    add(client, "2023010101", 5, ["1-2"], "leave", weekday="星期五")
    add(client, "2023010101", 5, ["5-6"], "absent", weekday="星期三")
    add(client, "2023010101", 7, ["3-4"], "absent")    # 同一人第7周也有
    add(client, "2023010102", 7, ["1-2"], "absent")    # 另一人只有第7周
    resp = client.get("/export", params={"week": 5})
    assert resp.status_code == 200
    wb = load_workbook(io.BytesIO(resp.content))
    rows = list(wb.active.iter_rows(values_only=True))
    header = rows[0]
    assert list(header) == ["序号", "学号", "姓名", "班级", "旷课（合计）",
                            "请假（合计）", "课时（缺勤记录）", "备注"]
    by_sid = {r[1]: r for r in rows[1:]}
    assert list(by_sid.keys()) == ["2023010101"]       # 只含该周有记录的学生
    row = by_sid["2023010101"]
    assert row[4] == 2 and row[5] == 2                 # 旷课 2 学时、请假 2 学时
    # 指定周次时明细不带周次前缀，星期之间两个空格分隔
    assert row[6] == "星期三：5-6（旷）  星期五：1-2（假）"
    wb.close()


def test_export_detail_same_day_multiple(tmp_path):
    """同一天内多次缺勤用“，”连接节次。"""
    client = make_client(tmp_path)
    import_roster(client)
    add(client, "2023010101", 5, ["5-6"], "leave", weekday="星期一")
    add(client, "2023010101", 5, ["9-10"], "absent", weekday="星期一")
    add(client, "2023010101", 5, ["1-2"], "absent", weekday="星期二")
    resp = client.get("/export", params={"week": 5})
    wb = load_workbook(io.BytesIO(resp.content))
    rows = list(wb.active.iter_rows(values_only=True))
    detail = {r[1]: r for r in rows[1:]}["2023010101"][6]
    assert detail == "星期一：5-6（假），9-10（旷）  星期二：1-2（旷）"
    wb.close()


def test_export_class_label_format(tmp_path):
    """班级列 = 专业 + 年级后两位 + 班级号，如 软件工程（中外合作办学）241。"""
    client = make_client(tmp_path)
    import_roster(client, [HEADER,
                           ["王五", "2024010101", "软件工程（中外合作办学）", "1班", "2024"]])
    add(client, "2024010101", 5, ["1-2"], "absent")
    resp = client.get("/export", params={"week": 5})
    wb = load_workbook(io.BytesIO(resp.content))
    rows = list(wb.active.iter_rows(values_only=True))
    assert rows[1][3] == "软件工程（中外合作办学）241"
    wb.close()


def test_same_period_different_weekday_allowed(tmp_path):
    client = make_client(tmp_path)
    import_roster(client)
    r1 = add(client, "2023010101", 5, ["1-2"], "leave", weekday="星期一")
    r2 = add(client, "2023010101", 5, ["1-2"], "absent", weekday="星期五")
    assert "登记成功 1 条" in loc(r1)
    assert "登记成功 1 条" in loc(r2)   # 不同星期的同节次可分别登记
    page = client.get("/student/2023010101", params={"week": 5})
    assert "请假 <b class=\"leave\">2 学时</b>" in page.text
    assert "旷课 <b class=\"absent\">2 学时</b>" in page.text


def test_invalid_weekday_rejected(tmp_path):
    client = make_client(tmp_path)
    import_roster(client)
    resp = add(client, "2023010101", 5, ["1-2"], "leave", weekday="星期八")
    assert "星期必须是星期一至星期日" in loc(resp)


def test_student_page_shows_weekday(tmp_path):
    client = make_client(tmp_path)
    import_roster(client)
    add(client, "2023010101", 5, ["1-2"], "leave", weekday="星期五")
    page = client.get("/student/2023010101")
    assert 'name="weekday"' in page.text          # 登记表单有星期选择
    assert "<td>星期五</td>" in page.text         # 记录列表展示星期


def test_db_migration_adds_weekday(tmp_path):
    """旧版数据库（无 weekday 列）升级后历史记录保留。"""
    import sqlite3

    from app import db as dbmod
    path = tmp_path / "old.db"
    conn = sqlite3.connect(str(path))
    conn.executescript(
        "CREATE TABLE absence_records ("
        " id INTEGER PRIMARY KEY AUTOINCREMENT, student_id TEXT NOT NULL,"
        " week INTEGER NOT NULL, period TEXT NOT NULL,"
        " type TEXT NOT NULL CHECK (type IN ('absent','leave')),"
        " note TEXT, created_at TEXT NOT NULL,"
        " UNIQUE (student_id, week, period));"
        " INSERT INTO absence_records (student_id, week, period, type, note, created_at)"
        " VALUES ('2023010101', 3, '1-2', 'leave', '', '2026-01-01 00:00:00');")
    conn.commit()
    conn.close()

    conn2 = dbmod.connect(path)
    dbmod.init(conn2)
    rows = conn2.execute("SELECT * FROM absence_records").fetchall()
    assert len(rows) == 1
    assert rows[0]["weekday"] == ""
    assert rows[0]["period"] == "1-2"
    conn2.close()


def test_export_week_without_records(tmp_path):
    client = make_client(tmp_path)
    import_roster(client)
    add(client, "2023010101", 5, ["1-2"], "leave")
    resp = client.get("/export", params={"week": 9})
    assert resp.status_code == 200
    wb = load_workbook(io.BytesIO(resp.content))
    rows = list(wb.active.iter_rows(values_only=True))
    assert len(rows) == 1          # 只有表头，无学生
    wb.close()


def test_index_shows_export_weeks(tmp_path):
    client = make_client(tmp_path)
    import_roster(client)
    add(client, "2023010101", 5, ["1-2"], "leave")
    home = client.get("/")
    assert 'value="5">第 5 周' in home.text
    assert "仅导出有缺勤记录的学生" in home.text


def test_export_empty_roster(tmp_path):
    client = make_client(tmp_path)
    resp = client.get("/export")
    assert resp.status_code == 200
