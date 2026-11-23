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

def add(client, sid, week, periods, rtype, note=""):
    return client.post(f"/student/{sid}/records",
                       data={"week": week, "rtype": rtype, "note": note,
                             "period": periods},
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
    assert list(header[:5]) == ["姓名", "学号", "专业", "班级", "年级"]
    assert "第5周 请假(学时)" in header and "第5周 旷课(学时)" in header
    assert list(header[-2:]) == ["请假合计(学时)", "旷课合计(学时)"]

    by_sid = {r[1]: r for r in rows[1:]}
    zhang = by_sid["2023010101"]
    assert zhang[-2] == 4 and zhang[-1] == 2   # 请假 4 学时，旷课 2 学时
    li = by_sid["2023010102"]
    assert li[-2] == 0 and li[-1] == 2
    wb.close()


def test_export_empty_roster(tmp_path):
    client = make_client(tmp_path)
    resp = client.get("/export")
    assert resp.status_code == 200
