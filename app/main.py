"""FastAPI 路由：登录、导入、检索、学生详情、缺勤登记、统计导出。"""
import hashlib
import hmac
import os
import urllib.parse
from pathlib import Path

from fastapi import FastAPI, File, Form, Request, UploadFile
from fastapi.responses import RedirectResponse, Response
from fastapi.templating import Jinja2Templates

from . import db, services

ROOT = Path(__file__).resolve().parent.parent
TEMPLATES = Jinja2Templates(directory=str(Path(__file__).parent / "templates"))

USERNAME = "bfmzdx"
PASSWORD = "wclshidashuaige"
# 登录凭据的摘要，作为会话 Cookie 值（不存明文密码）
SESSION_VALUE = hashlib.sha256(f"{USERNAME}:{PASSWORD}".encode()).hexdigest()
# 无需登录即可访问的路径
PUBLIC_PATHS = {"/login", "/logout"}


def _redirect(path: str, msg: str = "", err: str = "") -> RedirectResponse:
    params = {}
    if msg:
        params["msg"] = msg
    if err:
        params["err"] = err
    if params:
        path = path + ("&" if "?" in path else "?") + urllib.parse.urlencode(params)
    return RedirectResponse(path, status_code=303)


def create_app(db_path=None) -> FastAPI:
    app = FastAPI(title="学生缺勤登记系统")
    path = Path(db_path or os.environ.get("ABSENCE_DB") or ROOT / "data" / "absence.db")
    conn = db.connect(path)
    db.init(conn)
    app.state.conn = conn
    templates = TEMPLATES

    def page(request: Request, name: str, ctx: dict):
        ctx.update({"msg": request.query_params.get("msg", ""),
                    "err": request.query_params.get("err", ""),
                    "logged_in": _is_authed(request)})
        return templates.TemplateResponse(request, name, ctx)

    def _is_authed(request: Request) -> bool:
        return hmac.compare_digest(request.cookies.get("session", ""),
                                   SESSION_VALUE)

    # ---------- 登录保护 ----------
    @app.middleware("http")
    async def require_login(request: Request, call_next):
        path = request.url.path
        if path in PUBLIC_PATHS or _is_authed(request):
            return await call_next(request)
        return RedirectResponse("/login", status_code=303)

    @app.get("/login")
    def login_page(request: Request):
        if _is_authed(request):
            return RedirectResponse("/", status_code=303)
        return templates.TemplateResponse(request, "login.html",
                                          {"err": request.query_params.get("err", "")})

    @app.post("/login")
    async def login_submit(request: Request):
        form = await request.form()
        username = form.get("username", "")
        password = form.get("password", "")
        if hmac.compare_digest(username, USERNAME) and \
                hmac.compare_digest(password, PASSWORD):
            resp = RedirectResponse("/", status_code=303)
            resp.set_cookie("session", SESSION_VALUE, httponly=True, samesite="lax")
            return resp
        return templates.TemplateResponse(
            request, "login.html", {"err": "账号或密码错误"})

    @app.get("/logout")
    def logout():
        resp = RedirectResponse("/login", status_code=303)
        resp.delete_cookie("session")
        return resp

    # ---------- 首页：导入 + 检索 + 名单 ----------
    @app.get("/")
    def index(request: Request):
        students = services.list_students(app.state.conn)
        return page(request, "index.html",
                    {"students": students, "count": len(students),
                     "export_weeks": services.all_weeks(app.state.conn),
                     "periods": services.PERIODS})

    @app.post("/import")
    async def import_roster(file: UploadFile = File(...)):
        data = await file.read()
        try:
            result = services.import_students(app.state.conn, data)
        except services.ImportError_ as e:
            return _redirect("/", err=str(e))
        return _redirect("/", msg=f"导入成功，共 {result['total']} 名学生")

    @app.get("/search")
    def search(request: Request, q: str = ""):
        q = q.strip()
        if not q:
            return _redirect("/", err="请输入姓名或学号")
        results = services.search_students(app.state.conn, q)
        if not results:
            return _redirect("/", err=f"未找到与“{q}”匹配的学生")
        if len(results) == 1:
            return _redirect(f"/student/{results[0]['student_id']}")
        return page(request, "candidates.html",
                    {"results": results, "query": q})

    # ---------- 学生详情 ----------
    @app.get("/student/{student_id}")
    def student_detail(request: Request, student_id: str):
        student = services.get_student(app.state.conn, student_id)
        if student is None:
            return _redirect("/", err="该学生不存在（可能已被新的信息表覆盖）")
        records = services.list_records(app.state.conn, student_id)
        summary = services.weekly_summary(app.state.conn, student_id)
        weeks = sorted(summary.keys(), reverse=True)
        try:
            week = int(request.query_params.get("week", weeks[0] if weeks else 1))
        except ValueError:
            week = weeks[0] if weeks else 1
        current = summary.get(week, {"leave": 0, "absent": 0})
        # 按周分组记录
        grouped = {}
        for r in records:
            grouped.setdefault(r["week"], []).append(r)
        return page(request, "student.html",
                    {"student": student, "grouped": grouped,
                     "summary": summary, "weeks": weeks,
                     "week": week, "current": current,
                     "periods": services.PERIODS,
                     "weekdays": services.WEEKDAYS,
                     "type_label": services.TYPE_LABEL})

    @app.post("/student/{student_id}/records")
    async def add_records(request: Request, student_id: str):
        if services.get_student(app.state.conn, student_id) is None:
            return _redirect("/", err="该学生不存在")
        form = await request.form()
        week = form.get("week", "")
        weekday = form.get("weekday", "")
        rtype = form.get("rtype", "")
        note = form.get("note", "")
        periods = form.getlist("period")
        try:
            result = services.add_records(
                app.state.conn, student_id, week, weekday, periods, rtype, note)
        except ValueError as e:
            return _redirect(f"/student/{student_id}", err=str(e))
        msg = f"登记成功 {result['created']} 条"
        if result["skipped"]:
            msg += f"，跳过已存在的 {result['skipped']} 条（如需修改请直接改类型）"
        return _redirect(f"/student/{student_id}", msg=msg)

    @app.post("/records/{record_id}/type")
    def update_type(record_id: int, student_id: str = Form(...), rtype: str = Form(...)):
        try:
            ok = services.update_record_type(app.state.conn, record_id, rtype)
        except ValueError as e:
            return _redirect(f"/student/{student_id}", err=str(e))
        if not ok:
            return _redirect(f"/student/{student_id}", err="记录不存在")
        return _redirect(f"/student/{student_id}", msg="类型已更新")

    @app.post("/records/{record_id}/delete")
    def delete_record(record_id: int, student_id: str = Form(...)):
        ok = services.delete_record(app.state.conn, record_id)
        if not ok:
            return _redirect(f"/student/{student_id}", err="记录不存在")
        return _redirect(f"/student/{student_id}", msg="记录已删除")

    # ---------- 导出 ----------
    @app.get("/export")
    def export(week: int = None):
        if week is not None and week < 1:
            return _redirect("/", err="周次必须是正整数")
        data = services.export_report(app.state.conn, week)
        filename = urllib.parse.quote("缺勤统计报表.xlsx")
        return Response(
            content=data,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition":
                     f"attachment; filename*=UTF-8''{filename}"})

    return app


app = create_app()
