# -*- coding: utf-8 -*-
"""
志愿时长系统 — 纯 API 服务（FastAPI）

- 不渲染任何 HTML，仅提供 JSON / 文件上传接口。
- 全局 CORS 允许任意来源，便于通过 file:// 打开的本地 index.html 调用。
- 数据逻辑复用 services/excel_service.py；路径来自 config.BaseConfig。

启动（示例）：
    uvicorn api:app --host 127.0.0.1 --port 8000 --reload

或：
    python api.py
"""
from __future__ import annotations

import os
import re
import uuid
import mimetypes
import json
import hashlib
import base64
from contextlib import asynccontextmanager
from typing import Any, Dict, Optional
from urllib.parse import quote

import time

from pydantic import BaseModel

from fastapi import Depends, FastAPI, File, Form, Header, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import Response
from starlette.concurrency import run_in_threadpool

from config import BaseConfig
from services import excel_service

cfg = BaseConfig()


def _ensure_dirs() -> None:
    os.makedirs(cfg.DATA_DIR, exist_ok=True)
    os.makedirs(cfg.PROOF_UPLOAD_DIR, exist_ok=True)
    os.makedirs(cfg.TOPUP_PROOF_INTERNAL_DIR, exist_ok=True)
    os.makedirs(cfg.TOPUP_PROOF_EXTERNAL_DIR, exist_ok=True)
    _ensure_admin_users_file()


def _ensure_admin_users_file() -> None:
    """
    初始化管理员账号文件（不存在时创建）。
    文件支持两种密码字段：
    - password: 明文（仅本地开发建议）
    - password_sha256: sha256 哈希（更推荐）
    """
    if os.path.isfile(cfg.ADMIN_USERS_PATH):
        return
    default_data = {
        "users": [
            {"username": "admin", "password": "admin123"},
        ]
    }
    with open(cfg.ADMIN_USERS_PATH, "w", encoding="utf-8") as f:
        json.dump(default_data, f, ensure_ascii=False, indent=2)


def _sha256_text(raw: str) -> str:
    return hashlib.sha256((raw or "").encode("utf-8")).hexdigest()


def _load_admin_users() -> list[dict[str, str]]:
    """
    每次登录时动态读取账号文件，支持你直接改文件增删用户而无需重启。
    """
    _ensure_admin_users_file()
    try:
        with open(cfg.ADMIN_USERS_PATH, "r", encoding="utf-8") as f:
            data = json.load(f)
    except Exception as exc:
        raise HTTPException(
            status_code=500,
            detail={"ok": False, "message": f"管理员账号文件读取失败: {exc}", "data": []},
        )
    users = data.get("users", [])
    if not isinstance(users, list):
        raise HTTPException(
            status_code=500,
            detail={"ok": False, "message": "管理员账号文件格式错误：users 必须是数组", "data": []},
        )
    cleaned: list[dict[str, str]] = []
    for u in users:
        if not isinstance(u, dict):
            continue
        username = str(u.get("username", "")).strip()
        pwd = str(u.get("password", ""))
        pwd_hash = str(u.get("password_sha256", "")).strip().lower()
        if not username:
            continue
        cleaned.append({"username": username, "password": pwd, "password_sha256": pwd_hash})
    return cleaned


def _parse_positive_int(field_name: str, raw: str) -> int:
    raw = (raw or "").strip()
    if not re.fullmatch(r"\d+", raw or ""):
        raise HTTPException(
            status_code=400,
            detail={"ok": False, "message": f"{field_name} 必须是数字"},
        )
    return int(raw)


@asynccontextmanager
async def lifespan(app: FastAPI):
    _ensure_dirs()
    yield


app = FastAPI(
    title="志愿时长公示 API",
    description="前后端分离：查询总表、提交反馈与证明文件",
    version="2.0.0",
    lifespan=lifespan,
)

# ---------------------------------------------------------------------------
# CORS：允许所有来源（含 origin 为 null 的 file:// 页面）
# 注意：allow_origins=["*"] 时 allow_credentials 须为 False（规范限制）
# ---------------------------------------------------------------------------
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=False,
    allow_methods=["*"],
    allow_headers=["*"],
)


def _secure_filename(name: str) -> str:
    base = os.path.basename(name or "")
    base = re.sub(r"[^\w\u4e00-\u9fff.\-]", "_", base)
    return (base[:180] or "file").strip("._") or "file"


def _proof_allowed(filename: str) -> bool:
    if not filename or "." not in filename:
        return False
    ext = filename.rsplit(".", 1)[1].lower()
    return ext in cfg.ALLOWED_PROOF_EXTENSIONS


@app.get("/health")
def health() -> Dict[str, str]:
    return {"status": "ok"}


@app.get("/api/search")
def api_search(q: str = "", name: str = "", student_id: str = "", grade: str = "") -> Dict[str, Any]:
    """
    志愿时长查询接口。
    """
    if not grade:
        path = cfg.MASTER_EXCEL_PATH
    else:
        path = cfg.get_master_excel_path(grade)

    if name and student_id:
        ok, msg, rows = excel_service.search_student_by_name_and_id(
            path, name=name, student_id=student_id
        )
    else:
        if not q:
            raise HTTPException(
                status_code=400,
                detail={"ok": False, "message": "请同时输入姓名与学号", "data": []},
            )
        ok, msg, rows = excel_service.search_student(path, q)

    if not ok:
        raise HTTPException(status_code=400, detail={"ok": False, "message": msg, "data": []})

    return {"ok": True, "message": "查询成功", "data": rows}


@app.get("/api/master-last-upload")
def api_master_last_upload(grade: str = "") -> Dict[str, Any]:
    """
    返回指定年级总表的最后上传时间（用于查询页面提示）。
    """
    grade = (grade or "").strip()
    if not grade:
        raise HTTPException(status_code=400, detail={"ok": False, "message": "缺少年级参数"})

    path = cfg.get_master_excel_path(grade)
    if not os.path.isfile(path):
        return {"ok": True, "grade": grade, "exists": False, "last_uploaded": "尚未上传"}

    mtime = os.path.getmtime(path)
    dt_str = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(mtime))
    return {"ok": True, "grade": grade, "exists": True, "last_uploaded": dt_str}


# ---------------------------------------------------------------------------
# 管理员认证与文件管理（原型：固定账号密码 + 内存 token）
# ---------------------------------------------------------------------------

# 管理员 Token 存储： {token: expiry_timestamp}
_ADMIN_TOKENS: dict[str, float] = {}
_SENIOR_TOKENS: dict[str, float] = {}  # 高级管理员专用 Token
ADMIN_TOKEN_TTL_SECONDS = 3600 * 2  # 2小时有效期


class AdminLoginRequest(BaseModel):
    username: str
    password: str


class TopupDeleteRequest(BaseModel):
    submit_time: str
    name: str


class TopupBatchDeleteRequest(BaseModel):
    items: list[TopupDeleteRequest]


class TopupUpdateRequest(BaseModel):
    items: list[TopupDeleteRequest]
    status: Optional[str] = None
    admin_note: Optional[str] = None


def _extract_bearer_token(authorization: Optional[str]) -> str:
    if not authorization:
        raise HTTPException(status_code=401, detail={"ok": False, "message": "未登录", "data": []})
    if not authorization.lower().startswith("bearer "):
        raise HTTPException(
            status_code=401,
            detail={"ok": False, "message": "认证格式错误，请使用 Bearer Token", "data": []},
        )
    token = authorization.split(" ", 1)[1].strip()
    if not token:
        raise HTTPException(status_code=401, detail={"ok": False, "message": "认证为空", "data": []})
    return token


def require_admin(authorization: Optional[str] = Header(default=None)) -> None:
    token = _extract_bearer_token(authorization)
    exp = _ADMIN_TOKENS.get(token)
    if not exp or exp < time.time():
        raise HTTPException(status_code=401, detail={"ok": False, "message": "登录已过期或 token 无效", "data": []})


class SeniorLoginRequest(BaseModel):
    username: str
    password: str


def require_senior_admin(authorization: Optional[str] = Header(default=None)) -> None:
    """
    不仅需要普通管理员登录，还需要通过高级管理员二次验证。
    """
    token = _extract_bearer_token(authorization)
    # 同时检查普通 Token 和高级 Token (这里我们让高级 Token 独立管理)
    exp = _SENIOR_TOKENS.get(token)
    if not exp or exp < time.time():
        raise HTTPException(status_code=403, detail={"ok": False, "message": "需要高级管理员权限或权限已过期"})


@app.post("/api/admin/login")
def admin_login(req: AdminLoginRequest) -> Dict[str, Any]:
    username = (req.username or "").strip()
    password = req.password or ""
    users = _load_admin_users()

    matched = False
    for u in users:
        if username != u["username"]:
            continue
        # 优先使用 password_sha256 校验；否则使用明文 password
        if u.get("password_sha256"):
            if _sha256_text(password) == u["password_sha256"]:
                matched = True
                break
        else:
            if password == u.get("password", ""):
                matched = True
                break

    if not matched:
        raise HTTPException(status_code=401, detail={"ok": False, "message": "管理员账号或密码错误", "data": []})

    token = uuid.uuid4().hex
    _ADMIN_TOKENS[token] = time.time() + ADMIN_TOKEN_TTL_SECONDS
    return {"ok": True, "message": "登录成功", "token": token}


@app.post("/api/admin/senior-login")
def senior_admin_login(req: SeniorLoginRequest) -> Dict[str, Any]:
    """
    高级管理员验证接口
    账号：admin123
    密码：admin456
    """
    if req.username == "admin123" and req.password == "admin456":
        token = uuid.uuid4().hex
        _SENIOR_TOKENS[token] = time.time() + ADMIN_TOKEN_TTL_SECONDS
        return {"ok": True, "message": "高级管理员验证成功", "senior_token": token}
    else:
        raise HTTPException(status_code=401, detail={"ok": False, "message": "高级账号或密码错误"})


def _master_allowed(filename: str) -> bool:
    if not filename or "." not in filename:
        return False
    ext = filename.rsplit(".", 1)[1].lower()
    return ext in cfg.ALLOWED_MASTER_EXTENSIONS


@app.post("/api/admin/upload-master")
async def admin_upload_master(
    grade: str = Form(""),  # 新增年级参数
    file: UploadFile = File(...),
    _=Depends(require_senior_admin),  # 改为 senior 校验
) -> Dict[str, Any]:
    if not _master_allowed(file.filename or ""):
        raise HTTPException(
            status_code=400,
            detail={"ok": False, "message": f"仅允许上传 {', '.join(sorted(cfg.ALLOWED_MASTER_EXTENSIONS))} 格式的总表", "data": []},
        )

    _ensure_dirs()
    body = await file.read()
    max_bytes = getattr(cfg, "MAX_CONTENT_LENGTH", 16 * 1024 * 1024)
    if len(body) > max_bytes:
        raise HTTPException(status_code=400, detail={"ok": False, "message": "文件过大", "data": []})

    if grade:
        save_path = cfg.get_master_excel_path(grade)
        # 保存一份源文件供管理员下载
        raw_path = os.path.join(cfg.DATA_DIR, f"学生志愿时长总表_{grade}_raw.xlsx")
        with open(raw_path, "wb") as f:
            f.write(body)
    else:
        save_path = cfg.MASTER_EXCEL_PATH

    ok = excel_service.merge_and_save_master(body, save_path)
    if not ok:
        raise HTTPException(status_code=400, detail={"ok": False, "message": "解析或合并分表失败，请检查文件格式"})

    # 另存为一份 CSV 以便快速查询（固定公式为数值）
    excel_service.save_master_as_csv(save_path)

    return {"ok": True, "message": f"年级 {grade or '默认'} 的总表上传并成功合并所有分表"}


@app.post("/api/admin/append-durations")
async def admin_append_durations(
    pending_file: UploadFile = File(...),
    master_file: UploadFile = File(...),
    multi_sheet: bool = Form(False),
    _=Depends(require_admin),
) -> Dict[str, Any]:
    """
    时长录入功能：将待录入时长表中的时长追加到总表中，并返回更新后的文件 Base64。
    """
    if not _master_allowed(pending_file.filename or "") or not _master_allowed(master_file.filename or ""):
        raise HTTPException(
            status_code=400,
            detail={"ok": False, "message": "仅允许上传 Excel 格式文件"},
        )

    pending_body = await pending_file.read()
    master_body = await master_file.read()
    
    # 获取活动名称（待录入表文件名，去后缀）
    activity_name = os.path.splitext(pending_file.filename or "新活动")[0]
    
    updated_bytes, failed_names, success_count, grade_success_counts = await run_in_threadpool(
        excel_service.append_durations_to_master,
        pending_body,
        master_body,
        activity_name,
        multi_sheet,
    )
    
    if updated_bytes is None:
        raise HTTPException(
            status_code=400,
            detail={"ok": False, "message": "处理失败，请检查文件格式或内容是否正确"},
        )

    # 编码为 Base64 以便随 JSON 返回
    file_b64 = base64.b64encode(updated_bytes).decode("utf-8")
    
    return {
        "ok": True,
        "failed_names": failed_names,
        "success_count": success_count,
        "grade_success_counts": grade_success_counts,
        "file_b64": file_b64,
        "suggested_filename": f"{os.path.splitext(master_file.filename or 'total')[0]}_更新后.xlsx"
    }


@app.post("/api/admin/merge-grades")
async def admin_merge_grades(
    file: UploadFile = File(...),
    _=Depends(require_admin),
) -> Response:
    """
    上传单份 Excel 总表并返回合并分表后的文件。
    """
    if not _master_allowed(file.filename or ""):
        raise HTTPException(
            status_code=400,
            detail={"ok": False, "message": "仅允许上传 Excel 格式文件"},
        )

    body = await file.read()
    
    # 只需要合并一份文件内的分表，复用 merge_multiple_excels_to_bytes 即可
    merged_bytes = excel_service.merge_multiple_excels_to_bytes([body])
    
    if not merged_bytes:
        raise HTTPException(status_code=400, detail={"ok": False, "message": "合并失败，请检查文件内容是否正确"})

    # 使用原始文件名加上 _merged 后缀，或者固定名称
    original_name = os.path.splitext(file.filename or "result")[0]
    filename = f"{original_name}_合并结果.xlsx"
    filename_enc = quote(filename, safe="")
    headers = {
        "Content-Disposition": f"attachment; filename*=UTF-8''{filename_enc}",
    }
    return Response(
        content=merged_bytes,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers=headers,
    )


@app.get("/api/public/master-file-status")
def public_master_file_status(grade: str) -> Dict[str, Any]:
    """
    公开接口：获取指定年级总表的最后更新时间。
    """
    if grade not in ["2022", "2023", "2024", "2025"]:
        raise HTTPException(status_code=400, detail={"ok": False, "message": "无效的年级"})
        
    path = cfg.get_master_excel_path(grade)
    if os.path.isfile(path):
        mtime = os.path.getmtime(path)
        dt_str = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(mtime))
        return {"ok": True, "last_uploaded": dt_str, "exists": True}
    else:
        return {"ok": True, "last_uploaded": "尚未上传", "exists": False}


@app.get("/api/admin/master-files")
def admin_master_files(_=Depends(require_admin)) -> Dict[str, Any]:
    """
    管理员查看各年级总表的最后更新时间。
    """
    grades = ["2022", "2023", "2024", "2025"]
    data = []
    for g in grades:
        path = cfg.get_master_excel_path(g)
        if os.path.isfile(path):
            mtime = os.path.getmtime(path)
            dt_str = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(mtime))
            data.append({"grade": g, "last_uploaded": dt_str, "exists": True})
        else:
            data.append({"grade": g, "last_uploaded": "尚未上传", "exists": False})
    return {"ok": True, "data": data}


@app.get("/api/admin/download-master")
def admin_download_master(grade: str, _=Depends(require_admin)) -> Response:
    """
    管理员下载指定年级的总表。
    优先提供高级管理员上传的原始（未合并）文件。
    """
    # 优先提供原始文件
    raw_path = os.path.join(cfg.DATA_DIR, f"学生志愿时长总表_{grade}_raw.xlsx")
    if os.path.isfile(raw_path):
        with open(raw_path, "rb") as f:
            content = f.read()
        filename = f"学生志愿时长总表_{grade}_原始文件.xlsx"
        filename_enc = quote(filename, safe="")
        headers = {
            "Content-Disposition": f"attachment; filename*=UTF-8''{filename_enc}",
        }
        return Response(
            content=content,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers=headers,
        )

    # 如果原始文件不存在（可能是旧数据），则提供合并后的总表
    path = cfg.get_master_excel_path(grade)
    if not os.path.isfile(path):
        raise HTTPException(status_code=404, detail={"ok": False, "message": f"{grade}级总表尚未上传"})

    with open(path, "rb") as f:
        content = f.read()

    filename = f"学生志愿时长总表_{grade}.xlsx"
    filename_enc = quote(filename, safe="")
    headers = {
        "Content-Disposition": f"attachment; filename*=UTF-8''{filename_enc}",
    }
    return Response(
        content=content,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers=headers,
    )


@app.get("/api/admin/download-feedback")
def admin_download_feedback(_=Depends(require_admin)) -> Response:
    # 兼容旧接口：仍然返回“反馈汇总”，若你不再使用可忽略该路由
    data = excel_service.export_feedback_summary_bytes(cfg.FEEDBACK_EXCEL_PATH)
    if not data:
        raise HTTPException(status_code=404, detail={"ok": False, "message": "当前还没有反馈记录可导出", "data": []})

    # 注意：Starlette 在响应头编码时要求 latin-1，直接写中文文件名可能导致 UnicodeEncodeError。
    # 这里使用 RFC 5987：filename*=UTF-8''<percent-encoded>
    filename = "学生反馈汇总.xlsx"
    filename_enc = quote(filename, safe="")
    headers = {
        "Content-Disposition": f"attachment; filename*=UTF-8''{filename_enc}",
    }
    return Response(
        content=data,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers=headers,
    )


@app.get("/api/admin/download-topup")
def admin_download_topup(_=Depends(require_admin)) -> Response:
    """
    下载“学生时长补录汇总.xlsx”（校内/校外追加写入同一份文件）。
    """
    data = excel_service.export_topup_summary_bytes(cfg.TOPUP_EXCEL_PATH)
    if not data:
        raise HTTPException(status_code=404, detail={"ok": False, "message": "当前还没有补录记录可导出", "data": []})

    filename = "学生时长补录汇总.xlsx"
    filename_enc = quote(filename, safe="")
    headers = {
        "Content-Disposition": f"attachment; filename*=UTF-8''{filename_enc}",
    }
    return Response(
        content=data,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers=headers,
    )


@app.get("/api/download-master-raw")
def download_master_raw(grade: str) -> Response:
    """
    允许学生下载管理员上传的源 Excel 文件（针对 2022/2023 级）。
    """
    if grade not in ["2022", "2023"]:
        raise HTTPException(status_code=403, detail={"ok": False, "message": "该年级总表不支持直接下载"})
    
    # 源文件保存在 DATA_DIR 下，文件名为：学生志愿时长总表_{grade}_raw.xlsx
    # 注意：我们需要在上传时同时保存一份 _raw 后缀的文件
    raw_path = os.path.join(cfg.DATA_DIR, f"学生志愿时长总表_{grade}_raw.xlsx")
    
    if not os.path.isfile(raw_path):
        # 如果不存在 _raw，尝试返回现有的（可能是合并后的，或者旧逻辑上传的）
        raw_path = cfg.get_master_excel_path(grade)
        
    if not os.path.isfile(raw_path):
        raise HTTPException(status_code=404, detail={"ok": False, "message": "该年级总表文件尚未上传"})

    with open(raw_path, "rb") as f:
        data = f.read()

    filename = f"{grade}级志愿时长总表.xlsx"
    filename_enc = quote(filename, safe="")
    headers = {
        "Content-Disposition": f"attachment; filename*=UTF-8''{filename_enc}",
    }
    return Response(
        content=data,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers=headers,
    )


@app.get("/api/admin/topup-records")
def admin_topup_records(
    limit: int = 200,
    _=Depends(require_admin),
) -> Dict[str, Any]:
    """
    管理员查看补录记录（JSON）。
    默认返回最近 200 条。
    """
    if limit <= 0:
        limit = 200
    limit = min(limit, 1000)

    df = excel_service.load_topup_dataframe(cfg.TOPUP_EXCEL_PATH)
    if df.empty:
        return {"ok": True, "message": "暂无补录记录", "data": []}

    # 兼容 Excel 列缺失情况
    for c in excel_service.TOPUP_COLUMNS:
        if c not in df.columns:
            df[c] = ""

    # 按提交时间倒序（字符串时间格式 YYYY-MM-DD HH:MM:SS 可直接字典序）
    if "提交时间" in df.columns:
        df = df.sort_values(by="提交时间", ascending=False, na_position="last")

    # 前端表格展示使用的字段子集
    cols = [
        "提交时间",
        "补录类型",
        "姓名",
        "班级",
        "年级",
        "志愿活动所处学年",
        "活动名称",
        "时长",
        "活动项数",
        "微信号联系方式",
        "处理状态",
        "管理员备注",
        "证明文件相对路径",
        "本人活动照片相对路径",
        "备注",
    ]
    view = df[cols].head(limit).copy()
    view = view.fillna("")

    records: list[dict[str, Any]] = []
    for _, row in view.iterrows():
        records.append({k: (row[k].item() if hasattr(row[k], "item") else row[k]) for k in cols})

    return {"ok": True, "message": "查询成功", "data": records}


@app.post("/api/admin/delete-topup")
def admin_delete_topup(
    req: TopupDeleteRequest,
    _=Depends(require_admin),
) -> Dict[str, Any]:
    """
    管理员删除补录记录。
    """
    ok = excel_service.delete_topup_row(
        cfg.TOPUP_EXCEL_PATH,
        submit_time=req.submit_time,
        name=req.name
    )
    if not ok:
        raise HTTPException(status_code=404, detail={"ok": False, "message": "未找到匹配记录或删除失败"})

    return {"ok": True, "message": "删除成功"}


@app.post("/api/admin/delete-topup-batch")
def admin_delete_topup_batch(
    req: TopupBatchDeleteRequest,
    _=Depends(require_admin),
) -> Dict[str, Any]:
    """
    管理员批量删除补录记录。
    """
    items = [{"submit_time": i.submit_time, "name": i.name} for i in req.items]
    count = excel_service.delete_topup_rows_batch(cfg.TOPUP_EXCEL_PATH, items)
    
    return {"ok": True, "message": f"成功删除 {count} 条记录"}


@app.post("/api/admin/update-topup-batch")
def admin_update_topup_batch(
    req: TopupUpdateRequest,
    _=Depends(require_admin),
) -> Dict[str, Any]:
    """
    管理员批量更新补录记录的状态或管理员备注。
    """
    items = [{"submit_time": i.submit_time, "name": i.name} for i in req.items]
    count = excel_service.update_topup_rows_batch(
        cfg.TOPUP_EXCEL_PATH, 
        items,
        status=req.status,
        admin_note=req.admin_note
    )
    
    return {"ok": True, "message": f"成功更新 {count} 条记录"}


@app.get("/api/admin/uploaded-file")
def admin_uploaded_file(
    rel_path: str,
    _=Depends(require_admin),
) -> Response:
    """
    管理员读取上传附件（用于网页预览/下载）。
    仅允许访问 BASE_DIR 下 uploads/proofs 目录中的文件，防止路径穿越。
    """
    rel_path = (rel_path or "").strip()
    if not rel_path:
        raise HTTPException(status_code=400, detail={"ok": False, "message": "缺少 rel_path 参数", "data": []})

    base_dir = os.path.abspath(cfg.BASE_DIR)
    uploads_root = os.path.abspath(cfg.PROOF_UPLOAD_DIR)
    target_abs = os.path.abspath(os.path.join(base_dir, rel_path))

    # 安全限制：只能访问 uploads/proofs 下的文件
    if not target_abs.startswith(uploads_root + os.sep):
        raise HTTPException(status_code=403, detail={"ok": False, "message": "无权访问该文件", "data": []})
    if not os.path.isfile(target_abs):
        raise HTTPException(status_code=404, detail={"ok": False, "message": "文件不存在", "data": []})

    with open(target_abs, "rb") as f:
        data = f.read()

    mime, _ = mimetypes.guess_type(target_abs)
    media_type = mime or "application/octet-stream"
    filename = os.path.basename(target_abs)
    filename_enc = quote(filename, safe="")
    headers = {
        "Content-Disposition": f"inline; filename*=UTF-8''{filename_enc}",
    }
    return Response(content=data, media_type=media_type, headers=headers)


@app.post("/api/feedback")
async def api_feedback(
    name: str = Form(...),
    class_name: str = Form(...),
    grade: str = Form(...),
    note: str = Form(""),
    proof: UploadFile = File(...),
) -> Dict[str, Any]:
    """
    接收反馈表单与证明文件：追加写入反馈汇总 Excel，文件落盘到 uploads/proofs。
    """
    name = (name or "").strip()
    class_name = (class_name or "").strip()
    grade = (grade or "").strip()
    note = (note or "").strip()

    if not name or not class_name or not grade:
        raise HTTPException(
            status_code=400,
            detail={"ok": False, "message": "请填写姓名、班级、年级"},
        )

    if not proof.filename:
        raise HTTPException(
            status_code=400,
            detail={"ok": False, "message": "请上传证明文件"},
        )

    if not _proof_allowed(proof.filename):
        raise HTTPException(
            status_code=400,
            detail={
                "ok": False,
                "message": f"证明文件类型不允许，允许：{', '.join(sorted(cfg.ALLOWED_PROOF_EXTENSIONS))}",
            },
        )

    _ensure_dirs()
    safe = _secure_filename(proof.filename)
    unique_name = f"{uuid.uuid4().hex}_{safe}"
    dest_abs = os.path.join(cfg.PROOF_UPLOAD_DIR, unique_name)

    body = await proof.read()
    max_bytes = getattr(cfg, "MAX_CONTENT_LENGTH", 16 * 1024 * 1024)
    if len(body) > max_bytes:
        raise HTTPException(
            status_code=400,
            detail={"ok": False, "message": "文件过大"},
        )

    with open(dest_abs, "wb") as f:
        f.write(body)

    try:
        rel = os.path.relpath(dest_abs, cfg.BASE_DIR)
    except ValueError:
        rel = dest_abs

    excel_service.append_feedback_row(
        cfg.FEEDBACK_EXCEL_PATH,
        name=name,
        clazz=class_name,
        grade=grade,
        note=note,
        proof_relative_path=rel,
    )

    return {"ok": True, "message": "反馈已提交，感谢你的配合"}


# ---------------------------------------------------------------------------
# 时长补录：校内 / 校外
# ---------------------------------------------------------------------------


@app.post("/api/topup/internal")
async def api_topup_internal(
    name: str = Form(...),
    class_name: str = Form(...),
    grade: str = Form(...),
    academic_year: str = Form(...),
    activity_name: str = Form(...),
    duration: str = Form(...),
    wechat: str = Form(...),
    note: str = Form(""),
    proof: list[UploadFile] = File(...), # 改为列表支持多图
) -> Dict[str, Any]:
    """
    校内时长补录：
    - 证明文件：proof（图片列表，最多9张）
    """
    name = (name or "").strip()
    class_name = (class_name or "").strip()
    grade = (grade or "").strip()
    academic_year = (academic_year or "").strip()
    activity_name = (activity_name or "").strip()
    wechat = (wechat or "").strip()
    note = (note or "").strip()

    if not name or not class_name or not grade or not academic_year or not activity_name:
        raise HTTPException(status_code=400, detail={"ok": False, "message": "请完善必填字段", "data": []})
    if not wechat:
        raise HTTPException(status_code=400, detail={"ok": False, "message": "请填写微信号联系方式", "data": []})

    duration_i = _parse_positive_int("时长", duration)

    if not proof or not proof[0].filename:
        raise HTTPException(status_code=400, detail={"ok": False, "message": "请上传证明图片", "data": []})
    
    if len(proof) > 9:
        raise HTTPException(status_code=400, detail={"ok": False, "message": "证明图片最多上传9张", "data": []})

    proof_dir = cfg.TOPUP_PROOF_INTERNAL_DIR
    os.makedirs(proof_dir, exist_ok=True)
    
    proof_rel_paths = []
    max_bytes = getattr(cfg, "MAX_CONTENT_LENGTH", 16 * 1024 * 1024)

    for p in proof:
        if not _proof_allowed(p.filename or ""):
            continue
        
        safe = _secure_filename(p.filename or "file")
        unique_name = f"{uuid.uuid4().hex}_{safe}"
        dest_abs = os.path.join(proof_dir, unique_name)

        body = await p.read()
        if len(body) > max_bytes:
            continue

        with open(dest_abs, "wb") as f:
            f.write(body)

        try:
            rel = os.path.relpath(dest_abs, cfg.BASE_DIR)
            proof_rel_paths.append(rel)
        except ValueError:
            proof_rel_paths.append(dest_abs)

    if not proof_rel_paths:
        raise HTTPException(status_code=400, detail={"ok": False, "message": "没有有效的图片被上传", "data": []})

    # 多图路径用逗号分隔存储
    proof_rel_str = ",".join(proof_rel_paths)

    excel_service.append_topup_row(
        cfg.TOPUP_EXCEL_PATH,
        topup_type="校内",
        name=name,
        clazz=class_name,
        grade=grade,
        academic_year=academic_year,
        activity_name=activity_name,
        duration=duration_i,
        item_count=None,
        description="",
        proof_relative_path=proof_rel_str,
        photo_relative_path="",
        wechat=wechat,
        note=note,
    )

    return {"ok": True, "message": "已上传，青协办公室会定期检查并录入，如证明不符合规定会由微信通知"}


@app.post("/api/topup/external")
async def api_topup_external(
    name: str = Form(...),
    class_name: str = Form(...),
    grade: str = Form(...),
    academic_year: str = Form(...),
    activity_name: str = Form(...),
    duration: str = Form(...),
    item_count: str = Form(...),
    description: str = Form(...),
    wechat: str = Form(...),
    note: str = Form(""),
    proof: list[UploadFile] = File(...), # 多图支持
    photo: list[UploadFile] = File(...), # 多图支持
) -> Dict[str, Any]:
    """
    校外时长补录：
    - 证明文件：proof（图片列表）
    - 本人活动照片：photo（图片列表）
    """
    name = (name or "").strip()
    class_name = (class_name or "").strip()
    grade = (grade or "").strip()
    academic_year = (academic_year or "").strip()
    activity_name = (activity_name or "").strip()
    wechat = (wechat or "").strip()
    note = (note or "").strip()
    description = (description or "").strip()

    if not name or not class_name or not grade or not academic_year or not activity_name:
        raise HTTPException(status_code=400, detail={"ok": False, "message": "请完善必填字段", "data": []})
    if not wechat:
        raise HTTPException(status_code=400, detail={"ok": False, "message": "请填写微信号联系方式", "data": []})
    if not description:
        raise HTTPException(status_code=400, detail={"ok": False, "message": "请填写活动描述", "data": []})

    duration_i = _parse_positive_int("时长", duration)
    item_count_i = _parse_positive_int("活动项数", item_count)

    if not proof or not proof[0].filename:
        raise HTTPException(status_code=400, detail={"ok": False, "message": "请上传校外证明图片", "data": []})
    if not photo or not photo[0].filename:
        raise HTTPException(status_code=400, detail={"ok": False, "message": "请上传本人活动照片", "data": []})
    
    if len(proof) > 9 or len(photo) > 9:
        raise HTTPException(status_code=400, detail={"ok": False, "message": "证明或照片最多上传9张", "data": []})

    proof_dir = cfg.TOPUP_PROOF_EXTERNAL_DIR
    os.makedirs(proof_dir, exist_ok=True)
    max_bytes = getattr(cfg, "MAX_CONTENT_LENGTH", 16 * 1024 * 1024)

    # 保存证明文件列表
    proof_rel_paths = []
    for p in proof:
        if not _proof_allowed(p.filename or ""): continue
        safe = _secure_filename(p.filename or "file")
        unique = f"{uuid.uuid4().hex}_{safe}"
        dest = os.path.join(proof_dir, unique)
        content = await p.read()
        if len(content) > max_bytes: continue
        with open(dest, "wb") as f: f.write(content)
        try:
            rel = os.path.relpath(dest, cfg.BASE_DIR)
            proof_rel_paths.append(rel)
        except ValueError:
            proof_rel_paths.append(dest)

    # 保存照片文件列表
    photo_rel_paths = []
    for p in photo:
        if not _proof_allowed(p.filename or ""): continue
        safe = _secure_filename(p.filename or "file")
        unique = f"{uuid.uuid4().hex}_{safe}"
        dest = os.path.join(proof_dir, unique)
        content = await p.read()
        if len(content) > max_bytes: continue
        with open(dest, "wb") as f: f.write(content)
        try:
            rel = os.path.relpath(dest, cfg.BASE_DIR)
            photo_rel_paths.append(rel)
        except ValueError:
            photo_rel_paths.append(dest)

    if not proof_rel_paths or not photo_rel_paths:
        raise HTTPException(status_code=400, detail={"ok": False, "message": "没有有效的图片被上传", "data": []})

    excel_service.append_topup_row(
        cfg.TOPUP_EXCEL_PATH,
        topup_type="校外",
        name=name,
        clazz=class_name,
        grade=grade,
        academic_year=academic_year,
        activity_name=activity_name,
        duration=duration_i,
        item_count=item_count_i,
        description=description,
        proof_relative_path=",".join(proof_rel_paths),
        photo_relative_path=",".join(photo_rel_paths),
        wechat=wechat,
        note=note,
    )

    return {"ok": True, "message": "已上传，青协办公室会定期检查并录入，如证明不符合规定会由微信通知"}


# 便于直接 python api.py 启动（开发）
if __name__ == "__main__":
    import uvicorn

    uvicorn.run(
        "api:app",
        host="0.0.0.0",
        port=int(os.environ.get("PORT", "8000")),
        reload=False,
    )
