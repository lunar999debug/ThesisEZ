"""ThesisEZ Web Demo —— 在线版一键论文格式化。

部署形态：用户上传「输入.docx」+ 填封面信息 → 后端跑 pipeline → 返回 ThesisEZ.docx

关键设计：
  - 每个请求开独立工作区 /tmp/jobs/<uuid>/，并发安全
  - 跑完立即清掉用户数据，不存盘
  - 单 IP 每天限次（防恶意刷 API key）
  - 上传文件大小硬上限 5 MB（够装一篇带图论文）
"""
from __future__ import annotations

import os
import shutil
import subprocess
import sys
import uuid
import threading
import time
from datetime import datetime, timezone
from pathlib import Path

import yaml
from flask import (
    Flask,
    request,
    send_file,
    render_template,
    jsonify,
    abort,
)

# ──────────────────── 路径常量 ──────────────────────
HERE = Path(__file__).resolve().parent
SOURCE_DIR = HERE / "ThesisEZ"           # 原始 ThesisEZ 代码目录
TEMPLATE_SRC = SOURCE_DIR / "template"   # Word 模板源目录
JOBS_ROOT = Path("/tmp/jobs")
JOBS_ROOT.mkdir(parents=True, exist_ok=True)

# ──────────────────── 限流配置 ──────────────────────
MAX_UPLOAD_BYTES = 5 * 1024 * 1024  # 5 MB
DAILY_QUOTA_PER_IP = 5              # 每个 IP 每天最多 5 次
JOB_TIMEOUT_SEC = 240               # 单次任务最长 4 分钟

# 简单的内存限流表：{ip: {"date": "YYYY-MM-DD", "count": int}}
_quota_lock = threading.Lock()
_quota: dict[str, dict] = {}


def _check_quota(ip: str) -> tuple[bool, int]:
    """检查额度但不增计数 —— 返回 (是否放行, 已用次数)。"""
    today = datetime.now(timezone.utc).strftime("%Y-%m-%d")
    with _quota_lock:
        entry = _quota.get(ip)
        if entry is None or entry["date"] != today:
            _quota[ip] = {"date": today, "count": 0}
            entry = _quota[ip]
        if entry["count"] >= DAILY_QUOTA_PER_IP:
            return False, entry["count"]
        return True, entry["count"]


def _consume_quota(ip: str) -> None:
    """成功跑完后才计数。这样失败请求不会扣额度。"""
    today = datetime.now(timezone.utc).strftime("%Y-%m-%d")
    with _quota_lock:
        entry = _quota.get(ip)
        if entry is None or entry["date"] != today:
            _quota[ip] = {"date": today, "count": 1}
        else:
            entry["count"] += 1


# ──────────────────── Flask 应用 ──────────────────────
app = Flask(__name__, template_folder="templates", static_folder="static")
app.config["MAX_CONTENT_LENGTH"] = MAX_UPLOAD_BYTES


@app.route("/")
def index() -> str:
    return render_template("index.html", quota=DAILY_QUOTA_PER_IP)


@app.route("/health")
def health() -> tuple[str, int]:
    return "ok", 200


@app.errorhandler(413)
def too_large(_e):
    return (
        jsonify(error=f"文件太大，上限 {MAX_UPLOAD_BYTES // 1024 // 1024} MB"),
        413,
    )


@app.route("/api/render", methods=["POST"])
def api_render():
    """核心接口：上传 docx + 封面信息 → 返回渲染后的 docx。"""
    # 1. 限流
    ip = request.headers.get("X-Forwarded-For", request.remote_addr or "unknown")
    ip = ip.split(",")[0].strip()
    ok, remaining = _check_quota(ip)
    if not ok:
        return (
            jsonify(error=f"今天的演示额度用完了（每 IP 每天 {DAILY_QUOTA_PER_IP} 次），明天再来"),
            429,
        )

    # 2. 校验输入
    if "draft" not in request.files:
        return jsonify(error="缺少草稿文件"), 400
    draft = request.files["draft"]
    if not draft.filename or not draft.filename.lower().endswith(".docx"):
        return jsonify(error="请上传 .docx 文件"), 400

    # 封面信息（可选，没填就用占位）
    cover_info = {
        "title_zh": request.form.get("title_zh", "").strip() or "演示论文标题",
        "title_en": request.form.get("title_en", "").strip() or "Demo Thesis Title",
        "author": request.form.get("author", "").strip() or "演示作者",
        "student_id": request.form.get("student_id", "").strip() or "0000000000",
        "advisor": request.form.get("advisor", "").strip() or "演示导师",
        "major": request.form.get("major", "").strip() or "演示专业",
        "school": request.form.get("school", "").strip() or "演示学院",
        "submit_date": request.form.get("submit_date", "").strip() or "2026-06",
    }

    # 3. 起独立工作区
    job_id = uuid.uuid4().hex[:12]
    job_dir = JOBS_ROOT / job_id
    try:
        _prepare_job_dir(job_dir, draft, cover_info)
    except Exception as e:
        _cleanup_job(job_dir)
        return jsonify(error=f"工作区准备失败：{e}"), 500

    # 4. 跑 pipeline
    try:
        out_path = _run_pipeline(job_dir)
    except subprocess.TimeoutExpired:
        _cleanup_job(job_dir)
        return jsonify(error=f"渲染超时（>{JOB_TIMEOUT_SEC}s），可能论文太长，请缩短后重试"), 504
    except RuntimeError as e:
        _cleanup_job(job_dir)
        return jsonify(error=f"渲染失败：{e}"), 500

    # 5. 流回结果，再清场
    try:
        # 把文件读进内存，立即删工作区
        data = out_path.read_bytes()
    finally:
        _cleanup_job(job_dir)

    # 成功了才扣额度
    _consume_quota(ip)

    # 用 BytesIO 流式返回
    from io import BytesIO
    buf = BytesIO(data)
    buf.seek(0)
    return send_file(
        buf,
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        as_attachment=True,
        download_name="ThesisEZ.docx",
    )


# ──────────────────── pipeline 实现 ──────────────────
def _prepare_job_dir(job_dir: Path, draft_file, cover_info: dict) -> None:
    """把每次请求所需的全部文件复制到独立工作区。"""
    job_dir.mkdir(parents=True, exist_ok=False)

    # 1) 拷代码（aichuli.py + render.py）
    for fname in ("aichuli.py", "render.py"):
        shutil.copy(SOURCE_DIR / fname, job_dir / fname)

    # 2) 拷模板目录
    shutil.copytree(TEMPLATE_SRC, job_dir / "template")

    # 3) 写 info.yaml（用表单数据）
    (job_dir / "info.yaml").write_text(
        yaml.safe_dump(cover_info, allow_unicode=True, sort_keys=False),
        encoding="utf-8",
    )

    # 4) 存用户上传的草稿
    draft_file.save(job_dir / "输入.docx")

    # 5) 建空的图表/ 目录（避免代码找不到）
    (job_dir / "图表").mkdir(exist_ok=True)


def _run_pipeline(job_dir: Path) -> Path:
    """在 job_dir 里跑 aichuli + render，返回输出 docx 路径。"""
    env = os.environ.copy()
    env["PYTHONIOENCODING"] = "utf-8"
    env["PYTHONUTF8"] = "1"
    # DEEPSEEK_API_KEY 已在容器环境变量里，直接继承

    # 步骤 1：AI 解析（直接 import 跑会污染全局 cwd，所以走子进程）
    r1 = subprocess.run(
        [sys.executable, "-c",
         "import sys; sys.path.insert(0, '.'); import aichuli; aichuli.process()"],
        cwd=str(job_dir),
        env=env,
        capture_output=True,
        text=True,
        timeout=JOB_TIMEOUT_SEC,
    )
    if r1.returncode != 0:
        raise RuntimeError(
            f"AI 解析失败：{(r1.stderr or r1.stdout)[-500:]}"
        )

    # 步骤 2：渲染
    r2 = subprocess.run(
        [sys.executable, "render.py", "--no-bibtex"],
        cwd=str(job_dir),
        env=env,
        capture_output=True,
        text=True,
        timeout=JOB_TIMEOUT_SEC,
    )
    if r2.returncode != 0:
        raise RuntimeError(
            f"渲染引擎失败：{(r2.stderr or r2.stdout)[-500:]}"
        )

    out = job_dir / "ThesisEZ.docx"
    if not out.is_file():
        raise RuntimeError("没有生成 ThesisEZ.docx")
    return out


def _cleanup_job(job_dir: Path) -> None:
    if job_dir.exists():
        shutil.rmtree(job_dir, ignore_errors=True)


# ──────────────────── 启动 ──────────────────────
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8000))
    app.run(host="0.0.0.0", port=port, debug=False)
