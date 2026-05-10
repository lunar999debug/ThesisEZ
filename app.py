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
MAX_UPLOAD_BYTES = 20 * 1024 * 1024  # 20 MB —— 草稿 + 多张图表一起估算
ALLOWED_FIGURE_EXTS = {".png", ".jpg", ".jpeg", ".pdf", ".csv"}
MAX_FIGURES = 30  # 单次上传图表最多 30 个文件
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


@app.route("/sample")
def sample_draft():
    """提供示例草稿下载 —— 让没论文的评审也能马上试用。"""
    sample_path = SOURCE_DIR / "输入.docx"
    if not sample_path.is_file():
        abort(404)
    return send_file(
        sample_path,
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        as_attachment=True,
        download_name="示例草稿.docx",
    )


# 启动时预生成示例图表 zip（只打一次，请求时直接吞）
import zipfile
from io import BytesIO as _BytesIO

def _build_sample_figures_zip() -> bytes:
    fig_dir = SOURCE_DIR / "图表"
    if not fig_dir.is_dir():
        return b""
    buf = _BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for p in sorted(fig_dir.iterdir()):
            if p.is_file():
                # 压在压缩包里保留 「图表/xxx」 路径，让用户解压后原样拖进项目
                zf.write(p, arcname=f"图表/{p.name}")
    return buf.getvalue()

_SAMPLE_FIGURES_ZIP = _build_sample_figures_zip()


@app.route("/sample-figures")
def sample_figures():
    """提供示例图表压缩包下载。"""
    if not _SAMPLE_FIGURES_ZIP:
        abort(404)
    return send_file(
        _BytesIO(_SAMPLE_FIGURES_ZIP),
        mimetype="application/zip",
        as_attachment=True,
        download_name="示例图表.zip",
    )


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

    # 可选：图表文件（多个）
    figure_files = request.files.getlist("figures")
    figure_files = [f for f in figure_files if f and f.filename]
    if len(figure_files) > MAX_FIGURES:
        return jsonify(error=f"最多上传 {MAX_FIGURES} 个图表文件"), 400
    for f in figure_files:
        ext = Path(f.filename).suffix.lower()
        if ext not in ALLOWED_FIGURE_EXTS:
            return (
                jsonify(error=f"不支持的图表格式：{f.filename}（只接受 png/jpg/jpeg/pdf/csv）"),
                400,
            )

    # 封面信息（可选，没填就用占位）。
    # ☆ 字段名必须严格对齐 aichuli.py 中 SETUP_TEMPLATE 使用的中文 key。
    cover_info = {
        "姓名":          request.form.get("name_zh", "").strip() or "演示同学",
        "Author":        request.form.get("name_en", "").strip() or "Demo Student",
        "学号":          request.form.get("student_id", "").strip() or "520000000000",
        "导师":          request.form.get("advisor_zh", "").strip() or "演示导师",
        "Supervisor":    request.form.get("advisor_en", "").strip() or "Demo Advisor",
        "中文学院名":    request.form.get("school_zh", "").strip() or "演示学院",
        "英文学院名":    request.form.get("school_en", "").strip() or "Demo School",
        "专业名称":      request.form.get("major", "").strip() or "演示专业",
        "申请学位层次":  request.form.get("degree", "").strip() or "学士",
        "年份":          request.form.get("year", "").strip() or "2026",
        "中文论文标题":  request.form.get("title_zh", "").strip() or "演示论文标题",
        "英文论文标题":  request.form.get("title_en", "").strip() or "Demo Thesis Title",
    }

    # 3. 起独立工作区
    job_id = uuid.uuid4().hex[:12]
    job_dir = JOBS_ROOT / job_id
    try:
        _prepare_job_dir(job_dir, draft, cover_info, figure_files)
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
def _prepare_job_dir(job_dir: Path, draft_file, cover_info: dict, figure_files: list) -> None:
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

    # 5) 图表目录——创建后把用户上传的图表文件全部起名存进去。
    fig_dir = job_dir / "图表"
    fig_dir.mkdir(exist_ok=True)
    for f in figure_files:
        # 只取原始文件名的 basename，防止路径穿越
        safe_name = Path(f.filename).name
        f.save(fig_dir / safe_name)


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
