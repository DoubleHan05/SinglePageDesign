#!/usr/bin/env python3
"""Flask 服务：拖拽上传 Excel → 预览 HTML → 生成并下载 PNG

注意：Session 存在进程内存中，多 worker 部署时会串。
生产环境请使用单 worker 运行（如 `gunicorn -w 1 server:app`），
或后续接入 Redis / 文件缓存后再横向扩展。
"""
import os
import re
import threading
import time
import uuid
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from flask import Flask, abort, jsonify, request, send_file, send_from_directory
from werkzeug.utils import secure_filename

from generate import TEMPLATE_PATH, html_to_png, read_excel, render

HERE = Path(__file__).parent
LOGO_DIR = HERE / "logo"
UPLOAD_DIR = HERE / "output" / "_web"
UPLOAD_DIR.mkdir(parents=True, exist_ok=True)

# 上传大小上限（20 MB），避免大文件耗内存
MAX_UPLOAD_BYTES = 20 * 1024 * 1024
# Session 保留时长
SESSION_TTL_SEC = 3600
# 后台清理线程周期
CLEANUP_INTERVAL_SEC = 300

app = Flask(__name__, static_folder=None)
app.config["MAX_CONTENT_LENGTH"] = MAX_UPLOAD_BYTES

# 内存中记录已生成的 session
# sid -> {"html_path": Path, "excel_path": Path, "parsed_data": list, "ts": float}
SESSIONS: Dict[str, Dict[str, Any]] = {}
SESSIONS_LOCK = threading.Lock()


# ────────────────────────────── 工具函数 ──────────────────────────────

def _is_within(base: Path, target: Path) -> bool:
    """判断 target 是否在 base 目录下（防路径穿越，同时支持相同前缀的兄弟目录）。"""
    try:
        base_r = base.resolve()
        target_r = target.resolve()
        return os.path.commonpath([str(base_r), str(target_r)]) == str(base_r)
    except (ValueError, OSError):
        return False


def _cleanup_expired() -> None:
    now = time.time()
    with SESSIONS_LOCK:
        stale = [sid for sid, s in SESSIONS.items() if now - s["ts"] > SESSION_TTL_SEC]
        for sid in stale:
            SESSIONS.pop(sid, None)
    # 磁盘上的旧 session 目录也一并清理
    if UPLOAD_DIR.exists():
        for child in UPLOAD_DIR.iterdir():
            try:
                if child.is_dir() and now - child.stat().st_mtime > SESSION_TTL_SEC:
                    _rmtree(child)
            except OSError:
                pass


def _rmtree(p: Path) -> None:
    """轻量递归删除，忽略权限错误。"""
    for child in p.iterdir() if p.is_dir() else []:
        if child.is_dir():
            _rmtree(child)
        else:
            try: child.unlink()
            except OSError: pass
    try: p.rmdir()
    except OSError: pass


def _start_cleanup_thread() -> None:
    def _loop():
        while True:
            time.sleep(CLEANUP_INTERVAL_SEC)
            try:
                _cleanup_expired()
            except Exception as e:
                print(f"cleanup error: {e}")
    t = threading.Thread(target=_loop, name="session-cleanup", daemon=True)
    t.start()


def _model_from_data(data: List[Tuple[str, str]]) -> str:
    for k, v in data:
        if k.strip() == "型号" and v:
            return v.strip()
    return ""


def _rewrite_html_for_preview(html: str, sid: str) -> str:
    """把模板里的相对图片路径改写为服务器可访问的 URL。"""
    html = html.replace('src="../logo/', 'src="/logo/')

    def _sub(m: "re.Match[str]") -> str:
        src = m.group(1)
        if src.startswith(("http://", "https://", "data:", "/")):
            return m.group(0)
        return f'src="/session/{sid}/asset/{src}"'

    return re.sub(r'src="([^"]+)"', _sub, html)


def _rewrite_html_for_offline(html: str, sid: str, session_base: Path) -> str:
    """把 /logo/ 和 /session/<sid>/asset/ 前缀改为 file:// URI，供离线渲染。"""
    html = html.replace('src="/logo/', f'src="{LOGO_DIR.as_uri()}/')
    prefix = f"/session/{sid}/asset/"

    def _sub(m: "re.Match[str]") -> str:
        src = m.group(1)
        if src.startswith(prefix):
            return f'src="{(session_base / src[len(prefix):]).as_uri()}"'
        return m.group(0)

    return re.sub(r'src="([^"]+)"', _sub, html)


# ────────────────────────────── 路由 ──────────────────────────────

@app.get("/")
def index():
    return (HERE / "web" / "index.html").read_text(encoding="utf-8")


@app.get("/logo/<path:name>")
def logo(name):
    return send_from_directory(LOGO_DIR, name)


@app.get("/web/<path:name>")
def web_asset(name):
    return send_from_directory(HERE / "web", name)


@app.get("/template.xlsx")
def download_template():
    tpl = HERE / "samples" / "产品单页模板.xlsx"
    if not tpl.exists():
        return jsonify({"error": "模板文件不存在"}), 404
    return send_file(tpl, as_attachment=True, download_name="产品单页模板.xlsx")


@app.get("/session/<sid>/asset/<path:name>")
def session_asset(sid, name):
    """预览 HTML 里引用的图片（相对路径），从上传目录里取。"""
    if sid not in SESSIONS:
        abort(404)
    base = SESSIONS[sid]["excel_path"].parent
    target = base / name
    if not _is_within(base, target) or not target.exists():
        abort(404)
    return send_file(target)


@app.post("/upload")
def upload():
    if "file" not in request.files:
        return jsonify({"error": "未收到文件"}), 400
    f = request.files["file"]

    # 先用原始文件名判扩展名（secure_filename 会剥离中文，可能把扩展名一起丢掉）
    original = f.filename or ""
    ext = os.path.splitext(original)[1].lower()
    if ext not in (".xlsx", ".xlsm"):
        return jsonify({"error": "请上传 .xlsx 文件"}), 400

    # 再清洗一份用于落盘。清洗后无有效名或缺扩展名时兜底
    cleaned = secure_filename(original)
    stem = os.path.splitext(cleaned)[0] or "upload"
    safe_name = f"{stem}{ext}"

    sid = uuid.uuid4().hex[:12]
    sess_dir = UPLOAD_DIR / sid
    sess_dir.mkdir(parents=True, exist_ok=True)
    excel_path = sess_dir / safe_name
    f.save(excel_path)

    try:
        data = read_excel(excel_path, image_out_dir=sess_dir / "_embed")
        # 绝对路径的嵌入图片改为相对 session 目录的路径
        fixed: List[Tuple[str, str]] = []
        for k, v in data:
            if v and Path(v).is_absolute():
                try:
                    v = str(Path(v).resolve().relative_to(sess_dir.resolve()))
                except ValueError:
                    pass
            fixed.append((k, v))
        data = fixed
    except Exception as e:
        return jsonify({"error": f"解析 Excel 失败：{e}"}), 500

    with SESSIONS_LOCK:
        SESSIONS[sid] = {
            "session_dir": sess_dir,
            "excel_path": excel_path,
            "parsed_data": data,       # 缓存解析结果，避免下载时重复解析
            "ts": time.time(),
        }
    return jsonify({"sid": sid, "preview_url": f"/preview/{sid}"})


def _render_preview_html(sid: str, theme: str) -> Optional[Path]:
    """按主题渲染并落盘一份预览 HTML，返回文件路径。"""
    if sid not in SESSIONS:
        return None
    sess = SESSIONS[sid]
    template = TEMPLATE_PATH.read_text(encoding="utf-8")
    html = render(sess["parsed_data"], template, theme=theme)
    html = _rewrite_html_for_preview(html, sid)
    path = sess["session_dir"] / f"page.{theme}.html"
    path.write_text(html, encoding="utf-8")
    return path


@app.get("/render/<sid>")
def render_page(sid):
    """返回带 html2canvas 自动截图逻辑的产品页，由用户浏览器完成截图并下载。"""
    if sid not in SESSIONS:
        abort(404)
    sess = SESSIONS[sid]
    theme = request.args.get("theme", "blue")
    if theme not in ("blue", "light"):
        theme = "blue"

    model = _model_from_data(sess.get("parsed_data", [])) or sess["excel_path"].stem
    filename = f"{model}-产品单页-{datetime.now().strftime('%Y%m%d')}.png"

    path = _render_preview_html(sid, theme)
    if path is None:
        abort(404)
    html = path.read_text(encoding="utf-8")

    # 在 </body> 前注入自动截图脚本
    inject = f"""
<script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script>
<script>
window.addEventListener('load', function() {{
  var page = document.querySelector('.page') || document.body;
  html2canvas(page, {{
    scale: 2,
    useCORS: true,
    allowTaint: true,
    backgroundColor: null,
    width: page.scrollWidth,
    height: page.scrollHeight,
    windowWidth: page.scrollWidth,
    windowHeight: page.scrollHeight,
  }}).then(function(canvas) {{
    var a = document.createElement('a');
    a.download = {repr(filename)};
    a.href = canvas.toDataURL('image/png');
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(function() {{ window.close(); }}, 500);
  }});
}});
</script>"""
    html = html.replace("</body>", inject + "\n</body>")
    return html


@app.get("/preview/<sid>")
def preview(sid):
    if sid not in SESSIONS:
        abort(404)
    theme = request.args.get("theme", "blue")
    if theme not in ("blue", "light"):
        theme = "blue"
    path = _render_preview_html(sid, theme)
    if path is None:
        abort(404)
    return path.read_text(encoding="utf-8")


@app.get("/download-image/<sid>")
def download_image(sid):
    if sid not in SESSIONS:
        abort(404)
    sess = SESSIONS[sid]
    theme = request.args.get("theme", "blue")
    if theme not in ("blue", "light"):
        theme = "blue"

    # 按主题重新渲染一份预览 HTML
    preview_path = _render_preview_html(sid, theme)
    if preview_path is None:
        abort(404)
    png_path = preview_path.with_suffix(".png")

    # 渲染时把 URL 换成 file://，避免依赖 HTTP
    offline_html = _rewrite_html_for_offline(
        preview_path.read_text(encoding="utf-8"),
        sid,
        sess["excel_path"].parent,
    )
    offline_path = preview_path.with_name(f"page.{theme}.offline.html")
    offline_path.write_text(offline_html, encoding="utf-8")

    if not html_to_png(offline_path, png_path):
        return jsonify({"error": "图片生成失败，请检查是否安装了 weasyprint 或 playwright"}), 500

    # 型号从缓存里取，避免二次解析 xlsx
    model = _model_from_data(sess.get("parsed_data", [])) or sess["excel_path"].stem
    filename = f"{model}-产品单页-{datetime.now().strftime('%Y%m%d')}.png"
    return send_file(png_path, mimetype="image/png",
                     as_attachment=True,
                     download_name=filename)


@app.errorhandler(413)
def request_too_large(e):
    mb = MAX_UPLOAD_BYTES // (1024 * 1024)
    return jsonify({"error": f"文件超过 {mb}MB 上限"}), 413


_start_cleanup_thread()

if __name__ == "__main__":
    print("🚀 服务已启动：http://127.0.0.1:5001")
    app.run(host="127.0.0.1", port=5001, debug=False)
