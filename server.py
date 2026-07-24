#!/usr/bin/env python3
"""Flask 服务：拖拽上传 Excel → 预览 HTML → 生成并下载 PNG"""
import io
import time
import uuid
from pathlib import Path

from flask import Flask, request, jsonify, send_file, abort, send_from_directory

from generate import read_excel, render, html_to_png, TEMPLATE_PATH

HERE = Path(__file__).parent
LOGO_DIR = HERE / "logo"
UPLOAD_DIR = HERE / "output" / "_web"
UPLOAD_DIR.mkdir(parents=True, exist_ok=True)

app = Flask(__name__, static_folder=None)

# 内存中记录已生成的 session
SESSIONS = {}  # sid -> {"html_path": Path, "excel_path": Path, "ts": float}


def _cleanup(max_age_sec=3600):
    now = time.time()
    for sid in list(SESSIONS.keys()):
        if now - SESSIONS[sid]["ts"] > max_age_sec:
            SESSIONS.pop(sid, None)


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
    target = (base / name).resolve()
    if not str(target).startswith(str(base.resolve())):
        abort(403)
    if not target.exists():
        abort(404)
    return send_file(target)


@app.post("/upload")
def upload():
    _cleanup()
    if "file" not in request.files:
        return jsonify({"error": "未收到文件"}), 400
    f = request.files["file"]
    if not f.filename.lower().endswith((".xlsx", ".xlsm")):
        return jsonify({"error": "请上传 .xlsx 文件"}), 400

    sid = uuid.uuid4().hex[:12]
    sess_dir = UPLOAD_DIR / sid
    sess_dir.mkdir(parents=True, exist_ok=True)
    excel_path = sess_dir / f.filename
    f.save(excel_path)

    try:
        # 嵌入图片解压到 session 目录下的 _embed
        data = read_excel(excel_path, image_out_dir=sess_dir / "_embed")
        # 将绝对路径的嵌入图片改为相对 session 目录的路径
        fixed = []
        for k, v in data:
            if v and Path(v).is_absolute():
                try:
                    rel = str(Path(v).resolve().relative_to(sess_dir.resolve()))
                    v = rel
                except ValueError:
                    pass
            fixed.append((k, v))
        data = fixed
        template = TEMPLATE_PATH.read_text(encoding="utf-8")
        html = render(data, template)
    except Exception as e:
        return jsonify({"error": f"解析 Excel 失败：{e}"}), 500

    # 把模板里对 ../logo/ 与相对资源路径改写成服务器可访问的 URL
    html = html.replace('src="../logo/', 'src="/logo/')
    # 相对图片路径（如 assets/printer.svg 或 _embed/xx.png）转成 session 资源 URL
    import re
    def _rewrite(m):
        src = m.group(1)
        if src.startswith(("http://", "https://", "data:", "/")):
            return m.group(0)
        return f'src="/session/{sid}/asset/{src}"'
    html = re.sub(r'src="([^"]+)"', _rewrite, html)

    html_path = sess_dir / "page.html"
    html_path.write_text(html, encoding="utf-8")

    SESSIONS[sid] = {"html_path": html_path, "excel_path": excel_path, "ts": time.time()}
    return jsonify({"sid": sid, "preview_url": f"/preview/{sid}"})


@app.get("/preview/<sid>")
def preview(sid):
    if sid not in SESSIONS:
        abort(404)
    return SESSIONS[sid]["html_path"].read_text(encoding="utf-8")


@app.get("/download-image/<sid>")
def download_image(sid):
    if sid not in SESSIONS:
        abort(404)
    sess = SESSIONS[sid]
    html_path = sess["html_path"]
    png_path = html_path.with_suffix(".png")

    # 生成图片时，让引用的资源用文件系统路径（避免依赖 HTTP）
    # 为此重新渲染一份"离线 HTML"
    offline_html = html_path.read_text(encoding="utf-8")
    offline_html = offline_html.replace('src="/logo/', f'src="{(LOGO_DIR).as_uri()}/')
    base = sess["excel_path"].parent
    import re
    def _abs(m):
        src = m.group(1)
        if src.startswith((f"/session/{sid}/asset/",)):
            rel = src[len(f"/session/{sid}/asset/"):]
            return f'src="{(base / rel).as_uri()}"'
        return m.group(0)
    offline_html = re.sub(r'src="([^"]+)"', _abs, offline_html)
    offline_path = html_path.with_name("page.offline.html")
    offline_path.write_text(offline_html, encoding="utf-8")

    ok = html_to_png(offline_path, png_path)
    if not ok:
        return jsonify({"error": "图片生成失败，请检查是否安装了 weasyprint 或 playwright"}), 500

    # 从 Excel 里取型号；找不到则退回到文件名
    from datetime import datetime
    model = ""
    try:
        for k, v in read_excel(sess["excel_path"]):
            if k.strip() == "型号" and v:
                model = v.strip()
                break
    except Exception:
        pass
    if not model:
        model = sess["excel_path"].stem
    date_str = datetime.now().strftime("%Y%m%d")
    filename = f"{model}-产品单页-{date_str}.png"
    return send_file(png_path, mimetype="image/png",
                     as_attachment=True,
                     download_name=filename)


if __name__ == "__main__":
    print("🚀 服务已启动：http://127.0.0.1:5001")
    app.run(host="127.0.0.1", port=5001, debug=False)
