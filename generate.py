#!/usr/bin/env python3
"""
产品单页生成工具

用法:
    python generate.py <excel_path> [--out-dir <dir>] [--no-image]

Excel 格式（两列，无表头强制）：
    第一列 = 参数名（key）
    第二列 = 参数值（value）

支持的 key（大小写敏感，前后空格自动去除）：
    品牌LOGO字母, 品牌名称, 公司名, 股票代码, 型号, 表格标题,
    产品图, 细节图,
    特点1, 特点2, ...           （按序号顺序渲染为橙色菱形项目符号列表）
    场景1, 场景2, ...           （格式："仓|仓储"，圆内文字|下方说明）
    参数.<参数名>              （渲染到右侧参数表中；未填写用 "/"）

未在识别列表中的、以 "参数." 开头的行都会被当作参数表一行。
其余 key 会直接替换到模板占位符。
"""
import argparse
import os
import re
import sys
from html import escape
from pathlib import Path

import openpyxl

HERE = Path(__file__).parent
TEMPLATE_PATH = HERE / "templates" / "template.html"


def _extract_cell_embedded_images(xlsx_path: Path, ws_name: str, out_dir: Path):
    """
    解析 WPS / 新版 Excel 的"在单元格中放置图片"（DISPIMG 公式 + xl/cellimages.xml）。
    返回 {row_1based: [saved_path]}。
    """
    import zipfile, re as _re
    from xml.etree import ElementTree as ET

    row_map = {}
    try:
        with zipfile.ZipFile(xlsx_path) as z:
            names = z.namelist()
            if "xl/cellimages.xml" not in names:
                return row_map

            # 1) cellimages.xml: 图片名 -> 内部 embed id
            ci_xml = z.read("xl/cellimages.xml")
            # 2) cellimages.xml.rels: embed id -> media 文件路径
            rels_name = "xl/_rels/cellimages.xml.rels"
            if rels_name not in names:
                return row_map
            rels_xml = z.read(rels_name)

            ns = {
                "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
                "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
                "a":   "http://schemas.openxmlformats.org/drawingml/2006/main",
                "r":   "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
                "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
            }

            # embed_id -> media 路径
            rels_root = ET.fromstring(rels_xml)
            id_to_media = {}
            for rel in rels_root.findall("rel:Relationship", ns):
                rid = rel.get("Id"); target = rel.get("Target", "")
                if target.startswith("../"):
                    target = "xl/" + target[3:]
                elif not target.startswith("xl/"):
                    target = "xl/" + target
                id_to_media[rid] = target

            # cellimages.xml：picture name -> embed id
            ci_root = ET.fromstring(ci_xml)
            name_to_media = {}
            for pic in ci_root.iter("{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}pic"):
                nv = pic.find("xdr:nvPicPr/xdr:cNvPr", ns)
                blip = pic.find("xdr:blipFill/a:blip", ns)
                if nv is None or blip is None: continue
                pic_name = nv.get("name") or nv.get("descr")
                embed = blip.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed")
                if pic_name and embed and embed in id_to_media:
                    name_to_media[pic_name] = id_to_media[embed]

            # 找到工作表 sheet1.xml（简单起见默认第一张；ws_name 参数保留但暂不使用）
            sheet_xml = None
            for n in names:
                if n.startswith("xl/worksheets/sheet") and n.endswith(".xml"):
                    sheet_xml = n; break
            if not sheet_xml:
                return row_map

            root = ET.fromstring(z.read(sheet_xml))
            for c in root.iter("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}c"):
                ref = c.get("r", "")  # e.g. "B7"
                f = c.find("main:f", ns)
                if f is None or not f.text: continue
                m = _re.search(r'DISPIMG\(\s*"([^"]+)"', f.text)
                if not m: continue
                pic_name = m.group(1)
                media_rel = name_to_media.get(pic_name)
                if not media_rel or media_rel not in names: continue

                # 解析行号
                row_m = _re.search(r"(\d+)$", ref)
                if not row_m: continue
                row_1 = int(row_m.group(1))

                # 导出图片
                out_dir.mkdir(parents=True, exist_ok=True)
                ext = os.path.splitext(media_rel)[1] or ".png"
                fpath = out_dir / f"cell_{pic_name}{ext}"
                with z.open(media_rel) as src, open(fpath, "wb") as dst:
                    dst.write(src.read())
                row_map.setdefault(row_1, []).append(fpath)
    except Exception as e:
        print(f"⚠️  解析单元格嵌入图片失败：{e}")
    return row_map


def _extract_embedded_images(ws, out_dir: Path):
    """把 sheet 中的嵌入图片按行号提取出来，返回 {row_1based: [saved_path, ...]}."""
    out_dir.mkdir(parents=True, exist_ok=True)
    row_map = {}
    images = getattr(ws, "_images", []) or []
    for idx, img in enumerate(images):
        # openpyxl 的图片锚点通常是 OneCellAnchor / TwoCellAnchor
        row_1based = None
        try:
            anchor = img.anchor
            frm = getattr(anchor, "_from", None) or getattr(anchor, "from_", None)
            if frm is not None:
                row_1based = int(frm.row) + 1  # openpyxl 用 0-based
        except Exception:
            row_1based = None
        if row_1based is None:
            continue

        # 图片二进制
        try:
            blob = img._data() if callable(getattr(img, "_data", None)) else img.ref.read()
        except Exception:
            try:
                blob = img.ref.read()
            except Exception:
                continue

        # 后缀
        ext = ".png"
        fmt = (getattr(img, "format", "") or "").lower()
        if fmt in ("jpg", "jpeg"): ext = ".jpg"
        elif fmt == "gif": ext = ".gif"
        elif fmt == "bmp": ext = ".bmp"
        elif fmt == "webp": ext = ".webp"
        else:
            # 从 header 猜
            if blob[:3] == b"\xff\xd8\xff": ext = ".jpg"
            elif blob[:4] == b"GIF8": ext = ".gif"
            elif blob[:8].startswith(b"\x89PNG"): ext = ".png"

        fname = f"embed_{idx}_r{row_1based}{ext}"
        fpath = out_dir / fname
        with open(fpath, "wb") as f:
            f.write(blob)

        row_map.setdefault(row_1based, []).append(fpath)
    return row_map


def read_excel(path: Path, image_out_dir: Path = None):
    """
    读取 Excel。Sheet1 是产品数据，Sheet2 是参数表（无需加"参数."前缀）。
    若某一行第二列为空且该行嵌入了图片，则用图片路径作为值。
    image_out_dir：把嵌入图片解出到该目录（默认 excel 所在目录/_embed）。
    """
    wb = openpyxl.load_workbook(path, data_only=True)

    if image_out_dir is None:
        image_out_dir = path.parent / "_embed"

    data = []

    # ─── Sheet1：产品数据 ───
    ws1 = wb.worksheets[0]
    row_images_1 = _extract_embedded_images(ws1, image_out_dir)
    cell_images_1 = _extract_cell_embedded_images(path, ws1.title, image_out_dir)
    for r, lst in cell_images_1.items():
        row_images_1.setdefault(r, []).extend(lst)

    for row_idx, row in enumerate(ws1.iter_rows(values_only=True), start=1):
        if not row:
            continue
        key = row[0]
        val = row[1] if len(row) > 1 else None
        if key is None:
            continue
        key = str(key).strip()
        if not key:
            continue
        val = "" if val is None else str(val).strip()

        # 图片列（产品图 / 细节图N）优先使用嵌入图片
        if re.match(r"^(产品图|细节图\s*\d*)$", key):
            imgs = row_images_1.get(row_idx, [])
            if imgs and not val:
                val = str(imgs[0])

        data.append((key, val))

    # ─── Sheet2：参数表（如果存在） ───
    if len(wb.worksheets) >= 2:
        ws2 = wb.worksheets[1]
        # 跳过表头（第一行通常是"参数名称/参数值"）
        for row_idx, row in enumerate(ws2.iter_rows(values_only=True, min_row=2), start=2):
            if not row:
                continue
            key = row[0]
            val = row[1] if len(row) > 1 else None
            if key is None:
                continue
            key = str(key).strip()
            if not key:
                continue
            val = "" if val is None else str(val).strip()
            # 自动加上"参数."前缀，统一后续处理逻辑
            data.append((f"参数.{key}", val))

    return data


def build_features(items):
    if not items:
        return '<li style="list-style:none;color:#999">（未提供产品特点）</li>'
    return "\n".join(f"<li>{escape(x)}</li>" for x in items)


def build_scenes(items):
    if not items:
        return ""
    parts = []
    for raw in items:
        if "|" in raw:
            ch, label = raw.split("|", 1)
        else:
            ch, label = raw, raw
        parts.append(
            f'<div class="scene"><div class="circle">{escape(ch.strip())}</div>'
            f'<div class="label">{escape(label.strip())}</div></div>'
        )
    return "\n".join(parts)


def build_param_rows(rows):
    if not rows:
        return '<tr><td class="k">（无参数）</td><td class="v">/</td></tr>'
    out = []
    for k, v in rows:
        v = v if v else "/"
        out.append(
            f'<tr><td class="k">{escape(k)}</td><td class="v">{escape(v)}</td></tr>'
        )
    return "\n".join(out)


def render(data, template):
    features = []
    scenes = []
    params = []
    scalars = {
        "MODEL": "",
        "PRODUCT_IMG": "",
    }
    key_to_slot = {
        "型号": "MODEL",
        "产品图": "PRODUCT_IMG",
    }
    detail_imgs = []

    for key, val in data:
        if key in key_to_slot:
            scalars[key_to_slot[key]] = val
        elif re.match(r"^特点\s*\d+$", key):
            if val:
                features.append(val)
        elif re.match(r"^场景\s*\d+$", key):
            if val:
                scenes.append(val)
        elif re.match(r"^细节图\s*\d*$", key):
            if val:
                detail_imgs.append(val)
        elif key.startswith("参数."):
            pname = key[len("参数."):].strip()
            if pname:
                params.append((pname, val))
        else:
            # 未识别的 key，也当参数表处理
            params.append((key, val))

    html = template
    for k, v in scalars.items():
        html = html.replace("{{" + k + "}}", escape(v) if k != "PRODUCT_IMG" else v)
    html = html.replace("{{FEATURES}}", build_features(features))
    html = html.replace("{{SCENES}}", build_scenes(scenes))
    html = html.replace("{{PARAM_ROWS}}", build_param_rows(params))
    if not detail_imgs:
        detail_html = '<div class="product-img" style="color:#bbb;font-size:12px">（未提供细节图）</div>'
    else:
        detail_html = "\n".join(
            f'<div class="product-img"><img src="{src}" alt="detail"></div>'
            for src in detail_imgs
        )
    html = html.replace("{{DETAIL_IMGS}}", detail_html)
    return html


def html_to_png(html_path: Path, png_path: Path):
    """优先使用 Playwright（渲染最接近浏览器），退回到 WeasyPrint + PyMuPDF。"""
    # 尝试 Playwright
    try:
        from playwright.sync_api import sync_playwright  # type: ignore
        last_err = None
        # 依次尝试：系统 Chrome → 系统 Edge → 自带 chromium
        for launcher in [
            lambda p: p.chromium.launch(channel="chrome"),
            lambda p: p.chromium.launch(channel="msedge"),
            lambda p: p.chromium.launch(),
        ]:
            try:
                with sync_playwright() as p:
                    browser = launcher(p)
                    page = browser.new_page(viewport={"width": 1062, "height": 800},
                                            device_scale_factor=2)
                    page.goto(html_path.absolute().as_uri())
                    page.wait_for_load_state("networkidle")
                    page.screenshot(path=str(png_path), full_page=True)
                    browser.close()
                return True
            except Exception as e:
                last_err = e
                continue
        raise last_err or RuntimeError("no browser channel available")
    except Exception as e:
        print(f"ℹ️  Playwright 不可用（{e.__class__.__name__}），改用 WeasyPrint。")

    # 退回：WeasyPrint → PDF → PyMuPDF → PNG
    try:
        from weasyprint import HTML  # type: ignore
        import fitz  # type: ignore
        pdf_bytes = HTML(filename=str(html_path)).write_pdf()
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        # 用 tmp 路径生成再原子替换，避开某些文件系统的删除限制
        import tempfile as _tf
        _tmp_png = Path(_tf.mkstemp(suffix=".png")[1])
        # 拼接所有页为单张长图
        zoom = 2.0
        mat = fitz.Matrix(zoom, zoom)
        pixmaps = [page.get_pixmap(matrix=mat, alpha=False) for page in doc]
        # 单页情况：直接保存并裁掉底部空白
        if len(pixmaps) == 1:
            pixmaps[0].save(str(_tmp_png))
            try:
                from PIL import Image, ImageChops
                im = Image.open(str(_tmp_png)).convert("RGB")
                bg = Image.new("RGB", im.size, (244, 246, 248))
                diff = ImageChops.difference(im, bg)
                bbox = diff.getbbox()
                if bbox:
                    pad = 20
                    im = im.crop((0, 0, im.width, min(im.height, bbox[3] + pad)))
                im.save(str(png_path))
            except Exception:
                # 回退：直接覆盖写
                with open(_tmp_png, "rb") as _f, open(png_path, "wb") as _g:
                    _g.write(_f.read())
        else:
            # 竖向拼接
            from PIL import Image
            import io
            imgs = [Image.open(io.BytesIO(pm.tobytes("png"))) for pm in pixmaps]
            w = max(i.width for i in imgs)
            h = sum(i.height for i in imgs)
            combo = Image.new("RGB", (w, h), "white")
            y = 0
            for im in imgs:
                combo.paste(im, (0, y)); y += im.height
            combo.save(str(_tmp_png))
            with open(_tmp_png, "rb") as _f, open(png_path, "wb") as _g:
                _g.write(_f.read())
        doc.close()
        return True
    except Exception as e:
        print(f"⚠️  图片生成失败：{e}")
        return False


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("excel", help="输入 Excel 文件路径")
    ap.add_argument("--out-dir", default=str(HERE / "output"),
                    help="输出目录（默认 ./output）")
    ap.add_argument("--no-image", action="store_true", help="仅生成 HTML，不出图")
    args = ap.parse_args()

    excel_path = Path(args.excel).expanduser().resolve()
    if not excel_path.exists():
        print(f"❌ Excel 不存在：{excel_path}"); sys.exit(1)

    out_dir = Path(args.out_dir); out_dir.mkdir(parents=True, exist_ok=True)
    template = TEMPLATE_PATH.read_text(encoding="utf-8")
    data = read_excel(excel_path)
    html = render(data, template)

    stem = excel_path.stem
    html_path = out_dir / f"{stem}.html"
    html_path.write_text(html, encoding="utf-8")
    print(f"✅ HTML: {html_path}")

    if not args.no_image:
        png_path = out_dir / f"{stem}.png"
        if html_to_png(html_path, png_path):
            print(f"✅ PNG:  {png_path}")


if __name__ == "__main__":
    main()
