# 快麦产品单页生成工具

一个通过 Excel 生成产品单页 HTML/PNG 的小工具。用户在网页上拖入 Excel，即可预览效果并下载图片。

## 目录结构

```
SinglePageDesign/
├── server.py                    # Flask 服务（网页端入口）
├── generate.py                  # 核心：读 Excel → 生成 HTML → 截图为 PNG
├── build_sample_xlsx.py         # 生成示例 Excel（KM360N）
├── templates/
│   └── template.html            # 单页 HTML 模板
├── web/
│   ├── index.html               # 前端页面（拖拽上传/预览/下载）
│   └── excel-icon.svg
├── logo/
│   ├── kuaimai_logo.svg
│   └── guangyun_logo.png
├── samples/
│   ├── build_template.py        # 生成"下载模板"用的 Excel
│   ├── 产品单页模板.xlsx        # 用户下载的空白模板
│   └── KM360N.xlsx              # 示例数据
└── output/                      # 生成的 HTML / PNG
```

## 环境准备

```bash
cd "SinglePageDesign"
python3 -m pip install flask openpyxl pillow --break-system-packages
```

图片渲染二选一（用于把 HTML 转成 PNG）：

- **推荐 Playwright**（自带 Chromium，效果最好）
  ```bash
  python3 -m pip install playwright --break-system-packages
  python3 -m playwright install chromium
  ```
  如果本机已装 Chrome/Edge，脚本会优先直接调用，不必等 Chromium 下载。
- **或 WeasyPrint**（无浏览器，装系统库更简单）
  ```bash
  brew install pango
  python3 -m pip install weasyprint pymupdf --break-system-packages
  ```

## 运行网页版

```bash
python3 server.py
```

浏览器打开 http://127.0.0.1:5001，流程：

1. 点顶部"下载 Excel 模板"拿到 `产品单页模板.xlsx`
2. 打开填写内容（图片可以直接在 B 列插入）
3. 保存后拖到页面中央的上传框（或点击选择）
4. 右侧 iframe 展示预览
5. 确认无误 → 点右上角"确认生成并下载图片"
6. 下载得到 `型号-产品单页-YYYYMMDD.png`

## 命令行用法

```bash
# 只出 HTML
python3 generate.py samples/KM360N.xlsx --no-image

# 出 HTML + PNG
python3 generate.py samples/KM360N.xlsx

# 指定输出目录
python3 generate.py samples/KM360N.xlsx --out-dir ~/Desktop/pages
```

## Excel 结构

### Sheet1「产品数据」

第一列参数名（不要改），第二列参数值。

| 参数名 | 说明 |
| --- | --- |
| 型号 | 产品型号，显示在左上角 |
| 产品图 | 主图，可填路径或在 B 列插入图片 |
| 细节图1 / 细节图2 / … | 底部细节图，支持任意张数 |
| 特点1 / 特点2 / … | 产品特点列表 |
| 场景1 / 场景2 / … | 应用场景，格式 `仓\|仓储`（竖线前=圆圈内字，后=下方标签） |

### Sheet2「参数表」

两列：参数名称、参数值。无需任何前缀，方便直接粘贴。空值页面上显示 `/`。

### Sheet3「使用说明」

填写指引，可保留。

## 图片支持

三种方式任选：

1. 相对/绝对路径直接填在 B 列
2. 常规「浮动图片」：Excel 插入图片，锚点放在对应行
3. WPS/新版 Excel 的「在单元格中放置图片」（DISPIMG 公式）

## 输出文件命名

下载图片名格式：`<型号>-产品单页-<YYYYMMDD>.png`。型号取自 Excel 中「型号」字段。
