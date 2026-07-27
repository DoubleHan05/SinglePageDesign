# 快麦产品单页生成工具

一个通过 Excel 生成产品单页 HTML/PNG 的小工具。用户在网页上拖入 Excel，即可预览效果并下载图片。

**项目地址**：https://github.com/DoubleHan05/SinglePageDesign

**License**：[MIT](./LICENSE)

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
│   ├── kuaimai_logo.svg          # 快麦彩色 logo（浅灰底用）
│   ├── kuaimai_logo_white.svg    # 快麦白色 logo（蓝底用）
│   └── guangyun_logo.png         # 光云 logo（蓝底时经 CSS 滤镜转白）
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

图片截图由**用户浏览器**完成（html2canvas），服务器无需安装 Playwright 或 WeasyPrint。

## 运行网页版

```bash
python3 server.py
```

浏览器打开 http://127.0.0.1:8382，流程：

1. 点顶部"下载 Excel 模板"拿到 `产品单页模板.xlsx`
2. 打开填写内容（图片可以直接在 B 列插入）
3. 保存后拖到页面中央的上传框（或点击选择）
4. 右侧 iframe 展示预览
5. 预览栏右侧可切换背景（**蓝底 `#1489E8`** 默认 / **浅灰底 `#F4F6F8`**）
6. 确认无误 → 点右上角"确认生成并下载图片"
7. 浏览器自动打开新标签页，用 html2canvas 截图，下载 `型号-产品单页-YYYYMMDD.png` 后自动关闭

### 主题说明

- **蓝底**：页面背景 `#1489E8`，顶部"快麦打印机 / 股票代码：688365"文字变为白色，两个 logo 自动换成白色版
- **浅灰底**：页面背景 `#F4F6F8`，顶部使用彩色 logo 和深色文字

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

## 服务器部署

### 一、准备工作

服务器建议 Ubuntu 20.04+ / CentOS 8+ / Debian 11+，Python ≥ 3.9。

```bash
# 系统依赖（Ubuntu/Debian）
sudo apt update
sudo apt install -y python3 python3-pip python3-venv git \
    libpango-1.0-0 libpangoft2-1.0-0 fonts-noto-cjk

# CentOS/RHEL
sudo dnf install -y python3 python3-pip git pango pango-devel google-noto-sans-cjk-fonts
```

中文字体是必须的，否则渲染出来的图片会显示成方块。

### 二、拉取代码 & 安装依赖

```bash
git clone https://github.com/DoubleHan05/SinglePageDesign.git /opt/kuaimai-single-page
cd /opt/kuaimai-single-page

python3 -m venv .venv
source .venv/bin/activate
pip install flask openpyxl pillow weasyprint pymupdf gunicorn
```

如果偏好用 Playwright 出图：

```bash
pip install playwright
python -m playwright install --with-deps chromium
```

### 三、监听地址 / 启动脚本

推荐通过 `check_server.py` 启动，无需修改 `server.py`：

```bash
python3 check_server.py
```

`check_server.py` 会自动定位脚本所在目录作为项目根（用 `Path(__file__).resolve().parent`），
因此**放在任何路径下都能工作**，例如 ECS 常见约定 `/data/project/<任意名>/code/src/check_server.py`。

常用环境变量：

| 变量 | 默认值 | 说明 |
| --- | --- | --- |
| `HOST` | `0.0.0.0` | 监听地址 |
| `PORT` | `8382` | 监听端口 |
| `WORKERS` | `1` | gunicorn worker 数，Session 存内存，建议保持 1 |
| `TIMEOUT` | `120` | gunicorn 请求超时（秒），首次出图较慢 |
| `DEBUG` | `0` | 设为 `1` 则用 Flask 内置 server 并开启 debug |

启动逻辑：
- 检测到 `gunicorn` 时直接 `os.execv` 到 gunicorn（access/error 日志走 stdout，便于统一收集）
- 找不到 `gunicorn` 或 `DEBUG=1` 时回退到 Flask 内置 server

如需手动跑 gunicorn，可参考等价命令：

```bash
gunicorn -w 1 -b 0.0.0.0:8382 --timeout 120 server:app
```

> **注意**：Session 保存在进程内存中，多 worker 会导致预览/下载分散到不同进程后失效。请使用 `WORKERS=1` 或 `-w 1`，或后续接入 Redis / 文件缓存后再横向扩展。


### 四、systemd 常驻

新建 `/etc/systemd/system/single-page.service`：

```ini
[Unit]
Description=Kuaimai single-page generator
After=network.target

[Service]
Type=simple
User=www-data
WorkingDirectory=/opt/kuaimai-single-page
Environment="PATH=/opt/kuaimai-single-page/.venv/bin"
Environment="HOST=127.0.0.1" "PORT=5001" "WORKERS=1" "TIMEOUT=120"
ExecStart=/opt/kuaimai-single-page/.venv/bin/python3 /opt/kuaimai-single-page/check_server.py
Restart=on-failure
RestartSec=3

[Install]
WantedBy=multi-user.target
```

启用：

```bash
sudo systemctl daemon-reload
sudo systemctl enable --now single-page
sudo systemctl status single-page
```

### 五、Nginx 反向代理

```nginx
server {
    listen 80;
    server_name your-domain.example.com;

    client_max_body_size 20M;   # 允许上传较大的 xlsx / 图片

    location / {
        proxy_pass         http://127.0.0.1:5001;
        proxy_http_version 1.1;
        proxy_set_header   Host $host;
        proxy_set_header   X-Real-IP $remote_addr;
        proxy_set_header   X-Forwarded-For $proxy_add_x_forwarded_for;
        proxy_read_timeout 120s;
    }
}
```

用 Certbot 加 HTTPS：

```bash
sudo certbot --nginx -d your-domain.example.com
```

### 六、目录权限

服务运行用户（示例里是 `www-data`）需要读写 `output/` 目录：

```bash
sudo chown -R www-data:www-data /opt/kuaimai-single-page/output
```

### 七、更新部署

```bash
cd /opt/kuaimai-single-page
git pull
source .venv/bin/activate
pip install -r requirements.txt  # 如果新增了依赖
sudo systemctl restart single-page
```

### 八、常见问题排查

| 现象 | 原因 / 解决 |
| --- | --- |
| 图片里文字变方块 | 服务器未安装中文字体，装 `fonts-noto-cjk` |
| 生成图片 500 | Playwright/WeasyPrint 都没装，或依赖库缺失 |
| 上传大文件 413 | 调大 Nginx `client_max_body_size` |
| Gunicorn worker 超时 | 加大 `--timeout`，或调低 worker 并发 |
| session 数据变多 | `output/_web/` 会累积上传记录，可定期清理 |
