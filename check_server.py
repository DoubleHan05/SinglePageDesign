#!/usr/bin/env python3
"""ECS 启动脚本

部署路径约定：/data/project/xxxx/code/src/check_server.py
此文件作为 ECS 框架的服务入口，负责启动 Flask 应用。

启动方式：
    python3 /data/project/xxxx/code/src/check_server.py

环境变量（可选）：
    PORT   监听端口，默认 5001
    HOST   监听地址，默认 0.0.0.0
    DEBUG  设置为 1 时开启 Flask debug 模式
"""
import os
import sys
from pathlib import Path

# 让 server.py 及其依赖能被找到
BASE_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(BASE_DIR))

# 切到项目根目录，保证相对路径（templates/、logo/、samples/ 等）能被读到
os.chdir(BASE_DIR)


def main() -> None:
    host = os.environ.get("HOST", "0.0.0.0")
    port = int(os.environ.get("PORT", "5001"))
    debug = os.environ.get("DEBUG", "0") == "1"

    # 生产环境优先用 gunicorn；如果检测到 gunicorn 就 exec 到它
    if not debug and _try_exec_gunicorn(host, port):
        return

    # 否则使用 Flask 自带 server（仅开发 / 无 gunicorn 场景）
    from server import app
    print(f"[check_server] starting Flask on {host}:{port} (debug={debug})")
    app.run(host=host, port=port, debug=debug)


def _try_exec_gunicorn(host: str, port: int) -> bool:
    """尝试用 gunicorn 启动。找不到则返回 False，回退到 Flask 内置 server。"""
    import shutil
    gunicorn = shutil.which("gunicorn")
    if not gunicorn:
        return False
    args = [
        gunicorn,
        "-w", os.environ.get("WORKERS", "1"),   # 单 worker：session 存内存，多 worker 会失效
        "-b", f"{host}:{port}",
        "--timeout", os.environ.get("TIMEOUT", "120"),
        "--access-logfile", "-",
        "--error-logfile", "-",
        "server:app",
    ]
    print(f"[check_server] exec gunicorn: {' '.join(args)}")
    os.execv(gunicorn, args)
    return True  # 实际不会走到这里，os.execv 会替换进程


if __name__ == "__main__":
    main()
