#!/usr/bin/env python3
"""ECS 启动脚本

部署路径约定：/data/project/xxxx/code/src/check_server.py
此文件可以放在独立目录里，通过环境变量 PROJECT_ROOT 指定项目根目录。

环境变量：
    PROJECT_ROOT  项目根路径（必须包含 server.py / templates/ / logo/ 等）
                  默认值：脚本所在目录的上一级（../）
    PORT          监听端口，默认 5001
    HOST          监听地址，默认 0.0.0.0
    WORKERS       gunicorn worker 数，默认 1
    TIMEOUT       gunicorn 超时秒数，默认 120
    DEBUG         设为 1 时回退到 Flask 内置 server 并开启 debug 模式

启动示例：
    # check_server.py 和项目同级目录
    PROJECT_ROOT=/data/project/xxxx/code/src python3 /data/project/xxxx/code/src/check_server.py

    # 或直接用同目录的默认推断
    python3 /data/project/xxxx/code/src/check_server.py
"""
import os
import sys
from pathlib import Path

_SCRIPT_DIR = Path(__file__).resolve().parent

# PROJECT_ROOT 优先用环境变量；未设时：
#   - 若 server.py 与本脚本同目录，则用脚本目录
#   - 否则上移一级（适合 check_server.py 在独立子目录的场景）
_env_root = os.environ.get("PROJECT_ROOT")
if _env_root:
    BASE_DIR = Path(_env_root).resolve()
elif (_SCRIPT_DIR / "server.py").exists():
    BASE_DIR = _SCRIPT_DIR
else:
    BASE_DIR = _SCRIPT_DIR.parent

if not (BASE_DIR / "server.py").exists():
    print(f"[check_server] ERROR: 找不到 server.py，请设置环境变量 PROJECT_ROOT 指向项目根目录")
    print(f"[check_server] 当前推断根目录：{BASE_DIR}")
    sys.exit(1)

sys.path.insert(0, str(BASE_DIR))
os.chdir(BASE_DIR)


def main() -> None:
    host = os.environ.get("HOST", "0.0.0.0")
    port = int(os.environ.get("PORT", "5001"))
    debug = os.environ.get("DEBUG", "0") == "1"

    print(f"[check_server] 项目根目录：{BASE_DIR}")

    if not debug and _try_exec_gunicorn(host, port):
        return

    from server import app
    print(f"[check_server] starting Flask on {host}:{port} (debug={debug})")
    app.run(host=host, port=port, debug=debug)


def _try_exec_gunicorn(host: str, port: int) -> bool:
    import shutil
    gunicorn = shutil.which("gunicorn")
    if not gunicorn:
        return False
    args = [
        gunicorn,
        "-w", os.environ.get("WORKERS", "1"),
        "-b", f"{host}:{port}",
        "--timeout", os.environ.get("TIMEOUT", "120"),
        "--access-logfile", "-",
        "--error-logfile", "-",
        "server:app",
    ]
    print(f"[check_server] exec gunicorn: {' '.join(args)}")
    os.execv(gunicorn, args)
    return True


if __name__ == "__main__":
    main()
