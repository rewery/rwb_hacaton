"""Один файл для запуска сайта: python run.py (или двойной клик по START.bat).

Что делает при каждом запуске:
  1. Создаёт папку .venv рядом с проектом (один раз) и ставит библиотеки из requirements.txt
     (интернет нужен только при первом запуске или если requirements.txt изменился).
  2. Запускает сервер и открывает сайт в браузере.
"""
from __future__ import annotations

import hashlib
import json
import os
import subprocess
import sys
import threading
import time
import urllib.request
import webbrowser
from pathlib import Path

ROOT = Path(__file__).resolve().parent
VENV = ROOT / ".venv"
STAMP = VENV / ".requirements.sha256"
MIN_PY = (3, 9)

for _s in (sys.stdout, sys.stderr):
    try:
        _s.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass


def say(msg: str = "") -> None:
    print(msg, flush=True)


def venv_python() -> Path:
    return VENV / ("Scripts/python.exe" if os.name == "nt" else "bin/python")


def in_venv() -> bool:
    try:
        return Path(sys.prefix).resolve() == VENV.resolve()
    except OSError:
        return False


def req_hash() -> str:
    return hashlib.sha256((ROOT / "requirements.txt").read_bytes()).hexdigest()


def ensure_environment() -> None:
    """Создать .venv и поставить зависимости; затем перезапустить себя внутри .venv."""
    if sys.version_info < MIN_PY:
        raise SystemExit(f"Нужен Python {MIN_PY[0]}.{MIN_PY[1]}+, сейчас {sys.version.split()[0]}. Скачайте: https://www.python.org/downloads/")
    if in_venv():
        return
    py = venv_python()
    if not py.exists():
        say("Первый запуск: создаю окружение .venv …")
        import venv
        venv.create(VENV, with_pip=True, clear=False)
    if not STAMP.exists() or STAMP.read_text().strip() != req_hash():
        say("Устанавливаю библиотеки (нужен интернет, 1–2 минуты) …")
        cmd = [str(py), "-m", "pip", "install", "--disable-pip-version-check", "-r", str(ROOT / "requirements.txt")]
        if subprocess.call(cmd) != 0:
            raise SystemExit("Не удалось установить библиотеки. Проверьте интернет и запустите снова.")
        STAMP.write_text(req_hash())
        say("Готово.")
    raise SystemExit(subprocess.call([str(py), str(Path(__file__).resolve()), *sys.argv[1:]]))


def port_free(host: str, port: int) -> bool:
    import socket
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.settimeout(0.3)
        return s.connect_ex((host, port)) != 0


def our_site_on(host: str, port: int) -> bool:
    try:
        with urllib.request.urlopen(f"http://{host}:{port}/api/health", timeout=1) as r:
            return bool(json.loads(r.read().decode("utf-8")).get("ok"))
    except Exception:
        return False


def open_browser_when_ready(host: str, port: int) -> None:
    for _ in range(60):
        if not port_free(host, port):
            break
        time.sleep(0.25)
    webbrowser.open(f"http://{host}:{port}")


def main() -> None:
    ensure_environment()
    sys.path.insert(0, str(ROOT))
    os.chdir(ROOT)
    import excel_io
    import uvicorn

    cfg = excel_io.load_config()
    host, port = cfg["host"], int(cfg["port"])

    if not port_free(host, port):
        if our_site_on(host, port):
            say(f"Сайт уже запущен: http://{host}:{port}")
            if cfg.get("open_browser", True):
                webbrowser.open(f"http://{host}:{port}")
            return
        for p in range(port + 1, port + 20):
            if port_free(host, p):
                say(f"Порт {port} занят другой программой, использую {p}.")
                port = p
                break
        else:
            raise SystemExit(f"Свободный порт рядом с {port} не найден. Измените port в config.json.")

    path = excel_io.resolve(cfg["workbook"])
    say(f"  Excel-книга: {path}" + ("" if path.exists() else "   <-- ФАЙЛ НЕ НАЙДЕН"))

    say(f"\nСайт: http://{host}:{port}   (остановить: Ctrl+C или закрыть это окно)\n")
    if cfg.get("open_browser", True):
        threading.Thread(target=open_browser_when_ready, args=(host, port), daemon=True).start()
    uvicorn.run("main:app", host=host, port=port, log_level="warning")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        pass
    except SystemExit as e:
        if e.code not in (None, 0):
            say(str(e.code) if isinstance(e.code, str) else "")
        raise
    except Exception:
        import traceback
        traceback.print_exc()
        input("\nОшибка. Нажмите Enter, чтобы закрыть окно …")
        sys.exit(1)
