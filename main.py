"""Сервер сайта (FastAPI). Запуск: двойной клик по START.bat или `python run.py`."""
from __future__ import annotations

import threading
from typing import Any

from fastapi import Body, FastAPI, HTTPException, Request
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles

import engine
import excel_io as xio

ROOT = xio.ROOT
CONFIG = xio.load_config()
WORKBOOK = xio.resolve(CONFIG["workbook"])  # единственная книга Excel с расчётной моделью
INDEX = ROOT / "static" / "index.html"

app = FastAPI(title="Mesh Delivery", docs_url="/api/docs")
app.mount("/static", StaticFiles(directory=ROOT / "static"), name="static")

_lock = threading.Lock()
_cache: dict[str, Any] = {"sig": None, "payload": None}


@app.middleware("http")
async def no_cache(request: Request, call_next):
    resp = await call_next(request)
    # Правите app.js / style.css — обновите страницу, и изменения видны сразу.
    resp.headers["Cache-Control"] = "no-store"
    return resp


def signature() -> str:
    return xio.file_sig(WORKBOOK)


def get_payload() -> dict[str, Any]:
    sig = signature()
    with _lock:
        if _cache["sig"] != sig or _cache["payload"] is None:
            try:
                payload = xio.build_model_payload(WORKBOOK)
            except FileNotFoundError as exc:
                raise HTTPException(status_code=404, detail=str(exc)) from exc
            except PermissionError as exc:
                raise HTTPException(status_code=423, detail="Excel занят другим процессом. Сохраните и закройте книгу, затем обновите страницу.") from exc
            except Exception as exc:
                raise HTTPException(status_code=500, detail=f"Не удалось прочитать Excel: {exc}") from exc
            payload["version"] = sig
            payload["model_file"] = WORKBOOK.name
            _cache["sig"], _cache["payload"] = sig, payload
        return _cache["payload"]


@app.get("/")
def home():
    if not INDEX.exists():
        raise HTTPException(status_code=404, detail="Не найден static/index.html")
    return FileResponse(INDEX)


store = engine.Store(ROOT / "state.json")


def run(fn, *a, **k):
    payload = get_payload()
    if payload.get("missing"):
        raise HTTPException(status_code=409, detail="В выбранной книге нет листов расчётной модели: " + ", ".join(payload["missing"]))
    try:
        return fn(engine.Cfg(payload["sheets"]), *a, **k)
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc


# ------------------------------------------------------------ сквозной процесс
@app.post("/api/receive")
def receive(b: dict = Body(...)):
    return run(store.receive, b["pvz"], b["sku"], b["qty"], b.get("seller", "S1"), b.get("size", "M"), b.get("event_id"))


@app.post("/api/orders")
def create_order(b: dict = Body(...)):
    return run(store.create_order, b["pvz_a"], b["pvz_b"], b["sku"], b["qty"], b.get("n"), b.get("pvz_b_ok", True),
               b.get("other_city", False), b.get("seller", "S1"), b.get("event_id"))


@app.post("/api/orders/{oid}/{action}")
def order_action(oid: str, action: str, b: dict = Body(default={})):
    return run(store.act, oid, action, b.get("fail", False), b.get("mode", "stay"), b.get("event_id"))


@app.post("/api/time/advance")
def advance(b: dict = Body(...)):
    return store.advance(b.get("days", 1), b.get("event_id"))


@app.post("/api/expire")
def expire(b: dict = Body(default={})):
    return run(store.remove_expired, b.get("event_id"))


@app.post("/api/reset")
def reset():
    store.reset()
    return {"ok": True}


@app.get("/api/state")
def state():
    return run(store.state)


@app.get("/api/analytics")
def analytics():
    return run(store.analytics)


# ------------------------------------------------------------ данные Excel
@app.get("/api/data")
def api_data():
    return get_payload()


@app.get("/api/health")
def health():
    return {"ok": True, "workbook": WORKBOOK.name, "workbook_exists": WORKBOOK.exists()}


if __name__ == "__main__":
    import run as launcher
    launcher.main()
