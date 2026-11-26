"""Чтение Excel-книг и исходников проекта. Не зависит от FastAPI — удобно тестировать отдельно.

Одна книга — data/Mesh_проверка_модели.xlsx (путь задан в config.json, ключ workbook).
Из неё считаются тарифы и все расчёты сайта. Служебные листы проверки модели на сайт не попадают.
"""
from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Any

import openpyxl

ROOT = Path(__file__).resolve().parent

DEFAULT_CONFIG = {
    "host": "127.0.0.1",
    "port": 8000,
    "open_browser": True,
    "workbook": "data/Mesh_проверка_модели.xlsx",
}
# Служебные листы (проверка модели, сравнение с ТЗ): сайт их не читает и пользователю не отдаёт.
HIDDEN_SHEETS = {"Замечания", "Соответствие ТЗ", "Стресс-тесты"}

# Листы, без которых расчётная модель не работает (движок читает их по адресам ячеек).
REQUIRED_MODEL_SHEETS = ["Параметры", "Плечи", "Товары", "Заказы", "Сравнение", "Возвраты",
                         "ПВЗ_экономика", "Продавец_экономика", "Сценарии_и_риски"]


def load_config() -> dict[str, Any]:
    cfg = dict(DEFAULT_CONFIG)
    path = ROOT / "config.json"
    if path.exists():
        try:
            cfg.update(json.loads(path.read_text(encoding="utf-8-sig")))
        except Exception as exc:  # битый JSON не должен ронять сайт
            print(f"[config] config.json не прочитан ({exc}), использую значения по умолчанию")
    return cfg


def resolve(rel: str) -> Path:
    p = Path(rel)
    return p if p.is_absolute() else ROOT / p


def val(value: Any, default: Any = None) -> Any:
    if value is None:
        return default
    if isinstance(value, (str, int, float, bool)):
        return value
    if hasattr(value, "isoformat"):
        return value.isoformat()
    return str(value)


def number(value: Any, default: float = 0.0) -> float:
    try:
        return float(value) if value is not None else default
    except (TypeError, ValueError):
        return default


def file_sig(path: Path) -> str:
    try:
        st = path.stat()
        return f"{st.st_mtime_ns}:{st.st_size}"
    except OSError:
        return "missing"


def sheet_rows(wb, name: str):
    if name not in wb.sheetnames:
        return []
    return list(wb[name].iter_rows(values_only=True))


def read_all_sheets(path: Path) -> dict[str, list[list[Any]]]:
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    try:
        return {ws.title: [[val(c) for c in row] for row in ws.iter_rows(values_only=True)] for ws in wb.worksheets}
    finally:
        wb.close()


# ---------------------------------------------------------------- модель
def build_model_payload(path: Path) -> dict[str, Any]:
    if not path.exists():
        raise FileNotFoundError(f"Не найден файл: {path}")
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    try:
        names = list(wb.sheetnames)
        missing = [s for s in REQUIRED_MODEL_SHEETS if s not in names]
        if missing:
            # Книга — не расчётная модель (например, книга проверки). Не считаем нули, а честно сообщаем.
            return {"workbook": path.name, "missing": missing, "sheets": {},
                    "orders": [], "products": [], "parameters": [], "params": {}, "pvz": [],
                    "comparison": [], "returns": [], "metrics": {"orders": 0, "old_total": 0, "new_total": 0,
                                                                  "saving": 0, "saving_pct": 0, "routes": {}, "products": 0},
                    "uncomputed": False}

        params, param_map = [], {}
        for row in sheet_rows(wb, "Параметры")[4:]:
            if not row or not row[0]:
                continue
            name, value = str(row[0]), val(row[1])
            if value is not None:
                params.append({"name": name, "value": value, "unit": val(row[2], ""), "comment": val(row[3], "")})
                param_map[name] = value

        def p(name, default=0):
            return number(param_map.get(name), default)

        products = []
        for row in sheet_rows(wb, "Товары")[3:]:
            if not row or not row[0]:
                continue
            products.append({"sku": val(row[0]), "name": val(row[1]), "group": val(row[2]),
                             "storage": val(row[3]), "shelf_days": val(row[4], 0), "orders_per_day": val(row[5], 0)})

        orders, route_counts, pvz_counts = [], {}, {}
        for row in sheet_rows(wb, "Заказы")[3:]:
            if not row or row[0] is None:
                continue
            route = val(row[9], "Не рассчитано") if len(row) > 9 else "Не рассчитано"
            delivery_pvz = val(row[5], "")
            route_counts[str(route)] = route_counts.get(str(route), 0) + 1
            pvz_counts[str(delivery_pvz)] = pvz_counts.get(str(delivery_pvz), 0) + 1
            orders.append({
                "id": val(row[0]), "day": val(row[1]), "sku": val(row[2]), "group": val(row[3]),
                "shelf_days": val(row[4]), "delivery_pvz": delivery_pvz, "qty": val(row[6]),
                "returned": bool(number(row[7])), "storage": val(row[8]), "route": route,
                "orders_in_direction": val(row[10]), "trips": val(row[11]), "transport": val(row[12]),
                "old_cost": val(row[13]), "new_cost": val(row[14]), "saving": val(row[15]),
                "seller_storage": val(row[16]), "pvz_margin": val(row[17]),
            })

        comparison, comparison_map = [], {}
        for row in sheet_rows(wb, "Сравнение")[4:]:
            if row and row[0] is not None:
                label, value = str(row[0]), val(row[1])
                comparison.append({"label": label, "value": value})
                comparison_map[label] = value

        def metric(label, fallback=0):
            return number(comparison_map.get(label), fallback)

        old_total = metric("Расходы МП — старая схема (всё через хаб)", sum(number(x["old_cost"]) for x in orders))
        new_total = metric("Расходы МП — новая схема (mesh по правилам)", sum(number(x["new_cost"]) for x in orders))
        saving = metric("Экономия, ₽", old_total - new_total)
        saving_pct = metric("Экономия, %", saving / old_total if old_total else 0)

        returns = [{"label": val(r[0]), "value": val(r[1])} for r in sheet_rows(wb, "Возвраты")[4:] if r and r[0] is not None]

        sheets = {ws.title: [[val(c) for c in row] for row in ws.iter_rows(values_only=True)]
                  for ws in wb.worksheets if ws.title != "Заказы" and ws.title not in HIDDEN_SHEETS}
        k2 = sheets.get("К2_Панель", [])
        uncomputed = not (len(k2) > 4 and len(k2[4]) > 1 and k2[4][1] is not None)
        pvz = [{"code": c, "orders": n} for c, n in sorted(pvz_counts.items()) if c not in ("", "None")]
        return {
            "workbook": path.name, "missing": [],
            "metrics": {"orders": len(orders), "old_total": old_total, "new_total": new_total, "saving": saving,
                        "saving_pct": saving_pct, "routes": route_counts, "products": len(products)},
            "products": products, "orders": orders, "parameters": params,
            "params": {
                "trip_cost": p("Остановок в рейсе") * p("Стоимость одной остановки")
                             + p("Длина рейса") * p("Стоимость километра")
                             + p("Простой курьера за рейс") * p("Час работы курьера"),
                "min_orders": p("Минимум заказов, чтобы рейс поехал"),
                "trip_windows": p("Окон рейсов в день"),
                "hub_cost": p("Хаб + магистраль на заказ"),
            },
            "pvz": pvz, "comparison": comparison, "returns": returns, "sheets": sheets, "uncomputed": uncomputed,
        }
    finally:
        wb.close()

