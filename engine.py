"""Сквозной процесс Mesh: приёмка -> хранение -> заказ -> сборка -> рейс -> выдача -> возврат.
Все тарифы читаются из Excel (листы «Параметры», «Плечи», «Продавец_экономика», «Сценарии_и_риски»).
Состояние (остатки, заказы, проводки) хранится в state.json, Excel не изменяется."""
import json, threading
from pathlib import Path


class Cfg:
    def __init__(self, sheets): self.s = sheets
    def c(self, sh, a):
        try: return self.s[sh][int(a[1:]) - 1][ord(a[0]) - 65]
        except Exception: return None
    def n(self, sh, a):
        try: return float(self.c(sh, a))
        except (TypeError, ValueError): return 0.0
    def P(self, a): return self.n("Параметры", a)
    def SE(self, a): return self.n("Продавец_экономика", a)
    def hub(self): return sum(self.n("Плечи", f"B{r}") for r in range(5, 9))
    def hub_return(self): return sum(self.n("Плечи", f"B{r}") for r in (6, 7, 8))
    def trip(self): return self.P("B10") * self.P("B11") + self.P("B12") * self.P("B13") + self.P("B14") * self.P("B15")
    def fail_extra(self):
        e = self.n("Сценарии_и_риски", "E10")
        return e - (self.hub() + self.P("B27")) if e else 82.0
    def seller_store(self, q, d):
        r, f, t, m = self.P("B28"), self.P("B29"), self.P("B37"), self.P("B38")
        return q * (max(0, min(d, t) - f) * r + max(0, d - t) * r * m)
    def removal(self, q): return q * self.SE("B15")


class Store:
    def __init__(self, path):
        self.path, self.lock, self.d = Path(path), threading.RLock(), None
        if self.path.exists():
            try: self.d = json.loads(self.path.read_text(encoding="utf-8"))
            except Exception: self.d = None
        if not self.d: self.reset(False)

    def reset(self, save=True):
        self.d = {"day": 0, "seq": 0, "batches": [], "orders": {}, "ledger": [], "events": {}, "sizes": {}}
        if save: self.save()

    def save(self):
        try: self.path.write_text(json.dumps(self.d, ensure_ascii=False), encoding="utf-8")
        except OSError: pass

    def once(self, eid, fn):
        with self.lock:
            if eid and eid in self.d["events"]: return {**self.d["events"][eid], "duplicate": True}
            r = fn()
            if eid: self.d["events"][eid] = r
            self.save()
            return r

    def L(self, order, party, kind, rev=0.0, cost=0.0):
        self.d["ledger"].append({"order": order, "party": party, "kind": kind, "rev": round(rev, 2), "cost": round(cost, 2), "day": self.d["day"]})

    def onhand(self, pvz): return sum(b["left"] for b in self.d["batches"] if b["pvz"] == pvz and not b["removed"])

    def new_batch(self, pvz, sku, q, seller):
        self.d["seq"] += 1
        b = {"id": f"B{self.d['seq']}", "pvz": pvz, "sku": sku, "left": q, "res": 0, "day": self.d["day"], "seller": seller, "removed": False}
        self.d["batches"].append(b)
        return b

    # --- приёмка ---
    def receive(self, cfg, pvz, sku, qty, seller="S1", size="M", event_id=None):
        def go():
            q = int(qty)
            if q <= 0: raise ValueError("Количество должно быть больше 0")
            if size not in ("S", "M", "L"): raise ValueError("Размер ПВЗ: S, M или L")
            lim = cfg.P({"S": "B41", "M": "B42", "L": "B43"}[size]); on = self.onhand(pvz)
            if on + q > lim: raise ValueError(f"Нет места на ПВЗ {pvz} ({size}): на полке {on:.0f} из {lim:.0f}, приёмка {q} шт. отклонена")
            self.d["sizes"][pvz] = size
            b = self.new_batch(pvz, sku, q, seller)
            plan = {d: round(cfg.seller_store(q, d) + (cfg.removal(q) if d > cfg.P("B39") else 0), 2) for d in (4, 30, 60, 61)}
            return {"batch": b["id"], "plan": plan, "message": f"Принято {q} шт на ПВЗ {pvz}. Первые {cfg.P('B29'):.0f} дн. хранение платит МП, дальше продавец: за 30 дн. {plan[30]} ₽, за 60 дн. {plan[60]} ₽, после 60 дн. {plan[61]} ₽ с вывозом"}
        return self.once(event_id, go)

    # --- заказ и маршрутизация ---
    def create_order(self, cfg, pvz_a, pvz_b, sku, qty, n=None, pvz_b_ok=True, other_city=False, seller="S1", event_id=None):
        def go():
            q, day = int(qty), self.d["day"]
            n_ = int(n) if n else int(cfg.P("B16"))
            if q <= 0: raise ValueError("Количество должно быть больше 0")
            bs = sorted([b for b in self.d["batches"] if b["pvz"] == pvz_a and b["sku"] == sku and not b["removed"] and day - b["day"] <= cfg.P("B39")], key=lambda b: b["day"])
            avail = sum(b["left"] - b["res"] for b in bs); why = []
            if other_city: why.append("межгород: покупатель в другом городе")
            if avail < q: why.append(f"на ПВЗ {pvz_a} доступно {avail:.0f} из {q} шт")
            if not pvz_b_ok: why.append(f"ПВЗ {pvz_b} недоступен")
            if n_ < cfg.P("B17"): why.append(f"в рейсе {n_} < минимума {cfg.P('B17'):.0f}")
            if cfg.P("B23") > 6: why.append("целевой SLA > 6 ч")
            route, alloc = ("hub" if why else "mesh"), []
            if route == "mesh":
                rem = q
                for b in bs:
                    t = min(rem, b["left"] - b["res"])
                    if t > 0: b["res"] += t; alloc.append({"b": b["id"], "q": t}); rem -= t
            self.d["seq"] += 1
            o = {"id": f"O{self.d['seq']}", "sku": sku, "qty": q, "pvz_a": pvz_a, "pvz_b": pvz_b, "seller": seller, "n": n_, "route": route, "why": why, "alloc": alloc, "parts": [], "status": "CREATED", "created": day}
            self.d["orders"][o["id"]] = o
            return {"order": o["id"], "route": route, "message": f"Заказ {o['id']}: маршрут {'Mesh' if route == 'mesh' else 'через хаб'}" + (f" ({'; '.join(why)})" if why else ", резерв поставлен")}
        return self.once(event_id, go)

    def act(self, cfg, oid, action, fail=False, mode="stay", event_id=None):
        def go():
            o = self.d["orders"].get(oid)
            if not o: raise ValueError("Заказ не найден")
            st, mesh, day = o["status"], o["route"] == "mesh", self.d["day"]
            def need(s):
                if st != s: raise ValueError(f"Действие «{action}» недоступно в статусе {st}")
            if action == "assemble":
                need("CREATED")
                if not mesh: raise ValueError("Сборка нужна только для Mesh")
                o["status"] = "ASSEMBLED"
            elif action == "dispatch":
                need("ASSEMBLED" if mesh else "CREATED")
                if mesh:
                    for a in o["alloc"]:
                        b = next(x for x in self.d["batches"] if x["id"] == a["b"])
                        b["left"] -= a["q"]; b["res"] -= a["q"]; o["parts"].append({"q": a["q"], "d": day - b["day"]})
                    if fail: o["route"], o["parts"], o["why"] = "hub_fallback", [], ["сбой ПВЗ Б — заказ переведён на хаб"]
                o["status"] = "IN_TRANSIT"
            elif action == "deliver":
                need("IN_TRANSIT"); self.book_delivery(cfg, o); o["status"] = "DELIVERED"
            elif action == "return":
                need("DELIVERED"); self.book_return(cfg, o, mode); o["status"] = "RETURNED"
            else: raise ValueError("Неизвестное действие")
            return {"order": oid, "status": o["status"], "route": o["route"], "message": f"{oid}: {o['status']}"}
        return self.once(event_id, go)

    # --- начисления ---
    def seller_lines(self, cfg, q):
        g = cfg.SE("B4") * q
        return g, cfg.SE("B6") * q, g * cfg.SE("B7"), g * cfg.SE("B8"), cfg.SE("B9") + cfg.SE("B10")

    def book_delivery(self, cfg, o):
        q, oid, out, rate, free = o["qty"], o["id"], cfg.P("B27"), cfg.P("B28"), cfg.P("B29")
        pa, pb, sl = f"pvz:{o['pvz_a']}", f"pvz:{o['pvz_b']}", f"seller:{o['seller']}"
        g, cogs, comm, oth, prep = self.seller_lines(cfg, q); st = 0.0
        if o["route"] == "mesh":
            mps = sum(p["q"] * rate * min(p["d"], free) for p in o["parts"])
            st = sum(cfg.seller_store(p["q"], p["d"]) for p in o["parts"])
            self.L(oid, "mp", "Рейс ПВЗ→ПВЗ (доля)", cost=cfg.trip() / o["n"])
            self.L(oid, "mp", "Выдача на ПВЗ Б", cost=out)
            self.L(oid, "mp", "Приёмка и сборка на ПВЗ А", cost=q * (cfg.P("B25") + cfg.P("B26")))
            self.L(oid, "mp", f"Хранение за счёт МП (до {free:.0f} дн.)", cost=mps)
            self.L(oid, pa, "Приёмка и сборка", rev=q * (cfg.P("B25") + cfg.P("B26")), cost=q * (cfg.P("B31") + cfg.P("B32")))
            self.L(oid, pa, "Хранение (МП + продавец)", rev=mps + st, cost=sum(p["q"] * p["d"] * cfg.P("B34") for p in o["parts"]))
        else:
            self.L(oid, "mp", "Хаб + магистраль", cost=cfg.hub())
            self.L(oid, "mp", "Выдача на ПВЗ Б", cost=out)
            if o["route"] == "hub_fallback": self.L(oid, "mp", "Сбой ПВЗ Б: доп. расходы", cost=cfg.fail_extra())
        self.L(oid, pb, "Выдача заказа", rev=out, cost=cfg.P("B33"))
        for k, rv, cs in (("Продажа (GMV)", g, 0), ("Себестоимость товара", 0, cogs), ("Комиссия МП", 0, comm), ("Прочие удержания", 0, oth), ("Подготовка и подвоз", 0, prep), ("Хранение на ПВЗ", 0, st)):
            self.L(oid, sl, k, rev=rv, cost=cs)

    def book_return(self, cfg, o, mode):
        q, oid, mesh = o["qty"], o["id"], o["route"] == "mesh"
        pb = f"pvz:{o['pvz_b']}"; g, cogs, comm, oth, _ = self.seller_lines(cfg, q)
        if mesh and mode == "stay":
            lim = cfg.P({"S": "B41", "M": "B42", "L": "B43"}[self.d["sizes"].get(o["pvz_b"], "M")])
            if self.onhand(o["pvz_b"]) + q > lim: mode = "ride"
        rc = cfg.hub_return() if not mesh else q * cfg.P("B25") + {"stay": 0, "ride": cfg.trip() / o["n"], "van": cfg.trip()}[mode]
        self.L(oid, "mp", f"Возврат ({'хаб' if not mesh else mode})", cost=rc)
        self.L(oid, pb, "Приёмка возврата", rev=q * cfg.P("B25"), cost=q * cfg.P("B31"))
        self.L(oid, f"seller:{o['seller']}", "Возврат: продажа аннулирована", rev=-g, cost=-(cogs + comm + oth))
        if mesh and mode == "stay": self.new_batch(o["pvz_b"], o["sku"], q, o["seller"])

    # --- время и неликвид ---
    def advance(self, days, event_id=None):
        def go():
            self.d["day"] += int(days)
            return {"day": self.d["day"], "message": f"День {self.d['day']}"}
        return self.once(event_id, go)

    def remove_expired(self, cfg, event_id=None):
        def go():
            n = 0
            for b in self.d["batches"]:
                age = self.d["day"] - b["day"]
                if b["removed"] or b["res"] or b["left"] <= 0 or age <= cfg.P("B39"): continue
                q, pa, sl = b["left"], f"pvz:{b['pvz']}", f"seller:{b['seller']}"
                mps, st, rm = q * cfg.P("B28") * cfg.P("B29"), cfg.seller_store(q, age), cfg.removal(q)
                self.L(b["id"], "mp", "Хранение за счёт МП", cost=mps)
                self.L(b["id"], pa, "Хранение неликвида", rev=mps + st, cost=q * age * cfg.P("B34"))
                self.L(b["id"], pa, "Вывоз неликвида", rev=rm)
                self.L(b["id"], sl, "Хранение неликвида", cost=st); self.L(b["id"], sl, "Вывоз неликвида", cost=rm)
                b["removed"], b["left"], n = True, 0, n + 1
            return {"removed": n, "message": f"Вывезено партий: {n}"}
        return self.once(event_id, go)

    # --- представление ---
    def state(self, cfg):
        day, free, T = self.d["day"], cfg.P("B29"), cfg.P("B37")
        def tier(a): return "бесплатно (МП)" if a <= free else "платит продавец" if a <= T else "тариф ×%g" % cfg.P("B38") if a <= cfg.P("B39") else "неликвид"
        bs = [{**b, "age": day - b["day"], "tier": tier(day - b["day"])} for b in self.d["batches"]]
        stock = {}
        for b in bs:
            if b["removed"]: continue
            s = stock.setdefault(b["pvz"], {"pvz": b["pvz"], "phys": 0, "res": 0, "expired": 0})
            s["phys"] += b["left"]; s["res"] += b["res"]
            if b["age"] > cfg.P("B39"): s["expired"] += b["left"]
        for k, s in stock.items():
            s["limit"] = cfg.P({"S": "B41", "M": "B42", "L": "B43"}[self.d["sizes"].get(k, "M")])
            s["avail"] = s["phys"] - s["res"] - s["expired"]
        return {"day": day, "batches": bs, "stock": list(stock.values()), "orders": list(self.d["orders"].values()), "ledger": self.d["ledger"][-300:]}

    def analytics(self, cfg):
        pvz, seller, per = {}, {}, {}
        mp = 0.0
        for e in self.d["ledger"]:
            kind, _, code = e["party"].partition(":")
            if kind == "mp": mp += e["cost"]
            elif kind == "pvz":
                r = pvz.setdefault(code, {"pvz": code, "rev": 0.0, "cost": 0.0}); r["rev"] += e["rev"]; r["cost"] += e["cost"]
            elif kind == "seller":
                r = seller.setdefault(code, {"seller": code, "gmv": 0.0, "rev": 0.0, "cost": 0.0, "orders": set()})
                r["rev"] += e["rev"]; r["cost"] += e["cost"]; r["orders"].add(e["order"])
                if e["kind"].startswith("Продажа"): r["gmv"] += e["rev"]
                per[e["order"]] = per.get(e["order"], 0.0) + e["rev"] - e["cost"]
        done = [o for o in self.d["orders"].values() if o["status"] in ("DELIVERED", "RETURNED")]
        base = len(done) * (cfg.hub() + cfg.P("B27")) + sum(cfg.hub_return() for o in done if o["status"] == "RETURNED")
        rows = lambda d: [{k: (round(v, 2) if isinstance(v, float) else v) for k, v in r.items() if k != "orders"} | ({"orders": len(r["orders"]), "profit": round(r["rev"] - r["cost"], 2)} if "orders" in r else {"profit": round(r["rev"] - r["cost"], 2)}) for r in d.values()]
        return {"pvz": rows(pvz), "seller": rows(seller), "seller_orders": [{"order": k, "profit": round(v, 2)} for k, v in sorted(per.items())],
                "mp": {"orders": len(done), "spend": round(mp, 2), "baseline_hub": round(base, 2), "saving": round(base - mp, 2)}}
